/* Interface do Anonimizador — integra core.js com PapaParse/SheetJS. */

const Core = window.AnonimizadorCore;

const estado = {
  anon: { arquivo: null, colunasArquivo: null, colunasAlvo: null, colunasFiltro: null, valoresUnicos: null, encoding: null },
  desanon: { arquivo: null, mapa: null },
};

// ---------------------------------------------------------------------------
// Utilidades gerais
// ---------------------------------------------------------------------------
function ehExcel(nomeArquivo) {
  return /\.(xlsx|xls)$/i.test(nomeArquivo);
}

function formatarTempo(segundos) {
  segundos = Math.max(Math.round(segundos), 0);
  if (segundos < 60) return `${segundos}s`;
  const m = Math.floor(segundos / 60);
  const s = segundos % 60;
  if (m < 60) return `${m}m${String(s).padStart(2, "0")}s`;
  const h = Math.floor(m / 60);
  return `${h}h${String(m % 60).padStart(2, "0")}m`;
}

async function detectarEncoding(file) {
  const tamanhoAmostra = 1_000_000;
  const blobInicio = file.slice(0, tamanhoAmostra);
  const blobFim = file.size > tamanhoAmostra ? file.slice(file.size - tamanhoAmostra, file.size) : null;

  async function utf8Estrito(blob) {
    if (!blob) return true;
    const buffer = await blob.arrayBuffer();
    try {
      new TextDecoder("utf-8", { fatal: true }).decode(buffer);
      return true;
    } catch {
      return false;
    }
  }

  const ok = (await utf8Estrito(blobInicio)) && (await utf8Estrito(blobFim));
  return ok ? "utf-8" : "windows-1252";
}

function baixarTexto(texto, nomeArquivo, tipoMime) {
  const bom = "\uFEFF"; // BOM para o Excel abrir acentos corretamente
  const blob = new Blob([bom + texto], { type: tipoMime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nomeArquivo;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

// ---------------------------------------------------------------------------
// Leitura de amostra (CSV via PapaParse / Excel via SheetJS)
// ---------------------------------------------------------------------------
function lerAmostraCsv(file, sep, encoding, nLinhas = 2000) {
  // Importante: usar "preview" SEM um callback "chunk" faz o PapaParse
  // tentar ler o arquivo INTEIRO antes de aplicar o corte, o que travava
  // (e falhava silenciosamente) em arquivos grandes. Por isso paramos o
  // parser manualmente assim que juntamos linhas suficientes.
  return new Promise((resolve, reject) => {
    let linhas = [];
    let campos = [];
    let resolvido = false;

    function finalizar() {
      if (resolvido) return;
      resolvido = true;
      resolve({ colunas: campos, linhas: linhas.slice(0, nLinhas) });
    }

    Papa.parse(file, {
      header: true,
      delimiter: sep,
      encoding,
      skipEmptyLines: true,
      chunk: (results, parser) => {
        if (campos.length === 0) campos = results.meta.fields || [];
        linhas.push(...results.data);
        if (linhas.length >= nLinhas) {
          parser.abort();
          finalizar();
        }
      },
      complete: finalizar,
      error: reject,
    });
  });
}

function lerAmostraExcel(file, nLinhas = 2000) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const linhas = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
        const colunas = linhas.length > 0 ? Object.keys(linhas[0]) : [];
        resolve({ colunas, linhas: linhas.slice(0, nLinhas), todasLinhas: linhas });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ---------------------------------------------------------------------------
// Coleta de valores únicos (para os filtros) — full scan do arquivo
// ---------------------------------------------------------------------------
function coletarValoresUnicos(file, sep, encoding, colunasFiltro, aoProgredir) {
  if (Object.keys(colunasFiltro).length === 0) return Promise.resolve({});

  const unicos = {};
  for (const col of Object.keys(colunasFiltro)) unicos[col] = new Set();

  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      delimiter: sep,
      encoding,
      skipEmptyLines: true,
      chunk: (results, parser) => {
        for (const linha of results.data) {
          for (const col of Object.keys(colunasFiltro)) {
            const v = linha[col];
            if (v !== null && v !== undefined && String(v).trim() !== "") unicos[col].add(String(v));
          }
        }
        if (aoProgredir) aoProgredir(Math.min(results.meta.cursor / file.size, 1));
      },
      complete: () => {
        const saida = {};
        for (const [col, set] of Object.entries(unicos)) saida[col] = Array.from(set).sort();
        resolve(saida);
      },
      error: reject,
    });
  });
}

// ---------------------------------------------------------------------------
// Multiselect customizado (chips + painel de opções com busca)
// ---------------------------------------------------------------------------
function criarMultiselect(container, opcoes, rotulo) {
  container.innerHTML = "";
  const wrap = document.createElement("div");
  wrap.className = "multiselect";

  const titulo = document.createElement("label");
  titulo.className = "titulo-filtro";
  titulo.textContent = rotulo;
  wrap.appendChild(titulo);

  const botao = document.createElement("div");
  botao.className = "multiselect-botao";
  botao.textContent = "Selecionar valores...";
  wrap.appendChild(botao);

  const painel = document.createElement("div");
  painel.className = "multiselect-painel";
  const busca = document.createElement("input");
  busca.className = "multiselect-busca";
  busca.placeholder = "Buscar...";
  painel.appendChild(busca);

  const listaOpcoes = document.createElement("div");
  painel.appendChild(listaOpcoes);
  wrap.appendChild(painel);
  container.appendChild(wrap);

  const selecionados = new Set();

  function renderOpcoes(filtro = "") {
    listaOpcoes.innerHTML = "";
    const filtroNorm = filtro.toLowerCase();
    for (const opcao of opcoes) {
      if (filtroNorm && !String(opcao).toLowerCase().includes(filtroNorm)) continue;
      const item = document.createElement("div");
      item.className = "opcao";
      const chk = document.createElement("input");
      chk.type = "checkbox";
      chk.checked = selecionados.has(opcao);
      chk.addEventListener("change", () => {
        if (chk.checked) selecionados.add(opcao);
        else selecionados.delete(opcao);
        renderBotao();
      });
      const lbl = document.createElement("span");
      lbl.textContent = opcao;
      item.appendChild(chk);
      item.appendChild(lbl);
      item.addEventListener("click", (ev) => {
        if (ev.target !== chk) chk.click();
      });
      listaOpcoes.appendChild(item);
    }
  }

  function renderBotao() {
    botao.innerHTML = "";
    if (selecionados.size === 0) {
      botao.textContent = "Selecionar valores...";
      return;
    }
    for (const val of selecionados) {
      const chip = document.createElement("span");
      chip.className = "chip";
      chip.innerHTML = `<span>${val}</span><span class="x" data-val="${val}">×</span>`;
      chip.querySelector(".x").addEventListener("click", (ev) => {
        ev.stopPropagation();
        selecionados.delete(val);
        renderBotao();
        renderOpcoes(busca.value);
      });
      botao.appendChild(chip);
    }
  }

  botao.addEventListener("click", () => {
    document.querySelectorAll(".multiselect-painel.aberto").forEach((p) => {
      if (p !== painel) p.classList.remove("aberto");
    });
    painel.classList.toggle("aberto");
    if (painel.classList.contains("aberto")) busca.focus();
  });

  busca.addEventListener("input", () => renderOpcoes(busca.value));
  document.addEventListener("click", (ev) => {
    if (!wrap.contains(ev.target)) painel.classList.remove("aberto");
  });

  renderOpcoes();
  return { getSelecionados: () => Array.from(selecionados) };
}

document.addEventListener("DOMContentLoaded", () => {
  configurarAbas();
  configurarUploadAnon();
  configurarUploadDesanon();
});

function configurarAbas() {
  const botoes = document.querySelectorAll(".aba-botao");
  botoes.forEach((btn) => {
    btn.addEventListener("click", () => {
      botoes.forEach((b) => b.classList.remove("ativa"));
      btn.classList.add("ativa");
      document.querySelectorAll(".aba-conteudo").forEach((c) => c.classList.remove("visivel"));
      document.getElementById(btn.dataset.aba).classList.add("visivel");
    });
  });
}

// ---------------------------------------------------------------------------
// Aba Anonimizar
// ---------------------------------------------------------------------------
function configurarUploadAnon() {
  const area = document.getElementById("upload-anon-area");
  const input = document.getElementById("upload-anon-input");

  area.addEventListener("click", () => input.click());
  ["dragover", "dragenter"].forEach((ev) =>
    area.addEventListener(ev, (e) => {
      e.preventDefault();
      area.classList.add("arrastando");
    })
  );
  ["dragleave", "drop"].forEach((ev) =>
    area.addEventListener(ev, (e) => {
      e.preventDefault();
      area.classList.remove("arrastando");
    })
  );
  area.addEventListener("drop", (e) => {
    if (e.dataTransfer.files.length > 0) receberArquivoAnon(e.dataTransfer.files[0]);
  });
  input.addEventListener("change", (e) => {
    if (e.target.files.length > 0) receberArquivoAnon(e.target.files[0]);
  });

  document.getElementById("btn-iniciar-anon").addEventListener("click", iniciarAnonimizacao);
}

async function receberArquivoAnon(file) {
  estado.anon.arquivo = file;
  const info = document.getElementById("upload-anon-info");
  info.innerHTML = `📄 <strong>${file.name}</strong> — ${(file.size / 1_048_576).toFixed(1)} MB`;
  info.classList.remove("oculto");

  document.getElementById("bloco-config-anon").classList.add("oculto");
  document.getElementById("resultado-anon").classList.add("oculto");
  document.getElementById("erro-analise-anon").classList.add("oculto");
  document.getElementById("aviso-sem-colunas-anon").classList.add("oculto");

  const LIMITE_SEGURO_SEM_STREAMING = 150 * 1024 * 1024; // 150MB
  const avisoTamanho = document.getElementById("aviso-tamanho-arquivo-anon");
  if (!window.showSaveFilePicker && file.size > LIMITE_SEGURO_SEM_STREAMING) {
    avisoTamanho.innerHTML =
      `⚠️ Este arquivo tem ${(file.size / 1_048_576).toFixed(0)}MB e seu navegador não suporta o modo de gravação ` +
      `direta em disco (mais seguro para arquivos grandes). Há um risco real de a aba travar por falta de memória. ` +
      `<strong>Recomendamos fortemente usar Google Chrome ou Microsoft Edge atualizados</strong> para arquivos acima de 150MB.`;
    avisoTamanho.classList.remove("oculto");
  } else {
    avisoTamanho.classList.add("oculto");
  }

  const spinner = document.getElementById("spinner-analise-anon");
  spinner.classList.remove("oculto");
  spinner.innerHTML = `<span class="spinner"></span> Analisando colunas e valores...`;

  try {
    let colunas, linhasAmostra, encoding;
    const sep = document.getElementById("separador-anon").value;
    const encSelecionado = document.getElementById("encoding-anon").value;

    if (ehExcel(file.name)) {
      const r = await lerAmostraExcel(file);
      colunas = r.colunas;
      linhasAmostra = r.linhas;
      encoding = null;
    } else {
      encoding = encSelecionado === "auto" ? await detectarEncoding(file) : encSelecionado;
      const r = await lerAmostraCsv(file, sep, encoding);
      colunas = r.colunas;
      linhasAmostra = r.linhas;
    }

    const { colunasAlvo, colunasFiltro } = Core.analisarAmostra(colunas, linhasAmostra);

    spinner.innerHTML = `<span class="spinner"></span> Coletando valores para os filtros...`;
    const valoresUnicos = ehExcel(file.name)
      ? coletarValoresUnicosExcel((await lerAmostraExcel(file, Infinity)).todasLinhas, colunasFiltro)
      : await coletarValoresUnicos(file, sep, encoding, colunasFiltro);

    estado.anon.colunasArquivo = colunas;
    estado.anon.colunasAlvo = colunasAlvo;
    estado.anon.colunasFiltro = colunasFiltro;
    estado.anon.valoresUnicos = valoresUnicos;
    estado.anon.encoding = encoding;

    renderizarConfigAnon();
  } catch (err) {
    console.error(err);
    document.getElementById("erro-analise-anon").textContent = "Erro ao analisar o arquivo: " + err.message;
    document.getElementById("erro-analise-anon").classList.remove("oculto");
  } finally {
    spinner.classList.add("oculto");
  }
}

function coletarValoresUnicosExcel(todasLinhas, colunasFiltro) {
  const unicos = {};
  for (const col of Object.keys(colunasFiltro)) unicos[col] = new Set();
  for (const linha of todasLinhas) {
    for (const col of Object.keys(colunasFiltro)) {
      const v = linha[col];
      if (v !== null && v !== undefined && String(v).trim() !== "") unicos[col].add(String(v));
    }
  }
  const saida = {};
  for (const [col, set] of Object.entries(unicos)) saida[col] = Array.from(set).sort();
  return saida;
}

function renderizarConfigAnon() {
  const { colunasArquivo, colunasAlvo, colunasFiltro, valoresUnicos } = estado.anon;
  const bloco = document.getElementById("bloco-config-anon");

  if (Object.keys(colunasAlvo).length === 0) {
    document.getElementById("aviso-sem-colunas-anon").classList.remove("oculto");
    return;
  }
  bloco.classList.remove("oculto");

  const grade = document.getElementById("grade-colunas-anon");
  grade.innerHTML = "";
  const checkboxes = [];
  for (const col of colunasArquivo) {
    const idx = colunasAlvo[col];
    const item = document.createElement("div");
    item.className = "checkbox-item";
    const chk = document.createElement("input");
    chk.type = "checkbox";
    chk.checked = true;
    chk.id = `col_final_${col}`;
    const lbl = document.createElement("label");
    lbl.htmlFor = chk.id;
    lbl.innerHTML = idx !== undefined ? `${col} <span class="marcador-alvo">🔒</span>` : col;
    item.appendChild(chk);
    item.appendChild(lbl);
    grade.appendChild(item);
    checkboxes.push(chk);
  }

  const chkTodas = document.getElementById("chk-selecionar-todas-anon");
  chkTodas.checked = true;
  chkTodas.onchange = () => checkboxes.forEach((c) => (c.checked = chkTodas.checked));

  const blocoFiltros = document.getElementById("bloco-filtros-anon");
  const gradeFiltros = document.getElementById("grade-filtros-anon");
  gradeFiltros.innerHTML = "";
  const multiselects = {};
  const entradas = Object.entries(colunasFiltro);
  if (entradas.length > 0) {
    blocoFiltros.classList.remove("oculto");
    for (const [col, nome] of entradas) {
      const div = document.createElement("div");
      gradeFiltros.appendChild(div);
      multiselects[col] = criarMultiselect(div, valoresUnicos[col] || [], `${nome.charAt(0).toUpperCase() + nome.slice(1)} (${col})`);
    }
  } else {
    blocoFiltros.classList.add("oculto");
  }

  estado.anon._checkboxes = checkboxes;
  estado.anon._multiselects = multiselects;
}

async function criarEscritorSaida(nomeArquivoSugerido) {
  // Grava direto em disco, em streaming, quando o navegador suporta (Chrome/
  // Edge) -- evita reter o CSV inteiro na memória, o que causava lentidão
  // crescente (pressão de coletor de lixo) em arquivos grandes. Se não
  // suportado, cai no modo antigo (acumula e baixa como Blob ao final --
  // mais lento e usa mais memória, mas ainda funciona).
  if (window.showSaveFilePicker) {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: nomeArquivoSugerido,
        types: [{ description: "CSV", accept: { "text/csv": [".csv"] } }],
      });
      const writable = await handle.createWritable();
      return {
        modo: "stream",
        async escrever(texto) {
          await writable.write(texto);
        },
        async finalizar() {
          await writable.close();
        },
        async cancelar() {
          await writable.abort();
        },
      };
    } catch (e) {
      if (e.name === "AbortError") throw e; // usuário cancelou o diálogo de propósito
      // qualquer outro erro: cai no fallback abaixo
    }
  }

  const partes = [];
  return {
    modo: "blob",
    async escrever(texto) {
      partes.push(texto);
    },
    async finalizar() {
      const blob = new Blob(["\uFEFF", ...partes], { type: "text/csv;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = nomeArquivoSugerido;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 5000);
    },
    async cancelar() {
      partes.length = 0;
    },
  };
}

async function iniciarAnonimizacao() {
  const btn = document.getElementById("btn-iniciar-anon");
  const colunasFinais = new Set();
  for (const chk of estado.anon._checkboxes) {
    if (chk.checked) colunasFinais.add(chk.id.replace("col_final_", ""));
  }
  if (colunasFinais.size === 0) {
    alert("Selecione ao menos uma coluna para o arquivo final.");
    return;
  }

  const filtros = {};
  for (const [col, ms] of Object.entries(estado.anon._multiselects)) {
    const sel = ms.getSelecionados();
    if (sel.length > 0) filtros[col] = sel;
  }

  const LIMITE_SEGURO_SEM_STREAMING = 150 * 1024 * 1024;
  if (!window.showSaveFilePicker && estado.anon.arquivo.size > LIMITE_SEGURO_SEM_STREAMING) {
    const continuar = confirm(
      "Este arquivo é grande e seu navegador não suporta o modo mais seguro de gravação. " +
      "Existe um risco real de a aba travar por falta de memória.\n\n" +
      "Recomendamos cancelar e usar Google Chrome ou Microsoft Edge atualizados.\n\n" +
      "Quer continuar mesmo assim?"
    );
    if (!continuar) return;
  }

  // Pede o local do arquivo de saída JÁ NO CLIQUE (precisa ser assim para o
  // diálogo nativo de "Salvar como" funcionar -- navegadores só permitem
  // abri-lo em resposta direta a um gesto do usuário).
  let escritor;
  try {
    escritor = await criarEscritorSaida("dados_anonimizados.csv");
  } catch (e) {
    if (e.name === "AbortError") return; // usuário cancelou o diálogo -- só volta
    console.error(e);
    alert("Não foi possível preparar o arquivo de saída: " + e.message);
    return;
  }

  btn.disabled = true;
  document.getElementById("resultado-anon").classList.add("oculto");
  const progressoWrap = document.getElementById("progresso-anon");
  progressoWrap.classList.remove("oculto");
  if (escritor.modo === "blob") {
    document.getElementById("aviso-modo-blob-anon").classList.remove("oculto");
  }

  try {
    const resultado = await processarAnonimizacao(colunasFinais, filtros, escritor);
    await escritor.finalizar();
    exibirResultadoAnon(resultado);
  } catch (err) {
    console.error(err);
    await escritor.cancelar();
    alert("Erro durante a anonimização: " + err.message);
  } finally {
    btn.disabled = false;
    progressoWrap.classList.add("oculto");
  }
}

function processarAnonimizacao(colunasFinais, filtros, escritor) {
  const file = estado.anon.arquivo;
  const sep = document.getElementById("separador-anon").value;
  const encoding = estado.anon.encoding;
  const colunasAlvo = estado.anon.colunasAlvo;
  const colunasArquivo = estado.anon.colunasArquivo;
  const colunasSaida = colunasArquivo.filter((c) => colunasFinais.has(c));
  const colunasAlvoFiltradas = {};
  for (const [c, idx] of Object.entries(colunasAlvo)) {
    if (colunasFinais.has(c)) colunasAlvoFiltradas[c] = idx;
  }

  const mapeador = new Core.MapeadorAnonimizacao();
  const inicio = performance.now();
  let linhasLidas = 0;
  let linhasMantidas = 0;
  let primeiroChunk = true;
  const promessasEscrita = [];

  const filtrar = Object.keys(filtros).length > 0
    ? (linha) => {
        for (const [col, valores] of Object.entries(filtros)) {
          if (!valores.includes(String(linha[col] ?? ""))) return false;
        }
        return true;
      }
    : null;

  function transformarLinhas(linhasBrutas) {
    linhasLidas += linhasBrutas.length;
    let linhas = filtrar ? linhasBrutas.filter(filtrar) : linhasBrutas;

    for (const linha of linhas) {
      for (const [col, idx] of Object.entries(colunasAlvoFiltradas)) {
        linha[col] = mapeador.anonimizarValor(linha[col], idx);
      }
    }
    linhasMantidas += linhas.length;

    const linhasArray = linhas.map((l) => colunasSaida.map((c) => (l[c] === undefined || l[c] === null ? "" : l[c])));
    const texto = primeiroChunk
      ? Papa.unparse({ fields: colunasSaida, data: linhasArray }, { delimiter: sep, newline: "\n" })
      : Papa.unparse(linhasArray, { delimiter: sep, newline: "\n" });
    primeiroChunk = false;
    // Escreve e DESCARTA o texto imediatamente -- nada de acumular em array.
    promessasEscrita.push(escritor.escrever(texto + "\n"));
  }

  if (ehExcel(file.name)) {
    return lerAmostraExcel(file, Infinity).then(async (r) => {
      transformarLinhas(r.todasLinhas);
      await Promise.all(promessasEscrita);
      atualizarProgressoAnon(1, linhasMantidas, linhasLidas, inicio);
      return finalizarAnonimizacao(mapeador, linhasMantidas, linhasLidas);
    });
  }

  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      delimiter: sep,
      encoding,
      skipEmptyLines: true,
      chunk: (results, parser) => {
        try {
          transformarLinhas(results.data);
          const fracao = Math.min(results.meta.cursor / file.size, 1);
          atualizarProgressoAnon(fracao, linhasMantidas, linhasLidas, inicio);
        } catch (e) {
          parser.abort();
          reject(e);
        }
      },
      complete: async () => {
        if (linhasMantidas === 0) {
          reject(new Error("Nenhuma linha correspondeu ao filtro selecionado."));
          return;
        }
        try {
          await Promise.all(promessasEscrita);
          resolve(finalizarAnonimizacao(mapeador, linhasMantidas, linhasLidas));
        } catch (e) {
          reject(e);
        }
      },
      error: reject,
    });
  });
}

function finalizarAnonimizacao(mapeador, linhasMantidas, linhasLidas) {
  const mapeamento = mapeador.construirMapeamento();
  const mapaCsv = Papa.unparse(
    {
      fields: ["coluna", "valor_original", "codigo"],
      data: mapeamento.map((l) => [l.coluna, l.valor_original, l.codigo]),
    },
    { newline: "\n" }
  );
  return { mapaCsv, totalLinhas: linhasMantidas, totalLidas: linhasLidas, totalMapeados: mapeamento.length };
}

function atualizarProgressoAnon(fracao, mantidas, lidas, inicioMs) {
  const preenchido = document.querySelector("#progresso-anon .preenchido");
  preenchido.style.width = `${(fracao * 100).toFixed(0)}%`;
  const elapsed = (performance.now() - inicioMs) / 1000;
  const eta = fracao > 0.02 ? formatarTempo(elapsed / fracao - elapsed) : "calculando...";
  document.querySelector("#progresso-anon .progresso-texto").textContent =
    `Anonimizando... ${(fracao * 100).toFixed(1)}% — ${mantidas.toLocaleString("pt-BR")} linhas mantidas de ${lidas.toLocaleString("pt-BR")} lidas — tempo restante: ${eta}`;
}

function exibirResultadoAnon(r) {
  const div = document.getElementById("resultado-anon");
  div.classList.remove("oculto");
  document.getElementById("metrica-linhas-anon").textContent = r.totalLinhas.toLocaleString("pt-BR");
  document.getElementById("metrica-mapeados-anon").textContent = r.totalMapeados.toLocaleString("pt-BR");

  document.getElementById("btn-download-mapa-anon").onclick = () =>
    baixarTexto(r.mapaCsv, "mapeamento_anonimizacao.csv", "text/csv;charset=utf-8");
}

// ---------------------------------------------------------------------------
// Aba Desanonimizar
// ---------------------------------------------------------------------------
function configurarUploadDesanon() {
  const areaArquivo = document.getElementById("upload-desanon-area");
  const inputArquivo = document.getElementById("upload-desanon-input");
  const areaMapa = document.getElementById("upload-mapa-area");
  const inputMapa = document.getElementById("upload-mapa-input");

  areaArquivo.addEventListener("click", () => inputArquivo.click());
  inputArquivo.addEventListener("change", (e) => {
    if (e.target.files.length > 0) {
      estado.desanon.arquivo = e.target.files[0];
      document.getElementById("upload-desanon-info").innerHTML =
        `📄 <strong>${e.target.files[0].name}</strong> — ${(e.target.files[0].size / 1_048_576).toFixed(1)} MB`;
      document.getElementById("upload-desanon-info").classList.remove("oculto");
      verificarProntoDesanon();
    }
  });

  areaMapa.addEventListener("click", () => inputMapa.click());
  inputMapa.addEventListener("change", (e) => {
    if (e.target.files.length > 0) {
      estado.desanon.arquivoMapa = e.target.files[0];
      document.getElementById("upload-mapa-info").innerHTML = `📄 <strong>${e.target.files[0].name}</strong>`;
      document.getElementById("upload-mapa-info").classList.remove("oculto");
      verificarProntoDesanon();
    }
  });

  document.getElementById("btn-iniciar-desanon").addEventListener("click", iniciarDesanonimizacao);
}

function verificarProntoDesanon() {
  const pronto = estado.desanon.arquivo && estado.desanon.arquivoMapa;
  document.getElementById("btn-iniciar-desanon").disabled = !pronto;
}

function lerMapaCsv(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results) => resolve(results.data),
      error: reject,
    });
  });
}

async function iniciarDesanonimizacao() {
  const btn = document.getElementById("btn-iniciar-desanon");

  let escritor;
  try {
    escritor = await criarEscritorSaida("dados_desanonimizados.csv");
  } catch (e) {
    if (e.name === "AbortError") return;
    console.error(e);
    alert("Não foi possível preparar o arquivo de saída: " + e.message);
    return;
  }

  btn.disabled = true;
  document.getElementById("resultado-desanon").classList.add("oculto");
  const progressoWrap = document.getElementById("progresso-desanon");
  progressoWrap.classList.remove("oculto");
  if (escritor.modo === "blob") {
    document.getElementById("aviso-modo-blob-desanon").classList.remove("oculto");
  }

  try {
    const linhasMapeamento = await lerMapaCsv(estado.desanon.arquivoMapa);
    const mapaReverso = Core.construirMapaReverso(linhasMapeamento);

    const file = estado.desanon.arquivo;
    const sep = document.getElementById("separador-desanon").value;
    let encoding = document.getElementById("encoding-desanon").value;
    if (encoding === "auto" && !ehExcel(file.name)) encoding = await detectarEncoding(file);

    let colunasArquivo, todasLinhasExcel;
    if (ehExcel(file.name)) {
      const r = await lerAmostraExcel(file, Infinity);
      colunasArquivo = r.colunas;
      todasLinhasExcel = r.todasLinhas;
    } else {
      const r = await lerAmostraCsv(file, sep, encoding, 1000);
      colunasArquivo = r.colunas;
    }

    const candidatos = Core.identificarColunas(colunasArquivo);
    const colunasParaRestaurar = {};
    for (const [col, idx] of Object.entries(candidatos)) {
      const nome = Core.COLUNAS_ALVO[idx].nome;
      if (mapaReverso[nome]) colunasParaRestaurar[col] = nome;
    }

    if (Object.keys(colunasParaRestaurar).length === 0) {
      throw new Error("Nenhuma coluna do arquivo bate com o mapeamento enviado.");
    }

    const inicio = performance.now();
    let linhasProcessadas = 0;
    let primeiroChunk = true;
    const promessasEscrita = [];

    function transformar(linhasBrutas) {
      linhasProcessadas += linhasBrutas.length;
      for (const linha of linhasBrutas) {
        for (const [col, nome] of Object.entries(colunasParaRestaurar)) {
          linha[col] = Core.desanonimizarValor(linha[col], mapaReverso[nome]);
        }
      }
      const linhasArray = linhasBrutas.map((l) => colunasArquivo.map((c) => (l[c] === undefined || l[c] === null ? "" : l[c])));
      const texto = primeiroChunk
        ? Papa.unparse({ fields: colunasArquivo, data: linhasArray }, { delimiter: sep, newline: "\n" })
        : Papa.unparse(linhasArray, { delimiter: sep, newline: "\n" });
      primeiroChunk = false;
      promessasEscrita.push(escritor.escrever(texto + "\n"));
    }

    if (ehExcel(file.name)) {
      transformar(todasLinhasExcel);
      await Promise.all(promessasEscrita);
      atualizarProgressoDesanon(1, linhasProcessadas, inicio);
    } else {
      await new Promise((resolve, reject) => {
        Papa.parse(file, {
          header: true,
          delimiter: sep,
          encoding,
          skipEmptyLines: true,
          chunk: (results, parser) => {
            try {
              transformar(results.data);
              atualizarProgressoDesanon(Math.min(results.meta.cursor / file.size, 1), linhasProcessadas, inicio);
            } catch (e) {
              parser.abort();
              reject(e);
            }
          },
          complete: resolve,
          error: reject,
        });
      });
      await Promise.all(promessasEscrita);
    }

    await escritor.finalizar();
    document.getElementById("resultado-desanon").classList.remove("oculto");
    document.getElementById("metrica-linhas-desanon").textContent = linhasProcessadas.toLocaleString("pt-BR");
  } catch (err) {
    console.error(err);
    await escritor.cancelar();
    alert("Erro durante a desanonimização: " + err.message);
  } finally {
    btn.disabled = false;
    progressoWrap.classList.add("oculto");
  }
}

function atualizarProgressoDesanon(fracao, linhas, inicioMs) {
  const preenchido = document.querySelector("#progresso-desanon .preenchido");
  preenchido.style.width = `${(fracao * 100).toFixed(0)}%`;
  const elapsed = (performance.now() - inicioMs) / 1000;
  const eta = fracao > 0.02 ? formatarTempo(elapsed / fracao - elapsed) : "calculando...";
  document.querySelector("#progresso-desanon .progresso-texto").textContent =
    `Desanonimizando... ${(fracao * 100).toFixed(1)}% — ${linhas.toLocaleString("pt-BR")} linhas — tempo restante: ${eta}`;
}
