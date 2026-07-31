/* Interface do Anonimizador — integra core.js com PapaParse/SheetJS. */

const Core = window.AnonimizadorCore;

const COLUNA_ORIGEM = "Arquivo_Origem";
const LIMITE_SEGURO_SEM_STREAMING = 150 * 1024 * 1024; // 150MB
const LIMITE_RECOMENDADO_TOTAL = 2 * 1024 * 1024 * 1024; // 2GB

const estado = {
  anon: { arquivos: [], colunasArquivo: null, colunasAlvo: null, colunasFiltro: null, valoresUnicos: null, encoding: null },
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

// ---------------------------------------------------------------------------
// Diagnóstico amigável: em vez de mostrar um erro técnico quando a leitura
// dá errado, tenta identificar se o problema é separador ou codificação, e
// recomenda a mudança específica.
// ---------------------------------------------------------------------------
async function lerAmostraTextoBruto(file, encoding, tamanhoBytes = 65536) {
  const blob = file.slice(0, tamanhoBytes);
  const buffer = await blob.arrayBuffer();
  const label = (encoding || "utf-8").toLowerCase();
  try {
    return new TextDecoder(label === "auto" ? "utf-8" : label).decode(buffer);
  } catch {
    return new TextDecoder("utf-8").decode(buffer);
  }
}

function contarDelimitadores(linha) {
  return {
    ";": (linha.match(/;/g) || []).length,
    ",": (linha.match(/,/g) || []).length,
    "\t": (linha.match(/\t/g) || []).length,
    "|": (linha.match(/\|/g) || []).length,
  };
}

function diagnosticarLeitura(textoAmostra, colunas, sepAtual) {
  // 1) Poucas colunas detectadas costuma ser separador errado -- confere
  // se outro separador comum aparece muito mais vezes na primeira linha.
  if (colunas.length <= 1) {
    const primeiraLinha = (textoAmostra.split(/\r?\n/)[0] || "");
    const contagens = contarDelimitadores(primeiraLinha);
    delete contagens[sepAtual];
    const [melhorSep, melhorContagem] = Object.entries(contagens).sort((a, b) => b[1] - a[1])[0];
    if (melhorContagem > 0) {
      const nomeSep = melhorSep === "\t" ? "Tab" : `"${melhorSep}"`;
      return {
        tipo: "separador",
        mensagem: `O separador selecionado não parece bater com o arquivo -- ele parece usar ${nomeSep}. Troque em "Opções avançadas" e tente de novo.`,
      };
    }
  }

  // 2) Caracteres de acentuação corrompidos indicam codificação errada.
  const temSubstituicao = textoAmostra.includes("\uFFFD");
  const temMojibake = /[ÃÂ][\u0080-\u00BF]/.test(textoAmostra);
  if (temSubstituicao || temMojibake) {
    return {
      tipo: "codificacao",
      mensagem: `A codificação selecionada não parece bater com o arquivo -- apareceram caracteres estranhos no lugar de acentos. Tente "Windows-1252 / CP1252" em "Opções avançadas" e tente de novo.`,
    };
  }

  return null;
}

function mostrarErroLeitura(elementoId, mensagem) {
  const el = document.getElementById(elementoId);
  el.innerHTML = `<i class="ti ti-alert-triangle"></i> ${mensagem}`;
  el.classList.remove("oculto");
}

function erroAmigavel(mensagem) {
  const e = new Error(mensagem);
  e.amigavel = true;
  return e;
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
// Coleta de valores únicos (para os filtros) — feita em coletarValoresUnicosMultiplos,
// que também cobre o caso de um único arquivo.
// ---------------------------------------------------------------------------

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
  const botoes = document.querySelectorAll(".mode-btn");
  botoes.forEach((btn) => {
    btn.addEventListener("click", () => {
      botoes.forEach((b) => b.classList.remove("is-active"));
      btn.classList.add("is-active");
      document.querySelectorAll(".mode-panel").forEach((c) => c.classList.remove("is-visible"));
      document.getElementById(btn.dataset.modo).classList.add("is-visible");
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
    if (e.dataTransfer.files.length > 0) adicionarArquivosAnon(Array.from(e.dataTransfer.files));
  });
  input.addEventListener("change", (e) => {
    if (e.target.files.length > 0) adicionarArquivosAnon(Array.from(e.target.files));
    input.value = ""; // permite selecionar o mesmo arquivo de novo depois de remover
  });

  document.getElementById("btn-iniciar-anon").addEventListener("click", iniciarAnonimizacao);
}

function adicionarArquivosAnon(novosArquivos) {
  // Evita duplicar o mesmo arquivo (mesmo nome + tamanho) se selecionado de novo.
  for (const novo of novosArquivos) {
    const jaExiste = estado.anon.arquivos.some((a) => a.name === novo.name && a.size === novo.size);
    if (!jaExiste) estado.anon.arquivos.push(novo);
  }
  renderizarListaArquivosAnon();
  analisarArquivosAnon();
}

function removerArquivoAnon(indice) {
  estado.anon.arquivos.splice(indice, 1);
  renderizarListaArquivosAnon();
  if (estado.anon.arquivos.length > 0) {
    analisarArquivosAnon();
  } else {
    document.getElementById("bloco-config-anon").classList.add("oculto");
    document.getElementById("resultado-anon").classList.add("oculto");
  }
}

function renderizarListaArquivosAnon() {
  const lista = document.getElementById("lista-arquivos-anon");
  const arquivos = estado.anon.arquivos;

  if (arquivos.length === 0) {
    lista.classList.add("oculto");
    lista.innerHTML = "";
    return;
  }

  lista.classList.remove("oculto");
  lista.innerHTML = "";
  let tamanhoTotal = 0;

  arquivos.forEach((arquivo, i) => {
    tamanhoTotal += arquivo.size;
    const item = document.createElement("div");
    item.className = "item-arquivo";
    item.innerHTML =
      `<span class="nome"><i class="ti ti-file-text"></i> ${arquivo.name} — ${(arquivo.size / 1_048_576).toFixed(1)} MB</span>` +
      `<button class="remover" title="Remover">✕</button>`;
    item.querySelector(".remover").addEventListener("click", () => removerArquivoAnon(i));
    lista.appendChild(item);
  });

  const total = document.createElement("div");
  total.className = "total";
  const totalGB = tamanhoTotal / (1024 * 1024 * 1024);
  total.textContent = `${arquivos.length} arquivo(s) — total: ${totalGB >= 1 ? totalGB.toFixed(2) + " GB" : (tamanhoTotal / 1_048_576).toFixed(1) + " MB"}`;
  lista.appendChild(total);

  const avisoTamanho = document.getElementById("aviso-tamanho-arquivo-anon");
  if (!window.showSaveFilePicker && tamanhoTotal > LIMITE_SEGURO_SEM_STREAMING) {
    avisoTamanho.innerHTML =
      `<i class="ti ti-alert-triangle"></i> O total selecionado passa de 150MB e seu navegador não suporta o modo de gravação direta em disco ` +
      `(mais seguro para arquivos grandes). Há um risco real de a aba travar por falta de memória. ` +
      `<strong>Recomendamos fortemente usar Google Chrome ou Microsoft Edge atualizados.</strong>`;
    avisoTamanho.classList.remove("oculto");
  } else if (tamanhoTotal > LIMITE_RECOMENDADO_TOTAL) {
    avisoTamanho.innerHTML =
      `<i class="ti ti-alert-triangle"></i> O total selecionado passa de 2GB. Deve funcionar em Chrome/Edge, mas o processamento pode demorar vários minutos.`;
    avisoTamanho.classList.remove("oculto");
  } else {
    avisoTamanho.classList.add("oculto");
  }
}

async function obterColunasArquivo(file, sep, encoding) {
  if (ehExcel(file.name)) {
    const r = await lerAmostraExcel(file, 1);
    return r.colunas;
  }
  const r = await lerAmostraCsv(file, sep, encoding, 1);
  return r.colunas;
}

async function analisarArquivosAnon() {
  const arquivos = estado.anon.arquivos;
  if (arquivos.length === 0) return;

  document.getElementById("bloco-config-anon").classList.add("oculto");
  document.getElementById("resultado-anon").classList.add("oculto");
  document.getElementById("erro-analise-anon").classList.add("oculto");
  document.getElementById("aviso-sem-colunas-anon").classList.add("oculto");

  const spinner = document.getElementById("spinner-analise-anon");
  spinner.classList.remove("oculto");
  spinner.innerHTML = `<span class="spinner"></span> Analisando colunas e valores...`;

  try {
    const primeiro = arquivos[0];
    let colunas, linhasAmostra, encoding;
    const sep = document.getElementById("separador-anon").value;
    const encSelecionado = document.getElementById("encoding-anon").value;

    if (ehExcel(primeiro.name)) {
      const r = await lerAmostraExcel(primeiro);
      colunas = r.colunas;
      linhasAmostra = r.linhas;
      encoding = null;
    } else {
      encoding = encSelecionado === "auto" ? await detectarEncoding(primeiro) : encSelecionado;
      const r = await lerAmostraCsv(primeiro, sep, encoding);
      colunas = r.colunas;
      linhasAmostra = r.linhas;

      // Antes de seguir, confere se o resultado tem cara de separador ou
      // codificação errados -- e, se tiver, avisa exatamente o que trocar
      // em vez de deixar o erro estourar mais na frente.
      const textoBruto = await lerAmostraTextoBruto(primeiro, encoding);
      const diagnostico = diagnosticarLeitura(textoBruto, colunas, sep);
      if (diagnostico) {
        spinner.classList.add("oculto");
        mostrarErroLeitura("erro-analise-anon", diagnostico.mensagem);
        return;
      }
    }

    // Se houver mais de um arquivo, todos precisam ter exatamente as
    // mesmas colunas do primeiro -- senão o empilhamento fica sem sentido.
    if (arquivos.length > 1) {
      const divergentes = [];
      for (let i = 1; i < arquivos.length; i++) {
        const colsOutro = await obterColunasArquivo(arquivos[i], sep, encoding);
        const igual = colunas.length === colsOutro.length && colunas.every((c, idx) => c === colsOutro[idx]);
        if (!igual) divergentes.push(arquivos[i].name);
      }
      if (divergentes.length > 0) {
        throw erroAmigavel(
          `Estes arquivos têm colunas diferentes do primeiro (${primeiro.name}) e não podem ser empilhados: ${divergentes.join(", ")}. Remova-os ou ajuste as colunas antes de continuar.`
        );
      }
    }

    const { colunasAlvo, colunasFiltro } = Core.analisarAmostra(colunas, linhasAmostra);

    spinner.innerHTML = `<span class="spinner"></span> Coletando valores para os filtros...`;
    const valoresUnicos = await coletarValoresUnicosMultiplos(arquivos, sep, encoding, colunasFiltro);

    // Coluna sintética com o nome do arquivo de origem -- útil sobretudo
    // quando há mais de um arquivo empilhado, mas fica disponível sempre.
    const colunasArquivoFinal = [...colunas, COLUNA_ORIGEM];
    const colunasFiltroFinal = { ...colunasFiltro, [COLUNA_ORIGEM]: "arquivo de origem" };
    valoresUnicos[COLUNA_ORIGEM] = arquivos.map((a) => a.name);

    estado.anon.colunasArquivo = colunasArquivoFinal;
    estado.anon.colunasAlvo = colunasAlvo;
    estado.anon.colunasFiltro = colunasFiltroFinal;
    estado.anon.valoresUnicos = valoresUnicos;
    estado.anon.encoding = encoding;

    renderizarConfigAnon();
  } catch (err) {
    console.error(err);
    const mensagem = err.amigavel
      ? err.message
      : `Não conseguimos ler esse arquivo. Verifique se o separador e a codificação em "Opções avançadas" correspondem ao arquivo, ou tente novamente.`;
    mostrarErroLeitura("erro-analise-anon", mensagem);
  } finally {
    spinner.classList.add("oculto");
  }
}

async function coletarValoresUnicosMultiplos(arquivos, sep, encoding, colunasFiltro) {
  if (Object.keys(colunasFiltro).length === 0) return {};
  const unicos = {};
  for (const col of Object.keys(colunasFiltro)) unicos[col] = new Set();

  for (const arquivo of arquivos) {
    if (ehExcel(arquivo.name)) {
      const r = await lerAmostraExcel(arquivo, Infinity);
      for (const linha of r.todasLinhas) {
        for (const col of Object.keys(colunasFiltro)) {
          const v = linha[col];
          if (v !== null && v !== undefined && String(v).trim() !== "") unicos[col].add(String(v));
        }
      }
    } else {
      await new Promise((resolve, reject) => {
        Papa.parse(arquivo, {
          header: true,
          delimiter: sep,
          encoding,
          skipEmptyLines: true,
          chunk: (results) => {
            for (const linha of results.data) {
              for (const col of Object.keys(colunasFiltro)) {
                const v = linha[col];
                if (v !== null && v !== undefined && String(v).trim() !== "") unicos[col].add(String(v));
              }
            }
          },
          complete: resolve,
          error: reject,
        });
      });
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

  const tamanhoTotal = estado.anon.arquivos.reduce((s, a) => s + a.size, 0);
  if (!window.showSaveFilePicker && tamanhoTotal > LIMITE_SEGURO_SEM_STREAMING) {
    const continuar = confirm(
      "O total dos arquivos é grande e seu navegador não suporta o modo mais seguro de gravação. " +
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

async function processarAnonimizacao(colunasFinais, filtros, escritor) {
  const arquivos = estado.anon.arquivos;
  const sep = document.getElementById("separador-anon").value;
  const encoding = estado.anon.encoding;
  const colunasAlvo = estado.anon.colunasAlvo;
  const colunasArquivo = estado.anon.colunasArquivo; // já inclui COLUNA_ORIGEM
  const colunasSaida = colunasArquivo.filter((c) => colunasFinais.has(c));
  const colunasAlvoFiltradas = {};
  for (const [c, idx] of Object.entries(colunasAlvo)) {
    if (colunasFinais.has(c)) colunasAlvoFiltradas[c] = idx;
  }
  const precisaOrigem = colunasFinais.has(COLUNA_ORIGEM) || Boolean(filtros[COLUNA_ORIGEM]);

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

  function transformarLinhas(linhasBrutas, nomeArquivoAtual) {
    linhasLidas += linhasBrutas.length;

    if (precisaOrigem) {
      for (const linha of linhasBrutas) linha[COLUNA_ORIGEM] = nomeArquivoAtual;
    }

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

  const tamanhoTotal = arquivos.reduce((s, a) => s + a.size, 0) || 1;
  let bytesAnteriores = 0;

  for (const arquivo of arquivos) {
    if (ehExcel(arquivo.name)) {
      const r = await lerAmostraExcel(arquivo, Infinity);
      transformarLinhas(r.todasLinhas, arquivo.name);
      bytesAnteriores += arquivo.size;
      atualizarProgressoAnon(bytesAnteriores / tamanhoTotal, linhasMantidas, linhasLidas, inicio);
    } else {
      const bytesAntesDesteArquivo = bytesAnteriores;
      await new Promise((resolve, reject) => {
        Papa.parse(arquivo, {
          header: true,
          delimiter: sep,
          encoding,
          skipEmptyLines: true,
          chunk: (results, parser) => {
            try {
              transformarLinhas(results.data, arquivo.name);
              const fracaoGeral = (bytesAntesDesteArquivo + results.meta.cursor) / tamanhoTotal;
              atualizarProgressoAnon(Math.min(fracaoGeral, 1), linhasMantidas, linhasLidas, inicio);
            } catch (e) {
              parser.abort();
              reject(e);
            }
          },
          complete: resolve,
          error: reject,
        });
      });
      bytesAnteriores += arquivo.size;
    }
  }

  if (linhasMantidas === 0) {
    throw new Error("Nenhuma linha correspondeu ao filtro selecionado.");
  }
  await Promise.all(promessasEscrita);
  return finalizarAnonimizacao(mapeador, linhasMantidas, linhasLidas);
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
        `<i class="ti ti-file-text"></i> <strong>${e.target.files[0].name}</strong> — ${(e.target.files[0].size / 1_048_576).toFixed(1)} MB`;
      document.getElementById("upload-desanon-info").classList.remove("oculto");
      verificarProntoDesanon();
    }
  });

  areaMapa.addEventListener("click", () => inputMapa.click());
  inputMapa.addEventListener("change", (e) => {
    if (e.target.files.length > 0) {
      estado.desanon.arquivoMapa = e.target.files[0];
      document.getElementById("upload-mapa-info").innerHTML = `<i class="ti ti-file-text"></i> <strong>${e.target.files[0].name}</strong>`;
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

function ehDocx(nome) {
  return /\.docx$/i.test(nome);
}
function ehPptx(nome) {
  return /\.pptx$/i.test(nome);
}

async function iniciarDesanonimizacaoDocumento(file) {
  const btn = document.getElementById("btn-iniciar-desanon");
  btn.disabled = true;
  document.getElementById("resultado-desanon").classList.add("oculto");
  const progressoWrap = document.getElementById("progresso-desanon");
  progressoWrap.classList.remove("oculto");
  const barra = document.querySelector("#progresso-desanon .preenchido");
  const texto = document.querySelector("#progresso-desanon .progresso-texto");
  barra.style.width = "5%";
  texto.textContent = "Lendo mapeamento...";

  try {
    const linhasMapeamento = await lerMapaCsv(estado.desanon.arquivoMapa);
    const arrayBuffer = await file.arrayBuffer();

    const aoProgredir = (fracao) => {
      barra.style.width = `${Math.round(fracao * 100)}%`;
      texto.textContent = `Processando documento... ${Math.round(fracao * 100)}%`;
    };

    let resultado, nomeSaida;
    if (ehDocx(file.name)) {
      resultado = await window.DesanonimizadorDocumentos.desanonimizarDocx(arrayBuffer, linhasMapeamento, aoProgredir);
      nomeSaida = "documento_desanonimizado.docx";
    } else {
      resultado = await window.DesanonimizadorDocumentos.desanonimizarPptx(arrayBuffer, linhasMapeamento, aoProgredir);
      nomeSaida = "apresentacao_desanonimizada.pptx";
    }

    const url = URL.createObjectURL(resultado.blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = nomeSaida;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 5000);

    document.getElementById("resultado-desanon").classList.remove("oculto");
    document.getElementById("rotulo-metrica-desanon").textContent = "Substituições feitas no texto";
    document.getElementById("metrica-linhas-desanon").textContent = resultado.relatorio.substituicoes.length.toLocaleString("pt-BR");

    const detalhes = document.getElementById("detalhes-desanon");
    detalhes.innerHTML = "";
    if (resultado.relatorio.naoEncontrados.size > 0) {
      const aviso = document.createElement("div");
      aviso.className = "callout callout-warning";
      aviso.style.marginBottom = "10px";
      aviso.innerHTML =
        `<i class="ti ti-alert-triangle"></i> ${resultado.relatorio.naoEncontrados.size} referência(s) citada(s) no texto não foram encontradas no mapeamento (ficaram como estavam): ` +
        Array.from(resultado.relatorio.naoEncontrados).join(", ");
      detalhes.appendChild(aviso);
    }
    if (resultado.numImagens > 0) {
      const info = document.createElement("div");
      info.className = "callout callout-info";
      info.style.marginBottom = "10px";
      info.innerHTML = `<i class="ti ti-info-circle"></i> O documento tem ${resultado.numImagens} imagem(ns) embutida(s) (ex: gráficos salvos como figura). Texto dentro de imagens não é substituído automaticamente.`;
      detalhes.appendChild(info);
    }
    if (typeof resultado.numGraficos === "number" && resultado.numGraficos > 0) {
      const info = document.createElement("div");
      info.className = "callout callout-info";
      info.style.marginBottom = "10px";
      info.innerHTML = `<i class="ti ti-chart-bar"></i> ${resultado.numGraficos} gráfico(s) nativo(s) encontrado(s), ${resultado.graficosAtualizados} atualizado(s) automaticamente.`;
      detalhes.appendChild(info);
    }
  } catch (err) {
    console.error(err);
    alert("Erro durante a desanonimização: " + err.message);
  } finally {
    btn.disabled = false;
    progressoWrap.classList.add("oculto");
  }
}

async function iniciarDesanonimizacao() {
  const file = estado.desanon.arquivo;

  if (ehDocx(file.name) || ehPptx(file.name)) {
    return iniciarDesanonimizacaoDocumento(file);
  }

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
  document.getElementById("rotulo-metrica-desanon").textContent = "Linhas processadas";
  document.getElementById("detalhes-desanon").innerHTML = "";
  document.getElementById("erro-analise-desanon").classList.add("oculto");
  const progressoWrap = document.getElementById("progresso-desanon");
  progressoWrap.classList.remove("oculto");
  if (escritor.modo === "blob") {
    document.getElementById("aviso-modo-blob-desanon").classList.remove("oculto");
  }

  try {
    const linhasMapeamento = await lerMapaCsv(estado.desanon.arquivoMapa);
    const mapaReverso = Core.construirMapaReverso(linhasMapeamento);

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

      const textoBruto = await lerAmostraTextoBruto(file, encoding);
      const diagnostico = diagnosticarLeitura(textoBruto, colunasArquivo, sep);
      if (diagnostico) throw erroAmigavel(diagnostico.mensagem);
    }

    const candidatos = Core.identificarColunas(colunasArquivo);
    const colunasParaRestaurar = {};
    for (const [col, idx] of Object.entries(candidatos)) {
      const nome = Core.COLUNAS_ALVO[idx].nome;
      if (mapaReverso[nome]) colunasParaRestaurar[col] = nome;
    }

    if (Object.keys(colunasParaRestaurar).length === 0) {
      throw erroAmigavel("Nenhuma coluna do arquivo bate com o mapeamento enviado.");
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
    const mensagem = err.amigavel
      ? err.message
      : `Não conseguimos processar esse arquivo. Verifique se o separador e a codificação em "Opções avançadas" correspondem ao arquivo, ou tente novamente.`;
    mostrarErroLeitura("erro-analise-desanon", mensagem);
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
