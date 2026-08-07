/* Interface do Anonimizador — integra core.js com PapaParse/SheetJS. */

const Core = window.AnonimizadorCore;

const COLUNA_ORIGEM = "Arquivo_Origem";
const LIMITE_SEGURO_SEM_STREAMING = 150 * 1024 * 1024; // 150MB
const LIMITE_RECOMENDADO_TOTAL = 2 * 1024 * 1024 * 1024; // 2GB

const estado = {
  anon: { arquivos: [], colunasArquivo: null, colunasAlvo: null, colunasFiltro: null, valoresUnicos: null, encoding: null },
  desanon: { arquivo: null, arquivosMapa: [] },
};

// ---------------------------------------------------------------------------
// Utilidades gerais
// ---------------------------------------------------------------------------
function ehExcel(nomeArquivo) {
  return /\.(xlsx|xls)$/i.test(nomeArquivo);
}
function ehCsvOuTxt(nomeArquivo) {
  return /\.(csv|txt)$/i.test(nomeArquivo);
}

function obterSeparador(elementoId) {
  // O <option value="\t"> em HTML é literalmente os caracteres barra + t
  // (2 caracteres), não um tab de verdade (1 caractere, byte 9) -- sem essa
  // conversão, "Tab" nunca funcionava de fato como separador.
  const valor = document.getElementById(elementoId).value;
  return valor === "\\t" ? "\t" : valor;
}

// ---------------------------------------------------------------------------
// Nome do arquivo de saída (baseado no nome original, não mais genérico)
// ---------------------------------------------------------------------------
function nomeBaseSemExtensao(nomeArquivo) {
  return nomeArquivo.replace(/\.[^./\\]+$/, "");
}

function prefixoComumEntreNomes(nomes) {
  if (nomes.length === 0) return "";
  let prefixo = nomes[0];
  for (let i = 1; i < nomes.length && prefixo; i++) {
    const atual = nomes[i];
    let j = 0;
    while (j < prefixo.length && j < atual.length && prefixo[j] === atual[j]) j++;
    prefixo = prefixo.slice(0, j);
  }
  return prefixo.replace(/[\s_\-.(]+$/, "");
}

function nomeSaidaAnonimizado(arquivos) {
  const nomesBase = arquivos.map((a) => nomeBaseSemExtensao(a.name));
  if (nomesBase.length === 1) return `${nomesBase[0]}_anonimizado.csv`;
  let comum = prefixoComumEntreNomes(nomesBase);
  if (comum.length < 3) comum = `${arquivos.length}_arquivos`;
  return `${comum}_empilhado.csv`;
}

function nomeSaidaDesanonimizado(nomeOriginal, extensao) {
  return `${nomeBaseSemExtensao(nomeOriginal)}_desanonimizado.${extensao}`;
}

// ---------------------------------------------------------------------------
// Diretiva "sep=X" do Excel: escrita no arquivo de saída, garante que o
// Excel já abra com cada valor na célula certa, sem precisar de "Texto
// para colunas" manual -- funciona independente do idioma/config regional
// do Excel de quem for abrir. Também sabemos reconhecer e remover essa
// linha na leitura, senão nossa própria ferramenta trataria "sep=X" como
// se fosse a primeira linha de dados ao reprocessar um arquivo já
// anonimizado (ex: na hora de desanonimizar).
function linhaDiretivaSep(sep) {
  return `sep=${sep}\n`;
}

async function removerDiretivaSepSeExistir(file) {
  if (ehExcel(file.name)) return { arquivo: file, separadorDetectado: null };

  const amostraBuffer = await file.slice(0, 64).arrayBuffer();
  const bytes = new Uint8Array(amostraBuffer);
  // Blob.text()/TextDecoder removem o BOM automaticamente do texto
  // decodificado, mas ele continua fisicamente presente nos BYTES do
  // arquivo -- sem contar esses 3 bytes à parte, o corte abaixo (que usa
  // file.slice, que opera em bytes) ficaria desalinhado.
  const temBom = bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
  const offsetBom = temBom ? 3 : 0;

  const amostraTexto = new TextDecoder("utf-8").decode(amostraBuffer);
  const match = amostraTexto.match(/^sep=(.)\r?\n/);
  if (!match) return { arquivo: file, separadorDetectado: null };

  const arquivoLimpo = new File([file.slice(offsetBom + match[0].length)], file.name, { type: file.type });
  return { arquivo: arquivoLimpo, separadorDetectado: match[1] };
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
  const tamanhoAmostra = 700_000;
  const blobInicio = file.slice(0, tamanhoAmostra);
  const blobFim = file.size > tamanhoAmostra ? file.slice(Math.max(0, file.size - 300_000), file.size) : null;

  const buffers = [await blobInicio.arrayBuffer()];
  if (blobFim) buffers.push(await blobFim.arrayBuffer());

  // Detecção estatística de codificação (mesma família de biblioteca que
  // usamos na versão Python, chardet) -- muito mais confiável que só
  // testar se decodifica como UTF-8 estrito, principalmente pra separar
  // CP1252/Latin-1 de UTF-8 quando o acento é a única pista.
  try {
    const tamanhoTotal = buffers.reduce((s, b) => s + b.byteLength, 0);
    const combinado = new Uint8Array(tamanhoTotal);
    let offset = 0;
    for (const buf of buffers) {
      combinado.set(new Uint8Array(buf), offset);
      offset += buf.byteLength;
    }

    const deteccao = window.Chardet ? window.Chardet.detect(combinado) : null;
    if (deteccao) {
      const nome = deteccao.toLowerCase();
      if (nome.startsWith("utf-8") || nome === "ascii" || nome === "us-ascii") return "utf-8";
      if (nome.startsWith("iso-8859") || nome.startsWith("windows-125")) return "windows-1252";
      // Codificação exótica que não esperamos para bases BR -- cai no
      // teste estrito abaixo em vez de arriscar um nome desconhecido.
    }
  } catch (e) {
    console.error("Falha na detecção estatística de codificação, usando fallback:", e);
  }

  async function utf8Estrito(buffer) {
    try {
      new TextDecoder("utf-8", { fatal: true }).decode(buffer);
      return true;
    } catch {
      return false;
    }
  }
  const ok = (await Promise.all(buffers.map(utf8Estrito))).every(Boolean);
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
  return {
    getSelecionados: () => Array.from(selecionados),
    selecionar: (valor) => {
      if (!opcoes.includes(valor)) return;
      selecionados.add(valor);
      renderBotao();
      renderOpcoes(busca.value);
    },
  };
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

  // Trocar separador/codificação reanalisa o(s) arquivo(s) já selecionado(s)
  // automaticamente -- sem isso, o usuário precisava re-subir o arquivo pra
  // uma escolha diferente ter efeito.
  ["separador-anon", "encoding-anon"].forEach((id) => {
    document.getElementById(id).addEventListener("change", () => {
      if (estado.anon.arquivos.length > 0) analisarArquivosAnon();
    });
  });

  // Botão explícito, como reforço/garantia -- caso a atualização automática
  // acima não pareça ter efeito (ex: cache de uma versão antiga da página),
  // clicar aqui força a releitura de qualquer forma.
  document.getElementById("btn-atualizar-leitura-anon").addEventListener("click", () => {
    if (estado.anon.arquivos.length > 0) {
      analisarArquivosAnon();
    } else {
      alert("Suba um arquivo primeiro.");
    }
  });
}

async function adicionarArquivosAnon(novosArquivos) {
  // Evita duplicar o mesmo arquivo (mesmo nome + tamanho) se selecionado de novo.
  for (let novo of novosArquivos) {
    const jaExiste = estado.anon.arquivos.some((a) => a.name === novo.name && a.size === novo.size);
    if (jaExiste) continue;

    // Se o arquivo já tiver a diretiva "sep=X" (ex: um arquivo que a
    // própria ferramenta gerou antes), remove essa linha antes de
    // qualquer leitura, e já ajusta o separador selecionado de acordo.
    const { arquivo: limpo, separadorDetectado } = await removerDiretivaSepSeExistir(novo);
    novo = limpo;
    if (separadorDetectado) {
      const valorSelect = separadorDetectado === "\t" ? "\\t" : separadorDetectado;
      const select = document.getElementById("separador-anon");
      if (Array.from(select.options).some((o) => o.value === valorSelect)) select.value = valorSelect;
    }

    estado.anon.arquivos.push(novo);
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
    const sep = obterSeparador("separador-anon");
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

      // Confere se o resultado tem cara de separador ou codificação
      // errados -- se tiver, avisa exatamente o que trocar, mas SEM
      // bloquear o fluxo: o usuário ainda pode ajustar o mapeamento na
      // próxima etapa, ou trocar o separador/codificação acima (o que
      // reanalisa automaticamente).
      const textoBruto = await lerAmostraTextoBruto(primeiro, encoding);
      const diagnostico = diagnosticarLeitura(textoBruto, colunas, sep);
      if (diagnostico) {
        mostrarErroLeitura("erro-analise-anon", diagnostico.mensagem);
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

const NOME_EXIBICAO_CATEGORIA = {
  fornecedor: "Fornecedor", ean: "EAN", marca: "Marca", sku: "SKU",
  canal: "Canal", uf: "UF", nivel1: "Nível 1", nivel2: "Nível 2",
};

function renderizarConfigAnon() {
  const { colunasAlvo } = estado.anon;
  const bloco = document.getElementById("bloco-config-anon");

  if (Object.keys(colunasAlvo).length === 0) {
    document.getElementById("aviso-sem-colunas-anon").classList.remove("oculto");
  } else {
    document.getElementById("aviso-sem-colunas-anon").classList.add("oculto");
  }
  bloco.classList.remove("oculto");

  renderizarMapeamentoCategorias();
  renderizarColunasEFiltros();

  document.getElementById("btn-recalcular-mapeamento").onclick = recalcularAPartirDoMapeamento;
}

function renderizarMapeamentoCategorias() {
  const { colunasArquivo, colunasAlvo } = estado.anon;
  const colunasDisponiveis = colunasArquivo.filter((c) => c !== COLUNA_ORIGEM);

  // Inverte colunasAlvo (coluna -> indice) para indice -> [colunas], para
  // pre-selecionar a sugestao automatica em cada categoria.
  const colunasPorIndice = {};
  for (const [col, idx] of Object.entries(colunasAlvo)) {
    if (!colunasPorIndice[idx]) colunasPorIndice[idx] = [];
    colunasPorIndice[idx].push(col);
  }

  const grade = document.getElementById("grade-mapeamento-anon");
  grade.innerHTML = "";
  const mapeamentoMultiselects = {};

  for (const [idxStr, info] of Object.entries(Core.COLUNAS_ALVO)) {
    const idx = Number(idxStr);
    const item = document.createElement("div");
    item.className = "mapeamento-item";
    grade.appendChild(item);

    const ms = criarMultiselect(item, colunasDisponiveis, NOME_EXIBICAO_CATEGORIA[info.nome] || info.nome);
    const sugeridas = colunasPorIndice[idx] || [];
    for (const col of sugeridas) ms.selecionar(col);

    mapeamentoMultiselects[idx] = ms;
  }

  estado.anon._mapeamentoMultiselects = mapeamentoMultiselects;
}

async function recalcularAPartirDoMapeamento() {
  const btn = document.getElementById("btn-recalcular-mapeamento");
  const erroEl = document.getElementById("erro-mapeamento-anon");
  erroEl.classList.add("oculto");

  const novoColunasAlvo = {};
  const colunaJaUsadaEm = {};
  for (const [idxStr, ms] of Object.entries(estado.anon._mapeamentoMultiselects)) {
    const idx = Number(idxStr);
    for (const col of ms.getSelecionados()) {
      if (colunaJaUsadaEm[col] !== undefined && colunaJaUsadaEm[col] !== idx) {
        const nomeA = NOME_EXIBICAO_CATEGORIA[Core.COLUNAS_ALVO[colunaJaUsadaEm[col]].nome];
        const nomeB = NOME_EXIBICAO_CATEGORIA[Core.COLUNAS_ALVO[idx].nome];
        erroEl.innerHTML = `<i class="ti ti-alert-triangle"></i> A coluna "${col}" foi selecionada para ${nomeA} e para ${nomeB} ao mesmo tempo. Cada coluna só pode corresponder a uma variável.`;
        erroEl.classList.remove("oculto");
        return;
      }
      colunaJaUsadaEm[col] = idx;
      novoColunasAlvo[col] = idx;
    }
  }

  estado.anon.colunasAlvo = novoColunasAlvo;

  btn.disabled = true;
  const textoOriginal = btn.innerHTML;
  btn.innerHTML = `<span class="spinner"></span> Atualizando...`;

  try {
    const arquivos = estado.anon.arquivos;
    const sep = obterSeparador("separador-anon");
    const encoding = estado.anon.encoding;
    const amostraPrimeiro = ehExcel(arquivos[0].name)
      ? (await lerAmostraExcel(arquivos[0])).linhas
      : (await lerAmostraCsv(arquivos[0], sep, encoding)).linhas;

    const colunasFiltro = {};
    for (const [col, idx] of Object.entries(novoColunasAlvo)) {
      const nome = Core.COLUNAS_ALVO[idx].nome;
      if (!Core.COLUNAS_FILTRAVEIS.has(nome)) continue;
      if (Core.colunaEhNumerica(amostraPrimeiro.map((l) => l[col]))) continue;
      colunasFiltro[col] = nome;
    }
    Object.assign(colunasFiltro, identificarColunasFiltroExtraDisponiveis());

    const valoresUnicos = await coletarValoresUnicosMultiplos(arquivos, sep, encoding, colunasFiltro);
    valoresUnicos[COLUNA_ORIGEM] = arquivos.map((a) => a.name);
    colunasFiltro[COLUNA_ORIGEM] = "arquivo de origem";

    estado.anon.colunasFiltro = colunasFiltro;
    estado.anon.valoresUnicos = valoresUnicos;

    renderizarColunasEFiltros();
  } catch (err) {
    console.error(err);
    erroEl.innerHTML = `<i class="ti ti-alert-triangle"></i> Não foi possível atualizar: ${err.message}`;
    erroEl.classList.remove("oculto");
  } finally {
    btn.disabled = false;
    btn.innerHTML = textoOriginal;
  }
}

function identificarColunasFiltroExtraDisponiveis() {
  // Reaproveita a mesma deteccao de filtros extra (ex: Data) usada na
  // analise inicial, sem depender de uma categoria da tabela principal.
  const jaMapeadas = new Set(Object.keys(estado.anon.colunasAlvo));
  return Core.identificarColunasFiltroExtra(estado.anon.colunasArquivo, jaMapeadas);
}

function renderizarColunasEFiltros() {
  const { colunasArquivo, colunasAlvo, colunasFiltro, valoresUnicos } = estado.anon;

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
    lbl.innerHTML = idx !== undefined ? `${col} <i class="ti ti-lock locked-mark"></i>` : col;
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
      // BOM UTF-8 no início -- sem isso, o Excel (e outros programas) não
      // detecta a codificação certa e mostra acento errado, mesmo o
      // conteúdo estando corretamente em UTF-8 por dentro.
      await writable.write("\uFEFF");
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
    escritor = await criarEscritorSaida(nomeSaidaAnonimizado(estado.anon.arquivos));
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
  const sep = obterSeparador("separador-anon");
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
  const promessasEscrita = [escritor.escrever(linhaDiretivaSep(sep))];

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
  inputArquivo.addEventListener("change", async (e) => {
    if (e.target.files.length > 0) {
      const arquivoOriginal = e.target.files[0];
      const erroFormato = document.getElementById("erro-formato-desanon");

      if (!ehCsvOuTxt(arquivoOriginal.name) && !ehExcel(arquivoOriginal.name) && !ehDocx(arquivoOriginal.name) && !ehPptx(arquivoOriginal.name) && !ehHtml(arquivoOriginal.name)) {
        estado.desanon.arquivo = null;
        document.getElementById("upload-desanon-info").classList.add("oculto");
        erroFormato.innerHTML = `<i class="ti ti-alert-triangle"></i> Esse formato de arquivo não é suportado aqui. Formatos aceitos: CSV, TXT, Excel, Word, PowerPoint e HTML.`;
        erroFormato.classList.remove("oculto");
        verificarProntoDesanon();
        return;
      }
      erroFormato.classList.add("oculto");

      const { arquivo: limpo, separadorDetectado } = await removerDiretivaSepSeExistir(arquivoOriginal);
      estado.desanon.arquivo = limpo;
      if (separadorDetectado) {
        const valorSelect = separadorDetectado === "\t" ? "\\t" : separadorDetectado;
        const select = document.getElementById("separador-desanon");
        if (Array.from(select.options).some((o) => o.value === valorSelect)) select.value = valorSelect;
      }
      document.getElementById("upload-desanon-info").innerHTML =
        `<i class="ti ti-file-text"></i> <strong>${limpo.name}</strong> — ${(limpo.size / 1_048_576).toFixed(1)} MB`;
      document.getElementById("upload-desanon-info").classList.remove("oculto");
      verificarProntoDesanon();
    }
  });

  areaMapa.addEventListener("click", () => inputMapa.click());
  areaMapa.addEventListener("dragover", (e) => { e.preventDefault(); areaMapa.classList.add("arrastando"); });
  areaMapa.addEventListener("dragleave", () => areaMapa.classList.remove("arrastando"));
  areaMapa.addEventListener("drop", (e) => {
    e.preventDefault();
    areaMapa.classList.remove("arrastando");
    if (e.dataTransfer.files.length > 0) adicionarArquivosMapa(Array.from(e.dataTransfer.files));
  });
  inputMapa.addEventListener("change", (e) => {
    if (e.target.files.length > 0) adicionarArquivosMapa(Array.from(e.target.files));
    inputMapa.value = "";
  });

  document.getElementById("btn-iniciar-desanon").addEventListener("click", iniciarDesanonimizacao);
}

function adicionarArquivosMapa(novosArquivos) {
  for (const novo of novosArquivos) {
    const jaExiste = estado.desanon.arquivosMapa.some((a) => a.name === novo.name && a.size === novo.size);
    if (!jaExiste) estado.desanon.arquivosMapa.push(novo);
  }
  renderizarListaArquivosMapa();
  verificarProntoDesanon();
}

function removerArquivoMapa(indice) {
  estado.desanon.arquivosMapa.splice(indice, 1);
  renderizarListaArquivosMapa();
  verificarProntoDesanon();
}

function renderizarListaArquivosMapa() {
  const lista = document.getElementById("upload-mapa-info");
  const arquivos = estado.desanon.arquivosMapa;
  if (arquivos.length === 0) {
    lista.classList.add("oculto");
    lista.innerHTML = "";
    return;
  }
  lista.classList.remove("oculto");
  lista.innerHTML = "";
  arquivos.forEach((arquivo, i) => {
    const item = document.createElement("div");
    item.className = "item-arquivo";
    item.innerHTML =
      `<span class="nome"><i class="ti ti-file-text"></i> ${arquivo.name}</span>` +
      `<button class="remover" title="Remover">✕</button>`;
    item.querySelector(".remover").addEventListener("click", () => removerArquivoMapa(i));
    lista.appendChild(item);
  });
}

function verificarProntoDesanon() {
  const pronto = estado.desanon.arquivo && estado.desanon.arquivosMapa.length > 0;
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

async function lerMapasCsv(arquivos) {
  // Combina vários mapeamentos em um só -- necessário quando o time
  // comercial anonimiza arquivos separados (cada um com seu próprio
  // mapeamento) e depois junta os resultados numa análise conjunta: pra
  // desanonimizar essa análise conjunta, precisa dos mapeamentos de todos
  // os arquivos originais ao mesmo tempo.
  const todasLinhas = [];
  const vistos = new Map(); // "coluna|codigo" -> valor_original
  const conflitos = [];
  for (const arquivo of arquivos) {
    const linhas = await lerMapaCsv(arquivo);
    for (const linha of linhas) {
      const chave = `${linha.coluna}|${linha.codigo}`;
      if (vistos.has(chave) && vistos.get(chave) !== linha.valor_original) {
        conflitos.push(`"${linha.codigo}" (${linha.coluna}) aponta para "${vistos.get(chave)}" em um mapeamento e "${linha.valor_original}" em outro (${arquivo.name})`);
      }
      vistos.set(chave, linha.valor_original);
      todasLinhas.push(linha);
    }
  }
  return { linhas: todasLinhas, conflitos };
}

function ehDocx(nome) {
  return /\.docx$/i.test(nome);
}
function ehPptx(nome) {
  return /\.pptx$/i.test(nome);
}
function ehHtml(nome) {
  return /\.html?$/i.test(nome);
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
    const { linhas: linhasMapeamento, conflitos } = await lerMapasCsv(estado.desanon.arquivosMapa);

    const aoProgredir = (fracao) => {
      barra.style.width = `${Math.round(fracao * 100)}%`;
      texto.textContent = `Processando documento... ${Math.round(fracao * 100)}%`;
    };

    let resultado, nomeSaida, blob;
    if (ehHtml(file.name)) {
      texto.textContent = "Lendo arquivo HTML...";
      const encSelecionado = document.getElementById("encoding-desanon").value;
      const encoding = encSelecionado === "auto" ? await detectarEncoding(file) : encSelecionado;
      const buffer = await file.arrayBuffer();
      const textoHtml = new TextDecoder(encoding).decode(buffer);
      resultado = await window.DesanonimizadorDocumentos.desanonimizarHtml(textoHtml, linhasMapeamento);
      nomeSaida = nomeSaidaDesanonimizado(file.name, file.name.match(/\.htm$/i) ? "htm" : "html");
      blob = new Blob(["\uFEFF" + resultado.html], { type: "text/html;charset=utf-8" });
    } else {
      const arrayBuffer = await file.arrayBuffer();
      if (ehDocx(file.name)) {
        resultado = await window.DesanonimizadorDocumentos.desanonimizarDocx(arrayBuffer, linhasMapeamento, aoProgredir);
        nomeSaida = nomeSaidaDesanonimizado(file.name, "docx");
      } else {
        resultado = await window.DesanonimizadorDocumentos.desanonimizarPptx(arrayBuffer, linhasMapeamento, aoProgredir);
        nomeSaida = nomeSaidaDesanonimizado(file.name, "pptx");
      }
      blob = resultado.blob;
    }

    const url = URL.createObjectURL(blob);
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
    if (conflitos.length > 0) {
      const avisoConflito = document.createElement("div");
      avisoConflito.className = "callout callout-warning";
      avisoConflito.style.marginBottom = "10px";
      avisoConflito.innerHTML = `<i class="ti ti-alert-triangle"></i> ${conflitos.length} código(s) aparecem com valores diferentes entre os mapeamentos enviados (usado o último encontrado): ${conflitos.join("; ")}`;
      detalhes.appendChild(avisoConflito);
    }
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

  if (ehDocx(file.name) || ehPptx(file.name) || ehHtml(file.name)) {
    return iniciarDesanonimizacaoDocumento(file);
  }

  const btn = document.getElementById("btn-iniciar-desanon");

  let escritor;
  try {
    escritor = await criarEscritorSaida(nomeSaidaDesanonimizado(file.name, "csv"));
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
    const { linhas: linhasMapeamento, conflitos } = await lerMapasCsv(estado.desanon.arquivosMapa);
    if (conflitos.length > 0) {
      mostrarErroLeitura("erro-analise-desanon", `${conflitos.length} código(s) aparecem com valores diferentes entre os mapeamentos enviados (usado o último encontrado): ${conflitos.join("; ")}`);
      document.getElementById("erro-analise-desanon").classList.remove("callout-danger");
      document.getElementById("erro-analise-desanon").classList.add("callout-warning");
    }
    const mapaReverso = Core.construirMapaReverso(linhasMapeamento);

    const sep = obterSeparador("separador-desanon");
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
    const promessasEscrita = [escritor.escrever(linhaDiretivaSep(sep))];

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
