/*
 * Desanonimizador de documentos Word (.docx) e PowerPoint (.pptx).
 * Porta fiel da lógica já validada em Python (python-docx/python-pptx):
 * reconhece códigos em 3 formatos -- completo, abreviado e em lista --
 * e faz a substituição preservando formatação sempre que possível.
 *
 * DOCX/PPTX são arquivos ZIP com XML dentro -- por isso usamos JSZip para
 * abrir, e o DOMParser/XMLSerializer nativos do navegador para editar o
 * XML sem quebrar a formatação do documento.
 */

const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
};

const INDICE_POR_ROTULO = {
  marca: 3,
  fabricante: 1,
  fornecedor: 1,
  sku: 4,
  ean: 2,
  canal: 5,
  uf: 6,
  nivel1: 7,
  nivel2: 8,
  "nivel 1": 7,
  "nivel 2": 8,
};
const BASE_INDICE_DOC = 1_000_000_000;

function normalizarChaveDoc(texto) {
  return String(texto).trim().toLowerCase();
}

// ---------------------------------------------------------------------------
// Núcleo de substituição de texto (comum a docx e pptx)
// ---------------------------------------------------------------------------
function construirProcessadorTexto(mapaCodigoParaValor) {
  const relatorio = { substituicoes: [], naoEncontrados: new Set() };

  const rotulos = Object.keys(INDICE_POR_ROTULO).sort((a, b) => b.length - a.length);
  const rotulosRegex = rotulos.map((r) => r.replace(/\s+/g, "\\s+")).join("|");

  // Rótulo (singular/plural) + lista de números (cobre "Marca 197" e
  // "Marcas 213, 069 e 029").
  const padraoLista = new RegExp(`\\b(${rotulosRegex})(s?)\\s+((?:\\d{1,10}(?:\\s*,\\s*|\\s+e\\s+)?)+)`, "gi");
  // Código completo aparecendo sozinho, sem rótulo na frente.
  const padraoCodigoBruto = /\b([1-6]\d{9})\b/g;

  function resolver(rotulo, plural, numeroStr) {
    const indice = INDICE_POR_ROTULO[rotulo.toLowerCase()];
    const numero = parseInt(numeroStr, 10);
    const codigo = numero >= BASE_INDICE_DOC ? numero : indice * BASE_INDICE_DOC + numero;
    const valor = mapaCodigoParaValor.get(String(codigo));
    const chaveOriginal = `${rotulo}${plural} ${numeroStr}`;
    if (valor === undefined) {
      relatorio.naoEncontrados.add(chaveOriginal);
      return numeroStr;
    }
    relatorio.substituicoes.push([chaveOriginal, valor]);
    return valor;
  }

  function processarTexto(texto) {
    if (!texto) return texto;
    let novo = texto.replace(padraoLista, (match, rotulo, plural, bloco) => {
      const novoBloco = bloco.replace(/\d{1,10}/g, (num) => resolver(rotulo, plural, num));
      return `${rotulo}${plural} ${novoBloco}`;
    });
    novo = novo.replace(padraoCodigoBruto, (match, codigo) => {
      const valor = mapaCodigoParaValor.get(codigo);
      if (valor === undefined) return match;
      relatorio.substituicoes.push([match, valor]);
      return valor;
    });
    return novo;
  }

  return { processarTexto, relatorio };
}

function construirMapaCodigoParaValor(linhasMapeamento) {
  const mapa = new Map();
  for (const linha of linhasMapeamento) {
    mapa.set(String(linha.codigo).trim(), String(linha.valor_original));
  }
  return mapa;
}

// ---------------------------------------------------------------------------
// Processa um único XML: agrupa elementos de texto por parágrafo, tenta
// primeiro elemento a elemento (preserva formatação), com fallback para o
// parágrafo inteiro quando o trecho está quebrado em vários elementos.
// ---------------------------------------------------------------------------
function processarXmlTexto(xmlString, processarTexto, tagTexto, tagParagrafo) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlString, "application/xml");

  const erro = doc.getElementsByTagName("parsererror")[0];
  if (erro) throw new Error("Falha ao interpretar XML do documento.");

  const paragrafos = doc.getElementsByTagNameNS(tagParagrafo.ns, tagParagrafo.local);

  for (let i = 0; i < paragrafos.length; i++) {
    const paragrafo = paragrafos[i];
    const elementosTexto = Array.from(paragrafo.getElementsByTagNameNS(tagTexto.ns, tagTexto.local));
    if (elementosTexto.length === 0) continue;

    // 1) tentativa elemento a elemento -- preserva formatação
    for (const el of elementosTexto) {
      const original = el.textContent;
      const novo = processarTexto(original);
      if (novo !== original) el.textContent = novo;
    }

    // 2) fallback: concatena tudo, reprocessa; se ainda mudar, é porque o
    // trecho atravessava mais de um elemento -- aplica no primeiro e
    // limpa os demais (perde formatação fina só nesse parágrafo).
    const textoAtual = elementosTexto.map((el) => el.textContent).join("");
    const textoNovo = processarTexto(textoAtual);
    if (textoNovo !== textoAtual) {
      elementosTexto[0].textContent = textoNovo;
      for (let j = 1; j < elementosTexto.length; j++) elementosTexto[j].textContent = "";
    }
  }

  return new XMLSerializer().serializeToString(doc);
}

// Substituição direta (sem agrupamento por parágrafo) -- usada para valores
// de gráfico (categorias/séries), que já são nós de texto autocontidos.
function processarXmlValoresDiretos(xmlString, processarTexto, tagValor) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlString, "application/xml");
  const elementos = doc.getElementsByTagNameNS(tagValor.ns, tagValor.local);
  for (let i = 0; i < elementos.length; i++) {
    const el = elementos[i];
    const original = el.textContent;
    const novo = processarTexto(original);
    if (novo !== original) el.textContent = novo;
  }
  return new XMLSerializer().serializeToString(doc);
}

// ---------------------------------------------------------------------------
// DOCX
// ---------------------------------------------------------------------------
async function desanonimizarDocx(arrayBuffer, linhasMapeamento, aoProgredir) {
  const mapa = construirMapaCodigoParaValor(linhasMapeamento);
  const { processarTexto, relatorio } = construirProcessadorTexto(mapa);

  const zip = await JSZip.loadAsync(arrayBuffer);
  const nomesArquivo = Object.keys(zip.files).filter(
    (n) => /^word\/(document|header\d*|footer\d*)\.xml$/.test(n)
  );

  let processados = 0;
  for (const nome of nomesArquivo) {
    const xml = await zip.file(nome).async("string");
    const novoXml = processarXmlTexto(
      xml, processarTexto,
      { ns: NS.w, local: "t" },
      { ns: NS.w, local: "p" }
    );
    zip.file(nome, novoXml);
    processados++;
    if (aoProgredir) aoProgredir(processados / nomesArquivo.length);
  }

  const numImagens = Object.keys(zip.files).filter((n) => /^word\/media\//.test(n)).length;

  const blob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  return { blob, relatorio, numImagens };
}

// ---------------------------------------------------------------------------
// PPTX
// ---------------------------------------------------------------------------
async function desanonimizarPptx(arrayBuffer, linhasMapeamento, aoProgredir) {
  const mapa = construirMapaCodigoParaValor(linhasMapeamento);
  const { processarTexto, relatorio } = construirProcessadorTexto(mapa);

  const zip = await JSZip.loadAsync(arrayBuffer);
  const nomesSlides = Object.keys(zip.files).filter((n) => /^ppt\/(slides|notesSlides)\/[^/]+\.xml$/.test(n));
  const nomesCharts = Object.keys(zip.files).filter((n) => /^ppt\/(charts|embeddings)\/.*chart\d*\.xml$/.test(n));

  const totalEtapas = nomesSlides.length + nomesCharts.length || 1;
  let processados = 0;

  for (const nome of nomesSlides) {
    const xml = await zip.file(nome).async("string");
    const novoXml = processarXmlTexto(
      xml, processarTexto,
      { ns: NS.a, local: "t" },
      { ns: NS.a, local: "p" }
    );
    zip.file(nome, novoXml);
    processados++;
    if (aoProgredir) aoProgredir(processados / totalEtapas);
  }

  let graficosAtualizados = 0;
  for (const nome of nomesCharts) {
    try {
      const xml = await zip.file(nome).async("string");
      // <c:v> guarda tanto o texto de categorias/séries quanto valores
      // numéricos de dados -- como só textos batem no padrão de código
      // anonimizado, é seguro reaproveitar o mesmo processador aqui.
      const novoXml = processarXmlValoresDiretos(xml, processarTexto, { ns: "http://schemas.openxmlformats.org/drawingml/2006/chart", local: "v" });
      if (novoXml !== xml) graficosAtualizados++;
      zip.file(nome, novoXml);
    } catch (e) {
      // best-effort: se algum chart XML vier malformado, ignora e segue
    }
    processados++;
    if (aoProgredir) aoProgredir(processados / totalEtapas);
  }

  const numImagens = Object.keys(zip.files).filter((n) => /^ppt\/media\//.test(n)).length;

  const blob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
  return { blob, relatorio, numImagens, numGraficos: nomesCharts.length, graficosAtualizados };
}

if (typeof window !== "undefined") {
  window.DesanonimizadorDocumentos = { desanonimizarDocx, desanonimizarPptx };
}
if (typeof module !== "undefined" && module.exports) {
  module.exports = { desanonimizarDocx, desanonimizarPptx, construirProcessadorTexto, construirMapaCodigoParaValor };
}
