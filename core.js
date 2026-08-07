/*
 * Núcleo de anonimização — Scanntech.
 * Porta fiel da lógica Python já validada (app.py do anonimizador Streamlit).
 * Roda 100% no navegador: sem servidor, sem limite de memória de hospedagem.
 */

// ---------------------------------------------------------------------------
// Configuração (idêntica ao app.py)
// ---------------------------------------------------------------------------
const BASE_INDICE = 1_000_000_000; // até 999.999.999 valores únicos por coluna

const COLUNAS_ALVO = {
  1: {
    nome: "fornecedor",
    aliases: [
      "fornecedor", "fabricante", "proveedor", "supplier", "manufacturer",
      "nome fornecedor", "razao social fornecedor", "razao social",
      "nome fabricante", "cod fornecedor", "codigo fornecedor",
    ],
  },
  2: {
    nome: "ean",
    aliases: [
      "ean", "codigo ean", "cod ean", "codigo de barras", "codigo barras",
      "barcode", "gtin", "nome ean",
    ],
  },
  3: {
    nome: "marca",
    aliases: ["marca", "brand", "nome marca"],
  },
  4: {
    nome: "sku",
    aliases: [
      "sku", "nome sku", "cod sku", "codigo sku", "descricao sku",
      "produto", "nome produto", "descricao produto", "descricao",
      "nome do produto", "item", "nome item", "nome do item",
      "nombre sku",
    ],
  },
  5: {
    nome: "canal",
    aliases: ["canal", "channel", "pdv canal", "tipo canal"],
  },
  6: {
    nome: "uf",
    aliases: ["uf", "estado", "state", "unidade federativa"],
  },
  7: {
    nome: "nivel1",
    aliases: ["nivel 1", "nivel1", "n1"],
  },
  8: {
    nome: "nivel2",
    aliases: ["nivel 2", "nivel2", "n2"],
  },
};

// Colunas oferecidas no filtro de valores (categóricas, baixa cardinalidade).
const COLUNAS_FILTRAVEIS = new Set(["fornecedor", "marca", "canal", "uf", "nivel1", "nivel2"]);

// Filtros que NUNCA são anonimizados (ex: Data) -- casamento só por nome.
const FILTROS_EXTRA = {
  data: ["data", "date", "dt venda", "data venda", "dt_venda", "data referencia"],
};

const IDX_EAN = 2;
const IDX_SKU = 4;

// ---------------------------------------------------------------------------
// Normalização de texto (equivalente a unicodedata + encode ascii ignore)
// ---------------------------------------------------------------------------
function normalizar(texto) {
  if (texto === null || texto === undefined) return "";
  return String(texto)
    .trim()
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[^\x00-\x7F]/g, ""); // remove tudo que não é ASCII (equivalente a encode('ascii','ignore'))
}

function normalizarChave(texto) {
  return normalizar(texto).replace(/[^a-z0-9]/g, "");
}

function padronizarCanal(valor) {
  const v = normalizar(valor).replace(/ /g, "_").replace(/-/g, "_");
  if (["1_a_4", "1_4", "1a4"].includes(v)) return "canal_1_4";
  if (["5_a_9", "5_9", "5a9"].includes(v)) return "canal_5_9";
  if (["10", "10_mais", "10_ou_mais", "10plus", "10+"].includes(v)) return "canal_10_mais";
  if (["atacarejo", "atacado", "cash_carry", "cash_and_carry"].includes(v)) return "canal_atacarejo";
  return v;
}

// ---------------------------------------------------------------------------
// Detecção de colunas por nome (exata ou substring de 4+ caracteres)
// ---------------------------------------------------------------------------
function identificarColunas(colunasArquivo) {
  const aliasesPorIndice = {};
  for (const [idx, info] of Object.entries(COLUNAS_ALVO)) {
    aliasesPorIndice[idx] = info.aliases.map(normalizarChave);
  }

  const todosAliases = [];
  for (const [idx, aliases] of Object.entries(aliasesPorIndice)) {
    for (const alias of aliases) {
      todosAliases.push([alias, Number(idx)]);
    }
  }
  todosAliases.sort((a, b) => b[0].length - a[0].length);

  const encontradas = {};
  for (const coluna of colunasArquivo) {
    const colNorm = normalizarChave(coluna);
    let indiceEncontrado = null;

    for (const [idx, aliases] of Object.entries(aliasesPorIndice)) {
      if (aliases.includes(colNorm)) {
        indiceEncontrado = Number(idx);
        break;
      }
    }

    if (indiceEncontrado === null) {
      for (const [alias, idx] of todosAliases) {
        if (alias.length >= 4 && colNorm.includes(alias)) {
          indiceEncontrado = idx;
          break;
        }
      }
    }

    if (indiceEncontrado !== null) {
      encontradas[coluna] = indiceEncontrado;
    }
  }

  return encontradas;
}

function identificarColunasFiltroExtra(colunasArquivo, colunasJaUsadas) {
  const encontradas = {};
  for (const [nomeFiltro, aliases] of Object.entries(FILTROS_EXTRA)) {
    const aliasesNorm = aliases.map(normalizarChave);
    for (const coluna of colunasArquivo) {
      if (colunasJaUsadas.has(coluna) || coluna in encontradas) continue;
      const colNorm = normalizarChave(coluna);
      const casou =
        aliasesNorm.includes(colNorm) ||
        aliasesNorm.some((a) => a.length >= 4 && colNorm.includes(a));
      if (casou) {
        encontradas[coluna] = nomeFiltro;
        break;
      }
    }
  }
  return encontradas;
}

// ---------------------------------------------------------------------------
// Detecção de valor numérico (aceita formato BR: 1.234,56)
// ---------------------------------------------------------------------------
function tentarNumeroBR(valor) {
  if (valor === null || valor === undefined) return null;
  const texto = String(valor).trim();
  if (texto === "") return null;

  let n = Number(texto);
  if (!isNaN(n)) return n;

  const textoBr = texto.replace(/\./g, "").replace(",", ".");
  n = Number(textoBr);
  return isNaN(n) ? null : n;
}

function colunaEhNumerica(valores, limiar = 0.9) {
  const naoVazios = valores.filter((v) => v !== null && v !== undefined && String(v).trim() !== "");
  if (naoVazios.length === 0) return false;
  let numericos = 0;
  for (const v of naoVazios) {
    if (tentarNumeroBR(v) !== null) numericos += 1;
  }
  return numericos / naoVazios.length >= limiar;
}

// ---------------------------------------------------------------------------
// Refinamento por conteúdo (a mesma regra de "primeira ocorrência" do app.py)
// ---------------------------------------------------------------------------
function refinarColunasPorValor(amostraLinhas, colunasCandidatas) {
  const candidatosPorIndice = {};
  for (const [col, idx] of Object.entries(colunasCandidatas)) {
    if (!candidatosPorIndice[idx]) candidatosPorIndice[idx] = [];
    candidatosPorIndice[idx].push(col);
  }

  const colunasAlvoFinal = {};

  for (const [idxStr, cols] of Object.entries(candidatosPorIndice)) {
    const idx = Number(idxStr);

    if (idx === IDX_EAN) {
      if (cols.length > 0) colunasAlvoFinal[cols[0]] = idx;
      continue;
    }

    if (idx === IDX_SKU) {
      // SKU pode ter uma versão numérica (às vezes é o próprio EAN) e
      // versões textuais (a descrição -- ex: NOMBRE_SKU, NOME_SKU, "Nome
      // SKU" podem até coexistir como colunas diferentes no mesmo
      // arquivo). Todas as variantes de texto são sensíveis e devem ser
      // anonimizadas; a numérica fica restrita à primeira ocorrência
      // (mesma lógica das demais categorias, para não pegar métrica
      // derivada por engano, ex: N1_DP_SKU).
      let primeiraNumerica = null;
      for (const col of cols) {
        const valores = amostraLinhas.map((l) => l[col]);
        if (colunaEhNumerica(valores)) {
          if (primeiraNumerica === null) primeiraNumerica = col;
        } else {
          colunasAlvoFinal[col] = idx;
        }
      }
      if (primeiraNumerica !== null) colunasAlvoFinal[primeiraNumerica] = idx;
      continue;
    }

    // Demais categorias: só a primeira ocorrência NÃO numérica.
    for (const col of cols) {
      const valores = amostraLinhas.map((l) => l[col]);
      if (!colunaEhNumerica(valores)) {
        colunasAlvoFinal[col] = idx;
        break;
      }
    }
  }

  return colunasAlvoFinal;
}

// ---------------------------------------------------------------------------
// Analisa o arquivo (amostra) para detectar colunas-alvo e filtráveis
// ---------------------------------------------------------------------------
function analisarAmostra(colunasArquivo, amostraLinhas) {
  const candidatos = identificarColunas(colunasArquivo);
  const colunasAlvo = refinarColunasPorValor(amostraLinhas, candidatos);

  const colunasFiltro = {};
  for (const [col, idx] of Object.entries(colunasAlvo)) {
    const nome = COLUNAS_ALVO[idx].nome;
    if (!COLUNAS_FILTRAVEIS.has(nome)) continue;
    const valores = amostraLinhas.map((l) => l[col]);
    if (colunaEhNumerica(valores)) continue; // nunca filtra coluna numérica
    colunasFiltro[col] = nome;
  }

  const colunasFiltroExtra = identificarColunasFiltroExtra(colunasArquivo, new Set(Object.keys(colunasAlvo)));
  Object.assign(colunasFiltro, colunasFiltroExtra);

  return { colunasAlvo, colunasFiltro };
}

// ---------------------------------------------------------------------------
// Anonimização (mapeamento incremental, mesmo esquema de código do app.py)
// ---------------------------------------------------------------------------
// Alfabeto sem caracteres ambíguos (sem 0/O, sem 1/I) -- ~32 símbolos.
// Código final: 1 dígito de categoria (1-8) + 11 caracteres aleatórios
// deste alfabeto = 12 caracteres, 32^11 (~3,7×10^16) combinações por
// categoria. Substituiu o esquema sequencial antigo (idx*BASE + contador)
// porque, quando o time comercial anonimiza vários arquivos separados e
// depois junta os resultados numa análise conjunta, códigos sequenciais
// SEMPRE colidem entre arquivos diferentes (ambos começam do 1) -- com
// código aleatório desse tamanho, colisão entre arquivos é praticamente
// impossível.
const ALFABETO_CODIGO = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ";
const TAMANHO_SUFIXO_ALEATORIO = 11;

function gerarSufixoAleatorio() {
  let s = "";
  for (let i = 0; i < TAMANHO_SUFIXO_ALEATORIO; i++) {
    s += ALFABETO_CODIGO[Math.floor(Math.random() * ALFABETO_CODIGO.length)];
  }
  return s;
}

class MapeadorAnonimizacao {
  constructor() {
    this.mapas = {}; // idx -> Map(chaveNormalizada -> codigo)
    this.codigosUsados = {}; // idx -> Set(codigo) -- evita colisão dentro do mesmo arquivo
    for (const idx of Object.keys(COLUNAS_ALVO)) {
      this.mapas[idx] = new Map();
      this.codigosUsados[idx] = new Set();
    }
  }

  anonimizarValor(valorOriginal, idx) {
    const vStr = valorOriginal === null || valorOriginal === undefined ? "" : String(valorOriginal).trim();
    if (vStr === "") return valorOriginal; // mantém vazio como estava

    const nomeColuna = COLUNAS_ALVO[idx].nome;
    const ehCanal = nomeColuna === "canal";
    const chave = ehCanal ? padronizarCanal(vStr) : normalizar(vStr);

    const mapa = this.mapas[idx];
    let codigo = mapa.get(chave);
    if (codigo === undefined) {
      const usados = this.codigosUsados[idx];
      do {
        codigo = `${idx}${gerarSufixoAleatorio()}`;
      } while (usados.has(codigo));
      usados.add(codigo);
      mapa.set(chave, codigo);
    }
    return codigo;
  }

  construirMapeamento() {
    // valor_original sempre em MAIÚSCULA na referência, para ficar legível
    // ao desanonimizar depois.
    const linhas = [];
    for (const [idxStr, mapa] of Object.entries(this.mapas)) {
      const idx = Number(idxStr);
      const nomeColuna = COLUNAS_ALVO[idx].nome;
      for (const [chave, codigo] of mapa.entries()) {
        linhas.push({ coluna: nomeColuna, valor_original: chave.toUpperCase(), codigo });
      }
    }
    return linhas;
  }
}

// ---------------------------------------------------------------------------
// Desanonimização
// ---------------------------------------------------------------------------
function construirMapaReverso(linhasMapeamento) {
  const mapaReverso = {}; // nomeColuna -> Map(codigoStr -> valorOriginal)
  for (const linha of linhasMapeamento) {
    const coluna = linha.coluna;
    if (!mapaReverso[coluna]) mapaReverso[coluna] = new Map();
    mapaReverso[coluna].set(String(linha.codigo), linha.valor_original);
  }
  return mapaReverso;
}

function tentarNumerico(valor) {
  // Converte para número real qualquer valor restaurado que "parece
  // número", preservando texto genuíno (ex: nome de marca).
  const n = tentarNumeroBR(valor);
  if (n === null) return valor;
  return Number.isInteger(n) ? n : n;
}

function desanonimizarValor(valorAtual, mapaColuna) {
  if (valorAtual === null || valorAtual === undefined) return valorAtual;
  const chave = String(valorAtual).trim();
  if (chave === "") return valorAtual;
  const original = mapaColuna.get(chave);
  if (original === undefined) return valorAtual;
  return tentarNumerico(original);
}

// ---------------------------------------------------------------------------
// Exports (Node para testes locais / navegador via <script> global)
// ---------------------------------------------------------------------------
const AnonimizadorCore = {
  BASE_INDICE,
  COLUNAS_ALVO,
  COLUNAS_FILTRAVEIS,
  FILTROS_EXTRA,
  normalizar,
  normalizarChave,
  padronizarCanal,
  identificarColunas,
  identificarColunasFiltroExtra,
  tentarNumeroBR,
  colunaEhNumerica,
  refinarColunasPorValor,
  analisarAmostra,
  MapeadorAnonimizacao,
  construirMapaReverso,
  tentarNumerico,
  desanonimizarValor,
};

if (typeof module !== "undefined" && module.exports) {
  module.exports = AnonimizadorCore;
}
if (typeof window !== "undefined") {
  window.AnonimizadorCore = AnonimizadorCore;
}
