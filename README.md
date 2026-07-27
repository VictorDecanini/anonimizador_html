# Anonimizador de Dados — versão HTML (100% no navegador)

Anonimiza fornecedor, EAN, marca, SKU, canal, UF e nível 1. Processa **inteiramente no seu navegador** — nenhum arquivo é enviado a nenhum servidor. Isso elimina de vez o problema de limite de memória que tínhamos no Streamlit Cloud (que trava em ~1GB).

## Como rodar (localmente, sem instalar nada)

Basta abrir o `index.html` com duplo-clique. Funciona offline, sem instalar Python, sem terminal, sem `pip install`.

**Atenção:** por segurança dos navegadores, arraste a pasta inteira (`index.html`, `app.js`, `core.js`, `style.css`, `lib/`) — não mova só o `index.html` sozinho, ele precisa dos outros arquivos ao lado.

## Como ter um link compartilhável (GitHub Pages — grátis)

1. Cria um repositório no GitHub (pode reaproveitar um jeito parecido com o que já fizemos para o Streamlit)
2. Sobe esses 4 itens para a raiz do repositório: `index.html`, `app.js`, `core.js`, `style.css`, e a pasta `lib/` completa
3. No repositório, vai em **Settings → Pages**
4. Em "Source", escolhe a branch `main` e a pasta `/ (root)`
5. Salva. Em alguns minutos, o GitHub te dá um link do tipo `https://seu-usuario.github.io/nome-do-repositorio/`

Esse link é permanente, grátis, e qualquer pessoa do seu time pode acessar.

## Requisito de navegador — isso é importante

Para arquivos grandes (**recomendado: acima de 150MB**), use **Google Chrome ou Microsoft Edge atualizados**. Esses navegadores suportam gravar o resultado direto no disco durante o processamento, o que:
- Evita qualquer travamento por falta de memória
- É a única forma testada e validada para arquivos de 600MB+

Em outros navegadores (Firefox, Safari), a ferramenta ainda funciona, mas guarda o resultado na memória até o fim — para arquivos muito grandes, isso pode travar a aba. O próprio app avisa isso na tela quando detecta essa situação, e pede confirmação antes de seguir.

## O que foi testado (com arquivo real de ~600MB, 4,2 milhões de linhas)

- Detecção de colunas (incluindo casos como "Nome SKU", SKU numérico vs. textual, NIVEL1/NIVEL2, filtro por Data excluindo Periodo)
- Anonimização completa + filtro por valor
- Ciclo completo anonimizar → desanonimizar, com conferência linha a linha contra o arquivo original
- Suporte a Excel (.xlsx/.xls) e a CSVs em CP1252/Latin1 (acentuação)
- Tempo observado: ~25-30s para analisar e ~90-95s para processar um arquivo de 600MB no modo de streaming (Chrome/Edge)

## Desanonimizar Word e PowerPoint

A aba "Desanonimizar" também aceita **.docx** e **.pptx** diretamente — sobe o documento anonimizado + o mesmo CSV de referência, e a ferramenta substitui os códigos no texto, preservando a formatação (negrito, etc.) sempre que o trecho estiver dentro de uma única formatação. Reconhece 3 formas de código: completo (`Marca 3000000051`), abreviado (`Marca 197`) e em lista (`Marcas 213, 069 e 029`).

Em PowerPoint, também atualiza categorias e nomes de série de **gráficos nativos** (não os salvos como imagem).

**Limitações conhecidas:**
- **PDF não é suportado.** O formato guarda texto por posição de caractere, sem uma estrutura de parágrafo real como Word/PowerPoint — trocar um código por um texto de tamanho diferente arrisca desalinhar o documento inteiro. Por isso, optamos por não implementar em vez de arriscar um resultado quebrado.
- Números soltos sem o rótulo da categoria por perto (ex: só "029" sem "Marca" na frente) não são resolvidos automaticamente.
- Texto **dentro de imagens** (ex: um gráfico salvo como figura, print de tela) não é substituído — só texto real e gráficos nativos.

## Estrutura dos arquivos

- `index.html` — estrutura da página
- `style.css` — visual (mesma identidade executiva do app Streamlit: navy + azul Scanntech)
- `core.js` — lógica de anonimização de CSV/Excel (portada e testada a partir da versão Python)
- `docx-pptx.js` — lógica de desanonimização de Word/PowerPoint
- `app.js` — interface: upload, análise automática, filtros, barra de progresso, gravação em streaming
- `lib/` — bibliotecas de terceiros (PapaParse para CSV, SheetJS para Excel, JSZip para Word/PowerPoint), incluídas localmente para não depender de nenhum serviço externo no ar

## Limitações conhecidas (anonimização)

- Sem a opção de "selecionar todas" reaproveitar mapeamento anterior entre execuções (mesma decisão que você já tinha confirmado no Streamlit: cada anonimização começa do zero)
- O arquivo de mapeamento (referência) ainda é baixado do jeito tradicional (não em streaming) — não é um problema, porque esse arquivo é sempre muito menor que o principal
