# Gerador de Convites v2 – SisConec.TA 2026

Versão funcional do `geradorconvites.html`, sem alterar o arquivo original.

## Conteúdo

- **index.html** – Interface igual à original, com IDs e scripts para funcionamento completo.
- **app.js** – Lógica: abas, carregar planilha (xlsx/csv), gerar PDFs em lote ou selecionados, convite rápido, log de gerações.

## Funcionalidades

1. **Lote – Planilha**
   - Escolher arquivo .xlsx ou .csv e clicar em **Carregar Planilha** (ou selecionar o arquivo no input).
   - Tabela preenchida com ID, Nome, Tratamento, Instituição, Categoria, Modelo (select) e Status.
   - Checkbox no cabeçalho para marcar/desmarcar todos.
   - **Gerar Todos (PDF)** – gera um PDF por linha e baixa um ZIP.
   - **Gerar Selecionados (PDF)** – só para as linhas com checkbox marcado.
   - **Baixar ZIP com todos os PDFs** – mesmo que “Gerar Todos”.

2. **Convite Rápido**
   - Formulário: Nome completo, Tratamento, Cargo, Instituição, Modelo.
   - Pré-visualização da saudação atualizada ao digitar.
   - **Gerar e Baixar PDF Agora** – gera um único PDF e faz o download.

3. **Log & Pendências**
   - Registro em memória (sessão) de cada geração: data/hora, ID/Nome, modelo, status e arquivo.

## Uso

- Abra **v2/index.html** no navegador (recomendado: servidor local, ex. `npx serve` na raiz do projeto ou em `v2`).
- O template de convite é carregado de **../templates/templateAutoridades.html** e a imagem de fundo de **../bg.jpg** (URL absoluta no deploy para funcionar no GitHub Pages).

## Dependências (CDN)

- xlsx, jsPDF, html2canvas, JSZip, FileSaver.js (incluídas no index.html).
