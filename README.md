# Gerador de Convites

Aplicação em HTML + CSS + JavaScript que gera convites em PDF a partir de templates HTML e de dados vindos de planilha (.xlsx) ou entrada manual. Tudo roda no navegador; não é necessário servidor.

## Estrutura do repositório

```
GerarConvites/
├── index.html          # Página principal
├── styles.css          # Estilos da interface
├── app.js              # Lógica do gerador
├── bg.jpg              # Imagem de fundo (usada pelos templates)
├── templates/          # Templates de convite (.html)
│   └── templateAutoridades.html
└── README.md
```

## Como usar

1. **Abra** `index.html` no navegador.
2. **Template**: no select, escolha o modelo (ex.: "Convite Autoridades"). Os campos esperados aparecem abaixo.
3. **Dados**:
   - **Preencher manualmente**: preencha os campos exibidos; use "+ Adicionar outro convite" para mais de um.
   - **Planilha .xlsx**: selecione um arquivo cuja primeira linha seja o cabeçalho com os nomes das colunas (podem ser em português; o sistema faz a correspondência).
4. Clique em **"Gerar convites (PDF) e baixar .zip"**. Será baixado um `.zip` com um PDF por convite; os arquivos são nomeados pelo **Nome Completo** do convidado quando informado.

## Adicionar um novo template

1. Crie um arquivo `.html` na pasta **`templates/`** (ex.: `templates/templateBasico.html`).
2. Use placeholders no formato **`{{NOME_DO_CAMPO}}`** (ex.: `{{Nome Completo}}`, `{{Data}}`).
3. A imagem de fundo pode ser referenciada como `url('bg.jpg')` (resolvida a partir da pasta raiz quando o PDF é gerado).
4. No **`index.html`**, adicione uma nova opção no `<select id="templateSelect">`:

   ```html
   <option value="templates/templateBasico.html">Convite Básico</option>
   ```

Ao selecionar o novo template, os campos do formulário manual e a validação da planilha passam a usar os placeholders desse arquivo.

## Hospedagem

Publique a pasta completa em qualquer hospedagem estática (GitHub Pages, Netlify, etc.). Dependências (xlsx, jsPDF, html2canvas, JSZip, FileSaver) são carregadas por CDN.
