/**
 * Gerador de Convites
 * Fluxo: escolher template → preencher manualmente ou via .xlsx → gerar PDFs em .zip
 * Templates são arquivos .html com placeholders {{NOME}}, {{DATA}}, etc.
 */

(function () {
  'use strict';

  const TEMPLATES = {
    'templateAutoridades.html': 'Convite Autoridades'
  };

  // Fallback quando fetch falha (ex.: abrir index.html por file://)
  const EMBEDDED_TEMPLATES = {
    'templates/templateAutoridades.html': '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"/><meta name="viewport" content="width=device-width, initial-scale=1.0"/><title>Ofício – Convite Sisconec.TA 2026</title><style>*{box-sizing:border-box}body{margin:0;min-height:297mm;width:210mm;font-family:Georgia,serif;padding:25mm;color:#222;font-size:12pt;line-height:1.5;background:url(\'../bg.jpg\') center/cover no-repeat}.oficio-header{font-weight:bold;margin-bottom:1.5em}.oficio-data{margin-bottom:2em}.saudacao{margin-bottom:0.5em}.destinatario{margin-bottom:1.5em}.assunto{font-weight:bold;margin-bottom:1.5em}.corpo p{margin:0 0 1em 0;text-align:justify}.url-line{text-align:right;margin-top:2em}</style></head><body><div class="oficio"><p class="oficio-header">OFÍCIO Nº36/2026 – SIsLAB/Rede SisAssistiva</p><p class="oficio-data">Uberlândia, {{data em que o convite foi gerado}}</p><p class="saudacao">Prezado {{Tratamento}}.</p><p class="destinatario">{{Nome Completo}}</p><p class="destinatario">{{Cargo}} - {{Instituição}}</p><p class="assunto">Assunto: Convite para participação no Sisconec.TA 2026</p><div class="corpo"><p>É com satisfação que o SIsLAB – Laboratório Integrador da Rede SisAssistiva em articulação com o Ministério da Ciência, Tecnologia e Inovação (MCTI), por meio da Secretaria de Ciência e Tecnologia para o Desenvolvimento Social – SEDES, convida {{Tratamento}} para participar do Sisconec.TA 2026 – Evento Nacional de Inovação Tecnológica Assistiva, que acontecerá nos dias 20 e 21 de março de 2026 na Arena Sabiazinho, localizada em Uberlândia/MG.</p><p>O Sisconec.TA 2026 será um evento voltado à apresentação das inovações apoiadas pelo Edital FINEP 2022 – Tecnologia Assistiva, com foco na articulação de parcerias e na efetiva transferência das tecnologias desenvolvidas pela Rede para a sociedade. O encontro dará visibilidade aos projetos e promoverá diálogo com agências de fomento, como a FINEP, além de representantes da indústria e do poder público, visando ampliar o acesso da população às soluções geradas no âmbito da Rede SisAssistiva.</p><p class="url-line">As inscrições deverão ser realizadas por meio do hotsite oficial do evento: https://sisconec-ta.cintespbr.org/</p></div></div></body></html>'
  };

  let templateHtml = null;
  let placeholderNames = []; // campos esperados pelo template (extraídos do HTML)
  let dataRows = [];
  let currentDataSource = 'manual';
  let manualRowCount = 0;
  const AUTO_DATE_PLACEHOLDER = 'data em que o convite foi gerado';
  let xlsxAllRows = [];
  let xlsxSelectedIndices = new Set();

  const el = {
    templateSelect: document.getElementById('templateSelect'),
    templateStatus: document.getElementById('templateStatus'),
    templatePlaceholders: document.getElementById('templatePlaceholders'),
    dataSourceHint: document.getElementById('dataSourceHint'),
    manualFieldsContainer: document.getElementById('manualFieldsContainer'),
    manualActions: document.getElementById('manualActions'),
    addManualRow: document.getElementById('addManualRow'),
    xlsxFile: document.getElementById('xlsxFile'),
    xlsxPreview: document.getElementById('xlsxPreview'),
    xlsxPreviewContent: document.getElementById('xlsxPreviewContent'),
    xlsxRecordList: document.getElementById('xlsxRecordList'),
    xlsxSelectAll: document.getElementById('xlsxSelectAll'),
    btnGenerate: document.getElementById('btnGenerate'),
    generateStatus: document.getElementById('generateStatus'),
    renderContainer: document.getElementById('renderContainer'),
  };

  // --- Extrair placeholders do HTML do template: {{NOME}} -> ['NOME', ...]
  function extractPlaceholders(html) {
    const set = new Set();
    const re = /\{\{([^}]+)\}\}/g;
    let m;
    while ((m = re.exec(html)) !== null) set.add(m[1].trim());
    return Array.from(set);
  }

  // --- Carregar template ao selecionar
  el.templateSelect.addEventListener('change', function () {
    const file = this.value;
    if (!file) {
      templateHtml = null;
      placeholderNames = [];
      showStatus(el.templateStatus, '', '');
      el.templatePlaceholders.textContent = '';
      el.templatePlaceholders.classList.add('hide');
      el.dataSourceHint.classList.remove('hide');
      el.manualFieldsContainer.innerHTML = '<p class="template-placeholders">Selecione um template para exibir os campos.</p>';
      el.manualActions.classList.add('hide');
      dataRows = [];
      updateGenerateButton();
      return;
    }

    showStatus(el.templateStatus, 'Carregando template...', 'info');
    fetch(file)
      .then(function (r) {
        if (!r.ok) throw new Error('Arquivo não encontrado (' + r.status + ')');
        return r.text();
      })
      .then(function (html) {
        var names = extractPlaceholders(html);
        if (names.length === 0) {
          // Resposta não é o template (ex.: página de erro); usar cópia embutida
          var embedded = EMBEDDED_TEMPLATES[file];
          if (embedded) {
            html = embedded;
            names = extractPlaceholders(embedded);
          }
        }
        templateHtml = html;
        placeholderNames = names;
        showStatus(el.templateStatus, 'Template carregado. Campos: ' + placeholderNames.join(', '), 'success');
        el.templatePlaceholders.textContent = 'Campos esperados: ' + placeholderNames.join(', ');
        el.templatePlaceholders.classList.remove('hide');
        el.dataSourceHint.classList.add('hide');
        manualRowCount = 0;
        buildManualForm();
        updateGenerateButton();
      })
      .catch(function (err) {
        var embedded = EMBEDDED_TEMPLATES[file];
        if (embedded) {
          templateHtml = embedded;
          placeholderNames = extractPlaceholders(embedded);
          showStatus(el.templateStatus, 'Template carregado (cópia local). Campos: ' + placeholderNames.join(', '), 'success');
          el.templatePlaceholders.textContent = 'Campos esperados: ' + placeholderNames.join(', ');
          el.templatePlaceholders.classList.remove('hide');
          el.dataSourceHint.classList.add('hide');
          manualRowCount = 0;
          buildManualForm();
          updateGenerateButton();
        } else {
          templateHtml = null;
          placeholderNames = [];
          showStatus(el.templateStatus, 'Erro ao carregar template: ' + (err.message || err), 'error');
          el.manualFieldsContainer.innerHTML = '<p class="template-placeholders">Selecione um template para exibir os campos.</p>';
          el.manualActions.classList.add('hide');
          updateGenerateButton();
        }
      });
  });

  // --- Montar formulário manual: um bloco por convite, com um input por placeholder (exceto data automática)
  function getPlaceholdersForForm() {
    return placeholderNames.filter(function (n) { return n !== AUTO_DATE_PLACEHOLDER; });
  }

  function buildManualForm() {
    var namesForForm = getPlaceholdersForForm();
    if (!namesForForm.length) return;
    el.manualFieldsContainer.innerHTML = '';
    el.manualActions.classList.remove('hide');

    for (let r = 0; r < Math.max(1, manualRowCount); r++) {
      appendManualRow(r);
    }
    if (manualRowCount === 0) manualRowCount = 1;
  }

  function appendManualRow(index) {
    var namesForForm = getPlaceholdersForForm();
    const row = document.createElement('div');
    row.className = 'manual-row';
    row.setAttribute('data-row-index', index);
    const title = document.createElement('h4');
    title.textContent = 'Convite ' + (index + 1);
    row.appendChild(title);

    const grid = document.createElement('div');
    grid.className = 'field-grid';
    namesForForm.forEach(function (name) {
      const field = document.createElement('div');
      field.className = 'field';
      const lbl = document.createElement('label');
      lbl.textContent = name;
      lbl.setAttribute('for', 'manual_' + index + '_' + name);
      const input = document.createElement('input');
      input.type = 'text';
      input.id = 'manual_' + index + '_' + name;
      input.setAttribute('data-row', index);
      input.setAttribute('data-field', name);
      field.appendChild(lbl);
      field.appendChild(input);
      grid.appendChild(field);
    });
    row.appendChild(grid);

    const btnRemove = document.createElement('button');
    btnRemove.type = 'button';
    btnRemove.className = 'btn-remove-row';
    btnRemove.title = 'Remover este convite';
    btnRemove.textContent = '✕';
    row.appendChild(btnRemove);

    btnRemove.addEventListener('click', function () {
      const rows = el.manualFieldsContainer.querySelectorAll('.manual-row');
      if (rows.length <= 1) return;
      row.remove();
      reindexManualRows();
      updateGenerateButton();
    });

    el.manualFieldsContainer.appendChild(row);
    row.querySelector('input').addEventListener('input', updateGenerateButton);
  }

  function reindexManualRows() {
    const rows = el.manualFieldsContainer.querySelectorAll('.manual-row');
    rows.forEach(function (r, i) {
      r.setAttribute('data-row-index', i);
      r.querySelector('h4').textContent = 'Convite ' + (i + 1);
      r.querySelectorAll('input').forEach(function (inp) {
        inp.setAttribute('data-row', i);
        inp.id = 'manual_' + i + '_' + inp.getAttribute('data-field');
      });
    });
    manualRowCount = rows.length;
  }

  el.addManualRow.addEventListener('click', function () {
    const nextIndex = el.manualFieldsContainer.querySelectorAll('.manual-row').length;
    manualRowCount = nextIndex + 1;
    appendManualRow(nextIndex);
    updateGenerateButton();
  });

  el.manualFieldsContainer.addEventListener('input', function () {
    updateGenerateButton();
  });

  // --- Coletar dados do formulário manual
  function collectManualData() {
    dataRows = [];
    el.manualFieldsContainer.querySelectorAll('.manual-row').forEach(function (row) {
      const obj = {};
      let hasValue = false;
      row.querySelectorAll('input[data-field]').forEach(function (inp) {
        const key = inp.getAttribute('data-field');
        const val = (inp.value || '').trim();
        obj[key] = val;
        if (val) hasValue = true;
      });
      if (hasValue) dataRows.push(obj);
    });
    return dataRows;
  }

  // --- Tabs: Manual | XLSX
  document.querySelectorAll('#dataSourceTabs .tab').forEach(function (tab) {
    tab.addEventListener('click', function () {
      const t = this.getAttribute('data-tab');
      document.querySelectorAll('#dataSourceTabs .tab').forEach(function (x) {
        x.classList.toggle('active', x.getAttribute('data-tab') === t);
      });
      document.querySelectorAll('.tab-content').forEach(function (c) {
        if (c.id === 'tab-manual' || c.id === 'tab-xlsx') {
          c.classList.toggle('active', c.id === 'tab-' + t);
        }
      });
      currentDataSource = t;
      updateGenerateButton();
      if (t === 'manual') collectManualData();
    });
  });

  // --- Planilha XLSX
  el.xlsxFile.addEventListener('change', function () {
    const file = this.files[0];
    if (!file) {
      xlsxAllRows = [];
      xlsxSelectedIndices.clear();
      dataRows = [];
      el.xlsxPreview.style.display = 'none';
      if (el.xlsxRecordList) el.xlsxRecordList.innerHTML = '';
      updateGenerateButton();
      return;
    }
    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        if (json.length < 2) {
          xlsxAllRows = [];
          xlsxSelectedIndices.clear();
          dataRows = [];
          el.xlsxPreviewContent.textContent = 'Planilha vazia ou sem dados (mínimo: cabeçalho + 1 linha).';
          el.xlsxPreview.style.display = 'block';
          if (el.xlsxRecordList) el.xlsxRecordList.innerHTML = '';
        } else {
          const headers = json[0].map(function (h) { return String(h || '').trim() || 'Coluna'; });
          xlsxAllRows = json.slice(1).map(function (row) {
            const obj = {};
            headers.forEach(function (h, i) {
              obj[h] = row[i] != null ? String(row[i]).trim() : '';
            });
            return obj;
          });
          xlsxSelectedIndices.clear();
          for (var i = 0; i < xlsxAllRows.length; i++) xlsxSelectedIndices.add(i);
          el.xlsxPreviewContent.textContent = 'Colunas: ' + headers.join(', ') + '\n\nTotal de linhas: ' + xlsxAllRows.length;
          el.xlsxPreview.style.display = 'block';
          buildXlsxRecordList();
      updateDataRowsFromXlsxSelection();
    }
    updateGenerateButton();
    updateXlsxSelectAllLabel();
  } catch (err) {
        xlsxAllRows = [];
        xlsxSelectedIndices.clear();
        dataRows = [];
        el.xlsxPreviewContent.textContent = 'Erro: ' + (err.message || err);
        el.xlsxPreview.style.display = 'block';
        if (el.xlsxRecordList) el.xlsxRecordList.innerHTML = '';
        updateGenerateButton();
      }
    };
    reader.readAsArrayBuffer(file);
  });

  function updateDataRowsFromXlsxSelection() {
    if (currentDataSource !== 'xlsx') return;
    dataRows = xlsxAllRows.filter(function (_, i) { return xlsxSelectedIndices.has(i); });
  }

  function buildXlsxRecordList() {
    if (!el.xlsxRecordList) return;
    el.xlsxRecordList.innerHTML = '';
    xlsxAllRows.forEach(function (row, i) {
      var nome = getValueForPlaceholder(row, 'Nome Completo');
      if (!nome || nome.trim() === '') nome = 'Convite ' + (i + 1);
      var li = document.createElement('li');
      var label = document.createElement('label');
      label.className = 'xlsx-record-item';
      var cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = xlsxSelectedIndices.has(i);
      cb.setAttribute('data-xlsx-index', i);
      cb.addEventListener('change', function () {
        var idx = parseInt(this.getAttribute('data-xlsx-index'), 10);
        if (this.checked) xlsxSelectedIndices.add(idx);
        else xlsxSelectedIndices.delete(idx);
        updateDataRowsFromXlsxSelection();
        updateGenerateButton();
      });
      label.appendChild(cb);
      label.appendChild(document.createTextNode(' ' + nome));
      li.appendChild(label);
      el.xlsxRecordList.appendChild(li);
    });
    updateXlsxSelectAllLabel();
  }

  if (el.xlsxSelectAll) {
    el.xlsxSelectAll.addEventListener('click', function () {
      var allSelected = xlsxAllRows.length > 0 && xlsxSelectedIndices.size === xlsxAllRows.length;
      if (allSelected) {
        xlsxSelectedIndices.clear();
      } else {
        xlsxSelectedIndices.clear();
        for (var i = 0; i < xlsxAllRows.length; i++) xlsxSelectedIndices.add(i);
      }
      buildXlsxRecordList();
      updateDataRowsFromXlsxSelection();
      updateGenerateButton();
    });
  }

  function updateXlsxSelectAllLabel() {
    if (!el.xlsxSelectAll) return;
    var allSelected = xlsxAllRows.length > 0 && xlsxSelectedIndices.size === xlsxAllRows.length;
    el.xlsxSelectAll.textContent = allSelected ? 'Desmarcar todos' : 'Selecionar todos';
  }

  function updateGenerateButton() {
    if (currentDataSource === 'manual') collectManualData();
    else if (currentDataSource === 'xlsx') updateDataRowsFromXlsxSelection();
    const hasTemplate = !!templateHtml;
    const hasData = dataRows.length > 0;
    el.btnGenerate.disabled = !hasTemplate || !hasData;
    updateGenerateButtonLabel();
    if (currentDataSource === 'xlsx') updateXlsxSelectAllLabel();
  }

  function updateGenerateButtonLabel() {
    if (!el.btnGenerate) return;
    if (currentDataSource === 'xlsx' && dataRows.length > 0) {
      el.btnGenerate.textContent = 'Gerar selecionados (' + dataRows.length + ') e baixar .zip';
    } else {
      el.btnGenerate.textContent = 'Gerar convites (PDF) e baixar .zip';
    }
  }

  // --- Normalizar nome para comparação (maiúsculas, sem acentos)
  function normalizeKey(str) {
    if (str == null || str === '') return '';
    return String(str)
      .toUpperCase()
      .trim()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
  }

  function formatDate(d) {
    var day = ('0' + d.getDate()).slice(-2);
    var month = ('0' + (d.getMonth() + 1)).slice(-2);
    return day + '/' + month + '/' + d.getFullYear();
  }

  // --- Obter valor do row para um placeholder (aceita chave com nome diferente na planilha)
  function getValueForPlaceholder(row, placeholderName) {
    if (placeholderName === AUTO_DATE_PLACEHOLDER) {
      var v = row[placeholderName];
      if (v === undefined || v === null || String(v).trim() === '') return formatDate(new Date());
      return v;
    }
    var val = row[placeholderName];
    if (val !== undefined && val !== null) return val;
    var normPlaceholder = normalizeKey(placeholderName);
    if (!normPlaceholder) return '';
    for (var k in row) {
      if (Object.prototype.hasOwnProperty.call(row, k) && normalizeKey(k) === normPlaceholder)
        return row[k];
    }
    return '';
  }

  // --- Nome do arquivo PDF a partir do convidado (Nome Completo), com fallback
  function getFileNameForRow(row, index) {
    var nome = getValueForPlaceholder(row, 'Nome Completo');
    if (nome == null) nome = '';
    nome = String(nome).trim();
    if (nome) {
      nome = nome.replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim();
      if (nome.length > 80) nome = nome.substring(0, 80);
    }
    return nome || ('convite_' + (index + 1));
  }

  // --- Substituir placeholders: {{NOME}} -> valor (usa placeholders do template e normalização)
  function fillTemplate(html, row) {
    var out = html;
    // Primeiro: substituir por cada chave do row (para compatibilidade)
    Object.keys(row).forEach(function (key) {
      var placeholder = '{{' + key + '}}';
      out = out.split(placeholder).join(String(row[key] != null ? row[key] : ''));
    });
    // Depois: substituir placeholders que ainda restaram, buscando por nome normalizado
    var re = /\{\{([^}]+)\}\}/g;
    out = out.replace(re, function (match, name) {
      var n = name.trim();
      return String(getValueForPlaceholder(row, n));
    });
    return out;
  }

  // --- Gerar PDF a partir do HTML completo do template (usa iframe para manter estilos e background)
  function htmlToPdf(fullHtml) {
    return new Promise(function (resolve, reject) {
      const iframe = document.createElement('iframe');
      iframe.style.cssText = 'position:absolute;left:-9999px;width:210mm;height:297mm;';
      document.body.appendChild(iframe);

      const doc = iframe.contentDocument;
      doc.open();
      doc.write(fullHtml);
      doc.close();

      function capture() {
        const body = doc.body;
        if (!body) return reject(new Error('Iframe body not ready'));
        html2canvas(body, {
          scale: 2,
          useCORS: true,
          allowTaint: true,
          logging: false,
          width: body.scrollWidth,
          height: body.scrollHeight,
          windowWidth: body.scrollWidth,
          windowHeight: body.scrollHeight,
        })
          .then(function (canvas) {
            document.body.removeChild(iframe);
            const imgData = canvas.toDataURL('image/jpeg', 0.95);
            const pdf = new jspdf.jsPDF({
              orientation: canvas.width > canvas.height ? 'landscape' : 'portrait',
              unit: 'mm',
              format: 'a4',
            });
            const pageW = pdf.internal.pageSize.getWidth();
            const pageH = pdf.internal.pageSize.getHeight();
            const ratio = Math.min(pageW / canvas.width, pageH / canvas.height) * 0.95;
            const w = canvas.width * ratio;
            const h = canvas.height * ratio;
            pdf.addImage(imgData, 'JPEG', (pageW - w) / 2, (pageH - h) / 2, w, h);
            resolve(pdf.output('arraybuffer'));
          })
          .catch(function (err) {
            document.body.removeChild(iframe);
            reject(err);
          });
      }

      // Aguardar documento e recursos (ex.: bg.jpg) carregarem antes de capturar
      function waitAndCapture() {
        setTimeout(capture, 600);
      }
      if (doc.readyState === 'complete') {
        waitAndCapture();
      } else {
        iframe.contentWindow.onload = waitAndCapture;
      }
    });
  }

  // --- Botão Gerar
  el.btnGenerate.addEventListener('click', async function () {
    if (currentDataSource === 'manual') collectManualData();
    if (!templateHtml || dataRows.length === 0) return;

    el.btnGenerate.disabled = true;
    showStatus(el.generateStatus, 'Gerando ' + dataRows.length + ' convite(s)...', 'info');

    const zip = new JSZip();
    const folder = zip.folder('convites');
    const usedNames = {};

    try {
      for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
        const filledHtml = fillTemplate(templateHtml, row);
        const pdfBuffer = await htmlToPdf(filledHtml);
        var baseName = getFileNameForRow(row, i);
        var fileName = baseName + '.pdf';
        var count = 1;
        while (usedNames[fileName]) {
          count++;
          fileName = baseName + ' (' + count + ').pdf';
        }
        usedNames[fileName] = true;
        folder.file(fileName, pdfBuffer);
      }
      const zipBlob = await zip.generateAsync({ type: 'blob' });
      saveAs(zipBlob, 'convites.zip');
      showStatus(el.generateStatus, 'Download iniciado: convites.zip (' + dataRows.length + ' PDFs).', 'success');
    } catch (err) {
      showStatus(el.generateStatus, 'Erro ao gerar PDFs: ' + (err.message || err), 'error');
    } finally {
      el.btnGenerate.disabled = false;
      updateGenerateButton();
    }
  });

  function showStatus(elm, message, type) {
    if (!elm) return;
    elm.textContent = message;
    elm.className = 'status ' + (type || '');
    elm.style.display = message ? 'block' : 'none';
  }
})();
