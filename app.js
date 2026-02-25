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

  let templateHtml = null;
  let placeholderNames = []; // campos esperados pelo template (extraídos do HTML)
  let dataRows = [];
  let currentDataSource = 'manual';
  let manualRowCount = 0;

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
      .then(function (r) { return r.text(); })
      .then(function (html) {
        templateHtml = html;
        placeholderNames = extractPlaceholders(html);
        showStatus(el.templateStatus, 'Template carregado. Campos: ' + placeholderNames.join(', '), 'success');
        el.templatePlaceholders.textContent = 'Campos esperados: ' + placeholderNames.join(', ');
        el.templatePlaceholders.classList.remove('hide');
        el.dataSourceHint.classList.add('hide');
        manualRowCount = 0;
        buildManualForm();
        updateGenerateButton();
      })
      .catch(function (err) {
        templateHtml = null;
        placeholderNames = [];
        showStatus(el.templateStatus, 'Erro ao carregar template: ' + (err.message || err), 'error');
        el.manualFieldsContainer.innerHTML = '<p class="template-placeholders">Selecione um template para exibir os campos.</p>';
        el.manualActions.classList.add('hide');
        updateGenerateButton();
      });
  });

  // --- Montar formulário manual: um bloco por convite, com um input por placeholder
  function buildManualForm() {
    if (!placeholderNames.length) return;
    el.manualFieldsContainer.innerHTML = '';
    el.manualActions.classList.remove('hide');

    for (let r = 0; r < Math.max(1, manualRowCount); r++) {
      appendManualRow(r);
    }
    if (manualRowCount === 0) manualRowCount = 1;
  }

  function appendManualRow(index) {
    const row = document.createElement('div');
    row.className = 'manual-row';
    row.setAttribute('data-row-index', index);
    const title = document.createElement('h4');
    title.textContent = 'Convite ' + (index + 1);
    row.appendChild(title);

    const grid = document.createElement('div');
    grid.className = 'field-grid';
    placeholderNames.forEach(function (name) {
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
      dataRows = [];
      el.xlsxPreview.style.display = 'none';
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
          dataRows = [];
          el.xlsxPreviewContent.textContent = 'Planilha vazia ou sem dados (mínimo: cabeçalho + 1 linha).';
          el.xlsxPreview.style.display = 'block';
        } else {
          const headers = json[0].map(function (h) { return String(h || '').trim() || 'Coluna'; });
          dataRows = json.slice(1).map(function (row) {
            const obj = {};
            headers.forEach(function (h, i) {
              obj[h] = row[i] != null ? String(row[i]).trim() : '';
            });
            return obj;
          });
          el.xlsxPreviewContent.textContent = 'Colunas: ' + headers.join(', ') + '\n\nLinhas: ' + dataRows.length + '\n\nPrimeira linha: ' + JSON.stringify(dataRows[0] || {}, null, 2);
          el.xlsxPreview.style.display = 'block';
        }
        updateGenerateButton();
      } catch (err) {
        dataRows = [];
        el.xlsxPreviewContent.textContent = 'Erro: ' + (err.message || err);
        el.xlsxPreview.style.display = 'block';
        updateGenerateButton();
      }
    };
    reader.readAsArrayBuffer(file);
  });

  function updateGenerateButton() {
    if (currentDataSource === 'manual') collectManualData();
    const hasTemplate = !!templateHtml;
    const hasData = dataRows.length > 0;
    el.btnGenerate.disabled = !hasTemplate || !hasData;
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

  // --- Obter valor do row para um placeholder (aceita chave com nome diferente na planilha)
  function getValueForPlaceholder(row, placeholderName) {
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
  function getBaseUrl() {
    var href = window.location.href.replace(/#.*$/, '').replace(/\?.*$/, '');
    return href.indexOf('/') === -1 ? href : href.substring(0, href.lastIndexOf('/') + 1);
  }

  function resolveBackgroundUrl(html) {
    var base = getBaseUrl();
    var bgUrl = base + 'bg.jpg';
    return html.replace(/url\s*\(\s*['"]?(?:\.\.\/)?bg\.jpg['"]?\s*\)/gi, 'url("' + bgUrl + '")');
  }

  function htmlToPdf(fullHtml) {
    fullHtml = resolveBackgroundUrl(fullHtml);
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
