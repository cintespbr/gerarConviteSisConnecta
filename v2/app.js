/**
 * Gerador de Convites v2 – SisConec.TA 2026
 * Lógica completa: abas, planilha, geração PDF, convite rápido, log.
 */
(function () {
  'use strict';

  const AUTO_DATE_PLACEHOLDER = 'data em que o convite foi gerado';
  const isV2 = window.location.pathname.indexOf('/v2') !== -1;
  const assetPrefix = isV2 ? '../' : '';

  var EMBEDDED_TEMPLATE = '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"/><title>Ofício – Sisconec.TA 2026</title><style>*{box-sizing:border-box}body{margin:0;min-height:297mm;width:210mm;font-family:Georgia,serif;padding:25mm;color:#222;font-size:12pt;line-height:1.5}.oficio-header{font-weight:bold;margin-bottom:1.5em}.oficio-data{margin-bottom:2em}.saudacao{margin-bottom:0.5em}.destinatario{margin-bottom:1.5em}.assunto{font-weight:bold;margin-bottom:1.5em}.corpo p{margin:0 0 1em 0;text-align:justify}.url-line{text-align:right;margin-top:2em}</style></head><body><div class="oficio"><p class="oficio-header">OFÍCIO Nº36/2026 – SIsLAB/Rede SisAssistiva</p><p class="oficio-data">Uberlândia, {{data em que o convite foi gerado}}</p><p class="saudacao">Prezado {{Tratamento}}.</p><p class="destinatario">{{Nome Completo}} <br><i> {{Cargo}} - {{Instituição}}</i></p><p class="assunto">Assunto: Convite para participação no Sisconec.TA 2026</p><div class="corpo"><p>É com satisfação que o SIsLAB – Laboratório Integrador da Rede SisAssistiva em articulação com o Ministério da Ciência, Tecnologia e Inovação (MCTI), por meio da Secretaria de Ciência e Tecnologia para o Desenvolvimento Social – SEDES, convida {{Tratamento}} para participar do Sisconec.TA 2026 – Evento Nacional de Inovação Tecnológica Assistiva, que acontecerá nos dias 20 e 21 de março de 2026 na Arena Sabiazinho, localizada em Uberlândia/MG.</p><p>O Sisconec.TA 2026 será um evento voltado à apresentação das inovações apoiadas pelo Edital FINEP 2022 – Tecnologia Assistiva, com foco na articulação de parcerias e na efetiva transferência das tecnologias desenvolvidas pela Rede para a sociedade.</p><p class="url-line">As inscrições deverão ser realizadas por meio do hotsite oficial do evento: https://sisconec-ta.cintespbr.org/</p></div></div></body></html>';

  let templateHtml = null;
  let loteRows = [];
  let logEntries = [];

  const el = {
    tabs: document.querySelectorAll('.tab[data-tab]'),
    tabContents: document.querySelectorAll('.tab-content'),
    xlsxFile: document.getElementById('xlsxFile'),
    btnCarregarPlanilha: document.getElementById('btnCarregarPlanilha'),
    btnGerarTodos: document.getElementById('btnGerarTodos'),
    btnGerarSelecionados: document.getElementById('btnGerarSelecionados'),
    btnBaixarZip: document.getElementById('btnBaixarZip'),
    loteSelectAll: document.getElementById('loteSelectAll'),
    loteTableBody: document.getElementById('loteTableBody'),
    rapidoTemplateSelect: document.getElementById('rapidoTemplateSelect'),
    rapidoTemplateStatus: document.getElementById('rapidoTemplateStatus'),
    rapidoFieldsContainer: document.getElementById('rapidoFieldsContainer'),
    btnGerarRapido: document.getElementById('btnGerarRapido'),
    rapidoPreviewPlaceholder: document.getElementById('rapidoPreviewPlaceholder'),
    rapidoPreviewIframeWrap: document.getElementById('rapidoPreviewIframeWrap'),
    rapidoPreviewIframe: document.getElementById('rapidoPreviewIframe'),
    logTableBody: document.getElementById('logTableBody'),
  };

  let rapidoTemplateHtml = null;
  let rapidoPlaceholderNames = [];

  // --- Abas
  function showTab(tabId) {
    el.tabs.forEach(function (t) {
      t.classList.toggle('active', t.getAttribute('data-tab') === tabId);
    });
    el.tabContents.forEach(function (c) {
      c.classList.toggle('hidden', c.id !== tabId);
    });
  }

  el.tabs.forEach(function (tab) {
    tab.addEventListener('click', function () {
      showTab(tab.getAttribute('data-tab'));
    });
  });

  // --- Carregar template (Modelo 02 – Autoridades)
  function loadTemplate(callback) {
    if (templateHtml) {
      callback(null, templateHtml);
      return;
    }
    var url = assetPrefix + 'templates/templateAutoridades.html';
    fetch(url)
      .then(function (r) {
        if (!r.ok) throw new Error('Template não encontrado');
        return r.text();
      })
      .then(function (html) {
        templateHtml = html;
        callback(null, html);
      })
      .catch(function (err) {
        if (EMBEDDED_TEMPLATE) {
          templateHtml = EMBEDDED_TEMPLATE;
          callback(null, EMBEDDED_TEMPLATE);
        } else {
          callback(err);
        }
      });
  }

  // --- Base URL para imagens (GH Pages subpath)
  function getBaseUrl() {
    var href = window.location.href.replace(/#.*$/, '').replace(/\?.*$/, '');
    var base = href.indexOf('/') === -1 ? href : href.substring(0, href.lastIndexOf('/') + 1);
    if (isV2) base = base + '../';
    return base;
  }

  function resolveBackgroundUrl(html) {
    var base = getBaseUrl();
    var bgUrl = base + 'bg.jpg';
    return html.replace(/url\s*\(\s*['"]?(?:\.\.\/)?bg\.jpg['"]?\s*\)/gi, 'url("' + bgUrl + '")');
  }

  // --- Data formatada
  function formatDate(d) {
    var day = ('0' + d.getDate()).slice(-2);
    var month = ('0' + (d.getMonth() + 1)).slice(-2);
    return day + '/' + month + '/' + d.getFullYear();
  }

  var MESES_EXTENSO = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
  function formatDateLong(d) {
    if (!d) d = new Date();
    var dia = d.getDate();
    var mes = MESES_EXTENSO[d.getMonth()];
    var ano = d.getFullYear();
    return dia + ' de ' + mes + ' de ' + ano;
  }

  function formatDateTime(d) {
    return d.getFullYear() + '-' +
      ('0' + (d.getMonth() + 1)).slice(-2) + '-' +
      ('0' + d.getDate()).slice(-2) + ' ' +
      ('0' + d.getHours()).slice(-2) + ':' +
      ('0' + d.getMinutes()).slice(-2);
  }

  // --- Normalizar chave (planilha pode ter nomes diferentes)
  function normalizeKey(str) {
    if (str == null || str === '') return '';
    return String(str).toUpperCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }

  function getValue(row, key) {
    var val = row[key];
    if (val !== undefined && val !== null) return val;
    var n = normalizeKey(key);
    if (!n) return '';
    for (var k in row) {
      if (Object.prototype.hasOwnProperty.call(row, k) && normalizeKey(k) === n) return row[k];
    }
    return '';
  }

  // --- Row para o template: Nome Completo, Tratamento, Cargo, Instituição, data
  function rowToTemplateData(row) {
    var nome = String(getValue(row, 'Nome Completo') || getValue(row, 'Nome') || '').trim();
    var tratamento = String(getValue(row, 'Tratamento') || '').trim();
    var cargo = String(getValue(row, 'Cargo') || '').trim();
    var instituicao = String(getValue(row, 'Instituição') || '').trim();
    return {
      'Nome Completo': nome,
      'Tratamento': tratamento,
      'Cargo': cargo,
      'Instituição': instituicao,
      [AUTO_DATE_PLACEHOLDER]: formatDateLong(new Date())
    };
  }

  function fillTemplate(html, data) {
    var out = html;
    Object.keys(data).forEach(function (key) {
      var ph = '{{' + key + '}}';
      out = out.split(ph).join(String(data[key] != null ? data[key] : ''));
    });
    var re = /\{\{([^}]+)\}\}/g;
    out = out.replace(re, function (match, name) {
      var n = name.trim();
      if (n === AUTO_DATE_PLACEHOLDER) return formatDateLong(new Date());
      return String(getValue(data, n) || '');
    });
    return out;
  }

  function htmlToPdf(fullHtml) {
    fullHtml = resolveBackgroundUrl(fullHtml);
    return new Promise(function (resolve, reject) {
      var iframe = document.createElement('iframe');
      iframe.style.cssText = 'position:absolute;left:-9999px;width:210mm;height:297mm;';
      document.body.appendChild(iframe);
      var doc = iframe.contentDocument;
      doc.open();
      doc.write(fullHtml);
      doc.close();

      function capture() {
        var body = doc.body;
        if (!body) return reject(new Error('Iframe não pronto'));
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
            var imgData = canvas.toDataURL('image/jpeg', 0.95);
            var pdf = new jspdf.jsPDF({
              orientation: canvas.width > canvas.height ? 'landscape' : 'portrait',
              unit: 'mm',
              format: 'a4',
            });
            var pageW = pdf.internal.pageSize.getWidth();
            var pageH = pdf.internal.pageSize.getHeight();
            var ratio = Math.min(pageW / canvas.width, pageH / canvas.height) * 0.95;
            var w = canvas.width * ratio;
            var h = canvas.height * ratio;
            pdf.addImage(imgData, 'JPEG', (pageW - w) / 2, (pageH - h) / 2, w, h);
            resolve(pdf.output('arraybuffer'));
          })
          .catch(function (err) {
            document.body.removeChild(iframe);
            reject(err);
          });
      }
      setTimeout(capture, 600);
    });
  }

  function addLog(idNome, modelo, status, arquivo) {
    logEntries.unshift({
      dt: formatDateTime(new Date()),
      idNome: idNome,
      modelo: modelo,
      status: status,
      arquivo: arquivo || '—'
    });
    renderLog();
  }

  function renderLog() {
    if (!el.logTableBody) return;
    el.logTableBody.innerHTML = '';
    if (logEntries.length === 0) {
      var tr = document.createElement('tr');
      tr.innerHTML = '<td colspan="5" style="color:#7f8c8d;">Nenhuma geração nesta sessão.</td>';
      el.logTableBody.appendChild(tr);
      return;
    }
    logEntries.forEach(function (e) {
      var tr = document.createElement('tr');
      if (e.status.indexOf('Erro') !== -1) tr.style.background = '#fff3cd';
      tr.innerHTML =
        '<td>' + e.dt + '</td>' +
        '<td>' + escapeHtml(e.idNome) + '</td>' +
        '<td>' + escapeHtml(e.modelo) + '</td>' +
        '<td>' + escapeHtml(e.status) + '</td>' +
        '<td>' + escapeHtml(e.arquivo) + '</td>';
      el.logTableBody.appendChild(tr);
    });
  }

  function escapeHtml(s) {
    if (s == null) return '';
    var div = document.createElement('div');
    div.textContent = s;
    return div.innerHTML;
  }

  // --- Planilha: carregar e preencher tabela
  function parseFile(file, callback) {
    var name = (file.name || '').toLowerCase();
    if (name.endsWith('.csv')) {
      var reader = new FileReader();
      reader.onload = function () {
        try {
          var text = reader.result;
          var lines = text.split(/\r?\n/).filter(function (l) { return l.trim(); });
          if (lines.length < 2) return callback(new Error('CSV precisa de cabeçalho e ao menos uma linha'));
          var headers = lines[0].split(',').map(function (h) { return h.trim().replace(/^"|"$/g, ''); });
          var rows = lines.slice(1).map(function (line) {
            var obj = {};
            var vals = line.match(/("([^"]*)")|([^,]+)/g) || [];
            vals = vals.map(function (v) { return (v || '').trim().replace(/^"|"$/g, ''); });
            headers.forEach(function (h, i) { obj[h] = vals[i] != null ? vals[i] : ''; });
            return obj;
          });
          callback(null, rows);
        } catch (e) {
          callback(e);
        }
      };
      reader.readAsText(file, 'UTF-8');
      return;
    }
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: 'array' });
        var sheet = workbook.Sheets[workbook.SheetNames[0]];
        var json = XLSX.utils.sheet_to_json(sheet);
        if (!json.length) return callback(new Error('Planilha vazia'));
        callback(null, json);
      } catch (err) {
        callback(err);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function buildLoteTable(rows) {
    loteRows = rows.map(function (r, i) {
      return {
        id: String(i + 1).padStart(3, '0'),
        row: r,
        selected: true,
        modelo: '2',
        status: 'Não gerado'
      };
    });
    renderLoteTable();
  }

  function renderLoteTable() {
    if (!el.loteTableBody) return;
    el.loteTableBody.innerHTML = '';
    if (loteRows.length === 0) {
      var tr = document.createElement('tr');
      tr.innerHTML = '<td colspan="8" style="color:#7f8c8d;">Carregue uma planilha para exibir os convidados.</td>';
      el.loteTableBody.appendChild(tr);
      return;
    }
    var allChecked = loteRows.every(function (r) { return r.selected; });
    if (el.loteSelectAll) el.loteSelectAll.checked = allChecked;

    loteRows.forEach(function (item, idx) {
      var tr = document.createElement('tr');
      tr.setAttribute('data-idx', idx);
      var nome = getValue(item.row, 'Nome Completo') || getValue(item.row, 'Nome') || '';
      var tratamento = getValue(item.row, 'Tratamento') || '';
      var instituicao = getValue(item.row, 'Instituição') || '';
      var categoria = getValue(item.row, 'Categoria') || '';
      var cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = item.selected;
      cb.addEventListener('change', function () {
        item.selected = cb.checked;
        updateSelectAllState();
      });
      var sel = document.createElement('select');
      sel.innerHTML = '<option value="1">Modelo 01 – Padrão</option><option value="2" selected>Modelo 02 – Autoridades</option><option value="3">Modelo 03 – Rede</option>';
      sel.value = item.modelo;
      sel.addEventListener('change', function () { item.modelo = sel.value; });
      tr.innerHTML =
        '<td></td>' +
        '<td>' + escapeHtml(item.id) + '</td>' +
        '<td>' + escapeHtml(nome) + '</td>' +
        '<td>' + escapeHtml(tratamento) + '</td>' +
        '<td>' + escapeHtml(instituicao) + '</td>' +
        '<td>' + escapeHtml(categoria) + '</td>' +
        '<td></td>' +
        '<td class="status-pendente">' + escapeHtml(item.status) + '</td>';
      tr.querySelector('td').appendChild(cb);
      tr.querySelectorAll('td')[6].appendChild(sel);
      el.loteTableBody.appendChild(tr);
    });
  }

  function updateSelectAllState() {
    if (!el.loteSelectAll) return;
    var allChecked = loteRows.length > 0 && loteRows.every(function (r) { return r.selected; });
    el.loteSelectAll.checked = allChecked;
  }

  if (el.loteSelectAll) {
    el.loteSelectAll.addEventListener('change', function () {
      var checked = this.checked;
      loteRows.forEach(function (r) { r.selected = checked; });
      renderLoteTable();
    });
  }

  if (el.xlsxFile) {
    el.xlsxFile.addEventListener('change', function () {
      var file = this.files[0];
      if (!file) return;
      parseFile(file, function (err, rows) {
        if (err) {
          alert('Erro ao ler arquivo: ' + (err.message || err));
          return;
        }
        buildLoteTable(rows);
      });
    });
  }

  if (el.btnCarregarPlanilha) {
    el.btnCarregarPlanilha.addEventListener('click', function () {
      el.xlsxFile && el.xlsxFile.click();
    });
  }

  function getSelectedLoteRows() {
    return loteRows.filter(function (r) { return r.selected; });
  }

  function gerarPdfLote(onlySelected, done) {
    var rows = onlySelected ? getSelectedLoteRows() : loteRows;
    if (!rows.length) {
      alert(onlySelected ? 'Nenhum convite selecionado.' : 'Carregue uma planilha e tenha ao menos um convidado.');
      if (done) done();
      return;
    }
    loadTemplate(function (err, html) {
      if (err) {
        alert('Erro ao carregar template: ' + (err.message || err));
        if (done) done();
        return;
      }
      var zip = new JSZip();
      var folder = zip.folder('convites');
      var usedNames = {};
      var total = rows.length;
      var next = 0;

      function doOne() {
        if (next >= total) {
          zip.generateAsync({ type: 'blob' }).then(function (blob) {
            saveAs(blob, 'convites.zip');
            rows.forEach(function (item) {
              item.status = 'Gerado';
              var nome = getValue(item.row, 'Nome Completo') || getValue(item.row, 'Nome') || item.id;
              addLog(item.id + ' – ' + nome, 'Modelo ' + item.modelo, 'Gerado com sucesso', 'Convite_' + item.id + '.pdf');
            });
            renderLoteTable();
            if (done) done();
          });
          return;
        }
        var item = rows[next];
        var data = rowToTemplateData(item.row);
        var filled = fillTemplate(html, data);
        htmlToPdf(filled)
          .then(function (buf) {
            var nome = (getValue(item.row, 'Nome Completo') || getValue(item.row, 'Nome') || '').trim();
            var baseName = nome.replace(/[\\/:*?"<>|]/g, '_').substring(0, 80) || ('convite_' + item.id);
            var fileName = baseName + '.pdf';
            var c = 1;
            while (usedNames[fileName]) {
              fileName = baseName + ' (' + c + ').pdf';
              c++;
            }
            usedNames[fileName] = true;
            folder.file(fileName, buf);
            next++;
            doOne();
          })
          .catch(function (err) {
            addLog(item.id, 'Modelo ' + item.modelo, 'Erro: ' + (err.message || err), '—');
            item.status = 'Erro';
            next++;
            doOne();
          });
      }
      doOne();
    });
  }

  if (el.btnGerarTodos) {
    el.btnGerarTodos.addEventListener('click', function () {
      this.disabled = true;
      gerarPdfLote(false, function () { el.btnGerarTodos.disabled = false; });
    });
  }
  if (el.btnGerarSelecionados) {
    el.btnGerarSelecionados.addEventListener('click', function () {
      this.disabled = true;
      gerarPdfLote(true, function () { el.btnGerarSelecionados.disabled = false; });
    });
  }
  if (el.btnBaixarZip) {
    el.btnBaixarZip.addEventListener('click', function () {
      this.disabled = true;
      gerarPdfLote(false, function () { el.btnBaixarZip.disabled = false; });
    });
  }

  // --- Convite Rápido: template primeiro, depois campos dinâmicos
  function extractPlaceholders(html) {
    var set = {};
    var re = /\{\{([^}]+)\}\}/g;
    var m;
    while ((m = re.exec(html)) !== null) set[m[1].trim()] = true;
    return Object.keys(set);
  }

  function getRapidoPlaceholdersForForm() {
    return rapidoPlaceholderNames.filter(function (n) { return n !== AUTO_DATE_PLACEHOLDER; });
  }

  function showRapidoStatus(msg, type) {
    if (!el.rapidoTemplateStatus) return;
    el.rapidoTemplateStatus.textContent = msg;
    el.rapidoTemplateStatus.style.display = msg ? 'block' : 'none';
    el.rapidoTemplateStatus.style.color = type === 'error' ? '#e74c3c' : '#7f8c8d';
  }

  function buildRapidoForm() {
    if (!el.rapidoFieldsContainer) return;
    var names = getRapidoPlaceholdersForForm();
    if (!names.length) {
      el.rapidoFieldsContainer.innerHTML = '<p style="color:#7f8c8d;">Nenhum campo editável neste template (ou template não carregado).</p>';
      el.btnGerarRapido.disabled = true;
      return;
    }
    el.rapidoFieldsContainer.innerHTML = '';
    names.forEach(function (name) {
      var isTratamento = name === 'Tratamento';
      var div = document.createElement('div');
      div.style.marginBottom = '14px';
      var label = document.createElement('label');
      label.textContent = name + (name === 'Nome Completo' || name === 'Instituição' ? ' *' : '');
      label.style.display = 'block';
      label.style.marginBottom = '4px';
      label.style.fontSize = '0.875rem';
      label.style.color = '#7f8c8d';
      div.appendChild(label);
      var input;
      if (isTratamento) {
        input = document.createElement('select');
        input.innerHTML = '<option>Sr(a).</option><option>Prof.</option><option selected>Profa.</option><option>Dr(a).</option><option>Ministro(a)</option><option>Diretor(a)</option>';
      } else {
        input = document.createElement('input');
        input.type = 'text';
        input.placeholder = 'Ex: ' + (name === 'Nome Completo' ? 'Prof. João da Silva' : name === 'Instituição' ? 'FINEP' : name === 'Cargo' ? 'Coordenador de Inovação' : '');
      }
      input.setAttribute('data-field', name);
      input.style.width = '100%';
      input.style.maxWidth = '420px';
      input.style.padding = '10px';
      input.style.border = '1px solid #ccc';
      input.style.borderRadius = '5px';
      input.style.boxSizing = 'border-box';
      div.appendChild(input);
      el.rapidoFieldsContainer.appendChild(div);
    });
    el.btnGerarRapido.disabled = false;
    updateRapidoPreview();
  }

  function updateRapidoPreview() {
    if (!el.rapidoPreviewIframe || !el.rapidoPreviewIframeWrap || !el.rapidoPreviewPlaceholder) return;
    if (!rapidoTemplateHtml) {
      el.rapidoPreviewPlaceholder.style.display = 'block';
      el.rapidoPreviewIframeWrap.style.display = 'none';
      return;
    }
    el.rapidoPreviewPlaceholder.style.display = 'none';
    el.rapidoPreviewIframeWrap.style.display = 'block';
    var data = collectRapidoData();
    var filled = fillTemplate(rapidoTemplateHtml, data);
    filled = resolveBackgroundUrl(filled);
    var iframe = el.rapidoPreviewIframe;
    var doc = iframe.contentDocument || iframe.contentWindow.document;
    doc.open();
    doc.write(filled);
    doc.close();
  }

  function collectRapidoData() {
    var data = {};
    if (!el.rapidoFieldsContainer) return data;
    el.rapidoFieldsContainer.querySelectorAll('[data-field]').forEach(function (inp) {
      var key = inp.getAttribute('data-field');
      data[key] = (inp.value || '').trim();
    });
    data[AUTO_DATE_PLACEHOLDER] = formatDateLong(new Date());
    return data;
  }

  if (el.rapidoTemplateSelect) {
    el.rapidoTemplateSelect.addEventListener('change', function () {
      var path = this.value;
      rapidoTemplateHtml = null;
      rapidoPlaceholderNames = [];
      el.rapidoFieldsContainer.innerHTML = '<p style="color:#7f8c8d;">Selecione um template acima para exibir os campos.</p>';
      el.btnGerarRapido.disabled = true;
      showRapidoStatus('', '');
      if (!path) return;
      showRapidoStatus('Carregando template...', '');
      var url = assetPrefix + path;
      fetch(url)
        .then(function (r) {
          if (!r.ok) throw new Error('Template não encontrado');
          return r.text();
        })
        .then(function (html) {
          rapidoTemplateHtml = html;
          rapidoPlaceholderNames = extractPlaceholders(html);
          showRapidoStatus('Template carregado. Campos: ' + getRapidoPlaceholdersForForm().join(', '), '');
          buildRapidoForm();
        })
        .catch(function (err) {
          if (path.indexOf('templateAutoridades') !== -1 && EMBEDDED_TEMPLATE) {
            rapidoTemplateHtml = EMBEDDED_TEMPLATE;
            rapidoPlaceholderNames = extractPlaceholders(EMBEDDED_TEMPLATE);
            showRapidoStatus('Template carregado (cópia local).', '');
            buildRapidoForm();
          } else {
            showRapidoStatus('Erro: ' + (err.message || err), 'error');
            el.rapidoFieldsContainer.innerHTML = '<p style="color:#e74c3c;">Erro ao carregar template. Verifique se está usando um servidor local ou publique no GitHub Pages.</p>';
          }
        });
    });
  }

  if (el.rapidoFieldsContainer) {
    el.rapidoFieldsContainer.addEventListener('input', updateRapidoPreview);
    el.rapidoFieldsContainer.addEventListener('change', updateRapidoPreview);
  }

  if (el.btnGerarRapido) {
    el.btnGerarRapido.addEventListener('click', function () {
      if (!rapidoTemplateHtml) {
        alert('Selecione um template primeiro.');
        return;
      }
      var data = collectRapidoData();
      var nome = (data['Nome Completo'] || '').trim();
      if (!nome) {
        alert('Preencha o campo Nome Completo.');
        return;
      }
      this.disabled = true;
      var filled = fillTemplate(rapidoTemplateHtml, data);
      htmlToPdf(filled)
        .then(function (buf) {
          var blob = new Blob([buf], { type: 'application/pdf' });
          saveAs(blob, 'Convite_' + nome.replace(/[\\/:*?"<>|]/g, '_').substring(0, 60) + '.pdf');
          addLog(nome, 'Convite Rápido', 'Gerado com sucesso', 'Convite_' + nome.substring(0, 30) + '.pdf');
          el.btnGerarRapido.disabled = false;
        })
        .catch(function (err) {
          alert('Erro ao gerar PDF: ' + (err.message || err));
          addLog(nome || '—', 'Convite Rápido', 'Erro: ' + (err.message || err), '—');
          el.btnGerarRapido.disabled = false;
        });
    });
  }

  renderLog();
})();
