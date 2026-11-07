/* Erweiterter Sammeldruck der SRG-MSP-Spesenquittung – 2026-11-07.
 * Die bisherigen Druckmodi und das lokale Archiv bleiben unverändert.
 * Keine Datenübertragung: Es werden nur vorhandene Browser-Archivdaten gelesen. */
(function () {
  'use strict';
  var $ = function (id) { return document.getElementById(id); };
  var ARCHIVE_KEY = 'srgMspSpesenReceiptArchiveV1';
  var MODES = {
    'a5':     { cls: 'a5', count: 1, page: 'A5 landscape' },
    'a4':     { cls: 'a4-2', count: 2, page: 'A4 portrait' },
    'a4-1':   { cls: 'a4-one', count: 1, page: 'A4 landscape' },
    'a4-4':   { cls: 'a4-4', count: 4, page: 'A4 portrait' },
    'a6':     { cls: 'a6', count: 1, page: 'A6 landscape' }
  };
  var format = $('printFormat');
  var primaryPrint = $('printBtn');
  var archiveBox = $('receiptArchive');
  if (!format || !primaryPrint || !$('receipt')) return;

  /* Die alten Auswahlwerte a5 und a4 bleiben für bestehende Abläufe erhalten. */
  var a5 = format.querySelector('option[value="a5"]');
  var a4 = format.querySelector('option[value="a4"]');
  if (a5) a5.textContent = 'A5 – 1 Quittung';
  if (a4) a4.textContent = 'A4 – 2 Quittungen (A5)';
  [['a4-1','A4 – 1 Quittung (quer)'],
   ['a4-4','A4 – 4 Quittungen (A6)'],
   ['a6','A6 – 1 Quittung']].forEach(function (pair) {
     if (!format.querySelector('option[value="' + pair[0] + '"]')) {
       var opt = document.createElement('option');
       opt.value = pair[0]; opt.textContent = pair[1];
       format.appendChild(opt);
     }
   });

  var settings = format.closest('.printformat');
  var copyLabel = document.createElement('label');
  copyLabel.className = 'batch-copy-option';
  var copyCheckbox = document.createElement('input');
  copyCheckbox.type = 'checkbox';
  copyCheckbox.id = 'printWithCopy';
  copyLabel.appendChild(copyCheckbox);
  copyLabel.appendChild(document.createTextNode(' Mit Kopie je Spiel'));
  if (settings) settings.appendChild(copyLabel);

  var archiveActions = document.querySelector('#archivePanel .archive-actions');
  var archivePrint = document.createElement('button');
  archivePrint.type = 'button';
  archivePrint.id = 'printSelectedArchiveBtn';
  archivePrint.className = 'secondary';
  archivePrint.disabled = true;
  archivePrint.textContent = 'Ausgewählte drucken (0)';
  if (archiveActions) archiveActions.appendChild(archivePrint);

  var pages = document.createElement('div');
  pages.id = 'batchPrintPages';
  pages.setAttribute('aria-hidden', 'true');
  document.body.appendChild(pages);

  function loadArchive() {
    try {
      var records = JSON.parse(localStorage.getItem(ARCHIVE_KEY) || '[]');
      return Array.isArray(records) ? records.filter(function (r) {
        return r && r.id && r.controls && r.format === 'srg-msp-spesenquittung';
      }) : [];
    } catch (err) { return []; }
  }

  function selectedIds() {
    if (!archiveBox) return [];
    return Array.from(archiveBox.querySelectorAll('input[data-archive-select]:checked'))
      .map(function (box) { return box.dataset.archiveSelect; });
  }

  function updateArchivePrintCount() {
    var n = selectedIds().length;
    archivePrint.disabled = n === 0;
    archivePrint.textContent = 'Ausgewählte drucken (' + n + ')';
  }

  if (archiveBox) {
    archiveBox.addEventListener('change', function () { queueMicrotask(updateArchivePrintCount); }, true);
    if (typeof MutationObserver !== 'undefined') {
      new MutationObserver(updateArchivePrintCount).observe(archiveBox, {childList:true, subtree:true});
    }
  }
  updateArchivePrintCount();

  function syncLiveInputs(source, clone) {
    var original = source.querySelectorAll('input, textarea, select');
    var mirrored = clone.querySelectorAll('input, textarea, select');
    original.forEach(function (field, i) {
      var target = mirrored[i];
      if (!target) return;
      if (field.type === 'checkbox' || field.type === 'radio') target.checked = field.checked;
      else if (field.tagName === 'SELECT') target.selectedIndex = field.selectedIndex;
      else target.value = field.value;
    });
  }

  function getControl(record, id) {
    var value = record && record.controls && record.controls[id];
    return value ? String(value.value == null ? '' : value.value) : '';
  }

  function prepareArchiveClone(clone, record) {
    clone.querySelectorAll('input[id], select[id], textarea[id]').forEach(function (field) {
      var saved = record.controls[field.id];
      if (!saved) return;
      if (saved.kind === 'checked') field.checked = !!saved.value;
      else field.value = saved.value == null ? '' : String(saved.value);
    });

    var summary = record.summary || {};
    var role = String(summary.archiveRole || 'Schiedsrichter');
    var title = clone.querySelector('#receiptTitle');
    if (title) title.textContent = 'Spesenquittung für ' + role;
    var feeLabel = clone.querySelector('#srExpenseLabel');
    if (feeLabel && role !== 'Schiedsrichter') {
      feeLabel.textContent = 'Aufwandsentschädigung / Spesen für ' + role;
    }

    /* Zeilen-Sichtbarkeit aus der jeweiligen gespeicherten Quittung ableiten
       statt versehentlich vom momentan geöffneten Spiel zu übernehmen. */
    var has = function (id) { return getControl(record, id).trim() !== ''; };
    var otherRow = clone.querySelector('#otherCostRow');
    if (otherRow) {
      otherRow.hidden = !has('other');
      otherRow.classList.toggle('print-hide-empty', !has('other'));
    }
    var srTravel = clone.querySelector('#srTravelRow');
    if (srTravel) srTravel.classList.toggle('print-hide-empty', !(has('srKm') || has('srTravel')));
    var team = clone.querySelector('#teamTravelRow');
    if (team) team.hidden = !(has('teamKm') || has('teamTravel'));
    var team2 = clone.querySelector('#teamTravelRow2');
    if (team2) team2.hidden = !(has('team2Km') || has('team2Travel'));
    ['sr','sra1','sra2','team','team2'].forEach(function (prefix) {
      var ratePart = clone.querySelector('#' + prefix + 'RatePart');
      if (ratePart) ratePart.hidden = !has(prefix + 'Km');
    });
  }

  function newReceipt(record, isCopy) {
    var source = $('receipt');
    var clone = source.cloneNode(true);
    clone.removeAttribute('id');
    clone.classList.add('batch-receipt');
    syncLiveInputs(source, clone);
    if (record) prepareArchiveClone(clone, record);
    if (isCopy) {
      var stamp = document.createElement('span');
      stamp.className = 'batch-copy-stamp';
      stamp.textContent = 'KOPIE';
      clone.appendChild(stamp);
    }
    return clone;
  }

  function cleanPrint() {
    document.body.classList.remove('batch-print', 'print-a4', 'print-a5');
    pages.innerHTML = '';
    var style = $('batchPageStyle');
    if (style) style.remove();
  }

  function printReceipts(records) {
    if (!records.length) return;
    var mode = MODES[format.value] || MODES.a5;
    var copies = copyCheckbox.checked;
    var jobs = [];
    records.forEach(function (record) {
      jobs.push({record:record, copy:false});
      if (copies) jobs.push({record:record, copy:true});
    });
    cleanPrint();
    var docFrag = document.createDocumentFragment();
    for (var i = 0; i < jobs.length; i += mode.count) {
      var sheetPage = document.createElement('div');
      sheetPage.className = 'batch-print-page batch-layout-' + mode.cls;
      jobs.slice(i, i + mode.count).forEach(function (job) {
        var slot = document.createElement('div');
        slot.className = 'batch-print-slot';
        slot.appendChild(newReceipt(job.record, job.copy));
        sheetPage.appendChild(slot);
      });
      docFrag.appendChild(sheetPage);
    }
    pages.appendChild(docFrag);

    var style = document.createElement('style');
    style.id = 'batchPageStyle';
    style.textContent = '@page{size:' + mode.page + ';margin:0}';
    document.head.appendChild(style);
    document.body.classList.add('print-a4', 'batch-print');

    /* Aufbau vor dem systemeigenen Druckdialog abschließen (iPhone + Browser).
       Kein automatischer Download; der Nutzer wählt Druck oder PDF selbst. */
    requestAnimationFrame(function () {
      requestAnimationFrame(function () { window.print(); });
    });
  }

  /* Capture verhindert, dass der bisherige Einzel-Druckhandler doppelt startet. */
  document.addEventListener('click', function (event) {
    var button = event.target.closest && event.target.closest('#printBtn, #printSelectedArchiveBtn');
    if (!button) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    if (button.id === 'printBtn') {
      if (typeof window.__srgMarkPrinted === 'function') window.__srgMarkPrinted();
      printReceipts([null]); // aktuelle, gegebenenfalls noch ungespeicherte Quittung
      return;
    }
    var ids = new Set(selectedIds());
    var selected = loadArchive().filter(function (r) { return ids.has(r.id); })
      .sort(function (a,b) {
        var da = String(a.summary && a.summary.date || '');
        var db = String(b.summary && b.summary.date || '');
        return da.localeCompare(db) || String(a.id).localeCompare(String(b.id));
      });
    if (!selected.length) {
      alert('Bitte zuerst gespeicherte Spiele im Archiv auswählen.');
      return;
    }
    printReceipts(selected);
  }, true);

  window.addEventListener('afterprint', cleanPrint);
})();
