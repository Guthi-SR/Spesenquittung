(() => {
  'use strict';

  const VERSION = '69';
  const SHEETJS_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
  let xlsxPromise = null;

  function setStatus(text, kind='') {
    const el = document.getElementById('sraDirectoryStatus');
    if (!el) return;
    el.textContent = text;
    el.className = 'status' + (kind ? ' ' + kind : '');
  }

  function n(v) {
    return String(v == null ? '' : v).replace(/\r/g,'').replace(/\u00a0/g,' ').replace(/[ \t]+/g,' ').trim();
  }

  function keyName(v) {
    return n(v).toLocaleLowerCase('de-DE').replace(/\s+/g,' ');
  }

  function loadXlsx() {
    if (window.XLSX && window.XLSX.utils) return Promise.resolve(window.XLSX);
    if (xlsxPromise) return xlsxPromise;

    xlsxPromise = new Promise((resolve, reject) => {
      const existing = document.querySelector('script[data-sra-xlsx]');
      if (existing) {
        existing.addEventListener('load', () => resolve(window.XLSX), {once:true});
        existing.addEventListener('error', reject, {once:true});
        return;
      }
      const s = document.createElement('script');
      s.src = SHEETJS_URL;
      s.async = true;
      s.dataset.sraXlsx = VERSION;
      s.onload = () => window.XLSX ? resolve(window.XLSX) : reject(new Error('XLSX nicht verfügbar'));
      s.onerror = () => reject(new Error('SheetJS konnte nicht geladen werden'));
      document.head.appendChild(s);
    });
    return xlsxPromise;
  }

  function headerKey(v) {
    return n(v).toLocaleLowerCase('de-DE')
      .replace(/[._\-\/]+/g,' ')
      .replace(/\s+/g,' ')
      .trim();
  }

  function findColumn(row, aliases) {
    if (!row || typeof row !== 'object') return '';
    const normalized = new Map(Object.keys(row).map(k => [headerKey(k), k]));
    for (const alias of aliases) {
      const hit = normalized.get(headerKey(alias));
      if (hit != null) return hit;
    }
    return '';
  }

  async function exportExcel() {
    try {
      const XLSX = await loadXlsx();
      const records = (typeof loadSraDirectory === 'function' ? loadSraDirectory() : [])
        .slice()
        .sort((a,b) => (a.name||'').localeCompare(b.name||'','de'));

      const rows = records.map(r => ({
        'Name': r.name || '',
        'Heimatadresse': r.address || '',
        'Telefon': r.phone || '',
        'E-Mail': r.email || ''
      }));

      const ws = XLSX.utils.json_to_sheet(rows, {
        header:['Name','Heimatadresse','Telefon','E-Mail'],
        skipHeader:false
      });
      ws['!cols'] = [{wch:28},{wch:42},{wch:22},{wch:34}];

      const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:D1');
      for (let r=1; r<=range.e.r; r++) {
        ['C','D'].forEach(col => {
          const cell = ws[col + (r+1)];
          if (cell && cell.v != null) {
            cell.t = 's';
            cell.v = String(cell.v);
          }
        });
      }

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'SRA-Verzeichnis');

      const infoWs = XLSX.utils.aoa_to_sheet([
        ['SRA-Verzeichnis SRG Main-Spessart'],
        ['Spaltenüberschriften bitte nicht löschen.'],
        ['Beim Import werden vorhandene Einträge über den Namen erkannt und aktualisiert.'],
        ['Leere Zellen überschreiben vorhandene Angaben nicht.']
      ]);
      infoWs['!cols'] = [{wch:90}];
      XLSX.utils.book_append_sheet(wb, infoWs, 'Hinweise');

      const d = new Date();
      const stamp = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
      XLSX.writeFile(wb, 'SRA-Verzeichnis_SRG-MSP_' + stamp + '.xlsx', {compression:true});

      setStatus(records.length
        ? records.length + ' SRA wurden als Excel-Datei exportiert.'
        : 'Leere Excel-Vorlage für das SRA-Verzeichnis wurde exportiert.', 'ok');
    } catch (err) {
      console.error(err);
      setStatus('Excel-Export konnte nicht gestartet werden. Bitte Internetverbindung prüfen und erneut versuchen.', 'warn');
    }
  }

  async function importExcel(file) {
    if (!file) return;
    try {
      const XLSX = await loadXlsx();
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, {type:'array', cellDates:false, raw:false});
      const sheetName = wb.SheetNames.includes('SRA-Verzeichnis') ? 'SRA-Verzeichnis' : wb.SheetNames[0];
      if (!sheetName) throw new Error('Keine Tabelle gefunden');

      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {
        defval:'',
        raw:false,
        blankrows:false
      });

      if (!rows.length) {
        setStatus('Die Excel-Datei enthält keine SRA-Einträge.', 'warn');
        return;
      }

      const first = rows[0] || {};
      const nameCol = findColumn(first, ['Name','Vor- und Zuname','Vorname Nachname','SRA','Schiedsrichter-Assistent']);
      const addressCol = findColumn(first, ['Heimatadresse','Adresse','Anschrift','Wohnadresse']);
      const phoneCol = findColumn(first, ['Telefon','Telefonnummer','Handy','Mobil','Mobilnummer']);
      const emailCol = findColumn(first, ['E-Mail','Email','E Mail','Mail','E-Mail-Adresse']);

      if (!nameCol) {
        setStatus('Excel-Import nicht möglich: Es wurde keine Spalte „Name“ gefunden.', 'warn');
        return;
      }

      const records = typeof loadSraDirectory === 'function' ? loadSraDirectory() : [];
      const byName = new Map(records.map((r,i) => [keyName(r.name), i]));
      const now = new Date().toISOString();
      let added=0, updated=0, skipped=0;

      for (const row of rows) {
        const name = n(row[nameCol]);
        if (!name) { skipped++; continue; }

        const incoming = {
          name,
          address: addressCol ? n(row[addressCol]) : '',
          phone: phoneCol ? n(String(row[phoneCol] || '')) : '',
          email: emailCol ? n(row[emailCol]) : ''
        };

        const k = keyName(name);
        const idx = byName.get(k);

        if (idx == null) {
          const item = {
            id:'sra_' + Date.now() + '_' + Math.random().toString(36).slice(2,8),
            name: incoming.name,
            address: incoming.address,
            phone: incoming.phone,
            email: incoming.email,
            createdAt: now,
            updatedAt: now
          };
          records.push(item);
          byName.set(k, records.length-1);
          added++;
        } else {
          const item = records[idx];
          let changed = false;
          ['name','address','phone','email'].forEach(field => {
            if (incoming[field] && incoming[field] !== item[field]) {
              item[field] = incoming[field];
              changed = true;
            }
          });
          if (changed) {
            item.updatedAt = now;
            updated++;
          }
        }
      }

      if (typeof writeSraDirectory === 'function') writeSraDirectory(records);
      else {
        localStorage.setItem('spesenSraDirectoryV1', JSON.stringify(records));
        if (typeof renderSraDirectory === 'function') renderSraDirectory();
      }

      setStatus(
        'Excel-Import abgeschlossen: ' + added + ' neu, ' + updated + ' aktualisiert' +
        (skipped ? ', ' + skipped + ' Zeile' + (skipped===1?'':'n') + ' ohne Namen übersprungen' : '') + '.',
        'ok'
      );
      const details = document.getElementById('sraDirectoryDetails');
      if (details) details.open = true;
    } catch (err) {
      console.error(err);
      setStatus('Excel-Datei konnte nicht eingelesen werden. Bitte eine gültige .xlsx- oder .xls-Datei verwenden.', 'warn');
    }
  }

  function init() {
    const details = document.getElementById('sraDirectoryDetails');
    if (!details || document.getElementById('exportSraExcelBtn')) return;

    const actions = details.querySelector('.sra-dir-toolbar .actions');
    const clearBtn = document.getElementById('clearSraDirectoryBtn');
    if (!actions) return;

    const exportBtn = document.createElement('button');
    exportBtn.id = 'exportSraExcelBtn';
    exportBtn.type = 'button';
    exportBtn.className = 'secondary';
    exportBtn.textContent = 'Excel exportieren';

    const importBtn = document.createElement('button');
    importBtn.id = 'importSraExcelBtn';
    importBtn.type = 'button';
    importBtn.className = 'light';
    importBtn.textContent = 'Excel importieren';

    const input = document.createElement('input');
    input.id = 'sraExcelFileInput';
    input.type = 'file';
    input.hidden = true;
    input.accept = '.xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel';

    if (clearBtn) {
      actions.insertBefore(exportBtn, clearBtn);
      actions.insertBefore(importBtn, clearBtn);
      actions.insertBefore(input, clearBtn);
    } else {
      actions.append(exportBtn, importBtn, input);
    }

    const hint = details.querySelector('.hint');
    if (hint && !hint.dataset.excelHint) {
      hint.dataset.excelHint = '1';
      hint.insertAdjacentHTML('beforeend',
        '<br><strong>Excel:</strong> Das komplette SRA-Verzeichnis kann exportiert, in Excel bearbeitet und wieder importiert werden. Beim Import werden vorhandene SRA anhand des Namens aktualisiert und neue ergänzt.'
      );
    }

    exportBtn.addEventListener('click', exportExcel);
    importBtn.addEventListener('click', () => input.click());
    input.addEventListener('change', () => {
      const file = input.files && input.files[0];
      importExcel(file).finally(() => { input.value = ''; });
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, {once:true});
  else init();
})();