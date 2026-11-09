(() => {
  'use strict';
  const status = document.getElementById('status');
  const retry = document.getElementById('retry');
  const show = message => { if (status) status.textContent = message; };
  const detail = e => String((e && (e.stack || e.message)) || e);
  let active = false;
  let handlerResult = null;
  let initializing = false;

  function rowRangeFromAddress(address) {
    const local = String(address || '').split('!').pop().replace(/\$/g, '').toUpperCase();
    const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(local);
    if (!m) return null;
    const columnNumber = letters => [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
    const left = columnNumber(m[1]);
    const right = columnNumber(m[3] || m[1]);
    if (left > 4 || right < 4) return null;
    const start = Math.max(7, Number(m[2]));
    const end = Number(m[4] || m[2]);
    if (end < start) return null;
    return { start, end };
  }

  async function onChanged(event) {
    try {
      const rows = rowRangeFromAddress(event.address);
      if (!rows) return;
      if (rows.end - rows.start > 99) {
        show('Detected large edit including column D, rows ' + rows.start + '–' + rows.end + '.\nSkipped reading more than 100 rows.\nNo worksheet changes.');
        return;
      }
      await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem('Company Master');
        const range = sheet.getRange('D' + rows.start + ':D' + rows.end);
        range.load('values');
        await context.sync();
        const entries = range.values.map((v, i) => 'D' + (rows.start + i) + ': ' + String(v[0] == null || v[0] === '' ? '(blank)' : v[0]));
        show('DETECTED COMPANY NAME EDIT\n' + entries.join('\n') + '\n\nRead-only: no cells changed.');
      });
    } catch (error) { show('CHANGE EVENT ERROR:\n' + detail(error)); }
  }

  async function initialize() {
    if (active) { show('MONITORING: Company Master, column D, row 7 onward.\nEdit a company name to test.'); return; }
    if (initializing) return;
    initializing = true;
    show('Diagnostics V4: initializing Office connection...');
    let timer;
    try {
      if (typeof Office === 'undefined' || typeof Office.onReady !== 'function') throw new Error('Office.onReady unavailable.');
      const info = await Promise.race([
        Office.onReady(),
        new Promise((_, reject) => { timer = setTimeout(() => reject(new Error('Office.onReady timed out after 15 seconds.')), 15000); })
      ]);
      clearTimeout(timer);
      if (!info || info.host !== Office.HostType.Excel) throw new Error('Not running inside Excel. Host: ' + String(info && info.host));
      if (typeof Excel === 'undefined') throw new Error('Excel JavaScript API unavailable.');
      show('Excel connected. Registering Company Master change listener...');
      await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItemOrNullObject('Company Master');
        sheet.load('isNullObject');
        await context.sync();
        if (sheet.isNullObject) throw new Error('Worksheet "Company Master" was not found.');
        handlerResult = sheet.onChanged.add(onChanged);
        await context.sync();
      });
      active = true;
      show('MONITORING: Company Master, column D, row 7 onward.\nChange D7 or D8 to test.\nRead-only: no worksheet changes.');
    } catch (error) {
      show('DIAGNOSTIC ERROR:\n' + detail(error));
    } finally {
      clearTimeout(timer);
      initializing = false;
    }
  }
  if (retry) retry.addEventListener('click', initialize);
  initialize();
})();
