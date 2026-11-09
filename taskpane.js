(function () {
  'use strict';
  const status = document.getElementById('status');
  function show(message) { status.textContent = message; }
  window.addEventListener('error', e => show('JavaScript error: ' + e.message));
  window.addEventListener('unhandledrejection', e => show('Promise error: ' + String(e.reason)));
  if (!window.Office) { show('Office.js unavailable. Check connection or open inside Excel.'); return; }
  show('Office.js loaded. Waiting for Excel...');
  const watchdog = setTimeout(() => show('Excel initialization timed out. Close and reopen the add-in; check whether Office.js is blocked.'), 15000);
  Office.onReady(async info => {
    clearTimeout(watchdog);
    if (info.host !== Office.HostType.Excel) { show('This page must run inside Excel.'); return; }
    try {
      await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem('Company Master');
        sheet.onChanged.add(async event => {
          try {
            if (event.changeType && event.changeType !== 'RangeEdited') return;
            const address = event.address.split('!').pop().replace(/\$/g, '');
            const match = /^D([0-9]+)$/i.exec(address);
            if (!match || Number(match[1]) < 7) return;
            await Excel.run(async ctx => {
              const cell = ctx.workbook.worksheets.getItem('Company Master').getRange(address);
              cell.load('values');
              await ctx.sync();
              show('Detected company in ' + address + ': ' + String(cell.values[0][0] || '(blank)'));
            });
          } catch (e) { show('Detection error: ' + (e.message || String(e))); }
        });
        await context.sync();
      });
      show('Monitoring Company Master column D (rows 7+). Edit one company name to test.');
    } catch (e) { show('Excel connection error: ' + (e.message || String(e))); }
  }).catch(e => { clearTimeout(watchdog); show('Office initialization error: ' + String(e)); });
})();
