(() => {
  'use strict';
  const status = document.getElementById('status');
  const show = (message) => { status.textContent = message; };
  const detail = (error) => String(error && (error.stack || error.message) || error);
  let active = false;
  let listener = null;
  async function initialize() {
    if (active) { show('Already monitoring Company Master column D.'); return; }
    show('Diagnostic: JavaScript running.\nOffice global: ' + typeof Office + '\nExcel global: ' + typeof Excel + '\nWaiting for Office.onReady (max 12 seconds)...');
    if (typeof Office === 'undefined' || typeof Office.onReady !== 'function') {
      show('ERROR: Office.js did not expose Office.onReady. Check Office.js network access and add-in host.'); return;
    }
    let timeout;
    try {
      const info = await Promise.race([
        Office.onReady(),
        new Promise((_, reject) => { timeout = setTimeout(() => reject(new Error('Office.onReady timed out after 12 seconds. Excel host did not signal readiness.')), 12000); })
      ]);
      clearTimeout(timeout);
      show('Office.onReady completed.\nHost: ' + String(info && info.host) + '\nPlatform: ' + String(info && info.platform) + '\nExcel global: ' + typeof Excel);
      if (!info || info.host !== Office.HostType.Excel) {
        show('ERROR: Host is not Excel. Host: ' + String(info && info.host)); return;
      }
      if (typeof Excel === 'undefined') { show('ERROR: Excel JavaScript API is unavailable in this session.'); return; }
      await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItemOrNullObject('Company Master');
        sheet.load('isNullObject');
        await context.sync();
        if (sheet.isNullObject) throw new Error('Company Master worksheet not found.');
        listener = sheet.onChanged.add(async event => {
          try {
            const match = /^D(\d+)(?::D\d+)?$/i.exec(event.address.replace(/^.*!/, '').replace(/\$/g, ''));
            if (!match || Number(match[1]) < 7) return;
            const row = Number(match[1]);
            await Excel.run(async ctx => {
              const ws = ctx.workbook.worksheets.getItem('Company Master');
              const cell = ws.getRange('D' + row);
              cell.load('values');
              await ctx.sync();
              show('Detected edit at D' + row + '\nCompany: ' + String(cell.values[0][0] || '(blank)') + '\nNo cells were changed by this add-in.');
            });
          } catch (error) { show('Event error: ' + detail(error)); }
        });
        await context.sync();
      });
      active = true;
      show('MONITORING: Company Master, column D, row 7 onward.\nEdit a company name to test.\nRead-only: no worksheet changes.');
    } catch (error) {
      clearTimeout(timeout);
      show('DIAGNOSTIC ERROR:\n' + detail(error) + '\n\nIf Office.onReady times out, verify the add-in manifest SourceLocation, reload Excel, and test again.');
    }
  }
  document.getElementById('retry').addEventListener('click', initialize);
  initialize();
})();
