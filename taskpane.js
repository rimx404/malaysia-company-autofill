(() => {
  'use strict';

  const WORKER = 'https://malaysia-company-research.rimli271.workers.dev';
  const SHEET = 'Company Master';
  const status = document.getElementById('status');
  const results = document.getElementById('results');
  const retry = document.getElementById('retry');

  let active = false;
  let initializing = false;
  let handlerResult = null;
  let requestNumber = 0;

  const show = message => {
    if (status) status.textContent = message;
  };

  const detail = error =>
    String((error && (error.message || error.stack)) || error);

  function clearResults() {
    results.replaceChildren();
  }

  function rowRangeFromAddress(address) {
    const local = String(address || '')
      .split('!').pop()
      .replace(/\$/g, '')
      .toUpperCase();

    const match = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(local);
    if (!match) return null;

    const columnNumber = letters =>
      [...letters].reduce(
        (n, c) => n * 26 + c.charCodeAt(0) - 64,
        0
      );

    const left = columnNumber(match[1]);
    const right = columnNumber(match[3] || match[1]);

    if (left > 4 || right < 4) return null;

    const start = Math.max(7, Number(match[2]));
    const end = Number(match[4] || match[2]);

    return end >= start ? { start, end } : null;
  }

  function addText(parent, tag, value, className) {
    const element = document.createElement(tag);
    element.textContent = value;
    if (className) element.className = className;
    parent.appendChild(element);
    return element;
  }

  function renderCandidates(company, row, candidates) {
    clearResults();

    addText(results, 'h2', 'Review research for D' + row);
    addText(results, 'p', 'Searched: ' + company);

    if (!candidates.length) {
      addText(results, 'p', 'No candidates returned. Nothing was changed.');
      return;
    }

    candidates.forEach(candidate => {
      const card = document.createElement('div');
      card.className = 'candidate';

      addText(card, 'h3', String(candidate.name || 'Unnamed candidate'));
      addText(card, 'p', String(candidate.description || 'No description available.'));
      addText(card, 'p',
        'Wikidata ID: ' + String(candidate.wikidata_id || 'Not provided'),
        'small'
      );

      if (
        typeof candidate.source === 'string' &&
        /^https:\/\/(www\.)?wikidata\.org\/entity\/Q\d+$/i.test(candidate.source)
      ) {
        const link = document.createElement('a');
        link.href = candidate.source;
        link.target = '_blank';
        link.rel = 'noopener noreferrer';
        link.textContent = 'View Wikidata source';
        card.appendChild(link);
      }

      addText(card, 'p',
        'Not verified as the correct Malaysian company.',
        'small'
      );

      const select = document.createElement('button');
      select.textContent = 'Select for review';
      select.addEventListener('click', () => {
        show(
          'SELECTED FOR REVIEW\n' +
          'Excel row: D' + row + '\n' +
          'Company searched: ' + company + '\n' +
          'Candidate: ' + String(candidate.name || '') + '\n' +
          'Wikidata ID: ' + String(candidate.wikidata_id || '') +
          '\n\nNo Excel cells changed. Further verification required.'
        );
      });
      card.appendChild(select);

      results.appendChild(card);
    });

    const reject = document.createElement('button');
    reject.textContent = 'Reject all candidates';
    reject.addEventListener('click', () => {
      clearResults();
      show('All candidates rejected for D' + row +
        '.\nNo Excel cells changed.');
    });
    results.appendChild(reject);
  }

  async function research(company, row) {
    const currentRequest = ++requestNumber;
    clearResults();

    show(
      'RESEARCHING\n' +
      'Cell: D' + row + '\n' +
      'Company: ' + company + '\n' +
      'Source: free Wikidata search\n' +
      'No Excel cells will be changed.'
    );

    try {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 20000);

      let response;
      try {
        // The Worker should support GET /research?company=...
        const url = WORKER + '/research?company=' +
          encodeURIComponent(company);

        response = await fetch(url, {
          method: 'GET',
          signal: controller.signal,
          headers: { Accept: 'application/json' }
        });
      } finally {
        clearTimeout(timeout);
      }

      if (!response.ok) {
        throw new Error('Research service returned HTTP ' + response.status);
      }

      const data = await response.json();
      if (currentRequest !== requestNumber) return;

      const candidates = Array.isArray(data.candidates)
        ? data.candidates
        : [];

      if (data.status && data.status !== 'candidates_for_review') {
        show(
          'RESEARCH RESPONSE\n' +
          'Status: ' + data.status + '\n' +
          'Company: ' + company + '\n' +
          'No Excel cells changed.'
        );
      } else {
        show(
          'RESEARCH COMPLETE\n' +
          'Company: ' + company + '\n' +
          'Candidates: ' + candidates.length + '\n' +
          'Review the results below.\n' +
          'No Excel cells changed.'
        );
      }

      renderCandidates(company, row, candidates);
    } catch (error) {
      if (currentRequest !== requestNumber) return;
      show(
        'RESEARCH ERROR\n' + detail(error) +
        '\n\nThe Worker may need a matching /research endpoint ' +
        'and CORS permission for GitHub Pages.\n' +
        'No Excel cells changed.'
      );
    }
  }

  async function onChanged(event) {
    try {
      const rows = rowRangeFromAddress(event.address);
      if (!rows) return;

      if (rows.end !== rows.start) {
        ++requestNumber;
        clearResults();
        show(
          'Detected multiple edited rows: ' +
          rows.start + '–' + rows.end +
          '.\nFor safety, research one company at a time.' +
          '\nNo Excel cells changed.'
        );
        return;
      }

      const row = rows.start;

      await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem(SHEET);
        const cell = sheet.getRange('D' + row);
        cell.load('values');
        await context.sync();

        const company = String(cell.values[0][0] || '').trim();

        if (!company) {
          ++requestNumber;
          clearResults();
          show('D' + row + ' is blank.\nNo Excel cells changed.');
          return;
        }

        await research(company, row);
      });
    } catch (error) {
      show('CHANGE EVENT ERROR\n' + detail(error));
    }
  }

  async function initialize() {
    if (active) {
      show(
        'MONITORING: Company Master column D, row 7 onward.' +
        '\nEdit a company name to begin research.'
      );
      return;
    }

    if (initializing) return;
    initializing = true;
    show('V5: Connecting to Excel...');

    let timer;

    try {
      if (typeof Office === 'undefined') {
        throw new Error('Office.js unavailable.');
      }

      const info = await Promise.race([
        Office.onReady(),
        new Promise((_, reject) => {
          timer = setTimeout(
            () => reject(new Error('Office connection timed out.')),
            15000
          );
        })
      ]);

      if (!info || info.host !== Office.HostType.Excel) {
        throw new Error('Excel host unavailable.');
      }

      await Excel.run(async context => {
        const sheet = context.workbook.worksheets
          .getItemOrNullObject(SHEET);

        sheet.load('isNullObject');
        await context.sync();

        if (sheet.isNullObject) {
          throw new Error('Company Master worksheet not found.');
        }

        handlerResult = sheet.onChanged.add(onChanged);
        await context.sync();
      });

      active = true;

      show(
        'V5 READY — REVIEW-FIRST MODE\n' +
        'Monitoring Company Master column D, row 7 onward.\n' +
        'Edit D7 to start free company research.\n' +
        'No automatic Excel changes.'
      );
    } catch (error) {
      show('V5 CONNECTION ERROR\n' + detail(error));
    } finally {
      clearTimeout(timer);
      initializing = false;
    }
  }

  if (retry) retry.addEventListener('click', initialize);
  initialize();
})();
