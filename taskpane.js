(() => {
  'use strict';

  const WORKER =
    'https://malaysia-company-research.rimli271.workers.dev';

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
    if (results) results.replaceChildren();
  }

  function rowRangeFromAddress(address) {
    const local = String(address || '')
      .split('!').pop()
      .replace(/\$/g, '')
      .toUpperCase();

    const match =
      /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(local);

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

    if (!results) {
      show('ERROR: Results panel not found.');
      return;
    }

    addText(results, 'h2', 'Review research for D' + row);
    addText(results, 'p', 'Searched: ' + company);

    if (!candidates.length) {
      addText(
        results,
        'p',
        'No candidates found. No Excel cells changed.'
      );
      return;
    }

    const cards = [];

    candidates.forEach(candidate => {
      const card = document.createElement('div');
      card.className = 'candidate';

      addText(
        card,
        'h3',
        String(candidate.name || 'Unnamed candidate')
      );

      addText(
        card,
        'p',
        String(
          candidate.description || 'No description available.'
        )
      );

      addText(
        card,
        'p',
        'Wikidata ID: ' +
          String(candidate.wikidata_id || 'Not provided'),
        'small'
      );

      if (
        typeof candidate.source === 'string' &&
        /^https:\/\/(www\.)?wikidata\.org\/(wiki\/|entity\/)Q\d+$/i
          .test(candidate.source)
      ) {
        const link = document.createElement('a');
        link.href = candidate.source;
        link.target = '_blank';
        link.rel = 'noopener noreferrer';
        link.textContent = 'View Wikidata source';
        card.appendChild(link);
      }

      addText(
        card,
        'p',
        'Not yet verified as the correct Malaysian company.',
        'small'
      );

      const select = document.createElement('button');
      select.type = 'button';
      select.textContent = 'Select for review';

      select.addEventListener('click', () => {
        // Clear previous selection indicators.
        cards.forEach(item => {
          item.card.style.border = '';
          item.card.style.background = '';
          item.button.disabled = false;
          item.button.textContent = 'Select for review';

          const oldNotice =
            item.card.querySelector('.selection-notice');

          if (oldNotice) oldNotice.remove();
        });

        // Highlight selected candidate.
        card.style.border = '2px solid #16865d';
        card.style.background = '#eefaf3';

        select.textContent = '✓ Selected for review';
        select.disabled = true;

        // Show confirmation immediately below the button.
        const notice = document.createElement('p');
        notice.className = 'selection-notice';
        notice.style.padding = '10px';
        notice.style.borderRadius = '8px';
        notice.style.background = '#dff5e8';
        notice.style.color = '#145a32';
        notice.style.fontWeight = '600';

        notice.textContent =
          'Selected: ' +
          String(candidate.name || '') +
          ' (Wikidata ' +
          String(candidate.wikidata_id || '') +
          '). Further verification required. ' +
          'No Excel cells changed.';

        card.appendChild(notice);

        show(
          'SELECTED FOR REVIEW\n' +
          'Excel row: D' + row + '\n' +
          'Company searched: ' + company + '\n' +
          'Candidate: ' +
          String(candidate.name || '') + '\n' +
          'Wikidata ID: ' +
          String(candidate.wikidata_id || '') +
          '\n\nSelection recorded in this panel only.' +
          '\nNo Excel cells changed.'
        );
      });

      card.appendChild(select);
      results.appendChild(card);

      cards.push({ card, button: select });
    });

    const reject = document.createElement('button');
    reject.type = 'button';
    reject.textContent = 'Reject all candidates';

    reject.addEventListener('click', () => {
      ++requestNumber;
      clearResults();

      show(
        'ALL CANDIDATES REJECTED\n' +
        'Cell: D' + row + '\n' +
        'No Excel cells changed.'
      );
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
      const timeout = setTimeout(
        () => controller.abort(),
        20000
      );

      let response;

      try {
        const url =
          WORKER +
          '/research?company=' +
          encodeURIComponent(company);

        response = await fetch(url, {
          method: 'GET',
          signal: controller.signal,
          headers: {
            Accept: 'application/json'
          }
        });
      } finally {
        clearTimeout(timeout);
      }

      if (!response.ok) {
        throw new Error(
          'Research service returned HTTP ' + response.status
        );
      }

      const data = await response.json();

      if (currentRequest !== requestNumber) return;

      const candidates = Array.isArray(data.candidates)
        ? data.candidates
        : [];

      if (data.status &&
          data.status !== 'candidates_for_review') {
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
          'Select the correct company below.\n' +
          'No Excel cells changed.'
        );
      }

      renderCandidates(company, row, candidates);

    } catch (error) {
      if (currentRequest !== requestNumber) return;

      clearResults();

      show(
        'RESEARCH ERROR\n' +
        detail(error) +
        '\n\nNo Excel cells changed.'
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
          'Multiple rows edited: ' +
          rows.start + '–' + rows.end +
          '.\nResearch one company at a time.' +
          '\nNo Excel cells changed.'
        );
        return;
      }

      const row = rows.start;
      let company = '';

      await Excel.run(async context => {
        const sheet =
          context.workbook.worksheets.getItem(SHEET);

        const cell = sheet.getRange('D' + row);
        cell.load('values');

        await context.sync();

        company = String(
          cell.values[0][0] == null
            ? ''
            : cell.values[0][0]
        ).trim();
      });

      if (!company) {
        ++requestNumber;
        clearResults();

        show(
          'D' + row +
          ' is blank.\nNo Excel cells changed.'
        );
        return;
      }

      await research(company, row);

    } catch (error) {
      show(
        'CHANGE EVENT ERROR\n' + detail(error)
      );
    }
  }

  async function initialize() {
    if (active) {
      show(
        'MONITORING: Company Master column D, ' +
        'row 7 onward.\nEdit a company name to begin research.'
      );
      return;
    }

    if (initializing) return;

    initializing = true;
    show('V6: Connecting to Excel...');

    let timer;

    try {
      if (
        typeof Office === 'undefined' ||
        typeof Office.onReady !== 'function'
      ) {
        throw new Error('Office.js unavailable.');
      }

      const info = await Promise.race([
        Office.onReady(),
        new Promise((_, reject) => {
          timer = setTimeout(
            () => reject(
              new Error('Office connection timed out.')
            ),
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
          throw new Error(
            'Company Master worksheet not found.'
          );
        }

        handlerResult = sheet.onChanged.add(onChanged);
        await context.sync();
      });

      active = true;

      show(
        'V6 READY — REVIEW-FIRST MODE\n' +
        'Monitoring Company Master column D, row 7 onward.\n' +
        'Edit a company name to begin research.\n' +
        'Select a candidate to see visible confirmation.\n' +
        'No automatic Excel changes.'
      );

    } catch (error) {
      show(
        'V6 CONNECTION ERROR\n' + detail(error)
      );
    } finally {
      clearTimeout(timer);
      initializing = false;
    }
  }

  if (retry) {
    retry.addEventListener('click', initialize);
  }

  initialize();
})();
