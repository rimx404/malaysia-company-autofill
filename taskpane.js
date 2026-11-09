(() => {
  'use strict';

  const WORKER =
    'https://malaysia-company-research.rimli271.workers.dev';

  const SHEET = 'Company Master';
  const FIRST_ROW = 7;
  const NOTE_COLUMN = 'AP';

  const status = document.getElementById('status');
  const results = document.getElementById('results');
  const retry = document.getElementById('retry');

  let active = false;
  let initializing = false;
  let handlerResult = null;
  let requestNumber = 0;

  function show(message) {
    if (status) status.textContent = message;
  }

  function errorText(error) {
    return String(
      (error && (error.message || error.stack)) || error
    );
  }

  function clearResults() {
    if (results) results.replaceChildren();
  }

  function addText(parent, tag, value, className) {
    const element = document.createElement(tag);
    element.textContent = String(value);

    if (className) element.className = className;

    parent.appendChild(element);
    return element;
  }

  function columnNumber(letters) {
    return [...letters].reduce(
      (number, letter) =>
        number * 26 + letter.charCodeAt(0) - 64,
      0
    );
  }

  function rowFromAddress(address) {
    const local = String(address || '')
      .split('!')
      .pop()
      .replace(/\$/g, '')
      .toUpperCase();

    const match =
      /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(local);

    if (!match) return null;

    const left = columnNumber(match[1]);
    const right = columnNumber(match[3] || match[1]);

    if (left > 4 || right < 4) return null;

    const first = Number(match[2]);
    const last = Number(match[4] || match[2]);

    if (last < FIRST_ROW) return null;

    const start = Math.max(FIRST_ROW, first);

    if (start !== last) {
      return { multiple: true, start, last };
    }

    return { multiple: false, row: start };
  }

  async function readCompany(row) {
    return Excel.run(async context => {
      const sheet = context.workbook.worksheets.getItem(SHEET);
      const cell = sheet.getRange('D' + row);

      cell.load('values');
      await context.sync();

      return String(cell.values[0][0] ?? '').trim();
    });
  }

  async function approveAndSave(
    company,
    row,
    candidate,
    button,
    message
  ) {
    if (!candidate || !candidate.wikidata_id) {
      show('Cannot save: candidate has no Wikidata ID.');
      return;
    }

    const id = String(candidate.wikidata_id);

    if (!/^Q\d+$/.test(id)) {
      show('Cannot save: invalid Wikidata ID.');
      return;
    }

    const source = 'https://www.wikidata.org/entity/' + id;

    const note = [
      'UNVERIFIED RESEARCH CANDIDATE',
      'Searched company: ' + company,
      'Candidate: ' + String(candidate.name || ''),
      'Description: ' + String(candidate.description || ''),
      'Wikidata ID: ' + id,
      'Source: ' + source,
      'Malaysia-specific identity NOT verified.',
      'Further verification required before relying on this data.'
    ].join('\n');

    button.disabled = true;
    button.textContent = 'Checking Excel...';

    try {
      const outcome = await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem(SHEET);

        const nameCell = sheet.getRange('D' + row);
        const noteCell = sheet.getRange(NOTE_COLUMN + row);

        nameCell.load('values');
        noteCell.load('values,formulas');

        await context.sync();

        const currentCompany =
          String(nameCell.values[0][0] ?? '').trim();

        if (currentCompany !== company) {
          return {
            saved: false,
            reason: 'Company name changed. Research again.'
          };
        }

        const existing = noteCell.values[0][0];
        const formula = noteCell.formulas[0][0];

        if (
          (existing !== null &&
           existing !== undefined &&
           String(existing).trim() !== '') ||
          (typeof formula === 'string' &&
           formula.startsWith('='))
        ) {
          return {
            saved: false,
            reason:
              'AP' + row +
              ' already contains data or a formula. Nothing overwritten.'
          };
        }

        noteCell.values = [[note]];
        await context.sync();

        return { saved: true };
      });

      if (!outcome.saved) {
        show('SAVE BLOCKED\n' + outcome.reason);
        message.textContent = outcome.reason;
        button.disabled = false;
        button.textContent = 'Approve & save note';
        return;
      }

      show(
        'SAVE COMPLETE\n' +
        'Company: ' + company + '\n' +
        'Saved to: ' + NOTE_COLUMN + row + '\n' +
        'Unverified candidate and source recorded.\n' +
        'All other cells left unchanged.'
      );

      message.textContent =
        'Saved to ' + NOTE_COLUMN + row +
        '. This is an unverified research note, not a confirmed company profile.';

      button.textContent = 'Saved to Excel';
    } catch (error) {
      const explanation = errorText(error);

      show(
        'SAVE ERROR\n' +
        explanation + '\n\n' +
        'If Excel reports insufficient permissions, ' +
        'the add-in manifest must be updated to ' +
        'ReadWriteDocument.'
      );

      message.textContent =
        'Save was not confirmed. Check the error above.';

      button.disabled = false;
      button.textContent = 'Approve & save note';
    }
  }

  function renderCandidates(company, row, candidates, token) {
    clearResults();

    if (!results) {
      show('Missing HTML element: results');
      return;
    }

    addText(results, 'h2', 'Review research for D' + row);
    addText(results, 'p', 'Searched: ' + company);

    if (!candidates.length) {
      addText(
        results,
        'p',
        'No candidates found. Nothing was changed.'
      );
      return;
    }

    candidates.forEach(candidate => {
      const card = document.createElement('div');
      card.className = 'candidate';

      addText(
        card,
        'h3',
        candidate.name || 'Unnamed candidate'
      );

      addText(
        card,
        'p',
        candidate.description || 'No description available.'
      );

      addText(
        card,
        'p',
        'Wikidata ID: ' +
          String(candidate.wikidata_id || 'Unavailable'),
        'small'
      );

      const id = String(candidate.wikidata_id || '');

      if (/^Q\d+$/.test(id)) {
        const link = document.createElement('a');

        link.href =
          'https://www.wikidata.org/entity/' + id;

        link.target = '_blank';
        link.rel = 'noopener noreferrer';
        link.textContent = 'Open Wikidata source';

        card.appendChild(link);
      }

      addText(
        card,
        'p',
        'Not verified as the correct Malaysian company.',
        'small'
      );

      const select = document.createElement('button');
      select.type = 'button';
      select.textContent = 'Select for review';

      const message = addText(
        card,
        'p',
        '',
        'small'
      );

      select.addEventListener('click', () => {
        if (token !== requestNumber) {
          show('These results are outdated. Research again.');
          return;
        }

        card.style.border = '2px solid #13865c';
        card.style.background = '#edf8f2';

        select.disabled = true;
        select.textContent = 'Selected for review';

        message.textContent =
          'Selected: ' +
          String(candidate.name || '') +
          '. This is not a verified Malaysian company match.';

        const approve = document.createElement('button');
        approve.type = 'button';
        approve.textContent = 'Approve & save note';

        approve.addEventListener('click', async () => {
          if (token !== requestNumber) {
            show('Research changed. Please select a current result.');
            return;
          }

          await approveAndSave(
            company,
            row,
            candidate,
            approve,
            message
          );
        });

        card.appendChild(approve);

        show(
          'SELECTED FOR REVIEW\n' +
          'Company: ' + company + '\n' +
          'Candidate: ' + String(candidate.name || '') + '\n' +
          'Wikidata ID: ' + id + '\n\n' +
          'Tap Approve & save note to write the unverified ' +
          'source information into AP' + row + '.\n' +
          'No other cells will be changed.'
        );
      });

      card.appendChild(select);
      results.appendChild(card);
    });

    const reject = document.createElement('button');
    reject.type = 'button';
    reject.textContent = 'Reject all candidates';

    reject.addEventListener('click', () => {
      if (token !== requestNumber) return;

      ++requestNumber;
      clearResults();

      show(
        'All candidates rejected for D' + row +
        '.\nNo Excel cells changed.'
      );
    });

    results.appendChild(reject);
  }

  async function research(company, row) {
    const token = ++requestNumber;

    clearResults();

    show(
      'RESEARCHING\n' +
      'Cell: D' + row + '\n' +
      'Company: ' + company + '\n' +
      'Source: Wikidata via Cloudflare Worker\n' +
      'No Excel cells changed.'
    );

    const controller = new AbortController();

    const timeout = setTimeout(
      () => controller.abort(),
      20000
    );

    try {
      const url =
        WORKER +
        '/research?company=' +
        encodeURIComponent(company);

      const response = await fetch(url, {
        method: 'GET',
        headers: {
          Accept: 'application/json'
        },
        signal: controller.signal
      });

      if (!response.ok) {
        throw new Error(
          'Research service returned HTTP ' + response.status
        );
      }

      const data = await response.json();

      if (token !== requestNumber) return;

      const candidates = Array.isArray(data.candidates)
        ? data.candidates
        : [];

      show(
        'RESEARCH COMPLETE\n' +
        'Company: ' + company + '\n' +
        'Candidates: ' + candidates.length + '\n' +
        'Review and select a candidate below.\n' +
        'No Excel cells changed.'
      );

      renderCandidates(company, row, candidates, token);
    } catch (error) {
      if (token !== requestNumber) return;

      clearResults();

      show(
        'RESEARCH ERROR\n' +
        errorText(error) +
        '\nNo Excel cells changed.'
      );
    } finally {
      clearTimeout(timeout);
    }
  }

  async function onChanged(event) {
    try {
      const location = rowFromAddress(event.address);

      if (!location) return;

      if (location.multiple) {
        ++requestNumber;
        clearResults();

        show(
          'Multiple company rows edited.\n' +
          'Please research one company at a time.\n' +
          'No Excel cells changed.'
        );

        return;
      }

      const row = location.row;
      const company = await readCompany(row);

      if (!company) {
        ++requestNumber;
        clearResults();

        show(
          'D' + row + ' is blank.\n' +
          'No Excel cells changed.'
        );

        return;
      }

      await research(company, row);
    } catch (error) {
      show(
        'CHANGE EVENT ERROR\n' +
        errorText(error)
      );
    }
  }

  async function initialize() {
    if (active) {
      show(
        'V6 READY\n' +
        'Monitoring Company Master column D, row 7 onward.\n' +
        'Edit D7 to research a company.\n' +
        'Approval required before saving.'
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
        throw new Error('Office.js is unavailable.');
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

      if (
        !info ||
        info.host !== Office.HostType.Excel
      ) {
        throw new Error('Excel host unavailable.');
      }

      if (typeof Excel === 'undefined') {
        throw new Error('Excel JavaScript API unavailable.');
      }

      await Excel.run(async context => {
        const sheet = context.workbook.worksheets
          .getItemOrNullObject(SHEET);

        sheet.load('isNullObject');
        await context.sync();

        if (sheet.isNullObject) {
          throw new Error(
            'Worksheet "Company Master" not found.'
          );
        }

        handlerResult = sheet.onChanged.add(onChanged);

        await context.sync();
      });

      active = true;

      show(
        'V6 READY — SAFE APPROVAL MODE\n' +
        'Monitoring Company Master column D, row 7 onward.\n' +
        'Edit D7 to start research.\n' +
        'Select a candidate, then approve a note.\n' +
        'Existing data and formulas are protected.'
      );
    } catch (error) {
      show(
        'V6 CONNECTION ERROR\n' +
        errorText(error)
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
