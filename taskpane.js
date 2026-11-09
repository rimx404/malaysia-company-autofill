(() => {
  'use strict';

  const WORKER =
    'https://malaysia-company-research.rimli271.workers.dev';

  const WIKIDATA_API =
    'https://www.wikidata.org/w/api.php';

  const SHEET = 'Company Master';
  const FIRST_ROW = 7;
  const NOTE_COLUMN = 'AP';

  const status = document.getElementById('status');
  const results = document.getElementById('results');
  const retry = document.getElementById('retry');

  let active = false;
  let initializing = false;
  let requestNumber = 0;

  function show(message) {
    if (status) status.textContent = message;
  }

  function errorText(error) {
    return String(error?.message || error);
  }

  function clearResults() {
    if (results) results.replaceChildren();
  }

  function addText(parent, tag, value, className) {
    const element = document.createElement(tag);
    element.textContent = String(value ?? '');

    if (className) element.className = className;

    parent.appendChild(element);
    return element;
  }

  function addLink(parent, text, url) {
    const link = document.createElement('a');
    link.href = url;
    link.target = '_blank';
    link.rel = 'noopener noreferrer';
    link.textContent = text;
    parent.appendChild(link);
    return link;
  }

  function columnNumber(letters) {
    return [...letters].reduce(
      (number, letter) =>
        number * 26 + letter.charCodeAt(0) - 64,
      0
    );
  }

  function changedRow(address) {
    const local = String(address || '')
      .split('!')
      .pop()
      .replace(/\$/g, '')
      .toUpperCase();

    const match =
      /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(local);

    if (!match) return null;

    const firstColumn = columnNumber(match[1]);
    const lastColumn = columnNumber(match[3] || match[1]);

    if (firstColumn > 4 || lastColumn < 4) {
      return null;
    }

    const firstRow = Number(match[2]);
    const lastRow = Number(match[4] || match[2]);

    if (lastRow < FIRST_ROW) return null;

    const start = Math.max(firstRow, FIRST_ROW);

    if (start !== lastRow) {
      return { multiple: true };
    }

    return { row: start, multiple: false };
  }

  async function readCompany(row) {
    return Excel.run(async context => {
      const sheet =
        context.workbook.worksheets.getItem(SHEET);

      const cell = sheet.getRange('D' + row);
      cell.load('values');

      await context.sync();

      return String(cell.values[0][0] ?? '').trim();
    });
  }

  async function wikidataRequest(parameters) {
    const query = new URLSearchParams({
      action: 'wbgetentities',
      format: 'json',
      origin: '*',
      ...parameters
    });

    const controller = new AbortController();
    const timeout = setTimeout(
      () => controller.abort(),
      15000
    );

    try {
      const response = await fetch(
        WIKIDATA_API + '?' + query.toString(),
        {
          method: 'GET',
          signal: controller.signal
        }
      );

      if (!response.ok) {
        throw new Error(
          'Wikidata returned HTTP ' + response.status
        );
      }

      const data = await response.json();

      if (data.error) {
        throw new Error(
          data.error.info || 'Wikidata API error'
        );
      }

      return data.entities || {};
    } finally {
      clearTimeout(timeout);
    }
  }

  function statementValues(entity, property) {
    const statements = entity?.claims?.[property] || [];

    const valid = statements.filter(statement => {
      return (
        statement.rank !== 'deprecated' &&
        statement.mainsnak?.snaktype === 'value' &&
        statement.mainsnak?.datavalue
      );
    });

    const preferred = valid.filter(
      statement => statement.rank === 'preferred'
    );

    const chosen = preferred.length ? preferred : valid;

    return chosen
      .map(statement =>
        statement.mainsnak.datavalue.value
      )
      .slice(0, 5);
  }

  function itemIds(values) {
    return values
      .map(value => {
        if (value?.id && /^Q\d+$/.test(value.id)) {
          return value.id;
        }

        if (
          value?.['entity-type'] === 'item' &&
          Number.isInteger(value['numeric-id'])
        ) {
          return 'Q' + value['numeric-id'];
        }

        return null;
      })
      .filter(Boolean);
  }

  async function fetchCompanyDetails(id) {
    if (!/^Q\d+$/.test(id)) {
      throw new Error('Invalid Wikidata ID.');
    }

    const entities = await wikidataRequest({
      ids: id,
      props: 'claims|labels|descriptions',
      languages: 'en'
    });

    const entity = entities[id];

    if (!entity || entity.missing !== undefined) {
      throw new Error('Wikidata entity not found.');
    }

    const fields = [
      {
        column: 'E',
        title: 'Company Type',
        property: 'P31'
      },
      {
        column: 'F',
        title: 'Company Headquarters',
        property: 'P159'
      },
      {
        column: 'G',
        title: 'Industry',
        property: 'P452'
      },
      {
        column: 'K',
        title: 'Company Website',
        property: 'P856'
      }
    ];

    const raw = {};

    for (const field of fields) {
      raw[field.property] =
        statementValues(entity, field.property);
    }

    const relatedIds = [
      ...new Set([
        ...itemIds(raw.P31),
        ...itemIds(raw.P159),
        ...itemIds(raw.P452)
      ])
    ];

    let related = {};

    if (relatedIds.length) {
      related = await wikidataRequest({
        ids: relatedIds.join('|'),
        props: 'labels',
        languages: 'en',
        languagefallback: '1'
      });
    }

    return fields.map(field => {
      const values = raw[field.property];

      let displayValues = [];

      if (field.property === 'P856') {
        displayValues = values.filter(value => {
          if (typeof value !== 'string') return false;

          try {
            const url = new URL(value);
            return (
              url.protocol === 'https:' ||
              url.protocol === 'http:'
            );
          } catch {
            return false;
          }
        });
      } else {
        displayValues = itemIds(values).map(relatedId => {
          return (
            related[relatedId]?.labels?.en?.value ||
            relatedId
          );
        });
      }

      return {
        ...field,
        values: [...new Set(displayValues)],
        source:
          'https://www.wikidata.org/wiki/' +
          id + '#' + field.property
      };
    });
  }

  function renderDetails(parent, fields) {
    const section = document.createElement('div');

    section.style.marginTop = '14px';
    section.style.padding = '12px';
    section.style.border = '1px solid #cbd5df';
    section.style.borderRadius = '8px';

    addText(
      section,
      'h3',
      'V7 — Company Details Preview'
    );

    addText(
      section,
      'p',
      'Wikidata-reported values only. ' +
      'These have not been verified for Malaysia.'
    );

    fields.forEach(field => {
      const item = document.createElement('div');

      item.style.padding = '10px 0';
      item.style.borderBottom =
        '1px solid #e5e7eb';

      addText(
        item,
        'strong',
        field.column + ' — ' + field.title
      );

      const value = field.values.length
        ? field.values.join('; ')
        : 'Not available from this Wikidata record';

      addText(item, 'p', value);

      addLink(
        item,
        'View Wikidata property source',
        field.source
      );

      addText(
        item,
        'p',
        'Status: Not independently verified',
        'small'
      );

      section.appendChild(item);
    });

    addText(
      section,
      'p',
      'Preview only — columns E, F, G and K ' +
      'have not been changed.'
    );

    parent.appendChild(section);
  }

  async function approveAndSave(
    company,
    row,
    candidate,
    button,
    message
  ) {
    const id = String(candidate.wikidata_id || '');

    if (!/^Q\d+$/.test(id)) {
      show('Cannot save: invalid Wikidata ID.');
      return;
    }

    const note = [
      'UNVERIFIED RESEARCH CANDIDATE',
      'Searched company: ' + company,
      'Candidate: ' + String(candidate.name || ''),
      'Description: ' +
        String(candidate.description || ''),
      'Wikidata ID: ' + id,
      'Source: https://www.wikidata.org/entity/' + id,
      'Malaysia-specific identity NOT verified.',
      'Further verification required before relying on this data.'
    ].join('\n');

    button.disabled = true;
    button.textContent = 'Checking Excel...';

    try {
      const outcome = await Excel.run(async context => {
        const sheet =
          context.workbook.worksheets.getItem(SHEET);

        const nameCell = sheet.getRange('D' + row);
        const noteCell =
          sheet.getRange(NOTE_COLUMN + row);

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
              ' already contains data or a formula. ' +
              'Nothing overwritten.'
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
        'Saved unverified note to AP' + row + '.\n' +
        'Columns E, F, G and K unchanged.'
      );

      message.textContent =
        'Saved to AP' + row +
        '. Company details remain preview-only.';

      button.textContent = 'Saved to Excel';
    } catch (error) {
      show('SAVE ERROR\n' + errorText(error));

      message.textContent =
        'Save not confirmed. See error above.';

      button.disabled = false;
      button.textContent = 'Approve & save note';
    }
  }

  function renderCandidates(company, row, candidates, token) {
    clearResults();

    if (!results) return;

    addText(
      results,
      'h2',
      'Review research for D' + row
    );

    addText(
      results,
      'p',
      'Searched: ' + company
    );

    if (!candidates.length) {
      addText(
        results,
        'p',
        'No candidates found. Nothing changed.'
      );
      return;
    }

    candidates.forEach(candidate => {
      const card = document.createElement('div');
      card.className = 'candidate';

      const id = String(
        candidate.wikidata_id || ''
      );

      addText(
        card,
        'h3',
        candidate.name || 'Unnamed candidate'
      );

      addText(
        card,
        'p',
        candidate.description ||
          'No description available.'
      );

      addText(
        card,
        'p',
        'Wikidata ID: ' + id,
        'small'
      );

      if (/^Q\d+$/.test(id)) {
        addLink(
          card,
          'Open Wikidata source',
          'https://www.wikidata.org/wiki/' + id
        );
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

      select.addEventListener('click', async () => {
        if (token !== requestNumber) {
          show('Results outdated. Research again.');
          return;
        }

        if (!/^Q\d+$/.test(id)) {
          show('Invalid candidate ID.');
          return;
        }

        card.style.border = '2px solid #13865c';
        card.style.background = '#edf8f2';

        select.disabled = true;
        select.textContent = 'Selected for review';

        message.textContent =
          'Selected: ' +
          String(candidate.name || '') +
          '. Malaysian identity not verified.';

        const approve = document.createElement('button');
        approve.type = 'button';
        approve.textContent = 'Approve & save note';

        approve.addEventListener('click', async () => {
          if (token !== requestNumber) {
            show('Research changed. Select again.');
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

        const details = document.createElement('div');
        addText(
          details,
          'p',
          'Loading company details from Wikidata...'
        );

        card.appendChild(details);

        show(
          'V7: Candidate selected.\n' +
          'Loading details for ' + id + '.\n' +
          'No Excel cells changed.'
        );

        try {
          const fields = await fetchCompanyDetails(id);

          if (token !== requestNumber) return;

          details.replaceChildren();
          renderDetails(details, fields);

          show(
            'V7 DETAILS READY\n' +
            'Company: ' + company + '\n' +
            'Wikidata: ' + id + '\n' +
            'Review E, F, G and K below.\n' +
            'These are unverified preview values.\n' +
            'No company-detail cells changed.'
          );
        } catch (error) {
          if (token !== requestNumber) return;

          details.replaceChildren();

          addText(
            details,
            'p',
            'Details unavailable: ' + errorText(error)
          );

          addText(
            details,
            'p',
            'Candidate selection and AP note saving ' +
            'remain available.'
          );

          show(
            'V7 DETAIL LOOKUP ERROR\n' +
            errorText(error) + '\n' +
            'No company-detail cells changed.'
          );
        }
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
        'All candidates rejected.\n' +
        'No Excel cells changed.'
      );
    });

    results.appendChild(reject);
  }

  async function research(company, row) {
    const token = ++requestNumber;

    clearResults();

    show(
      'V7 RESEARCHING\n' +
      'Cell: D' + row + '\n' +
      'Company: ' + company
    );

    const controller = new AbortController();
    const timeout = setTimeout(
      () => controller.abort(),
      20000
    );

    try {
      const response = await fetch(
        WORKER +
          '/research?company=' +
          encodeURIComponent(company),
        {
          headers: {
            Accept: 'application/json'
          },
          signal: controller.signal
        }
      );

      if (!response.ok) {
        throw new Error(
          'Research HTTP ' + response.status
        );
      }

      const data = await response.json();

      if (token !== requestNumber) return;

      const candidates = Array.isArray(data.candidates)
        ? data.candidates
        : [];

      show(
        'V7 RESEARCH COMPLETE\n' +
        'Candidates: ' + candidates.length + '\n' +
        'Select one to preview company details.'
      );

      renderCandidates(
        company,
        row,
        candidates,
        token
      );
    } catch (error) {
      if (token !== requestNumber) return;

      clearResults();

      show(
        'RESEARCH ERROR\n' + errorText(error)
      );
    } finally {
      clearTimeout(timeout);
    }
  }

  async function onChanged(event) {
    try {
      const location = changedRow(event.address);

      if (!location) return;

      if (location.multiple) {
        ++requestNumber;
        clearResults();

        show(
          'Multiple company rows edited.\n' +
          'Please research one row at a time.'
        );

        return;
      }

      const row = location.row;
      const company = await readCompany(row);

      if (!company) {
        ++requestNumber;
        clearResults();

        show(
          'D' + row + ' is blank.'
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
        'V7 READY\n' +
        'Monitoring Company Master D7 onward.\n' +
        'Edit D7 to start research.'
      );
      return;
    }

    if (initializing) return;

    initializing = true;
    show('V7: Connecting to Excel...');

    try {
      await Office.onReady();

      if (typeof Excel === 'undefined') {
        throw new Error(
          'Excel JavaScript API unavailable.'
        );
      }

      await Excel.run(async context => {
        const sheet = context.workbook.worksheets
          .getItemOrNullObject(SHEET);

        sheet.load('isNullObject');
        await context.sync();

        if (sheet.isNullObject) {
          throw new Error(
            'Company Master sheet not found.'
          );
        }

        sheet.onChanged.add(onChanged);
        await context.sync();
      });

      active = true;

      show(
        'V7 READY — COMPANY DETAILS PREVIEW\n' +
        'Monitoring Company Master column D, ' +
        'row 7 onward.\n' +
        'Edit D7 to research.\n' +
        'Select a candidate to preview E, F, G, K.\n' +
        'Company details will NOT be written.\n' +
        'AP note saving remains approval-only.'
      );
    } catch (error) {
      show(
        'V7 CONNECTION ERROR\n' +
        errorText(error)
      );
    } finally {
      initializing = false;
    }
  }

  if (retry) {
    retry.addEventListener('click', initialize);
  }

  initialize();
})();
