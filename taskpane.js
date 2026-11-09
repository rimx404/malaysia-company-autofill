(() => {
  'use strict';

  const VERSION = '8';
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
    if (typeof url !== 'string') return null;

    try {
      const parsed = new URL(url);

      if (
        parsed.protocol !== 'https:' &&
        parsed.protocol !== 'http:'
      ) {
        return null;
      }

      const link = document.createElement('a');
      link.href = parsed.href;
      link.target = '_blank';
      link.rel = 'noopener noreferrer';
      link.textContent = text;
      parent.appendChild(link);
      return link;
    } catch {
      return null;
    }
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
    const lastColumn = columnNumber(
      match[3] || match[1]
    );

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

  function validId(id) {
    return /^Q\d+$/.test(String(id || ''));
  }

  function uniqueStrings(values) {
    if (!Array.isArray(values)) return [];

    return [
      ...new Set(
        values
          .filter(value => typeof value === 'string')
          .map(value => value.trim())
          .filter(Boolean)
      )
    ];
  }

  function makeField(column, title, values, source) {
    return {
      column,
      title,
      values: uniqueStrings(values),
      source
    };
  }

  function globalFields(candidate) {
    const profile = candidate.global_profile || {};
    const id = String(candidate.wikidata_id || '');
    const sources = candidate.field_sources || {};

    function source(property, key) {
      return sources[key] ||
        'https://www.wikidata.org/wiki/' +
        id + '#' + property;
    }

    return [
      makeField(
        'E',
        'Company Type',
        profile.company_type,
        source('P31', 'company_type')
      ),
      makeField(
        'F',
        'Company Headquarters',
        profile.headquarters,
        source('P159', 'headquarters')
      ),
      makeField(
        'G',
        'Industry',
        profile.industry,
        source('P452', 'industry')
      ),
      makeField(
        'K',
        'Company Website',
        profile.websites,
        source('P856', 'websites')
      )
    ];
  }

  function renderGlobalDetails(parent, candidate) {
    const section = document.createElement('div');

    section.style.marginTop = '14px';
    section.style.padding = '12px';
    section.style.border = '1px solid #cbd5df';
    section.style.borderRadius = '8px';

    addText(
      section,
      'h3',
      'V8 — Global Company Background'
    );

    addText(
      section,
      'p',
      'Information reported by Wikidata. ' +
      'Not independently verified.'
    );

    const profile = candidate.global_profile || {};

    const fields = globalFields(candidate);

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

      addText(
        item,
        'p',
        field.values.length
          ? field.values.join('; ')
          : 'Not available from this source'
      );

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

    const extra = [
      ['Country', profile.country],
      ['Country of origin', profile.country_of_origin]
    ];

    extra.forEach(([title, values]) => {
      const item = document.createElement('div');
      item.style.padding = '10px 0';

      addText(item, 'strong', title);

      addText(
        item,
        'p',
        uniqueStrings(values).join('; ') ||
          'Not available from this source'
      );

      section.appendChild(item);
    });

    addText(
      section,
      'p',
      'Preview only. Columns E, F, G and K ' +
      'have not been changed.'
    );

    parent.appendChild(section);
  }

  function renderMalaysiaDetails(parent, candidate) {
    const profile = candidate.malaysia_profile || {};

    const section = document.createElement('div');

    section.style.marginTop = '14px';
    section.style.padding = '12px';
    section.style.border = '1px solid #cbd5df';
    section.style.borderRadius = '8px';

    addText(
      section,
      'h3',
      'Malaysia-Specific Research'
    );

    addText(
      section,
      'p',
      'This section is separate from the ' +
      'global company background.'
    );

    const fields = [
      [
        'Malaysia presence',
        profile.malaysia_presence_status
      ],
      [
        'Malaysia company name',
        profile.malaysia_company_name
      ],
      [
        'Malaysia headquarters',
        profile.malaysia_headquarters
      ],
      [
        'Malaysia office address',
        profile.malaysia_office_address
      ],
      [
        'Malaysia website',
        profile.malaysia_website
      ],
      [
        'Malaysia careers website',
        profile.malaysia_careers_url
      ]
    ];

    fields.forEach(([title, value]) => {
      const item = document.createElement('div');

      item.style.padding = '8px 0';
      item.style.borderBottom =
        '1px solid #e5e7eb';

      addText(item, 'strong', title);

      addText(
        item,
        'p',
        value || 'Not independently verified'
      );

      if (
        typeof value === 'string' &&
        /^https?:\/\//i.test(value)
      ) {
        addLink(item, 'Open source website', value);
      }

      section.appendChild(item);
    });

    const evidence = Array.isArray(
      profile.malaysia_evidence
    )
      ? profile.malaysia_evidence
      : [];

    addText(section, 'h4', 'Malaysia Evidence');

    if (!evidence.length) {
      addText(
        section,
        'p',
        'No independent Malaysia-specific ' +
        'evidence available from this research.'
      );
    } else {
      evidence.forEach(item => {
        const block = document.createElement('div');
        block.style.padding = '8px 0';

        addText(
          block,
          'strong',
          item.type || 'Research signal'
        );

        addText(
          block,
          'p',
          item.note ||
            'This signal is not independent verification.'
        );

        addLink(
          block,
          'View source',
          item.source_url
        );

        section.appendChild(block);
      });
    }

    addText(
      section,
      'p',
      profile.research_note ||
      'Verify the Malaysian legal entity, ' +
      'office and careers website separately.'
    );

    addText(
      section,
      'p',
      'Malaysia presence verified: ' +
      (
        candidate.verified_malaysia_presence === true
          ? 'Reported verified by source'
          : 'No'
      )
    );

    parent.appendChild(section);
  }

  function buildNote(company, candidate) {
    const id = String(candidate.wikidata_id || '');
    const global = candidate.global_profile || {};
    const malaysia = candidate.malaysia_profile || {};

    function values(items) {
      return uniqueStrings(items).join('; ') ||
        'Not available';
    }

    return [
      'UNVERIFIED COMPANY RESEARCH — V8',
      'Searched company: ' + company,
      'Candidate: ' + String(candidate.name || ''),
      'Description: ' +
        String(candidate.description || ''),
      'Wikidata ID: ' + id,
      'Source: https://www.wikidata.org/wiki/' + id,
      '',
      'GLOBAL BACKGROUND (UNVERIFIED)',
      'Company type: ' + values(global.company_type),
      'Headquarters: ' + values(global.headquarters),
      'Industry: ' + values(global.industry),
      'Country: ' + values(global.country),
      'Websites: ' + values(global.websites),
      '',
      'MALAYSIA-SPECIFIC RESEARCH',
      'Presence: ' +
        String(
          malaysia.malaysia_presence_status ||
          'Not independently verified'
        ),
      'Malaysia company name: ' +
        String(
          malaysia.malaysia_company_name ||
          'Not verified'
        ),
      'Malaysia office address: ' +
        String(
          malaysia.malaysia_office_address ||
          'Not verified'
        ),
      'Malaysia website: ' +
        String(
          malaysia.malaysia_website ||
          'Not verified'
        ),
      'Malaysia careers: ' +
        String(
          malaysia.malaysia_careers_url ||
          'Not verified'
        ),
      '',
      'REVIEW REQUIRED',
      'Global information does not prove ' +
        'Malaysian office presence.',
      'Further verification required.'
    ].join('\n');
  }

  async function approveAndSave(
    company,
    row,
    candidate,
    button,
    message,
    token
  ) {
    const id = String(candidate.wikidata_id || '');

    if (!validId(id)) {
      message.textContent =
        'Cannot save: invalid Wikidata ID.';
      return;
    }

    if (token !== requestNumber) {
      message.textContent =
        'Research changed. Select again.';
      return;
    }

    const note = buildNote(company, candidate);

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

        if (token !== requestNumber) {
          return {
            saved: false,
            reason: 'Research changed. Nothing saved.'
          };
        }

        const currentCompany =
          String(nameCell.values[0][0] ?? '').trim();

        if (currentCompany !== company) {
          return {
            saved: false,
            reason:
              'Company name changed. Research again.'
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
        'V8 SAVE COMPLETE\n' +
        'Research note saved to AP' + row + '.\n' +
        'E, F, G and K remain unchanged.'
      );

      message.textContent =
        'Research note saved to AP' + row +
        '. Company details remain preview-only.';

      button.textContent = 'Saved to Excel';
    } catch (error) {
      show('SAVE ERROR\n' + errorText(error));

      message.textContent =
        'Save not confirmed: ' + errorText(error);

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

    addText(results, 'p', 'Searched: ' + company);

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

      if (validId(id)) {
        addLink(
          card,
          'Open Wikidata source',
          candidate.source_url ||
            'https://www.wikidata.org/wiki/' + id
        );
      }

      addText(
        card,
        'p',
        'Candidate identity and Malaysian ' +
        'operations require verification.',
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
          show('Results outdated. Research again.');
          return;
        }

        if (!validId(id)) {
          message.textContent =
            'Invalid candidate ID.';
          return;
        }

        card.style.border = '2px solid #13865c';
        card.style.background = '#edf8f2';

        select.disabled = true;
        select.textContent = 'Selected for review';

        message.textContent =
          'Selected: ' +
          String(candidate.name || '') +
          '. Malaysia-specific information ' +
          'is not independently verified.';

        const approve = document.createElement('button');
        approve.type = 'button';
        approve.textContent = 'Approve & save note';

        approve.addEventListener('click', async () => {
          await approveAndSave(
            company,
            row,
            candidate,
            approve,
            message,
            token
          );
        });

        card.appendChild(approve);

        const details = document.createElement('div');

        renderGlobalDetails(details, candidate);
        renderMalaysiaDetails(details, candidate);

        card.appendChild(details);

        show(
          'V8 DETAILS READY\n' +
          'Company: ' + company + '\n' +
          'Wikidata: ' + id + '\n' +
          'Global and Malaysia-specific previews ready.\n' +
          'No company-detail cells changed.'
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
      'V8 RESEARCHING\n' +
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

      if (data.status === 'research_error') {
        throw new Error(
          data.message || 'Worker research failed.'
        );
      }

      const candidates = Array.isArray(data.candidates)
        ? data.candidates
        : [];

      show(
        'V8 RESEARCH COMPLETE\n' +
        'Worker: ' + (data.version || 'unknown') + '\n' +
        'Candidates: ' + candidates.length + '\n' +
        'Select one to review global and ' +
        'Malaysia-specific information.'
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

        show('D' + row + ' is blank.');
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
        'V8 READY\n' +
        'Monitoring Company Master D7 onward.\n' +
        'Edit D7 to start research.'
      );
      return;
    }

    if (initializing) return;

    initializing = true;
    show('V8: Connecting to Excel...');

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
          throw
