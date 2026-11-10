(() => {
  'use strict';

  // Malaysia Company Research task pane V10.2.
  // Compatible with Worker V10.2's distinct MY: local record IDs.
  // Only the single Worker-designated Malaysia save target can be approved.
  // L/M/N are the sole exception: global-parent metrics, labelled in AP.

  const VERSION = '10.2';
  const WORKER =
    'https://malaysia-company-research.rimli271.workers.dev';

  const SHEET = 'Company Master';
  const HEADER_ROW = 6;
  const FIRST_ROW = 7;
  const NOTE_COLUMN = 'AP';
  const PLACEHOLDER = '[verify office address]';

  // These are the 23 columns selected for the Company Master research system.
  const SELECTED_COLUMNS = [
    'A', 'B', 'C', 'D', 'E', 'F', 'G', 'I', 'J', 'K', 'L', 'M',
    'N', 'O', 'P', 'Q', 'R', 'S', 'AE', 'AF', 'AG', 'AH', 'AI'
  ];

  // A, B, C, D and AE are deliberately not written by this add-in.
  // C (Fortune Global 500 rank) remains unfilled until a dated official source
  // is supplied by the Worker. A/B/D/AE are user-maintained or protected.
  const WRITABLE_COLUMNS = [
    'E', 'F', 'G', 'I', 'J', 'K', 'L', 'M', 'N', 'O',
    'P', 'Q', 'R', 'S', 'AF', 'AG', 'AH', 'AI'
  ];

  const GLOBAL_COLUMNS = new Set(['L', 'M', 'N']);
  const REPORTING_DATE_COLUMNS = new Set(['L', 'M', 'N', 'O']);
  const ALLOWED_STATUSES = new Set([
    'official',
    'official_lower_bound',
    'reported_needs_review',
    'derived_fx_estimate',
    'review_required'
  ]);

  const status = document.getElementById('status');
  const results = document.getElementById('results');
  const retry = document.getElementById('retry');

  let active = false;
  let starting = false;
  let requestNumber = 0;

  function show(message) {
    if (status) status.textContent = message;
  }

  function clearResults() {
    if (results) results.replaceChildren();
  }

  function errorText(error) {
    return String(error?.message || error);
  }

  function addText(parent, tag, value, className) {
    const node = document.createElement(tag);
    node.textContent = String(value ?? '');
    if (className) node.className = className;
    parent.appendChild(node);
    return node;
  }

  function addLink(parent, label, value) {
    if (typeof value !== 'string' || !value.trim()) return null;

    try {
      const url = new URL(value);
      if (url.protocol !== 'https:' && url.protocol !== 'http:') {
        return null;
      }

      const anchor = document.createElement('a');
      anchor.href = url.href;
      anchor.target = '_blank';
      anchor.rel = 'noopener noreferrer';
      anchor.textContent = label;
      parent.appendChild(anchor);
      return anchor;
    } catch {
      return null;
    }
  }

  function panel(parent, heading) {
    const node = document.createElement('section');
    node.style.padding = '12px';
    node.style.marginTop = '14px';
    node.style.border = '1px solid #cbd5df';
    node.style.borderRadius = '8px';
    addText(node, 'h3', heading);
    parent.appendChild(node);
    return node;
  }

  function displayValue(value) {
    if (value === null || value === undefined || value === '') {
      return 'Not available';
    }
    if (Array.isArray(value)) {
      const parts = value.map(x => String(x ?? '').trim()).filter(Boolean);
      return parts.length ? [...new Set(parts)].join('; ') : 'Not available';
    }
    if (typeof value === 'number') {
      return Number.isFinite(value) ? value.toLocaleString('en-US') : 'Not available';
    }
    if (typeof value === 'object') {
      return JSON.stringify(value);
    }
    return String(value);
  }

  function addField(parent, label, value, options = {}) {
    const block = document.createElement('div');
    block.style.padding = '8px 0';
    block.style.borderBottom = '1px solid #e5e7eb';
    addText(block, 'strong', label);
    addText(block, 'p', displayValue(value));
    if (options.scope) {
      addText(block, 'p', 'Scope: ' + options.scope, 'small');
    }
    if (options.notes) {
      addText(block, 'p', options.notes, 'small');
    }
    if (options.source) {
      addLink(block, 'Open source', options.source);
    }
    parent.appendChild(block);
    return block;
  }

  function uniqueStrings(value) {
    const values = Array.isArray(value) ? value : [value];
    return [...new Set(values
      .filter(item => typeof item === 'string')
      .map(item => item.trim())
      .filter(Boolean))];
  }

  function columnNumber(letters) {
    return [...letters].reduce(
      (number, letter) => number * 26 + letter.charCodeAt(0) - 64,
      0
    );
  }

  function changedRow(address) {
    const local = String(address || '')
      .split('!')
      .pop()
      .replace(/\$/g, '')
      .toUpperCase();

    const match = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(local);
    if (!match) return null;

    const firstColumn = columnNumber(match[1]);
    const lastColumn = columnNumber(match[3] || match[1]);

    if (firstColumn > 4 || lastColumn < 4) return null;

    const firstRow = Number(match[2]);
    const lastRow = Number(match[4] || match[2]);
    if (lastRow < FIRST_ROW) return null;

    const start = Math.max(firstRow, FIRST_ROW);
    if (start !== lastRow) return { multiple: true };
    return { row: start, multiple: false };
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

  function validGlobalParentId(id) {
    return /^Q\d+$/.test(String(id || ''));
  }

  function validMalaysiaSaveTarget(candidate, data) {
    if (!candidate || !data || data.version !== VERSION) return false;

    const targetId = String(data.save_target_candidate_id || '');
    const candidateId = String(candidate.candidate_id || '');
    const parentId = String(candidate.global_parent_wikidata_id || '');
    const identity = candidate.record_identity || {};

    return Boolean(
      targetId &&
      candidateId === targetId &&
      candidate.candidate_role === 'MALAYSIA_SAVE_CANDIDATE' &&
      candidate.save_eligible === true &&
      (candidate.record_type === 'malaysia_network' ||
        candidate.record_type === 'malaysia_company') &&
      identity.scope === 'Malaysia' &&
      identity.record_id === candidateId &&
      identity.type === candidate.record_type &&
      candidate.verified_malaysia_presence === true &&
      evidenceItems(candidate).some(item =>
        item?.verified === true &&
        typeof item.source_url === 'string' &&
        item.source_url.startsWith('https://')
      ) &&
      identity.parent_wikidata_id === parentId &&
      identity.parent_reference_id === 'GLOBAL:' + parentId &&
      candidate.global_parent_candidate_id === 'GLOBAL:' + parentId &&
      data.save_policy?.write_from_candidate_id_only === candidateId &&
      data.save_policy?.local_record_only === true &&
      data.save_policy?.reference_only_candidates_must_never_be_saved === true &&
      data.save_policy?.global_parent_is_reference_only === true &&
      data.approval_required === true &&
      data.excel_unchanged === true &&
      validGlobalParentId(parentId)
    );
  }

  function getSingleSaveTarget(data) {
    if (!data || data.version !== VERSION || !Array.isArray(data.candidates)) {
      return null;
    }

    const eligible = data.candidates.filter(candidate =>
      candidate?.save_eligible === true &&
      candidate?.candidate_role === 'MALAYSIA_SAVE_CANDIDATE'
    );

    if (eligible.length !== 1) return null;
    if (data.save_target_candidate_id !== eligible[0].candidate_id) return null;
    return validMalaysiaSaveTarget(eligible[0], data) ? eligible[0] : null;
  }

  function normalizeHeader(value) {
    return String(value ?? '')
      .toLowerCase()
      .replace(/\b(code|symbol)\b/g, '')
      .replace(/[^a-z0-9]/g, '')
      .replace(/s$/, '');
  }

  const HEADER_ALIASES = {
    F: ['Company Headquarter', 'Company Headquarters'],
    L: [
      'Company Revenue (code:MYR symbol:RM)',
      'Company Revenue (MYR)',
      'Company Revenue MYR'
    ],
    M: ['Company Revenue (USD)', 'Company Revenue USD'],
    N: ['Employees (Global)', 'Global Employees'],
    O: ['Employees (MY)', 'Employees Malaysia', 'Malaysia Employees'],
    AF: ['Nearby Rental Location', 'Suggested Rental Location'],
    AG: ['Rent Budget (MYR/mo)', 'Rent Budget MYR/mo', 'Monthly Rent Budget (MYR)'],
    AH: ['Nearest Rail / Transit', 'Nearest Rail Transit', 'Nearest Transit'],
    AI: ['Commute (min)', 'Commute Minutes', 'Commute (minutes)']
  };

  function headerMatches(actual, expected, column) {
    const actualNorm = normalizeHeader(actual);
    const choices = [expected, ...(HEADER_ALIASES[column] || [])]
      .filter(Boolean)
      .map(normalizeHeader);
    return Boolean(actualNorm && choices.includes(actualNorm));
  }

  function normalizeProposedValue(value) {
    if (typeof value === 'number') {
      return Number.isFinite(value) ? value : null;
    }
    if (typeof value === 'string') {
      const cleaned = value.trim();
      return cleaned ? cleaned : null;
    }
    if (typeof value === 'boolean') return value;
    return null;
  }

  function validHttpsSource(value) {
    try {
      const url = new URL(String(value || ''));
      return url.protocol === 'https:';
    } catch {
      return false;
    }
  }

  function hasReportingDate(item) {
    const sourceText = [
      item?.reporting_date,
      item?.as_of,
      item?.date,
      item?.notes
    ].filter(Boolean).join(' ');

    return /(?:as of|reporting date|published|dated|\bon\b)\s*:?\s*(?:\+)?\d{4}(?:-\d{2}(?:-\d{2})?)?/i.test(sourceText);
  }

  function proposedValues(candidate) {
    const rows = Array.isArray(candidate.column_research)
      ? candidate.column_research
      : [];

    const proposed = {};
    const audit = [];

    for (const item of rows) {
      if (!item || !WRITABLE_COLUMNS.includes(item.column)) continue;

      const column = item.column;
      const value = normalizeProposedValue(item.value);
      const isGlobal = GLOBAL_COLUMNS.has(column);
      const expectedScope = isGlobal ? 'global_group_metric' : 'malaysia_company';
      const scopeOK = item.scope === expectedScope;
      const statusOK = ALLOWED_STATUSES.has(item.status);
      const sourceOK = validHttpsSource(item.source_url);
      const dateOK = !REPORTING_DATE_COLUMNS.has(column) || hasReportingDate(item);

      if (value !== null && scopeOK && statusOK && sourceOK && dateOK) {
        proposed[column] = {
          value,
          status: item.status,
          source_url: item.source_url || null,
          scope: item.scope,
          notes: String(item.notes || ''),
          header: String(item.header || '')
        };
      }

      audit.push({
        column,
        header: String(item.header || ''),
        value,
        status: String(item.status || 'unknown'),
        scope: String(item.scope || 'not specified'),
        scope_ok: scopeOK,
        status_ok: statusOK,
        source_ok: sourceOK,
        reporting_date_ok: dateOK,
        source_url: item.source_url || null,
        notes: String(item.notes || '')
      });
    }

    return { proposed, audit };
  }

  function evidenceItems(candidate) {
    const items = candidate?.malaysia_profile?.malaysia_evidence;
    return Array.isArray(items) ? items : [];
  }

  function buildResearchNote(company, candidate, proposed, audit) {
    const p = candidate.malaysia_profile || {};
    const lines = [
      'MALAYSIA COMPANY MASTER — APPROVED RESEARCH NOTE V10.2',
      'Saved at (UTC): ' + new Date().toISOString(),
      'Company entered in D: ' + company,
      'Malaysia record name: ' + String(candidate.name || ''),
      'Malaysia record ID: ' + String(candidate.candidate_id || ''),
      'Record type: ' + String(candidate.record_type || ''),
      'Global parent: ' + String(candidate.global_parent_name || ''),
      'Global parent Wikidata ID: ' + String(candidate.global_parent_wikidata_id || ''),
      'Global parent record is reference-only; no separate global row is saved.',
      '',
      'GLOBAL PARENT METRICS — SAVED ON THIS MALAYSIA COMPANY ROW ONLY'
    ];

    for (const column of ['L', 'M', 'N']) {
      const item = proposed[column];
      if (item) {
        lines.push(
          column + ' — ' + item.header + ': ' + String(item.value) +
          ' [GLOBAL PARENT; ' + item.status + ']' +
          (item.notes ? ' — ' + item.notes : '') +
          (item.source_url ? ' — Source: ' + item.source_url : '')
        );
      } else {
        lines.push(column + ': no supported global-parent figure saved.');
      }
    }

    lines.push('', 'MALAYSIA-SPECIFIC VALUES');
    for (const [column, item] of Object.entries(proposed)) {
      if (GLOBAL_COLUMNS.has(column)) continue;
      lines.push(
        column + ' — ' + item.header + ': ' + String(item.value) +
        ' [' + item.status + ']' +
        (item.notes ? ' — ' + item.notes : '') +
        (item.source_url ? ' — Source: ' + item.source_url : '')
      );
    }

    lines.push('', 'COLUMN RESEARCH AUDIT');
    for (const item of audit) {
      lines.push(
        item.column + ' — ' + item.header +
        ' | status=' + item.status +
        ' | scope=' + item.scope +
        ' | value=' + (item.value === null ? 'not available' : String(item.value)) +
        ' | ' + item.notes +
        (item.source_url ? ' | source=' + item.source_url : '')
      );
    }

    lines.push('', 'MALAYSIA SOURCE EVIDENCE');
    for (const item of evidenceItems(candidate)) {
      lines.push(
        (item.verified === true ? 'CHECKED: ' : 'REVIEW ONLY: ') +
        String(item.type || 'Source') + ' — ' +
        String(item.note || '') +
        (item.source_url ? ' — ' + item.source_url : '')
      );
    }

    lines.push(
      '',
      'IMPORTANT LIMITATIONS',
      'L/M/N are global-parent metrics, not Malaysia subsidiary figures; their scope is explicitly labelled here.',
      'O is Malaysia workforce only. Financial and employee figures without a reporting date were not saved.',
      'Converted revenue is an estimate using the stated historical FX rate, not separately reported revenue.',
      'Visa sponsorship must be supported by explicit employer/job evidence.',
      'Rental amounts and commute estimates must be supported by current, comparable sources.',
      'Unsupported fields were left blank; no values were invented.',
      'Existing Excel values and formulas were not overwritten.'
    );

    return lines.join('\n').slice(0, 30000);
  }

  function renderGlobalProfile(parent, candidate) {
    const g = candidate.global_profile || {};
    const section = panel(parent, 'Global parent — reference information');
    addText(section, 'p', 'This global record is for context only. It is never an Excel save candidate.');
    addField(section, 'Global company type', g.company_type);
    addField(section, 'Global headquarters', g.headquarters);
    addField(section, 'Global industry', g.industry);
    addField(section, 'Country', g.country);
    addField(section, 'Country of origin', g.country_of_origin);
    addField(section, 'Global websites', g.websites);
    const globalId = candidate.global_parent_wikidata_id || candidate.wikidata_id;
    if (validGlobalParentId(globalId)) {
      addLink(section, 'Open global Wikidata record', 'https://www.wikidata.org/wiki/' + globalId);
    }
  }

  function renderMalaysiaProfile(parent, candidate) {
    const p = candidate.malaysia_profile || {};
    const section = panel(parent, 'Malaysia-specific company information');
    addField(section, 'Malaysia presence', p.malaysia_presence_status);
    addField(section, 'Malaysia company / network name', p.malaysia_company_name);
    addField(section, 'Malaysia main office', p.malaysia_headquarters);
    addField(section, 'Malaysia office address', p.malaysia_office_address);
    addField(section, 'Malaysia website', p.malaysia_website);
    addField(section, 'Malaysia careers website', p.malaysia_careers_url);
    addField(section, 'Malaysia industry', p.industry_local);
    addField(section, 'Malaysia employees', p.employees_my, {
      notes: p.employees_my_scope || ''
    });
    addField(section, 'Research note', p.research_note);

    addText(section, 'h4', 'Malaysia source evidence');
    const items = evidenceItems(candidate);
    if (!items.length) {
      addText(section, 'p', 'No Malaysia evidence was supplied.');
    }

    items.forEach(item => {
      const label = (item?.verified === true ? 'Checked: ' : 'Review only: ') +
        String(item?.type || 'Source');
      addText(section, 'strong', label);
      addText(section, 'p', String(item?.note || ''));
      addLink(section, 'View source', item?.source_url);
    });
  }

  function renderColumnPreview(parent, candidate, proposal) {
    const section = panel(parent, '23-column Company Master research preview');
    addText(
      section,
      'p',
      'Only supported values with an allowed source scope can be proposed. Unsupported fields remain blank.'
    );

    const rows = Array.isArray(candidate.column_research)
      ? candidate.column_research
      : [];
    const byColumn = new Map(rows.map(item => [item.column, item]));

    SELECTED_COLUMNS.forEach(column => {
      const item = byColumn.get(column) || {};
      const proposed = proposal.proposed[column];
      const block = document.createElement('div');
      block.style.padding = '9px 0';
      block.style.borderBottom = '1px solid #e5e7eb';

      addText(block, 'strong', column + ' — ' + (item.header || 'Selected column'));

      if (proposed) {
        const scopeLabel = GLOBAL_COLUMNS.has(column)
          ? 'GLOBAL PARENT METRIC'
          : 'MALAYSIA-SPECIFIC';
        addText(block, 'p', displayValue(proposed.value) + '  |  ' + scopeLabel);
        addText(block, 'p', 'Status: ' + proposed.status, 'small');
        if (proposed.notes) addText(block, 'p', proposed.notes, 'small');
        addLink(block, 'Open source', proposed.source_url);
      } else {
        const value = item.value === null || item.value === undefined || item.value === ''
          ? 'No supported value returned — will not be written.'
          : REPORTING_DATE_COLUMNS.has(column) && !hasReportingDate(item)
            ? 'Not eligible to save: reporting date is missing. The agreed rule requires a dated financial or employee figure.'
            : !validHttpsSource(item.source_url)
              ? 'Not eligible to save: no valid HTTPS source URL was supplied.'
              : 'Not eligible to save: status or source scope did not pass the safety checks.';
        addText(block, 'p', value);
        if (item.status) addText(block, 'p', 'Status: ' + item.status, 'small');
        if (item.scope) addText(block, 'p', 'Scope: ' + item.scope, 'small');
        if (item.notes) addText(block, 'p', item.notes, 'small');
        addLink(block, 'Open source', item.source_url);
      }

      section.appendChild(block);
    });

    addText(
      section,
      'p',
      'Protected: A, B, D and AE. C remains unchanged until a dated official Fortune Global 500 source is available.'
    );
  }

  function renderSavePanel(parent, company, row, candidate, data, token, proposal) {
    const section = panel(parent, 'Approval — save to this Malaysia company row');
    addText(
      section,
      'p',
      'Only this Malaysia record can be saved. The global parent remains reference-only, not a separate Excel row.'
    );
    addText(section, 'p', 'Malaysia save ID: ' + candidate.candidate_id);
    addText(section, 'p', 'Global parent reference: ' + candidate.global_parent_candidate_id, 'small');

    addText(
      section,
      'p',
      'L, M and N remain numeric where possible; their GLOBAL parent scope, reporting dates, conversion details and sources are recorded explicitly in AP' + row + '.'
    );

    const save = document.createElement('button');
    save.type = 'button';
    save.textContent = 'Approve & save Malaysia research';
    section.appendChild(save);

    const message = addText(section, 'p', '', 'small');
    const placeholderCheckbox = document.createElement('input');
    let allowPlaceholder = false;

    if (proposal.proposed.J) {
      const label = document.createElement('label');
      label.style.display = 'block';
      label.style.padding = '10px 0';
      placeholderCheckbox.type = 'checkbox';
      placeholderCheckbox.checked = false;
      placeholderCheckbox.style.marginRight = '8px';
      label.appendChild(placeholderCheckbox);
      label.appendChild(document.createTextNode(
        'I explicitly allow replacing the exact text ' + PLACEHOLDER +
        ' in J' + row + ' only. No other existing value may be replaced.'
      ));
      section.appendChild(label);
    }

    const hasValues = Object.keys(proposal.proposed).length > 0;
    if (!hasValues) {
      save.disabled = true;
      message.textContent = 'No supported Malaysia or approved global-parent values are available to save.';
      return;
    }

    save.addEventListener('click', async () => {
      allowPlaceholder = placeholderCheckbox.checked === true;
      await saveMalaysia(
        company,
        row,
        candidate,
        data,
        proposal,
        token,
        allowPlaceholder,
        save,
        message
      );
    });
  }

  function renderLocalCandidate(parent, company, row, candidate, data, token) {
    const proposal = proposedValues(candidate);
    renderGlobalProfile(parent, candidate);
    renderMalaysiaProfile(parent, candidate);
    renderColumnPreview(parent, candidate, proposal);
    renderSavePanel(parent, company, row, candidate, data, token, proposal);
  }

  function renderReferenceCandidate(parent, candidate) {
    renderGlobalProfile(parent, candidate);
    const notice = panel(parent, 'Reference-only record');
    addText(
      notice,
      'p',
      'This candidate cannot be approved or saved. Select the separate Malaysia-specific record to review and save Malaysian operations.'
    );
    addText(notice, 'p', 'Candidate ID: ' + String(candidate.candidate_id || 'Not supplied'));
    addText(notice, 'p', 'Save eligible: no');
  }

  async function saveMalaysia(
    company,
    row,
    candidate,
    data,
    proposal,
    token,
    permitPlaceholder,
    button,
    message
  ) {
    if (token !== requestNumber) {
      message.textContent = 'Research changed. Nothing saved.';
      return;
    }

    if (!validMalaysiaSaveTarget(candidate, data) ||
        data.save_target_candidate_id !== candidate.candidate_id) {
      message.textContent = 'Save blocked: this is not the Worker-designated Malaysia save target.';
      show('SAVE BLOCKED\nOnly the exact Malaysia save_target_candidate_id may be approved.');
      return;
    }

    const note = buildResearchNote(
      company,
      candidate,
      proposal.proposed,
      proposal.audit
    );

    button.disabled = true;
    button.textContent = 'Checking Excel cells...';

    try {
      const outcome = await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem(SHEET);
        const nameCell = sheet.getRange('D' + row);
        nameCell.load('values');

        const headerChecks = Object.entries(proposal.proposed).map(([column, item]) => {
          const range = sheet.getRange(column + HEADER_ROW);
          range.load('values');
          return {
            column,
            expected: item.header || data.column_headers?.[column] || '',
            range
          };
        });

        const columns = [...Object.keys(proposal.proposed), NOTE_COLUMN];
        const cells = columns.map(column => {
          const range = sheet.getRange(column + row);
          range.load('values,formulas');
          return { column, range };
        });

        await context.sync();

        if (token !== requestNumber ||
            String(nameCell.values[0][0] ?? '').trim() !== company) {
          return { blocked: 'Company name or research changed. Nothing saved.' };
        }

        for (const check of headerChecks) {
          const actual = String(check.range.values[0][0] ?? '').trim();
          if (!check.expected || !headerMatches(actual, check.expected, check.column)) {
            return {
              blocked:
                'Header mismatch at ' + check.column + HEADER_ROW +
                '. Expected: "' + check.expected + '"; found: "' + actual + '". Nothing saved.'
            };
          }
        }

        const saved = [];
        const skipped = [];

        for (const cell of cells) {
          const value = cell.range.values[0][0];
          const formula = cell.range.formulas[0][0];
          const hasFormula = typeof formula === 'string' && formula.startsWith('=');
          const isEmpty = (value === null || value === undefined || value === '') && !hasFormula;
          const exactPlaceholder =
            cell.column === 'J' &&
            permitPlaceholder &&
            !hasFormula &&
            value === PLACEHOLDER;

          if (!isEmpty && !exactPlaceholder) {
            skipped.push(cell.column + row);
            continue;
          }

          const valueToWrite = cell.column === NOTE_COLUMN
            ? note
            : proposal.proposed[cell.column].value;

          cell.range.values = [[valueToWrite]];

          // Keep L/M/N numeric for formulas, while displaying GLOBAL in-cell.
          if (GLOBAL_COLUMNS.has(cell.column) && typeof valueToWrite === 'number') {
            const formats = {
              L: '"GLOBAL RM "#,##0.00',
              M: '"GLOBAL USD "#,##0.00',
              N: '"GLOBAL "#,##0'
            };
            cell.range.numberFormat = [[formats[cell.column]]];
          }

          saved.push(cell.column + row);
        }

        if (saved.length) await context.sync();
        return { saved, skipped };
      });

      if (outcome.blocked) {
        show('SAVE BLOCKED\n' + outcome.blocked);
        message.textContent = outcome.blocked;
        button.disabled = false;
        button.textContent = 'Approve & save Malaysia research';
        return;
      }

      show(
        'V10.2 MALAYSIA SAVE COMPLETE\n' +
        'Saved: ' + (outcome.saved.join(', ') || 'none') +
        '\nSkipped (existing values/formulas): ' +
        (outcome.skipped.join(', ') || 'none') +
        '\nL/M/N global-parent scope and sources are documented in AP' + row + '.' +
        '\nOther cells untouched.'
      );

      message.textContent =
        'Saved: ' + (outcome.saved.join(', ') || 'none') +
        '. Skipped: ' + (outcome.skipped.join(', ') || 'none') +
        '. Check AP' + row + ' for the complete research audit.';
      button.textContent = 'Save checked';
    } catch (error) {
      show('SAVE ERROR\n' + errorText(error));
      message.textContent = 'Save not confirmed. Check Excel before retrying.';
      button.disabled = false;
      button.textContent = 'Approve & save Malaysia research';
    }
  }

  function renderCandidates(company, row, data, token) {
    clearResults();
    if (!results) return;

    addText(results, 'h2', 'Review research for D' + row);
    addText(results, 'p', 'Searched: ' + company);
    addText(results, 'p', 'Worker version: ' + data.version);
    addText(results, 'p', 'Save target status: ' + String(data.save_target_status || 'not supplied'));

    const target = getSingleSaveTarget(data);
    if (!target) {
      const warning = panel(results, 'Approval unavailable');
      addText(
        warning,
        'p',
        'The Worker did not return exactly one valid Malaysia save target matching save_target_candidate_id. You may review the results, but nothing can be approved or written.'
      );
    }

    const candidates = Array.isArray(data.candidates) ? data.candidates : [];
    if (!candidates.length) {
      addText(results, 'p', 'No candidates found. Nothing changed in Excel.');
      return;
    }

    candidates.forEach(candidate => {
      const card = document.createElement('article');
      card.className = 'candidate';

      const isTarget = Boolean(target && candidate.candidate_id === target.candidate_id &&
        validMalaysiaSaveTarget(candidate, data));
      const isReference = candidate.candidate_role === 'GLOBAL_REFERENCE_ONLY' ||
        candidate.candidate_role === 'SEARCH_REFERENCE_ONLY' ||
        candidate.save_eligible !== true;

      addText(card, 'h3', candidate.name || 'Unnamed candidate');
      addText(card, 'p', candidate.description || 'No description available.');
      addText(card, 'p', 'Candidate ID: ' + String(candidate.candidate_id || 'Not supplied'), 'small');
      addText(card, 'p', 'Role: ' + String(candidate.candidate_role || 'Unspecified'), 'small');

      if (candidate.global_parent_wikidata_id) {
        addText(card, 'p', 'Global parent Wikidata ID: ' + candidate.global_parent_wikidata_id, 'small');
      } else if (candidate.wikidata_id) {
        addText(card, 'p', 'Wikidata ID: ' + candidate.wikidata_id, 'small');
      }

      if (candidate.source_url) addLink(card, 'Open source record', candidate.source_url);

      if (isTarget) {
        addText(card, 'p', 'This is the only eligible Malaysia save target.', 'small');
      } else if (isReference) {
        addText(card, 'p', 'Reference only — cannot be saved to Excel.', 'small');
      } else {
        addText(card, 'p', 'Not approved as the designated save target.', 'small');
      }

      const select = document.createElement('button');
      select.type = 'button';
      select.textContent = 'Review this record';
      card.appendChild(select);

      const details = document.createElement('div');
      card.appendChild(details);

      select.addEventListener('click', () => {
        if (token !== requestNumber) {
          show('Research expired. Edit the company cell to research again.');
          return;
        }

        select.disabled = true;
        select.textContent = 'Record opened';
        details.replaceChildren();

        if (isTarget) {
          renderLocalCandidate(details, company, row, candidate, data, token);
          show(
            'V10.2 MALAYSIA PREVIEW READY\n' +
            'Malaysia record: ' + candidate.candidate_id +
            '\nGlobal parent is reference-only.' +
            '\nReview the 23 columns before approval.'
          );
        } else {
          renderReferenceCandidate(details, candidate);
          show('REFERENCE ONLY\nThis record cannot be saved to Excel.');
        }
      });

      results.appendChild(card);
    });

    const reject = document.createElement('button');
    reject.type = 'button';
    reject.textContent = 'Reject all candidates';
    reject.addEventListener('click', () => {
      if (token !== requestNumber) return;
      ++requestNumber;
      clearResults();
      show('Candidates rejected. No Excel cells changed.');
    });
    results.appendChild(reject);
  }

  async function research(company, row, token) {
    clearResults();
    show('V10.2 RESEARCHING\nCell: D' + row + '\nCompany: ' + company);

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 25000);

    try {
      const response = await fetch(
        WORKER + '/research?company=' + encodeURIComponent(company),
        {
          method: 'GET',
          headers: { Accept: 'application/json' },
          signal: controller.signal
        }
      );

      if (!response.ok) {
        throw new Error('Research HTTP ' + response.status);
      }

      const data = await response.json();
      if (token !== requestNumber) return;

      if (data.version !== VERSION) {
        throw new Error(
          'Worker version mismatch. Expected V' + VERSION +
          ', received V' + String(data.version || 'unknown') +
          '. Deploy the matching Worker before using this task pane.'
        );
      }

      if (data.status !== 'candidates_for_review') {
        throw new Error(data.message || 'Worker did not return candidates_for_review.');
      }

      const target = getSingleSaveTarget(data);
      show(
        'V10.2 RESEARCH COMPLETE\n' +
        'Candidates: ' + (Array.isArray(data.candidates) ? data.candidates.length : 0) +
        '\nMalaysia save target: ' + (target ? target.candidate_id : 'none / ambiguous') +
        '\nReview the results. Excel has not been changed.'
      );

      renderCandidates(company, row, data, token);
    } catch (error) {
      if (token === requestNumber) {
        clearResults();
        show('RESEARCH ERROR\n' + errorText(error));
      }
    } finally {
      clearTimeout(timeout);
    }
  }

  async function onChanged(event) {
    const location = changedRow(event.address);
    if (!location) return;

    const token = ++requestNumber;
    clearResults();

    if (location.multiple) {
      show('Multiple company rows changed. Please edit/research one company row at a time.');
      return;
    }

    try {
      const company = await readCompany(location.row);
      if (token !== requestNumber) return;

      if (!company) {
        show('D' + location.row + ' is blank.');
        return;
      }

      await research(company, location.row, token);
    } catch (error) {
      if (token === requestNumber) show('CHANGE EVENT ERROR\n' + errorText(error));
    }
  }

  async function initialize() {
    if (active) {
      show('V10.2 READY\nMonitoring Company Master column D, row 7 onward.');
      return;
    }
    if (starting) return;

    starting = true;
    show('V10.2: Connecting to Excel...');

    try {
      await Office.onReady();
      if (typeof Excel === 'undefined') {
        throw new Error('Excel JavaScript API unavailable. Open this add-in inside Excel.');
      }

      await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItemOrNullObject(SHEET);
        sheet.load('isNullObject');
        await context.sync();

        if (sheet.isNullObject) {
          throw new Error('Worksheet "' + SHEET + '" was not found.');
        }

        sheet.onChanged.add(onChanged);
        await context.sync();
      });

      active = true;
      show(
        'V10.2 READY — MALAYSIA COMPANY RESEARCH\n' +
        'Monitoring Company Master column D, row 7 onward.\n' +
        'Edit one company name to research.\n' +
        'Review the Malaysia-specific save target before approval.\n' +
        'L/M/N may contain clearly labelled global-parent metrics.\n' +
        'Existing values and formulas will not be overwritten.'
      );
    } catch (error) {
      show('V10.2 CONNECTION ERROR\n' + errorText(error));
    } finally {
      starting = false;
    }
  }

  if (retry) retry.addEventListener('click', initialize);
  initialize();
})();
