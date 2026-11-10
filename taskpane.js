(() => {
  'use strict';

  // Malaysia Company Research task pane V10.4 — individual company approval.
  // Researches all nonblank Company Master rows, then presents a separate review
  // and approval button for each company. Each approval writes only that row.
  // A separate Company Master worksheet backup and audit entries are created
  // for each company approval. L/M/N are the only GLOBAL PARENT metric exception.

  const VERSION = '10.4';
  const WORKER_VERSION = '10.3';
  const WORKER =
    'https://malaysia-company-research.rimli271.workers.dev';

  const SHEET = 'Company Master';
  const HEADER_ROW = 6;
  const FIRST_ROW = 7;
  const NOTE_COLUMN = 'AP';
  const AUDIT_SHEET = 'Company Research Audit';
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
  let batchButton = null;
  let activeBatch = null;

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
    if (!candidate || !data || data.version !== WORKER_VERSION) return false;

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
    if (!data || data.version !== WORKER_VERSION || !Array.isArray(data.candidates)) {
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
      'MALAYSIA COMPANY MASTER — APPROVED RESEARCH NOTE V10.4',
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
      'L/M/N are global-parent metrics, not Malaysia subsidiary figures.',
      'O is Malaysia workforce only.',
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

  function renderColumnPreview(parent, candidate, proposal, row, snapshot) {
    const section = panel(parent, '23-column Company Master research preview — row ' + row);
    addText(
      section,
      'p',
      'This company’s review covers all selected columns. Only eligible empty cells can be filled after this company’s individual approval; existing values and formulas are preserved.'
    );

    const rows = Array.isArray(candidate.column_research)
      ? candidate.column_research
      : [];
    const byColumn = new Map(rows.map(item => [item.column, item]));
    const rowValues = snapshot?.values?.[row - 1] || [];
    const rowFormulas = snapshot?.formulas?.[row - 1] || [];

    SELECTED_COLUMNS.forEach(column => {
      const item = byColumn.get(column) || {};
      const proposed = proposal.proposed[column];
      const index = columnNumber(column) - 1;
      const currentValue = rowValues[index];
      const currentFormula = rowFormulas[index];
      const hasFormula = typeof currentFormula === 'string' && currentFormula.startsWith('=');
      const isEmpty = (currentValue === null || currentValue === undefined || currentValue === '') && !hasFormula;
      const exactPlaceholder = column === 'J' && currentValue === PLACEHOLDER && !hasFormula;
      const block = document.createElement('div');
      block.style.padding = '9px 0';
      block.style.borderBottom = '1px solid #e5e7eb';

      const globalLabel = GLOBAL_COLUMNS.has(column)
        ? ' — GLOBAL PARENT / NOT MALAYSIA HEADCOUNT OR SUBSIDIARY REVENUE'
        : '';
      addText(block, 'strong', column + ' — ' + (item.header || 'Selected column') + globalLabel);

      if (proposed) {
        const scopeLabel = GLOBAL_COLUMNS.has(column)
          ? 'GLOBAL PARENT METRIC — NOT MALAYSIA ENTITY'
          : 'MALAYSIA-SPECIFIC';
        addText(block, 'p', displayValue(proposed.value) + '  |  ' + scopeLabel);
        addText(block, 'p', 'Research status: ' + proposed.status, 'small');
        if (proposed.notes) addText(block, 'p', proposed.notes, 'small');
        if (proposed.source_url) addLink(block, 'Open source', proposed.source_url);

        if (isEmpty) {
          addText(block, 'p', 'Action: eligible to fill after this company’s individual approval.', 'small');
        } else if (exactPlaceholder) {
          addText(block, 'p', 'Action: exact placeholder only; requires the separate row-specific checkbox below.', 'small');
        } else {
          addText(block, 'p', 'Action: existing value or formula detected — will be skipped and preserved.', 'small');
          addText(block, 'p', 'Current Excel value: ' + displayValue(currentValue), 'small');
        }
      } else {
        let reason;
        if (item.status === 'protected') {
          reason = 'Protected column — never autofilled.';
        } else if (item.value === null || item.value === undefined || item.value === '') {
          reason = 'No supported value returned — will not be written.';
        } else if (REPORTING_DATE_COLUMNS.has(column) && !hasReportingDate(item)) {
          reason = 'Not eligible to save: reporting date is missing. The dated-figure rule remains in force.';
        } else if (!validHttpsSource(item.source_url)) {
          reason = 'Not eligible to save: no valid HTTPS source URL was supplied.';
        } else {
          reason = 'Not eligible to save: status or source scope did not pass the safety checks.';
        }
        addText(block, 'p', reason);
        if (item.status) addText(block, 'p', 'Research status: ' + item.status, 'small');
        if (item.scope) addText(block, 'p', 'Scope: ' + item.scope, 'small');
        if (item.notes) addText(block, 'p', item.notes, 'small');
        if (item.source_url) addLink(block, 'Open source', item.source_url);
        if (!isEmpty) {

          addText(block, 'p', 'Current Excel value: ' + displayValue(currentValue) + ' — preserved.', 'small');
        }
      }

      section.appendChild(block);
    });

    addText(
      section,
      'p',
      'Protected: A, B, C, D and AE. H, T:AD and AJ:AO are outside this research set. All remain unchanged.'
    );
  }

  function ensureBatchButton() {
    if (batchButton || !results?.parentElement) return;
    batchButton = document.createElement('button');
    batchButton.type = 'button';
    batchButton.textContent = 'Research all company rows';
    batchButton.style.display = 'block';
    batchButton.style.width = '100%';
    batchButton.style.margin = '10px 0';
    batchButton.style.padding = '12px 14px';
    batchButton.style.border = '0';
    batchButton.style.borderRadius = '6px';
    batchButton.style.background = '#137f5b';
    batchButton.style.color = '#ffffff';
    batchButton.style.fontWeight = '600';
    batchButton.style.cursor = 'pointer';
    batchButton.addEventListener('click', researchAllCompanies);
    results.parentElement.insertBefore(batchButton, results);
  }

  async function readCompanyMasterSnapshot() {
    return Excel.run(async context => {
      const sheet = context.workbook.worksheets.getItem(SHEET);
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load('isNullObject,rowIndex,rowCount');
      await context.sync();

      if (used.isNullObject) {
        return { lastRow: HEADER_ROW, values: [], formulas: [], companies: [] };
      }

      const lastRow = Math.max(HEADER_ROW, used.rowIndex + used.rowCount);
      const source = sheet.getRange('A1:AP' + lastRow);
      source.load('values,formulas');
      await context.sync();

      const companies = [];
      for (let row = FIRST_ROW; row <= lastRow; row++) {
        const company = String(source.values[row - 1]?.[columnNumber('D') - 1] ?? '').trim();
        if (!company) continue;
        companies.push({
          row,
          company,
          officeAddressValue: source.values[row - 1]?.[columnNumber('J') - 1],
          officeAddressFormula: source.formulas[row - 1]?.[columnNumber('J') - 1]
        });
      }

      return {
        lastRow,
        values: source.values,
        formulas: source.formulas,
        companies
      };
    });
  }

  async function fetchResearchData(company, token) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 45000);
    try {
      const response = await fetch(
        WORKER + '/research?company=' + encodeURIComponent(company),
        {
          method: 'GET',
          headers: { Accept: 'application/json' },
          cache: 'no-store',
          signal: controller.signal
        }
      );

      if (!response.ok) throw new Error('Research HTTP ' + response.status);
      const data = await response.json();
      if (token !== requestNumber) return null;

      if (data.version !== WORKER_VERSION) {
        throw new Error(
          'Worker version mismatch. Task pane expects V' + WORKER_VERSION +
          ', received V' + String(data.version || 'unknown') + '.'
        );
      }
      if (data.status !== 'candidates_for_review') {
        throw new Error(data.message || 'Worker did not return candidates_for_review.');
      }
      return data;
    } finally {
      clearTimeout(timeout);
    }
  }

  function canWriteCompanyItem(state, item) {
    if (!state || !item?.target || item.saved || item.completed || item.uncertain) {
      return false;
    }

    const rowValues = state.snapshot.values[item.row - 1] || [];
    const rowFormulas = state.snapshot.formulas[item.row - 1] || [];
    const permission = state.placeholderPermissions.get(item.row)?.checked === true;

    for (const column of Object.keys(item.proposal.proposed)) {
      const index = columnNumber(column) - 1;
      const value = rowValues[index];
      const formula = rowFormulas[index];
      const hasFormula = typeof formula === 'string' && formula.startsWith('=');
      const empty = (value === null || value === undefined || value === '') && !hasFormula;
      if (empty) return true;

      if (column === 'J' && permission && value === PLACEHOLDER && !hasFormula) {
        return true;
      }
    }
    return false;
  }

  function updateIndividualApprovalStatus(state) {
    const saved = state.items.filter(item => item.saved).length;
    const pending = state.items.filter(item =>
      item.target && !item.saved && !item.completed && !item.uncertain
    ).length;
    const errors = state.items.filter(item => item.error).length;
    const uncertain = state.items.filter(item => item.uncertain).length;
    const reResearch = state.items.filter(item => item.requiresResearch).length;

    show(
      'V10.4 INDIVIDUAL APPROVAL REVIEW\n' +
      'Companies researched: ' + state.snapshot.companies.length +
      '\nSaved individually: ' + saved +
      '\nAwaiting individual approval: ' + pending +
      '\nResearch errors: ' + errors +
      '\nCompanies requiring fresh research: ' + reResearch +
      '\nSave results requiring workbook inspection: ' + uncertain +
      '\nEach company has its own approval, backup and audit entries. No bulk approval is available.'
    );
  }

  function renderBatchReview(state) {
    clearResults();
    if (!results) return;

    const { items, snapshot } = state;
    addText(results, 'h2', 'Company-by-company review');
    addText(results, 'p', 'Worker version: ' + WORKER_VERSION);
    addText(results, 'p', 'Company rows researched: ' + snapshot.companies.length);
    addText(
      results,
      'p',
      'Research is performed together for convenience, but each company must be approved and saved separately. Approving one company never saves another.'
    );
    addText(
      results,
      'p',
      'Each individual approval creates its own Company Master worksheet backup and its own source-by-source Company Research Audit entries.'
    );

    items.forEach(item => {
      const card = panel(results, 'D' + item.row + ' — ' + item.company);

      if (item.error) {
        addText(card, 'p', 'Research failed: ' + item.error);
        addText(card, 'p', 'No Excel changes have been made for this company.');
        return;
      }

      if (!item.target) {
        addText(card, 'p', 'No unique verified Malaysia save candidate was returned.');
        addText(card, 'p', 'Global and other search candidates are reference-only and cannot be saved.');
        const refs = Array.isArray(item.data?.candidates) ? item.data.candidates : [];
        refs.slice(0, 3).forEach(candidate => {
          addText(
            card,
            'p',
            String(candidate.name || 'Unnamed candidate') + ' — ' +
            String(candidate.display_label || 'Reference only')
          );
          if (candidate.source_url) addLink(card, 'Open reference source', candidate.source_url);
        });
        return;
      }

      addText(card, 'p', 'Malaysia save candidate: ' + item.target.name);
      addText(card, 'p', 'Malaysia record ID: ' + item.target.candidate_id, 'small');
      addText(
        card,
        'p',
        'Global parent: ' + String(item.target.global_parent_name || 'Not supplied') +
        ' — reference only, except for the explicitly labelled L/M/N metrics.',
        'small'
      );

      renderGlobalProfile(card, item.target);
      renderMalaysiaProfile(card, item.target);
      renderColumnPreview(card, item.target, item.proposal, item.row, snapshot);

      const rowValues = snapshot.values[item.row - 1] || [];
      const apValue = rowValues[columnNumber(NOTE_COLUMN) - 1];
      if (apValue !== null && apValue !== undefined && apValue !== '') {
        addText(
          card,
          'p',
          'AP' + item.row + ' already contains a note. It will be preserved; the new research note will not replace it.',
          'small'
        );
      } else {
        addText(
          card,
          'p',
          'AP' + item.row + ' is empty. A research note may be added as part of this company’s individual approval.',
          'small'
        );
      }

      const jProposal = item.proposal.proposed.J;
      const jValue = rowValues[columnNumber('J') - 1];
      const jFormula = (snapshot.formulas[item.row - 1] || [])[columnNumber('J') - 1];
      const jHasFormula = typeof jFormula === 'string' && jFormula.startsWith('=');

      if (jProposal && jValue === PLACEHOLDER && !jHasFormula) {
        const label = document.createElement('label');
        label.style.display = 'block';
        label.style.padding = '10px 0';

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = false;
        checkbox.style.marginRight = '8px';

        label.appendChild(checkbox);
        label.appendChild(document.createTextNode(
          'Allow replacing only the exact text ' + PLACEHOLDER +
          ' in J' + item.row + '. No other existing value may be replaced.'
        ));
        card.appendChild(label);
        state.placeholderPermissions.set(item.row, checkbox);
      }

      const approval = document.createElement('div');
      approval.style.marginTop = '14px';
      approval.style.paddingTop = '12px';
      approval.style.borderTop = '1px solid #cbd5df';
      addText(approval, 'strong', 'Individual approval — ' + item.company);
      addText(
        approval,
        'p',
        'This button saves only ' + item.company + ' in Company Master row ' + item.row +
        '. It will not save or approve any other company.'
      );

      const approve = document.createElement('button');
      approve.type = 'button';
      approve.textContent = 'Approve & save ' + item.company;
      approve.disabled = !canWriteCompanyItem(state, item);
      approve.style.display = 'block';
      approve.style.width = '100%';
      approve.style.padding = '13px';
      approve.style.border = '0';
      approve.style.borderRadius = '6px';
      approve.style.background = approve.disabled ? '#9ca3af' : '#137f5b';
      approve.style.color = '#ffffff';
      approve.style.fontWeight = '600';

      const message = addText(
        approval,
        'p',
        approve.disabled
          ? 'No eligible empty fields are available yet. If J contains the exact placeholder, select its checkbox to enable approval.'
          : 'Ready for this company’s individual approval.',
        'small'
      );

      item.approveButton = approve;
      item.approvalMessage = message;

      approve.addEventListener('click', () => saveCompany(state, item, approve, message));

      const placeholderCheckbox = state.placeholderPermissions.get(item.row);
      if (placeholderCheckbox) {
        placeholderCheckbox.addEventListener('change', () => {
          if (item.saved || item.completed || item.uncertain || item.saving) return;
          approve.disabled = !canWriteCompanyItem(state, item);
          message.textContent = approve.disabled
            ? 'Select the exact-placeholder checkbox only if you want to replace that placeholder.'
            : 'Ready for this company’s individual approval.';
        });
      }

      approval.appendChild(approve);
      card.appendChild(approval);
    });

    updateIndividualApprovalStatus(state);
  }

  async function researchAllCompanies() {
    const token = ++requestNumber;
    activeBatch = null;
    clearResults();
    ensureBatchButton();
    if (batchButton) {
      batchButton.disabled = true;
      batchButton.textContent = 'Researching all company rows…';
    }

    try {
      show('V10.4 RESEARCH ALL COMPANIES\nReading Company Master rows…\nExcel has not been changed.');
      const snapshot = await readCompanyMasterSnapshot();
      if (token !== requestNumber) return;

      if (!snapshot.companies.length) {
        show('No company names found in Company Master column D, row 7 onward.');
        return;
      }

      const items = [];
      for (let i = 0; i < snapshot.companies.length; i++) {
        if (token !== requestNumber) return;
        const companyRow = snapshot.companies[i];
        show(
          'V10.4 RESEARCH ALL COMPANIES\n' +
          'Company ' + (i + 1) + ' of ' + snapshot.companies.length +
          '\nD' + companyRow.row + ': ' + companyRow.company +
          '\nResearching this company. Excel has not been changed.'
        );

        try {
          const data = await fetchResearchData(companyRow.company, token);
          if (token !== requestNumber) return;
          const target = getSingleSaveTarget(data);
          items.push({
            row: companyRow.row,
            company: companyRow.company,
            data,
            target,
            proposal: target ? proposedValues(target) : { proposed: {}, audit: [] },
            saved: false,
            completed: false,
            saving: false,
            uncertain: false,
            requiresResearch: false
          });
        } catch (error) {
          items.push({ row: companyRow.row, company: companyRow.company, error: errorText(error) });
        }
      }

      if (token !== requestNumber) return;
      activeBatch = {
        token,
        snapshot,
        items,
        placeholderPermissions: new Map()
      };
      renderBatchReview(activeBatch);

      const targets = items.filter(item => item.target).length;
      const failed = items.filter(item => item.error).length;
      show(
        'V10.4 INDIVIDUAL APPROVAL REVIEW READY\n' +
        'Company rows: ' + snapshot.companies.length +
        '\nVerified Malaysia save targets: ' + targets +
        '\nResearch errors: ' + failed +
        '\nEach company has its own approval button. Excel has not been changed.'
      );
    } catch (error) {
      if (token === requestNumber) {
        clearResults();
        show('BATCH RESEARCH ERROR\n' + errorText(error));
      }
    } finally {
      if (token === requestNumber && batchButton) {
        batchButton.disabled = false;
        batchButton.textContent = 'Research all company rows';
      }
    }
  }

  function currentRowCompany(values, row) {
    return String(values[row - 1]?.[columnNumber('D') - 1] ?? '').trim();
  }

  const AUDIT_HEADERS = [
    'Run ID', 'Approved UTC', 'Company Master Row', 'Searched Company',
    'Malaysia Record ID', 'Global Parent Wikidata ID', 'Excel Column',
    'Existing Value', 'Proposed Value', 'Save Result', 'Data Scope',
    'Research Status', 'Source URL', 'Research Notes / Reporting Date',
    'Written UTC', 'Column Header'
  ];

  function excelSafe(value) {
    if (value === null || value === undefined) return '';
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    const text = String(value);
    // Prevent sourced text from being evaluated as an Excel formula.
    return /^[\s]*[=+\-@]/.test(text) ? "'" + text : text;
  }

  function sameCell(a, b) {
    return a === b || (a == null && b == null);
  }

  async function saveCompany(state, item, button, message) {
    if (!state || state.token !== requestNumber || state !== activeBatch) {
      message.textContent = 'This review has expired. Nothing was saved; research again.';
      return;
    }
    if (!item || item.saved || item.completed || item.uncertain || item.saving) {
      message.textContent = 'This company is already saved, completed, or awaiting workbook inspection.';
      return;
    }
    if (!item.target || !validMalaysiaSaveTarget(item.target, item.data)) {
      message.textContent = 'The Malaysia save candidate did not pass validation. Nothing was saved.';
      return;
    }
    if (!canWriteCompanyItem(state, item)) {
      message.textContent = 'No eligible empty fields are available for this company. Nothing was saved.';
      button.disabled = true;
      return;
    }

    item.saving = true;
    button.disabled = true;
    button.textContent = 'Saving ' + item.company + ' only…';
    message.textContent = 'Creating this company’s backup and audit. Other company rows will not be written.';

    let backupNameConfirmed = null;
    let auditPrepared = false;

    try {
      const outcome = await Excel.run(async context => {
        const sheets = context.workbook.worksheets;
        const source = sheets.getItem(SHEET);

        const used = source.getUsedRangeOrNullObject(true);
        used.load('isNullObject,rowIndex,rowCount');
        await context.sync();

        const currentLastRow = used.isNullObject
          ? HEADER_ROW
          : Math.max(HEADER_ROW, used.rowIndex + used.rowCount);
        const lastRow = Math.max(state.snapshot.lastRow, currentLastRow);
        const current = source.getRange('A1:AP' + lastRow);
        current.load('values,formulas');
        await context.sync();

        if (state.token !== requestNumber || state !== activeBatch) {
          return { blocked: 'The review expired before approval. Research again.' };
        }

        const row = item.row;
        const currentCompany = currentRowCompany(current.values, row);
        if (currentCompany !== item.company) {
          return {
            blocked: 'Company name in D' + row + ' changed after research. Research this row again.'
          };
        }

        if (!validMalaysiaSaveTarget(item.target, item.data)) {
          return { blocked: 'Malaysia record validation failed for row ' + row + '.' };
        }

        // Validate only this company's proposed headers.
        for (const [column, proposal] of Object.entries(item.proposal.proposed)) {
          const expected = proposal.header || item.data.column_headers?.[column] || '';
          const actual = current.values[HEADER_ROW - 1]?.[columnNumber(column) - 1];
          if (!expected || !headerMatches(actual, expected, column)) {
            return {
              blocked: 'Header mismatch at ' + column + HEADER_ROW +
                '. No cells were written for ' + item.company + '.'
            };
          }
        }

        // Recheck only this company's row. Other companies' rows are not part
        // of this approval and cannot be included in its write list.
        const fieldsToCheck = new Set([...WRITABLE_COLUMNS, NOTE_COLUMN]);
        for (const column of fieldsToCheck) {
          const index = columnNumber(column) - 1;
          if (!sameCell(
                current.values[row - 1]?.[index],
                state.snapshot.values[row - 1]?.[index]
              ) ||
              !sameCell(
                current.formulas[row - 1]?.[index],
                state.snapshot.formulas[row - 1]?.[index]
              )) {
            return {
              blocked: column + row +
                ' changed after preview. Research this company again; nothing was written.'
            };
          }
        }

        // Check the audit worksheet. An existing but empty audit sheet gets
        // its header row; a non-empty sheet must have the exact expected schema.
        const auditExisting = sheets.getItemOrNullObject(AUDIT_SHEET);
        auditExisting.load('isNullObject');
        await context.sync();

        let auditNextRow = 2;
        let writeAuditHeaders = auditExisting.isNullObject;

        if (!auditExisting.isNullObject) {
          const auditUsed = auditExisting.getUsedRangeOrNullObject(true);
          auditUsed.load('isNullObject,rowIndex,rowCount');
          await context.sync();

          if (auditUsed.isNullObject) {
            writeAuditHeaders = true;
            auditNextRow = 2;
          } else {
            const first = auditExisting.getRange('A1:P1');
            first.load('values,formulas');
            await context.sync();

            const validHeaders = AUDIT_HEADERS.every((header, index) =>
              first.values[0][index] === header &&
              !(typeof first.formulas[0][index] === 'string' &&
                first.formulas[0][index].startsWith('='))
            );

            if (!validHeaders) {
              return {
                blocked: 'Existing "' + AUDIT_SHEET +
                  '" has different headers. No Company Master cells were written.'
              };
            }

            auditNextRow = Math.max(
              2,
              auditUsed.rowIndex + auditUsed.rowCount + 1
            );
          }
        }

        const now = new Date().toISOString();
        const runId = 'V104-' + now.replace(/[^0-9]/g, '') + '-R' + row;
        const rowValues = current.values[row - 1] || [];
        const rowFormulas = current.formulas[row - 1] || [];
        const research = new Map(
          (item.target.column_research || []).map(value => [value.column, value])
        );
        const permission = state.placeholderPermissions.get(row)?.checked === true;

        const writes = [];
        const skipped = [];
        const auditRecords = [];

        for (const column of SELECTED_COLUMNS) {

          const index = columnNumber(column) - 1;
          const oldValue = rowValues[index];
          const oldFormula = rowFormulas[index];
          const hasFormula = typeof oldFormula === 'string' && oldFormula.startsWith('=');
          const empty = (oldValue === null || oldValue === undefined || oldValue === '') &&
            !hasFormula;
          const exactPlaceholder = column === 'J' && permission &&
            !hasFormula && oldValue === PLACEHOLDER;
          const proposed = item.proposal.proposed[column];
          const raw = research.get(column) || {};
          let disposition;

          if (!proposed) {
            disposition = 'NOT_PROPOSED';
          } else if (!empty && !exactPlaceholder) {
            disposition = 'SKIPPED_EXISTING';
            skipped.push(column + row);
          } else {
            disposition = 'PREPARED_NOT_CONFIRMED';
            writes.push({
              row,
              column,
              value: proposed.value
            });
          }

          auditRecords.push({
            disposition,
            cells: [
              runId,
              now,
              row,
              item.company,
              item.target.candidate_id,
              item.target.global_parent_wikidata_id,
              column,
              hasFormula ? '[EXISTING FORMULA PROTECTED]' : (oldValue ?? ''),
              proposed ? proposed.value : (raw.value ?? ''),
              disposition,
              proposed?.scope || raw.scope || '',
              proposed?.status || raw.status || '',
              proposed?.source_url || raw.source_url || '',
              proposed?.notes || raw.notes || '',
              '',
              raw.header || proposed?.header || ''
            ].map(excelSafe)
          });
        }

        // AP is an optional summary only. Existing AP notes and formulas are
        // always preserved. Detailed source records are stored in the audit.
        const apIndex = columnNumber(NOTE_COLUMN) - 1;
        const apOld = rowValues[apIndex];
        const apFormula = rowFormulas[apIndex];
        const apEmpty = (apOld === null || apOld === undefined || apOld === '') &&
          !(typeof apFormula === 'string' && apFormula.startsWith('='));

        if (apEmpty) {
          writes.push({
            row,
            column: NOTE_COLUMN,
            value: buildResearchNote(
              item.company,
              item.target,
              item.proposal.proposed,
              item.proposal.audit
            )
          });
        } else {
          skipped.push(NOTE_COLUMN + row);
        }

        if (auditRecords.length > 30000) {
          return { blocked: 'Audit would exceed the safe batch size. Nothing was written.' };
        }

        if (!writes.length) {
          return { noChanges: true, skipped };
        }

        // One worksheet backup for THIS company's approval only.
        const stamp = now.replace(/\D/g, '').slice(0, 17);
        const backupName = 'CM104_' + row + '_' + stamp;
        if (backupName.length > 31) {
          return { blocked: 'Could not create a valid backup worksheet name. Nothing was written.' };
        }

        const collision = sheets.getItemOrNullObject(backupName);
        collision.load('isNullObject');
        await context.sync();

        if (!collision.isNullObject) {
          return {
            blocked: 'This company’s backup name already exists. Nothing was written; run research again.'
          };
        }

        const backup = source.copy(Excel.WorksheetPositionType.after, source);
        backup.name = backupName;
        await context.sync();
        backupNameConfirmed = backupName;

        // Create/append only this company's audit rows before Master writes.
        const audit = auditExisting.isNullObject
          ? sheets.add(AUDIT_SHEET)
          : sheets.getItem(AUDIT_SHEET);

        if (writeAuditHeaders) {
          audit.getRange('A1:P1').values = [AUDIT_HEADERS];
        }

        const lastAuditRow = auditNextRow + auditRecords.length - 1;
        audit.getRange('A' + auditNextRow + ':P' + lastAuditRow).values =
          auditRecords.map(record => record.cells);

        await context.sync();
        auditPrepared = true;

        // Critical isolation: every write below uses this one approved row.
        // No loop over the other researched companies exists in this function.
        for (const write of writes) {
          if (write.row !== row) {
            throw new Error('Safety check failed: attempted write outside approved company row.');
          }
          source.getRange(write.column + write.row).values =
            [[excelSafe(write.value)]];
        }
        await context.sync();

        const savedAt = new Date().toISOString();
        audit.getRange('J' + auditNextRow + ':J' + lastAuditRow).values =
          auditRecords.map(record => [
            record.disposition === 'PREPARED_NOT_CONFIRMED'
              ? 'SAVED'
              : record.disposition
          ]);
        audit.getRange('O' + auditNextRow + ':O' + lastAuditRow).values =
          auditRecords.map(record => [
            record.disposition === 'PREPARED_NOT_CONFIRMED'
              ? savedAt
              : ''
          ]);
        await context.sync();

        return {
          saved: writes.map(write => write.column + write.row),
          skipped,
          backupName,
          auditRows: auditRecords.length
        };
      });

      if (outcome.blocked) {
        item.saving = false;
        item.completed = true;
        item.requiresResearch = true;
        show('SAVE BLOCKED — ' + item.company + '\n' + outcome.blocked);
        message.textContent = outcome.blocked + ' Run Research all company rows again before trying this company.';
        button.disabled = true;
        button.textContent = 'Re-research required — ' + item.company;
        updateIndividualApprovalStatus(state);
        return;
      }

      if (outcome.noChanges) {
        item.saving = false;
        item.completed = true;
        button.disabled = true;
        button.textContent = 'No changes needed — ' + item.company;
        message.textContent =
          'No eligible empty cells were available. Existing values and formulas were preserved. No backup or audit entry was created.';
        updateIndividualApprovalStatus(state);
        return;
      }

      item.saving = false;
      item.saved = true;
      button.disabled = true;
      button.textContent = 'Saved ✓ — ' + item.company;
      message.textContent =
        'Saved only ' + item.company + ' (row ' + item.row + '). ' +
        'Backup: ' + outcome.backupName +
        '. Audit rows: ' + outcome.auditRows +
        '. Cells saved: ' + outcome.saved.join(', ') +
        '. Existing cells skipped: ' + (outcome.skipped.join(', ') || 'none') + '.';

      show(
        'INDIVIDUAL SAVE COMPLETE\n' +
        'Company: ' + item.company + '\n' +
        'Company Master row: ' + item.row + '\n' +
        'Backup: ' + outcome.backupName + '\n' +
        'Audit: ' + AUDIT_SHEET + ' (' + outcome.auditRows + ' records)\n' +
        'Saved cells: ' + outcome.saved.join(', ') + '\n' +
        'Skipped existing cells: ' + (outcome.skipped.join(', ') || 'none') + '\n' +
        'All other company rows remain pending and unchanged.'
      );

      updateIndividualApprovalStatus(state);
    } catch (error) {
      item.saving = false;
      item.uncertain = true;
      button.disabled = true;
      button.textContent = 'Check workbook before retrying — ' + item.company;
      const detail =
        'Save outcome is not confirmed for ' + item.company + '. ' +
        (backupNameConfirmed ? 'Backup exists: ' + backupNameConfirmed + '. ' : '') +
        (auditPrepared
          ? 'Audit was prepared; check for PREPARED_NOT_CONFIRMED entries.'
          : 'Audit preparation was not confirmed.') +
        ' Inspect this company row, its backup and Company Research Audit before retrying. ' +
        errorText(error);

      show('SAVE NOT CONFIRMED\n' + detail);
      message.textContent = detail;
      updateIndividualApprovalStatus(state);
    }
  }

  function onChanged(event) {
    const location = changedRow(event.address);
    if (!location) return;

    ++requestNumber;
    activeBatch = null;
    clearResults();
    if (location.multiple) {
      show('Company names changed in multiple rows. Click Research all company rows to prepare a fresh consolidated review.');
    } else {
      show('D' + location.row + ' changed. Click Research all company rows to refresh the consolidated review. No Excel cells were changed by research.');
    }
  }

  async function initialize() {
    if (active) {
      ensureBatchButton();
      show('V10.4 READY\nResearch all company rows, then approve each company separately.');
      return;
    }
    if (starting) return;

    starting = true;
    show('V10.4: Connecting to Excel...');

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
      ensureBatchButton();
      show(
        'V10.4 READY — MALAYSIA COMPANY RESEARCH\n' +
        'Research all company rows, then approve each company separately.\n' +
        'Each company approval creates its own worksheet backup before writing.\n' +
        'Only verified Malaysia-local candidates may be saved.\n' +
        'L/M/N are GLOBAL PARENT figures; their sources go to the audit.\n' +
        'Existing values, formulas, number formats and AP notes are preserved.'
      );
    } catch (error) {
      show('V10.4 CONNECTION ERROR\n' + errorText(error));
    } finally {
      starting = false;
    }
  }

  if (retry) retry.addEventListener('click', initialize);
  ensureBatchButton();
  initialize();
})();
