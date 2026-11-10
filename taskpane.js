(() => {
  'use strict';

  // Malaysia Company Research task pane V10.3.
  // Batch-researches all nonblank Company Master rows, then presents one review
  // and one approval. A full Company Master snapshot is created before writes.
  // Only Worker-designated Malaysia-local records can be saved. L/M/N are the
  // sole global-parent-metric exception and are explicitly labelled as GLOBAL.

  const VERSION = '10.3';
  const WORKER =
    'https://malaysia-company-research.rimli271.workers.dev';

  const SHEET = 'Company Master';
  const HEADER_ROW = 6;
  const FIRST_ROW = 7;
  const NOTE_COLUMN = 'AP';
  const BACKUP_PREFIX = 'CMBackup103_';
  const PLACEHOLDER = '[verify office address]';

  const SELECTED_COLUMNS = [
    'A', 'B', 'C', 'D', 'E', 'F', 'G', 'I', 'J', 'K', 'L', 'M',
    'N', 'O', 'P', 'Q', 'R', 'S', 'AE', 'AF', 'AG', 'AH', 'AI'
  ];

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
      'MALAYSIA COMPANY MASTER — APPROVED RESEARCH NOTE V10.3',
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
      'One consolidated review covers all selected columns. Only eligible empty cells can be filled; existing values and formulas are preserved.'
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
          addText(block, 'p', 'Action: eligible to fill after the single batch approval.', 'small');
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

      if (data.version !== VERSION) {
        throw new Error(
          'Worker version mismatch. Task pane expects V' + VERSION +
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

  function renderBatchReview(state) {
    clearResults();
    if (!results) return;

    const { items, snapshot, token } = state;
    addText(results, 'h2', 'Consolidated Company Master review');
    addText(results, 'p', 'Worker version: ' + VERSION);
    addText(results, 'p', 'Company rows researched: ' + snapshot.companies.length);
    addText(results, 'p', 'No Excel cells have been changed. Review every row below before the single approval.');

    let targetCount = 0;
    let proposalCount = 0;

    items.forEach(item => {
      const card = panel(results, 'D' + item.row + ' — ' + item.company);
      if (item.error) {
        addText(card, 'p', 'Research failed: ' + item.error);
        addText(card, 'p', 'No Excel changes will be made for this row.');
        return;
      }

      if (!item.target) {
        addText(card, 'p', 'No unique verified Malaysia save candidate was returned.');
        addText(card, 'p', 'Any global or search candidates are reference-only and cannot be written to Excel.');
        const refs = Array.isArray(item.data?.candidates) ? item.data.candidates : [];
        refs.slice(0, 3).forEach(candidate => {
          addText(card, 'p', String(candidate.name || 'Unnamed candidate') + ' — ' + String(candidate.display_label || 'Reference only'));
          if (candidate.source_url) addLink(card, 'Open reference source', candidate.source_url);
        });
        return;
      }

      targetCount++;
      proposalCount += Object.keys(item.proposal.proposed).length;
      addText(card, 'p', 'Malaysia save candidate: ' + item.target.name);
      addText(card, 'p', 'Malaysia record ID: ' + item.target.candidate_id, 'small');
      addText(card, 'p', 'Global parent: ' + String(item.target.global_parent_name || 'Not supplied') + ' — reference only.', 'small');

      renderGlobalProfile(card, item.target);
      renderMalaysiaProfile(card, item.target);
      renderColumnPreview(card, item.target, item.proposal, item.row, snapshot);

      const rowValues = snapshot.values[item.row - 1] || [];
      const rowFormulas = snapshot.formulas[item.row - 1] || [];
      const apValue = rowValues[columnNumber(NOTE_COLUMN) - 1];
      if (apValue !== null && apValue !== undefined && apValue !== '') {
        addText(card, 'p', 'AP' + item.row + ' already contains a note. It will be preserved; the new research note will not replace it.', 'small');
      } else {
        addText(card, 'p', 'AP' + item.row + ' is empty; the source-aware research audit can be added there.', 'small');
      }

      const jProposal = item.proposal.proposed.J;
      const jValue = rowValues[columnNumber('J') - 1];
      const jFormula = rowFormulas[columnNumber('J') - 1];
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
    });

    const approval = panel(results, 'One consolidated approval');
    addText(
      approval,
      'p',
      'One approval will first create a separate backup worksheet containing the full Company Master snapshot (A1:AP through the last used row), then fill eligible empty cells across all reviewed rows.'
    );
    addText(
      approval,
      'p',
      'Existing values and formulas are skipped. Only the exact J placeholder can be replaced, and only when its row-specific checkbox is selected.'
    );
    addText(
      approval,
      'p',
      'L, M and N are the only global-parent exception. They will be labelled GLOBAL PARENT and their scope, date, source and conversion notes will be recorded in AP where AP is empty.'
    );

    const approve = document.createElement('button');
    approve.type = 'button';
    approve.textContent = 'Approve & save all eligible fields';
    approve.disabled = targetCount === 0 || proposalCount === 0;
    approve.style.display = 'block';
    approve.style.width = '100%';
    approve.style.padding = '13px';
    approve.style.border = '0';
    approve.style.borderRadius = '6px';
    approve.style.background = approve.disabled ? '#9ca3af' : '#137f5b';
    approve.style.color = '#ffffff';
    approve.style.fontWeight = '600';
    approval.appendChild(approve);

    const message = addText(approval, 'p', '', 'small');
    if (approve.disabled) {
      message.textContent = 'Approval unavailable: no verified Malaysia save target with supported fields was returned.';
    } else {
      message.textContent =
        'Verified Malaysia save targets: ' + targetCount +
        '. Supported field proposals: ' + proposalCount +
        '. The workbook remains unchanged until you press this button.';
      approve.addEventListener('click', () => saveBatch(state, approve, message));
    }
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
      show('V10.3 BATCH RESEARCH\nReading Company Master rows…\nExcel has not been changed.');
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
          'V10.3 BATCH RESEARCH\n' +
          'Company ' + (i + 1) + ' of ' + snapshot.companies.length +
          '\nD' + companyRow.row + ': ' + companyRow.company +
          '\nResearching all selected columns. Excel has not been changed.'
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
            proposal: target ? proposedValues(target) : { proposed: {}, audit: [] }
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
        'V10.3 CONSOLIDATED REVIEW READY\n' +
        'Company rows: ' + snapshot.companies.length +
        '\nVerified Malaysia save targets: ' + targets +
        '\nResearch errors: ' + failed +
        '\nReview all rows, then use the single approval button. Excel has not been changed.'
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

  async function saveBatch(state, button, message) {
    if (!state || state.token !== requestNumber || state !== activeBatch) {
      message.textContent = 'Research changed or expired. Nothing saved; run the batch again.';
      return;
    }

    button.disabled = true;
    button.textContent = 'Validating rows and headers…';
    let backupCreatedName = null;

    try {
      const outcome = await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem(SHEET);
        const used = sheet.getUsedRangeOrNullObject(true);
        used.load('isNullObject,rowIndex,rowCount');
        await context.sync();

        const currentLastRow = used.isNullObject
          ? HEADER_ROW
          : Math.max(HEADER_ROW, used.rowIndex + used.rowCount);
        const lastRow = Math.max(state.snapshot.lastRow, currentLastRow);
        const currentRange = sheet.getRange('A1:AP' + lastRow);
        currentRange.load('values,formulas');
        await context.sync();

        if (state.token !== requestNumber || state !== activeBatch) {
          return { blocked: 'Research changed. Nothing saved.' };
        }

        const expectedCompanies = new Map(state.snapshot.companies.map(x => [x.row, x.company]));
        for (let row = FIRST_ROW; row <= lastRow; row++) {
          const expected = expectedCompanies.get(row) || '';
          const actual = currentRowCompany(currentRange.values, row);
          if (actual !== expected) {
            return {
              blocked: 'Company list changed at row ' + row + ' after research. Run Research all company rows again; nothing was saved.'
            };
          }
        }

        const headerChecks = new Map();
        for (const item of state.items) {
          if (!item.target) continue;
          if (!validMalaysiaSaveTarget(item.target, item.data) ||
              item.data.save_target_candidate_id !== item.target.candidate_id) {
            return { blocked: 'Save blocked for D' + item.row + ': Malaysia save-target validation failed.' };
          }
          for (const [column, proposal] of Object.entries(item.proposal.proposed)) {
            const expected = proposal.header || item.data.column_headers?.[column] || '';
            headerChecks.set(column, expected);
          }
        }

        for (const [column, expected] of headerChecks) {
          const actual = String(currentRange.values[HEADER_ROW - 1]?.[columnNumber(column) - 1] ?? '').trim();
          if (!expected || !headerMatches(actual, expected, column)) {
            return {
              blocked: 'Header mismatch at ' + column + HEADER_ROW +
                '. Expected: "' + expected + '"; found: "' + actual + '". Nothing was saved.'
            };
          }
        }

        const writes = [];
        const skipped = [];
        const targetRows = state.items.filter(item => item.target);

        for (const item of targetRows) {
          const row = item.row;
          const rowValues = currentRange.values[row - 1] || [];
          const rowFormulas = currentRange.formulas[row - 1] || [];
          const note = buildResearchNote(
            item.company,
            item.target,
            item.proposal.proposed,
            item.proposal.audit
          );

          for (const [column, proposed] of Object.entries(item.proposal.proposed)) {
            const index = columnNumber(column) - 1;
            const value = rowValues[index];
            const formula = rowFormulas[index];
            const hasFormula = typeof formula === 'string' && formula.startsWith('=');
            const isEmpty = (value === null || value === undefined || value === '') && !hasFormula;
            const permission = state.placeholderPermissions.get(row);
            const exactPlaceholder = column === 'J' &&
              permission?.checked === true && !hasFormula && value === PLACEHOLDER;

            if (!isEmpty && !exactPlaceholder) {
              skipped.push(column + row);
              continue;
            }

            let valueToWrite = proposed.value;
            let numberFormat = null;
            if (GLOBAL_COLUMNS.has(column)) {
              if (typeof valueToWrite === 'number') {
                const formats = {
                  L: '"GLOBAL PARENT RM "#,##0.00',
                  M: '"GLOBAL PARENT USD "#,##0.00',
                  N: '"GLOBAL PARENT EMPLOYEES "#,##0'
                };
                numberFormat = formats[column];
              } else if (typeof valueToWrite === 'string' &&
                  !/^GLOBAL PARENT\b/i.test(valueToWrite)) {
                valueToWrite = 'GLOBAL PARENT — ' + valueToWrite;
              }
            }

            writes.push({ row, column, value: valueToWrite, numberFormat });
          }

          const noteIndex = columnNumber(NOTE_COLUMN) - 1;
          const noteValue = rowValues[noteIndex];
          const noteFormula = rowFormulas[noteIndex];
          const noteHasFormula = typeof noteFormula === 'string' && noteFormula.startsWith('=');
          if ((noteValue === null || noteValue === undefined || noteValue === '') && !noteHasFormula) {
            writes.push({ row, column: NOTE_COLUMN, value: note, numberFormat: null });
          } else {
            skipped.push(NOTE_COLUMN + row);
          }
        }

        if (!writes.length) return { saved: [], skipped, backupName: null };

        const stamp = new Date().toISOString().replace(/\D/g, '').slice(0, 17);
        const backupName = BACKUP_PREFIX + stamp;
        const backupSheet = context.workbook.worksheets.add(backupName);
        const sourceRange = sheet.getRange('A1:AP' + lastRow);
        const backupRange = backupSheet.getRange('A1:AP' + lastRow);
        backupRange.copyFrom(sourceRange, Excel.RangeCopyType.all);
        backupSheet.getRange('AQ1:AV1').values = [[
          'Backup created (UTC)',
          'Source worksheet',
          'Source range',
          'Task pane version',
          'Worker version',
          'Company rows researched'
        ]];
        backupSheet.getRange('AQ2:AV2').values = [[
          new Date().toISOString(),
          SHEET,
          'A1:AP' + lastRow,
          VERSION,
          VERSION,
          state.snapshot.companies.length
        ]];
        backupSheet.getRange('AQ1:AV1').format.font.bold = true;
        await context.sync();
        backupCreatedName = backupName;

        const saved = [];
        for (const write of writes) {
          const cell = sheet.getRange(write.column + write.row);
          cell.values = [[write.value]];
          if (write.numberFormat) cell.numberFormat = [[write.numberFormat]];
          saved.push(write.column + write.row);
        }
        await context.sync();

        return { saved, skipped, backupName, lastRow };
      });

      if (outcome.blocked) {
        show('SAVE BLOCKED\n' + outcome.blocked);
        message.textContent = outcome.blocked;
        button.disabled = false;
        button.textContent = 'Approve & save all eligible fields';
        return;
      }

      if (!outcome.saved.length) {
        show(
          'V10.3 REVIEW COMPLETE — NO CELLS CHANGED\n' +
          'All proposed destination cells were already populated, contained formulas, or were not explicitly approved for placeholder replacement.'
        );
        message.textContent = 'Nothing needed saving, so no backup worksheet was created.';
        button.textContent = 'No changes needed';
        return;
      }

      activeBatch = null;
      show(
        'V10.3 BATCH SAVE COMPLETE\n' +
        'Backup worksheet: ' + outcome.backupName +
        '\nSnapshot: Company Master A1:AP' + outcome.lastRow +
        '\nSaved: ' + outcome.saved.join(', ') +
        '\nSkipped (existing values/formulas): ' + (outcome.skipped.join(', ') || 'none') +
        '\nL/M/N are labelled GLOBAL PARENT; scope, dates and sources are recorded in AP where AP was empty.' +
        '\nOther worksheets and unselected Company Master columns were not written.'
      );
      message.textContent =
        'Saved ' + outcome.saved.length + ' cells. Backup created: ' + outcome.backupName +
        '. Skipped ' + outcome.skipped.length + ' existing/protected cells.';
      button.textContent = 'Batch save complete';
    } catch (error) {
      const backupText = backupCreatedName
        ? ' Backup snapshot exists as ' + backupCreatedName + '.'
        : ' Backup was not confirmed; no save should be assumed.';
      show('BATCH SAVE ERROR\n' + errorText(error) + backupText);
      message.textContent = 'Save not confirmed. Check Company Master and the backup worksheet before retrying.';
      button.disabled = false;
      button.textContent = 'Approve & save all eligible fields';
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
      show('V10.3 READY\nUse Research all company rows for one consolidated review and approval.');
      return;
    }
    if (starting) return;

    starting = true;
    show('V10.3: Connecting to Excel...');

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
        'V10.3 READY — MALAYSIA COMPANY RESEARCH\n' +
        'Use Research all company rows for a consolidated review.\n' +
        'One approval creates a Company Master backup snapshot first.\n' +
        'Only verified Malaysia-local candidates may be saved.\n' +
        'L/M/N are the clearly labelled GLOBAL PARENT exception.\n' +
        'Existing values and formulas are preserved.'
      );
    } catch (error) {
      show('V10.3 CONNECTION ERROR\n' + errorText(error));
    } finally {
      starting = false;
    }
  }

  if (retry) retry.addEventListener('click', initialize);
  ensureBatchButton();
  initialize();
})();
