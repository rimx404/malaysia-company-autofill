(() => {
  'use strict';

  // Malaysia Company Detector V10.5.2
  // Research many companies, approve each eligible field without Excel writes,
  // then save approved fields for ONE company after a backup and an audit.
  // Company Master headers, formats, formulas and existing values are protected.

  const VERSION = '10.5.2';
  const WORKER_VERSION = '10.3';
  const WORKER = 'https://malaysia-company-research.rimli271.workers.dev';
  const MASTER = 'Company Master';
  const AUDIT = 'Company Research Audit';
  const FIRST_ROW = 7;
  const HEADER_ROW = 6;
  const PLACEHOLDER = '[verify office address]';

  const COLUMNS = 'A B C D E F G I J K L M N O P Q R S AE AF AG AH AI'.split(' ');
  const WRITABLE = new Set('E F G I J K L M N O P Q R S AF AG AH AI'.split(' '));
  const GLOBAL = new Set(['L', 'M', 'N']);
  const DATED = new Set(['L', 'M', 'N', 'O']);

  const STATUSES = new Set([
    'official', 'official_lower_bound', 'reported_needs_review',
    'derived_fx_estimate', 'review_required'
  ]);

  const AUDIT_HEADERS = [
    'Run ID', 'Approved UTC', 'Company Master Row', 'Searched Company',
    'Malaysia Record ID', 'Global Parent Wikidata ID', 'Excel Column',
    'Existing Value', 'Proposed Value', 'Save Result', 'Data Scope',
    'Research Status', 'Source URL', 'Research Notes / Reporting Date',
    'Written UTC', 'Column Header'
  ];

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
    AG: [
      'Rent Budget (MYR/mo)',
      'Rent Budget MYR/mo',
      'Monthly Rent Budget (MYR)'
    ],
    AH: ['Nearest Rail / Transit', 'Nearest Rail Transit', 'Nearest Transit'],
    AI: ['Commute (min)', 'Commute Minutes', 'Commute (minutes)']
  };

  const status = document.getElementById('status');
  const results = document.getElementById('results');
  const retry = document.getElementById('retry');

  let initialized = false;
  let initializing = false;
  let researchButton = null;
  let generation = 0;
  let currentReview = null;

  function say(message) {
    if (status) status.textContent = message;
  }

  function add(parent, tag, content, className) {
    const element = document.createElement(tag);
    if (content !== undefined) {
      element.textContent = String(content);
    }
    if (className) element.className = className;
    parent.appendChild(element);
    return element;
  }

  function panel(parent, heading) {
    const card = add(parent, 'section');
    card.style.padding = '12px';
    card.style.margin = '14px 0';
    card.style.border = '1px solid #cbd5df';
    card.style.borderRadius = '8px';
    add(card, 'h3', heading);
    return card;
  }

  function link(parent, label, address) {
    if (typeof address !== 'string') return;

    try {
      const url = new URL(address);
      if (url.protocol !== 'https:') return;

      const anchor = add(parent, 'a', label);
      anchor.href = url.href;
      anchor.target = '_blank';
      anchor.rel = 'noopener noreferrer';
    } catch {
      // Missing or invalid links are not displayed.
    }
  }

  function display(value) {
    if (value === undefined || value === null || value === '') {
      return 'Not available';
    }
    if (Array.isArray(value)) {
      return value.map(String).join('; ') || 'Not available';
    }
    if (typeof value === 'number') {
      return value.toLocaleString('en-US');
    }
    if (typeof value === 'object') {
      return JSON.stringify(value);
    }
    return String(value);
  }

  function field(parent, label, value, url) {
    const block = add(parent, 'div');
    block.style.margin = '8px 0';
    add(block, 'strong', label);
    add(block, 'p', display(value));
    link(block, 'Source', url);
  }

  function indexOfColumn(column) {
    return [...column].reduce(
      (n, c) => n * 26 + c.charCodeAt(0) - 64,
      0
    ) - 1;
  }

  function same(a, b) {
    return a === b || (a == null && b == null);
  }

  function safeExcel(value) {
    if (value === null || value === undefined) return '';
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    if (typeof value === 'boolean') return value;

    const text = String(value);

    // Prevent sourced strings from being interpreted as formulas.
    return /^\s*[=+\-@]/.test(text) ? "'" + text : text;
  }

  function https(url) {
    try {
      return new URL(String(url || '')).protocol === 'https:';
    } catch {
      return false;
    }
  }

  function normalizeHeader(value) {
    return String(value || '')
      .toLowerCase()
      .replace(/\b(code|symbol)\b/g, '')
      .replace(/[^a-z0-9]/g, '')
      .replace(/s$/, '');
  }

  function headerMatches(actual, expected, column) {
    const normalized = normalizeHeader(actual);

    return Boolean(
      normalized &&
      [expected, ...(HEADER_ALIASES[column] || [])]
        .some(value =>
          value && normalizeHeader(value) === normalized
        )
    );
  }

  function hasReportingDate(item) {
    const text = [
      item?.reporting_date,
      item?.as_of,
      item?.date,
      item?.notes
    ].filter(Boolean).join(' ');

    return /(?:as of|reporting date|published|dated|\bon\b)\s*:?\s*\+?\d{4}(?:-\d{2})?(?:-\d{2})?/i
      .test(text);
  }

  function evidenceItems(candidate) {
    const evidence =
      candidate?.malaysia_profile?.malaysia_evidence;

    return Array.isArray(evidence) ? evidence : [];
  }

  // ==================================================
  // VERIFY THE MALAYSIA SAVE TARGET
  // ==================================================

  function verifiedMalaysiaTarget(data) {
    if (
      !data ||
      data.version !== WORKER_VERSION ||
      !Array.isArray(data.candidates)
    ) {
      return null;
    }

    const eligible = data.candidates.filter(
      candidate =>
        candidate?.save_eligible === true &&
        candidate?.candidate_role === 'MALAYSIA_SAVE_CANDIDATE'
    );

    if (eligible.length !== 1) return null;

    const candidate = eligible[0];
    const id = candidate.candidate_id;
    const parent = candidate.global_parent_wikidata_id;
    const identity = candidate.record_identity || {};
    const policy = data.save_policy || {};

    if (
      !id ||
      data.save_target_candidate_id !== id ||
      !/^Q\d+$/.test(String(parent || '')) ||

      !['malaysia_company', 'malaysia_network']
        .includes(candidate.record_type) ||

      candidate.verified_malaysia_presence !== true ||

      identity.record_id !== id ||
      identity.scope !== 'Malaysia' ||
      identity.type !== candidate.record_type ||
      identity.parent_wikidata_id !== parent ||
      identity.parent_reference_id !== 'GLOBAL:' + parent ||

      candidate.global_parent_candidate_id !== 'GLOBAL:' + parent ||

      policy.write_from_candidate_id_only !== id ||
      policy.local_record_only !== true ||
      policy.reference_only_candidates_must_never_be_saved !== true ||
      policy.global_parent_is_reference_only !== true ||

      data.approval_required !== true ||
      data.excel_unchanged !== true ||

      !evidenceItems(candidate).some(
        e => e.verified === true && https(e.source_url)
      )
    ) {
      return null;
    }

    return candidate;
  }

  function buildProposals(candidate) {
    const raw = new Map();
    const proposed = new Map();

    for (const record of candidate.column_research || []) {
      if (!record || !COLUMNS.includes(record.column)) continue;

      raw.set(record.column, record);

      const v = record.value;

      const validValue =
        (
          typeof v === 'number' &&
          Number.isFinite(v)
        ) ||
        (
          typeof v === 'string' &&
          Boolean(v.trim())
        );

      const scope = GLOBAL.has(record.column)
        ? 'global_group_metric'
        : 'malaysia_company';

      if (
        WRITABLE.has(record.column) &&
        validValue &&
        record.scope === scope &&
        STATUSES.has(record.status) &&
        https(record.source_url) &&
        (
          !DATED.has(record.column) ||
          hasReportingDate(record)
        )
      ) {
        proposed.set(record.column, record);
      }
    }

    return { raw, proposed };
  }

  function rowValues(state, item) {
    return state.snapshot.values[item.row - 1] || [];
  }

  function rowFormulas(state, item) {
    return state.snapshot.formulas[item.row - 1] || [];
  }

  // ==================================================
  // FIELD APPROVAL RULES
  // ==================================================

  function canApprove(state, item, column) {
    if (
      !state ||
      !item.target ||
      !WRITABLE.has(column) ||
      item.saved ||
      item.saving ||
      item.blocked ||
      item.uncertain ||
      !item.proposal?.proposed.has(column)
    ) {
      return false;
    }

    const index = indexOfColumn(column);
    const value = rowValues(state, item)[index];
    const formula = rowFormulas(state, item)[index];

    if (
      typeof formula === 'string' &&
      formula.startsWith('=')
    ) {
      return false;
    }

    if (
      value === '' ||
      value === null ||
      value === undefined
    ) {
      return true;
    }

    return (
      column === 'J' &&
      value === PLACEHOLDER &&
      item.allowJ === true
    );
  }

  function approvedColumns(state, item) {
    return [...item.approved].filter(
      column => canApprove(state, item, column)
    );
  }

  // FIX: Recalculate the enabled state AND color of
  // all approval and Save buttons.

  function refreshButtons(state, item) {
    for (const [column, button] of item.buttons) {
      const selected = item.approved.has(column);

      button.textContent = selected
        ? 'Approved ✓ (undo)'
        : 'Approve';

      button.disabled = !canApprove(state, item, column);

      button.style.background = button.disabled
        ? '#9ca3af'
        : selected
          ? '#e4f3e9'
          : '#137f5b';

      button.style.color = button.disabled
        ? '#ffffff'
        : selected
          ? '#126142'
          : '#ffffff';

      button.style.cursor = button.disabled
        ? 'not-allowed'
        : 'pointer';

      button.style.opacity = button.disabled
        ? '0.75'
        : '1';

      button.setAttribute(
        'aria-disabled',
        String(button.disabled)
      );
    }

    if (!item.saveButton) return;

    const count = approvedColumns(state, item).length;

    const enabled =
      count > 0 &&
      !item.saved &&
      !item.saving &&
      !item.blocked &&
      !item.uncertain;

    item.saveButton.disabled = !enabled;

    item.saveButton.style.background = enabled
      ? '#137f5b'
      : '#9ca3af';

    item.saveButton.style.color = '#ffffff';

    item.saveButton.style.cursor = enabled
      ? 'pointer'
      : 'not-allowed';

    item.saveButton.style.opacity = enabled
      ? '1'
      : '0.75';

    item.saveButton.setAttribute(
      'aria-disabled',
      String(!enabled)
    );

    if (
      item.saved ||
      item.uncertain ||
      item.blocked ||
      item.saving
    ) {
      return;
    }

    item.saveMessage.textContent = count
      ? count +
        ' column(s) approved. Nothing has been written. Save only when ready.'
      : 'No approved columns yet. Approve an eligible column above first.';
  }

  // ==================================================
  // RESEARCH PROFILE DISPLAY
  // ==================================================

  function showProfiles(card, target) {
    const p = target.malaysia_profile || {};
    const g = target.global_profile || {};

    field(card, 'Malaysia save record', target.name);
    field(card, 'Malaysia record ID', target.candidate_id);

    field(
      card,
      'Global parent — reference only',
      target.global_parent_name
    );

    field(
      card,
      'Malaysia presence',
      p.malaysia_presence_status
    );

    field(
      card,
      'Malaysia company/network',
      p.malaysia_company_name
    );

    field(
      card,
      'Malaysia head office',
      p.malaysia_headquarters
    );

    field(
      card,
      'Malaysia main office address',
      p.malaysia_office_address
    );

    field(
      card,
      'Malaysia website',
      p.malaysia_website,
      p.malaysia_website
    );

    field(
      card,
      'Malaysia careers',
      p.malaysia_careers_url,
      p.malaysia_careers_url
    );

    field(
      card,
      'Malaysia industry',
      p.industry_local
    );

    field(
      card,
      'Malaysia workforce — display only until dated',
      p.employees_my
    );

    field(
      card,
      'Global headquarters — reference only',
      g.headquarters
    );

    field(
      card,
      'Global industry — reference only',
      g.industry
    );

    field(
      card,
      'Global websites — reference only',
      g.websites
    );

    for (const e of evidenceItems(target)) {
      add(
        card,
        'p',
        (
          e.verified === true
            ? 'Verified evidence: '
            : 'Review evidence: '
        ) +
        String(e.note || e.type || 'Source')
      );

      link(card, 'View source', e.source_url);
    }
  }

  // ==================================================
  // SHOW ALL REQUESTED COLUMN BUTTONS
  // ==================================================

  function showColumn(card, state, item, column) {
    const record =
      item.proposal.raw.get(column) || {};

    const proposal =
      item.proposal.proposed.get(column);

    // Display the exact existing Excel header.
    // No header is renamed or overwritten.
    const block = panel(
      card,
      column + ' — ' +
      display(
        state.snapshot.values[HEADER_ROW - 1]?.[
          indexOfColumn(column)
        ]
      )
    );

    if (GLOBAL.has(column)) {
      add(
        block,
        'p',
        'GLOBAL PARENT metric, not Malaysia entity data.'
      );
    }

    const old = rowValues(state, item)[
      indexOfColumn(column)
    ];

    const formula = rowFormulas(state, item)[
      indexOfColumn(column)
    ];

    if (proposal) {
      field(
        block,
        'Research proposal',
        proposal.value,
        proposal.source_url
      );

      add(
        block,
        'p',
        'Status: ' + proposal.status
      );

      if (proposal.notes) {
        add(block, 'p', proposal.notes);
      }

    } else {
      let reason =
        'No supported Worker value; nothing can be approved.';

      if (!WRITABLE.has(column)) {
        reason =
          'Protected column — cannot be saved by this add-in.';

      } else if (
        DATED.has(column) &&
        record.value != null &&
        !hasReportingDate(record)
      ) {
        reason =
          'Reporting date missing; cannot approve.';

      } else if (
        record.value != null &&
        record.value !== '' &&
        !https(record.source_url)
      ) {
        reason =
          'No HTTPS evidence source; cannot approve.';
      }

      add(block, 'p', reason);

      if (record.notes) {
        add(block, 'p', record.notes);
      }

      link(
        block,
        'Research source',
        record.source_url
      );
    }

    if (
      old !== '' &&
      old != null
    ) {
      add(
        block,
        'p',
        'Existing Excel value (preserved): ' +
        display(old)
      );
    }

    if (
      typeof formula === 'string' &&
      formula.startsWith('=')
    ) {
      add(
        block,
        'p',
        'Existing formula (protected).'
      );
    }

    // Office address J is the only permitted
    // existing-value exception.

    const jPlaceholder =
      column === 'J' &&
      proposal &&
      old === PLACEHOLDER &&
      !(
        typeof formula === 'string' &&
        formula.startsWith('=')
      );

    if (jPlaceholder) {
      const label = add(block, 'label');

      const checkbox = add(label, 'input');
      checkbox.type = 'checkbox';
      checkbox.checked = false;

      add(
        label,
        'span',
        ' Permit replacing only [verify office address] in J' +
        item.row
      );

      checkbox.addEventListener(
        'change',
        () => {
          item.allowJ = checkbox.checked;

          if (!checkbox.checked) {
            item.approved.delete('J');
          }

          refreshButtons(state, item);
        }
      );
    }

    // All requested research fields show an Approve button:
    // G I O P Q R S AF AG AH AI (and other writable fields).
    //
    // No value or existing data => button visible but DISABLED.
    // AE Tier => button visible but ALWAYS DISABLED.
    // A B C D => no button.

    if (
      !WRITABLE.has(column) &&
      column !== 'AE'
    ) {
      return;
    }

    if (!canApprove(state, item, column)) {
      let reason =
        'Approve is unavailable: no verified, writable proposal.';

      if (column === 'AE') {
        reason =
          'Tier AE is protected; approval is review-only and disabled.';

      } else if (
        proposal &&
        typeof formula === 'string' &&
        formula.startsWith('=')
      ) {
        reason =
          'Approve is disabled because this cell contains a formula.';

      } else if (
        proposal &&
        old !== '' &&
        old != null &&
        !jPlaceholder
      ) {
        reason =
          'Approve is disabled because an existing Excel value must be preserved.';

      } else if (
        jPlaceholder &&
        !item.allowJ
      ) {
        reason =
          'Tick the separate J placeholder checkbox to enable Approve.';
      }

      add(block, 'p', reason);
    }

    const button = add(
      block,
      'button',
      'Approve'
    );

    button.type = 'button';

    button.style.padding = '9px 14px';
    button.style.border = '0';
    button.style.borderRadius = '6px';
    button.style.fontWeight = '600';

    button.disabled =
      !canApprove(state, item, column);

    // Approve only changes the local selection.
    // This handler NEVER writes to Excel.

    button.addEventListener(
      'click',
      () => {
        if (!canApprove(state, item, column)) {
          return;
        }

        if (item.approved.has(column)) {
          item.approved.delete(column);
        } else {
          item.approved.add(column);
        }

        refreshButtons(state, item);
      }
    );

    item.buttons.set(column, button);
  }

  function renderReview(state) {
    if (!results) return;

    results.replaceChildren();

    add(
      results,
      'h2',
      'Individual column approvals — V' + VERSION
    );

    add(
      results,
      'p',
      'Research does not change Excel. Approve fields one at a time, then use the separate Save button for that ONE company.'
    );

    add(
      results,
      'p',
      'A/B/C/D have no Approve button. AE has a disabled, review-only button. Missing Fortune, visa, rental and commute data need verified Worker research.'
    );

    state.items.forEach(item => {
      const card = panel(
        results,
        'D' + item.row + ' — ' + item.name
      );

      if (item.error) {
        add(
          card,
          'p',
          'Research failed: ' + item.error
        );
        return;
      }

      if (!item.target) {
        add(
          card,
          'p',
          'No unique verified Malaysia save candidate. Nothing may be written for this row.'
        );
        return;
      }

      showProfiles(card, item.target);

      add(
        card,
        'h4',
        'Review the existing 23 Company Master columns'
      );

      COLUMNS.forEach(
        column => showColumn(
          card,
          state,
          item,
          column
        )
      );

      const ap = rowValues(state, item)[
        indexOfColumn('AP')
      ];

      add(
        card,
        'p',
        ap == null || ap === ''
          ? 'AP is blank; a source note may be added on save.'
          : 'AP already contains information and will never be overwritten.'
      );

      const saveBox = panel(
        card,
        'Save for ' + item.name + ' ONLY'
      );

      add(
        saveBox,
        'p',
        'Saving first copies the full Company Master worksheet, then appends audit entries before writing approved values.'
      );

      const button = add(
        saveBox,
        'button',
        'Save approved columns — ' + item.name
      );

      button.type = 'button';
      button.style.width = '100%';
      button.style.padding = '13px';
      button.style.border = '0';
      button.style.borderRadius = '6px';
      button.style.fontWeight = '600';

      item.saveButton = button;
      item.saveMessage = add(saveBox, 'p');

      button.addEventListener(
        'click',
        () => saveCompany(state, item)
      );

      refreshButtons(state, item);
    });
  }

  // ==================================================
  // RESEARCH COMPANY MASTER
  // ==================================================

  function ensureResearchButton() {
    if (
      researchButton ||
      !results?.parentElement
    ) {
      return;
    }

    researchButton =
      document.createElement('button');

    researchButton.type = 'button';
    researchButton.textContent = 'Research all company rows';

    researchButton.style.display = 'block';
    researchButton.style.width = '100%';
    researchButton.style.margin = '10px 0';
    researchButton.style.padding = '12px';
    researchButton.style.background = '#137f5b';
    researchButton.style.color = 'white';
    researchButton.style.border = '0';
    researchButton.style.borderRadius = '6px';

    researchButton.addEventListener(
      'click',
      researchAll
    );

    results.parentElement.insertBefore(
      researchButton,
      results
    );
  }

  async function snapshot() {
    return Excel.run(async context => {
      const sheet =
        context.workbook.worksheets.getItem(MASTER);

      const used =
        sheet.getUsedRangeOrNullObject(true);

      used.load(
        'isNullObject,rowIndex,rowCount'
      );

      await context.sync();

      const lastRow = used.isNullObject
        ? HEADER_ROW
        : Math.max(
            HEADER_ROW,
            used.rowIndex + used.rowCount
          );

      const range =
        sheet.getRange('A1:AP' + lastRow);

      range.load('values,formulas');

      await context.sync();

      const companies = [];

      for (
        let row = FIRST_ROW;
        row <= lastRow;
        row++
      ) {
        const name = String(
          range.values[row - 1]?.[
            indexOfColumn('D')
          ] ?? ''
        ).trim();

        if (name) {
          companies.push({
            row,
            name
          });
        }
      }

      return {
        values: range.values,
        formulas: range.formulas,
        companies,
        lastRow
      };
    });
  }

  async function fetchResearch(name, token) {
    const controller = new AbortController();

    const timeout = setTimeout(
      () => controller.abort(),
      45000
    );

    try {
      const response = await fetch(
        WORKER +
        '/research?company=' +
        encodeURIComponent(name),
        {
          headers: {
            Accept: 'application/json'
          },
          cache: 'no-store',
          signal: controller.signal
        }
      );

      if (!response.ok) {
        throw Error(
          'Research HTTP ' +
          response.status
        );
      }

      const data = await response.json();

      if (token !== generation) {
        return null;
      }

      if (data.version !== WORKER_VERSION) {
        throw Error(
          'Worker version must be 10.3.'
        );
      }

      if (
        data.status !== 'candidates_for_review'
      ) {
        throw Error(
          data.message ||
          'No review candidates.'
        );
      }

      return data;
    } finally {
      clearTimeout(timeout);
    }
  }

  async function researchAll() {
    const token = ++generation;
    currentReview = null;

    if (results) {
      results.replaceChildren();
    }

    if (researchButton) {
      researchButton.disabled = true;
    }

    try {
      say(
        'V' + VERSION +
        ': Reading Company Master. Research never writes Excel.'
      );

      const original = await snapshot();

      if (token !== generation) return;

      if (!original.companies.length) {
        say(
          'No companies in column D, row 7 onward.'
        );
        return;
      }

      const items = [];

      for (const company of original.companies) {
        if (token !== generation) return;

        say(
          'Researching D' +
          company.row +
          ' — ' +
          company.name +
          '. Excel unchanged.'
        );

        const item = {
          row: company.row,
          name: company.name,

          approved: new Set(),
          buttons: new Map(),

          allowJ: false,
          saved: false,
          saving: false,
          uncertain: false,
          blocked: false
        };

        try {
          const data = await fetchResearch(
            company.name,
            token
          );

          if (token !== generation) return;

          item.data = data;

          item.target = verifiedMalaysiaTarget(data);

          item.proposal = item.target
            ? buildProposals(item.target)
            : {
                raw: new Map(),
                proposed: new Map()
              };

        } catch (error) {
          item.error =
            String(error?.message || error);
        }

        items.push(item);
      }

      currentReview = {
        token,
        snapshot: original,
        items
      };

      renderReview(currentReview);

      say(
        'V' +
        VERSION +
        ' REVIEW READY: ' +
        items.length +
        ' companies researched. Approve fields individually. Excel unchanged.'
      );

    } catch (error) {
      say(
        'RESEARCH ERROR: ' +
        String(error?.message || error)
      );

    } finally {
      if (
        token === generation &&
        researchButton
      ) {
        researchButton.disabled = false;
      }
    }
  }

  // ==================================================
  // SOURCE NOTES
  // ==================================================

  function noteFor(item, selected) {
    const lines = [
      'MALAYSIA RESEARCH V' + VERSION,

      'Company Master row: ' + item.row,

      'Company: ' + item.name,

      'Malaysia record: ' +
      item.target.name +
      ' (' +
      item.target.candidate_id +
      ')',

      'Global parent: ' +
      String(item.target.global_parent_name || ''),

      'ONLY THESE FIELDS WERE APPROVED:'
    ];

    selected.forEach(column => {
      const p =
        item.proposal.proposed.get(column);

      lines.push(
        column +
        ' — ' +
        String(p.header || '') +
        ': ' +
        String(p.value) +
        ' | ' +
        (
          GLOBAL.has(column)
            ? 'GLOBAL PARENT'
            : 'MALAYSIA'
        ) +
        ' | ' +
        p.status +
        ' | ' +
        String(p.notes || '') +
        ' | ' +
        p.source_url
      );
    });

    lines.push(
      'Unapproved columns and existing data were not overwritten.'
    );

    return lines.join('\n').slice(0, 30000);
  }

  // ==================================================
  // SAVE ONE COMPANY'S APPROVED FIELDS ONLY
  // ==================================================

  async function saveCompany(state, item) {
    if (
      state !== currentReview ||
      state.token !== generation ||
      !item.target ||
      item.saved ||
      item.saving ||
      item.uncertain ||
      item.blocked
    ) {
      return;
    }

    if (
      verifiedMalaysiaTarget(item.data) !== item.target
    ) {
      item.saveMessage.textContent =
        'Malaysia candidate validation failed. No changes.';
      return;
    }

    const selected =
      approvedColumns(state, item);

    if (!selected.length) {
      refreshButtons(state, item);
      return;
    }

    item.saving = true;
    refreshButtons(state, item);

    item.saveMessage.textContent =
      'Checking this company, creating backup and audit…';

    let backupConfirmed = null;
    let auditConfirmed = false;

    try {
      const result = await Excel.run(async context => {
        const sheets = context.workbook.worksheets;

        const master =
          sheets.getItem(MASTER);

        const current = master.getRange(
          'A6:AP' + item.row
        );

        current.load('values,formulas');

        await context.sync();

        if (
          state !== currentReview ||
          state.token !== generation
        ) {
          return {
            blocked:
              'Review expired. Research again.'
          };
        }

        const headers = current.values[0];

        const values =
          current.values[item.row - HEADER_ROW];

        const formulas =
          current.formulas[item.row - HEADER_ROW];

        if (
          String(
            values[indexOfColumn('D')] ?? ''
          ).trim() !== item.name
        ) {
          return {
            blocked:
              'Company name changed. Research again.'
          };
        }

        // Recheck only the approved company's cells.
        // Other company rows cannot enter the write list.

        for (
          const column of [
            ...WRITABLE,
            'AP'
          ]
        ) {
          const i = indexOfColumn(column);

          if (
            !same(
              values[i],
              rowValues(state, item)[i]
            ) ||
            !same(
              formulas[i],
              rowFormulas(state, item)[i]
            )
          ) {
            return {
              blocked:
                column +
                item.row +
                ' changed since review. Research again.'
            };
          }
        }

        const writes = [];

        for (const column of selected) {
          const proposal =
            item.proposal.proposed.get(column);

          if (
            !proposal ||
            !WRITABLE.has(column)
          ) {
            return {
              blocked:
                'Unverified field ' + column
            };
          }

          const expected =
            proposal.header ||
            item.data.column_headers?.[column] ||
            '';

          if (
            !expected ||
            !headerMatches(
              headers[indexOfColumn(column)],
              expected,
              column
            ) ||
            !same(
              headers[indexOfColumn(column)],
              state.snapshot.values[
                HEADER_ROW - 1
              ]?.[
                indexOfColumn(column)
              ]
            )
          ) {
            return {
              blocked:
                'Header mismatch at ' +
                column +
                HEADER_ROW +
                '. No changes.'
            };
          }

          const i = indexOfColumn(column);
          const formula = formulas[i];
          const value = values[i];

          const blank =
            value === '' ||
            value == null;

          const allowPlaceholder =
            column === 'J' &&
            item.allowJ &&
            value === PLACEHOLDER;

          if (
            (
              typeof formula === 'string' &&
              formula.startsWith('=')
            ) ||
            (
              !blank &&
              !allowPlaceholder
            )
          ) {
            return {
              blocked:
                'Existing value/formula in ' +
                column +
                item.row +
                '. No changes.'
            };
          }

          writes.push({
            column,
            value: proposal.value
          });
        }

        if (!writes.length) {
          return {
            blocked:
              'No approved fields to save.'
          };
        }

        // Validate audit before making any changes.

        const existingAudit =
          sheets.getItemOrNullObject(AUDIT);

        existingAudit.load('isNullObject');
        await context.sync();

        let auditStart = 2;

        let writeHeaders =
          existingAudit.isNullObject;

        if (!existingAudit.isNullObject) {
          const usedAudit =
            existingAudit.getUsedRangeOrNullObject(
              true
            );

          usedAudit.load(
            'isNullObject,rowIndex,rowCount'
          );

          await context.sync();

          if (usedAudit.isNullObject) {
            writeHeaders = true;

          } else {
            const auditHeader =
              existingAudit.getRange(
                'A1:P1'
              );

            auditHeader.load(
              'values,formulas'
            );

            await context.sync();

            if (
              !AUDIT_HEADERS.every(
                (h, i) =>
                  auditHeader.values[0][i] === h &&
                  !(
                    typeof auditHeader.formulas[0][i] === 'string' &&
                    auditHeader.formulas[0][i].startsWith('=')
                  )
              )
            ) {
              return {
                blocked:
                  'Company Research Audit headers differ. Nothing written.'
              };
            }

            auditStart = Math.max(
              2,
              usedAudit.rowIndex +
              usedAudit.rowCount +
              1
            );
          }
        }

        const time =
          new Date().toISOString();

        const runId =
          'V1052-' +
          time.replace(/\D/g, '') +
          '-R' +
          item.row;

        const selectedSet =
          new Set(selected);

        const auditRows =
          COLUMNS.map(column => {
            const p =
              item.proposal.proposed.get(column);

            const raw =
              item.proposal.raw.get(column) || {};

            const i =
              indexOfColumn(column);

            const old =
              typeof formulas[i] === 'string' &&
              formulas[i].startsWith('=')
                ? '[EXISTING FORMULA PROTECTED]'
                : values[i] ?? '';

            return [
              runId,
              time,
              item.row,
              item.name,

              item.target.candidate_id,

              item.target.global_parent_wikidata_id,

              column,
              old,

              p?.value ?? raw.value ?? '',

              selectedSet.has(column)
                ? 'PREPARED_NOT_CONFIRMED'
                : 'NOT_APPROVED',

              p?.scope ?? raw.scope ?? '',

              p?.status ?? raw.status ?? '',

              p?.source_url ?? raw.source_url ?? '',

              p?.notes ?? raw.notes ?? '',

              '',

              p?.header ?? raw.header ?? ''
            ].map(safeExcel);
          });

        // Backup the ENTIRE Company Master worksheet
        // before the audit and the company-row write.

        const backupName =
          'CM52_R' +
          item.row +
          '_' +
          Date.now().toString(36);

        if (backupName.length > 31) {
          return {
            blocked:
              'Backup name too long.'
          };
        }

        const existingBackup =
          sheets.getItemOrNullObject(
            backupName
          );

        existingBackup.load('isNullObject');

        await context.sync();

        if (!existingBackup.isNullObject) {
          return {
            blocked:
              'Backup name already exists.'
          };
        }

        const copy = master.copy(
          Excel.WorksheetPositionType.after,
          master
        );

        copy.name = backupName;

        await context.sync();

        backupConfirmed = backupName;

        // Commit audit entries BEFORE Company Master writes.

        const audit = existingAudit.isNullObject
          ? sheets.add(AUDIT)
          : sheets.getItem(AUDIT);

        if (writeHeaders) {
          audit.getRange(
            'A1:P1'
          ).values = [
            AUDIT_HEADERS
          ];
        }

        const auditEnd =
          auditStart +
          auditRows.length -
          1;

        audit.getRange(
          'A' +
          auditStart +
          ':P' +
          auditEnd
        ).values = auditRows;

        await context.sync();

        auditConfirmed = true;

        // ONLY this one row and its approved columns
        // are written. Never changes cell formatting.

        for (const write of writes) {
          master.getRange(
            write.column +
            item.row
          ).values = [
            [
              safeExcel(write.value)
            ]
          ];
        }

        // Preserve all existing AP values and formulas.
        // Write an optional summary only if AP is blank.

        const ap = indexOfColumn('AP');

        if (
          (
            values[ap] == null ||
            values[ap] === ''
          ) &&
          !(
            typeof formulas[ap] === 'string' &&
            formulas[ap].startsWith('=')
          )
        ) {
          master.getRange(
            'AP' +
            item.row
          ).values = [
            [
              safeExcel(
                noteFor(item, selected)
              )
            ]
          ];
        }

        await context.sync();

        // Mark audit rows saved only after sync.

        const confirmedAt =
          new Date().toISOString();

        audit.getRange(
          'J' +
          auditStart +
          ':J' +
          auditEnd
        ).values = COLUMNS.map(
          column => [
            selectedSet.has(column)
              ? 'SAVED'
              : 'NOT_APPROVED'
          ]
        );

        audit.getRange(
          'O' +
          auditStart +
          ':O' +
          auditEnd
        ).values = COLUMNS.map(
          column => [
            selectedSet.has(column)
              ? confirmedAt
              : ''
          ]
        );

        await context.sync();

        return {
          backupName,
          saved: selected.map(
            column =>
              column +
              item.row
          )
        };
      });

      if (result.blocked) {
        item.blocked = true;

        item.saveButton.textContent =
          'Re-research required — ' +
          item.name;

        item.saveMessage.textContent =
          result.blocked +
          ' No Excel values were written.';

        say(
          'SAVE BLOCKED: ' +
          result.blocked
        );

      } else {
        item.saved = true;

        item.saveButton.textContent =
          'Saved ✓ — ' +
          item.name;

        item.saveMessage.textContent =
          'Saved only ' +
          item.name +
          ': ' +
          result.saved.join(', ') +
          '. Backup: ' +
          result.backupName +
          '. Audit: ' +
          AUDIT +
          '. Other companies were not saved.';

        say(
          'V' +
          VERSION +
          ' SAVED: ' +
          item.name +
          '. ' +
          result.saved.join(', ')
        );
      }

    } catch (error) {
      item.uncertain = true;

      item.saveButton.textContent =
        'Inspect workbook before retrying';

      item.saveMessage.textContent =
        'SAVE NOT CONFIRMED. ' +
        String(error?.message || error) +
        '. ' +
        (
          backupConfirmed
            ? 'Backup: ' + backupConfirmed + '. '
            : ''
        ) +
        (
          auditConfirmed
            ? 'Audit was prepared. '
            : ''
        ) +
        'Check Company Master and the audit; do not retry automatically.';

      say(
        'SAVE NOT CONFIRMED — check workbook before retrying.'
      );

    } finally {
      item.saving = false;
      refreshButtons(state, item);
    }
  }

  // ==================================================
  // CHANGE DETECTION
  // ==================================================

  function onChanged(event) {
    const address =
      String(event?.address || '')
        .split('!')
        .pop()
        .replace(/\$/g, '')
        .toUpperCase();

    const match =
      /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/
        .exec(address);

    if (!match) return;

    const firstCol =
      indexOfColumn(match[1]);

    const lastCol =
      indexOfColumn(
        match[3] || match[1]
      );

    const lastRow =
      Number(match[4] || match[2]);

    if (
      firstCol <= indexOfColumn('D') &&
      lastCol >= indexOfColumn('D') &&
      lastRow >= FIRST_ROW
    ) {
      ++generation;
      currentReview = null;

      if (results) {
        results.replaceChildren();
      }

      say(
        'Company names changed. Research again before approving or saving.'
      );
    }
  }

  // ==================================================
  // INITIALIZATION
  // ==================================================

  async function initialize() {
    if (initialized) {
      ensureResearchButton();

      say(
        'V' +
        VERSION +
        ' READY — individual column Approve buttons; separate company Save.'
      );

      return;
    }

    if (initializing) return;

    initializing = true;

    say(
      'V' +
      VERSION +
      ' connecting to Excel…'
    );

    try {
      await Office.onReady();

      if (typeof Excel === 'undefined') {
        throw Error(
          'Open this add-in inside Excel.'
        );
      }

      await Excel.run(async context => {
        const sheet =
          context.workbook.worksheets
            .getItemOrNullObject(
              MASTER
            );

        sheet.load('isNullObject');

        await context.sync();

        if (sheet.isNullObject) {
          throw Error(
            'Company Master worksheet was not found.'
          );
        }

        sheet.onChanged.add(onChanged);

        await context.sync();
      });

      initialized = true;

      ensureResearchButton();

      say(
        'V' +
        VERSION +
        ' READY. Research companies; Approve individual eligible columns; ' +
        'Save only that company. Existing Excel headers, values and formulas stay protected.'
      );

    } catch (error) {
      say(
        'CONNECTION ERROR: ' +
        String(error?.message || error)
      );

    } finally {
      initializing = false;
    }
  }

  if (retry) {
    retry.addEventListener(
      'click',
      initialize
    );
  }

  ensureResearchButton();
  initialize();

})();
