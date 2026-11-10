(() => {
  'use strict';

  // V9.1: Malaysia-only, approval-based saving.
  // Only E, F, J, K and AP may be written.
  const WORKER =
    'https://malaysia-company-research.rimli271.workers.dev';
  const SHEET = 'Company Master';
  const FIRST_ROW = 7;
  const PLACEHOLDER = '[verify office address]';

  const HEADERS = {
    E: 'Company Type',
    F: 'Company Headquarters',
    J: 'Office Address',
    K: 'Company Website'
  };

  const status = document.getElementById('status');
  const results = document.getElementById('results');
  const retry = document.getElementById('retry');

  let active = false;
  let starting = false;
  let requestNumber = 0;

  function show(message) {
    if (status) status.textContent = message;
  }

  function clear() {
    if (results) results.replaceChildren();
  }

  function errorText(error) {
    return String(error?.message || error);
  }

  function text(parent, tag, value, className) {
    const node = document.createElement(tag);
    node.textContent = String(value ?? '');
    if (className) node.className = className;
    parent.appendChild(node);
    return node;
  }

  function link(parent, label, value) {
    try {
      const url = new URL(String(value));
      if (!['https:', 'http:'].includes(url.protocol)) {
        return;
      }

      const a = document.createElement('a');
      a.href = url.href;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      a.textContent = label;
      parent.appendChild(a);
    } catch {
      // Missing or invalid URL.
    }
  }

  function strings(value) {
    const items = Array.isArray(value)
      ? value
      : [value];

    return [...new Set(
      items
        .filter(x => typeof x === 'string')
        .map(x => x.trim())
        .filter(Boolean)
    )];
  }

  function display(value) {
    return strings(value).join('; ') ||
      'Not available';
  }

  function panel(parent, heading) {
    const node = document.createElement('div');

    node.style.padding = '12px';
    node.style.marginTop = '14px';
    node.style.border = '1px solid #cbd5df';
    node.style.borderRadius = '8px';

    text(node, 'h3', heading);
    parent.appendChild(node);

    return node;
  }

  function field(parent, title, value) {
    const node = document.createElement('div');

    node.style.padding = '8px 0';
    node.style.borderBottom = '1px solid #e5e7eb';

    text(node, 'strong', title);
    text(node, 'p', display(value));

    parent.appendChild(node);
  }

  function columnNumber(letters) {
    return [...letters].reduce(
      (n, c) => n * 26 + c.charCodeAt(0) - 64,
      0
    );
  }

  function changedRow(address) {
    const local = String(address || '')
      .split('!')
      .pop()
      .replace(/\$/g, '')
      .toUpperCase();

    const m =
      /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(local);

    if (!m) return null;

    if (
      columnNumber(m[1]) > 4 ||
      columnNumber(m[3] || m[1]) < 4
    ) {
      return null;
    }

    const first = Number(m[2]);
    const last = Number(m[4] || m[2]);

    if (last < FIRST_ROW) return null;

    if (Math.max(first, FIRST_ROW) !== last) {
      return { multiple: true };
    }

    return { row: last };
  }

  async function readCompany(row) {
    return Excel.run(async context => {
      const cell = context.workbook.worksheets
        .getItem(SHEET)
        .getRange('D' + row);

      cell.load('values');
      await context.sync();

      return String(
        cell.values[0][0] ?? ''
      ).trim();
    });
  }

  function evidence(candidate, type) {
    const items =
      candidate.malaysia_profile?.malaysia_evidence;

    return Array.isArray(items) &&
      items.some(item =>
        item?.verified === true &&
        item.type === type &&
        typeof item.source_url === 'string' &&
        item.source_url.startsWith('https://')
      );
  }

  function verifiedPresence(candidate) {
    const items =
      candidate.malaysia_profile?.malaysia_evidence;

    return (
      candidate.verified_malaysia_presence === true &&
      Array.isArray(items) &&
      items.some(item =>
        item?.verified === true &&
        /official|bank negara|maybank/i.test(
          String(item.type || '')
        )
      )
    );
  }

  function verifiedWebsite(candidate, site) {
    try {
      const host =
        new URL(site).hostname.toLowerCase();

      const items =
        candidate.malaysia_profile?.malaysia_evidence ||
        [];

      return items.some(item => {
        if (
          item?.verified !== true ||
          !item.source_url
        ) {
          return false;
        }

        return new URL(
          item.source_url
        ).hostname.toLowerCase() === host;
      });
    } catch {
      return false;
    }
  }

  function verifiedAddress(candidate) {
    const id = String(
      candidate.wikidata_id || ''
    );

    if (id === 'Q488048') {
      return evidence(
        candidate,
        'Official Malaysia office directory'
      );
    }

    if (id === 'Q1364018') {
      return (
        evidence(
          candidate,
          'Current official Maybank contact page'
        ) ||
        evidence(
          candidate,
          'Maybank investor-relations renaming announcement'
        )
      );
    }

    // Other companies need separate
    // official address verification.
    return false;
  }

  function proposedValues(candidate) {
    const p =
      candidate.malaysia_profile || {};

    if (
      !/^Q\d+$/.test(
        String(candidate.wikidata_id || '')
      ) ||
      !verifiedPresence(candidate)
    ) {
      return {};
    }

    const proposed = {};

    const country = strings(
      candidate.global_profile?.country
    ).map(x => x.toLowerCase());

    const hq = String(
      p.malaysia_headquarters || ''
    ).trim();

    if (
      country.includes('malaysia') &&
      hq
    ) {
      proposed.E =
        'Malaysian-headquartered company';
    } else if (
      country.length &&
      !country.includes('malaysia')
    ) {
      proposed.E =
        'Global company with Malaysia operations';
    } else {
      proposed.E =
        'Company with verified Malaysia operations (origin unconfirmed)';
    }

    if (
      hq &&
      /malaysia/i.test(hq)
    ) {
      proposed.F = hq;
    }

    const address = String(
      p.malaysia_office_address || ''
    ).trim();

    if (
      address &&
      /malaysia/i.test(address) &&
      verifiedAddress(candidate)
    ) {
      proposed.J = address;
    }

    const site = String(
      p.malaysia_website || ''
    ).trim();

    if (
      site &&
      verifiedWebsite(candidate, site)
    ) {
      proposed.K = site;
    }

    return proposed;
  }

  function noteFor(company, candidate, proposed) {
    const p =
      candidate.malaysia_profile || {};

    const items =
      p.malaysia_evidence || [];

    const lines = [
      'MALAYSIA-SPECIFIC RESEARCH — APPROVAL-BASED V9.1',
      'Date (UTC): ' + new Date().toISOString(),
      'Entered company: ' + company,
      'Selected Wikidata candidate: ' +
        String(candidate.name || ''),
      'Source: https://www.wikidata.org/wiki/' +
        candidate.wikidata_id,

      'Malaysia entity: ' +
        display(p.malaysia_company_name),

      'Malaysia main office: ' +
        display(p.malaysia_headquarters),

      'Malaysia address: ' +
        display(p.malaysia_office_address),

      'Malaysia website: ' +
        display(p.malaysia_website),

      'Malaysia careers: ' +
        display(p.malaysia_careers_url),

      'Proposed E: ' + display(proposed.E),
      'Proposed F: ' + display(proposed.F),
      'Proposed J: ' + display(proposed.J),
      'Proposed K: ' + display(proposed.K),

      'Values saved only where blank.',
      'J placeholder replacement requires separate opt-in.',
      'Global headquarters and unrelated websites NOT saved.',
      'G, I, L, M, N, O, AF–AI NOT changed.',
      'Visa sponsorship NOT established.',

      'SUPPORTING OFFICIAL EVIDENCE:'
    ];

    if (Array.isArray(items)) {
      items
        .filter(x => x?.verified === true)
        .forEach(x => {
          lines.push(
            '- ' +
            String(x.type || 'Source') +
            ': ' +
            String(x.note || '') +
            ' ' +
            String(x.source_url || '')
          );
        });
    }

    return lines.join('\n').slice(0, 30000);
  }

  function preview(
    parent,
    company,
    candidate,
    row,
    token
  ) {
    const p =
      candidate.malaysia_profile || {};

    const g =
      candidate.global_profile || {};

    const global = panel(
      parent,
      'Global company — reference only'
    );

    text(
      global,
      'p',
      'Global values will NOT be written to Excel.'
    );

    field(
      global,
      'Company type (Wikidata)',
      g.company_type
    );

    field(
      global,
      'Global headquarters (not Malaysia office)',
      g.headquarters
    );

    field(
      global,
      'Global industry (not locally verified)',
      g.industry
    );

    field(
      global,
      'Global websites',
      g.websites
    );

    link(
      global,
      'Open Wikidata record',
      candidate.source_url
    );

    const local = panel(
      parent,
      'Malaysia-specific information'
    );

    field(
      local,
      'Malaysia presence',
      p.malaysia_presence_status
    );

    field(
      local,
      'Malaysia entity / brand',
      p.malaysia_company_name
    );

    field(
      local,
      'Malaysia main office',
      p.malaysia_headquarters
    );

    field(
      local,
      'Malaysia office address',
      p.malaysia_office_address
    );

    field(
      local,
      'Malaysia website',
      p.malaysia_website
    );

    link(
      local,
      'Open Malaysia website',
      p.malaysia_website
    );

    field(
      local,
      'Malaysia careers',
      p.malaysia_careers_url
    );

    link(
      local,
      'Open Malaysia careers',
      p.malaysia_careers_url
    );

    text(
      local,
      'h4',
      'Evidence'
    );

    const items = Array.isArray(
      p.malaysia_evidence
    )
      ? p.malaysia_evidence
      : [];

    if (!items.length) {
      text(
        local,
        'p',
        'No evidence supplied.'
      );
    }

    items.forEach(item => {
      text(
        local,
        'p',
        (
          item?.verified
            ? 'Checked: '
            : 'Review only: '
        ) +
        String(
          item?.note ||
          item?.type ||
          'Source'
        )
      );

      link(
        local,
        'View evidence',
        item?.source_url
      );
    });

    text(
      local,
      'p',
      p.research_note ||
        'Review official sources.'
    );

    const proposed =
      proposedValues(candidate);

    const save = panel(
      parent,
      'Proposed Malaysia-only Excel save'
    );

    text(
      save,
      'p',
      'Only empty cells are filled. Existing values and formulas are skipped.'
    );

    field(
      save,
      'E — Company Type',
      proposed.E || 'Not proposed'
    );

    field(
      save,
      'F — Company Headquarters (Malaysia main office)',
      proposed.F || 'Not proposed'
    );

    field(
      save,
      'G — Industry',
      'Not proposed; left unchanged'
    );

    field(
      save,
      'I — Office Location',
      'Left unchanged'
    );

    field(
      save,
      'J — Office Address',
      proposed.J ||
        'Not proposed: address lacks verified office evidence'
    );

    field(
      save,
      'K — Company Website (Malaysia)',
      proposed.K || 'Not proposed'
    );

    text(
      save,
      'p',
      'AP — Malaysia-only research note (if blank).'
    );

    text(
      save,
      'p',
      'L–O and AF–AI will NOT be changed.'
    );

    let replacePlaceholder = null;

    if (proposed.J) {
      const label =
        document.createElement('label');

      label.style.display = 'block';
      label.style.padding = '10px 0';

      replacePlaceholder =
        document.createElement('input');

      replacePlaceholder.type = 'checkbox';
      replacePlaceholder.checked = false;
      replacePlaceholder.style.marginRight = '8px';

      label.appendChild(replacePlaceholder);

      label.appendChild(
        document.createTextNode(
          'I explicitly allow replacing the exact text ' +
          PLACEHOLDER +
          ' in J' + row +
          ' only. Never replace any other existing value.'
        )
      );

      save.appendChild(label);
    }

    const button =
      document.createElement('button');

    button.type = 'button';
    button.textContent =
      'Approve Malaysia-only save';

    save.appendChild(button);

    const message = text(
      save,
      'p',
      '',
      'small'
    );

    if (!Object.keys(proposed).length) {
      button.disabled = true;

      message.textContent =
        'No supported Malaysia-specific values. Nothing to save.';
    } else {
      button.addEventListener('click', () =>
        saveMalaysia(
          company,
          row,
          candidate,
          proposed,
          token,
          replacePlaceholder?.checked === true,
          button,
          message
        )
      );
    }
  }

  async function saveMalaysia(
    company,
    row,
    candidate,
    proposed,
    token,
    permitPlaceholder,
    button,
    message
  ) {
    if (token !== requestNumber) {
      message.textContent =
        'Research changed. Nothing saved.';
      return;
    }

    button.disabled = true;
    button.textContent =
      'Checking Excel cells...';

    const note = noteFor(
      company,
      candidate,
      proposed
    );

    try {
      const outcome = await Excel.run(async context => {
        const sheet =
          context.workbook.worksheets.getItem(SHEET);

        const nameCell =
          sheet.getRange('D' + row);

        nameCell.load('values');

        const headerChecks =
          Object.entries(HEADERS).map(
            ([col, name]) => ({
              col,
              name,
              range: sheet.getRange(col + '6')
            })
          );

        headerChecks.forEach(x =>
          x.range.load('values')
        );

        const columns = [
          ...Object.keys(proposed),
          'AP'
        ];

        const cells = columns.map(col => ({
          col,
          range: sheet.getRange(col + row)
        }));

        cells.forEach(x =>
          x.range.load('values,formulas')
        );

        await context.sync();

        if (
          token !== requestNumber ||
          String(
            nameCell.values[0][0] ?? ''
          ).trim() !== company
        ) {
          return {
            blocked:
              'Company name or research changed. Nothing saved.'
          };
        }

        for (const item of headerChecks) {
          if (
            String(
              item.range.values[0][0] ?? ''
            ).trim() !== item.name
          ) {
            return {
              blocked:
                'Header mismatch at ' +
                item.col +
                '6. Nothing saved. Do not rename columns.'
            };
          }
        }

        const saved = [];
        const skipped = [];

        for (const item of cells) {
          const value =
            item.range.values[0][0];

          const formula =
            item.range.formulas[0][0];

          const hasFormula =
            typeof formula === 'string' &&
            formula.startsWith('=');

          const isEmpty =
            (
              value === null ||
              value === undefined ||
              value === ''
            ) &&
            !hasFormula;

          const allowExactPlaceholder =
            item.col === 'J' &&
            permitPlaceholder &&
            !hasFormula &&
            value === PLACEHOLDER;

          if (
            !isEmpty &&
            !allowExactPlaceholder
          ) {
            skipped.push(item.col + row);
            continue;
          }

          item.range.values = [[
            item.col === 'AP'
              ? note
              : proposed[item.col]
          ]];

          saved.push(item.col + row);
        }

        if (saved.length) {
          await context.sync();
        }

        return { saved, skipped };
      });

      if (outcome.blocked) {
        show(
          'SAVE BLOCKED\n' +
          outcome.blocked
        );

        message.textContent =
          outcome.blocked;

        button.disabled = false;
        button.textContent =
          'Approve Malaysia-only save';

        return;
      }

      show(
        'MALAYSIA SAVE COMPLETE\n' +
        'Saved: ' +
        (outcome.saved.join(', ') || 'none') +
        '\nSkipped: ' +
        (outcome.skipped.join(', ') || 'none') +
        '\nOther cells untouched.'
      );

      message.textContent =
        'Saved ' +
        (outcome.saved.join(', ') || 'none') +
        '. Skipped ' +
        (outcome.skipped.join(', ') || 'none') +
        '.';

      button.textContent = 'Save checked';

    } catch (error) {
      show(
        'SAVE ERROR\n' +
        errorText(error)
      );

      message.textContent =
        'Save not confirmed. Check Excel before retrying.';

      button.disabled = false;
      button.textContent =
        'Approve Malaysia-only save';
    }
  }

  function renderCandidates(
    company,
    row,
    candidates,
    token
  ) {
    clear();

    if (!results) return;

    text(
      results,
      'h2',
      'Review research for D' + row
    );

    text(
      results,
      'p',
      'Searched: ' + company
    );

    if (!candidates.length) {
      text(
        results,
        'p',
        'No candidates. Nothing written.'
      );

      return;
    }

    candidates.forEach(candidate => {
      const card =
        document.createElement('div');

      card.className = 'candidate';

      text(
        card,
        'h3',
        candidate.name || 'Unnamed candidate'
      );

      text(
        card,
        'p',
        candidate.description || 'No description'
      );

      text(
        card,
        'p',
        'Wikidata ID: ' +
        candidate.wikidata_id,
        'small'
      );

      text(
        card,
        'p',
        'Select only the correct parent company.',
        'small'
      );

      const button =
        document.createElement('button');

      button.type = 'button';
      button.textContent =
        'Select for review';

      card.appendChild(button);

      button.addEventListener('click', () => {
        if (token !== requestNumber) {
          show(
            'Research expired. Edit the company cell to try again.'
          );

          return;
        }

        if (
          !/^Q\d+$/.test(
            String(candidate.wikidata_id || '')
          )
        ) {
          return;
        }

        button.disabled = true;
        button.textContent =
          'Selected for review';

        card.style.border =
          '2px solid #13865c';

        preview(
          card,
          company,
          candidate,
          row,
          token
        );

        show(
          'V9.1 MALAYSIA PREVIEW READY\n' +
          'Global data is reference only.\n' +
          'Nothing saved.'
        );
      });

      results.appendChild(card);
    });

    const reject =
      document.createElement('button');

    reject.type = 'button';
    reject.textContent =
      'Reject all candidates';

    reject.addEventListener('click', () => {
      if (token !== requestNumber) return;

      ++requestNumber;
      clear();

      show(
        'Candidates rejected. No cells changed.'
      );
    });

    results.appendChild(reject);
  }

  async function research(company, row, token) {
    clear();

    show(
      'V9.1 RESEARCHING\n' +
      'Cell: D' + row +
      '\nCompany: ' + company
    );

    const controller =
      new AbortController();

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

      if (
        data.status !== 'candidates_for_review'
      ) {
        throw new Error(
          data.message ||
          'Research unavailable'
        );
      }

      const candidates =
        Array.isArray(data.candidates)
          ? data.candidates
          : [];

      show(
        'V9.1 RESEARCH COMPLETE\n' +
        'Worker: ' +
        (data.version || '?') +
        '\nCandidates: ' +
        candidates.length +
        '\nSelect for review.'
      );

      renderCandidates(
        company,
        row,
        candidates,
        token
      );

    } catch (error) {
      if (token === requestNumber) {
        clear();

        show(
          'RESEARCH ERROR\n' +
          errorText(error)
        );
      }
    } finally {
      clearTimeout(timeout);
    }
  }

  async function onChanged(event) {
    const location =
      changedRow(event.address);

    if (!location) return;

    const token = ++requestNumber;

    clear();

    if (location.multiple) {
      show(
        'Multiple company rows changed. Research one row at a time.'
      );

      return;
    }

    try {
      const company =
        await readCompany(location.row);

      if (token !== requestNumber) return;

      if (!company) {
        show(
          'D' + location.row + ' is blank.'
        );
      } else {
        await research(
          company,
          location.row,
          token
        );
      }
    } catch (error) {
      if (token === requestNumber) {
        show(
          'CHANGE ERROR\n' +
          errorText(error)
        );
      }
    }
  }

  async function initialize() {
    if (active) {
      show(
        'V9.1 READY\n' +
        'Monitoring Company Master D7 onward.'
      );

      return;
    }

    if (starting) return;

    starting = true;

    show(
      'V9.1: Connecting to Excel...'
    );

    try {
      await Office.onReady();

      if (typeof Excel === 'undefined') {
        throw new Error(
          'Open this add-in inside Excel Online.'
        );
      }

      await Excel.run(async context => {
        const sheet =
          context.workbook.worksheets
            .getItemOrNullObject(SHEET);

        sheet.load('isNullObject');

        await context.sync();

        if (sheet.isNullObject) {
          throw new Error(
            'Company Master not found.'
          );
        }

        sheet.onChanged.add(onChanged);

        await context.sync();
      });

      active = true;

      show(
        'V9.1 READY — MALAYSIA-ONLY SAVE\n' +
        'Edit D7 onward; select the correct company and review sources.\n' +
        'Approval can fill blank E, F, J, K, AP.\n' +
        'Only an additional checked box can replace the exact J placeholder.'
      );

    } catch (error) {
      show(
        'V9.1 CONNECTION ERROR\n' +
        errorText(error)
      );
    } finally {
      starting = false;
    }
  }

  if (retry) {
    retry.addEventListener(
      'click',
      initialize
    );
  }

  initialize();
})();
