(() => {
  'use strict';

  // V9: Global results are reference-only. Approved Malaysia data
  // can be written ONLY to empty E, F, K and AP cells.
  const WORKER = 'https://malaysia-company-research.rimli271.workers.dev';
  const SHEET = 'Company Master';
  const FIRST_ROW = 7;
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
  function text(parent, tag, value, cssClass) {
    const node = document.createElement(tag);
    node.textContent = String(value ?? '');
    if (cssClass) node.className = cssClass;
    parent.appendChild(node);
    return node;
  }
  function link(parent, label, value) {
    try {
      const url = new URL(String(value));
      if (url.protocol !== 'https:' && url.protocol !== 'http:') return;
      const node = document.createElement('a');
      node.href = url.href;
      node.target = '_blank';
      node.rel = 'noopener noreferrer';
      node.textContent = label;
      parent.appendChild(node);
    } catch { /* Invalid or missing URL. */ }
  }
  function strings(value) {
    const items = Array.isArray(value) ? value : [value];
    return [...new Set(items.filter(x => typeof x === 'string')
      .map(x => x.trim()).filter(Boolean))];
  }
  function display(value) {
    return strings(value).join('; ') || 'Not available';
  }
  function columnNumber(letters) {
    return [...letters].reduce((n, c) =>
      n * 26 + c.charCodeAt(0) - 64, 0);
  }
  function changedRow(address) {
    const local = String(address || '').split('!').pop()
      .replace(/\$/g, '').toUpperCase();
    const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(local);
    if (!m) return null;
    if (columnNumber(m[1]) > 4 || columnNumber(m[3] || m[1]) < 4) return null;
    const first = Number(m[2]);
    const last = Number(m[4] || m[2]);
    if (last < FIRST_ROW) return null;
    if (Math.max(first, FIRST_ROW) !== last) return { multiple: true };
    return { row: last };
  }
  async function readCompany(row) {
    return Excel.run(async context => {
      const cell = context.workbook.worksheets
        .getItem(SHEET).getRange('D' + row);
      cell.load('values');
      await context.sync();
      return String(cell.values[0][0] ?? '').trim();
    });
  }
  function field(parent, title, value, source) {
    const box = document.createElement('div');
    box.style.padding = '8px 0';
    box.style.borderBottom = '1px solid #e5e7eb';
    text(box, 'strong', title);
    text(box, 'p', display(value));
    if (source) link(box, 'View source', source);
    parent.appendChild(box);
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

  // Only treat the Worker's official-source evidence as sufficient
  // to propose a Malaysia-specific value. Wikidata alone is not.
  function verifiedEvidence(candidate) {
    const evidence = candidate.malaysia_profile?.malaysia_evidence;
    return candidate.verified_malaysia_presence === true &&
      Array.isArray(evidence) && evidence.some(x => x?.verified === true &&
        /official|bank negara|maybank/i.test(String(x.type || '')));
  }
  function hasVerifiedWebsite(candidate, site) {
    try {
      const host = new URL(site).hostname.toLowerCase();
      const evidence = candidate.malaysia_profile?.malaysia_evidence || [];
      return evidence.some(e => {
        if (e?.verified !== true || !e.source_url) return false;
        return new URL(e.source_url).hostname.toLowerCase() === host;
      });
    } catch { return false; }
  }
  function proposedValues(candidate) {
    const profile = candidate.malaysia_profile || {};
    if (!/^Q\d+$/.test(String(candidate.wikidata_id || '')) ||
        !verifiedEvidence(candidate)) return {};

    const proposed = {};
    const country = strings(candidate.global_profile?.country)
      .map(s => s.toLowerCase());
    const hq = String(profile.malaysia_headquarters || '').trim();
    if (country.includes('malaysia') && hq) {
      proposed.E = 'Malaysian-headquartered company';
    } else if (country.length && !country.includes('malaysia')) {
      proposed.E = 'Global company with Malaysia operations';
    } else {
      proposed.E = 'Company with verified Malaysia operations (origin unconfirmed)';
    }
    // Preserve the main-office qualification; never use global HQ values.
    if (hq && /malaysia/i.test(hq)) proposed.F = hq;
    const site = String(profile.malaysia_website || '').trim();
    if (site && hasVerifiedWebsite(candidate, site)) proposed.K = site;
    // G (industry) has no separately verified Malaysia source yet.
    // L, M, N, O, AF, AG, AH, AI have no supported data yet.
    return proposed;
  }

  function researchNote(company, candidate, proposed) {
    const p = candidate.malaysia_profile || {};
    const evidence = Array.isArray(p.malaysia_evidence) ? p.malaysia_evidence : [];
    const lines = [
      'MALAYSIA-SPECIFIC RESEARCH — REVIEWED',
      'Date (UTC): ' + new Date().toISOString(),
      'Entered company name: ' + company,
      'Selected global identity: ' + String(candidate.name || ''),
      'Wikidata: https://www.wikidata.org/wiki/' + candidate.wikidata_id,
      'Malaysia entity/brand: ' + display(p.malaysia_company_name),
      'Malaysia main office: ' + display(p.malaysia_headquarters),
      'Malaysia address: ' + display(p.malaysia_office_address),
      'Malaysia site: ' + display(p.malaysia_website),
      'Malaysia careers: ' + display(p.malaysia_careers_url),
      'Approved proposed E: ' + display(proposed.E),
      'Approved proposed F: ' + display(proposed.F),
      'Approved proposed K: ' + display(proposed.K),
      'E/F/K written only when their cells were blank.',
      'Industry, revenue, employees, rent, transit and commute NOT verified.',
      'Visa sponsorship NOT established.',
      'OFFICIAL RESEARCH EVIDENCE:'
    ];
    evidence.forEach(e => {
      if (e?.verified === true) {
        lines.push('- ' + String(e.type || 'Official source') + ': ' +
          String(e.note || '') + ' ' + String(e.source_url || ''));
      }
    });
    return lines.join('\n').slice(0, 30000);
  }

  function preview(parent, company, candidate, row, token) {
    const id = String(candidate.wikidata_id || '');
    const p = candidate.malaysia_profile || {};
    const g = candidate.global_profile || {};
    const global = panel(parent, 'Global company — reference only');
    text(global, 'p', 'These global values will NEVER be written to Company Master.');
    field(global, 'Company type reported by Wikidata', g.company_type);
    field(global, 'Global headquarters (not Malaysia office)', g.headquarters);
    field(global, 'Global industry (unverified for Malaysia)', g.industry);
    field(global, 'Global websites', g.websites);
    link(global, 'Open Wikidata record', candidate.source_url ||
      'https://www.wikidata.org/wiki/' + id);

    const local = panel(parent, 'Malaysia-specific information');
    field(local, 'Malaysia presence status', p.malaysia_presence_status);
    field(local, 'Malaysia entity / brand', p.malaysia_company_name);
    field(local, 'Malaysia main office', p.malaysia_headquarters);
    field(local, 'Malaysia office address', p.malaysia_office_address);
    field(local, 'Malaysia website', p.malaysia_website);
    link(local, 'Open Malaysia website', p.malaysia_website);
    field(local, 'Malaysia careers', p.malaysia_careers_url);
    link(local, 'Open Malaysia careers', p.malaysia_careers_url);
    text(local, 'h4', 'Official-source evidence');
    const evidence = Array.isArray(p.malaysia_evidence) ? p.malaysia_evidence : [];
    if (!evidence.length) text(local, 'p', 'No evidence returned.');
    evidence.forEach(e => {
      text(local, 'p', (e.verified ? 'Checked: ' : 'Review only: ') +
        String(e.note || e.type || 'Source'));
      link(local, 'View evidence', e.source_url);
    });
    text(local, 'p', p.research_note || 'Review the sources carefully.');

    const proposed = proposedValues(candidate);
    const approval = panel(parent, 'Proposed Malaysia-only Excel save');
    text(approval, 'p', 'Only empty cells can be filled. Existing values and formulas are skipped.');
    field(approval, 'E — Company Type', proposed.E || 'Not proposed');
    field(approval, 'F — Company Headquarters (Malaysia main office)',
      proposed.F || 'Not proposed');
    field(approval, 'G — Industry', 'Not proposed: no verified Malaysia-only value');
    field(approval, 'K — Company Website (Malaysia)', proposed.K || 'Not proposed');
    text(approval, 'p', 'L, M, N, O and AF–AI will NOT be written. ' +
      'Revenue, staff numbers, rent, transit and commute are not yet researched.');
    text(approval, 'p', 'AP — Malaysia research note and source links (if blank).');
    text(approval, 'p', 'D' + row + ' and every column header remain unchanged.');

    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = 'Approve Malaysia-only save';
    approval.appendChild(button);
    const message = text(approval, 'p', '', 'small');
    if (!Object.keys(proposed).length) {
      button.disabled = true;
      message.textContent = 'Save unavailable: no supported Malaysia-specific cells.';
    } else {
      button.addEventListener('click', () =>
        saveMalaysia(company, row, candidate, proposed, token, button, message));
    }
  }

  async function saveMalaysia(company, row, candidate, proposed, token, button, message) {
    if (token !== requestNumber) {
      message.textContent = 'Research changed. Nothing saved.';
      return;
    }
    button.disabled = true;
    button.textContent = 'Checking blank cells...';
    const note = researchNote(company, candidate, proposed);
    try {
      const outcome = await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem(SHEET);
        const nameCell = sheet.getRange('D' + row);
        nameCell.load('values');
        const columns = [...Object.keys(proposed), 'AP'];
        const cells = columns.map(column => ({
          column, range: sheet.getRange(column + row)
        }));
        cells.forEach(item => item.range.load('values,formulas'));
        await context.sync();
        if (token !== requestNumber ||
            String(nameCell.values[0][0] ?? '').trim() !== company) {
          return { blocked: 'Company changed. Research again.' };
        }
        const saved = [];
        const skipped = [];
        for (const item of cells) {
          const value = item.range.values[0][0];
          const formula = item.range.formulas[0][0];
          const occupied = (value !== null && value !== undefined &&
            String(value).trim() !== '') ||
            (typeof formula === 'string' && formula.startsWith('='));
          if (occupied) {
            skipped.push(item.column + row);
            continue;
          }
          item.range.values = [[item.column === 'AP' ? note : proposed[item.column]]];
          saved.push(item.column + row);
        }
        if (saved.length) await context.sync();
        return { saved, skipped };
      });
      if (outcome.blocked) {
        show('SAVE BLOCKED\n' + outcome.blocked);
        message.textContent = outcome.blocked;
        button.disabled = false;
        button.textContent = 'Approve Malaysia-only save';
        return;
      }
      show('MALAYSIA SAVE COMPLETE\nSaved: ' +
        (outcome.saved.join(', ') || 'none') + '\nSkipped: ' +
        (outcome.skipped.join(', ') || 'none') + '\nAll other cells unchanged.');
      message.textContent = 'Saved ' +
        (outcome.saved.join(', ') || 'no cells') + '. Skipped ' +
        (outcome.skipped.join(', ') || 'none') + '.';
      button.textContent = 'Save checked';
    } catch (error) {
      show('SAVE ERROR\n' + errorText(error));
      message.textContent = 'Save not confirmed. Check the status box.';
      button.disabled = false;
      button.textContent = 'Approve Malaysia-only save';
    }
  }

  function renderCandidates(company, row, candidates, token) {
    clear();
    if (!results) return;
    text(results, 'h2', 'Review research for D' + row);
    text(results, 'p', 'Searched: ' + company);
    if (!candidates.length) text(results, 'p', 'No candidates found. Nothing written.');
    candidates.forEach(candidate => {
      const card = document.createElement('div');
      card.className = 'candidate';
      text(card, 'h3', candidate.name || 'Unnamed candidate');
      text(card, 'p', candidate.description || 'No description');
      text(card, 'p', 'Wikidata ID: ' + String(candidate.wikidata_id || ''), 'small');
      text(card, 'p', 'Select only if this is the correct parent company.', 'small');
      const select = document.createElement('button');
      select.type = 'button';
      select.textContent = 'Select for review';
      card.appendChild(select);
      select.addEventListener('click', () => {
        if (token !== requestNumber) {
          show('Results expired. Edit D' + row + ' to research again.');
          return;
        }
        if (!/^Q\d+$/.test(String(candidate.wikidata_id || ''))) return;
        select.disabled = true;
        select.textContent = 'Selected for review';
        card.style.border = '2px solid #13865c';
        preview(card, company, candidate, row, token);
        show('V9 MALAYSIA PREVIEW READY\n' +
          'Global information is reference only.\n' +
          'Nothing has been saved.');
      });
      results.appendChild(card);
    });
    const reject = document.createElement('button');
    reject.type = 'button';
    reject.textContent = 'Reject all candidates';
    reject.addEventListener('click', () => {
      if (token !== requestNumber) return;
      ++requestNumber;
      clear();
      show('Candidates rejected. No cells changed.');
    });
    results.appendChild(reject);
  }

  async function research(company, row, token) {
    clear();
    show('V9 RESEARCHING\nCell: D' + row + '\nCompany: ' + company);
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 20000);
    try {
      const response = await fetch(WORKER + '/research?company=' +
        encodeURIComponent(company), {
          headers: { Accept: 'application/json' },
          signal: controller.signal
        });
      if (!response.ok) throw new Error('Research HTTP ' + response.status);
      const data = await response.json();
      if (token !== requestNumber) return;
      if (data.status !== 'candidates_for_review') {
        throw new Error(data.message || 'Research unavailable');
      }
      const candidates = Array.isArray(data.candidates) ? data.candidates : [];
      show('V9 RESEARCH COMPLETE\nWorker: ' + data.version +
        '\nCandidates: ' + candidates.length + '\nSelect one to preview.');
      renderCandidates(company, row, candidates, token);
    } catch (error) {
      if (token === requestNumber) {
        clear();
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
    clear();
    if (location.multiple) {
      show('Multiple rows changed. Research one company at a time.');
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
      if (token === requestNumber) show('CHANGE ERROR\n' + errorText(error));
    }
  }

  async function initialize() {
    if (active) {
      show('V9 READY\nMonitoring Company Master D7 onward.');
      return;
    }
    if (starting) return;
    starting = true;
    show('V9: Connecting to Excel...');
    try {
      await Office.onReady();
      if (typeof Excel === 'undefined') {
        throw new Error('Open this add-in inside Excel Online.');
      }
      await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItemOrNullObject(SHEET);
        sheet.load('isNullObject');
        await context.sync();
        if (sheet.isNullObject) throw new Error('Company Master not found.');
        sheet.onChanged.add(onChanged);
        await context.sync();
      });
      active = true;
      show('V9 READY — MALAYSIA-ONLY SAVE\n' +
        'Edit Company Master D7 onward to research.\n' +
        'Select a candidate, review Malaysia sources, then approve.\n' +
        'Only blank E, F, K, AP can be written. All other cells stay unchanged.');
    } catch (error) {
      show('V9 CONNECTION ERROR\n' + errorText(error));
    } finally {
      starting = false;
    }
  }

  if (retry) retry.addEventListener('click', initialize);
  initialize();
})();
