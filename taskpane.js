(() => {
  'use strict';

  const WORKER = 'https://malaysia-company-research.rimli271.workers.dev';
  const SHEET = 'Company Master';
  const FIRST_ROW = 7;
  const NOTE_COLUMN = 'AP';
  const status = document.getElementById('status');
  const results = document.getElementById('results');
  const retry = document.getElementById('retry');
  let active = false;
  let starting = false;
  let requestNumber = 0;

  function show(text) { if (status) status.textContent = text; }
  function clear() { if (results) results.replaceChildren(); }
  function errorText(error) { return String(error?.message || error); }

  function text(parent, tag, value, className) {
    const node = document.createElement(tag);
    node.textContent = String(value ?? '');
    if (className) node.className = className;
    parent.appendChild(node);
    return node;
  }

  function link(parent, label, url) {
    try {
      const parsed = new URL(String(url));
      if (!['https:', 'http:'].includes(parsed.protocol)) return;
      const a = document.createElement('a');
      a.textContent = label;
      a.href = parsed.href;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      parent.appendChild(a);
    } catch { /* No valid link available. */ }
  }

  function list(value) {
    if (value == null || value === '') return [];
    const items = Array.isArray(value) ? value : [value];
    return [...new Set(items.filter(x => typeof x === 'string')
      .map(x => x.trim()).filter(Boolean))];
  }

  function display(value) {
    return list(value).join('; ') || 'Not available / not verified';
  }

  function field(parent, title, value, url) {
    const box = document.createElement('div');
    box.style.padding = '9px 0';
    box.style.borderBottom = '1px solid #e5e7eb';
    text(box, 'strong', title);
    text(box, 'p', display(value));
    if (url) link(box, 'View source', url);
    parent.appendChild(box);
  }

  function panel(parent, title) {
    const box = document.createElement('div');
    box.style.marginTop = '14px';
    box.style.padding = '12px';
    box.style.border = '1px solid #cbd5df';
    box.style.borderRadius = '8px';
    text(box, 'h3', title);
    parent.appendChild(box);
    return box;
  }

  function changedRow(address) {
    const cell = String(address || '').split('!').pop()
      .replace(/\$/g, '').toUpperCase();
    const match = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(cell);
    if (!match) return null;
    const number = letters => [...letters].reduce(
      (n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
    if (number(match[1]) > 4 || number(match[3] || match[1]) < 4) {
      return null;
    }
    const first = Number(match[2]);
    const last = Number(match[4] || match[2]);
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

  function showProfile(parent, candidate) {
    const id = String(candidate.wikidata_id || '');
    const g = candidate.global_profile || {};
    const m = candidate.malaysia_profile || {};
    const property = code => 'https://www.wikidata.org/wiki/' + id + '#' + code;

    const global = panel(parent, 'Global / Main Company Background');
    text(global, 'p', 'Wikidata-reported information; not independently verified.');
    field(global, 'E — Company Type', g.company_type, property('P31'));
    field(global, 'F — Company Headquarters', g.headquarters, property('P159'));
    field(global, 'G — Industry', g.industry, property('P452'));
    field(global, 'K — Company Website', g.websites, property('P856'));
    list(g.websites).forEach(url => link(global, 'Open reported website', url));
    field(global, 'Country (reported)', g.country, property('P17'));
    field(global, 'Country of origin (reported)', g.country_of_origin);

    const malaysia = panel(parent, 'Separate Malaysia-Specific Research');
    text(malaysia, 'p', 'A global company record alone does not verify Malaysian operations.');
    field(malaysia, 'Malaysia presence status', m.malaysia_presence_status);
    field(malaysia, 'Malaysia entity name', m.malaysia_company_name);
    field(malaysia, 'Malaysia headquarters', m.malaysia_headquarters);
    field(malaysia, 'Malaysia office', m.malaysia_office_address);
    field(malaysia, 'Malaysia website', m.malaysia_website);
    list(m.malaysia_website).forEach(url => link(malaysia, 'Open Malaysia website', url));
    field(malaysia, 'Malaysia careers', m.malaysia_careers_url);
    list(m.malaysia_careers_url).forEach(url => link(malaysia, 'Open Malaysia careers', url));

    const evidence = Array.isArray(m.malaysia_evidence) ? m.malaysia_evidence : [];
    if (evidence.length) {
      text(malaysia, 'h4', 'Malaysia research signals');
      evidence.forEach(item => {
        if (typeof item === 'string') {
          text(malaysia, 'p', item);
        } else if (item && typeof item === 'object') {
          text(malaysia, 'p', String(item.note || item.type || 'Research signal'));
          link(malaysia, 'View evidence source', item.source_url);
        }
      });
    } else {
      text(malaysia, 'p', 'No independent Malaysia-specific evidence supplied.');
    }
    if (typeof m.research_note === 'string') text(malaysia, 'p', m.research_note);
    text(malaysia, 'p', candidate.verified_malaysia_presence === true
      ? 'Worker reports Malaysia presence verified; review its sources.'
      : 'Malaysia presence NOT independently verified by this research.');
    text(parent, 'p', 'Preview only: E, F, G and K have NOT been changed.');
  }

  function noteFor(company, candidate) {
    const g = candidate.global_profile || {};
    const m = candidate.malaysia_profile || {};
    return [
      'UNVERIFIED COMPANY RESEARCH — V8',
      'Search: ' + company,
      'Candidate: ' + String(candidate.name || ''),
      'Description: ' + String(candidate.description || ''),
      'Wikidata: https://www.wikidata.org/wiki/' + candidate.wikidata_id,
      '',
      'GLOBAL/MAIN COMPANY (WIKIDATA, UNVERIFIED)',
      'Company type: ' + display(g.company_type),
      'Headquarters: ' + display(g.headquarters),
      'Industry: ' + display(g.industry),
      'Country: ' + display(g.country),
      'Website: ' + display(g.websites),
      '',
      'MALAYSIA-SPECIFIC (NOT INDEPENDENTLY VERIFIED)',
      'Presence: ' + display(m.malaysia_presence_status),
      'Entity name: ' + display(m.malaysia_company_name),
      'Malaysia office: ' + display(m.malaysia_office_address),
      'Malaysia website: ' + display(m.malaysia_website),
      'Malaysia careers: ' + display(m.malaysia_careers_url),
      'Review official sources before using this information.'
    ].join('\n');
  }

  async function approveAndSave(company, row, candidate, token, button, message) {
    if (token !== requestNumber) {
      message.textContent = 'Research is outdated. Nothing saved.';
      return;
    }
    const note = noteFor(company, candidate);
    if (note.length > 32000) {
      message.textContent = 'Research note too long. Nothing saved.';
      return;
    }
    button.disabled = true;
    button.textContent = 'Checking Excel...';
    try {
      const result = await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem(SHEET);
        const nameCell = sheet.getRange('D' + row);
        const noteCell = sheet.getRange(NOTE_COLUMN + row);
        nameCell.load('values');
        noteCell.load('values,formulas');
        await context.sync();

        if (token !== requestNumber) return 'Research changed. Nothing saved.';
        if (String(nameCell.values[0][0] ?? '').trim() !== company) {
          return 'Company name changed. Nothing saved.';
        }
        const existing = noteCell.values[0][0];
        const formula = noteCell.formulas[0][0];
        if ((existing !== null && existing !== undefined && String(existing).trim()) ||
            (typeof formula === 'string' && formula.startsWith('='))) {
          return NOTE_COLUMN + row + ' already has data or a formula. Nothing overwritten.';
        }
        noteCell.values = [[note]];
        await context.sync();
        return null;
      });
      if (result) {
        show('SAVE BLOCKED\n' + result);
        message.textContent = result;
        button.disabled = false;
        button.textContent = 'Approve & save note';
      } else {
        show('NOTE SAVED TO ' + NOTE_COLUMN + row + '\nE, F, G and K unchanged.');
        message.textContent = 'Research note saved. Details remain preview-only.';
        button.textContent = 'Saved to Excel';
      }
    } catch (error) {
      show('SAVE ERROR\n' + errorText(error));
      message.textContent = 'Save not confirmed.';
      button.disabled = false;
      button.textContent = 'Approve & save note';
    }
  }

  function renderCandidates(company, row, candidates, token) {
    clear();
    if (!results) return;
    text(results, 'h2', 'Review research for D' + row);
    text(results, 'p', 'Searched: ' + company);
    if (!candidates.length) {
      text(results, 'p', 'No candidates found. No cells changed.');
      return;
    }
    candidates.forEach(candidate => {
      const card = document.createElement('div');
      card.className = 'candidate';
      text(card, 'h3', candidate.name || 'Unnamed candidate');
      text(card, 'p', candidate.description || 'No description available');
      const id = String(candidate.wikidata_id || '');
      text(card, 'p', 'Wikidata ID: ' + id, 'small');
      if (/^Q\d+$/.test(id)) {
        link(card, 'Open Wikidata record',
          candidate.source_url || 'https://www.wikidata.org/wiki/' + id);
      }
      text(card, 'p', 'Candidate identity and Malaysia presence require review.', 'small');
      const select = document.createElement('button');
      select.type = 'button';
      select.textContent = 'Select for review';
      card.appendChild(select);
      const message = text(card, 'p', '', 'small');
      select.addEventListener('click', () => {
        if (token !== requestNumber) {
          show('Research outdated. Edit the company cell to search again.');
          return;
        }
        if (!/^Q\d+$/.test(id)) {
          message.textContent = 'Invalid Wikidata ID.';
          return;
        }
        select.disabled = true;
        select.textContent = 'Selected for review';
        card.style.border = '2px solid #13865c';
        showProfile(card, candidate);
        const approve = document.createElement('button');
        approve.type = 'button';
        approve.textContent = 'Approve & save note';
        approve.addEventListener('click', () =>
          approveAndSave(company, row, candidate, token, approve, message));
        card.appendChild(approve);
        show('V8 DETAILS READY\nGlobal and Malaysia-specific previews shown.\nNo company details written.');
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
      show('All candidates rejected. No cells changed.');
    });
    results.appendChild(reject);
  }

  async function research(company, row) {
    const token = ++requestNumber;
    clear();
    show('V8 RESEARCHING\nCell: D' + row + '\nCompany: ' + company);
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 20000);
    try {
      const response = await fetch(WORKER + '/research?company=' +
        encodeURIComponent(company), {
          headers: { Accept: 'application/json' }, signal: controller.signal
        });
      if (!response.ok) throw new Error('Research HTTP ' + response.status);
      const data = await response.json();
      if (token !== requestNumber) return;
      if (data.status === 'research_error') {
        throw new Error(data.message || 'Worker research failed');
      }
      const candidates = Array.isArray(data.candidates) ? data.candidates : [];
      show('V8 RESEARCH COMPLETE\nWorker: ' + (data.version || 'unknown') +
        '\nCandidates: ' + candidates.length + '\nSelect one to review.');
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
    if (location.multiple) {
      ++requestNumber;
      clear();
      show('Multiple company rows edited. Research one row at a time.');
      return;
    }
    try {
      const company = await readCompany(location.row);
      if (!company) {
        ++requestNumber;
        clear();
        show('D' + location.row + ' is blank.');
      } else {
        await research(company, location.row);
      }
    } catch (error) {
      show('CHANGE EVENT ERROR\n' + errorText(error));
    }
  }

  async function initialize() {
    if (active) {
      show('V8 READY\nMonitoring Company Master D7 onward.');
      return;
    }
    if (starting) return;
    starting = true;
    show('V8: Connecting to Excel...');
    try {
      if (typeof Office === 'undefined') {
        throw new Error('Office.js is not available.');
      }
      await Office.onReady();
      if (typeof Excel === 'undefined') {
        throw new Error('Open this add-in inside Excel Online.');
      }
      await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItemOrNullObject(SHEET);
        sheet.load('isNullObject');
        await context.sync();
        if (sheet.isNullObject) {
          throw new Error('Company Master worksheet not found.');
        }
        sheet.onChanged.add(onChanged);
        await context.sync();
      });
      active = true;
      show('V8 READY — GLOBAL + MALAYSIA RESEARCH\n' +
        'Monitoring Company Master column D from row 7.\n' +
        'Edit a company name to research.\n' +
        'E, F, G, K are preview-only. AP note saving requires approval.');
    } catch (error) {
      show('V8 CONNECTION ERROR\n' + errorText(error));
    } finally {
      starting = false;
    }
  }

  if (retry) retry.addEventListener('click', initialize);
  initialize();
})();
