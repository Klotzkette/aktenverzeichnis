/* ============================================================
   Aktenverzeichnis — Kanzlei-Tool
   Core application logic
   ============================================================ */

// ── In-memory storage (persists for session) ──
const memoryStore = {};
const safeStorage = {
  getItem(k) {
    try {
      const ls = window['local' + 'Storage'];
      if (ls) return ls.getItem(k);
    } catch { /* unavailable */ }
    return memoryStore[k] || null;
  },
  setItem(k, v) {
    try {
      const ls = window['local' + 'Storage'];
      if (ls) ls.setItem(k, v);
    } catch { /* unavailable */ }
    memoryStore[k] = v;
  },
};

// ── State ──────────────────────────────────────────────────
const state = {
  caseRef: '',
  caseTitle: '',
  caseDesc: '',
  hauptakte: [],   // { nr, blatt, datum, vorgang, kategorie, essentialia, anmKanzlei, anmMandant }
  chronologie: [], // { nr, datum, blatt, beteiligte, vorgang, anmKanzlei, anmMandant }
  personen: [],    // { nr, name, adresse, rolle, blatt, anmKanzlei, anmMandant }
  nextNrHA: 1,
  nextNrCH: 1,
  nextNrPE: 1,
};

// ── Category Map ───────────────────────────────────────────
const CATEGORIES = [
  'Klageschrift', 'Klageerwiderung', 'Schriftsatz', 'Bescheid',
  'Widerspruchsbescheid', 'Vollmacht', 'Gutachten', 'Rechnung',
  'Vertrag', 'Beschluss', 'Urteil', 'Protokoll', 'Korrespondenz',
  'Mahnung', 'Abrechnung', 'Anlage', 'Sonstige'
];

const CAT_CSS = {
  'Klageschrift': 'cat-klageschrift', 'Klageerwiderung': 'cat-klageschrift',
  'Bescheid': 'cat-bescheid', 'Widerspruchsbescheid': 'cat-bescheid',
  'Vollmacht': 'cat-vollmacht',
  'Gutachten': 'cat-gutachten',
  'Schriftsatz': 'cat-schriftsatz',
  'Rechnung': 'cat-rechnung', 'Mahnung': 'cat-rechnung', 'Abrechnung': 'cat-rechnung',
  'Vertrag': 'cat-vertrag',
  'Beschluss': 'cat-beschluss',
  'Urteil': 'cat-urteil',
  'Protokoll': 'cat-protokoll',
  'Korrespondenz': 'cat-korrespondenz',
};

function catClass(cat) {
  return CAT_CSS[cat] || 'cat-sonstige';
}

// ── Init ───────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  loadSettings();
  bindEvents();
});

function bindEvents() {
  // Settings modal
  const modal = document.getElementById('settings-modal');
  document.getElementById('btn-settings').addEventListener('click', () => modal.classList.remove('hidden'));
  document.getElementById('settings-close').addEventListener('click', () => modal.classList.add('hidden'));
  modal.querySelector('.modal-backdrop').addEventListener('click', () => modal.classList.add('hidden'));
  document.getElementById('toggle-key').addEventListener('click', () => {
    const inp = document.getElementById('api-key');
    inp.type = inp.type === 'password' ? 'text' : 'password';
  });
  document.getElementById('save-settings').addEventListener('click', saveSettings);

  // Create case
  document.getElementById('btn-create-case').addEventListener('click', createCase);
  // Enter key in case fields
  ['case-ref', 'case-title'].forEach(id => {
    document.getElementById(id).addEventListener('keydown', (e) => {
      if (e.key === 'Enter') createCase();
    });
  });

  // Import old table from new-case screen
  document.getElementById('btn-import-old').addEventListener('click', () => {
    document.getElementById('xlsx-import-input').click();
  });

  // Tab switching
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => switchTab(tab.dataset.tab));
  });

  // Drop zone
  const dz = document.getElementById('drop-zone');
  const fi = document.getElementById('file-input');
  dz.addEventListener('dragover', (e) => { e.preventDefault(); dz.classList.add('drag-over'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
  dz.addEventListener('drop', (e) => {
    e.preventDefault();
    dz.classList.remove('drag-over');
    if (e.dataTransfer.files.length) handleFiles(Array.from(e.dataTransfer.files));
  });
  dz.addEventListener('click', (e) => {
    if (e.target.id !== 'btn-browse') fi.click();
  });
  document.getElementById('btn-browse').addEventListener('click', (e) => {
    e.stopPropagation();
    fi.click();
  });
  fi.addEventListener('change', (e) => {
    if (e.target.files.length) handleFiles(Array.from(e.target.files));
    fi.value = '';
  });

  // Add rows
  document.getElementById('btn-add-row-ha').addEventListener('click', () => {
    addHauptakteRow({});
    renderHauptakte();
  });
  document.getElementById('btn-add-row-ch').addEventListener('click', () => {
    addChronologieRow({});
    renderChronologie();
  });
  document.getElementById('btn-add-row-pe').addEventListener('click', () => {
    addPersonRow({});
    renderPersonen();
  });

  // Generate chronology
  document.getElementById('btn-gen-chrono').addEventListener('click', generateChronologie);

  // Import/Export XLSX
  document.getElementById('btn-import-xlsx').addEventListener('click', () => {
    document.getElementById('xlsx-import-input').click();
  });
  document.getElementById('xlsx-import-input').addEventListener('change', handleXLSXImport);
  document.getElementById('btn-export-xlsx').addEventListener('click', exportXLSX);
}

// ── Settings ───────────────────────────────────────────────
function loadSettings() {
  try {
    const provider = safeStorage.getItem('av_provider') || 'openai';
    const key = safeStorage.getItem('av_apikey') || '';
    document.getElementById('ai-provider').value = provider;
    document.getElementById('api-key').value = key;
  } catch (e) {}
}

function saveSettings() {
  try {
    safeStorage.setItem('av_provider', document.getElementById('ai-provider').value);
    safeStorage.setItem('av_apikey', document.getElementById('api-key').value.trim());
    const st = document.getElementById('settings-status');
    st.textContent = 'Gespeichert';
    st.style.color = 'var(--success)';
    setTimeout(() => { st.textContent = ''; }, 2000);
  } catch (e) {
    document.getElementById('settings-status').textContent = 'Fehler beim Speichern';
  }
}

function getAPIKey() {
  return document.getElementById('api-key').value.trim() || safeStorage.getItem('av_apikey') || '';
}

function getProvider() {
  return document.getElementById('ai-provider').value || safeStorage.getItem('av_provider') || 'openai';
}

// ── Case Creation ──────────────────────────────────────────
function createCase() {
  const ref = document.getElementById('case-ref').value.trim();
  const title = document.getElementById('case-title').value.trim();
  if (!ref && !title) {
    alert('Bitte mindestens Aktenzeichen oder Bezeichnung eingeben.');
    return;
  }
  state.caseRef = ref;
  state.caseTitle = title;
  state.caseDesc = document.getElementById('case-desc').value.trim();

  showCaseScreen();
}

function showCaseScreen() {
  document.getElementById('screen-new').classList.add('hidden');
  document.getElementById('screen-case').classList.remove('hidden');

  const badge = document.getElementById('case-badge');
  const label = [state.caseRef, state.caseTitle].filter(Boolean).join(' — ');
  if (label) {
    badge.textContent = label;
    badge.classList.add('visible');
  }

  updateCounts();
}

// ── Tab Switching ──────────────────────────────────────────
function switchTab(name) {
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.dataset.tab === name));
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.toggle('active', p.id === `panel-${name}`));
}

// ── File Handling ──────────────────────────────────────────
async function handleFiles(files) {
  const apiKey = getAPIKey();
  if (!apiKey) {
    document.getElementById('btn-settings').click();
    alert('Bitte zuerst einen API-Schlüssel in den Einstellungen hinterlegen.');
    return;
  }

  const bar = document.getElementById('processing-bar');
  const progress = document.getElementById('processing-progress');
  const text = document.getElementById('processing-text');
  bar.classList.remove('hidden');

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const pct = Math.round(((i) / files.length) * 100);
    progress.style.width = pct + '%';
    text.textContent = `${file.name} wird analysiert… (${i + 1}/${files.length})`;

    try {
      const content = await extractFileContent(file);
      const analysis = await analyzeDocument(content, file.name);

      // Add to Hauptakte
      if (analysis.entries) {
        for (const entry of analysis.entries) {
          addHauptakteRow({
            datum: entry.datum || '',
            vorgang: entry.vorgang || file.name.replace(/\.[^.]+$/, ''),
            kategorie: entry.kategorie || 'Sonstige',
            essentialia: entry.essentialia || '',
          });
        }
      } else {
        addHauptakteRow({
          datum: analysis.datum || '',
          vorgang: analysis.vorgang || file.name.replace(/\.[^.]+$/, ''),
          kategorie: analysis.kategorie || 'Sonstige',
          essentialia: analysis.essentialia || '',
        });
      }

      // Add persons if detected
      if (analysis.personen && analysis.personen.length > 0) {
        for (const p of analysis.personen) {
          // Avoid duplicates
          const exists = state.personen.some(
            ep => ep.name.toLowerCase() === (p.name || '').toLowerCase()
          );
          if (!exists && p.name) {
            addPersonRow({
              name: p.name,
              adresse: p.adresse || '',
              rolle: p.rolle || '',
              blatt: String(state.nextNrHA - 1),
            });
          }
        }
      }
    } catch (err) {
      console.error(`Error processing ${file.name}:`, err);
      // Still add with filename
      addHauptakteRow({
        vorgang: file.name.replace(/\.[^.]+$/, ''),
        kategorie: 'Sonstige',
        essentialia: `Fehler: ${err.message}`,
      });
    }
  }

  progress.style.width = '100%';
  text.textContent = `${files.length} Dokument${files.length > 1 ? 'e' : ''} verarbeitet`;
  setTimeout(() => { bar.classList.add('hidden'); progress.style.width = '0%'; }, 2000);

  renderHauptakte();
  renderPersonen();
  updateCounts();
  switchTab('hauptakte');
}

// ── File Content Extraction ────────────────────────────────
async function extractFileContent(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  const imageExts = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp', 'tif', 'tiff'];

  if (ext === 'pdf') {
    return await extractPDF(file);
  } else if (ext === 'docx' || ext === 'doc') {
    return await extractDOCX(file);
  } else if (ext === 'xlsx' || ext === 'xls') {
    return extractXLSXContent(file);
  } else if (imageExts.includes(ext)) {
    return await fileToBase64(file);
  }
  throw new Error('Nicht unterstütztes Format: ' + ext);
}

async function extractPDF(file) {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  let text = '';
  for (let i = 1; i <= Math.min(pdf.numPages, 20); i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    text += content.items.map(item => item.str).join(' ') + '\n';
  }
  return { type: 'text', content: text };
}

async function extractDOCX(file) {
  const buf = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer: buf });
  return { type: 'text', content: result.value };
}

function extractXLSXContent(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        let text = '';
        wb.SheetNames.forEach(name => {
          text += `--- ${name} ---\n${XLSX.utils.sheet_to_csv(wb.Sheets[name])}\n\n`;
        });
        resolve({ type: 'text', content: text });
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve({ type: 'image', content: reader.result });
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ── AI Analysis ────────────────────────────────────────────
async function analyzeDocument(fileData, fileName) {
  const provider = getProvider();
  const apiKey = getAPIKey();

  const systemPrompt = `Du bist ein juristischer Assistent für Aktenführung in einer deutschen Kanzlei.
Analysiere das bereitgestellte Dokument und extrahiere folgende Informationen:

1. **Datum** des Dokuments (Format: TT.MM.JJJJ)
2. **Vorgang** — kurze Beschreibung (z.B. "Klageschrift des Klägers", "Vollmacht", "Gutachten Prof. Dr. Meier")
3. **Kategorie** — eine der folgenden: ${CATEGORIES.join(', ')}
4. **Essentialia** — die wichtigsten inhaltlichen Punkte in 1-2 Sätzen
5. **Personen** — alle genannten Personen mit Name, Adresse (falls erkennbar) und Prozessrolle (Kläger, Beklagter, Zeuge, Sachverständiger, Richter, Rechtsanwalt, etc.)

Wenn das Dokument mehrere separate Dokumente/Schriftstücke enthält, gib für jedes einen eigenen Eintrag zurück.

Der Dateiname lautet: "${fileName}"

Antworte AUSSCHLIESSLICH mit validem JSON:
{
  "entries": [
    {
      "datum": "15.03.2026",
      "vorgang": "Klageschrift des Klägers",
      "kategorie": "Klageschrift",
      "essentialia": "Räumungsklage wegen Mietrückständen i.H.v. 5.400 EUR für den Zeitraum Jan-Jun 2026"
    }
  ],
  "personen": [
    { "name": "Max Müller", "adresse": "Berliner Str. 1, 10115 Berlin", "rolle": "Kläger" },
    { "name": "Hans Schmidt", "adresse": "", "rolle": "Beklagter" }
  ]
}

Wenn du Informationen nicht erkennen kannst, lasse das Feld leer. Erfinde NICHTS.`;

  const userMessage = fileData.type === 'image'
    ? 'Analysiere dieses Dokument/Bild und extrahiere die Akteninformationen.'
    : `Analysiere das folgende Dokument und extrahiere die Akteninformationen.\n\n--- DOKUMENTINHALT ---\n${fileData.content.substring(0, 30000)}`;

  if (provider === 'openai') {
    return await callOpenAI(apiKey, systemPrompt, userMessage, fileData);
  } else {
    return await callAnthropic(apiKey, systemPrompt, userMessage, fileData);
  }
}

async function callOpenAI(apiKey, system, user, fileData) {
  const messages = [{ role: 'system', content: system }];

  if (fileData.type === 'image') {
    messages.push({
      role: 'user',
      content: [
        { type: 'text', text: user },
        { type: 'image_url', image_url: { url: fileData.content } }
      ]
    });
  } else {
    messages.push({ role: 'user', content: user });
  }

  const res = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` },
    body: JSON.stringify({
      model: 'gpt-4o', messages, temperature: 0.1, max_tokens: 8000,
      response_format: { type: 'json_object' }
    })
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || res.statusText);
  }
  const data = await res.json();
  return JSON.parse(data.choices[0].message.content);
}

async function callAnthropic(apiKey, system, user, fileData) {
  const messages = [];

  if (fileData.type === 'image') {
    const match = fileData.content.match(/^data:(image\/[a-z+]+);base64,(.+)$/i);
    if (!match) throw new Error('Ungültiges Bildformat');
    messages.push({
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: match[1], data: match[2] } },
        { type: 'text', text: user }
      ]
    });
  } else {
    messages.push({ role: 'user', content: user });
  }

  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true'
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-20250514', max_tokens: 8000, system, messages
    })
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || res.statusText);
  }
  const data = await res.json();
  const text = data.content[0].text;
  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) throw new Error('Keine gültige JSON-Antwort erhalten.');
  return JSON.parse(jsonMatch[0]);
}

// ── Data Management ────────────────────────────────────────
function addHauptakteRow(data) {
  state.hauptakte.push({
    nr: state.nextNrHA++,
    blatt: data.blatt || '',
    datum: data.datum || '',
    vorgang: data.vorgang || '',
    kategorie: data.kategorie || '',
    essentialia: data.essentialia || '',
    anmKanzlei: data.anmKanzlei || '',
    anmMandant: data.anmMandant || '',
  });
}

function addChronologieRow(data) {
  state.chronologie.push({
    nr: state.nextNrCH++,
    datum: data.datum || '',
    blatt: data.blatt || '',
    beteiligte: data.beteiligte || '',
    vorgang: data.vorgang || '',
    anmKanzlei: data.anmKanzlei || '',
    anmMandant: data.anmMandant || '',
  });
}

function addPersonRow(data) {
  state.personen.push({
    nr: state.nextNrPE++,
    name: data.name || '',
    adresse: data.adresse || '',
    rolle: data.rolle || '',
    blatt: data.blatt || '',
    anmKanzlei: data.anmKanzlei || '',
    anmMandant: data.anmMandant || '',
  });
}

function removeHauptakteRow(nr) {
  state.hauptakte = state.hauptakte.filter(r => r.nr !== nr);
  renderHauptakte();
  updateCounts();
}

function removeChronologieRow(nr) {
  state.chronologie = state.chronologie.filter(r => r.nr !== nr);
  renderChronologie();
  updateCounts();
}

function removePersonRow(nr) {
  state.personen = state.personen.filter(r => r.nr !== nr);
  renderPersonen();
  updateCounts();
}

// ── Rendering ──────────────────────────────────────────────
function renderHauptakte() {
  const tbody = document.getElementById('tbody-hauptakte');
  const empty = document.getElementById('empty-hauptakte');
  tbody.innerHTML = '';

  if (state.hauptakte.length === 0) {
    empty.classList.remove('hidden');
    return;
  }
  empty.classList.add('hidden');

  state.hauptakte.forEach((row, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="col-nr">${idx + 1}</td>
      <td contenteditable="true" data-field="blatt" data-nr="${row.nr}" data-placeholder="—">${esc(row.blatt)}</td>
      <td contenteditable="true" data-field="datum" data-nr="${row.nr}" data-placeholder="TT.MM.JJJJ">${esc(row.datum)}</td>
      <td contenteditable="true" data-field="vorgang" data-nr="${row.nr}" data-placeholder="Vorgang…">${esc(row.vorgang)}</td>
      <td data-field="kategorie" data-nr="${row.nr}">
        <select class="cat-select" data-nr="${row.nr}">
          ${CATEGORIES.map(c => `<option value="${c}" ${c === row.kategorie ? 'selected' : ''}>${c}</option>`).join('')}
        </select>
      </td>
      <td contenteditable="true" data-field="essentialia" data-nr="${row.nr}" data-placeholder="Wesentliche Inhalte…">${esc(row.essentialia)}</td>
      <td contenteditable="true" data-field="anmKanzlei" data-nr="${row.nr}" data-placeholder="—">${esc(row.anmKanzlei)}</td>
      <td contenteditable="true" data-field="anmMandant" data-nr="${row.nr}" data-placeholder="—">${esc(row.anmMandant)}</td>
      <td class="col-actions">
        <button class="btn-danger-icon" title="Entfernen" onclick="removeHauptakteRow(${row.nr})">
          <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  // Bind blur events for saving
  tbody.querySelectorAll('[contenteditable]').forEach(cell => {
    cell.addEventListener('blur', () => {
      const nr = parseInt(cell.dataset.nr);
      const field = cell.dataset.field;
      const row = state.hauptakte.find(r => r.nr === nr);
      if (row) row[field] = cell.textContent.trim();
    });
  });

  // Bind category selects
  tbody.querySelectorAll('.cat-select').forEach(sel => {
    sel.addEventListener('change', () => {
      const nr = parseInt(sel.dataset.nr);
      const row = state.hauptakte.find(r => r.nr === nr);
      if (row) row.kategorie = sel.value;
    });
  });
}

function renderChronologie() {
  const tbody = document.getElementById('tbody-chronologie');
  const empty = document.getElementById('empty-chronologie');
  tbody.innerHTML = '';

  if (state.chronologie.length === 0) {
    empty.classList.remove('hidden');
    return;
  }
  empty.classList.add('hidden');

  state.chronologie.forEach((row, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="col-nr">${idx + 1}</td>
      <td contenteditable="true" data-field="datum" data-nr="${row.nr}" data-placeholder="TT.MM.JJJJ">${esc(row.datum)}</td>
      <td contenteditable="true" data-field="blatt" data-nr="${row.nr}" data-placeholder="—">${esc(row.blatt)}</td>
      <td contenteditable="true" data-field="beteiligte" data-nr="${row.nr}" data-placeholder="—">${esc(row.beteiligte)}</td>
      <td contenteditable="true" data-field="vorgang" data-nr="${row.nr}" data-placeholder="Vorgang…">${esc(row.vorgang)}</td>
      <td contenteditable="true" data-field="anmKanzlei" data-nr="${row.nr}" data-placeholder="—">${esc(row.anmKanzlei)}</td>
      <td contenteditable="true" data-field="anmMandant" data-nr="${row.nr}" data-placeholder="—">${esc(row.anmMandant)}</td>
      <td class="col-actions">
        <button class="btn-danger-icon" title="Entfernen" onclick="removeChronologieRow(${row.nr})">
          <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  tbody.querySelectorAll('[contenteditable]').forEach(cell => {
    cell.addEventListener('blur', () => {
      const nr = parseInt(cell.dataset.nr);
      const field = cell.dataset.field;
      const row = state.chronologie.find(r => r.nr === nr);
      if (row) row[field] = cell.textContent.trim();
    });
  });
}

function renderPersonen() {
  const tbody = document.getElementById('tbody-personen');
  const empty = document.getElementById('empty-personen');
  tbody.innerHTML = '';

  if (state.personen.length === 0) {
    empty.classList.remove('hidden');
    return;
  }
  empty.classList.add('hidden');

  state.personen.forEach((row, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="col-nr">${idx + 1}</td>
      <td contenteditable="true" data-field="name" data-nr="${row.nr}" data-placeholder="Name…">${esc(row.name)}</td>
      <td contenteditable="true" data-field="adresse" data-nr="${row.nr}" data-placeholder="Adresse…">${esc(row.adresse)}</td>
      <td contenteditable="true" data-field="rolle" data-nr="${row.nr}" data-placeholder="Rolle…">${esc(row.rolle)}</td>
      <td contenteditable="true" data-field="blatt" data-nr="${row.nr}" data-placeholder="—">${esc(row.blatt)}</td>
      <td contenteditable="true" data-field="anmKanzlei" data-nr="${row.nr}" data-placeholder="—">${esc(row.anmKanzlei)}</td>
      <td contenteditable="true" data-field="anmMandant" data-nr="${row.nr}" data-placeholder="—">${esc(row.anmMandant)}</td>
      <td class="col-actions">
        <button class="btn-danger-icon" title="Entfernen" onclick="removePersonRow(${row.nr})">
          <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  tbody.querySelectorAll('[contenteditable]').forEach(cell => {
    cell.addEventListener('blur', () => {
      const nr = parseInt(cell.dataset.nr);
      const field = cell.dataset.field;
      const row = state.personen.find(r => r.nr === nr);
      if (row) row[field] = cell.textContent.trim();
    });
  });
}

function updateCounts() {
  document.getElementById('count-hauptakte').textContent = state.hauptakte.length;
  document.getElementById('count-chronologie').textContent = state.chronologie.length;
  document.getElementById('count-personen').textContent = state.personen.length;
}

function esc(str) {
  if (!str) return '';
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

// ── Chronologie Generation ─────────────────────────────────
function generateChronologie() {
  // Build chronologie from hauptakte entries that have dates
  const entries = state.hauptakte
    .filter(r => r.datum)
    .map(r => ({
      datum: r.datum,
      blatt: r.blatt,
      beteiligte: '', // Could be enhanced with AI
      vorgang: r.vorgang,
      anmKanzlei: r.anmKanzlei,
      anmMandant: r.anmMandant,
    }));

  // Sort by date
  entries.sort((a, b) => {
    const da = parseDE(a.datum);
    const db = parseDE(b.datum);
    return (da || 0) - (db || 0);
  });

  // Reset chronologie
  state.chronologie = [];
  state.nextNrCH = 1;
  entries.forEach(e => addChronologieRow(e));

  renderChronologie();
  updateCounts();
  switchTab('chronologie');
}

function parseDE(dateStr) {
  if (!dateStr) return null;
  const m = dateStr.match(/(\d{1,2})\.(\d{1,2})\.(\d{2,4})/);
  if (!m) return null;
  const year = m[3].length === 2 ? 2000 + parseInt(m[3]) : parseInt(m[3]);
  return new Date(year, parseInt(m[2]) - 1, parseInt(m[1]));
}

// ── XLSX Import ────────────────────────────────────────────
function handleXLSXImport(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      const wb = XLSX.read(ev.target.result, { type: 'array' });

      // Try to find sheets by name patterns
      const sheetNames = wb.SheetNames;

      // Parse Hauptakte / Aktenübersicht
      const haSheet = findSheet(sheetNames, ['hauptakte', 'akten', 'übersicht', 'overview']);
      if (haSheet) {
        const data = XLSX.utils.sheet_to_json(wb.Sheets[haSheet], { header: 1 });
        importHauptakteData(data);
      }

      // Parse Chronologie
      const chSheet = findSheet(sheetNames, ['chrono', 'zeitl']);
      if (chSheet) {
        const data = XLSX.utils.sheet_to_json(wb.Sheets[chSheet], { header: 1 });
        importChronologieData(data);
      }

      // Parse Personen
      const peSheet = findSheet(sheetNames, ['person', 'beteili', 'partei']);
      if (peSheet) {
        const data = XLSX.utils.sheet_to_json(wb.Sheets[peSheet], { header: 1 });
        importPersonenData(data);
      }

      // If nothing found by name, try the first sheets
      if (!haSheet && sheetNames.length >= 2) {
        // Skip first sheet (usually cover page), import second as Hauptakte
        const data = XLSX.utils.sheet_to_json(wb.Sheets[sheetNames[1]], { header: 1 });
        importHauptakteData(data);
      }

      // Extract case info from Vorblatt if exists
      const coverSheet = findSheet(sheetNames, ['vorblatt', 'deckblatt', 'cover']);
      if (coverSheet) {
        const data = XLSX.utils.sheet_to_json(wb.Sheets[coverSheet], { header: 1 });
        extractCaseInfo(data);
      }

      renderHauptakte();
      renderChronologie();
      renderPersonen();
      updateCounts();

      // Show case screen if not already
      if (!state.caseRef && !state.caseTitle) {
        state.caseRef = 'Importiert';
        state.caseTitle = file.name.replace(/\.[^.]+$/, '');
      }
      showCaseScreen();

    } catch (err) {
      console.error('XLSX Import error:', err);
      alert('Fehler beim Importieren: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
  e.target.value = '';
}

function findSheet(names, keywords) {
  return names.find(n => keywords.some(k => n.toLowerCase().includes(k)));
}

function importHauptakteData(data) {
  if (data.length < 2) return; // Need at least header + 1 row
  const headers = data[0].map(h => String(h || '').toLowerCase());

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(c => !c)) continue;

    addHauptakteRow({
      blatt: findCol(headers, row, ['blatt']),
      datum: findCol(headers, row, ['datum', 'date']),
      vorgang: findCol(headers, row, ['vorgang', 'dokument', 'bezeichnung', 'beschreibung']),
      kategorie: findCol(headers, row, ['kategorie', 'art', 'typ', 'type']),
      essentialia: findCol(headers, row, ['essentialia', 'inhalt', 'zusammenfassung', 'content']),
      anmKanzlei: findCol(headers, row, ['anm. kanzlei', 'kanzlei', 'anmerkung kanzlei']),
      anmMandant: findCol(headers, row, ['anm. mandant', 'mandant', 'anmerkung mandant']),
    });
  }
}

function importChronologieData(data) {
  if (data.length < 2) return;
  const headers = data[0].map(h => String(h || '').toLowerCase());

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(c => !c)) continue;

    addChronologieRow({
      datum: findCol(headers, row, ['datum', 'date']),
      blatt: findCol(headers, row, ['blatt']),
      beteiligte: findCol(headers, row, ['beteiligte', 'partei', 'person']),
      vorgang: findCol(headers, row, ['vorgang', 'ereignis', 'beschreibung']),
      anmKanzlei: findCol(headers, row, ['anm. kanzlei', 'kanzlei']),
      anmMandant: findCol(headers, row, ['anm. mandant', 'mandant']),
    });
  }
}

function importPersonenData(data) {
  if (data.length < 2) return;
  const headers = data[0].map(h => String(h || '').toLowerCase());

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(c => !c)) continue;

    addPersonRow({
      name: findCol(headers, row, ['name', 'person', 'beteiligte']),
      adresse: findCol(headers, row, ['adresse', 'anschrift', 'address']),
      rolle: findCol(headers, row, ['rolle', 'funktion', 'prozessrolle', 'role']),
      blatt: findCol(headers, row, ['blatt']),
      anmKanzlei: findCol(headers, row, ['anm. kanzlei', 'kanzlei']),
      anmMandant: findCol(headers, row, ['anm. mandant', 'mandant']),
    });
  }
}

function extractCaseInfo(data) {
  // Try to find case reference and title from cover sheet
  for (const row of data) {
    if (!row) continue;
    for (let i = 0; i < row.length; i++) {
      const val = String(row[i] || '');
      if (val.toLowerCase().includes('aktenzeichen') || val.toLowerCase().includes('az.')) {
        state.caseRef = String(row[i + 1] || row[i] || '').replace(/aktenzeichen:?\s*/i, '').trim();
      }
      if (val.toLowerCase().includes('mandant') || val.toLowerCase().includes('sache')) {
        state.caseTitle = String(row[i + 1] || '').trim();
      }
    }
  }
}

function findCol(headers, row, keywords) {
  for (const kw of keywords) {
    const idx = headers.findIndex(h => h.includes(kw));
    if (idx >= 0 && row[idx] !== undefined && row[idx] !== null) {
      return String(row[idx]);
    }
  }
  return '';
}

// ── XLSX Export ─────────────────────────────────────────────
function exportXLSX() {
  const wb = XLSX.utils.book_new();

  // 1. Vorblatt (Cover Page)
  const coverData = [
    [''],
    ['Aktenverzeichnis'],
    [''],
    ['Aktenzeichen:', state.caseRef],
    ['Bezeichnung:', state.caseTitle],
    ['Beschreibung:', state.caseDesc],
    [''],
    ['Erstellt am:', new Date().toLocaleDateString('de-DE')],
    ['Dokumente:', String(state.hauptakte.length)],
    ['Personen:', String(state.personen.length)],
  ];
  const wsCover = XLSX.utils.aoa_to_sheet(coverData);
  wsCover['!cols'] = [{ wch: 18 }, { wch: 50 }];
  // Merge title cell
  wsCover['!merges'] = [{ s: { r: 1, c: 0 }, e: { r: 1, c: 1 } }];
  XLSX.utils.book_append_sheet(wb, wsCover, 'Vorblatt');

  // 2. Aktenübersicht Hauptakte
  const haHeaders = ['#', 'Blatt', 'Datum', 'Vorgang', 'Kategorie', 'Essentialia', 'Anm. Kanzlei', 'Anm. Mandant'];
  const haData = [haHeaders, ...state.hauptakte.map((r, i) => [
    i + 1, r.blatt, r.datum, r.vorgang, r.kategorie, r.essentialia, r.anmKanzlei, r.anmMandant
  ])];
  const wsHA = XLSX.utils.aoa_to_sheet(haData);
  wsHA['!cols'] = [
    { wch: 5 }, { wch: 8 }, { wch: 12 }, { wch: 35 }, { wch: 18 }, { wch: 40 }, { wch: 25 }, { wch: 25 }
  ];
  XLSX.utils.book_append_sheet(wb, wsHA, 'Aktenübersicht Hauptakte');

  // 3. Chronologie
  const chHeaders = ['#', 'Datum', 'Blatt', 'Beteiligte', 'Vorgang', 'Anm. Kanzlei', 'Anm. Mandant'];
  const chData = [chHeaders, ...state.chronologie.map((r, i) => [
    i + 1, r.datum, r.blatt, r.beteiligte, r.vorgang, r.anmKanzlei, r.anmMandant
  ])];
  const wsCH = XLSX.utils.aoa_to_sheet(chData);
  wsCH['!cols'] = [
    { wch: 5 }, { wch: 12 }, { wch: 8 }, { wch: 25 }, { wch: 40 }, { wch: 25 }, { wch: 25 }
  ];
  XLSX.utils.book_append_sheet(wb, wsCH, 'Chronologie');

  // 4. Personenübersicht
  const peHeaders = ['#', 'Name', 'Adresse', 'Prozessrolle', 'Blatt', 'Anm. Kanzlei', 'Anm. Mandant'];
  const peData = [peHeaders, ...state.personen.map((r, i) => [
    i + 1, r.name, r.adresse, r.rolle, r.blatt, r.anmKanzlei, r.anmMandant
  ])];
  const wsPE = XLSX.utils.aoa_to_sheet(peData);
  wsPE['!cols'] = [
    { wch: 5 }, { wch: 25 }, { wch: 35 }, { wch: 18 }, { wch: 8 }, { wch: 25 }, { wch: 25 }
  ];
  XLSX.utils.book_append_sheet(wb, wsPE, 'Personenübersicht');

  // Generate filename
  const fileName = `Aktenverzeichnis_${state.caseRef || 'Export'}_${new Date().toISOString().slice(0,10)}.xlsx`
    .replace(/[\/\\:*?"<>|]/g, '_');

  XLSX.writeFile(wb, fileName);
}

// ── Make functions globally accessible for onclick handlers ──
window.removeHauptakteRow = removeHauptakteRow;
window.removeChronologieRow = removeChronologieRow;
window.removePersonRow = removePersonRow;
