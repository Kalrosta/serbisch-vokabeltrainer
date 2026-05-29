/**
 * Serbisch B2 - Main App
 * Vanilla JS, ES Modules.
 */
import { fsrsNewCard, fsrsReview, fsrsSkip, fsrsRedistribute, isDue } from "./fsrs.js";

// ============ CONSTANTS ============
const DATA_VERSION_KEY = "srb_data_version";
const PROGRESS_KEY = "srb_progress_v1";
const SETTINGS_KEY = "srb_settings_v1";
const ERRORS_KEY = "srb_errors_v1";
const NOTES_KEY = "srb_notes_v1";
const META_KEY = "srb_meta_v1";
const DAILY_NEW_DEFAULT = 15;
const BACKUP_REMIND_DAYS = 7;
const BACKUP_REMIND_NEW = 100;
const DAY_MS = 86400000;

// ============ STATE ============
const state = {
  view: "home",
  words: [],         // [{ id, n, de, sl, sc, f, wa, fm, tb }, ...]
  themen: [],        // ["Arbeit & Organisation", ...]
  examples: {},      // { "de|sl": { ipf: [], pf: [] } }
  forms: {},         // { "de|sl": { ...form info... } }
  progress: {},      // { id: card }
  errors: [],        // [{ id, comment, ts }]
  notes: {},         // { wordId: noteString }
  meta: {            // misc app meta
    lastBackup: 0,
    learnedSinceBackup: 0,
    todayKey: "",    // YYYY-MM-DD of last session
  },
  settings: {
    direction: "de2sr",
    themen: null,    // null = all
    filters: ["a", "b", "c"],
    dailyNew: DAILY_NEW_DEFAULT,
  },
  session: null,     // { queue: [ids], idx, revealed, results, origCount, again: Set }
  loadError: null,
};

// ============ STORAGE ============
function loadStored() {
  try { state.progress = JSON.parse(localStorage.getItem(PROGRESS_KEY)) || {}; } catch { state.progress = {}; }
  try { state.errors = JSON.parse(localStorage.getItem(ERRORS_KEY)) || []; } catch { state.errors = []; }
  try { state.notes = JSON.parse(localStorage.getItem(NOTES_KEY)) || {}; } catch { state.notes = {}; }
  try { state.meta = { ...state.meta, ...(JSON.parse(localStorage.getItem(META_KEY)) || {}) }; } catch {}
  try {
    const s = JSON.parse(localStorage.getItem(SETTINGS_KEY));
    if (s) Object.assign(state.settings, s);
  } catch {}
}

function saveProgress() { localStorage.setItem(PROGRESS_KEY, JSON.stringify(state.progress)); }
function saveErrors()   { localStorage.setItem(ERRORS_KEY, JSON.stringify(state.errors)); }
function saveNotes()    { localStorage.setItem(NOTES_KEY, JSON.stringify(state.notes)); }
function saveMeta()     { localStorage.setItem(META_KEY, JSON.stringify(state.meta)); }
function saveSettings() { localStorage.setItem(SETTINGS_KEY, JSON.stringify(state.settings)); }

// ============ DATA LOADING ============
async function loadData() {
  try {
    const [wordsRes, examplesRes, formsRes] = await Promise.all([
      fetch("./data/words.json"),
      fetch("./data/examples.json"),
      fetch("./data/forms.json").catch(() => ({ ok: false })),
    ]);
    if (!wordsRes.ok) throw new Error("words.json konnte nicht geladen werden");
    const wordsData = await wordsRes.json();
    state.words = wordsData.words;
    state.themen = wordsData.themen;

    // Daten-Versionierung: wenn xlsx neu gebuildet wurde, alle progress.skip behalten,
    // aber unbekannte IDs (Wörter die nicht mehr existieren) werden später ignoriert.
    const storedVersion = localStorage.getItem(DATA_VERSION_KEY);
    if (storedVersion !== wordsData.version) {
      localStorage.setItem(DATA_VERSION_KEY, wordsData.version);
    }

    if (examplesRes.ok) state.examples = await examplesRes.json();
    if (formsRes && formsRes.ok) state.forms = await formsRes.json();
  } catch (e) {
    state.loadError = e.message || String(e);
    console.error("Data load failed:", e);
  }
}

// ============ HELPERS ============
const $ = (id) => document.getElementById(id);

function el(html) {
  const t = document.createElement("template");
  t.innerHTML = html.trim();
  return t.content.firstChild;
}

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[c]));
}

function todayKey(now = Date.now()) {
  const d = new Date(now);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function wordKey(w) {
  return `${w.de}|${w.sl}`;
}

function getCard(wordId) {
  let c = state.progress[wordId];
  if (!c) {
    c = fsrsNewCard();
    state.progress[wordId] = c;
  }
  return c;
}

function passesFilter(w) {
  if (!state.settings.filters.includes(w.f)) return false;
  if (state.settings.themen && !state.settings.themen.includes(w.tb)) return false;
  return true;
}

function getStats() {
  const now = Date.now();
  const candidates = state.words.filter(passesFilter);
  let due = 0, unseen = 0, learning = 0, mature = 0;
  for (const w of candidates) {
    const c = state.progress[w.id];
    if (!c || c.state === "new") {
      due += 1;
      unseen += 1;
      continue;
    }
    if (c.due <= now) due += 1;
    if (c.S < 21) learning += 1;
    else mature += 1;
  }
  return { due, unseen, learning, mature, total: candidates.length };
}

// ============ ICONS (inline SVG) ============
const ICONS = {
  back: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M19 12H5M12 19l-7-7 7-7"/></svg>',
  gear: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>',
  download: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>',
  skip: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><polyline points="13 17 18 12 13 7"/><polyline points="6 17 11 12 6 7"/></svg>',
  flag: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 15s1-1 4-1 5 2 8 2 4-1 4-1V3s-1 1-4 1-5-2-8-2-4 1-4 1z"/><line x1="4" y1="22" x2="4" y2="15"/></svg>',
  edit: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20h9"/><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/></svg>',
  bulb: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M9 18h6"/><path d="M10 22h4"/><path d="M12 2a7 7 0 0 0-4 12.7c.7.5 1 1.3 1 2.1V18h6v-1.2c0-.8.3-1.6 1-2.1A7 7 0 0 0 12 2z"/></svg>',
  close: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>',
};

// ============ RENDER ============
function render() {
  const app = $("app");
  app.innerHTML = "";
  if (state.loadError) {
    app.appendChild(el(`
      <div class="center-prose">
        <p class="empty-icon">!</p>
        <h2 class="brand">Daten konnten nicht geladen werden</h2>
        <p class="helper-text">${esc(state.loadError)}</p>
      </div>
    `));
    return;
  }
  if (state.view === "home") renderHome(app);
  else if (state.view === "learn") renderLearn(app);
  else if (state.view === "done") renderDone(app);
  else if (state.view === "settings") renderSettings(app);
  else if (state.view === "errors") renderErrors(app);
  window.scrollTo(0, 0);
}

// ---------- Home ----------
function renderHome(root) {
  const s = getStats();
  const todayLearned = state.meta.todayKey === todayKey();
  const needsBackup = needsBackupReminder();

  root.appendChild(el(`
    <div class="header">
      <div class="header-title">
        <h1 class="brand">srpski</h1>
        <span class="brand-sub">B2</span>
      </div>
      <button class="icon-btn" id="settings-btn" aria-label="Einstellungen">${ICONS.gear}</button>
    </div>
  `));
  $("settings-btn").onclick = () => goto("settings");

  // Hero: due heute
  root.appendChild(el(`
    <div class="hero">
      <p class="hero-num ${s.due === 0 ? "muted" : ""}">${s.due}</p>
      <p class="hero-label">${s.due === 1 ? "Karte heute fällig" : "Karten heute fällig"}</p>
      <div class="today-indicator ${todayLearned ? "active" : ""}">
        <span class="dot"></span>
        <span>${todayLearned ? "heute schon gelernt" : "heute noch nicht gelernt"}</span>
      </div>
    </div>
  `));

  // Stat-Triplet
  root.appendChild(el(`
    <div class="stat-grid">
      <div class="stat-mini">
        <div class="stat-mini-num">${s.unseen}</div>
        <div class="stat-mini-label">Neu</div>
      </div>
      <div class="stat-mini">
        <div class="stat-mini-num">${s.learning}</div>
        <div class="stat-mini-label">Lernend</div>
      </div>
      <div class="stat-mini">
        <div class="stat-mini-num">${s.mature}</div>
        <div class="stat-mini-label">Gemeistert</div>
      </div>
    </div>
  `));

  // Primary action stack
  const actions = el(`<div class="action-stack"></div>`);
  const startBtn = el(`<button class="btn btn-primary" ${s.due === 0 ? "disabled" : ""}>${s.due === 0 ? "Heute nichts fällig" : "Lernen starten"}</button>`);
  startBtn.onclick = () => startSession();
  actions.appendChild(startBtn);

  // Backup button - prominent if needed, ghost otherwise
  const backupBtn = el(`
    <button class="btn ${needsBackup ? "" : "btn-ghost"}" id="backup-btn">
      ${ICONS.download}
      <span>${needsBackup ? "Backup empfohlen" : "Backup speichern"}</span>
    </button>
  `);
  backupBtn.onclick = () => exportProgress();
  actions.appendChild(backupBtn);

  root.appendChild(actions);

  if (needsBackup) {
    root.appendChild(el(`
      <p class="helper-text center mt-1">
        Seit dem letzten Backup ${state.meta.learnedSinceBackup} neue Karten oder mehr als ${BACKUP_REMIND_DAYS} Tage vergangen.
      </p>
    `));
  }
}

function needsBackupReminder() {
  if (!state.meta.lastBackup) {
    // never backed up, but only remind if any progress
    return Object.keys(state.progress).length > 5;
  }
  const daysAgo = (Date.now() - state.meta.lastBackup) / DAY_MS;
  return daysAgo >= BACKUP_REMIND_DAYS || state.meta.learnedSinceBackup >= BACKUP_REMIND_NEW;
}

// ---------- Session start ----------
function startSession() {
  const now = Date.now();
  // Erst Stau nach Pause umverteilen
  const allCards = Object.values(state.progress);
  fsrsRedistribute(allCards, now);
  saveProgress();

  // Queue zusammenbauen: erst überfällige Reviews, dann frische neue Karten (cap dailyNew)
  const candidates = state.words.filter(passesFilter);
  const dueReviews = [];
  const newWords = [];
  for (const w of candidates) {
    const c = state.progress[w.id];
    if (!c || c.state === "new") {
      newWords.push(w);
    } else if (c.due <= now) {
      dueReviews.push(w);
    }
  }
  // New words: nicht nur C zuerst, sondern A/B/C gemischt zufällig
  shuffle(newWords);
  const newPicked = newWords.slice(0, state.settings.dailyNew);

  // Reviews: in zufälliger Reihenfolge
  shuffle(dueReviews);

  // Interleave: new + reviews abwechselnd, damit Robin nicht 50 reviews am Stück sieht
  const queue = interleave(dueReviews, newPicked);

  if (queue.length === 0) return;
  state.session = {
    queue: queue.map(w => w.id),
    idx: 0,
    revealed: false,
    results: [],
    origCount: queue.length,
    againSet: new Set(),
  };
  goto("learn");
}

function shuffle(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
}

function interleave(a, b) {
  const out = [];
  const ia = a.length, ib = b.length;
  if (ia === 0) return b;
  if (ib === 0) return a;
  const ratio = ia / (ia + ib);
  let pa = 0, pb = 0;
  for (let k = 0; k < ia + ib; k++) {
    if (pa >= ia) { out.push(b[pb++]); continue; }
    if (pb >= ib) { out.push(a[pa++]); continue; }
    const expectedA = (k + 1) * ratio;
    if (pa < expectedA) out.push(a[pa++]);
    else out.push(b[pb++]);
  }
  return out;
}

// ---------- Learn ----------
function renderLearn(root) {
  const sess = state.session;
  if (!sess) { goto("home"); return; }
  const wordId = sess.queue[sess.idx];
  const w = state.words.find(x => x.id === wordId);
  if (!w) { sess.idx += 1; renderLearn(root); return; }

  const done = sess.idx;
  const total = sess.origCount;
  const pct = total > 0 ? Math.min(100, (done / total) * 100) : 0;

  root.appendChild(el(`
    <div class="learn-bar">
      <button class="icon-btn" id="back-btn" aria-label="Abbrechen">${ICONS.back}</button>
      <div class="progress"><div class="progress-fill" style="width:${pct}%"></div></div>
      <span class="progress-count">${done + 1} / ${total}</span>
    </div>
  `));
  $("back-btn").onclick = () => abortSession();

  const dir = state.settings.direction;
  const filterClass = `filter-${w.f}`;
  const badgesHtml = `
    <div class="card-meta">
      <span class="badge ${filterClass}">${w.f.toUpperCase()}</span>
      ${w.wa ? `<span class="badge">${esc(w.wa)}</span>` : ""}
      <span class="badge">${esc(w.tb)}</span>
    </div>
  `;

  const cardEl = el(`<div class="card ${!sess.revealed ? "card-tap" : ""}"></div>`);
  cardEl.appendChild(el(badgesHtml));

  const body = el(`<div class="card-body"></div>`);
  const frontPrime = dir === "de2sr" ? w.de : w.sl;
  const frontSub = dir === "de2sr" ? "" : w.sc;
  const frontIsDe = dir === "de2sr";
  body.appendChild(el(`
    <p class="word-prime ${frontIsDe ? "" : "sans"}">${esc(frontPrime)}</p>
    ${frontSub ? `<p class="word-cir">${esc(frontSub)}</p>` : ""}
    ${frontIsDe && w.fm ? `<p class="form-hint">${esc(w.fm)}</p>` : ""}
  `));

  if (!sess.revealed) {
    body.appendChild(el(`<p class="tap-hint">Tipp auf die Karte</p>`));
  } else {
    body.appendChild(renderBack(w, dir));
  }
  cardEl.appendChild(body);

  if (!sess.revealed) {
    cardEl.onclick = () => { sess.revealed = true; render(); };
  }

  root.appendChild(cardEl);

  // Extras: skip + note + flag, immer sichtbar
  const note = state.notes[w.id];
  const extras = el(`
    <div class="card-extras">
      <button class="btn-extra" id="skip-btn" title="Wort kenne ich schon, 60 Tage pausieren">
        ${ICONS.skip}<span>Kenne ich schon</span>
      </button>
      <button class="btn-extra" id="note-btn" title="${note ? "Eselsbrücke bearbeiten" : "Eselsbrücke hinzufügen"}">
        ${note ? ICONS.edit : ICONS.bulb}<span>${note ? "Notiz" : "Notiz"}</span>
      </button>
      <button class="btn-extra" id="flag-btn" title="Fehler in diesem Eintrag melden">
        ${ICONS.flag}<span>Fehler</span>
      </button>
    </div>
  `);
  extras.querySelector("#skip-btn").onclick = () => handleSkip(w);
  extras.querySelector("#note-btn").onclick = () => openNoteModal(w);
  extras.querySelector("#flag-btn").onclick = () => openErrorModal(w);
  root.appendChild(extras);

  if (sess.revealed) {
    const rate = el(`
      <div class="btn-row rate-row">
        <button class="btn btn-again" data-r="1">
          <span>Nochmal</span><span class="btn-sub">again</span>
        </button>
        <button class="btn btn-hard" data-r="2">
          <span>Schwer</span><span class="btn-sub">hard</span>
        </button>
        <button class="btn btn-good" data-r="3">
          <span>Gut</span><span class="btn-sub">good</span>
        </button>
        <button class="btn btn-easy" data-r="4">
          <span>Easy</span><span class="btn-sub">easy</span>
        </button>
      </div>
    `);
    rate.querySelectorAll("button").forEach(btn => {
      btn.onclick = (e) => {
        e.stopPropagation();
        const r = parseInt(btn.dataset.r, 10);
        handleRate(w, r);
      };
    });
    root.appendChild(rate);
  }
}

function renderBack(w, dir) {
  const wrap = document.createDocumentFragment();
  wrap.appendChild(el(`<hr class="divider">`));

  if (dir === "de2sr") {
    // Show SR
    const pairParts = (w.sl || "").split(" / ");
    if (pairParts.length >= 2 && w.wa && /verb/i.test(w.wa)) {
      const cirParts = (w.sc || "").split(" / ");
      const formInfo = state.forms[wordKey(w)] || {};
      const presIpf = formInfo.pres1_ipf || "";
      const presPf = formInfo.pres1_pf || "";
      wrap.appendChild(el(`
        <div class="aspect-grid">
          <div class="aspect-box">
            <div class="aspect-label">Imperfektiv</div>
            <div class="aspect-form">${esc(pairParts[0].trim())}</div>
            ${cirParts[0] ? `<div class="aspect-cir">${esc(cirParts[0].trim())}</div>` : ""}
            ${presIpf ? `<div class="aspect-pres">1. P. Sg. ${esc(presIpf)}</div>` : ""}
          </div>
          <div class="aspect-box pf">
            <div class="aspect-label">Perfektiv</div>
            <div class="aspect-form">${esc(pairParts[1].trim())}</div>
            ${cirParts[1] ? `<div class="aspect-cir">${esc(cirParts[1].trim())}</div>` : ""}
            ${presPf ? `<div class="aspect-pres">1. P. Sg. ${esc(presPf)}</div>` : ""}
          </div>
        </div>
      `));
    } else {
      wrap.appendChild(el(`<p class="word-prime sans" style="font-size:30px;margin-top:0.5rem">${esc(w.sl)}</p>`));
      if (w.sc) wrap.appendChild(el(`<p class="word-cir">${esc(w.sc)}</p>`));
    }
    // Form-Anreicherung aus forms.json (sekundär)
    const formHint = lookupForm(w);
    if (formHint) wrap.appendChild(el(`<p class="form-hint">${esc(formHint)}</p>`));
  } else {
    // sr2de: show DE
    wrap.appendChild(el(`<p class="word-prime" style="font-size:30px;margin-top:0.5rem">${esc(w.de)}</p>`));
    if (w.fm) wrap.appendChild(el(`<p class="form-hint">${esc(w.fm)}</p>`));
  }

  // Examples
  const ex = state.examples[wordKey(w)];
  if (ex) {
    const exEl = el(`<div class="examples"></div>`);
    if (ex.ipf) {
      exEl.appendChild(el(`
        <div class="example-row">
          <span class="example-tag">IPF</span>
          <div class="example-content">
            <div class="example-sr">${esc(ex.ipf[0])}</div>
            <div class="example-de">${esc(ex.ipf[1])}</div>
          </div>
        </div>
      `));
    }
    if (ex.pf) {
      exEl.appendChild(el(`
        <div class="example-row">
          <span class="example-tag pf">PF</span>
          <div class="example-content">
            <div class="example-sr">${esc(ex.pf[0])}</div>
            <div class="example-de">${esc(ex.pf[1])}</div>
          </div>
        </div>
      `));
    }
    wrap.appendChild(exEl);
  }

  // Notiz/Eselsbrücke (persönlich)
  const note = state.notes[w.id];
  if (note && note.trim()) {
    wrap.appendChild(el(`
      <div class="note-block">
        <div class="note-label">${ICONS.bulb}<span>Eselsbrücke</span></div>
        <div class="note-text">${esc(note)}</div>
      </div>
    `));
  }

  return wrap;
}

function lookupForm(w) {
  const f = state.forms[wordKey(w)];
  if (!f) return "";
  // Format depends on Wortart
  if (f.mfn) return `m: ${f.mfn[0]} · f: ${f.mfn[1]} · n: ${f.mfn[2]}`;
  if (f.genus) {
    let s = `${f.genus}`;
    if (f.plural) s += ` · Pl. ${f.plural}`;
    if (f.gen) s += ` · Gen. ${f.gen}`;
    return s;
  }
  // Singleton-Verben (kein Aspektpaar): pres1 als form-hint
  // Aspektpaare bekommen die Anzeige in der aspect-grid, nicht hier
  if (f.pres1_ipf && !f.pres1_pf) return `1. P. Sg. ${f.pres1_ipf}`;
  return "";
}

// ---------- Rating handler ----------
function handleRate(w, rating) {
  const sess = state.session;
  if (!sess) return;
  const now = Date.now();
  const card = getCard(w.id);
  const updated = fsrsReview(card, rating, now);
  state.progress[w.id] = updated;

  sess.results.push({ id: w.id, rating });

  // Wenn Again: ans Ende der Queue setzen für Re-Show in derselben Session
  if (rating === 1) {
    sess.queue.push(w.id);
    sess.againSet.add(w.id);
  }

  // Backup-Counter
  state.meta.learnedSinceBackup = (state.meta.learnedSinceBackup || 0) + 1;
  state.meta.todayKey = todayKey(now);

  saveProgress();
  saveMeta();

  sess.idx += 1;
  sess.revealed = false;

  if (sess.idx >= sess.queue.length) {
    goto("done");
  } else {
    render();
  }
}

function handleSkip(w) {
  const sess = state.session;
  if (!sess) return;
  const now = Date.now();
  const card = getCard(w.id);
  state.progress[w.id] = fsrsSkip(card, now);
  saveProgress();

  state.meta.todayKey = todayKey(now);
  saveMeta();

  sess.results.push({ id: w.id, rating: 0 });
  sess.idx += 1;
  sess.revealed = false;

  if (sess.idx >= sess.queue.length) goto("done");
  else render();
}

function abortSession() {
  if (state.session && state.session.idx > 0) {
    if (!confirm("Session wirklich abbrechen? Bisheriger Fortschritt der Session ist gespeichert, aber die Reihenfolge geht verloren.")) return;
  }
  state.session = null;
  goto("home");
}

// ---------- Done ----------
function renderDone(root) {
  const r = state.session ? state.session.results : [];
  const skipped = r.filter(x => x.rating === 0).length;
  const again = r.filter(x => x.rating === 1).length;
  const hard = r.filter(x => x.rating === 2).length;
  const good = r.filter(x => x.rating === 3).length;
  const easy = r.filter(x => x.rating === 4).length;

  root.appendChild(el(`<div class="header"><div class="header-title"><h1 class="brand">srpski</h1><span class="brand-sub">geschafft</span></div></div>`));

  root.appendChild(el(`
    <div class="center-prose">
      <p class="hero-num">${r.length}</p>
      <p class="hero-label">Karten in dieser Session</p>
    </div>
  `));

  root.appendChild(el(`
    <div class="stat-grid">
      <div class="stat-mini"><div class="stat-mini-num">${easy}</div><div class="stat-mini-label">Easy</div></div>
      <div class="stat-mini"><div class="stat-mini-num">${good}</div><div class="stat-mini-label">Gut</div></div>
      <div class="stat-mini"><div class="stat-mini-num">${hard}</div><div class="stat-mini-label">Schwer</div></div>
    </div>
    <div class="stat-grid mt-1">
      <div class="stat-mini"><div class="stat-mini-num">${again}</div><div class="stat-mini-label">Nochmal</div></div>
      <div class="stat-mini"><div class="stat-mini-num">${skipped}</div><div class="stat-mini-label">Übersprungen</div></div>
      <div class="stat-mini"><div class="stat-mini-num">${state.session?.againSet?.size || 0}</div><div class="stat-mini-label">Wiederholt</div></div>
    </div>
  `));

  const backBtn = el(`<button class="btn btn-primary mt-3">Zur Startseite</button>`);
  backBtn.onclick = () => {
    state.session = null;
    goto("home");
  };
  root.appendChild(backBtn);
}

// ---------- Settings ----------
function renderSettings(root) {
  root.appendChild(el(`
    <div class="header">
      <button class="icon-btn" id="back-btn" aria-label="Zurück">${ICONS.back}</button>
      <div class="header-title"><h1 class="brand">Einstellungen</h1></div>
      <div style="width:40px"></div>
    </div>
  `));
  $("back-btn").onclick = () => goto("home");

  // Direction
  root.appendChild(el(`<div class="section-label">Lernrichtung</div>`));
  const dirBox = el(`<div class="chips"></div>`);
  for (const [val, lbl] of [["de2sr", "DE → SR"], ["sr2de", "SR → DE"]]) {
    const c = el(`<div class="chip ${state.settings.direction === val ? "active" : ""}">${lbl}</div>`);
    c.onclick = () => { state.settings.direction = val; saveSettings(); render(); };
    dirBox.appendChild(c);
  }
  root.appendChild(dirBox);

  // Filter
  root.appendChild(el(`<div class="section-label">Filter</div>`));
  const filterBox = el(`<div class="chips"></div>`);
  for (const [val, lbl] of [["a", "A — aktiv"], ["b", "B — passiv"], ["c", "C — neu"]]) {
    const isActive = state.settings.filters.includes(val);
    const c = el(`<div class="chip ${isActive ? "active" : ""}">${lbl}</div>`);
    c.onclick = () => {
      if (isActive && state.settings.filters.length > 1) {
        state.settings.filters = state.settings.filters.filter(x => x !== val);
      } else if (!isActive) {
        state.settings.filters = [...state.settings.filters, val];
      }
      saveSettings(); render();
    };
    filterBox.appendChild(c);
  }
  root.appendChild(filterBox);

  // Themen
  root.appendChild(el(`<div class="section-label">Themen</div>`));
  const themBox = el(`<div class="chips"></div>`);
  const allChip = el(`<div class="chip ${state.settings.themen === null ? "active" : ""}">Alle</div>`);
  allChip.onclick = () => { state.settings.themen = null; saveSettings(); render(); };
  themBox.appendChild(allChip);
  for (const t of state.themen) {
    const isActive = state.settings.themen && state.settings.themen.includes(t);
    const c = el(`<div class="chip chip-sm ${isActive ? "active" : ""}">${esc(t)}</div>`);
    c.onclick = () => {
      if (state.settings.themen === null) {
        state.settings.themen = state.themen.filter(x => x !== t);
      } else if (isActive) {
        state.settings.themen = state.settings.themen.filter(x => x !== t);
      } else {
        state.settings.themen = [...state.settings.themen, t];
      }
      if (state.settings.themen.length === state.themen.length) state.settings.themen = null;
      if (state.settings.themen && state.settings.themen.length === 0) state.settings.themen = [t];
      saveSettings(); render();
    };
    themBox.appendChild(c);
  }
  root.appendChild(themBox);

  // Neue Wörter pro Session
  root.appendChild(el(`<div class="section-label">Neue Wörter pro Session</div>`));
  const slider = el(`
    <div class="row">
      <span>Max <strong id="dn-val">${state.settings.dailyNew}</strong></span>
      <input type="range" min="5" max="50" step="5" value="${state.settings.dailyNew}" id="dn-slider">
    </div>
  `);
  root.appendChild(slider);
  $("dn-slider").oninput = (e) => {
    state.settings.dailyNew = parseInt(e.target.value, 10);
    $("dn-val").textContent = state.settings.dailyNew;
    saveSettings();
  };

  // Daten & Backup
  root.appendChild(el(`<div class="section-label">Daten</div>`));
  const seenCount = Object.values(state.progress).filter(c => c.reps > 0).length;
  root.appendChild(el(`<div class="row"><span>Gelernte Karten</span><span class="meta">${seenCount}</span></div>`));
  root.appendChild(el(`<div class="row"><span>Eigene Eselsbrücken</span><span class="meta">${Object.keys(state.notes).length}</span></div>`));

  const lastBackup = state.meta.lastBackup
    ? new Date(state.meta.lastBackup).toLocaleString("de-DE", { dateStyle: "short", timeStyle: "short" })
    : "nie";
  root.appendChild(el(`<div class="row"><span>Letztes Backup</span><span class="meta">${lastBackup}</span></div>`));

  const expBtn = el(`<button class="btn btn-ghost mt-1">${ICONS.download}<span>Fortschritt exportieren</span></button>`);
  expBtn.onclick = () => exportProgress();
  root.appendChild(expBtn);

  const impBtn = el(`<button class="btn btn-ghost">Fortschritt importieren</button>`);
  impBtn.onclick = () => importProgress();
  root.appendChild(impBtn);

  // Fehler-Liste
  root.appendChild(el(`<div class="section-label">Fehlermeldungen</div>`));
  root.appendChild(el(`<div class="row"><span>Gespeicherte Fehler</span><span class="meta">${state.errors.length}</span></div>`));
  if (state.errors.length > 0) {
    const errBtn = el(`<button class="btn btn-ghost">Fehler-Liste anzeigen</button>`);
    errBtn.onclick = () => goto("errors");
    root.appendChild(errBtn);
  }

  // Reset
  root.appendChild(el(`<div class="section-label">Gefahrenzone</div>`));
  const resetBtn = el(`<button class="btn btn-danger">Allen Fortschritt zurücksetzen</button>`);
  resetBtn.onclick = () => {
    if (!confirm("Bist du sicher? Aller Lernfortschritt geht verloren.")) return;
    if (!confirm("Wirklich? Das kann nicht rückgängig gemacht werden.")) return;
    state.progress = {};
    state.meta = { lastBackup: 0, learnedSinceBackup: 0, todayKey: "" };
    saveProgress();
    saveMeta();
    render();
  };
  root.appendChild(resetBtn);
}

// ---------- Errors view ----------
function renderErrors(root) {
  root.appendChild(el(`
    <div class="header">
      <button class="icon-btn" id="back-btn" aria-label="Zurück">${ICONS.back}</button>
      <div class="header-title"><h1 class="brand">Fehlermeldungen</h1></div>
      <div style="width:40px"></div>
    </div>
  `));
  $("back-btn").onclick = () => goto("settings");

  if (state.errors.length === 0) {
    root.appendChild(el(`<div class="center-prose"><p class="empty-icon">·</p><p class="helper-text">Keine Fehler gemeldet.</p></div>`));
    return;
  }

  root.appendChild(el(`<p class="helper-text">${state.errors.length} ${state.errors.length === 1 ? "Eintrag" : "Einträge"}. Beim Export wird die Datei zum Sharing erzeugt.</p>`));

  for (const err of state.errors.slice().reverse()) {
    const w = state.words.find(x => x.id === err.id);
    const wText = w ? `${w.de} → ${w.sl}` : `(ID ${err.id})`;
    const date = new Date(err.ts).toLocaleString("de-DE", { dateStyle: "short", timeStyle: "short" });
    const row = el(`
      <div class="row" style="flex-direction:column;align-items:stretch;gap:6px">
        <div class="row-spread">
          <strong>${esc(wText)}</strong>
          <span class="meta">${esc(date)}</span>
        </div>
        ${err.comment ? `<div class="helper-text">${esc(err.comment)}</div>` : ""}
      </div>
    `);
    root.appendChild(row);
  }

  const expBtn = el(`<button class="btn btn-ghost mt-2">${ICONS.download}<span>Fehler-Liste exportieren</span></button>`);
  expBtn.onclick = () => exportErrors();
  root.appendChild(expBtn);

  const clearBtn = el(`<button class="btn btn-danger mt-1">Fehler-Liste leeren</button>`);
  clearBtn.onclick = () => {
    if (!confirm("Alle Fehlermeldungen löschen?")) return;
    state.errors = [];
    saveErrors();
    render();
  };
  root.appendChild(clearBtn);
}

// ---------- Error modal ----------
function openErrorModal(w) {
  const existing = $("error-modal");
  if (existing) existing.remove();

  const backdrop = el(`
    <div class="modal-backdrop" id="error-modal">
      <div class="modal" onclick="event.stopPropagation()">
        <div class="row-spread">
          <h2>Fehler melden</h2>
          <button class="icon-btn" id="modal-close">${ICONS.close}</button>
        </div>
        <p class="helper-text" style="margin-bottom:0.75rem">
          <strong>${esc(w.de)} → ${esc(w.sl)}</strong>
        </p>
        <textarea id="error-comment" placeholder="Was stimmt nicht? (Tippfehler, falsche Übersetzung, ...)"></textarea>
        <div class="action-stack mt-2">
          <button class="btn btn-primary" id="modal-save">Speichern</button>
          <button class="btn btn-ghost" id="modal-cancel">Abbrechen</button>
        </div>
      </div>
    </div>
  `);
  backdrop.onclick = () => backdrop.remove();
  document.body.appendChild(backdrop);

  $("modal-close").onclick = () => backdrop.remove();
  $("modal-cancel").onclick = () => backdrop.remove();
  $("modal-save").onclick = () => {
    const comment = $("error-comment").value.trim();
    state.errors.push({
      id: w.id,
      de: w.de,
      sl: w.sl,
      comment,
      ts: Date.now(),
    });
    saveErrors();
    backdrop.remove();
  };
  // Auto-focus textarea
  setTimeout(() => $("error-comment").focus(), 50);
}

// ---------- Note modal ----------
function openNoteModal(w) {
  const existing = $("note-modal");
  if (existing) existing.remove();

  const currentNote = state.notes[w.id] || "";

  const backdrop = el(`
    <div class="modal-backdrop" id="note-modal">
      <div class="modal" onclick="event.stopPropagation()">
        <div class="row-spread">
          <h2>${currentNote ? "Eselsbrücke bearbeiten" : "Eselsbrücke hinzufügen"}</h2>
          <button class="icon-btn" id="note-close">${ICONS.close}</button>
        </div>
        <p class="helper-text" style="margin-bottom:0.75rem">
          <strong>${esc(w.de)} → ${esc(w.sl)}</strong>
        </p>
        <textarea id="note-text" placeholder="Notiere, was dir hilft, dieses Wort zu merken — Klang, Bild, Verbindung, Wortherkunft, persönliche Assoziation."></textarea>
        <div class="action-stack mt-2">
          <button class="btn btn-primary" id="note-save">Speichern</button>
          ${currentNote ? `<button class="btn btn-danger" id="note-delete">Löschen</button>` : ""}
          <button class="btn btn-ghost" id="note-cancel">Abbrechen</button>
        </div>
      </div>
    </div>
  `);
  backdrop.onclick = () => backdrop.remove();
  document.body.appendChild(backdrop);

  $("note-text").value = currentNote;
  $("note-close").onclick = () => backdrop.remove();
  $("note-cancel").onclick = () => backdrop.remove();
  $("note-save").onclick = () => {
    const text = $("note-text").value.trim();
    if (text) state.notes[w.id] = text;
    else delete state.notes[w.id];
    saveNotes();
    backdrop.remove();
    render();
  };
  const delBtn = $("note-delete");
  if (delBtn) delBtn.onclick = () => {
    delete state.notes[w.id];
    saveNotes();
    backdrop.remove();
    render();
  };
  setTimeout(() => $("note-text").focus(), 50);
}
function exportProgress() {
  const payload = {
    type: "serbisch-b2-backup",
    version: 1,
    timestamp: new Date().toISOString(),
    progress: state.progress,
    meta: state.meta,
    settings: state.settings,
    errors: state.errors,
    notes: state.notes,
  };
  downloadJson(payload, `serbisch-b2-backup-${todayKey()}.json`);
  state.meta.lastBackup = Date.now();
  state.meta.learnedSinceBackup = 0;
  saveMeta();
  render();
}

function exportErrors() {
  const payload = {
    type: "serbisch-b2-errors",
    timestamp: new Date().toISOString(),
    errors: state.errors,
  };
  downloadJson(payload, `serbisch-b2-fehler-${todayKey()}.json`);
}

function downloadJson(obj, filename) {
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 500);
}

function importProgress() {
  const inp = document.createElement("input");
  inp.type = "file";
  inp.accept = ".json,application/json";
  inp.onchange = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result);
        if (data.type !== "serbisch-b2-backup" || !data.progress) {
          alert("Datei sieht nicht nach einem Backup aus.");
          return;
        }
        if (!confirm("Aktuellen Fortschritt mit Backup ersetzen?")) return;
        state.progress = data.progress;
        if (data.meta) state.meta = { ...state.meta, ...data.meta };
        if (data.settings) Object.assign(state.settings, data.settings);
        if (data.errors) state.errors = data.errors;
        if (data.notes) state.notes = data.notes;
        saveProgress(); saveMeta(); saveSettings(); saveErrors(); saveNotes();
        render();
      } catch (err) {
        alert("Datei nicht lesbar: " + err.message);
      }
    };
    reader.readAsText(file);
  };
  inp.click();
}

// ---------- Navigation ----------
function goto(view) {
  state.view = view;
  render();
}

// ============ SW REGISTRATION ============
async function registerSW() {
  if (!("serviceWorker" in navigator)) return;
  try {
    await navigator.serviceWorker.register("./sw.js");
  } catch (e) {
    console.warn("SW registration failed:", e);
  }
}

// ============ INIT ============
async function init() {
  loadStored();
  await loadData();
  render();
  registerSW();
}

init();
