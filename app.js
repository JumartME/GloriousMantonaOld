const CACHE_KEY = "npcviewer:npcs:bootstrap:v1";

const els = {
  file: document.getElementById("file"),
  clear: document.getElementById("clear"),
  status: document.getElementById("status"),
  count: document.getElementById("count"),

  q: document.getElementById("q"),
  origin: document.getElementById("origin"),
  reputation: document.getElementById("reputation"),
  relation: document.getElementById("relation"),
  sort: document.getElementById("sort"),
  onlyRecruited: document.getElementById("onlyRecruited"),
  hideDefeated: document.getElementById("hideDefeated"),

  list: document.getElementById("list"),

  // modal
  modalEl: document.getElementById("npcModal"),
  modalTitle: document.getElementById("modalTitle"),
  modalSubtitle: document.getElementById("modalSubtitle"),
  modalImg: document.getElementById("modalImg"),
  kvIdentity: document.getElementById("kvIdentity"),
  kvSocial: document.getElementById("kvSocial"),
  kvStats: document.getElementById("kvStats"),
  kvGear: document.getElementById("kvGear"),
  skillsBox: document.getElementById("skillsBox"),
  descBox: document.getElementById("descBox"),
  modalFooter: document.getElementById("modalFooter"),
};

let dataset = [];
let view = [];
let modal = null;

function clean(v) { return String(v ?? "").trim(); }
function setStatus(msg) { els.status.textContent = msg || ""; }

function titleCase(s) {
  const t = clean(s);
  if (!t) return "";
  return t.charAt(0).toUpperCase() + t.slice(1).toLowerCase();
}

function normEnum(v, allowedLower, fallback) {
  const s = clean(v).toLowerCase();
  if (!s) return fallback;
  return allowedLower.includes(s) ? titleCase(s) : fallback;
}

function uniqueSorted(arr) {
  return [...new Set(arr.filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

function setOptions(select, values, allLabel) {
  select.innerHTML = "";
  const all = document.createElement("option");
  all.value = "";
  all.textContent = allLabel;
  select.appendChild(all);
  for (const v of values) {
    const o = document.createElement("option");
    o.value = v;
    o.textContent = v;
    select.appendChild(o);
  }
}

function loadCache() {
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch {
    return null;
  }
}
function saveCache(obj) {
  localStorage.setItem(CACHE_KEY, JSON.stringify(obj));
}
function clearCache() {
  localStorage.removeItem(CACHE_KEY);
}

/** Skip separator rows: Name filled but all other columns empty */
function isSeparatorRow(row) {
  const name = clean(row.Name);
  if (!name) return false;

  const fieldsToCheck = [
    "Gender","Age","Species","Concept","Description","Origin",
    "Equipment","Armor","Shield","Weapon","Magic","Special",
    "Healing","Agility","Strength","Dexterity","Stamina","Intelligence",
    "Perception","Will","Wits","Expression","Instinct","Presence","Wisdom",
    "Skills","Size","Health","Spirit","Reputation","Relation"
  ];

  return fieldsToCheck.every((k) => clean(row[k]) === "");
}

function rowToNpc(row) {
  const Name = clean(row.Name);
  if (!Name) return null;
  if (isSeparatorRow(row)) return null;

  const Origin = clean(row.Origin);

  const Reputation = normEnum(row.Reputation, ["hostile","friendly","neutral"], "Neutral");
  const Relation = normEnum(row.Relation, ["unknown","met","recruited","defeated"], "Unknown");

  const imageUrl = Origin
    ? `./images/${Origin}/${Name}.jpg`
    : `./images/${Name}.jpg`;

  return {
    id: Name.toLowerCase().replace(/\s+/g, "-"),
    Name, Origin, Reputation, Relation, imageUrl,

    Gender: clean(row.Gender),
    Age: clean(row.Age),
    Species: clean(row.Species),
    Concept: clean(row.Concept),
    Description: clean(row.Description),

    Equipment: clean(row.Equipment),
    Armor: clean(row.Armor),
    Shield: clean(row.Shield),
    Weapon: clean(row.Weapon),
    Magic: clean(row.Magic),
    Special: clean(row.Special),

    Healing: clean(row.Healing),
    Agility: clean(row.Agility),
    Strength: clean(row.Strength),
    Dexterity: clean(row.Dexterity),
    Stamina: clean(row.Stamina),
    Intelligence: clean(row.Intelligence),
    Perception: clean(row.Perception),
    Will: clean(row.Will),
    Wits: clean(row.Wits),
    Expression: clean(row.Expression),
    Instinct: clean(row.Instinct),
    Presence: clean(row.Presence),
    Wisdom: clean(row.Wisdom),

    Skills: clean(row.Skills),
    Size: clean(row.Size),
    Health: clean(row.Health),
    Spirit: clean(row.Spirit),
    MagicStat: clean(row["Magic "] ?? ""),
  };
}

async function parseExcelFile(file) {
  const { read, utils } = await import("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm");
  const buf = await file.arrayBuffer();
  const wb = read(buf);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = utils.sheet_to_json(ws, { defval: "" });
  const npcs = rows.map(rowToNpc).filter(Boolean);
  return { updatedAt: new Date().toISOString(), count: npcs.length, npcs };
}

function initFilters() {
  setOptions(els.origin, uniqueSorted(dataset.map(n => n.Origin)), "All Origins");
  setOptions(els.reputation, ["Hostile","Friendly","Neutral"], "All Reputation");
  setOptions(els.relation, ["Unknown","Met","Recruited","Defeated"], "All Relation");

  els.sort.innerHTML = "";
  const sorts = [
    ["name_asc", "Name (A–Z)"],
    ["name_desc", "Name (Z–A)"],
    ["origin_asc", "Origin (A–Z)"],
    ["relation", "Relation"],
    ["reputation", "Reputation"],
  ];
  for (const [v, label] of sorts) {
    const o = document.createElement("option");
    o.value = v;
    o.textContent = label;
    els.sort.appendChild(o);
  }
  els.sort.value = "name_asc";
}

function matches(n) {
  const q = els.q.value.trim().toLowerCase();
  const origin = els.origin.value;
  const rep = els.reputation.value;
  const rel = els.relation.value;

  if (q && !n.Name.toLowerCase().includes(q)) return false;
  if (origin && n.Origin !== origin) return false;
  if (rep && n.Reputation !== rep) return false;
  if (rel && n.Relation !== rel) return false;
  if (els.onlyRecruited.checked && n.Relation !== "Recruited") return false;
  if (els.hideDefeated.checked && n.Relation === "Defeated") return false;

  return true;
}

const REL_ORDER = ["Recruited", "Met", "Unknown", "Defeated"];
const REP_ORDER = ["Friendly", "Neutral", "Hostile"];

function sortKeyRelation(n) {
  const i = REL_ORDER.indexOf(n.Relation);
  return i === -1 ? 999 : i;
}
function sortKeyReputation(n) {
  const i = REP_ORDER.indexOf(n.Reputation);
  return i === -1 ? 999 : i;
}

function sortNpcs(arr) {
  const mode = els.sort.value;
  const copy = [...arr];
  copy.sort((a, b) => {
    if (mode === "name_asc") return a.Name.localeCompare(b.Name);
    if (mode === "name_desc") return b.Name.localeCompare(a.Name);
    if (mode === "origin_asc") {
      const c = (a.Origin || "").localeCompare(b.Origin || "");
      return c !== 0 ? c : a.Name.localeCompare(b.Name);
    }
    if (mode === "relation") {
      const c = sortKeyRelation(a) - sortKeyRelation(b);
      return c !== 0 ? c : a.Name.localeCompare(b.Name);
    }
    if (mode === "reputation") {
      const c = sortKeyReputation(a) - sortKeyReputation(b);
      return c !== 0 ? c : a.Name.localeCompare(b.Name);
    }
    return a.Name.localeCompare(b.Name);
  });
  return copy;
}

function nameClass(n) {
  // Hostile overrides
  if (n.Reputation === "Hostile") return "text-danger";

  // relation coloring
  if (n.Relation === "Recruited") return "text-success";
  if (n.Relation === "Met") return "text-primary";
  if (n.Relation === "Defeated") return "text-body-secondary";
  return "text-muted";
}

function pillHtml(text, cls) {
  if (!clean(text)) return "";
  return `<span class="badge rounded-pill ${cls}">${escapeHtml(text)}</span>`;
}

function escapeHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function buildRow(n, idx) {
  const a = document.createElement("a");
  a.className = "list-group-item list-group-item-action d-flex gap-3 align-items-center npc-row";
  a.href = "javascript:void(0)";
  a.dataset.index = String(idx);

  // thumb
  const img = document.createElement("img");
  img.className = "thumb";
  img.alt = n.Name;
  img.src = n.imageUrl;
  img.onerror = () => {
    img.removeAttribute("src");
    img.classList.add("missing");
    img.alt = `${n.Name} (missing image)`;
  };

  // text
  const wrap = document.createElement("div");
  wrap.className = "flex-grow-1";

  const top = document.createElement("div");
  top.className = "d-flex align-items-center justify-content-between gap-2";

  const name = document.createElement("div");
  name.className = `npc-name ${nameClass(n)}`;
  name.textContent = n.Name;

  const badges = document.createElement("div");
  badges.className = "d-flex gap-2 flex-wrap justify-content-end";
  badges.innerHTML =
    pillHtml(n.Origin, "text-bg-light border") +
    pillHtml(n.Reputation, n.Reputation === "Hostile" ? "text-bg-danger" : (n.Reputation === "Friendly" ? "text-bg-success" : "text-bg-secondary")) +
    pillHtml(n.Relation, n.Relation === "Recruited" ? "text-bg-success" : (n.Relation === "Met" ? "text-bg-primary" : (n.Relation === "Defeated" ? "text-bg-dark" : "text-bg-secondary")));

  top.appendChild(name);
  top.appendChild(badges);

  const meta = document.createElement("div");
  meta.className = "npc-meta";
  meta.textContent = [n.Species, n.Concept].filter(Boolean).join(" • ");

  wrap.appendChild(top);
  if (meta.textContent) wrap.appendChild(meta);

  a.appendChild(img);
  a.appendChild(wrap);

  a.addEventListener("click", () => openModal(idx));
  img.addEventListener("click", (e) => { e.stopPropagation(); openModal(idx); });

  return a;
}

function render() {
  const filtered = dataset.filter(matches);
  view = sortNpcs(filtered);

  els.count.textContent = `${view.length} shown / ${dataset.length} total`;
  els.list.innerHTML = "";

  view.forEach((n, i) => {
    els.list.appendChild(buildRow(n, i));
  });
}

/* ---------- Modal ---------- */

function kv(container, pairs) {
  container.innerHTML = "";
  const frag = document.createDocumentFragment();

  for (const [k, v] of pairs) {
    const val = clean(v);
    if (!val) continue;

    const row = document.createElement("div");
    row.className = "row g-2";

    const kc = document.createElement("div");
    kc.className = "col-5 col-xl-4 k";
    kc.textContent = k;

    const vc = document.createElement("div");
    vc.className = "col-7 col-xl-8 v";
    vc.textContent = val;

    row.appendChild(kc);
    row.appendChild(vc);
    frag.appendChild(row);
  }

  container.appendChild(frag);
}

function openModal(idx) {
  const n = view[idx];
  if (!n) return;

  els.modalTitle.textContent = n.Name;
  els.modalSubtitle.textContent = [n.Species, n.Origin, n.Concept].filter(Boolean).join(" • ");
  els.modalFooter.textContent = `Reputation: ${n.Reputation} • Relation: ${n.Relation}`;

  // image right
  els.modalImg.alt = n.Name;
  els.modalImg.src = n.imageUrl;
  els.modalImg.onerror = () => {
    els.modalImg.removeAttribute("src");
    els.modalImg.alt = `${n.Name} (missing image)`;
  };

  kv(els.kvIdentity, [
    ["Name", n.Name],
    ["Gender", n.Gender],
    ["Age", n.Age],
    ["Species", n.Species],
    ["Origin", n.Origin],
    ["Size", n.Size],
  ]);

  kv(els.kvSocial, [
    ["Reputation", n.Reputation],
    ["Relation", n.Relation],
    ["Presence", n.Presence],
    ["Wisdom", n.Wisdom],
    ["Expression", n.Expression],
    ["Instinct", n.Instinct],
  ]);

  kv(els.kvStats, [
    ["Health", n.Health],
    ["Spirit", n.Spirit],
    ["Magic", n.Magic],
    ["Healing", n.Healing],
    ["Agility", n.Agility],
    ["Strength", n.Strength],
    ["Dexterity", n.Dexterity],
    ["Stamina", n.Stamina],
    ["Intelligence", n.Intelligence],
    ["Perception", n.Perception],
    ["Will", n.Will],
    ["Wits", n.Wits],
  ]);

  kv(els.kvGear, [
    ["Equipment", n.Equipment],
    ["Armor", n.Armor],
    ["Shield", n.Shield],
    ["Weapon", n.Weapon],
    ["Special", n.Special],
  ]);

  els.skillsBox.textContent = clean(n.Skills) || "—";
  els.descBox.textContent = clean(n.Description) || "—";

  modal = bootstrap.Modal.getOrCreateInstance(els.modalEl);
  modal.show();
}

/* ---------- App lifecycle ---------- */

function applyData(json) {
  dataset = json?.npcs ?? [];
  initFilters();
  render();

  if (json) setStatus(`Loaded ${json.count} NPCs (cached: ${json.updatedAt})`);
  else setStatus("No cached data. Upload an Excel file to begin.");
}

function wireEvents() {
  els.file.addEventListener("change", async () => {
    const file = els.file.files?.[0];
    if (!file) return;

    setStatus("Reading Excel...");
    try {
      const json = await parseExcelFile(file);
      saveCache(json);
      applyData(json);
    } catch (e) {
      console.error(e);
      setStatus("Failed to read Excel. Make sure it's .xlsx with headers in row 1.");
    }
  });

  els.clear.addEventListener("click", () => {
    clearCache();
    applyData(null);
  });

  for (const el of [els.q, els.origin, els.reputation, els.relation, els.sort, els.onlyRecruited, els.hideDefeated]) {
    el.addEventListener("input", render);
    el.addEventListener("change", render);
  }
}

function main() {
  wireEvents();

  const cached = loadCache();
  applyData(cached);
}

main();
