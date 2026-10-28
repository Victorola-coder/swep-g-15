// Draws the PPCLFMVRS diagrams and crops the app screenshots for the report and slides.
// Usage: node build-figures.js   (needs Google Chrome; writes into ./assets)
//
// Diagrams are SVG, rendered to PNG by headless Chrome. Screenshots in
// docs/screenshots/ppclfmvrs are 1440 × 900 (scale 2); the app's content sits in a
// centred column, so the copies used in the documents are cropped to that column
// to keep the text readable when printed.
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");

const CHROME = process.env.CHROME ?? "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const ASSETS = path.join(__dirname, "assets");
const SHOTS = path.join(__dirname, "..", "..", "screenshots", "ppclfmvrs");
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), "ppcl-fig-"));
fs.mkdirSync(path.join(ASSETS, "shots"), { recursive: true });

function render(html, w, h, out, scale = 2) {
  const file = path.join(TMP, `${path.basename(out, ".png")}.html`);
  fs.writeFileSync(file, html);
  execFileSync(CHROME, [
    "--headless=new", "--disable-gpu", "--hide-scrollbars", `--user-data-dir=${path.join(TMP, "chrome")}`,
    `--window-size=${w},${h}`, `--force-device-scale-factor=${scale}`, `--screenshot=${out}`, `file://${file}`,
  ], { stdio: "ignore" });
  console.log("wrote", path.relative(process.cwd(), out));
}

// ---------- shared drawing helpers ----------
const INK = "#1f1a27", MUTED = "#5d5668", LINE = "#c9c3d1", GROUP = "#f1eff4", BOX = "#ffffff";
const PLUM = "#4a2a63", PLUM_FILL = "#ece4f3", CORAL = "#d9643f";
const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");

function box(x, y, w, h, title, lines = [], { fill = BOX, stroke = LINE, width = 2, titleColor = INK, sub = MUTED } = {}) {
  const cx = x + w / 2;
  const total = 30 + lines.length * 26;
  let ty = y + (h - total) / 2 + 24;
  let out = `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="12" fill="${fill}" stroke="${stroke}" stroke-width="${width}"/>`;
  out += `<text x="${cx}" y="${ty}" text-anchor="middle" font-size="24" font-weight="600" fill="${titleColor}">${esc(title)}</text>`;
  for (const l of lines) { ty += 27; out += `<text x="${cx}" y="${ty}" text-anchor="middle" font-size="19.5" fill="${sub}">${esc(l)}</text>`; }
  return out;
}
const group = (x, y, w, h, label) =>
  `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="14" fill="${GROUP}" stroke="${LINE}" stroke-width="2"/>` +
  `<text x="${x + 26}" y="${y + 36}" font-size="23" font-weight="600" fill="${INK}">${esc(label)}</text>`;
function arrow(x1, y1, x2, y2, { color = "#9b93a6", both = false, label, lx, ly, anchor = "start" } = {}) {
  let out = `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${color}" stroke-width="2.5" marker-end="url(#${color === CORAL ? "headCoral" : "head"})"${both ? ' marker-start="url(#tail)"' : ""}/>`;
  if (label) out += `<text x="${lx}" y="${ly}" font-size="19" fill="${MUTED}" text-anchor="${anchor}">${esc(label)}</text>`;
  return out;
}
const svgPage = (w, h, title, body) => `<!doctype html><html><head><meta charset="utf-8"><style>
  html,body{margin:0;background:#fff} svg{display:block;font-family:"Helvetica Neue",Helvetica,Arial,sans-serif}
</style></head><body><svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
<defs>
  <marker id="head" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="#9b93a6"/></marker>
  <marker id="tail" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="#9b93a6"/></marker>
  <marker id="headCoral" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="${CORAL}"/></marker>
</defs>
<text x="40" y="52" font-size="29" font-weight="600" fill="${INK}">${esc(title)}</text>
${body}</svg></body></html>`;

// ---------- Figure: system architecture ----------
{
  const W = 1340, H = 1030;
  let b = "";
  // Browser
  b += group(40, 85, 1260, 190, "In the browser (phone or computer)");
  b += box(70, 140, 375, 108, "Owner pages", ["report lost, ranked candidates,", "claim, handover code and QR"]);
  b += box(482, 140, 375, 108, "Finder pages", ["report found, write 2–3", "questions, choose a desk"]);
  b += box(895, 140, 375, 108, "Desk dashboard", ["statistics, confirm a code", "or scan the QR, failed claims"]);
  // Arrows browser <-> server
  b += arrow(330, 278, 330, 352, { label: "JSON requests + session cookie", lx: 345, ly: 322 });
  b += arrow(1010, 352, 1010, 278, { label: "allowlisted JSON only", lx: 995, ly: 322, anchor: "end" });
  // Server
  b += group(40, 355, 1260, 345, "On the server: Next.js route handlers (Vercel in production)");
  b += box(70, 410, 560, 105, "API routes (17)", ["session check, JSON-only requests,", "sign-in limit: 5 per account, 30 per IP"]);
  b += box(710, 410, 560, 105, "Allowlist serializers + leak guard", ["copy named fields only; throw if a", "private field would be sent"], { fill: PLUM_FILL, stroke: PLUM, width: 3, titleColor: PLUM });
  b += arrow(330, 518, 330, 566);
  b += arrow(1010, 566, 1010, 518);
  b += box(70, 568, 375, 108, "Matching engine", ["category gate; colour, place,", "date and keyword scores"]);
  b += box(482, 568, 375, 108, "Claims and verification", ["scrypt answer hashes, 3 / 5 limits,", "one-time handover codes"]);
  b += box(895, 568, 375, 108, "Reports and desk", ["validation, sensitive-text check,", "desk queries and statistics"]);
  // Server <-> DB
  b += arrow(670, 703, 670, 774, { both: true, label: "SQL in transactions with row locks", lx: 690, ly: 745 });
  // Database
  b += group(40, 777, 1260, 230, "PostgreSQL: one schema, ppclfmvrs, with 11 tables");
  b += `<text x="670" y="851" text-anchor="middle" font-size="19.5" fill="${MUTED}">users · sessions · lost_items · found_items · questions · matches · claims · claim_attempts · handover_codes · notifications · login_failures</text>`;
  b += box(70, 878, 580, 100, "PGlite (local and tests)", ["Postgres in the Node process; no server"]);
  b += box(690, 878, 580, 100, "Supabase (production)", ["transaction pooler, port 6543"]);
  render(svgPage(W, H, "Private details stay on the server; every response passes the leak guard", b), W, H, path.join(ASSETS, "architecture.png"));
}

// ---------- Figure: life of an item and a claim ----------
{
  const W = 1340, H = 640;
  let b = "";
  const steps = [
    ["Reported", ["lost or found,", "public fields only"]],
    ["Matched", ["scored 0–100%;", "35% or more shown"]],
    ["Claimed", ["owner picks a", "candidate"]],
    ["Verified", ["all answers right;", "code issued (72 h)"]],
    ["Handed over", ["code used once;", "both recovered"]],
  ];
  steps.forEach(([t, lines], i) => {
    const x = 40 + i * 258;
    const strong = i === 4;
    b += box(x, 105, 222, 130, t, lines, strong ? { fill: PLUM, stroke: PLUM, titleColor: "#fff", sub: "#e9e1f0" } : {});
    if (i < steps.length - 1) b += arrow(x + 224, 170, x + 256, 170);
  });
  // Branches (step centres: 151, 409, 667, 925, 1183)
  b += box(40, 333, 470, 122, "Closed", ["owner, finder or desk closed it, or an admin", "closed it after 7–365 days without activity"]);
  b += `<line x1="151" y1="238" x2="151" y2="330" stroke="#9b93a6" stroke-width="2.5" stroke-dasharray="7 6" marker-end="url(#head)"/>`;
  b += arrow(667, 238, 667, 330, { color: CORAL });
  b += box(538, 333, 258, 122, "Locked", ["3 wrong tries on the claim;", "owner must go to a desk"], { stroke: CORAL, titleColor: CORAL });
  b += arrow(905, 238, 905, 330);
  b += box(814, 333, 222, 122, "Cancelled", ["other open claims", "on the same item"]);
  b += arrow(985, 238, 1150, 330);
  b += box(1072, 333, 228, 122, "New code", ["if the owner lost it;", "old code revoked"]);
  b += `<text x="40" y="520" font-size="21" fill="${MUTED}">Limits: 3 answer tries per claim and 5 per user in any 24 hours. Every attempt is logged; failures appear on the desk dashboard.</text>`;
  b += `<text x="40" y="556" font-size="21" fill="${MUTED}">Before verification neither side sees the other. Afterwards the owner sees only where to collect; names and numbers are never shown.</text>`;
  b += `<text x="40" y="592" font-size="21" fill="${MUTED}">Claim: open → verified → completed (or locked, cancelled). Found item: open → pending handover → recovered.</text>`;
  render(svgPage(W, H, "The life of an item and a claim", b), W, H, path.join(ASSETS, "lifecycle.png"));
}

// ---------- Screenshots cropped to the app's content column ----------
const CROP = { x: 200, w: 1032, h: 900 }; // CSS pixels of the 1440 × 900 capture
for (const name of fs.readdirSync(SHOTS).filter((f) => f.endsWith(".png"))) {
  const src = path.join(SHOTS, name);
  const html = `<!doctype html><html><head><style>html,body{margin:0;overflow:hidden;background:#fff}
    img{position:absolute;left:-${CROP.x}px;top:0;width:1440px;height:900px}</style></head>
    <body><img src="file://${src}"></body></html>`;
  render(html, CROP.w, CROP.h, path.join(ASSETS, "shots", name), 1.5);
}
fs.rmSync(TMP, { recursive: true, force: true });
