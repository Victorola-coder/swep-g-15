// Builds the SCTPSTVS SWEP defence slides as a PowerPoint file (and a PDF beside it when LibreOffice is installed).
// Usage: node build-deck.js [output.pptx]   (default: ~/Desktop)
//
// The slides follow Dr Solanke's checklist: names on the cover, a short
// introduction, citations, a one-line caption on every image, mitigation
// for each limitation, and a thank-you slide at the end.
const fs = require("fs");
const os = require("os");
const path = require("path");

// pptxgenjs, and the theme helper's own dependencies, live in ../node_modules.
process.env.NODE_PATH = [path.join(__dirname, "..", "node_modules"), process.env.NODE_PATH].filter(Boolean).join(path.delimiter);
require("module").Module._initPaths();
const pptxgen = require("pptxgenjs");

process.env.PROJECT ??= "sctpstvs";
const group = require("../group");

const OUT = process.argv[2] ?? path.join(os.homedir(), "Desktop", "SCTPSTVS Defence Slides - SWEP Group 15.pptx");
const SHOTS = path.join(__dirname, "..", "..", "screenshots", "sctpstvs");
const ASSETS = path.join(__dirname, "assets");

// pptxgenjs cannot write theme colours, so the pptx skill's helper patches them in afterwards.
function findApplyTheme() {
  if (process.env.APPLY_THEME) return require(process.env.APPLY_THEME).applyTheme;
  const root = path.join(os.homedir(), ".claude", "skills", "synced");
  for (const dir of fs.existsSync(root) ? fs.readdirSync(root) : []) {
    const file = path.join(root, dir, "pptx", "scripts", "apply_theme.js");
    if (fs.existsSync(file)) return require(file).applyTheme;
  }
  return null;
}

// Deep blue and teal for transport, with an amber accent.
const THEME = {
  name: "SCTPSTVS Transit",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "16212E", lt1: "FFFFFF", dk2: "0E3A5C", lt2: "EDF3F8",
    accent1: "F2A93B", accent2: "0F766E", accent3: "1D4ED8", accent4: "C2410C", accent5: "5B6B7A", accent6: "B91C1C",
    hlink: "1D4ED8", folHlink: "6B21A8",
  },
};
const HEX = THEME.colors;
const GRID = "C8D3DD";

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 × 5.625 in
pres.title = "SCTPSTVS — SWEP Group 15 Project Defence";
pres.author = group.groupName;
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.35, w: 8.8, h: 1.8, fontFace: "Cambria", fontSize: 36, bold: true, color: C.background1, valign: "bottom", align: "left" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 3.3, w: 8.8, h: 1.2, fontFace: "Calibri", fontSize: 18, color: C.background2, valign: "top" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9, h: 0.7, fontFace: "Cambria", fontSize: 30, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: "SCTPSTVS · SWEP Group 15 · OAU", options: { x: 0.5, y: 5.2, w: 5, h: 0.3, fontFace: "Calibri", fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
  slideNumber: { x: 9.0, y: 5.2, w: 0.5, h: 0.3, fontFace: "Calibri", fontSize: 10, color: C.accent5, align: "right" },
});
pres.defineSlideMaster({
  title: "STATEMENT_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.45, w: 8.8, h: 0.8, fontFace: "Cambria", fontSize: 30, bold: true, color: C.background1, valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 9.0, y: 5.2, w: 0.5, h: 0.3, fontFace: "Calibri", fontSize: 10, color: C.background2, align: "right" },
});

// ---------- helpers ----------
const text = (slide, value, opts) => slide.addText(value, { isTextBox: true, fontFace: "Calibri", color: C.text1, margin: 0, valign: "top", ...opts });
const card = (slide, x, y, w, h, name, fill = C.background2) =>
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill }, objectName: name });
const pill = (slide, x, y, w, label, fill, color) =>
  slide.addText(label, { isTextBox: true, x, y, w, h: 0.3, fontFace: "Calibri", fontSize: 11, bold: true, color, fill: { color: fill }, align: "center", valign: "middle", margin: 0, rectRadius: 0.15, shape: pres.shapes.ROUNDED_RECTANGLE });
const image = (slide, file, x, y, w, h, alt) =>
  slide.addImage({ path: path.isAbsolute(file) ? file : path.join(SHOTS, file), x, y, w, h, altText: alt, sizing: { type: "contain", w, h } });
// A thin frame so light screenshots stand out from the white slide.
const framed = (slide, file, x, y, w, h, alt) => {
  slide.addShape(pres.shapes.RECTANGLE, { x: x - 0.02, y: y - 0.02, w: w + 0.04, h: h + 0.04, fill: { color: GRID }, line: { color: GRID }, objectName: `frame-${path.basename(file, ".png")}` });
  image(slide, file, x, y, w, h, alt);
};
let figNo = 0;
const fig = (textValue) => `Figure ${++figNo}: ${textValue}`;
// A one-line caption under every image (Dr Solanke: "a really short info on each image").
const caption = (slide, value, x, y, w, h = 0.25) => text(slide, value, { x, y, w, h, fontSize: 11, italic: true, color: C.accent5 });
// Every slide that uses a fact names its source in the same place.
const source = (slide, value) => text(slide, `Source: ${value}`, { x: 0.5, y: 4.88, w: 9, h: 0.25, fontSize: 10, color: C.accent5 });
const head = (t) => ({ text: t, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 } } });
const tableOpts = (extra) => ({ fontFace: "Calibri", color: HEX.dk1, border: { type: "solid", pt: 0.75, color: GRID }, valign: "middle", ...extra });
const bulletList = (items) => items.map((t, i, a) => ({ text: t, options: { bullet: true, breakLine: i < a.length - 1 } }));
const content = (title, section) => {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  return s;
};

// ---------- 1. Cover: title and names ----------
pres.addSection({ title: "Introduction" });
{
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Introduction" });
  pill(s, 0.6, 0.45, 4.2, "SWEP GROUP 15 · OBAFEMI AWOLOWO UNIVERSITY", C.accent1, C.text1);
  text(s, "Secure Campus Transport Payment and Single-Use Ticket Verification System (SCTPSTVS)", { x: 0.6, y: 0.9, w: 8.8, h: 1.2, fontFace: "Cambria", fontSize: 26, bold: true, color: C.background1, valign: "bottom" });
  text(s, "Pay for campus shuttle rides from a wallet; each signed ticket works once", { x: 0.6, y: 2.15, w: 8.8, h: 0.4, fontSize: 16, color: C.background2 });
  text(s, "PRESENTED BY", { x: 0.6, y: 2.75, w: 4, h: 0.25, fontSize: 10, bold: true, color: C.accent1 });
  if (group.hasMembers) {
    // Names in up to three columns.
    const cols = 3, perCol = Math.ceil(group.members.length / cols);
    for (let c = 0; c < cols; c++) {
      const names = group.members.slice(c * perCol, (c + 1) * perCol);
      if (!names.length) continue;
      text(s, names.map((m, i) => ({ text: m.name, options: { breakLine: i < names.length - 1 } })), { x: 0.6 + c * 2.95, y: 3.05, w: 2.85, h: 1.7, fontSize: 12, color: C.background1, paraSpaceAfter: 4 });
    }
  } else {
    text(s, "[Names of group members]", { x: 0.6, y: 3.05, w: 8.8, h: 0.35, fontSize: 14, color: C.background1 });
  }
  const footer = [group.supervisor && `Supervisor: ${group.supervisor}`, `Project defence${group.defenceDate ? `, ${group.defenceDate}` : ""}`].filter(Boolean).join(" · ");
  text(s, footer, { x: 0.6, y: 4.95, w: 8.8, h: 0.3, fontSize: 11, color: C.background2 });
  s.addNotes("Good morning. We are SWEP Group 15. Our project is SCTPSTVS, the Secure Campus Transport Payment and Single-Use Ticket Verification System. In the app we call it Campus Ride Pass. [Introduce the presenters.]");
}

// ---------- 2. Short introduction ----------
{
  const s = content("About this project", "Introduction");
  text(s, [
    { text: "SCTPSTVS lets a student pay a campus shuttle fare from a wallet on their phone and get a ticket with a signed QR code.", options: { breakLine: true } },
    { text: "The conductor scans it on the bus. Each ticket works once, on its route, until it expires.", options: { breakLine: true } },
    { text: "A digital signature stops fake tickets; the database stops a ticket being used twice." },
  ], { x: 0.5, y: 1.25, w: 3.9, h: 3.5, fontSize: 15, paraSpaceAfter: 12 });
  framed(s, "ticket-qr.png", 4.7, 1.25, 4.8, 3.0, "A student's QR ticket with its backup code");
  caption(s, fig("A ticket: signed QR code, backup code and expiry time."), 4.7, 4.32, 4.8);
  source(s, "Demo only: no real money moves, and the fares are example values, not OAU’s real fares.");
  s.addNotes("In short: a student pays a shuttle fare from a wallet on their phone and gets a ticket with a signed QR code. The conductor scans it on the bus, and each ticket works once, on its route, until it expires. A digital signature stops fake tickets, and the database stops a ticket being used twice. This is a working demo: no real money moves and the fares are example values.");
}

// ---------- 3. Problem ----------
{
  const s = content("Every common way of paying has a gap", "Introduction");
  const ways = [
    ["Cash on board", "Slow, and no record", "Change is needed at the door, and nobody knows how many fares were paid on each route."],
    ["Paper tickets", "Easy to copy", "A photocopy looks like the original, and a ticket can be shown twice."],
    ["Transfer alerts", "Checked by hand", "The conductor reads each payment alert, and a screenshot of an alert can be edited."],
  ];
  ways.forEach(([who, title, body], i) => {
    const x = 0.5 + i * 3.1;
    card(s, x, 1.25, 2.8, 2.85, `problem-card-${i + 1}`);
    pill(s, x + 0.25, 1.5, 2.3, who.toUpperCase(), C.text2, C.background1);
    text(s, title, { x: x + 0.25, y: 2.0, w: 2.3, h: 0.8, fontFace: "Cambria", fontSize: 18, bold: true, color: C.text2 });
    text(s, body, { x: x + 0.25, y: 2.85, w: 2.3, h: 1.1, fontSize: 14, color: C.text1 });
  });
  text(s, "A QR code alone does not fix this: it must be signed, and the system must remember which tickets have been used.", { x: 0.5, y: 4.3, w: 9, h: 0.55, fontSize: 15, italic: true, color: C.text2 });
  source(s, "ISO/IEC 18004:2015, QR Code bar code symbology specification.");
  s.addNotes("Every common way of paying has a gap. Cash on board is slow, needs change, and leaves no record. Paper tickets are easy to photocopy or show twice. Bank transfer alerts have to be checked by hand, and a screenshot can be edited. A plain QR code does not fix this on its own: a QR code is just a way of carrying data. It has to be signed, and the system has to remember which tickets were used.");
}

// ---------- 4. Implications for the campus ----------
{
  const s = content("What SCTPSTVS changes on campus", "Introduction");
  const items = [
    ["Students", "No cash or change; the ticket is on the phone before boarding"],
    ["Conductors", "One scan gives a large green or red result; no cash to count"],
    ["Transport office", "Live sales and rides per route, refunds and fare changes"],
    ["Fraud", "Copied, edited and reused tickets are refused and logged"],
    ["Accountability", "Every naira is a ledger row that cannot be edited"],
    ["Risks to manage", "Phones need data; the signing key must stay secret"],
  ];
  items.forEach(([title, body], i) => {
    const x = 0.5 + (i % 3) * 3.05, y = 1.25 + Math.floor(i / 3) * 1.75;
    const dark = i === 5;
    card(s, x, y, 2.85, 1.55, `impact-${i + 1}`, dark ? C.text2 : C.background2);
    text(s, title, { x: x + 0.2, y: y + 0.15, w: 2.45, h: 0.4, fontFace: "Cambria", fontSize: 16, bold: true, color: dark ? C.accent1 : C.text2 });
    text(s, body, { x: x + 0.2, y: y + 0.6, w: 2.45, h: 0.85, fontSize: 12.5, color: dark ? C.background1 : C.text1 });
  });
  s.addNotes("What does this change on campus? Students need no cash or change. Conductors get a clear green or red screen from one scan. The transport office sees sales and rides per route live. Cheating attempts are refused and logged, and every naira is a ledger row that cannot be edited. There are risks to manage: the conductor's phone needs data, and the signing key must stay secret.");
}

// ---------- 5. Aim and objectives ----------
{
  const s = content("Aim and objectives", "Introduction");
  card(s, 0.5, 1.25, 3.6, 3.7, "aim-card", C.text2);
  pill(s, 0.8, 1.5, 0.9, "AIM", C.accent1, C.text1);
  text(s, "Build a secure web app for paying campus transport fares, with tickets that are checked on board and used only once.", { x: 0.8, y: 2.0, w: 3.0, h: 2.6, fontFace: "Cambria", fontSize: 19, color: C.background1 });
  const objectives = [
    "A wallet built from a ledger that cannot be changed",
    "Fares set by the server, never by the phone",
    "Digitally signed QR tickets with a backup code",
    "One scan to check; each ticket accepted once",
    "Catch fake, reused, expired and wrong-route tickets",
    "Sales, refunds, fraud log and fares for the admin",
    "Prove it with tests, including simultaneous requests",
  ];
  objectives.forEach((o, i) => {
    const y = 1.25 + i * 0.53;
    s.addShape(pres.shapes.OVAL, { x: 4.45, y: y + 0.04, w: 0.36, h: 0.36, fill: { color: C.accent2 }, line: { color: C.accent2 }, objectName: `objective-dot-${i + 1}` });
    text(s, String(i + 1), { x: 4.45, y: y + 0.04, w: 0.36, h: 0.36, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" });
    text(s, o, { x: 4.95, y, w: 4.55, h: 0.45, fontSize: 14, valign: "middle" });
  });
  s.addNotes("Our aim was to build a secure web app for paying campus transport fares, with tickets that are checked on board and used only once. We broke that into seven objectives: the wallet, server-side fares, signed QR tickets, a one-scan check, catching every kind of bad ticket, the admin tools, and tests that prove it all works even when requests arrive at the same moment.");
}

// ---------- 6. Features and specifications ----------
{
  const s = content("Distinct features and specifications", "Introduction");
  text(s, bulletList([
    "Wallet with demo top-up and admin credits", "Fare always set by the server", "Ed25519-signed QR ticket + backup code",
    "First scan wins; later scans show ALREADY USED", "Five clear results for conductors", "Admin sales, refunds and fraud log",
    "Rate limits shared by every server", "Phone-first, light and dark themes",
  ]), { x: 0.5, y: 1.25, w: 4.0, h: 3.55, fontSize: 14, paraSpaceAfter: 6 });
  s.addTable([
    [head("Item"), head("Specification")],
    ["Platform", "Web app, any modern browser"],
    ["Roles", "Student, conductor, admin"],
    ["Ticket", "Ed25519 signature, 2-hour validity"],
    ["Backup code", "10 characters, 50 random bits"],
    ["Storage", "PostgreSQL (Supabase); PGlite locally"],
    ["Demo data", "8 routes, 5 vehicles, 4 accounts"],
    ["Tests", "35 automated, all passing"],
  ], tableOpts({ x: 4.7, y: 1.25, w: 4.8, colW: [1.4, 3.4], fontSize: 12.5, rowH: 0.42 }));
  s.addNotes("These are the features that make SCTPSTVS different, and on the right its specifications: a web app with three roles, tickets signed with Ed25519 and valid for two hours, a ten-character backup code, PostgreSQL storage, and 35 automated tests that all pass.");
}

// ---------- 7. Existing ways ----------
{
  const s = content("How it compares with existing ways of paying", "Introduction");
  const ours = (t) => ({ text: t, options: { bold: true, fill: { color: "FDF0DA" } } });
  s.addTable([
    [head("Question"), head("Cash fare"), head("Paper ticket"), head("Transfer alert"), head("QR in general"), head("SCTPSTVS")],
    ["Quick to board?", "No: change", "Yes", "No: read by hand", "Yes", ours("Yes: one scan")],
    ["Stops copies and reuse?", "No ticket", "No", "No: can be edited", "If checked once", ours("Yes: signed, single-use")],
    ["Sales per route?", "Hand count", "Stubs", "Bank statement", "Usually", ours("Yes, live")],
    ["Log of cheating?", "No", "No", "No", "Depends", ours("Yes, time and bus")],
    ["Works without data?", "Yes", "Yes", "No", "Some systems", ours("No: needs the server")],
  ], tableOpts({ x: 0.5, y: 1.25, w: 9, colW: [1.9, 1.15, 1.2, 1.55, 1.4, 1.8], fontSize: 12.5, rowH: 0.52 }));
  text(s, "This compares how each method works, not measured figures.", { x: 0.5, y: 4.45, w: 9, h: 0.3, fontSize: 13, italic: true, color: C.text2 });
  source(s, "ISO/IEC 18004:2015 for QR codes; methods compared as described in the report, Table 4.1.");
  s.addNotes("Here is how SCTPSTVS compares with the usual ways of paying. Cash and transfer alerts are slow at the door. Paper tickets and screenshots can be copied or edited. QR tickets in general are only as strong as their checking. SCTPSTVS is quick, stops copies and reuse, and records sales and cheating attempts. Its honest weakness: the conductor's phone needs data, because only the server knows if a ticket was used.");
}

// ---------- 8. Tools and stack ----------
{
  const s = content("Tools and technology stack", "Introduction");
  const tools = [
    ["Next.js 16 + React 19", "the web app and API"], ["TypeScript", "type-checked code"], ["Tailwind CSS", "styling and themes"],
    ["Node.js crypto", "Ed25519, scrypt, SHA-256"], ["qrcode", "draws QR tickets"], ["@zxing/browser", "camera QR scanner"],
    ["PostgreSQL", "Supabase, in production"], ["PGlite", "local database for tests"], ["Vitest", "automated tests"],
    ["Git + GitHub", "version control"], ["Vercel", "hosting"], ["ESLint", "code quality"],
  ];
  tools.forEach(([name, use], i) => {
    const x = 0.5 + (i % 4) * 2.3, y = 1.25 + Math.floor(i / 4) * 1.15;
    card(s, x, y, 2.1, 0.98, `tool-${i + 1}`);
    text(s, name, { x: x + 0.15, y: y + 0.12, w: 1.8, h: 0.4, fontFace: "Cambria", fontSize: 13.5, bold: true, color: C.text2 });
    text(s, use, { x: x + 0.15, y: y + 0.52, w: 1.8, h: 0.35, fontSize: 11.5 });
  });
  source(s, "Tools needed: a computer with Node.js 20.9+, npm, Git and a browser; Supabase and Vercel accounts to deploy.");
  s.addNotes("This is our stack. The app is built with Next.js, React and TypeScript, styled with Tailwind. Node's built-in crypto module gives us Ed25519 signatures, scrypt password hashing and SHA-256. The qrcode library draws tickets and zxing reads them with the camera. Data lives in PostgreSQL on Supabase, and PGlite runs the same SQL locally and in tests. To run it you only need Node.js, npm, Git and a browser.");
}

// ---------- 9. How a ticket works ----------
pres.addSection({ title: "Design" });
{
  const s = content("How a ticket works", "Design");
  const steps = [
    ["Top up", "Demo credit or an admin credit goes into the wallet ledger"],
    ["Buy", "The server charges the route’s fare under a row lock"],
    ["Sign", "The server signs the ticket with its Ed25519 private key"],
    ["Scan", "The conductor scans; the server checks signature, expiry and route"],
    ["Use once", "One guarded UPDATE marks it used; a second scan fails"],
  ];
  steps.forEach(([title, body], i) => {
    const x = 0.5 + i * 1.86, strong = i === 4;
    card(s, x, 1.55, 1.6, 2.3, `step-${i + 1}`, strong ? C.text2 : C.background2);
    text(s, String(i + 1), { x: x + 0.15, y: 1.7, w: 1.3, h: 0.5, fontFace: "Cambria", fontSize: 26, bold: true, color: strong ? C.accent1 : C.accent2 });
    text(s, title, { x: x + 0.15, y: 2.25, w: 1.3, h: 0.45, fontFace: "Cambria", fontSize: 15, bold: true, color: strong ? C.background1 : C.text2 });
    text(s, body, { x: x + 0.15, y: 2.75, w: 1.3, h: 1.0, fontSize: 11.5, color: strong ? C.background1 : C.text1 });
    if (i < steps.length - 1) s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: x + 1.63, y: 2.6, w: 0.2, h: 0.18, rotate: 90, fill: { color: C.accent5 }, line: { color: C.accent5 }, objectName: `arrow-${i + 1}` });
  });
  text(s, "The signature proves the server issued the ticket; the database proves it has not been used.", { x: 0.5, y: 4.15, w: 9, h: 0.4, fontSize: 15, italic: true, color: C.text2 });
  source(s, "Bernstein et al. (2012) for Ed25519; PostgreSQL Global Development Group (2024) for row locking.");
  s.addNotes("This is the life of one ride. The student tops up the wallet. Buying a ticket charges the route's fare, set by the server, under a lock on the wallet. The server signs the ticket with its private key. The conductor scans it, and the server checks the signature, the expiry and the route. Finally one guarded database update marks it used, so a second scan fails. The signature proves the server issued it; the database proves it has not been used.");
}

// ---------- 10. System architecture ----------
{
  const s = content("System architecture", "Design");
  image(s, path.join(ASSETS, "architecture.png"), 0.5, 1.15, 5.3, 3.65, "System architecture diagram");
  caption(s, fig("Phones talk to the Next.js server; PostgreSQL keeps every ticket and naira."), 0.5, 4.88, 5.3);
  text(s, [
    { text: "Phones and browsers: ", options: { bold: true } },
    { text: "student, conductor and admin pages; they hold no secrets.", options: { breakLine: true } },
    { text: "Server: ", options: { bold: true } },
    { text: "API routes check roles and origin; only the server holds the signing key.", options: { breakLine: true } },
    { text: "Database: ", options: { bold: true } },
    { text: "PostgreSQL on Supabase decides single use and guards the ledger. No state lives on the server, so several copies can run." },
  ], { x: 6.1, y: 1.25, w: 3.4, h: 3.6, fontSize: 13.5, paraSpaceAfter: 10 });
  s.addNotes("This is the system architecture. The student, conductor and admin pages run in the browser and hold no secrets. The Next.js server checks every user's role and the origin of every request, and only the server holds the private signing key. PostgreSQL on Supabase stores the tickets and the wallet ledger and enforces the single-use rule. Because no state lives in the server's memory, several copies of the server can run at once.");
}

// ---------- 11. Signed tickets ----------
{
  const s = content("A signature makes tickets unforgeable", "Design");
  card(s, 0.5, 1.25, 4.6, 3.45, "token-card", C.text2);
  text(s, "THE TICKET TOKEN", { x: 0.8, y: 1.45, w: 4, h: 0.3, fontSize: 11, bold: true, color: C.accent1 });
  [["SCT1", "version"], ["payload", "ticket ID, route, issued, expires, key ID"], ["signature", "Ed25519, made with the server’s private key"]].forEach(([part, meaning], i) => {
    const y = 1.85 + i * 0.72;
    pill(s, 0.8, y, 1.2, part, C.accent1, C.text1);
    text(s, meaning, { x: 2.15, y: y - 0.05, w: 2.75, h: 0.6, fontSize: 13, color: C.background1, valign: "middle" });
  });
  text(s, "Change one character and the check fails: INVALID.", { x: 0.8, y: 4.1, w: 4.1, h: 0.4, fontSize: 13, bold: true, color: C.background1 });
  text(s, [
    { text: "Nobody can forge a ticket, ", options: { bold: true } },
    { text: "extend its expiry or move it to another route without the private key.", options: { breakLine: true } },
    { text: "Why Ed25519, not a shared secret (HMAC)? ", options: { bold: true } },
    { text: "The checking key is public, so a future offline scanner could check tickets without being able to make them.", options: { breakLine: true } },
    { text: "Backup code: ", options: { bold: true } },
    { text: "10 characters with 50 random bits, for typing when the camera fails." },
  ], { x: 5.4, y: 1.25, w: 4.1, h: 3.5, fontSize: 13.5, paraSpaceAfter: 10 });
  source(s, "Bernstein et al. (2012), High-speed high-security signatures; Josefsson and Liusvaara (2017), RFC 8032.");
  s.addNotes("Every ticket is a token in three parts: a version, a payload with the ticket ID, route and times, and an Ed25519 signature made with the server's private key. Change one character and the check fails, so the conductor sees INVALID. We chose Ed25519 rather than a shared secret because the checking key can be public: a future offline scanner could check tickets without being able to make them. The backup code is for when the camera can't scan.");
}

// ---------- 12. Single use ----------
{
  const s = content("The database makes each ticket single-use", "Design");
  card(s, 0.5, 1.25, 5.3, 1.75, "sql-card");
  text(s, [
    "UPDATE tickets",
    "SET status = 'redeemed', redeemed_at = now,",
    "    redeemed_by = conductor, redeemed_vehicle = bus",
    "WHERE id = ticket AND status = 'active'",
    "  AND expires_at > now",
    "RETURNING id",
  ].map((t, i, a) => ({ text: t, options: { breakLine: i < a.length - 1 } })), { x: 0.7, y: 1.4, w: 5.0, h: 1.8, fontFace: "Courier New", fontSize: 11.5, color: C.text2 });
  text(s, [
    { text: "PostgreSQL locks the ticket’s row while it updates it. ", options: { bold: true } },
    { text: "A second update waits, re-checks the WHERE clause, finds the ticket no longer active and changes nothing." },
  ], { x: 0.5, y: 3.25, w: 5.3, h: 1.2, fontSize: 14 });
  card(s, 6.1, 1.25, 3.4, 3.45, "single-use-result", C.text2);
  text(s, "20", { x: 6.35, y: 1.45, w: 2.9, h: 0.8, fontFace: "Cambria", fontSize: 44, bold: true, color: C.accent1 });
  text(s, "simultaneous scans of one ticket", { x: 6.35, y: 2.25, w: 2.9, h: 0.5, fontSize: 14, color: C.background1 });
  text(s, "1 VALID", { x: 6.35, y: 2.95, w: 2.9, h: 0.5, fontFace: "Cambria", fontSize: 24, bold: true, color: C.background1 });
  text(s, "19 ALREADY USED, with the time and bus of the first scan", { x: 6.35, y: 3.5, w: 2.9, h: 0.9, fontSize: 13, color: C.background2 });
  source(s, "PostgreSQL Global Development Group (2024), PostgreSQL 17 Documentation, Chapter 13: Concurrency Control.");
  s.addNotes("A valid signature only proves the server issued a ticket. Whether it has been used lives in the database. The last step of every check is this one update: set the status to redeemed, but only where the ticket is still active and not expired. Postgres locks the row, so a second update waits, re-checks the condition, and changes nothing. In our test, twenty simultaneous scans of one ticket gave exactly one VALID and nineteen ALREADY USED.");
}

// ---------- 13. Money and other protections ----------
{
  const s = content("Money is safe, and attacks are logged", "Design");
  const cards = [
    ["Server sets the fare", "The phone sends only the route. A fare of 1 kobo sent by the client is ignored."],
    ["Ledger that only grows", "Triggers block UPDATE, DELETE and TRUNCATE, and any row that would make a balance negative."],
    ["Row locks", "A purchase locks the wallet’s owner; a refund locks the ticket. No double spending, no double refunds."],
  ];
  cards.forEach(([title, body], i) => {
    const x = 0.5 + i * 3.05;
    card(s, x, 1.25, 2.85, 1.85, `money-${i + 1}`);
    text(s, title, { x: x + 0.2, y: 1.38, w: 2.45, h: 0.4, fontFace: "Cambria", fontSize: 16, bold: true, color: C.text2 });
    text(s, body, { x: x + 0.2, y: 1.83, w: 2.45, h: 1.2, fontSize: 12.5 });
  });
  s.addTable([
    [head("Attack"), head("What stops it")],
    ["Guessing passwords", "5 tries per account, 30 per IP in 15 min; scrypt hashing"],
    ["Guessing backup codes", "About 10¹⁵ codes; 30 checks per conductor per minute"],
    ["Request from another website", "Refused with 403"],
  ], tableOpts({ x: 0.5, y: 3.3, w: 9, colW: [3.0, 6.0], fontSize: 12.5, rowH: 0.36 }));
  source(s, "Percival (2009) for scrypt; OWASP Foundation (2024), Cross-Site Request Forgery Prevention Cheat Sheet.");
  s.addNotes("Money is protected in three ways. The server sets the fare: the phone only says which route. The ledger only grows: database triggers refuse any edit or delete, and refuse any row that would make a balance negative. And row locks stop double spending and double refunds. Other attacks are limited too: password guessing is rate limited and passwords are hashed with scrypt, backup codes are too many to guess at 30 checks a minute, and requests from other websites are refused.");
}

// ---------- 14–17. The system, every screenshot captioned ----------
pres.addSection({ title: "The system" });
const triple = (title, shots, notes, src) => {
  const s = content(title, "The system");
  shots.forEach(([file, alt, cap, step], i) => {
    const x = 0.5 + i * 3.05;
    framed(s, file, x, 1.2, 2.9, 1.81, alt);
    caption(s, fig(cap), x, 3.08, 2.9, 0.45);
    text(s, step, { x, y: 3.6, w: 2.9, h: 1.1, fontSize: 12.5 });
  });
  if (src) source(s, src);
  s.addNotes(notes);
  return s;
};
triple("Student: top up, buy, show the ticket", [
  ["student-wallet.png", "Student wallet with demo top-up and tickets", "Wallet, demo top-up and tickets.", "1  Top up the wallet (demo credit, up to ₦20,000)."],
  ["student-buy.png", "Choosing a route before paying", "Choosing a route; the server sets the fare.", "2  Pick a route and pay from the wallet."],
  ["ticket-qr.png", "Full-screen QR ticket", "The QR ticket and backup code.", "3  Show the QR, or read out the backup code."],
], "As a student, you sign in with your matric number. You top up the wallet with demo credit, choose a route, and pay. The fare comes from the server. Then you press Show ticket QR and show it to the conductor, with the backup code underneath in case the camera can't scan. [Switch to the live demo here if possible.]",
"Demo fares are example values, not OAU’s real fares.");

{
  const s = content("Conductor: check a ticket", "The system");
  framed(s, "conductor.png", 0.5, 1.15, 6.0, 3.75, "Conductor screen with vehicle, scanner and backup-code box");
  caption(s, fig("The conductor screen on Bus 1: camera scanner and backup-code box."), 0.5, 4.93, 6.0);
  text(s, [
    { text: "1  Choose the vehicle you are on", options: { bold: true, breakLine: true } },
    { text: "2  Scan the QR, or type the backup code", options: { bold: true, breakLine: true } },
    { text: "3  Read the full-screen result", options: { bold: true, breakLine: true } },
    { text: "A ticket for another route shows WRONG ROUTE and stays unused. Each result has its own vibration pattern." },
  ], { x: 6.8, y: 1.2, w: 2.7, h: 3.7, fontSize: 13, paraSpaceAfter: 8 });
  s.addNotes("The conductor chooses the vehicle, then scans the QR with the phone camera or types the backup code. If the ticket is for another route, the result is WRONG ROUTE and the ticket stays unused, so the student can still ride the right bus. Each result also has its own vibration pattern, so the conductor can feel it.");
}

triple("Results a conductor reads at a glance", [
  ["verify-valid.png", "Green VALID result", "VALID: the first scan.", "Let the passenger board. The ticket is now used."],
  ["verify-already-used.png", "Red ALREADY USED result", "ALREADY USED: scanned again on Bus 3.", "Shows the time, bus and conductor of the first scan."],
  ["verify-invalid.png", "INVALID result for an edited ticket", "INVALID: the signed data was edited.", "Not issued by the system, or changed. Logged as fraud."],
], "These are the result screens, designed to be read at a glance. Green VALID means board. Red ALREADY USED shows when and where the ticket was first used: here it was used at 11:21 on Bus 1, then shown again on Bus 3. INVALID means the code was not issued by the system or was changed. Every refusal goes in the fraud log.");

{
  const s = content("Admin: sales, tickets and the fraud log", "The system");
  [["admin-overview.png", "Admin overview with sales and per-route figures", "Overview: sales, rides and figures per route."], ["admin-fraud.png", "Fraud log with reuse attempts and invalid signatures", "Fraud log: every refused check, newest first."]].forEach(([file, alt, cap], i) => {
    const x = 0.5 + i * 4.65;
    framed(s, file, x, 1.2, 4.35, 2.72, alt);
    caption(s, fig(cap), x, 3.98, 4.35);
  });
  text(s, "The admin can also refund unused tickets, change fares, credit wallets and assign vehicles to routes.", { x: 0.5, y: 4.4, w: 9, h: 0.4, fontSize: 13.5, italic: true, color: C.text2 });
  s.addNotes("The admin dashboard shows tickets sold, sales, rides and fraud events, with a table per route. The fraud log lists every refused check with who checked it and on which vehicle: reuse attempts, invalid signatures, unknown codes and wrong routes. The admin can also refund unused tickets, change fares, credit wallets and assign vehicles.");
}

// ---------- 18. App life cycle ----------
pres.addSection({ title: "Life cycle" });
{
  const s = content("App life cycle", "Life cycle");
  const stages = [
    ["Plan and design", "By 23 Oct"], ["Set-up", "23 Oct"], ["Core and tests", "23 Oct"], ["API", "23 Oct"], ["Screens", "23 Oct"],
    ["Polish and docs", "23 Oct"], ["PostgreSQL", "27 Oct"], ["Live testing", "27–28 Oct"], ["Deploy", "Next"], ["Maintain", "Ongoing"],
  ];
  s.addShape(pres.shapes.LINE, { x: 0.7, y: 2.05, w: 8.6, h: 0, line: { color: C.accent5, width: 1.5 }, objectName: "timeline" });
  stages.forEach(([name, when], i) => {
    const x = 0.5 + (i % 5) * 1.83, y = i < 5 ? 1.25 : 3.0;
    if (i === 5) s.addShape(pres.shapes.LINE, { x: 0.7, y: 3.8, w: 8.6, h: 0, line: { color: C.accent5, width: 1.5 }, objectName: "timeline-2" });
    const done = i < 8;
    s.addShape(pres.shapes.OVAL, { x: x + 0.62, y: y + 0.62, w: 0.36, h: 0.36, fill: { color: done ? C.accent2 : C.background1 }, line: { color: C.accent2, width: 1.5 }, objectName: `stage-${i + 1}` });
    text(s, String(i + 1), { x: x + 0.62, y: y + 0.62, w: 0.36, h: 0.36, fontSize: 11, bold: true, color: done ? C.background1 : C.accent2, align: "center", valign: "middle" });
    text(s, name, { x, y: y + 1.05, w: 1.6, h: 0.35, fontSize: 12.5, bold: true, color: C.text2, align: "center" });
    text(s, when, { x, y: y + 1.38, w: 1.6, h: 0.3, fontSize: 11, color: C.accent5, align: "center" });
  });
  source(s, `Dates from the repository history: ${group.repoUrl}`);
  s.addNotes("This is how we built it. We set up the app on 23 October, then built the core with its tests, the API, and the student, conductor and admin screens the same day, and wrote the documentation. On 27 October we moved storage from SQLite to PostgreSQL so it can run online, which exposed some faults we then fixed. We then tested it live on Supabase. Deployment to Vercel and maintenance come next.");
}

// ---------- 19. Life of a ticket ----------
{
  const s = content("The life of a ticket", "Life cycle");
  image(s, path.join(ASSETS, "ticket-lifecycle.png"), 0.5, 1.2, 9.0, 3.76, "Ticket states: bought, active, then redeemed, expired or refunded");
  caption(s, fig("A ticket ends in one of three states: redeemed, expired or refunded."), 0.5, 4.95, 9.0);
  s.addNotes("A ticket is bought and becomes active, valid for two hours by default. From there it ends in one of three ways: redeemed by the first scan on the right route, expired when the time runs out, or refunded by an admin while unused. An expired ticket can still be refunded. A scan on the wrong route leaves it active.");
}

// ---------- 20. Capacity ----------
{
  const s = content("What the system can handle", "Life cycle");
  const stats = [["20 → 1", "simultaneous scans, one VALID"], ["10 → 1", "purchases, money for one ticket"], ["30 / min", "checks per conductor"], ["2 hours", "default ticket life"]];
  stats.forEach(([big, label], i) => {
    const x = 0.5 + i * 2.3;
    card(s, x, 1.25, 2.1, 1.85, `capacity-${i + 1}`);
    text(s, big, { x: x + 0.1, y: 1.45, w: 1.9, h: 0.85, fontFace: "Cambria", fontSize: 30, bold: true, color: C.text2, align: "center", valign: "middle" });
    text(s, label, { x: x + 0.1, y: 2.35, w: 1.9, h: 0.6, fontSize: 13, align: "center" });
  });
  text(s, [
    { text: "Any number of server copies: ", options: { bold: true } },
    { text: "nothing is kept in server memory or files; each uses up to 5 database connections.", options: { breakLine: true } },
    { text: "Limits: ", options: { bold: true } },
    { text: "top-ups up to ₦20,000; fares ₦10 to ₦5,000; 5 sign-ins per account per 15 minutes.", options: { breakLine: true } },
    { text: "Not yet measured: ", options: { bold: true } },
    { text: "a load test with hundreds of users at once." },
  ], { x: 0.5, y: 3.35, w: 9, h: 1.45, fontSize: 13.5, paraSpaceAfter: 6 });
  source(s, "Automated tests and source code in the repository (sctpstvs/src/lib/server).");
  s.addNotes("What can it handle? Twenty simultaneous scans of one ticket give exactly one VALID. Ten purchases at once from a wallet with money for one ticket give exactly one ticket. A conductor can check 30 tickets a minute, and a ticket lasts two hours by default. Nothing is kept in server memory, so we can run as many copies of the server as we need. We have not yet run a load test with hundreds of users; that is open work.");
}

// ---------- 21. Tests ----------
pres.addSection({ title: "Results" });
{
  const s = content("35 automated tests, all passing", "Results");
  s.addChart(pres.charts.BAR, [{ name: "Tests", labels: ["Tickets and wallet", "Ticket tokens", "Sign-in and sessions", "API route"], values: [19, 7, 5, 4] }], {
    x: 0.5, y: 1.2, w: 5.8, h: 3.4, barDir: "bar",
    chartColors: [HEX.dk2], invertedColors: [HEX.dk2],
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: HEX.dk1, dataLabelFontSize: 12, dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.dk1, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
    valAxisHidden: true, valAxisMaxVal: 22, valAxisMinVal: 0,
    valGridLine: { style: "none" }, catGridLine: { style: "none" },
    showLegend: false, showTitle: true, title: "Automated tests per test file", titleFontSize: 13, titleColor: HEX.dk1, titleFontFace: "+mn-lt",
  });
  caption(s, fig("Number of tests in each of the four test files."), 0.5, 4.62, 5.8);
  card(s, 6.6, 1.25, 2.9, 3.55, "tests-note");
  text(s, "35 / 35", { x: 6.8, y: 1.45, w: 2.5, h: 0.6, fontFace: "Cambria", fontSize: 30, bold: true, color: C.text2 });
  text(s, "tests pass, in 14.4 s", { x: 6.8, y: 2.05, w: 2.5, h: 0.35, fontSize: 13 });
  text(s, bulletList(["Forged and edited tickets rejected", "Ledger never negative", "Fare from the server only", "Also run on a real PostgreSQL server"]), { x: 6.8, y: 2.6, w: 2.55, h: 2.1, fontSize: 12.5, paraSpaceAfter: 5 });
  s.addNotes("We wrote 35 automated tests in four files, and all of them pass, in about fourteen seconds. Nineteen test the tickets and the wallet, seven test the signed tokens, five test sign-in and sessions, and four call the real API route. Because the local database runs one query at a time, we also ran the suite against a real PostgreSQL server, where the requests truly overlap.");
}

// ---------- 22. Concurrency statement ----------
{
  const s = pres.addSlide({ masterName: "STATEMENT_DARK", sectionTitle: "Results" });
  s.addText("Simultaneous requests cannot cheat it", { placeholder: "title" });
  const tiles = [
    ["20 scans at once", "1 VALID, 19 ALREADY USED"],
    ["10 purchases, money for one", "exactly 1 ticket"],
    ["3 refunds of one ticket", "refunded once"],
    ["Live on Supabase: 20 HTTP scans", "1 VALID, 19 ALREADY USED"],
    ["Live: client sends a fare of 1 kobo", "real ₦150 charged"],
    ["Live: edited QR / other website", "INVALID / refused with 403"],
  ];
  tiles.forEach(([what, result], i) => {
    const x = 0.6 + (i % 2) * 4.5, y = 1.45 + Math.floor(i / 2) * 0.95;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 4.2, h: 0.8, rectRadius: 0.08, fill: { color: "1D5079" }, line: { color: "1D5079" }, objectName: `result-${i + 1}` });
    text(s, what, { x: x + 0.2, y: y + 0.08, w: 3.8, h: 0.32, fontSize: 13, color: C.background2 });
    text(s, result, { x: x + 0.2, y: y + 0.38, w: 3.8, h: 0.36, fontSize: 16, bold: true, color: C.accent1 });
  });
  text(s, "The first three come from the automated tests; the last three from the live app on Supabase.", { x: 0.6, y: 4.4, w: 8.8, h: 0.4, fontSize: 14, color: C.background1 });
  s.addNotes("The most important result: requests that arrive at the same moment cannot cheat the system. In the tests, twenty scans at once gave exactly one VALID; ten purchases with money for only one ticket gave one ticket; three refunds at once paid out once. We then tested the live app on Supabase: twenty simultaneous scans over the internet gave one VALID and nineteen ALREADY USED, a fare of one kobo sent by the client was ignored and the real 150 naira charged, an edited QR was INVALID, and a request from another website was refused.");
}

// ---------- 23. Errors found and fixes ----------
{
  const s = content("Errors we found and fixed", "Results");
  s.addTable([
    [head("Problem"), head("Fix")],
    ["Data in a local SQLite file would be lost online", "Moved all data to PostgreSQL (Supabase)"],
    ["Rate limits held in one server’s memory", "Counters moved into a database table"],
    ["Two debits could both pass the balance check", "Lock the wallet owner’s row first"],
    ["A ticket could be refunded twice", "Refund locks the ticket row (FOR UPDATE)"],
    ["Race tests did not really overlap", "warmPool() opens connections before each race"],
    ["Each server would make its own signing key", "SIGNING_PRIVATE_KEY required in production"],
  ], tableOpts({ x: 0.5, y: 1.2, w: 9, colW: [4.6, 4.4], fontSize: 13, rowH: 0.48 }));
  source(s, `Every fix is a commit in the repository: ${group.repoUrl}`);
  s.addNotes("Testing found real errors, mostly when we moved to PostgreSQL for online hosting. A local data file would be lost on a serverless host, so we moved to Postgres. Rate limits lived in one server's memory, so we moved them into the database. Postgres lets writes run side by side, so two debits could both pass a balance check, and two refunds could both pay out; we fixed both with row locks. Our race tests weren't really racing until we opened the connections first. And production now requires one fixed signing key from the settings, and says clearly if it is missing.");
}

// ---------- 24. Limitations and mitigation ----------
pres.addSection({ title: "Conclusion" });
{
  const s = content("Limitations and how we mitigate them", "Conclusion");
  s.addTable([
    [head("Limitation"), head("Mitigation")],
    ["No real payments: top-ups are demo credit", "Credit wallets only after a payment provider confirms payment"],
    ["Routes and fares are invented", "Admin enters the real ones from the transport office"],
    ["Checking needs a network connection", "Signatures can be checked offline; sync used tickets later"],
    ["Demo accounts exist in production", "Change or remove them before real testers"],
    ["A ticket screenshot can be passed on", "It still works only once"],
    ["Losing the signing key cancels unused tickets", "Keep it in the host’s secrets; key ID allows rotation"],
  ], tableOpts({ x: 0.5, y: 1.2, w: 9, colW: [3.9, 5.1], fontSize: 13, rowH: 0.5 }));
  s.addNotes("Every limitation has a mitigation. There are no real payments yet, so a payment provider must confirm each top-up before launch. The routes and fares are invented; the admin can enter the real ones in minutes. Checking needs a network, but signatures can already be checked offline, and used tickets could sync later. Demo accounts must be changed before real testers. A screenshot can be passed on, but it still works only once. And the signing key must be kept safe; the key ID in each token allows rotation later.");
}

// ---------- 25. Installation and deployment ----------
{
  const s = content("Installation and deployment", "Conclusion");
  card(s, 0.5, 1.2, 4.35, 3.55, "install-local");
  text(s, "Run it locally", { x: 0.75, y: 1.35, w: 3.9, h: 0.4, fontFace: "Cambria", fontSize: 17, bold: true, color: C.text2 });
  text(s, ["git clone the repository", "cd swep-g-15/sctpstvs", "npm install", "npm run dev", "npm test"].map((t, i, a) => ({ text: t, options: { breakLine: i < a.length - 1 } })), { x: 0.75, y: 1.85, w: 3.9, h: 2.0, fontFace: "Courier New", fontSize: 13, paraSpaceAfter: 6 });
  text(s, "No database server needed: PGlite runs locally.", { x: 0.75, y: 4.15, w: 3.9, h: 0.4, fontSize: 12, italic: true, color: C.text2 });
  card(s, 5.15, 1.2, 4.35, 3.55, "install-deploy", C.text2);
  text(s, "Put it online", { x: 5.4, y: 1.35, w: 3.9, h: 0.4, fontFace: "Cambria", fontSize: 17, bold: true, color: C.accent1 });
  text(s, [
    "Supabase: copy the Transaction pooler URL (port 6543)",
    "npm run gen-key for SIGNING_PRIVATE_KEY",
    "Vercel project, Root Directory sctpstvs",
    "Set DATABASE_URL and SIGNING_PRIVATE_KEY, then deploy",
  ].map((t, i, a) => ({ text: `${i + 1}.  ${t}`, options: { breakLine: i < a.length - 1 } })), { x: 5.4, y: 1.85, w: 3.9, h: 2.2, fontSize: 13, paraSpaceAfter: 8, color: C.background1 });
  text(s, "Optional: DATABASE_SCHEMA, TICKET_TTL_MINUTES", { x: 5.4, y: 4.15, w: 3.9, h: 0.4, fontSize: 12, italic: true, color: C.background2 });
  s.addNotes("Running it locally takes five commands, and no database server is needed because PGlite runs Postgres inside Node. To put it online: copy Supabase's transaction pooler URL, generate a signing key with npm run gen-key, create a Vercel project with the root directory set to sctpstvs, set the two settings, and deploy. The app creates its tables on the first request.");
}

// ---------- 26. Implemented and open ----------
{
  const s = content("Implemented and open features", "Conclusion");
  const columns = [
    ["Implemented", ["Accounts, roles, scrypt passwords", "Wallet ledger, demo top-up, credits", "Signed QR tickets and backup codes", "Conductor scanner, five results", "Single use under simultaneous scans", "Admin: sales, refunds, fraud, fares", "PostgreSQL for online hosting"], C.background2],
    ["Open", ["Real payments via a provider", "OAU’s real routes and fares", "Link to student records", "Self-registration, password reset", "Offline checking", "Signing-key rotation screen", "Load test with many users"], C.text2],
  ];
  columns.forEach(([title, items, fill], i) => {
    const x = 0.5 + i * 4.65, dark = fill === C.text2;
    card(s, x, 1.2, 4.35, 3.6, `features-${i + 1}`, fill);
    text(s, title, { x: x + 0.25, y: 1.35, w: 3.9, h: 0.4, fontFace: "Cambria", fontSize: 17, bold: true, color: dark ? C.accent1 : C.text2 });
    text(s, bulletList(items), { x: x + 0.25, y: 1.85, w: 3.9, h: 2.85, fontSize: 13, paraSpaceAfter: 5, color: dark ? C.background1 : C.text1 });
  });
  s.addNotes("On the left is everything implemented and working. On the right are the open items for a real launch: real payments through a provider, OAU's real routes and fares, linking to student records, self-registration and password reset, offline checking, a screen for rotating the signing key, and a load test.");
}

// ---------- 27. Team and links ----------
{
  const s = content("Our team and project links", "Conclusion");
  const rows = group.members.map((m) => [m.name, m.matric, m.department]);
  s.addTable([[head("Name"), head("Matric number"), head("Department")], ...rows], tableOpts({ x: 0.5, y: 1.2, w: 9, colW: [3.3, 2.4, 3.3], fontSize: 12, rowH: 0.42 }));
  text(s, [
    { text: "Source code, guide and full history: ", options: { bold: true } },
    { text: group.repoUrl, options: { hyperlink: { url: group.repoUrl }, color: HEX.accent3 } },
  ], { x: 0.5, y: 4.0, w: 9, h: 0.4, fontSize: 14 });
  s.addNotes("This is our team, and the link to the repository. The sctpstvs folder has the full source code, the installation and deployment guide, the security design, the tests, and the history of every change and fix.");
}

// ---------- 28. References ----------
{
  const s = content("References", "Conclusion");
  const refs = [
    "Bernstein, D. J., Duif, N., Lange, T., Schwabe, P. and Yang, B.-Y. (2012). High-speed high-security signatures. Journal of Cryptographic Engineering, 2(2), 77–89.",
    "Crockford, D. Base 32. crockford.com/base32.html",
    "ISO/IEC 18004:2015. QR Code bar code symbology specification.",
    "Josefsson, S. and Liusvaara, I. (2017). Edwards-Curve Digital Signature Algorithm (EdDSA), RFC 8032. IETF.",
    "OWASP Foundation (2024). Cross-Site Request Forgery Prevention Cheat Sheet.",
    "Percival, C. (2009). Stronger key derivation via sequential memory-hard functions. BSDCan 2009.",
    "PostgreSQL Global Development Group (2024). PostgreSQL 17 Documentation, Chapter 13: Concurrency Control.",
    "Supabase. Connecting to Postgres: connection pooler. supabase.com/docs",
  ];
  const half = Math.ceil(refs.length / 2);
  [refs.slice(0, half), refs.slice(half)].forEach((list, c) => {
    text(s, list.map((r, i) => ({ text: r, options: { breakLine: i < list.length - 1 } })), { x: 0.5 + c * 4.6, y: 1.2, w: 4.3, h: 3.75, fontSize: 13, paraSpaceAfter: 9 });
  });
  s.addNotes("These are the sources we cited. The full reference list is in our report.");
}

// ---------- 29. Thank you ----------
{
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Conclusion" });
  s.addText("Thank you", { placeholder: "title" });
  s.addText([
    { text: "SCTPSTVS sells campus transport tickets that cannot be forged and work only once, even when scanned twenty times at the same moment.", options: { breakLine: true } },
    { text: "We welcome your questions.", options: { bold: true, color: HEX.accent1 } },
  ], { placeholder: "body" });
  s.addNotes("To conclude: SCTPSTVS sells campus transport tickets that cannot be forged and work only once, even when scanned twenty times at the same moment, and it keeps a record of every naira and every cheating attempt. Thank you for listening. We welcome your questions, and we can show the live app.");
}

pres.writeFile({ fileName: OUT }).then(async () => {
  const applyTheme = findApplyTheme();
  if (applyTheme) await applyTheme(OUT, THEME);
  else console.warn("Theme helper not found: colours fall back to Office defaults. Set APPLY_THEME to its path.");
  console.log("wrote", OUT);
  const soffice = "/Applications/LibreOffice.app/Contents/MacOS/soffice";
  if (!fs.existsSync(soffice)) return console.warn("LibreOffice not found: skipping the PDF.");
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), "lo-"));
  require("child_process").execFileSync(soffice, ["--headless", `-env:UserInstallation=file://${profile}`, "--convert-to", "pdf", "--outdir", path.dirname(OUT), OUT], { stdio: "ignore" });
  console.log("wrote", OUT.replace(/\.pptx$/, ".pdf"));
});
