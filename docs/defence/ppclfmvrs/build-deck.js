// Builds the PPCLFMVRS SWEP defence slides as a PowerPoint file (and a PDF beside it).
// Usage: node build-deck.js [output.pptx]   (default: ~/Desktop)
// Group details come from ../group.json and ../members.json (PROJECT=ppclfmvrs).
// Run `node build-figures.js` first to draw the diagrams and crop the screenshots.
//
// The slides follow Dr Solanke's checklist: names on the cover, a short
// introduction, citations, a one-line caption on every image and chart,
// mitigation for each limitation, and a thank-you slide at the end.
process.env.PROJECT ??= "ppclfmvrs";
const fs = require("fs");
const os = require("os");
const path = require("path");
// Reuse ../node_modules, also for the pptx skill's apply_theme.js, which looks for pptxgenjs on NODE_PATH.
process.env.NODE_PATH = [path.join(__dirname, "..", "node_modules"), process.env.NODE_PATH].filter(Boolean).join(path.delimiter);
require("module").Module._initPaths();
const pptxgen = require("pptxgenjs");
const group = require("../group");

const OUT = process.argv[2] ?? path.join(os.homedir(), "Desktop", "PPCLFMVRS Defence Slides - SWEP Group 15.pptx");
const ASSETS = path.join(__dirname, "assets");
const SHOTS = path.join(ASSETS, "shots");

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

// Deep plum for privacy, with a warm coral accent.
const THEME = {
  name: "PPCLFMVRS Plum",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "231A2E", lt1: "FFFFFF", dk2: "3D2152", lt2: "F2EEF6",
    accent1: "EE8B62", accent2: "6B4A9A", accent3: "2563EB", accent4: "C2410C", accent5: "6A6275", accent6: "B91C1C",
    hlink: "2563EB", folHlink: "7C3AED",
  },
};
const HEX = THEME.colors;
const STATEMENT_BOX = "55356E";
const OURS_FILL = "FCE7DD";
const TABLE_LINE = "D6CEDF";

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 × 5.625 in
pres.title = "PPCLFMVRS — SWEP Group 15 Project Defence";
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
    { text: { text: "PPCLFMVRS · SWEP Group 15 · OAU", options: { x: 0.5, y: 5.2, w: 5, h: 0.3, fontFace: "Calibri", fontSize: 10, color: C.accent5, margin: 0 } } },
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
// A thin frame around screenshots, so their light-grey page reads as a screen.
const frame = (slide, x, y, w, h, name) =>
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { type: "none" }, line: { color: TABLE_LINE, width: 0.75 }, objectName: name });
let figNo = 0;
const fig = (textValue) => `Figure ${++figNo}: ${textValue}`;
// A one-line caption under every image (Dr Solanke: "a really short info on each image").
const caption = (slide, value, x, y, w) => text(slide, value, { x, y, w, h: 0.25, fontSize: 11, italic: true, color: C.accent5 });
// Every slide that uses a fact names its source in the same place.
const source = (slide, value) => text(slide, `Source: ${value}`, { x: 0.5, y: 4.88, w: 9, h: 0.25, fontSize: 10, color: C.accent5 });
const content = (title, section) => {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  return s;
};
const head = (t) => ({ text: t, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 } } });
const tableStyle = (extra) => ({ fontFace: "Calibri", color: HEX.dk1, border: { type: "solid", pt: 0.75, color: TABLE_LINE }, valign: "middle", ...extra });
const list = (items, opts = {}) => items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1, ...opts } }));

// Screenshots are cropped to the app's content column: 1032 × 900, so width / height = 1.147.
const SHOT_RATIO = 1032 / 900;

// ---------- 1. Cover: title and names ----------
pres.addSection({ title: "Introduction" });
{
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Introduction" });
  pill(s, 0.6, 0.45, 4.2, "SWEP GROUP 15 · OBAFEMI AWOLOWO UNIVERSITY", C.accent1, C.text1);
  text(s, "Privacy-Preserving Campus Lost-and-Found Matching and Verified Recovery System (PPCLFMVRS)", { x: 0.6, y: 0.85, w: 8.8, h: 1.25, fontFace: "Cambria", fontSize: 26, bold: true, color: C.background1, valign: "bottom" });
  text(s, "Lost items matched and returned only to their owners, without posting anyone’s details", { x: 0.6, y: 2.15, w: 8.8, h: 0.4, fontSize: 16, color: C.background2 });
  text(s, "PRESENTED BY", { x: 0.6, y: 2.75, w: 4, h: 0.25, fontSize: 10, bold: true, color: C.accent1 });
  if (group.hasMembers) {
    const cols = 3, perCol = Math.ceil(group.members.length / cols);
    for (let c = 0; c < cols; c++) {
      const names = group.members.slice(c * perCol, (c + 1) * perCol);
      if (!names.length) continue;
      text(s, names.map((m, i) => ({ text: m.name, options: { breakLine: i < names.length - 1 } })), { x: 0.6 + c * 2.95, y: 3.05, w: 2.85, h: 1.7, fontSize: 12, color: C.background1, paraSpaceAfter: 2 });
    }
  } else {
    text(s, "[Names of group members]", { x: 0.6, y: 3.05, w: 8.8, h: 0.35, fontSize: 14, color: C.background1 });
  }
  const footer = [group.supervisor && `Supervisor: ${group.supervisor}`, `Project defence${group.defenceDate ? `, ${group.defenceDate}` : ""}`].filter(Boolean).join(" · ");
  text(s, footer, { x: 0.6, y: 4.95, w: 8.8, h: 0.3, fontSize: 11, color: C.background2 });
  s.addNotes("Good morning. We are SWEP Group 15. Our project is PPCLFMVRS, the Privacy-Preserving Campus Lost-and-Found Matching and Verified Recovery System. In the app we simply call it OAU Lost and Found. [Introduce the presenters.]");
}

// ---------- 2. Short introduction ----------
{
  const s = content("About this project", "Introduction");
  text(s, [
    { text: "PPCLFMVRS is a web app that returns lost items on the OAU campus to their real owners.", options: { breakLine: true } },
    { text: "Today, finders post photos in WhatsApp groups. The photos show names, lock screens and card numbers, and anyone who sees them can claim the item.", options: { breakLine: true } },
    { text: "Our app matches lost and found reports automatically, and hands an item over only after the owner proves it is theirs." },
  ], { x: 0.5, y: 1.25, w: 4.3, h: 3.5, fontSize: 15, paraSpaceAfter: 12 });
  const w = 4.2, h = w / SHOT_RATIO;
  image(s, "home.png", 5.3, 1.15, w, h, "The home page of OAU Lost & Found");
  frame(s, 5.3, 1.15, w, h, "home-frame");
  caption(s, fig("The home page: how it works, and recently found items."), 5.3, 4.88, w);
  source(s, "Federal Republic of Nigeria (2023), Nigeria Data Protection Act.");
  s.addNotes("In short: PPCLFMVRS returns lost items on campus to their real owners. Today a finder usually posts a photo in WhatsApp groups. That photo can show the owner's name, lock screen or card numbers, which is personal data the Nigeria Data Protection Act of 2023 aims to protect, and anyone who sees it can claim the item. Our app matches reports automatically and hands an item over only after the owner proves it is theirs.");
}

// ---------- 3. Problem ----------
{
  const s = content("Posting a lost item exposes it to everyone", "Introduction");
  const cards = [
    ["Details exposed", "Photos show private details", "Names on ID cards, lock-screen pictures, card and account numbers go to whole groups."],
    ["False claims", "Anyone can claim it", "Whoever has seen the post can describe the item well enough to say it is theirs."],
    ["Matching by chance", "Owners must find the post", "An owner scrolls many groups and notice boards; nobody compares lost with found."],
  ];
  cards.forEach(([tag, headLine, body], i) => {
    const x = 0.5 + i * 3.1;
    card(s, x, 1.25, 2.8, 2.85, `problem-card-${i + 1}`);
    pill(s, x + 0.25, 1.5, 2.3, tag.toUpperCase(), C.text2, C.background1);
    text(s, headLine, { x: x + 0.25, y: 2.0, w: 2.3, h: 0.8, fontFace: "Cambria", fontSize: 18, bold: true, color: C.text2 });
    text(s, body, { x: x + 0.25, y: 2.85, w: 2.3, h: 1.15, fontSize: 14, color: C.text1 });
  });
  text(s, "And when an item is handed over, nobody records who took it or whether they proved it was theirs.", { x: 0.5, y: 4.3, w: 9, h: 0.55, fontSize: 15, italic: true, color: C.text2 });
  s.addNotes("This is the problem. When a finder posts a photo, private details go to the whole group: names on ID cards, lock-screen pictures, card numbers. Anyone who has seen the post can describe the item and claim it. Owners have to scroll through many groups and hope to see the right post. And when the item is handed over, nobody records who took it or whether they proved it was theirs.");
}

// ---------- 4. Implications for the campus ----------
{
  const s = content("What PPCLFMVRS changes on campus", "Introduction");
  const items = [
    ["Owners", "One place to report; likely matches are found and ranked for them"],
    ["Finders", "Hand items in without posting photos or sharing a phone number"],
    ["Privacy", "No more names, card numbers or lock screens posted to large groups"],
    ["Fraud", "False claimants must answer questions they cannot see: 3 tries per claim"],
    ["Desks", "Security Unit and SUB desks get a dashboard and a record of each handover"],
    ["Safety", "Items a finder keeps are handed over at a public campus place"],
  ];
  items.forEach(([h, body], i) => {
    const x = 0.5 + (i % 3) * 3.05, y = 1.25 + Math.floor(i / 3) * 1.75;
    const dark = i === 2;
    card(s, x, y, 2.85, 1.55, `impact-${i + 1}`, dark ? C.text2 : C.background2);
    text(s, h, { x: x + 0.2, y: y + 0.15, w: 2.45, h: 0.4, fontFace: "Cambria", fontSize: 16, bold: true, color: dark ? C.accent1 : C.text2 });
    text(s, body, { x: x + 0.2, y: y + 0.6, w: 2.45, h: 0.85, fontSize: 12.5, color: dark ? C.background1 : C.text1 });
  });
  source(s, "Federal Republic of Nigeria (2023), Nigeria Data Protection Act.");
  s.addNotes("What does this change on campus? Owners get one place to report, and matches are found for them. Finders can hand items in without posting photos or sharing their number. Private details stop being posted to large groups, in line with the Data Protection Act. False claimants have to answer questions they cannot see. The Security Unit and SUB desks get a dashboard and a record of every handover. And items a finder keeps are handed over in a public place. The risks to manage are weak questions, checking ID at the desk, and protecting the database.");
}

// ---------- 5. Aim and objectives ----------
{
  const s = content("Aim and objectives", "Introduction");
  card(s, 0.5, 1.25, 3.6, 3.7, "aim-card", C.text2);
  pill(s, 0.8, 1.5, 0.9, "AIM", C.accent1, C.text1);
  text(s, "Design and build a web app that matches lost and found items on the OAU campus and returns them only to verified owners, without exposing private details.", { x: 0.8, y: 2.0, w: 3.0, h: 2.8, fontFace: "Cambria", fontSize: 18, color: C.background1 });
  const objectives = [
    "Lost reports with public and owner-only private details",
    "Found reports with a generic description and 2–3 questions",
    "Automatic matching with ranked, explained candidates",
    "Verify ownership with hashed answers and attempt limits",
    "Hand over with a one-time code and QR code",
    "A desk dashboard for handovers and failed claims",
    "Privacy enforced in code: nothing private leaves the server",
  ];
  objectives.forEach((o, i) => {
    const y = 1.25 + i * 0.53;
    s.addShape(pres.shapes.OVAL, { x: 4.45, y: y + 0.04, w: 0.36, h: 0.36, fill: { color: C.accent2 }, line: { color: C.accent2 }, objectName: `objective-dot-${i + 1}` });
    text(s, String(i + 1), { x: 4.45, y: y + 0.04, w: 0.36, h: 0.36, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" });
    text(s, o, { x: 4.95, y, w: 4.55, h: 0.45, fontSize: 14, valign: "middle" });
  });
  s.addNotes("Our aim was to build a web app that matches lost and found items on campus and returns them only to verified owners, without exposing private details. We broke that into seven objectives: the two kinds of report, automatic matching, verification with hashed answers and limits, one-time handover codes, a desk dashboard, and privacy enforced in the code itself.");
}

// ---------- 6. Features and specifications ----------
{
  const s = content("Distinct features and specifications", "Introduction");
  text(s, list([
    "Private details only the owner sees", "Public text checked for numbers and contacts", "Automatic matching, ranked and explained",
    "Questions with hashed answers", "3 tries per claim, 5 per day", "One-time handover code and QR",
    "Desk dashboard and failed-claim log", "No photo uploads, by design",
  ]), { x: 0.5, y: 1.25, w: 3.9, h: 3.55, fontSize: 14, paraSpaceAfter: 6 });
  s.addTable([
    [head("Item"), head("Specification")],
    ["Platform", "Web app, any modern browser"],
    ["Catalogue", "9 categories, 16 colours, 26 OAU places"],
    ["Matching", "Category gate + 4 weighted factors"],
    ["Handover code", "8 characters, single use, 72 hours"],
    ["Storage", "PostgreSQL (Supabase), 11 tables"],
    ["Interface", "17 API routes behind a leak guard"],
    ["Tests", "51 automated, all passing"],
  ], { x: 4.7, y: 1.25, w: 4.8, colW: [1.5, 3.3], fontSize: 12.5, rowH: 0.43, ...tableStyle() });
  s.addNotes("These are the features that make PPCLFMVRS different: private details only the owner sees, a check that stops phone numbers and other identifying numbers in public text, ranked matching, hashed answers with strict limits, one-time codes, a desk dashboard, and no photos at all. On the right are the specifications: 26 campus places, 17 API routes, 11 database tables and 51 automated tests.");
}

// ---------- 7. Existing ways ----------
{
  const s = content("Existing ways expose details or rely on luck", "Introduction");
  const ours = (t) => ({ text: t, options: { bold: true, fill: { color: OURS_FILL } } });
  s.addTable([
    [head("Way"), head("Details private"), head("Checks ownership"), head("Matches lost with found"), head("Records handover")],
    ["WhatsApp groups, notice boards", "No: photos are public", "Informally, by chat", "By chance", "No"],
    ["A desk logbook", "Yes", "Staff judgement", "By hand", "Yes, written"],
    ["Generic lost-and-found apps", "Varies by app", "Varies by app", "Varies by app", "Varies by app"],
    [ours("PPCLFMVRS"), ours("Yes, no photos"), ours("Hashed questions, limits"), ours("Automatically, ranked"), ours("Yes, one-time code")],
  ], { x: 0.5, y: 1.3, w: 9, colW: [2.2, 1.7, 1.75, 1.75, 1.6], fontSize: 13, rowH: 0.55, ...tableStyle() });
  source(s, "Compared by design from common practice; not measured. Generic apps differ, so we do not rate them.");
  s.addNotes("We compared the usual ways. WhatsApp groups and notice boards reach many people, but the photos are public and ownership is checked informally. A desk logbook keeps details private, but owners must visit the desk and matching is by hand. Generic lost-and-found apps differ from app to app, so we do not rate them, and none is built for OAU's places and desks. PPCLFMVRS keeps details private, checks ownership with hashed questions, matches automatically and records every handover.");
}

// ---------- 8. Tools ----------
{
  const s = content("Tools and technology stack", "Introduction");
  const tools = [
    ["Next.js 16 + React 19", "pages and API routes"], ["TypeScript", "type-checked code"], ["Tailwind CSS", "styling, light and dark"],
    ["Node.js crypto", "scrypt, SHA-256, timing-safe"], ["PostgreSQL / Supabase", "all data in production"], ["postgres.js", "pooler-safe driver"],
    ["PGlite", "local database for tests"], ["qrcode", "handover QR codes"], ["OpenStreetMap", "place coordinates"],
    ["Vitest", "51 automated tests"], ["Git + GitHub", "version control"], ["Vercel", "hosting"],
  ];
  tools.forEach(([name, use], i) => {
    const x = 0.5 + (i % 4) * 2.3, y = 1.25 + Math.floor(i / 4) * 1.15;
    card(s, x, y, 2.1, 0.98, `tool-${i + 1}`);
    text(s, name, { x: x + 0.15, y: y + 0.12, w: 1.8, h: 0.4, fontFace: "Cambria", fontSize: 13.5, bold: true, color: C.text2 });
    text(s, use, { x: x + 0.15, y: y + 0.52, w: 1.8, h: 0.35, fontSize: 11.5 });
  });
  source(s, "Tools needed: a computer with Node.js 22+, npm, Git and a web browser; free Supabase and Vercel accounts to deploy.");
  s.addNotes("This is our stack. The app is built with Next.js, React and TypeScript, styled with Tailwind. Hashing uses Node's built-in crypto: scrypt for passwords and answers, SHA-256 for tokens and codes. Data is in PostgreSQL on Supabase, and PGlite runs the same database inside Node for local work and tests. To run it you only need Node.js, npm, Git and a browser.");
}

// ---------- 9. How an item gets home ----------
pres.addSection({ title: "Design" });
{
  const s = content("How a lost item gets home", "Design");
  const h = 3.55, w = h * 1340 / 640, x = (10 - w) / 2;
  image(s, path.join(ASSETS, "lifecycle.png"), x, 1.1, w, h, "Life of an item and a claim: reported, matched, claimed, verified, handed over");
  caption(s, fig("From report to handover, with the ways a claim can stop."), x, 1.1 + h + 0.08, w);
  s.addNotes("This is how a lost item gets home. Both sides report. The system scores each pair and shows candidates from 35 percent. The owner claims one, answers the finder's questions, and if every answer is right gets a one-time code, valid for 72 hours. The desk or finder enters the code and the item is recovered. Three wrong answers lock the claim. Other claims on the item are cancelled once one owner is verified, and an owner who loses the code can get a new one, which revokes the old one.");
}

// ---------- 10. Architecture ----------
{
  const s = content("System architecture", "Design");
  const h = 3.68, w = h * 1340 / 1030;
  image(s, path.join(ASSETS, "architecture.png"), 0.5, 1.12, w, h, "System architecture diagram");
  caption(s, fig("Browser, server and database; every response passes the leak guard."), 0.5, 4.88, 5.6);
  text(s, [
    { text: "In the browser: ", options: { bold: true } },
    { text: "owner, finder and desk pages; nothing private is sent to them.", options: { breakLine: true } },
    { text: "On the server: ", options: { bold: true } },
    { text: "17 API routes, matching, claims and desk code. Responses pass allowlists and the leak guard.", options: { breakLine: true } },
    { text: "In the database: ", options: { bold: true } },
    { text: "all data and limits in Postgres, so many servers can run at once." },
  ], { x: 5.45 + 0.35, y: 1.25, w: 3.7, h: 3.5, fontSize: 13.5, paraSpaceAfter: 10 });
  s.addNotes("This is the architecture. The browser pages only ever receive allowlisted JSON. On the server, 17 API routes check the session and that requests are JSON, then call the matching engine, the claims code or the desk code. Every response is built by allowlist serializers and checked by a leak guard. All data, including the attempt counters, is in Postgres: Supabase in production and PGlite locally.");
}

// ---------- 11. Match scoring ----------
{
  const s = content("Match scoring: public fields only", "Design");
  s.addTable([
    [head("Factor"), head("Weight"), head("How it scores")],
    ["Category", "Gate", "Must be the same"],
    ["Colour", "25%", "Same 1; confusable 0.5"],
    ["Place", "25%", "Same 1, falling to 0 at 900 m"],
    ["Date", "20%", "Day of loss or next day 1; 0 after two weeks"],
    ["Keywords", "30%", "Dice overlap of description words"],
  ], { x: 0.5, y: 1.25, w: 5.6, colW: [1.2, 0.9, 3.5], fontSize: 13, rowH: 0.5, ...tableStyle() });
  text(s, "Shown from 35%; both sides notified from 50%.", { x: 0.5, y: 4.4, w: 5.6, h: 0.35, fontSize: 13, italic: true, color: C.text2 });
  card(s, 6.4, 1.25, 3.1, 3.5, "worked-example", C.text2);
  text(s, "93%", { x: 6.65, y: 1.4, w: 2.6, h: 0.75, fontFace: "Cambria", fontSize: 40, bold: true, color: C.accent1 });
  text(s, "the demo’s true match", { x: 6.65, y: 2.15, w: 2.6, h: 0.3, fontSize: 13, color: C.background1 });
  text(s, [{ text: "0.25 + 0.25 + 0.20", options: { breakLine: true } }, { text: "+ 0.30 × 0.77" }], { x: 6.65, y: 2.55, w: 2.6, h: 0.5, fontFace: "Courier New", fontSize: 12, color: C.background1 });
  text(s, "Same colour, same place, found the next day, 77% word overlap. A white phone from Moremi Hall scores 37%.", { x: 6.65, y: 3.2, w: 2.6, h: 1.4, fontSize: 12.5, color: C.background2 });
  source(s, "Dice (1945) for word overlap; Sinnott (1984) for distances from OpenStreetMap coordinates.");
  s.addNotes("Matching uses public fields only, so private details cannot leak through scores. The category must match. Then colour and place count 25 percent each, the date 20, and description words 30, using the Dice overlap. Places are compared by real distance, falling to zero at 900 metres. In our demo, the true match scores 93 percent, and a white phone from Moremi Hall scores only 37.");
}

// ---------- 12. Proving ownership ----------
{
  const s = content("Proving ownership", "Design");
  const cols = [
    ["Hashed answers", ["Finder writes 2–3 questions", "Answers normalised, then scrypt with a salt", "Compared with timingSafeEqual", "All answers checked, even after a wrong one"]],
    ["Attempt limits", ["3 wrong tries lock the claim", "5 tries per user in any 24 hours", "Every attempt logged for the desk", "Rows locked in one transaction"]],
    ["Handover code", ["8 characters, no 0/O or 1/I/L", "About 40 bits; stored as SHA-256", "Valid 72 hours, shown once", "One UPDATE uses it: reuse gets 409"]],
  ];
  cols.forEach(([h, items], i) => {
    const x = 0.5 + i * 3.05;
    const dark = i === 1;
    card(s, x, 1.2, 2.85, 3.55, `verify-${i + 1}`, dark ? C.text2 : C.background2);
    text(s, h, { x: x + 0.2, y: 1.35, w: 2.45, h: 0.4, fontFace: "Cambria", fontSize: 17, bold: true, color: dark ? C.accent1 : C.text2 });
    text(s, list(items, { color: dark ? HEX.lt1 : HEX.dk1 }), { x: x + 0.2, y: 1.85, w: 2.5, h: 2.8, fontSize: 14, paraSpaceAfter: 8 });
  });
  source(s, "Percival and Josefsson (2016), RFC 7914 (scrypt); Node.js crypto docs; PostgreSQL 17 docs, Explicit Locking.");
  s.addNotes("Proving ownership has three parts. First, the finder's answers are normalised, so capitals and spaces don't matter, and stored only as salted scrypt hashes, compared in constant time. Second, the limits: three wrong tries lock a claim, and five tries per user per day stop someone fishing across many items. These are enforced with row locks in one transaction, so parallel requests can't sneak past. Third, the handover code: eight characters, about 40 bits, valid 72 hours, stored only as a hash, and used by a single UPDATE, so reusing it gets a 409.");
}

// ---------- 13. Privacy in code ----------
{
  const s = content("Privacy is enforced in code", "Design");
  s.addTable([
    [head("Data"), head("Who can see it")],
    ["Public fields", "Everyone; the owner (candidates)"],
    ["Owner’s private details", "The owner only"],
    ["Verification answers", "Nobody: salted hashes"],
    ["Handover code", "The owner, once"],
    ["Names, phones, emails, ids", "Nobody else, ever"],
    ["Failed attempts", "Desk staff and admins"],
  ], { x: 0.5, y: 1.2, w: 4.9, colW: [2.3, 2.6], fontSize: 12.5, rowH: 0.47, ...tableStyle() });
  const points = [
    ["Allowlist serializers", "copy named fields only; a new column cannot leak"],
    ["Leak guard", "throws if a forbidden field appears at any depth"],
    ["No photos", "the main way details leak, and a gift to false claimants"],
  ];
  points.forEach(([h, body], i) => {
    const y = 1.2 + i * 1.18;
    card(s, 5.7, y, 3.8, 1.02, `privacy-${i + 1}`, i === 1 ? C.text2 : C.background2);
    text(s, h, { x: 5.9, y: y + 0.1, w: 3.4, h: 0.35, fontFace: "Cambria", fontSize: 15, bold: true, color: i === 1 ? C.accent1 : C.text2 });
    text(s, body, { x: 5.9, y: y + 0.47, w: 3.4, h: 0.5, fontSize: 12.5, color: i === 1 ? C.background1 : C.text1 });
  });
  source(s, "OWASP (2023), API Security Top 10, API3:2023 Broken Object Property Level Authorization.");
  s.addNotes("Privacy is not just a promise; it is in the code. The table shows who can see what. Every response is built by allowlist serializers that copy named fields only, then a leak guard walks the response and throws if a forbidden field such as an answer hash, phone number or owner id appears anywhere. This guards against what OWASP calls broken object property level authorization. And we take no photos at all, because photos are how details leak and how false claimants learn the answers.");
}

// ---------- The system: screenshots with captions ----------
pres.addSection({ title: "The system" });
const screenSlide = (title, file, alt, figure, side, notes) => {
  const s = content(title, "The system");
  const h = 3.7, w = h * SHOT_RATIO;
  image(s, file, 0.5, 1.13, w, h, alt);
  frame(s, 0.5, 1.13, w, h, "screen-frame");
  caption(s, figure, 0.5, 4.88, w + 1);
  text(s, side, { x: 5.0, y: 1.25, w: 4.5, h: 3.5, fontSize: 14, paraSpaceAfter: 9, color: C.text1 });
  s.addNotes(notes);
  return s;
};
const pairSlide = (title, left, right, side, notes) => {
  const s = content(title, "The system");
  const h = 3.3, w = h * SHOT_RATIO;
  [left, right].forEach(([file, alt, cap], i) => {
    const x = 0.5 + i * (w + 0.3);
    image(s, file, x, 1.13, w, h, alt);
    frame(s, x, 1.13, w, h, `screen-frame-${i + 1}`);
    caption(s, fig(cap), x, 4.48, w);
  });
  text(s, side, { x: 0.5, y: 4.8, w: 9, h: 0.3, fontSize: 12.5, color: C.text2 });
  s.addNotes(notes);
  return s;
};
screenSlide("Owner: report a lost item", "lost-new-private.png", "The lost-item form with the private identifying details section",
  fig("The lost-item form; private details stay with the owner."),
  [
    { text: "Public details: ", options: { bold: true } },
    { text: "category, colour, brand, a short description, place, date and time window.", options: { breakLine: true } },
    { text: "Private details: ", options: { bold: true } },
    { text: "last digits of the IMEI, lock screen, contents, unique marks.", options: { breakLine: true } },
    { text: "Never shown ", options: { bold: true } },
    { text: "to finders, desks or the public, and never used for matching." },
  ],
  "The owner reports a lost item with public details, which are used for matching, and private details, which only the owner can ever see. They help the owner remember what to answer, and help staff if a dispute reaches the desk.");
screenSlide("Owner: ranked candidates", "candidates.png", "Ranked candidates: a black phone at 93% and a white phone at 37%",
  fig("Ranked candidates with a bar for each factor."),
  [
    { text: "Best match first, ", options: { bold: true } },
    { text: "with a bar for colour, place, date and description.", options: { breakLine: true } },
    { text: "93% ", options: { bold: true } },
    { text: "for the true match; 37% for a white phone from Moremi Hall.", options: { breakLine: true } },
    { text: "Claim only what is yours: ", options: { bold: true } },
    { text: "3 tries per claim, 5 per day." },
  ],
  "The owner opens the report and sees the candidates, best first, with the score for each factor. The black phone found in the library scores 93 percent. A white phone from Moremi Hall scores 37. The page reminds the owner that tries are limited, so they should only claim what they believe is theirs.");
pairSlide("Owner: answer the questions, get the code",
  ["claim-wrong.png", "A wrong answer refused, with 2 tries left", "A wrong answer: 2 tries left on the claim."],
  ["claim-verified.png", "A verified claim with the one-time code and its QR code", "Verified: where to collect, the code and QR."],
  "The code is shown once and stored only as a hash. Asking for a new one revokes the old one.",
  "The owner claims the item and answers the finder's questions. On the left, a wrong answer is refused and two tries are left. On the right, with the right answers, the claim is verified: the owner sees where to collect, the Security Unit desk at the Senate Building, plus a one-time code and its QR code. The code is shown only once.");
pairSlide("Finder: report a found item",
  ["found-warning.png", "The found-item form warning about an email address in the description", "An email address in public text is flagged."],
  ["found-new.png", "Verification questions and where the item is kept", "2–3 questions, and where the item is kept."],
  "The server refuses public text with phone numbers, emails, matric numbers or long numbers. Answers are hashed at once.",
  "The finder reports what they found with a generic description. If they type contact details, like the email address here, the form flags them and the server refuses it. Then they write two or three questions about things only the owner would know, and choose whether the item goes to a collection desk or stays with them until a meeting at a public place.");
pairSlide("Desk: dashboard and handover",
  ["desk-dashboard.png", "The Security Unit desk dashboard with statistics", "The desk dashboard, one owner waiting."],
  ["desk-confirmed.png", "The handover confirmed after the code is entered", "Code entered: handover confirmed."],
  "The officer checks the owner’s ID card, then types or scans the code. Entering it again gives “already used”.",
  "Desk officers see statistics, items waiting for collection, and items held at the desk. Further down is the failed-claim log. When the owner arrives, the officer checks their ID card and types or scans the code. The handover is confirmed and the item is marked recovered. The same code entered again is refused.");

// ---------- Life cycle ----------
pres.addSection({ title: "Life cycle" });
{
  const s = content("Development life cycle", "Life cycle");
  const stages = [
    ["Plan and design", "By 23 Oct"], ["Set-up", "23 Oct"], ["Matching engine", "23 Oct"], ["Data and security", "23 Oct"], ["Claims and codes", "23 Oct"],
    ["API and interface", "23 Oct"], ["README", "23 Oct"], ["Postgres + locks", "27 Oct"], ["Live test", "After 27 Oct"], ["Deploy, maintain", "Next"],
  ];
  s.addShape(pres.shapes.LINE, { x: 0.7, y: 2.05, w: 8.6, h: 0, line: { color: C.accent5, width: 1.5 }, objectName: "timeline" });
  stages.forEach(([name, when], i) => {
    const x = 0.5 + (i % 5) * 1.83, y = i < 5 ? 1.25 : 3.0;
    if (i === 5) s.addShape(pres.shapes.LINE, { x: 0.7, y: 3.8, w: 8.6, h: 0, line: { color: C.accent5, width: 1.5 }, objectName: "timeline-2" });
    const done = i < 9;
    s.addShape(pres.shapes.OVAL, { x: x + 0.62, y: y + 0.62, w: 0.36, h: 0.36, fill: { color: done ? C.accent2 : C.background1 }, line: { color: C.accent2, width: 1.5 }, objectName: `stage-${i + 1}` });
    text(s, String(i + 1), { x: x + 0.62, y: y + 0.62, w: 0.36, h: 0.36, fontSize: 11, bold: true, color: done ? C.background1 : C.accent2, align: "center", valign: "middle" });
    text(s, name, { x, y: y + 1.05, w: 1.6, h: 0.35, fontSize: 12.5, bold: true, color: C.text2, align: "center" });
    text(s, when, { x, y: y + 1.38, w: 1.6, h: 0.3, fontSize: 11, color: C.accent5, align: "center" });
  });
  source(s, `git log of the ppclfmvrs folder, ${group.repoUrl}`);
  s.addNotes("This is how we built it. After planning, we set up the app and built the matching engine, the data and security layer, claims and handover codes, the API and the interface, and the README, all on 23 October, testing each step. On 27 October we moved storage to Postgres with row locks and added the sign-in rate limit, so it can run on serverless hosting. We then tested the full flow live on Supabase. Hosting on Vercel and maintenance come next.");
}
{
  const s = content("What the system can handle", "Life cycle");
  const stats = [["3 / 5", "answer tries per claim / per day"], ["72 h", "life of a one-time code"], ["8.5 × 10¹¹", "possible handover codes"], ["Any", "number of server instances"]];
  stats.forEach(([big, label], i) => {
    const x = 0.5 + i * 2.3;
    card(s, x, 1.25, 2.1, 1.85, `capacity-${i + 1}`);
    text(s, big, { x: x + 0.1, y: 1.45, w: 1.9, h: 0.85, fontFace: "Cambria", fontSize: 30, bold: true, color: C.text2, align: "center", valign: "middle" });
    text(s, label, { x: x + 0.1, y: 2.35, w: 1.9, h: 0.6, fontSize: 13, align: "center" });
  });
  text(s, [
    { text: "Sign-in limit: ", options: { bold: true } },
    { text: "5 failures per account and 30 per IP address in 15 minutes, then 429.", options: { breakLine: true } },
    { text: "All state in Postgres, ", options: { bold: true } },
    { text: "limits included, so many server copies can run at once (as on Vercel); at most 5 connections each.", options: { breakLine: true } },
    { text: "26 places, 9 categories, 16 colours; ", options: { bold: true } },
    { text: "no load test yet, so no requests-per-second figure." },
  ], { x: 0.5, y: 3.35, w: 9, h: 1.45, fontSize: 13.5, paraSpaceAfter: 6 });
  s.addNotes("What can it handle? Each claim gets three tries and each user five per day. A code lives 72 hours and there are about 850 billion possible codes. Because all state, including the limits, is in Postgres, any number of server copies can run at once. Sign-in is limited to five failures per account and thirty per IP address in fifteen minutes. We have not run a load test, so we don't claim a requests-per-second figure.");
}

// ---------- Results ----------
pres.addSection({ title: "Results" });
{
  const s = content("51 automated tests, all passing", "Results");
  s.addChart(pres.charts.BAR, [{
    name: "Tests",
    labels: ["Claims and handover", "Matching", "Hashing and accounts", "Sign-in limit", "Privacy", "Text checks", "Demo data"],
    values: [19, 9, 6, 6, 5, 4, 2],
  }], {
    x: 0.5, y: 1.15, w: 5.8, h: 3.45, barDir: "bar", barGapWidthPct: 60,
    chartColors: [HEX.dk2], invertedColors: [HEX.dk2],
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: HEX.dk1, dataLabelFontSize: 12, dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.dk1, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
    valAxisHidden: true, valAxisMaxVal: 22, valAxisMinVal: 0,
    valGridLine: { style: "none" }, catGridLine: { style: "none" },
    showLegend: false, showTitle: true, title: "Automated tests by area", titleFontSize: 13, titleColor: HEX.dk1, titleFontFace: "+mn-lt",
  });
  caption(s, fig("The 51 tests in 7 files, grouped by what they check."), 0.5, 4.62, 5.8);
  card(s, 6.6, 1.25, 2.9, 3.55, "results-note");
  text(s, "51 / 51", { x: 6.8, y: 1.4, w: 2.5, h: 0.6, fontFace: "Cambria", fontSize: 30, bold: true, color: C.text2 });
  text(s, "pass in 12.6 s on PGlite, and on PostgreSQL 17", { x: 6.8, y: 2.0, w: 2.5, h: 0.55, fontSize: 12.5 });
  text(s, "Parallel tests", { x: 6.8, y: 2.65, w: 2.5, h: 0.35, fontFace: "Cambria", fontSize: 16, bold: true, color: C.text2 });
  text(s, "8 wrong answers at once, 6 uses of one code at once, 12 parallel passwords. With a lock removed, 5–6 of 6 guesses got through instead of 2.", { x: 6.8, y: 3.0, w: 2.5, h: 1.7, fontSize: 12, italic: true, color: C.text2 });
  source(s, "npx vitest run, 28 October 2026; ppclfmvrs README (concurrency checks on PostgreSQL 17).");
  s.addNotes("All 51 automated tests pass, in about 12 seconds. Most of them test claims and handover codes, because that is where fraud would happen. Several tests fire requests in parallel: eight wrong answers at once, six uses of one code at once, twelve wrong passwords at once. We ran them on a real PostgreSQL 17 server too, and as a check we removed each lock in turn: without the user lock, five or six of six parallel guesses got through instead of two. With the locks, the limits hold.");
}
{
  const s = pres.addSlide({ masterName: "STATEMENT_DARK", sectionTitle: "Results" });
  s.addText("Tested live on Supabase: every step held", { placeholder: "title" });
  const steps = [
    ["True match ranked first", "93%"], ["Wrong answer refused", "2 tries left"], ["Right answers", "Code + QR issued"],
    ["Desk enters the code", "Confirmed"], ["Same code again", "409 already used"], ["6th wrong password", "429 too many"],
  ];
  steps.forEach(([label, result], i) => {
    const x = 0.6 + (i % 2) * 4.5, y = 1.5 + Math.floor(i / 2) * 0.75;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 4.2, h: 0.6, rectRadius: 0.08, fill: { color: STATEMENT_BOX }, line: { color: STATEMENT_BOX }, objectName: `live-step-${i + 1}` });
    text(s, label, { x: x + 0.2, y, w: 2.3, h: 0.6, fontSize: 14.5, bold: true, color: C.background1, valign: "middle" });
    text(s, result, { x: x + 2.4, y, w: 1.65, h: 0.6, fontSize: 14.5, bold: true, color: C.accent1, valign: "middle", align: "right" });
  });
  text(s, [
    { text: "The same flow ran on the group’s production database through Supabase’s transaction pooler. " },
    { text: "Matching, verification, single-use codes and the sign-in limit all behaved as designed.", options: { bold: true, color: HEX.accent1 } },
  ], { x: 0.6, y: 3.9, w: 8.8, h: 1.0, fontSize: 15, color: C.background1 });
  s.addNotes("We also tested the whole flow live on Supabase, the production database. The true match ranked first at 93 percent. A wrong answer was refused with two tries left. The right answers gave a handover code. The desk confirmed it. Using the same code again was refused with 409, and the sixth wrong password was refused with 429. Every step behaved as designed.");
}
{
  const s = content("Errors we found and fixed", "Results");
  s.addTable([
    [head("Problem"), head("Fix")],
    ["Date fields could show a stale, prerendered date", "Read today’s date only in the browser"],
    ["Local SQLite file would lose data on serverless hosts", "Moved all data to PostgreSQL (Supabase)"],
    ["Parallel answers could pass the attempt limits", "Lock user, item and claim rows in one transaction"],
    ["A code used twice at once could work twice", "One conditional UPDATE uses the code"],
    ["Passwords could be guessed without limit", "Sign-in limit: 5 per account, 30 per IP"],
    ["Checking answers blocked other requests", "Async scrypt on Node’s thread pool"],
  ], { x: 0.5, y: 1.2, w: 9, colW: [4.6, 4.4], fontSize: 13, rowH: 0.48, ...tableStyle() });
  source(s, `Every fix is in the repository’s history: ${group.repoUrl}`);
  s.addNotes("We found and fixed real errors. Date fields could have baked in the date at build time, so the date is now read in the browser. Our first version kept data in a local SQLite file, which would be lost on serverless hosts, so we moved to Postgres. With many servers, parallel answers could have slipped past the limits, and a code could have been used twice, so we added row locks and a single conditional update. Sign-in had no rate limit, so we added one. And answer checking now uses async scrypt, so it doesn't block other requests.");
}

// ---------- Conclusion ----------
pres.addSection({ title: "Conclusion" });
{
  const s = content("Limitations and how we mitigate them", "Conclusion");
  s.addTable([
    [head("Limitation"), head("Mitigation")],
    ["No way to agree a time for finder-held items", "Public meeting place; next, in-app messages that block numbers"],
    ["Weak questions make false claims easier", "Example questions; 3 / 5 limits; desk checks the ID card"],
    ["Private details not encrypted at rest", "Own schema, off Supabase’s public API, owner-only; next, encrypt"],
    ["Desk handover trusts the officer", "Single-use codes; officer checks the owner’s ID card"],
    ["No sign-up; demo accounts seeded", "Treat as a test; connect to the university’s identity system"],
    ["Desk names and hours are examples", "Confirm with the Security Unit and Students’ Union"],
  ], { x: 0.5, y: 1.2, w: 9, colW: [3.6, 5.4], fontSize: 12.5, rowH: 0.5, ...tableStyle() });
  s.addNotes("Every limitation has a mitigation. Finder-held items have a public meeting place but no way to agree a time; next we add in-app messages that block phone numbers. Weak questions are countered by example questions, the attempt limits and the ID check. Private details are not yet encrypted at rest, but they sit in their own schema, off Supabase's public API, and only the owner sees them. Handover codes are single use. The demo accounts must be replaced with the university's identity system, and the desk details confirmed with the Security Unit.");
}
{
  const s = content("Installation and deployment", "Conclusion");
  card(s, 0.5, 1.2, 4.35, 3.55, "install-local");
  text(s, "Run it locally", { x: 0.75, y: 1.35, w: 3.9, h: 0.4, fontFace: "Cambria", fontSize: 17, bold: true, color: C.text2 });
  text(s, ["git clone the repository", "cd swep-g-15/ppclfmvrs", "npm install", "npm run dev", "npm test"].map((t, i, a) => ({ text: t, options: { breakLine: i < a.length - 1 } })), { x: 0.75, y: 1.85, w: 3.9, h: 2.0, fontFace: "Courier New", fontSize: 13, paraSpaceAfter: 6 });
  text(s, "No database server needed: PGlite runs locally.", { x: 0.75, y: 4.15, w: 3.9, h: 0.4, fontSize: 12, italic: true, color: C.text2 });
  card(s, 5.15, 1.2, 4.35, 3.55, "install-deploy", C.text2);
  text(s, "Put it online", { x: 5.4, y: 1.35, w: 3.9, h: 0.4, fontFace: "Cambria", fontSize: 17, bold: true, color: C.accent1 });
  text(s, [
    "Supabase: copy the Transaction pooler link (port 6543)",
    "Vercel: new project, Root Directory ppclfmvrs",
    "Set DATABASE_URL; DATABASE_SCHEMA is optional",
    "Deploy: tables are created automatically",
  ].map((t, i, a) => ({ text: `${i + 1}.  ${t}`, options: { breakLine: i < a.length - 1 } })), { x: 5.4, y: 1.85, w: 3.9, h: 2.7, fontSize: 13, paraSpaceAfter: 8, color: C.background1 });
  source(s, "Supabase docs, Connecting to your database (Supavisor transaction mode).");
  s.addNotes("To run it locally takes five commands: clone, open the folder, npm install, npm run dev, and npm test. No database server is needed. To put it online, copy Supabase's transaction pooler link on port 6543, create a Vercel project with the root directory set to ppclfmvrs, set DATABASE_URL, and deploy. The tables are created automatically.");
}
{
  const s = content("Implemented and open features", "Conclusion");
  const columns = [
    ["Implemented", ["Lost and found reports, private details", "Sensitive-text check", "Ranked automatic matching", "Hashed questions, 3 / 5 limits", "One-time code and QR", "Desk dashboard, QR scan, failed log", "Leak guard, sign-in limit, Postgres"], C.background2],
    ["Open", ["In-app messages for finder-held items", "Encrypt private details at rest", "Confirm desk details with Security Unit", "Sign-up via the university’s system", "SMS or email notifications", "Synonyms in matching (purse, wallet)", "More campus places than 26"], C.text2],
  ];
  columns.forEach(([h, items, fill], i) => {
    const x = 0.5 + i * 4.65, dark = fill === C.text2;
    card(s, x, 1.2, 4.35, 3.6, `features-${i + 1}`, fill);
    text(s, h, { x: x + 0.25, y: 1.35, w: 3.9, h: 0.4, fontFace: "Cambria", fontSize: 17, bold: true, color: dark ? C.accent1 : C.text2 });
    text(s, list(items, { color: dark ? HEX.lt1 : HEX.dk1 }), { x: x + 0.25, y: 1.85, w: 3.9, h: 2.85, fontSize: 13, paraSpaceAfter: 5 });
  });
  s.addNotes("On the left is everything implemented. On the right are the open items: in-app messages for finder-held items, encryption of private details at rest, confirming the real desk details with the Security Unit, sign-up through the university's identity system, SMS or email notifications, synonyms in matching, and more campus places.");
}
{
  const s = content("Our team and project links", "Conclusion");
  const rows = group.members.map((m) => [m.name, m.matric, m.department]);
  s.addTable([[head("Name"), head("Matric number"), head("Department")], ...rows], { x: 0.5, y: 1.2, w: 9, colW: [3.4, 2.3, 3.3], fontSize: 12, rowH: 0.42, ...tableStyle() });
  text(s, [
    { text: "Source code, guide and full history: ", options: { bold: true } },
    { text: group.repoUrl, options: { hyperlink: { url: group.repoUrl }, color: HEX.accent3 } },
    { text: "  (folder ppclfmvrs)" },
  ], { x: 0.5, y: 4.0, w: 9, h: 0.4, fontSize: 14 });
  s.addNotes("This is our team, and the link to the repository. It has the full source code in the ppclfmvrs folder, the installation and deployment guide, the tests, and the history of every change and fix.");
}
{
  const s = content("References", "Conclusion");
  const refs = [
    "Dice, L. R. (1945). Measures of the amount of ecologic association between species. Ecology, 26(3), 297–302.",
    "Federal Republic of Nigeria (2023). Nigeria Data Protection Act, 2023.",
    "ISO/IEC 18004:2015. QR Code bar code symbology specification.",
    "Node.js. Crypto module documentation (scrypt, timingSafeEqual).",
    "OpenStreetMap contributors (2026). OpenStreetMap. Open Database Licence.",
    "OWASP Foundation (2023). API Security Top 10, API3:2023 Broken Object Property Level Authorization.",
    "Percival, C. and Josefsson, S. (2016). The scrypt Password-Based Key Derivation Function, RFC 7914. IETF.",
    "PostgreSQL Global Development Group (2024). PostgreSQL 17 Documentation: Explicit Locking.",
    "Sinnott, R. W. (1984). Virtues of the haversine. Sky and Telescope, 68(2), 159.",
    "Supabase. Connecting to your database (Supavisor transaction mode).",
  ];
  const half = Math.ceil(refs.length / 2);
  [refs.slice(0, half), refs.slice(half)].forEach((items, c) => {
    text(s, items.map((r, i) => ({ text: r, options: { breakLine: i < items.length - 1 } })), { x: 0.5 + c * 4.6, y: 1.2, w: 4.3, h: 3.75, fontSize: 11.5, paraSpaceAfter: 8 });
  });
  s.addNotes("These are the sources we cited. The full reference list is in our report.");
}
{
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Conclusion" });
  s.addText("Thank you", { placeholder: "title" });
  s.addText([
    { text: "PPCLFMVRS matches lost and found items on campus and returns them only to owners who prove it, without posting anyone’s details.", options: { breakLine: true } },
    { text: "We welcome your questions.", options: { bold: true, color: HEX.accent1 } },
  ], { placeholder: "body" });
  s.addNotes("To conclude: PPCLFMVRS matches lost and found items on campus and returns them only to owners who prove it, without posting anyone's details. Thank you for listening. We welcome your questions, and we can show the app live.");
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
  fs.rmSync(profile, { recursive: true, force: true });
  console.log("wrote", OUT.replace(/\.pptx$/, ".pdf"));
});
