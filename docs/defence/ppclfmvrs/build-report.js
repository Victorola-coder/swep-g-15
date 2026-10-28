// Builds the PPCLFMVRS SWEP defence report as a Word document (and a PDF beside it).
// Usage: node build-report.js [output.docx]   (default: ~/Desktop)
// Group details come from ../group.json and ../members.json (PROJECT=ppclfmvrs).
// Run `node build-figures.js` first to draw the diagrams and crop the screenshots.
process.env.PROJECT ??= "ppclfmvrs";
const fs = require("fs");
const os = require("os");
const path = require("path");
const {
  AlignmentType, BorderStyle, Document, Footer, HeadingLevel, ImageRun, LevelFormat, Packer,
  PageNumber, Paragraph, ShadingType, Table, TableCell, TableLayoutType, TableOfContents, TableRow, TextRun, WidthType,
} = require(path.join(__dirname, "..", "node_modules", "docx"));
const group = require("../group");

const OUT = process.argv[2] ?? path.join(os.homedir(), "Desktop", "PPCLFMVRS Project Report - SWEP Group 15.docx");
const ASSETS = path.join(__dirname, "assets");
const SHOTS = path.join(ASSETS, "shots");
const FONT = "Times New Roman";
const CONTENT_WIDTH = 9026; // A4 width 11906 minus two 1440 margins, in DXA
const BLANK = "______________________________";
const HEAD_FILL = "E4DCEC";

// ---------- inline text: **bold**, *italic*, `code` ----------
function runs(text, base = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)/g;
  let last = 0;
  for (const m of text.matchAll(re)) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    const t = m[0];
    if (t.startsWith("**")) out.push(new TextRun({ text: t.slice(2, -2), bold: true, ...base }));
    else if (t.startsWith("`")) out.push(new TextRun({ text: t.slice(1, -1), font: "Courier New", size: 20, ...base }));
    else out.push(new TextRun({ text: t.slice(1, -1), italics: true, ...base }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...base }));
  return out;
}

const P = (text, opts = {}) => new Paragraph({ children: runs(text), spacing: { after: 160 }, alignment: AlignmentType.JUSTIFIED, ...opts });
const center = (text, opts = {}) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 }, ...opts, children: runs(text, opts.run) });
const H2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
// Sections start on a new page only where a break helps the reader; short ones run on.
const lead = (text) => P(text, { keepNext: true });
const S = (text, newPage = true) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: newPage, children: [new TextRun(text)] });

let listInstance = 0;
const bullets = (items) => items.map((t) => new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: runs(t) }));
const numbered = (items) => {
  const instance = ++listInstance;
  return items.map((t) => new Paragraph({ numbering: { reference: "numbers", level: 0, instance }, spacing: { after: 80 }, children: runs(t) }));
};

const border = { style: BorderStyle.SINGLE, size: 4, color: "808080" };
const borders = { top: border, bottom: border, left: border, right: border };
function table(header, rows, widths, { size = 22 } = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const scaled = widths.map((w) => Math.round((w / total) * CONTENT_WIDTH));
  scaled[scaled.length - 1] += CONTENT_WIDTH - scaled.reduce((a, b) => a + b, 0);
  // keepNext on every row but the last keeps a table on one page, together with its caption.
  const cell = (text, i, head) =>
    new TableCell({
      borders,
      width: { size: scaled[i], type: WidthType.DXA },
      shading: head ? { fill: HEAD_FILL, type: ShadingType.CLEAR, color: "auto" } : undefined,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [new Paragraph({ keepNext: true, spacing: { after: 0, line: 276 }, children: runs(text, { size, bold: head || undefined }) })],
    });
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    columnWidths: scaled,
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: header.map((h, i) => cell(h, i, true)) }),
      ...rows.map((r) => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, i, false)) })),
    ],
  });
}
const caption = (text) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 240 }, children: runs(text, { italics: true, size: 22 }) });

function figure(file, w, h, text) {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      keepNext: true,
      spacing: { before: 120, after: 0 },
      children: [new ImageRun({ type: "png", data: fs.readFileSync(file), transformation: { width: w, height: h }, altText: { title: text, description: text, name: path.basename(file) } })],
    }),
    caption(text),
  ];
}
// Screenshots are cropped to the app's content column: 1032 × 900, so width / height = 1.147.
const shot = (name, text, w = 540) => figure(path.join(SHOTS, name), w, Math.round(w / 1.1467), text);

// ---------- content ----------
const memberRows = group.hasMembers
  ? group.members.map((m, i) => [String(i + 1), m.name, m.matric, m.department])
  : [1, 2, 3, 4, 5].map((n) => [String(n), "", "", ""]);

const titlePage = [
  new Paragraph({ spacing: { before: 400 } }),
  center("**PRIVACY-PRESERVING CAMPUS LOST-AND-FOUND MATCHING AND VERIFIED RECOVERY SYSTEM (PPCLFMVRS)**", { run: { size: 32 } }),
  new Paragraph({ spacing: { after: 200 } }),
  center("A report submitted in fulfilment of the requirements of the Students’ Work Experience Programme (SWEP)"),
  new Paragraph({ spacing: { after: 200 } }),
  center(`**${group.groupName}**`, { run: { size: 28 } }),
  table(["S/N", "Name", "Matriculation number", "Department"], memberRows, [7, 38, 25, 30]),
  new Paragraph({ spacing: { after: 240 } }),
  ...(group.faculty ? [center(group.faculty)] : []),
  center(group.university),
  center(`Source code: ${group.repoUrl}`),
  new Paragraph({ spacing: { after: 240 } }),
  center(`Supervisor: ${group.supervisor || BLANK}`),
  center(`Date: ${group.defenceDate || BLANK}`),
];

const contents = [
  S("Table of Contents"),
  new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-2" }),
  P("*If the table above is empty, right-click it in Word and choose Update Field.*"),
];

const s1 = [
  S("1. Introduction"),
  P("PPCLFMVRS is a web app that helps students and staff of Obafemi Awolowo University (OAU) get lost items back without exposing anyone’s private details. In the app it is called **OAU Lost & Found**. Owners report what they lost and finders report what they found. The system matches the two kinds of report automatically and shows the owner a ranked list of likely matches. An item is handed over only after the claimant answers questions that only the real owner can answer, and then only with a one-time code, at a collection desk or a public meeting place."),
  P("Today, a finder usually posts a photo and a full description in WhatsApp groups. The photo can show the owner’s name, lock screen or card numbers, and anyone who sees it can then describe the item well enough to claim it. PPCLFMVRS keeps those details private and makes the claimant prove ownership."),
  H2("1.1 Aim"),
  P("To design and build a web app that matches lost and found items on the OAU campus and returns them only to verified owners, without exposing private details."),
  H2("1.2 Objectives"),
  ...numbered([
    "Let owners report lost items with public details, plus private identifying details that only they can see.",
    "Let finders report found items with a generic description and 2–3 verification questions.",
    "Match lost and found reports automatically and rank the candidates, with a clear score breakdown.",
    "Verify ownership with hashed answers and strict attempt limits.",
    "Hand items over with a one-time code and QR code, confirmed by a desk officer or the finder.",
    "Give desk officers a dashboard for handovers, items at the desk and failed claims.",
    "Enforce privacy in code, so that private details, answers and contact details never leave the server.",
  ]),
];

const s2 = [
  S("2. The Problem and the App’s Implications for the Campus"),
  H2("2.1 The problem"),
  ...bullets([
    "**Private details are posted publicly.** Photos of lost phones, ID cards and wallets show names, lock screens, and card and account numbers to everyone in the group.",
    "**False claims are easy.** Anyone who has seen the post can describe the item and say it is theirs.",
    "**Matching is left to chance.** An owner must scroll through many groups and notice boards and hope to see the post. Nobody compares lost reports with found reports.",
    "**There is no record.** Nobody records who collected an item or whether they proved it was theirs.",
  ]),
  H2("2.2 Implications for the campus"),
  table(
    ["Area", "What changes with PPCLFMVRS"],
    [
      ["Students and staff who lose items", "One place to report; likely matches are found for them, ranked and explained"],
      ["Finders", "Can hand an item in without posting photos or sharing their phone number"],
      ["Privacy", "Names, phone numbers, card numbers and lock screens are no longer posted to large groups, in line with the aims of the Nigeria Data Protection Act 2023"],
      ["Fraud", "A false claimant must answer questions they cannot see, with 3 tries per claim and 5 per day"],
      ["Security Unit and SUB desks", "A dashboard of items held, handovers waiting and failed claims, and a record of every handover"],
      ["Personal safety", "Items a finder keeps are handed over at a public campus place, not at someone’s room"],
    ],
    [32, 68],
  ),
  caption("Table 2.1: Implications of PPCLFMVRS for the OAU campus"),
  P("**Risks to manage.** Weak questions such as “What colour is it?” make false claims easier, so the form shows examples of good questions. Desk officers must still check the owner’s ID card. Private details are stored as plain text on the server, so the database must be protected. The demo accounts must be removed, and the desk names and hours confirmed with the Security Unit, before real use."),
];

const s3 = [
  S("3. Distinct Features and Specifications"),
  H2("3.1 Distinct features"),
  ...bullets([
    "**Lost and found reports** with category, colour, place (26 OAU places), date and, for lost items, a time window.",
    "**Private identifying details** (last digits of the serial or IMEI number, lock-screen picture, contents, unique marks) that only the owner can see.",
    "**Sensitive-text check:** public text containing a phone number, email, matriculation number or long number is refused.",
    "**Automatic matching:** ranked candidates with a score for each factor.",
    "**Verification questions** written by the finder; the answers are stored only as salted hashes.",
    "**One-time handover code and QR code**, valid for 72 hours.",
    "**Desk dashboard:** statistics, handover by code or camera QR scan, items held, and the failed-claim log.",
    "**In-app notifications** for both owner and finder.",
    "**No photo uploads**, by design.",
    "**Accessible design:** labelled fields, keyboard use, light and dark themes, and the Atkinson Hyperlegible font.",
  ]),
  H2("3.2 Specifications"),
  table(
    ["Item", "Specification"],
    [
      ["Type", "Web app; works in any modern browser on phones and computers"],
      ["Roles", "Member (student or staff), desk officer, admin"],
      ["Catalogue", "9 categories, 16 colours, 26 OAU places with OpenStreetMap coordinates, 2 collection desks"],
      ["Matching", "Category must match; colour 25%, place 25%, date 20%, description words 30%; shown from 35%"],
      ["Verification", "2–3 questions; scrypt hashes; 3 tries per claim and 5 per user in any 24 hours"],
      ["Handover code", "8 characters (about 40 bits), single use, valid 72 hours, stored as a SHA-256 hash"],
      ["Sign-in", "scrypt password hashes; at most 5 failures per account and 30 per IP address in 15 minutes"],
      ["Storage", "PostgreSQL (Supabase in production, PGlite locally): 11 tables in one schema"],
      ["Interface", "17 API routes; every response passes allowlist serializers and a leak guard"],
      ["Automated tests", "51, all passing"],
    ],
    [25, 75],
  ),
  caption("Table 3.1: PPCLFMVRS specifications"),
];

const s4 = [
  S("4. Comparison with Existing Ways", false),
  lead("Lost items on campus are usually handled in one of three ways. The table compares them with PPCLFMVRS by what each one does by design. We did not measure them, so the table gives no figures."),
  table(
    ["Feature", "WhatsApp groups and notice boards", "A desk logbook (e.g. Security Unit)", "Generic lost-and-found apps", "PPCLFMVRS"],
    [
      ["Reaches many people quickly", "Yes", "No: only people who visit the desk", "Only if many people use it", "Yes: a campus board of found items"],
      ["Keeps the owner’s details private", "No: photos and details are public", "Yes", "Varies by app", "Yes: no photos; private details owner-only"],
      ["Checks ownership", "Informally, by chat", "Staff judgement", "Varies by app", "Hashed questions with attempt limits"],
      ["Compares lost with found", "By chance", "By hand", "Varies by app", "Automatically, ranked with scores"],
      ["Stops repeated guessing", "No", "Depends on staff", "Varies by app", "3 per claim, 5 per day; failures logged"],
      ["Records the handover", "No", "Yes, a written entry", "Varies by app", "Yes: one-time code, logged"],
      ["Built for OAU places and desks", "Yes", "Yes", "No", "Yes"],
    ],
    [22, 19, 19, 17, 23],
    { size: 20 },
  ),
  caption("Table 4.1: PPCLFMVRS compared with existing ways (by design, not measured)"),
];

const s5 = [
  S("5. Tools and Technology Stack", false),
  table(
    ["Tool", "What we used it for"],
    [
      ["Next.js 16 and React 19", "The web app: its pages and its 17 API routes"],
      ["TypeScript", "The programming language, with type checks that catch mistakes early"],
      ["Tailwind CSS 4", "Styling, including light and dark themes"],
      ["Node.js crypto", "scrypt for passwords and answers; SHA-256 for session tokens and codes; timingSafeEqual for comparisons"],
      ["PostgreSQL (Supabase)", "Stores all data in production"],
      ["postgres.js", "Connects to Postgres through Supabase’s transaction pooler"],
      ["PGlite", "Postgres inside Node.js, for local use and tests with no database server"],
      ["qrcode", "Draws the QR code for the handover code"],
      ["OpenStreetMap", "Coordinates of the 26 campus places, used for distances"],
      ["Vitest", "Automated tests"],
      ["ESLint", "Checks code quality"],
      ["Git and GitHub", "Version control and sharing the code"],
      ["Vercel", "Hosting the app online"],
    ],
    [30, 70],
  ),
  caption("Table 5.1: Tools and technology used"),
  P("**Tools needed to run it:** a computer with Node.js 22 or newer, npm (comes with Node.js), Git, and a web browser. For deployment: a Supabase account (free) and a Vercel account (free). Scanning a QR code at the desk needs a browser with the BarcodeDetector API, such as Chrome on Android; typing the code works in any browser."),
];

const s6 = [
  S("6. System Architecture"),
  ...figure(path.join(ASSETS, "architecture.png"), 560, 430, "Figure 6.1: System architecture"),
  ...bullets([
    "**In the browser:** the owner, finder and desk pages ask the server for data in JSON and show it. Nothing private is ever sent to them.",
    "**On the server:** the API routes check the session and that each request is JSON, then call the matching, claims and report code. Every response is built by allowlist serializers and checked by the leak guard before it leaves.",
    "**In the database:** all data, including the attempt counters and sign-in failures, lives in Postgres, so several servers can run at once. Locally the same SQL runs on PGlite.",
  ]),
];

const s7 = [
  S("7. How Matching, Verification and Privacy Work"),
  H2("7.1 Match scoring"),
  lead("Every new report is scored against the open reports of the other kind, using public fields only, so private details cannot leak through scores."),
  table(
    ["Factor", "Weight", "How it is scored"],
    [
      ["Category", "Gate", "Must be the same, or the pair is never scored"],
      ["Colour", "25%", "1 for the same colour; 0.5 for colours people confuse (black/grey, silver/white, blue/navy…); 0.3 if either is multicoloured"],
      ["Place", "25%", "1 for the same place, falling to 0 at 900 m; distance from the places’ coordinates (haversine formula)"],
      ["Date", "20%", "1 if found on the day of the loss or the next day, then falling to 0 over two weeks; 0.5 if found one day before"],
      ["Keywords", "30%", "Dice overlap of the description words, 2·|A∩B| / (|A|+|B|), after removing common words and plurals"],
    ],
    [16, 12, 72],
  ),
  caption("Table 7.1: Match factors (Dice, 1945; Sinnott, 1984)"),
  P("Candidates scoring **35% or more** are shown to the owner, best first; a new pair scoring **50% or more** notifies both sides. In the demo, the black phone found in the library scores 100% for colour, place and date and 77% for words: 0.25 + 0.25 + 0.20 + 0.30 × 0.77 = **93%**. A white phone found at Moremi Hall scores only 37%."),
  H2("7.2 Hashed answers and attempt limits"),
  ...bullets([
    "Each answer is normalised (trimmed, lower-cased, spaces collapsed), so “Blue  STAR” matches “blue star”, then hashed with **scrypt** and a random salt (Percival and Josefsson, 2016). A slow hash matters because answers are short and guessable.",
    "When the owner answers, every answer is checked, even after a wrong one, and hashes are compared with `timingSafeEqual`, so timing does not reveal which answer was wrong.",
    "**3 wrong tries per claim** lock the claim; **5 tries per user in any 24 hours** stop someone fishing across many items. Every attempt is logged, and failures appear on the desk dashboard.",
    "An owner can only claim an item the matcher shows for one of their own lost reports, and a finder cannot claim their own item.",
    "The limits are enforced inside one database transaction that locks the user, item and claim rows, so parallel requests cannot get past them (PostgreSQL, 2024).",
  ]),
  H2("7.3 Single-use handover codes"),
  P("When the answers are right, the server issues an 8-character code from 31 characters with no look-alikes (no 0/O or 1/I/L): about 8.5 × 10¹¹ possible codes, or 40 bits. It is valid for 72 hours, shown once with its QR code (ISO/IEC 18004), and stored only as a SHA-256 hash. The desk officer of the holding desk, or the finder for items they keep, enters or scans it. One SQL statement marks it used only if it is unused, not replaced and not expired, so even two requests at the same moment cannot both succeed: the second gets **409 already used**. Asking for a new code revokes the old one."),
  H2("7.4 Allowlist serializers and the leak guard"),
  table(
    ["Data", "Who can see it"],
    [
      ["Public fields (category, colour, place, date, generic description)", "Everyone (found board); the owner (candidates)"],
      ["Owner’s private details", "The owner only"],
      ["Verification answers", "Nobody: salted hashes only"],
      ["Handover code", "The owner, once"],
      ["Names, phone numbers, emails, user ids", "Nobody else, ever"],
      ["Failed attempts", "Desk staff and admins, with the claimant’s matric number"],
    ],
    [55, 45],
  ),
  caption("Table 7.2: Who can see what"),
  P("Every response is built by **allowlist serializers** that copy named fields only, so a new database column cannot leak by accident. Then **assertNoLeaks** walks the whole response and throws an error if a forbidden field (answers, hashes, salts, phone, email, owner or finder ids) appears at any depth. This guards against what OWASP calls broken object property level authorization (OWASP, 2023). Someone else’s report returns 404, not 403, so ids cannot be probed."),
  H2("7.5 No photos, by design"),
  P("Photos are the main way details leak in WhatsApp groups: lock screens, names on ID cards and unique marks all show up in them. A photo also gives a false claimant everything needed to answer the questions. So the app takes no photos at all: a generic text description plus private questions is safer."),
];

const s8 = [
  S("8. How to Use the App", false),
  P("The screenshots below come from the app running on a computer with its demo data. The demo accounts and answers are listed in the project README."),
  H2("8.1 Owner: report a lost item"),
  ...numbered([
    "Sign in with your matriculation or staff number and press **Report lost**.",
    "Fill in the public details: category, colour, brand, a short description, where you last had it, the date and the time window.",
    "Add private details only you will see, such as the last digits of the IMEI or a sticker on the back, then press **Report lost item**.",
  ]),
  ...shot("lost-new-private.png", "Figure 8.1: The lost-item form. Private identifying details are never shown to anyone else."),
  H2("8.2 Owner: see ranked candidates"),
  lead("Open the report from **My reports**. Possible matches are listed best first, with a bar for each factor."),
  ...shot("candidates.png", "Figure 8.2: Ranked candidates. The true match scores 93%; a white phone from Moremi Hall scores 37%."),
  H2("8.3 Owner: claim and answer the questions"),
  lead("Press **This is mine, claim it** and answer the finder’s questions. A wrong answer uses one of the 3 tries on the claim."),
  ...shot("claim-wrong.png", "Figure 8.3: A wrong answer is refused, with 2 tries left on the claim."),
  H2("8.4 Owner: receive the handover code and QR"),
  lead("When all answers are right, the page shows where to collect the item, the one-time code and its QR code. Save them: the code is shown only once."),
  ...shot("claim-verified.png", "Figure 8.4: A verified claim with the meeting place, the one-time code and its QR code."),
  H2("8.5 Finder: report a found item with questions"),
  lead("Press **Report found**, give a generic description, and write 2–3 questions about things only the owner would know. If the description contains a phone number, email or identifying number, the form warns you and the server refuses it."),
  ...shot("found-warning.png", "Figure 8.5: The form flags an email address in the public description."),
  lead("Then choose where the item is: handed to a collection desk (recommended), or kept by you and handed over at a public place."),
  ...shot("found-new.png", "Figure 8.6: Verification questions and where the item is kept."),
  H2("8.6 Desk: the dashboard"),
  lead("Desk officers sign in and open **Desk**. The dashboard shows statistics, the handover box, items waiting for collection and items held at the desk."),
  ...shot("desk-dashboard.png", "Figure 8.7: The Security Unit desk dashboard, with one verified owner waiting."),
  lead("Further down, the failed claim attempts are listed with the claimant’s matric number, so repeated fishing can be followed up."),
  ...shot("desk-failed-log.png", "Figure 8.8: Items held at the desk and the failed-claim log."),
  H2("8.7 Desk: confirm a handover"),
  lead("Check the owner’s ID card, type their code or press **Scan QR code**, then press **Confirm handover**. The item is marked recovered."),
  ...shot("desk-confirmed.png", "Figure 8.9: The handover is confirmed and the code can no longer be used."),
];

const s9 = [
  S("9. App Life Cycle", false),
  H2("9.1 Development life cycle"),
  lead("We built PPCLFMVRS in short steps, testing each one before the next. Each stage below is a commit in the repository."),
  table(
    ["Stage", "Date (2026)", "What we did"],
    [
      ["1. Planning and design", "By 23 Oct", "Set the aim, the privacy rules and the data model"],
      ["2. Set-up", "23 Oct", "Created the app with Next.js, Tailwind CSS and Vitest"],
      ["3. Matching engine", "23 Oct", "Campus places, match scoring and the sensitive-text check, with tests"],
      ["4. Data and security", "23 Oct", "Database, password and answer hashing, sessions and attempt limits"],
      ["5. Claims and handover", "23 Oct", "Reports, matching on creation, serializers, verification and one-time codes"],
      ["6. API and interface", "23 Oct", "17 API routes; pages for owners, finders and desks; UI polish"],
      ["7. Documentation", "23 Oct", "Wrote the README: running, demo accounts, matching, privacy and API"],
      ["8. Production readiness", "27 Oct", "Moved to Postgres with row locks; added the sign-in rate limit"],
      ["9. Live test", "After 27 Oct", "Connected to the group’s Supabase database and tested the full flow"],
      ["10. Deployment", "Next", "Host on Vercel with Root Directory ppclfmvrs"],
      ["11. Maintenance", "Ongoing", "Pilot with the desks, feedback, open features"],
    ],
    [26, 16, 58],
  ),
  caption("Table 9.1: Development life cycle"),
  H2("9.2 Life of an item and a claim"),
  ...figure(path.join(ASSETS, "lifecycle.png"), 540, 258, "Figure 9.1: The life of an item and a claim"),
  P("A report is **open** until it is matched and claimed. A claim is **open** while the owner answers, **verified** when all answers are right, and **completed** when the code is used; the found item then becomes **recovered**. Three wrong tries **lock** the claim. When one owner is verified, other open claims on the item are **cancelled**. Reports can also be **closed**."),
];

const s10 = [
  S("10. What the System Can Handle", false),
  table(
    ["Measure", "Value"],
    [
      ["Campus places", "26, with distances from OpenStreetMap coordinates"],
      ["Questions per found item", "2–3; a question up to 150 characters, an answer up to 100"],
      ["Public description", "Up to 300 characters (lost) or 200 (found)"],
      ["Answer attempts", "3 per claim; 5 per user in any 24 hours"],
      ["Failed sign-ins", "5 per account and 30 per IP address in 15 minutes, then 429"],
      ["Handover code", "Single use, valid 72 hours; about 8.5 × 10¹¹ possible codes"],
      ["Session length", "7 days"],
      ["Database connections", "At most 5 per server instance, through Supabase’s pooler"],
      ["Server instances", "Any number at once (as on Vercel): all state, including limits, is in Postgres"],
      ["Requests per second", "Not measured: we did not run a load test"],
    ],
    [35, 65],
  ),
  caption("Table 10.1: Capacity and limits of the system, from the code"),
];

const s11 = [
  S("11. Tests Done and Results"),
  H2("11.1 Automated tests"),
  table(
    ["Area", "Tests", "What they check"],
    [
      ["Matching", "9", "The true pair ranks first; a different category never matches; colour, place, date and word factors; the 35% threshold"],
      ["Text checks", "4", "Answer normalisation; word tokens; phone numbers, emails, matric and long numbers are flagged"],
      ["Hashing and accounts", "6", "Salted answer hashes; scrypt passwords; sign-in ignoring case; session expiry; code format; attempt rules"],
      ["Claims and handover", "19", "Verification; 3 and 5 attempt limits; parallel answers; two owners racing; single-use codes, also in parallel; expiry; replaced codes; desk permissions; hashes only"],
      ["Privacy", "5", "Allowlists; no private details, hashes, codes, contacts or ids in any view; the guard catches forbidden keys at any depth"],
      ["Sign-in limit", "6", "401 then 429; the lock lifts after 15 minutes; other accounts unaffected; reset on success; 30 per IP; at most 5 of 12 parallel guesses checked"],
      ["Demo data", "2", "The true match ranks first; data is seeded once when servers start together"],
    ],
    [20, 9, 71],
  ),
  caption("Table 11.1: Automated tests (51 in 7 files, all passing)"),
  P("The 51 tests ran in 12.6 seconds on PGlite (`npx vitest run`). They also pass on a real PostgreSQL 17 server. Type checking, code linting and a production build are clean."),
  H2("11.2 Concurrency tests and lock-removal checks"),
  P("The tests send requests in parallel: 8 wrong answers at once on one claim, 6 at once across 6 claims, two owners verifying the same item, 6 uses of one code at once, a new code racing a handover, 12 wrong passwords at once, and 3 servers seeding at once. Because PGlite runs one query at a time, these tests were also run on a real PostgreSQL 17 server. As a check, each lock was removed in turn, and the matching test then failed: without the user lock, 5–6 of 6 parallel guesses got through instead of 2."),
  H2("11.3 Live test on Supabase"),
  table(
    ["Step", "Result"],
    [
      ["Owner opens the lost black phone", "The true match is ranked first at 93%"],
      ["Owner sends a wrong answer", "Refused; 2 attempts left"],
      ["Owner sends the right answers", "Verified; a one-time handover code and QR code are issued"],
      ["Desk officer enters the code", "Handover confirmed; the item is recovered"],
      ["The same code is entered again", "Refused with 409 (already used)"],
      ["A 6th wrong password in 15 minutes", "Refused with 429 (too many requests)"],
    ],
    [45, 55],
  ),
  caption("Table 11.2: The full flow tested live on the Supabase database"),
];

const s12 = [
  S("12. Errors Found and Fixes", false),
  table(
    ["Problem", "Cause", "Fix"],
    [
      ["Date fields could show a stale date", "The default date was read with new Date() while the page was prerendered", "Read today’s date only in the browser (a useToday hook)"],
      ["Data would be lost on serverless hosts", "The first version kept data in a local SQLite file", "Moved all data to PostgreSQL (Supabase), with PGlite locally"],
      ["Parallel answers could pass the attempt limits", "With many servers, two requests could read the same count", "Lock the user, item and claim rows in one transaction"],
      ["A code used twice at once could work twice", "Check and use were separate steps", "One conditional UPDATE marks the code used"],
      ["Demo data could be seeded twice", "Several servers could start together", "Set-up and seeding hold a Postgres advisory lock"],
      ["Passwords could be guessed without limit", "Sign-in had no rate limit", "5 failures per account and 30 per IP in 15 minutes, then 429"],
      ["Checking answers blocked other requests", "scrypt ran synchronously", "Use async scrypt, which runs on Node’s thread pool"],
    ],
    [32, 34, 34],
  ),
  caption("Table 12.1: Errors found and how we fixed them. Each fix is in the repository’s history."),
];

const s13 = [
  S("13. Installation and Deployment Guide"),
  H2("13.1 Run it on your computer"),
  ...numbered([
    "Install Node.js 22 or newer and Git.",
    `Download the code: \`git clone ${group.repoUrl}\``,
    "Open the app folder: `cd swep-g-15/ppclfmvrs`",
    "Install the packages: `npm install`",
    "Start the app: `npm run dev`, then open http://localhost:3000",
    "Run the tests: `npm test`",
  ]),
  P("No database server is needed locally: without `DATABASE_URL`, the app uses PGlite in its `data/pglite` folder and fills it with the demo accounts and reports."),
  H2("13.2 Put it online"),
  ...numbered([
    "Create a free Supabase project. In **Project Settings → Database → Connection string**, copy the **Transaction pooler** link (port 6543) and add your database password. It looks like `postgresql://postgres.<project-ref>:<password>@aws-0-<region>.pooler.supabase.com:6543/postgres`.",
    "Create a Vercel project from the GitHub repository, with **Root Directory** set to `ppclfmvrs`.",
    "In Vercel, add the settings below, then deploy.",
  ]),
  table(
    ["Setting", "Value"],
    [
      ["DATABASE_URL", "The Supabase Transaction pooler link (required)"],
      ["DATABASE_SCHEMA", "Optional; defaults to ppclfmvrs"],
    ],
    [35, 65],
  ),
  caption("Table 13.1: Production settings"),
  P("There are no other secrets to set. On first start the app creates its tables and seeds the demo data, under a lock so that several servers do not collide. The group’s three apps can share one Supabase database, each in its own schema. Do not add `ppclfmvrs` to Supabase’s exposed schemas. The demo accounts are seeded online too, so treat a deployment as a test until they are removed."),
];

const s14 = [
  S("14. Implemented and Open Features"),
  table(
    ["Feature", "Status"],
    [
      ["Lost and found reports with private details", "Implemented"],
      ["Sensitive-text check on public text", "Implemented"],
      ["Automatic matching with ranked candidates", "Implemented"],
      ["Hashed verification questions with attempt limits", "Implemented"],
      ["One-time handover code and QR code", "Implemented"],
      ["Desk dashboard with QR scan and failed-claim log", "Implemented"],
      ["In-app notifications", "Implemented"],
      ["Allowlist serializers and leak guard", "Implemented"],
      ["Sign-in rate limit", "Implemented"],
      ["PostgreSQL storage for serverless hosting", "Implemented"],
      ["In-app messages to agree a time for finder-held items", "Open"],
      ["Encryption of private details at rest", "Open"],
      ["Confirm real desk names, places and hours with the Security Unit", "Open"],
      ["Sign-up through the university’s identity system", "Open: demo accounts for now"],
      ["SMS or email notifications", "Open"],
      ["Synonyms in matching (e.g. purse and wallet)", "Open"],
      ["More campus places than the current 26", "Open"],
    ],
    [65, 35],
  ),
  caption("Table 14.1: Implemented features and open improvements"),
];

const s15 = [
  S("15. Conclusion"),
  P("PPCLFMVRS shows that lost items can be matched and returned on campus without posting anyone’s details. All seven objectives were met: reports are matched automatically, ownership is proved with hashed answers under strict limits, items are handed over with single-use codes, and privacy is enforced in code. 51 automated tests and a live test on Supabase confirm it."),
  H2("15.1 Limitations and mitigation"),
  table(
    ["Limitation", "Mitigation"],
    [
      ["No way to agree a meeting time for finder-held items without sharing contacts", "Meet at a public place chosen in the app; next, in-app messages that block phone numbers and emails"],
      ["Weak questions make false claims easier", "The form shows good examples; 3 and 5 attempt limits; the desk checks the ID card"],
      ["Private details are not encrypted at rest", "Kept in their own schema, off Supabase’s public API, and sent only to the owner; next, encrypt them"],
      ["A desk handover trusts the desk officer", "Codes are single use and the officer is told to check the owner’s ID card"],
      ["No sign-up; demo accounts are seeded", "Treat a deployment as a test; connect to the university’s identity system"],
      ["Desk names and hours are examples", "Confirm them with the Security Unit and the Students’ Union"],
      ["QR scanning needs a supported browser", "The code can always be typed in"],
    ],
    [45, 55],
  ),
  caption("Table 15.1: Limitations and how each is mitigated"),
  H2("15.2 Recommendations"),
  ...numbered([
    "Confirm the desk details and run a pilot with the Security Unit and the SUB information desk.",
    "Add in-app messages and encryption at rest before any real data is entered.",
    "Connect sign-in to the university’s identity system and remove the demo accounts.",
  ]),
];

const s16 = [
  S("16. Project Links and Group Members"),
  P(`**Source code and documentation:** ${group.repoUrl}`),
  P("The repository contains the full code, the installation and deployment guide, the list of tests, and the history of every change and fix. The app is in the `ppclfmvrs` folder."),
  table(["S/N", "Name", "Matriculation number", "Department"], memberRows, [7, 38, 25, 30]),
  caption("Table 16.1: Group members"),
];

const references = [
  S("References"),
  ...[
    "Braille Institute. *Atkinson Hyperlegible Next typeface*. https://www.brailleinstitute.org/freefont/",
    "Dice, L. R. (1945). Measures of the amount of ecologic association between species. *Ecology, 26*(3), 297–302.",
    "Federal Republic of Nigeria. (2023). *Nigeria Data Protection Act, 2023*.",
    "International Organization for Standardization. (2015). *ISO/IEC 18004:2015 Information technology: QR Code bar code symbology specification*.",
    "Node.js. *Crypto module documentation* (scrypt, timingSafeEqual). https://nodejs.org/api/crypto.html",
    "OpenStreetMap contributors. (2026). *OpenStreetMap* [Data set]. https://www.openstreetmap.org/copyright",
    "OWASP Foundation. (2023). *API3:2023 Broken Object Property Level Authorization*. OWASP API Security Top 10. https://owasp.org/API-Security/",
    "OWASP Foundation. *Authentication Cheat Sheet* and *Password Storage Cheat Sheet*. https://cheatsheetseries.owasp.org",
    "Percival, C., & Josefsson, S. (2016). *The scrypt Password-Based Key Derivation Function* (RFC 7914). IETF. https://www.rfc-editor.org/rfc/rfc7914",
    "PostgreSQL Global Development Group. (2024). *PostgreSQL 17 Documentation: Explicit Locking*. https://www.postgresql.org/docs/17/explicit-locking.html",
    "Sinnott, R. W. (1984). Virtues of the haversine. *Sky and Telescope, 68*(2), 159.",
    "Supabase. *Connecting to your database* (Supavisor transaction mode). https://supabase.com/docs/guides/database/connecting-to-postgres",
  ].map((r) => new Paragraph({ children: runs(r), spacing: { after: 160 }, indent: { left: 720, hanging: 720 } })),
];

const page = { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } };
const doc = new Document({
  creator: group.groupName,
  title: "PPCLFMVRS Project Report",
  description: `Privacy-Preserving Campus Lost-and-Found Matching and Verified Recovery System — ${group.groupName}, OAU`,
  features: { updateFields: true },
  styles: {
    default: { document: { run: { font: FONT, size: 24 }, paragraph: { spacing: { line: 360 } } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: FONT }, paragraph: { keepNext: true, keepLines: true, spacing: { before: 360, after: 240 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: FONT }, paragraph: { keepNext: true, keepLines: true, spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ],
  },
  sections: [
    {
      // titlePage: the cover gets its own (empty) footer, so it carries no page number.
      properties: { titlePage: true, page },
      footers: {
        first: new Footer({ children: [new Paragraph({ children: [] })] }),
        default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 20 })] })] }),
      },
      children: [
        ...titlePage, ...contents,
        ...s1, ...s2, ...s3, ...s4, ...s5, ...s6, ...s7, ...s8, ...s9, ...s10, ...s11, ...s12, ...s13, ...s14, ...s15, ...s16,
        ...references,
      ],
    },
  ],
});

// Writes the Word file, then a PDF beside it when LibreOffice is installed.
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("wrote", OUT);
  const soffice = "/Applications/LibreOffice.app/Contents/MacOS/soffice";
  if (!fs.existsSync(soffice)) return console.warn("LibreOffice not found: skipping the PDF.");
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), "lo-"));
  require("child_process").execFileSync(soffice, ["--headless", `-env:UserInstallation=file://${profile}`, "--convert-to", "pdf", "--outdir", path.dirname(OUT), OUT], { stdio: "ignore" });
  fs.rmSync(profile, { recursive: true, force: true });
  console.log("wrote", OUT.replace(/\.docx$/, ".pdf"));
});
