// Builds the SCTPSTVS SWEP defence report as a Word document (and a PDF beside it when LibreOffice is installed).
// Usage: node build-report.js [output.docx]   (default: ~/Desktop)
// Uses the docx package installed in ../node_modules and the shared group details in ../group.js.
const fs = require("fs");
const os = require("os");
const path = require("path");
const {
  AlignmentType, BorderStyle, Document, Footer, HeadingLevel, ImageRun, LevelFormat, Packer,
  PageNumber, Paragraph, ShadingType, Table, TableCell, TableLayoutType, TableOfContents, TableRow, TextRun, WidthType,
} = require("docx");

process.env.PROJECT ??= "sctpstvs";
const group = require("../group");

const OUT = process.argv[2] ?? path.join(os.homedir(), "Desktop", "SCTPSTVS Project Report - SWEP Group 15.docx");
const SHOTS = path.join(__dirname, "..", "..", "screenshots", "sctpstvs");
const ASSETS = path.join(__dirname, "assets");
const FONT = "Times New Roman";
const CONTENT_WIDTH = 9026; // A4 width 11906 minus two 1440 margins, in DXA
const BLANK = "______________________________";

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

let listInstance = 0;
const bullets = (items) => items.map((t) => new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: runs(t) }));
const numbered = (items) => {
  const instance = ++listInstance;
  return items.map((t) => new Paragraph({ numbering: { reference: "numbers", level: 0, instance }, spacing: { after: 80 }, children: runs(t) }));
};

// A shaded block of code, kept on one page.
function code(lines) {
  return lines.map((line, i) =>
    new Paragraph({
      keepNext: i < lines.length - 1,
      keepLines: true,
      shading: { fill: "F1F4F7", type: ShadingType.CLEAR, color: "auto" },
      spacing: { before: i === 0 ? 120 : 0, after: i === lines.length - 1 ? 200 : 0, line: 260 },
      indent: { left: 240, right: 240 },
      children: [new TextRun({ text: line || " ", font: "Courier New", size: 19 })],
    }),
  );
}

const border = { style: BorderStyle.SINGLE, size: 4, color: "808080" };
const borders = { top: border, bottom: border, left: border, right: border };
function table(header, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  const scaled = widths.map((w) => Math.round((w / total) * CONTENT_WIDTH));
  scaled[scaled.length - 1] += CONTENT_WIDTH - scaled.reduce((a, b) => a + b, 0);
  // keepNext on every row keeps a table on one page, together with the caption that follows it.
  const cell = (text, i, head, last) =>
    new TableCell({
      borders,
      width: { size: scaled[i], type: WidthType.DXA },
      shading: head ? { fill: "D6E4F0", type: ShadingType.CLEAR, color: "auto" } : undefined,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [new Paragraph({ keepNext: true, spacing: { after: 0, line: 276 }, children: runs(text, { size: 22, bold: head || undefined }) })],
    });
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    columnWidths: scaled,
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: header.map((h, i) => cell(h, i, true, false)) }),
      ...rows.map((r, j) => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, i, false, j === rows.length - 1)) })),
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
// Screenshots are 1440 × 900 (16:10).
const shot = (name, text) => figure(path.join(SHOTS, name), 470, 294, text);

// ---------- content ----------
// Sections start on a new page only where a break helps the reader; short ones run on.
const S = (text, newPage = true) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: newPage, children: [new TextRun(text)] });

const memberRows = group.hasMembers
  ? group.members.map((m, i) => [String(i + 1), m.name, m.matric, m.department])
  : [1, 2, 3, 4, 5].map((n) => [String(n), "", "", ""]);

const titlePage = [
  new Paragraph({ spacing: { before: 400 } }),
  center("**SECURE CAMPUS TRANSPORT PAYMENT AND SINGLE-USE TICKET VERIFICATION SYSTEM (SCTPSTVS)**", { run: { size: 32 } }),
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
  P("SCTPSTVS is a web app for paying campus shuttle fares at Obafemi Awolowo University (OAU). In the app it is called **Campus Ride Pass**. A student pays from a wallet on their phone and gets a ticket with a signed QR code. The conductor scans it on the bus. Each ticket works **once**, only on its route, and only until it expires, so fake, edited and reused tickets are caught."),
  P("The app is a working demonstration. No real money moves: wallet top-ups are a clearly labelled demo, and the eight routes and their fares are example values, not OAU’s real fares. Everything else, from the signatures to the single-use check, works as it would in a real system."),
  H2("1.1 Aim"),
  P("To design and build a secure web app for paying campus transport fares, with tickets that can be verified on board and used only once."),
  H2("1.2 Objectives"),
  ...numbered([
    "Give each student a wallet whose balance comes from a record of every payment that can never be changed.",
    "Sell tickets at the fare stored on the server, never at a price sent by the phone.",
    "Issue each ticket as a digitally signed QR code, with a short backup code for typing.",
    "Let a conductor check a ticket with one scan, and accept each ticket only once.",
    "Catch fake, edited, reused, expired and wrong-route tickets, and log every attempt.",
    "Give the transport office sales figures, refunds, a fraud log and control of fares.",
    "Prove the security rules with automated tests, including many requests arriving at the same moment.",
  ]),
];

const s2 = [
  S("2. The Problem and the App’s Implications for the Campus"),
  H2("2.1 The problem"),
  P("Every common way of paying a bus fare has a weakness:"),
  ...bullets([
    "**Cash on board** needs change, slows boarding, and leaves no record of how many fares were paid on each route.",
    "**Paper tickets** can be photocopied or used twice, and a conductor cannot tell a copy from the original.",
    "**Bank transfers and mobile money** make the conductor check each payment alert by hand, and a screenshot of an alert can be edited or shown twice.",
    "**The transport office** gets no reliable figures on sales, rides per route or attempts to cheat.",
  ]),
  H2("2.2 Implications for the campus"),
  table(
    ["Who or what", "What changes with SCTPSTVS"],
    [
      ["Students", "No cash or change needed; the ticket is on the phone before boarding; unused tickets can be refunded"],
      ["Conductors", "One scan gives a large green VALID or red ALREADY USED screen; no cash to count"],
      ["Transport office", "Live sales and rides per route, refunds, fare changes and wallet credits from one dashboard"],
      ["Fraud", "Copied, edited, reused, expired and wrong-route tickets are refused and logged with time, vehicle and conductor"],
      ["Accountability", "Every naira is a row in a record that cannot be edited or deleted, so balances can always be traced"],
      ["Planning", "Sales and rides per route show where demand is highest"],
    ],
    [28, 72],
  ),
  caption("Table 2.1: Implications of SCTPSTVS for the OAU campus"),
  H2("2.3 Risks to manage"),
  ...bullets([
    "**Network:** the conductor’s phone needs data to check tickets, because only the central database knows whether a ticket has been used.",
    "**The signing key:** anyone holding the private key could make tickets, so it lives only in the server’s settings, never in the code or the database.",
    "**Smartphones:** buying needs a phone or computer; the 10-character backup code can be written down and typed by the conductor.",
    "**Personal data:** the app stores names and matric numbers, so passwords are hashed with scrypt and only a hash of each session token is kept.",
    "**Real money:** before real payments, top-ups must come from a payment provider that confirms each payment.",
  ]),
];

const s3 = [
  S("3. Distinct Features and Specifications", false),
  H2("3.1 Distinct features"),
  ...bullets([
    "**Wallet:** a balance in naira, built from a ledger of top-ups, purchases and refunds; demo top-ups for students and credits by the admin.",
    "**Server-side pricing:** the phone sends only the route; the server charges the fare stored for that route.",
    "**Signed QR tickets:** each ticket carries an Ed25519 digital signature that only the server can make.",
    "**Single use:** the first scan marks the ticket used; any later scan shows ALREADY USED with the time and bus of the first scan.",
    "**Conductor screen:** camera scanner or typed backup code, and five clear results: VALID, ALREADY USED, EXPIRED, WRONG ROUTE and INVALID, each with its own vibration pattern.",
    "**Admin dashboard:** sales and redemption totals, per-route figures, ticket search and refunds, fraud log, fares, wallet credits and vehicles.",
    "**Accessible, phone-first design:** labelled fields, keyboard use, live result announcements, light and dark themes, and the Atkinson Hyperlegible typeface.",
  ]),
  H2("3.2 Specifications"),
  table(
    ["Item", "Specification"],
    [
      ["Type", "Web app; works in any modern browser on phones and computers"],
      ["Users", "Three roles: student, conductor and admin"],
      ["Ticket token", "`SCT1.<payload>.<signature>`, signed with Ed25519"],
      ["Ticket ID", "96 random bits"],
      ["Backup code", "10 characters (Crockford base32), 50 random bits, e.g. SW7H3-ANCH4"],
      ["Ticket validity", "2 hours after purchase by default (setting `TICKET_TTL_MINUTES`)"],
      ["Money", "Whole kobo (₦1 = 100 kobo); top-ups and credits up to ₦20,000 each; fares ₦10 to ₦5,000"],
      ["Demo data", "8 routes, 5 vehicles, 4 accounts (2 students, 1 conductor, 1 admin)"],
      ["Storage", "PostgreSQL: Supabase in production, PGlite locally; 8 tables in the schema `sctpstvs`"],
      ["Passwords and sessions", "scrypt with a random salt; 256-bit session token in an httpOnly cookie for 12 hours"],
      ["Automated tests", "35, all passing"],
    ],
    [28, 72],
  ),
  caption("Table 3.1: SCTPSTVS specifications"),
];

const s4 = [
  S("4. Comparison with Existing Ways of Paying", false),
  P("The table compares SCTPSTVS with the usual ways of paying for campus transport. It describes how each method works, not measured figures.", { keepNext: true }),
  table(
    ["Question", "Cash fare", "Paper ticket", "Transfer or mobile money", "QR tickets in general", "SCTPSTVS"],
    [
      ["Quick to board?", "Slow: change", "Yes", "Slow: alert checked by hand", "Yes", "Yes: one scan"],
      ["Stops copies and reuse?", "No ticket to copy", "No: copies look the same", "No: screenshots can be edited", "Only if each code is checked once", "Yes: signed and single-use"],
      ["Record of sales per route?", "Hand count", "Ticket stubs", "Bank statement only", "Usually", "Yes, live"],
      ["Log of cheating attempts?", "No", "No", "No", "Depends on the system", "Yes, with time and bus"],
      ["Refunds?", "By hand", "By hand", "By hand", "Depends on the system", "One click, recorded"],
      ["Works without data?", "Yes", "Yes", "No", "Some systems", "No: the check needs the server"],
      ["Needs a smartphone?", "No", "No", "Yes", "Phone or printout", "To buy; the code can be written down"],
    ],
    [20, 13, 15, 18, 17, 17],
  ),
  caption("Table 4.1: SCTPSTVS compared with common ways of paying for campus transport"),
  P("A QR code is only a way of carrying data (ISO/IEC 18004). It stops fraud only when the data is signed and the system remembers which tickets have been used. SCTPSTVS does both."),
];

const s5 = [
  S("5. Tools and Technology Stack"),
  table(
    ["Tool", "What we used it for"],
    [
      ["Next.js 16 and React 19", "The web app: pages for students, conductors and admins, and the API routes"],
      ["TypeScript", "The programming language, with type checks that catch mistakes early"],
      ["Tailwind CSS 4", "Styling, light and dark themes, layouts that work at phone width"],
      ["Node.js crypto module", "Ed25519 signatures, scrypt password hashing, SHA-256 and secure random numbers"],
      ["qrcode", "Drawing each ticket as a QR code"],
      ["@zxing/browser", "Reading QR codes with the phone camera"],
      ["PostgreSQL (Supabase)", "The production database, shared by the group’s three apps"],
      ["postgres.js", "Connecting to PostgreSQL through Supabase’s Transaction pooler"],
      ["PGlite", "PostgreSQL inside Node.js, for local use and tests with no database server"],
      ["Vitest", "Automated tests"],
      ["ESLint", "Checks code quality"],
      ["Git and GitHub", "Version control and sharing the code"],
      ["Vercel", "Hosting the app online"],
    ],
    [32, 68],
  ),
  caption("Table 5.1: Tools and technology used"),
  P("**Tools needed to run it:** a computer with Node.js 20.9 or newer (22 recommended), npm (comes with Node.js), Git and a web browser. A phone camera is needed only for QR scanning, which browsers allow on HTTPS pages or on localhost. For deployment: a Supabase account and a Vercel account (both have free plans)."),
];

const s6 = [
  S("6. System Architecture"),
  ...figure(path.join(ASSETS, "architecture.png"), 560, 405, "Figure 6.1: System architecture"),
  ...bullets([
    "**Phones and browsers:** the student, conductor and admin pages. They hold no secrets; they only show what the server sends.",
    "**On the server:** the API routes check the user’s role and the request’s origin, then call the ticket and wallet logic. Only the server holds the private signing key.",
    "**In the database:** PostgreSQL decides whether a ticket is still unused, and its triggers guard the wallet ledger. Because nothing is kept in server memory or files, several copies of the server can run at once.",
  ]),
];

const s7 = [
  S("7. How the Security Works"),
  H2("7.1 Signed tickets cannot be forged"),
  P("When a ticket is bought, the server makes a token in three parts:"),
  ...code(["SCT1.<payload>.<signature>"]),
  P("The payload holds the ticket ID, the route, the issue and expiry times and the key ID. The signature is an **Ed25519** digital signature (Bernstein et al., 2012; Josefsson and Liusvaara, 2017) made with a private key that only the server has. Changing any character of the payload or the signature makes the check fail, and the conductor sees **INVALID**. Without the private key, nobody can make a ticket, extend its expiry or move it to another route."),
  P("We chose Ed25519 over a shared-secret code (HMAC) because the key that *checks* a ticket is public (`/api/public-key`). A future offline scanner could check signatures without being able to create tickets."),
  H2("7.2 The database makes each ticket single-use"),
  P("A valid signature only proves the server issued the ticket. Whether it is still unused lives in the database. The last step of every check is one conditional update:"),
  ...code([
    "UPDATE sctpstvs.tickets",
    "SET status = 'redeemed', redeemed_at = $1,",
    "    redeemed_by = $2, redeemed_vehicle = $3",
    "WHERE id = $4 AND status = 'active' AND expires_at > $1",
    "RETURNING id",
  ]),
  P("PostgreSQL locks the ticket’s row while it updates it. A second update of the same row waits, then checks the WHERE clause again against the saved row; the status is no longer ‘active’, so it changes nothing (PostgreSQL Global Development Group, 2024). If two conductors scan the same ticket at the same instant, exactly one gets **VALID** and the other gets **ALREADY USED**, with the time and bus of the first scan."),
  H2("7.3 The wallet ledger can only grow"),
  ...bullets([
    "A balance is never stored; it is the sum of the wallet’s ledger rows, in whole kobo.",
    "Database triggers refuse any UPDATE, DELETE or TRUNCATE of the ledger, so no record can be changed or erased.",
    "Another trigger locks the owner’s row and refuses any new row that would take the balance below zero, even from a direct database insert.",
    "A CHECK rule keeps purchases negative and top-ups and refunds positive.",
  ]),
  H2("7.4 Row locks stop double spending and double refunds"),
  P("A purchase runs in one transaction that first locks the buyer’s row (`SELECT … FOR NO KEY UPDATE`), then checks the balance, writes the ledger row and creates the ticket. Any other purchase, top-up or refund for the same wallet waits for that lock, on any server. A refund locks the ticket’s row (`SELECT … FOR UPDATE`), so two admins clicking at once refund it only once."),
  H2("7.5 Other protections"),
  table(
    ["Attack", "What stops it"],
    [
      ["Sending a lower fare from the phone", "The server ignores it and charges the fare stored for the route"],
      ["Guessing passwords", "5 tries per account and 30 per IP address every 15 minutes; passwords hashed with scrypt (Percival, 2009)"],
      ["Guessing backup codes", "50 random bits (about 10¹⁵ codes) and 30 checks per conductor per minute"],
      ["Using a ticket on the wrong bus", "WRONG ROUTE; the ticket stays unused"],
      ["Requests from another website", "Refused with 403 when the Origin header is from another site (OWASP, 2024)"],
      ["Stolen session database", "Only a SHA-256 hash of each session token is stored"],
      ["Finding which matric numbers exist", "An unknown number still runs scrypt, so the response time gives nothing away"],
      ["Students opening admin pages", "Every API route checks the user’s role on the server"],
    ],
    [35, 65],
  ),
  caption("Table 7.1: Attacks and how SCTPSTVS stops them. Every refusal goes in the fraud log."),
];

const s8 = [
  S("8. How to Use the App"),
  P("The demo accounts are listed in the repository’s README and on the sign-in page. The screens below come from a local copy of the app with demo data."),
  H2("8.1 Sign in (student)"),
  P("Open the app, type your matric number and password, and press **Sign in**. Conductors and admins use their staff number."),
  ...shot("signin.png", "Figure 8.1: The sign-in page, with the demo accounts opened."),
  H2("8.2 Check the wallet and top up"),
  P("The wallet shows your balance. Press **+ ₦500**, **+ ₦1,000** or **+ ₦2,000**, or type another amount up to ₦20,000, to add demo credit. **Wallet history** lists every top-up, purchase and refund. Your tickets are on the right."),
  ...shot("student-wallet.png", "Figure 8.2: The student wallet, demo top-up and the list of tickets."),
  H2("8.3 Buy a ticket"),
  P("Under **Buy a ticket**, choose your route and press **Pay ₦… from wallet**. The fare comes from the server. If the balance is too low, the app says how much is missing."),
  ...shot("student-buy.png", "Figure 8.3: Choosing the Main Gate – SUB route before paying ₦200 from the wallet."),
  H2("8.4 Show the QR ticket"),
  P("Press **Show ticket QR** on an active ticket. Turn the screen brightness up and show it to the conductor. The backup code under the QR can be read out or typed if the camera cannot scan."),
  ...shot("ticket-qr.png", "Figure 8.4: The QR ticket with its backup code and expiry time."),
  H2("8.5 Verify a ticket (conductor)"),
  P("Sign in as a conductor and choose the vehicle you are on. Press **Start camera scanner** and point it at the QR, or type the backup code and press **Check ticket**."),
  ...shot("conductor.png", "Figure 8.5: The conductor screen on Bus 1, with the scanner and the backup-code box."),
  H2("8.6 Read the result"),
  P("The result fills the screen so it can be read at a glance. **VALID** means let the passenger board. **ALREADY USED** gives the time, bus and conductor of the first scan. **INVALID** means the code was not issued by the system or has been changed."),
  ...shot("verify-valid.png", "Figure 8.6: VALID: the first scan of a ticket."),
  ...shot("verify-already-used.png", "Figure 8.7: ALREADY USED: the same ticket scanned again, on another bus."),
  ...shot("verify-invalid.png", "Figure 8.8: INVALID: a ticket whose signed data was edited."),
  H2("8.7 See sales and tickets (admin)"),
  P("Sign in as the admin. **Overview** shows tickets sold, sales, rides, wallet totals, fraud events and a table per route. **Tickets** lists every ticket by status, route or search, with a **Refund** button on unused tickets."),
  ...shot("admin-overview.png", "Figure 8.9: The admin overview: sales, redemptions and figures per route."),
  ...shot("admin-tickets.png", "Figure 8.10: The ticket list, with one refunded, one active and two used tickets."),
  H2("8.8 Review the fraud log (admin)"),
  P("**Fraud log** lists every refused check, newest first: reuse attempts, invalid signatures, unknown codes, wrong routes, expired and refunded tickets, and rate-limit hits. Each row shows who checked it and on which vehicle."),
  ...shot("admin-fraud.png", "Figure 8.11: The fraud log after the demo checks."),
  H2("8.9 Change fares (admin)"),
  P("**Routes & fares** lets the admin change a fare, stop selling a route or add a new one. A new fare applies to new purchases only; sold tickets keep the fare that was paid."),
  ...shot("admin-fares.png", "Figure 8.12: Routes and fares. The fares are demo values, not OAU’s real fares."),
];

const s9 = [
  S("9. App Life Cycle", false),
  H2("9.1 Development life cycle"),
  P("We built SCTPSTVS in short steps, testing each one before the next. The dates come from the repository’s history.", { keepNext: true }),
  table(
    ["Stage", "Date (2026)", "What we did"],
    [
      ["1. Planning and design", "By 23 Oct", "Chose the project; designed tokens and the ledger"],
      ["2. Set-up", "23 Oct", "Next.js app with tests, QR and scanner libraries"],
      ["3. Core and tests", "23 Oct", "Signed single-use tickets, wallet ledger, tests"],
      ["4. API", "23 Oct", "Sign-in, wallet, tickets, verification, admin"],
      ["5. Screens", "23 Oct", "Student, conductor and admin screens"],
      ["6. Polish and documents", "23 Oct", "Phone spacing, app icon, README"],
      ["7. Production database", "27 Oct", "Moved from SQLite to PostgreSQL; fixed races"],
      ["8. Live testing", "27–28 Oct", "Tested on Supabase; ran it for group testers"],
      ["9. Deployment", "Next", "Host on Vercel with the Supabase database"],
      ["10. Maintenance", "Ongoing", "Real payments, real fares, new features"],
    ],
    [28, 15, 57],
  ),
  caption("Table 9.1: Development life cycle"),
  H2("9.2 Life of a ticket"),
  ...figure(path.join(ASSETS, "ticket-lifecycle.png"), 490, 205, "Figure 9.1: The states a ticket passes through"),
  P("A ticket is **bought** (the fare leaves the wallet) and becomes **active**. From there it ends in one of three ways: **redeemed** by the first scan on the right route, **expired** when its time runs out, or **refunded** by an admin while unused. A scan on the wrong route shows WRONG ROUTE and leaves the ticket active."),
  H2("9.3 Life of a check"),
  P("The conductor scans or types a code → the server checks the rate limit → it checks the signature → it finds the ticket → it checks the status, expiry and route → the guarded UPDATE marks it used → the phone shows the result. Every refusal on the way is written to the fraud log."),
];

const s10 = [
  S("10. What the System Can Handle", false),
  table(
    ["Measure", "Value"],
    [
      ["Simultaneous scans of one ticket", "20 tested: exactly 1 VALID"],
      ["Simultaneous purchases, one wallet", "10 tested: exactly 1 ticket"],
      ["Simultaneous refunds of one ticket", "3 tested: refunded once"],
      ["Copies of the server", "Any number: no state kept on a server"],
      ["Database connections per server", "Up to 5, through the pooler"],
      ["Checks per conductor", "30 per minute"],
      ["Sign-in attempts per 15 minutes", "5 per account, 30 per IP address"],
      ["Ticket validity", "2 hours by default"],
      ["Top-up or credit", "Up to ₦20,000 at a time"],
      ["Fares", "₦10 to ₦5,000 per route"],
      ["Possible backup codes", "2⁵⁰, about 1.1 × 10¹⁵"],
      ["Devices", "Phones and computers; camera needs HTTPS"],
    ],
    [45, 55],
  ),
  caption("Table 10.1: Capacity and limits of the system, from the code and the tests"),
  P("Each check is a few small database queries, so the database and the network set the real limit. We have not yet run a load test with hundreds of users at once; that is listed as open work in Section 14."),
];

const s11 = [
  S("11. Tests Done and Results"),
  H2("11.1 Automated tests"),
  table(
    ["Test file", "Tests", "What they check"],
    [
      ["Tickets and wallet", "19", "Fare from the route; insufficient funds; simultaneous purchases; ledger never negative and append-only; single use; backup code; expiry; tampered, forged and unknown codes; wrong route; simultaneous scans; refunds"],
      ["Ticket tokens", "7", "Signing and checking; edited payload or signature rejected; wrong key rejected; bad input; backup-code format"],
      ["Sign-in and sessions", "5", "scrypt hashing; right password only; sessions expire and can be revoked; database rate limiter, including a burst"],
      ["API route", "4", "Server fare charged when the phone sends 1 kobo; role rules; cross-site block; a ticket verifies exactly once"],
    ],
    [22, 9, 69],
  ),
  caption("Table 11.1: Automated tests (35 in total, all passing in 14.4 s on PGlite)"),
  P("Type checking and code linting also pass. PGlite runs one query at a time, so the same tests were also run against a real PostgreSQL server, where requests truly overlap."),
  H2("11.2 Concurrency results"),
  table(
    ["Test", "At the same moment", "Result"],
    [
      ["Scans of one ticket", "20", "Exactly 1 VALID and 19 ALREADY USED"],
      ["Two database connections redeem one ticket", "2", "1 VALID and 1 ALREADY USED"],
      ["Purchases of a ₦200 ticket with ₦250 in the wallet", "10", "Exactly 1 ticket; ₦50 left"],
      ["Direct ₦500 debits from a ₦2,000 wallet", "10", "Exactly 4 accepted; balance ₦0, never negative"],
      ["Refunds of one ticket", "3", "Refunded once; balance back to ₦2,000"],
      ["Rate-limit checks against a limit of 5", "8", "Exactly 5 allowed; 1 marked as the first refusal"],
    ],
    [46, 18, 36],
  ),
  caption("Table 11.2: Results of the concurrency tests"),
  H2("11.3 Live test on Supabase"),
  P("We also tested the running app over HTTP against the Supabase database:"),
  table(
    ["Test", "Result"],
    [
      ["20 simultaneous scans of one ticket", "1 VALID and 19 ALREADY USED"],
      ["Purchase with a fare of 1 kobo sent by the client", "Ignored; the real fare of ₦150 was charged"],
      ["QR code with edited data", "INVALID, logged as an invalid signature"],
      ["Request from another website", "Refused with 403"],
    ],
    [50, 50],
  ),
  caption("Table 11.3: Live test results on Supabase"),
  P("In the demo session used for the screenshots, a student with ₦100 was refused a ₦200 ticket with an *insufficient funds* message, and a Main Gate – SUB ticket checked on Bus 4 showed WRONG ROUTE and stayed unused."),
];

const s12 = [
  S("12. Errors Found and Fixes", false),
  table(
    ["Problem", "Cause", "Fix"],
    [
      ["Data would be lost on online hosts", "The first version kept data in a local SQLite file, which serverless hosts do not keep", "Moved all data to PostgreSQL (Supabase in production, PGlite locally)"],
      ["Rate limits not shared between servers", "Counters were kept in one server’s memory", "Counters moved into a database table, checked under a lock"],
      ["Two debits could both pass the balance check", "PostgreSQL lets writes run side by side, unlike SQLite’s single writer", "The no-negative trigger and purchases now lock the wallet owner’s row first"],
      ["A ticket could be refunded twice", "Two refunds could both read the ticket as unused", "The refund locks the ticket row with FOR UPDATE"],
      ["Race tests passed without really racing", "Requests queued for new database connections instead of overlapping", "Added `warmPool()` to open the connections before each race"],
      ["Each server copy would make its own signing key", "The key was generated into a local file", "Production now requires `SIGNING_PRIVATE_KEY`; `npm run gen-key` prints one"],
      ["Spacing too tight on phones", "Found when testing the screens at phone width", "Adjusted spacing at phone width and added an app icon"],
    ],
    [28, 36, 36],
  ),
  caption("Table 12.1: Errors found during testing and how we fixed them. The fixes are in the repository’s history."),
];

const s13 = [
  S("13. Installation and Deployment Guide"),
  H2("13.1 Run it on your computer"),
  ...numbered([
    "Install Node.js 20.9 or newer (22 recommended) and Git.",
    `Download the code: \`git clone ${group.repoUrl}\``,
    "Open the app folder: `cd swep-g-15/sctpstvs`",
    "Install the packages: `npm install`",
    "Start the app: `npm run dev`, then open http://localhost:3000",
    "Run the tests: `npm test`",
  ]),
  P("No database server is needed locally: the app uses PGlite in its `data` folder and makes a signing key there on first run. Delete `data` to start again with fresh demo data."),
  H2("13.2 Put it online"),
  ...numbered([
    "In Supabase, open **Project Settings → Database → Connection string → Transaction pooler** and copy the link (port 6543), adding the database password. It looks like `postgresql://postgres.<project-ref>:<password>@aws-0-<region>.pooler.supabase.com:6543/postgres`.",
    "Run `npm run gen-key` and copy the printed `SIGNING_PRIVATE_KEY` value. Keep it secret.",
    "Create a Vercel project from the GitHub repository, with **Root Directory** set to `sctpstvs`.",
    "In Vercel, add the settings below, then deploy.",
    "Sign in with the demo accounts, then change or remove them before real testers use the app.",
  ]),
  table(
    ["Setting", "Value"],
    [
      ["DATABASE_URL", "The Supabase Transaction pooler link (required)"],
      ["SIGNING_PRIVATE_KEY", "The key printed by `npm run gen-key` (required in production)"],
      ["DATABASE_SCHEMA", "Optional; defaults to sctpstvs"],
      ["TICKET_TTL_MINUTES", "Optional; how long a ticket stays valid, default 120"],
    ],
    [35, 65],
  ),
  caption("Table 13.1: Production settings"),
  P("On its first request the app creates its schema, tables and triggers and adds the demo data, all in one transaction under a lock, so two servers starting together cannot both do it. Changing the signing key later makes every unused ticket invalid."),
];

const s14 = [
  S("14. Implemented and Open Features", false),
  table(
    ["Feature", "Status"],
    [
      ["Accounts and roles, scrypt passwords, sessions", "Implemented"],
      ["Wallet with append-only ledger, demo top-up, admin credits", "Implemented"],
      ["Tickets at the server’s fare, Ed25519-signed QR and backup code", "Implemented"],
      ["Conductor scanner and five result screens", "Implemented"],
      ["Single-use check that holds under simultaneous scans", "Implemented"],
      ["Admin dashboard: sales, tickets, refunds, fraud log, fares, vehicles", "Implemented"],
      ["Rate limits shared by every server", "Implemented"],
      ["PostgreSQL storage for online hosting", "Implemented"],
      ["Real payments through a payment provider", "Open"],
      ["OAU’s real routes, fares and vehicles", "Open: needs the transport office"],
      ["Link accounts to university student records", "Open"],
      ["Self-registration and password reset", "Open"],
      ["Offline checking for areas with no network", "Open"],
      ["Signing-key rotation screen", "Open"],
      ["Load test with many users", "Open"],
    ],
    [65, 35],
  ),
  caption("Table 14.1: Implemented features and open improvements"),
];

const s15 = [
  S("15. Conclusion"),
  P("SCTPSTVS shows that campus transport fares can be paid and checked securely with a phone and a web app. All seven objectives were met. The tests prove the key promises: a ticket cannot be forged, it works only once even when scanned twenty times at the same moment, and a wallet can never be overspent."),
  H2("15.1 Limitations and mitigation"),
  table(
    ["Limitation", "Mitigation"],
    [
      ["No real payments yet: top-ups are demo credit", "Credit wallets only after a payment provider confirms the payment; admin credits stand in for cash paid at an office"],
      ["Routes and fares are invented", "The admin can set real routes and fares in minutes once the transport office supplies them"],
      ["Checking needs a network connection", "The signature can already be checked offline with the public key; an offline mode would sync used tickets later"],
      ["Demo accounts are created in production too", "Change their passwords or remove them before real testers use the app"],
      ["A screenshot of a ticket can be passed on", "It still works only once, so a ticket is never used twice"],
      ["Losing the signing key cancels unused tickets", "Keep it in the host’s secret settings; the key ID in each token allows rotation later"],
      ["No self-registration or password reset", "Link accounts to the university’s student records in a later version"],
    ],
    [42, 58],
  ),
  caption("Table 15.1: Limitations and how each is mitigated"),
  H2("15.2 Recommendations"),
  ...numbered([
    "Agree real routes and fares with the university’s transport office and enter them in the app.",
    "Connect a payment provider so top-ups move real money, confirmed by signed notifications.",
    "Run a pilot on one route, then a load test before opening it to all students.",
  ]),
];

const s16 = [
  S("16. Project Links and Group Members"),
  P(`**Source code and documentation:** ${group.repoUrl}`),
  P("The `sctpstvs` folder holds the full code, the installation and deployment guide, the security design, the list of tests and the history of every change and fix."),
  table(["S/N", "Name", "Matriculation number", "Department"], memberRows, [7, 38, 25, 30]),
  caption("Table 16.1: Group members"),
];

const references = [
  S("References"),
  ...[
    "Bernstein, D. J., Duif, N., Lange, T., Schwabe, P., & Yang, B.-Y. (2012). High-speed high-security signatures. *Journal of Cryptographic Engineering, 2*(2), 77–89.",
    "Braille Institute. (2019). *Atkinson Hyperlegible typeface*. https://www.brailleinstitute.org/freefont/",
    "Crockford, D. *Base 32*. https://www.crockford.com/base32.html",
    "International Organization for Standardization. (2015). *ISO/IEC 18004:2015 Information technology — Automatic identification and data capture techniques — QR Code bar code symbology specification*.",
    "Josefsson, S., & Liusvaara, I. (2017). *Edwards-Curve Digital Signature Algorithm (EdDSA)* (RFC 8032). Internet Engineering Task Force.",
    "OWASP Foundation. (2024). *Cross-Site Request Forgery Prevention Cheat Sheet*. https://cheatsheetseries.owasp.org/cheatsheets/Cross-Site_Request_Forgery_Prevention_Cheat_Sheet.html",
    "Percival, C. (2009). *Stronger key derivation via sequential memory-hard functions*. BSDCan 2009, Ottawa.",
    "PostgreSQL Global Development Group. (2024). *PostgreSQL 17 Documentation, Chapter 13: Concurrency Control*. https://www.postgresql.org/docs/17/mvcc.html",
    "Supabase. *Connecting to Postgres: connection pooler*. https://supabase.com/docs/guides/database/connecting-to-postgres",
  ].map((r) => new Paragraph({ children: runs(r), spacing: { after: 160 }, indent: { left: 720, hanging: 720 } })),
];

const page = { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } };
const doc = new Document({
  creator: group.groupName,
  title: "SCTPSTVS Project Report",
  description: `Secure Campus Transport Payment and Single-Use Ticket Verification System — ${group.groupName}, OAU`,
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
  console.log("wrote", OUT.replace(/\.docx$/, ".pdf"));
});
