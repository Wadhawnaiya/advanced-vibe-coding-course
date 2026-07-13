/**
 * Student-Only Workbook — Vibe Coding with Claude Code
 */
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
        BorderStyle, WidthType, ShadingType, PageNumber, PageBreak } = require('docx');
const fs = require('fs');

const C = {
  primary: "0B3D2E",
  accent: "1A7A5C",
  dark: "1C2833",
  muted: "5D6D7E",
  light: "E8F6F1",
  alt: "F4F6F7",
  white: "FFFFFF",
  border: "D5D8DC",
  softY: "FEF9E7",
  warn: "9A7D0A",
  softB: "EBF5FB",
};
const PAGE_W = 12240, PAGE_H = 15840, M = 1008, CW = PAGE_W - M * 2;
const nb = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const nbs = { top: nb, bottom: nb, left: nb, right: nb };
const tb = { style: BorderStyle.SINGLE, size: 4, color: C.border };
const tbs = { top: tb, bottom: tb, left: tb, right: tb };

const p = (text, o = {}) => new Paragraph({
  spacing: { after: o.after ?? 120, before: o.before ?? 0, line: 276 },
  alignment: o.align || AlignmentType.LEFT,
  children: [new TextRun({ text, font: "Arial", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || C.dark })],
});
const multi = (runs, o = {}) => new Paragraph({
  spacing: { after: o.after ?? 120, line: 276 },
  children: runs.map(r => new TextRun({ text: r.text, font: "Arial", size: r.size || 21, bold: r.bold, italics: r.italics, color: r.color || C.dark })),
});
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 160 },
  children: [new TextRun({ text: t, font: "Arial", size: 30, bold: true, color: C.primary })] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 220, after: 120 },
  children: [new TextRun({ text: t, font: "Arial", size: 24, bold: true, color: C.accent })] });
const h3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 160, after: 80 },
  children: [new TextRun({ text: t, font: "Arial", size: 21, bold: true, color: C.dark })] });
const b = (t, ref = "bullets") => new Paragraph({
  numbering: { reference: ref, level: 0 }, spacing: { after: 70, line: 276 },
  children: [new TextRun({ text: t, font: "Arial", size: 20, color: C.dark })],
});
const n = (t, ref = "numbers") => new Paragraph({
  numbering: { reference: ref, level: 0 }, spacing: { after: 70, line: 276 },
  children: [new TextRun({ text: t, font: "Arial", size: 20, color: C.dark })],
});
const sp = (a = 100) => new Paragraph({ spacing: { after: a }, children: [] });
const pb = () => new Paragraph({ children: [new PageBreak()] });

const callout = (title, lines, bg = C.light, tc = C.primary) => {
  const arr = Array.isArray(lines) ? lines : [lines];
  return new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: [CW],
    rows: [
      new TableRow({ children: [new TableCell({ borders: nbs, width: { size: CW, type: WidthType.DXA },
        shading: { fill: bg, type: ShadingType.CLEAR }, margins: { top: 90, bottom: 40, left: 140, right: 140 },
        children: [new Paragraph({ children: [new TextRun({ text: title, font: "Arial", size: 19, bold: true, color: tc })] })] })] }),
      ...arr.map((line, i) => new TableRow({ children: [new TableCell({ borders: nbs, width: { size: CW, type: WidthType.DXA },
        shading: { fill: bg, type: ShadingType.CLEAR },
        margins: { top: 20, bottom: i === arr.length - 1 ? 90 : 20, left: 140, right: 140 },
        children: [new Paragraph({ children: [new TextRun({ text: line, font: "Arial", size: 18, color: C.dark })] })] })] })),
    ],
  });
};

const code = (lines) => {
  const arr = Array.isArray(lines) ? lines : lines.split("\n");
  return new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: [CW],
    rows: arr.map((line, i) => new TableRow({ children: [new TableCell({
      borders: nbs, width: { size: CW, type: WidthType.DXA },
      shading: { fill: "1E2A32", type: ShadingType.CLEAR },
      margins: { top: i === 0 ? 70 : 16, bottom: i === arr.length - 1 ? 70 : 16, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: line || " ", font: "Consolas", size: 16, color: "E8F0F2" })] })],
    })] })),
  });
};

const cell = (text, o = {}) => new TableCell({
  borders: tbs, width: { size: o.w, type: WidthType.DXA },
  shading: { fill: o.fill || C.white, type: ShadingType.CLEAR },
  margins: { top: 50, bottom: 50, left: 70, right: 70 },
  children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: o.size || 17, bold: o.bold, color: o.color || C.dark })] })],
});
const table = (headers, rows, widths) => new Table({
  width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths,
  rows: [
    new TableRow({ children: headers.map((h, i) => cell(h, { w: widths[i], fill: C.primary, bold: true, color: C.white, size: 17 })) }),
    ...rows.map((r, ri) => new TableRow({ children: r.map((c, i) => cell(c, { w: widths[i], fill: ri % 2 ? C.white : C.alt, size: 16 })) })),
  ],
});

const line = (label) => new Paragraph({
  spacing: { after: 160 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.border, space: 1 } },
  children: [new TextRun({ text: label, font: "Arial", size: 18, color: C.muted })],
});

function numberingConfig() {
  const refs = [];
  for (let i = 0; i < 40; i++) {
    refs.push({
      reference: `bullets${i}`,
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
    });
    refs.push({
      reference: `numbers${i}`,
      levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
    });
  }
  refs.push({
    reference: "bullets",
    levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
  });
  refs.push({
    reference: "numbers",
    levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
  });
  return refs;
}

function moduleBlock(num, title, goals, youWill, lab, homework, prompts, bi) {
  return [
    h1(`Module ${num}: ${title}`),
    h2("Learning goals"),
    ...goals.map(g => b(g, `bullets${bi}`)),
    h2("You will be able to"),
    ...youWill.map(g => b(g, `bullets${bi + 1}`)),
    h2("Lab"),
    ...lab,
    h2("Copy-paste prompts"),
    ...prompts,
    h2("Homework"),
    ...homework,
    h2("My notes"),
    line("What I learned:"),
    line("What confused me:"),
    line("Questions for class:"),
    pb(),
  ];
}

async function main() {
  const children = [];

  // Cover
  children.push(
    sp(500),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 },
      children: [new TextRun({ text: "STUDENT WORKBOOK", font: "Arial", size: 22, bold: true, color: C.accent })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 16, color: C.primary, space: 8 } },
      children: [new TextRun({ text: "Vibe Coding with Claude Code", font: "Arial", size: 40, bold: true, color: C.primary })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 160, after: 80 },
      children: [new TextRun({ text: "From Zero Tech Knowledge to Your First Full-Stack App", font: "Arial", size: 24, color: C.dark })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
      children: [new TextRun({ text: "Student Edition — Practice, Labs, Prompts & Checklists", font: "Arial", size: 20, italics: true, color: C.muted })] }),
    sp(200),
    callout("How to use this workbook", [
      "Bring this to every class (print or digital).",
      "Fill the blanks. Write real notes — that is how learning sticks.",
      "Do labs in order. Do not skip Module 1.",
      "When stuck: use the 3-minute protocol (page in Tools section).",
      "You are the director. Claude Code is your builder. You own the result.",
    ]),
    sp(200),
    table(["Field", "Your details"], [
      ["My name", ""],
      ["Class / cohort", ""],
      ["Laptop OS (Windows / Mac / Linux)", ""],
      ["Claude account email", ""],
      ["GitHub username", ""],
      ["Capstone app idea (working title)", ""],
    ], [3600, 6624]),
    pb(),
  );

  // Welcome
  children.push(
    h1("Welcome — What You Are Learning"),
    p("This course teaches you to build real web applications using Claude Code — an AI that can create and edit project files and run commands — even if you have never coded before."),
    h2("What is vibe coding?"),
    p("You describe what you want in plain language. The AI writes and edits the code. You review, test, and steer. You do not need to memorize programming syntax. You do need clear thinking, good instructions, and the habit of checking that things actually work."),
    h2("What success looks like"),
    b("You can open a terminal and start Claude Code in a project folder."),
    b("You can explain frontend, backend, database, and deploy in simple words."),
    b("You can write a short plan (PRD) for an app idea."),
    b("You can direct Claude to build a working full-stack app."),
    b("You can put it on GitHub and deploy a public link."),
    b("You can demo your app and explain your decisions."),
    h2("Course map (10 weeks)"),
    table(["Week", "Focus", "You deliver"], [
      ["0", "Accounts & laptop ready", "Setup checklist done"],
      ["1", "Files, terminal, first Claude page", "Hello webpage"],
      ["2", "Prompt craft + plan mode", "Prompt journal + mini tool"],
      ["3", "How web apps work", "Portfolio site"],
      ["4", "Interactive pages", "Quiz or calculator app"],
      ["5", "Product thinking", "Capstone PRD"],
      ["6", "Data & storage", "Notes/todo with save"],
      ["7", "Full-stack MVP", "Core app flow works"],
      ["8", "Login + polish", "Auth or polish checkpoint"],
      ["9", "GitHub", "Public repo + README"],
      ["10", "Deploy + demo", "Live URL + 5-min demo"],
    ], [1200, 4000, 5024]),
    sp(120),
    callout("Mindset", [
      "Mistakes and red error messages are normal — they are clues, not failure.",
      "Small steps beat giant wishes.",
      "If you cannot explain what your app does, keep simplifying.",
    ], C.softY, C.warn),
    pb(),
  );

  // Week 0
  children.push(
    h1("Week 0 — Setup Checklist"),
    p("Complete before Module 1. Tick each box when done."),
    h2("Accounts"),
    b("I have a working email I can access."),
    b("I created / upgraded my Claude account as instructed (Pro/Max/Team — free plan alone is not enough for Claude Code)."),
    b("I can log into claude.ai in a browser."),
    b("I created a free GitHub account (github.com). Username: _______________"),
    h2("Computer"),
    b("Laptop has enough free space (about 10 GB free is safe)."),
    b("I know my OS: Windows / Mac / Linux (circle one)."),
    b("Browser is Chrome or Edge (updated)."),
    b("I can connect to stable internet."),
    h2("Practice"),
    b("I watched the short “What is a terminal?” video from class."),
    b("I know where my Documents / Downloads folders are."),
    h2("Reflection"),
    line("One app idea I might want to build someday:"),
    line("One fear I have about this course:"),
    pb(),
  );

  // Module 1
  children.push(...moduleBlock(
    1, "Digital Ground Zero — Files, Terminal, First Page",
    [
      "Understand files, folders, and paths.",
      "Open a terminal and run safe basic commands.",
      "Install and start Claude Code.",
      "Create your first webpage with Claude in English.",
    ],
    [
      "Open terminal and go to a folder.",
      "Run claude and log in.",
      "Generate and open a simple HTML page in the browser.",
    ],
    [
      p("1. Create this folder structure on your computer:"),
      code(["Documents (or Home)", "  └── vibe-coding-course", "        └── week1-hello"]),
      p("2. Open terminal in week1-hello."),
      p("3. Install Claude Code if not done (instructor will demo your OS)."),
      p("macOS / Linux:"),
      code(["curl -fsSL https://claude.ai/install.sh | bash"]),
      p("Windows PowerShell:"),
      code(["irm https://claude.ai/install.ps1 | iex"]),
      p("4. Check install:"),
      code(["claude --version"]),
      p("5. Start Claude and build:"),
      code(["claude"]),
      p("Then type:"),
      code(["make me a simple webpage that says Hello, my name is [YOUR NAME]", "and I am learning vibe coding. Use a friendly color background."]),
      p("6. Open the page in your browser (double-click the HTML file or ask Claude to open it)."),
    ],
    [
      b("Customize the page: add your favorite color and a short bio paragraph (ask Claude in English).", "bullets2"),
      b("Screenshot before and after.", "bullets2"),
      b("Write 3 sentences: What surprised me / What confused me / What I want next.", "bullets2"),
    ],
    [
      code(["make me a simple webpage that says Hello, my name is [Name]", "and I am learning vibe coding"]),
      sp(60),
      code(["change the background to soft blue and add a short bio paragraph about me"]),
    ],
    3
  ));

  // Module 2
  children.push(...moduleBlock(
    2, "Directing Claude — Prompt Craft & Plan Mode",
    [
      "Write clear prompts using the PROMPT formula.",
      "Use plan mode before big changes.",
      "Verify work instead of trusting “done.”",
    ],
    [
      "Turn a vague wish into a precise build request.",
      "Use /plan, /clear, /help confidently.",
    ],
    [
      p("PROMPT formula (memorize):"),
      table(["Letter", "Meaning", "Ask yourself"], [
        ["P", "Purpose", "What outcome do I want?"],
        ["R", "Rules / constraints", "What must stay simple?"],
        ["O", "Objects", "Screens, buttons, data?"],
        ["M", "Musts", "What must work to pass?"],
        ["P", "Proof", "How will I test it?"],
        ["T", "Tone / style", "Look and feel?"],
      ], [1000, 2800, 6424]),
      sp(80),
      p("Lab: Build a tip calculator"),
      n("Write a bad one-line prompt first (yes, intentionally bad).", "numbers4"),
      n("Rewrite it with PROMPT.", "numbers4"),
      n("Build with Claude.", "numbers4"),
      n("Partner tries to break it (empty input, weird numbers).", "numbers4"),
      n("Ask Claude to fix based on partner findings.", "numbers4"),
    ],
    [
      b("Prompt journal: rewrite 5 bad prompts into good ones (any life tools).", "bullets5"),
      b("Build a study timer page with PROMPT formula.", "bullets5"),
    ],
    [
      code(["/plan Create a tip calculator webpage. Inputs: bill amount and tip %.", "Show tip and total. Friendly errors for empty/negative bill.", "Large fonts. No frameworks. Wait for approval before coding."]),
      sp(60),
      code(["Implement the approved plan only. Then tell me exact steps to open and test it."]),
    ],
    6
  ));

  // Module 3
  children.push(...moduleBlock(
    3, "How Web Apps Work (No Coding Torture)",
    [
      "Name the parts of a full-stack app in plain language.",
      "Use the restaurant metaphor to explain frontend/backend/data.",
    ],
    [
      "Build a multi-page portfolio site with Claude.",
      "Explain your folder structure out loud.",
    ],
    [
      p("Restaurant metaphor — fill the blanks in class:"),
      table(["Restaurant", "Web app part", "My words"], [
        ["Dining room / menu", "Frontend", ""],
        ["Kitchen", "Backend", ""],
        ["Pantry / recipe book", "Database", ""],
        ["Waiter order", "API / request", ""],
        ["Opening the restaurant", "Deploy", ""],
      ], [2800, 2800, 4624]),
      sp(80),
      p("Lab prompt:"),
      code([
        "Create a simple multi-page personal portfolio site with:",
        "- index.html (home)",
        "- about.html",
        "- projects.html",
        "- contact.html (visual form is OK)",
        "Shared CSS. Mobile-friendly. Warm professional colors.",
        "Explain the folder structure to me like I am new.",
      ]),
    ],
    [
      b("Finish portfolio site with your real bio and one project idea.", "bullets7"),
      b("Self-quiz: define frontend, backend, database, deploy in one sentence each.", "bullets7"),
    ],
    [
      code(["explain this project folder to me like I am completely new to coding"]),
    ],
    8
  ));

  // Module 4
  children.push(...moduleBlock(
    4, "Interactive Frontend — Pages That Do Things",
    [
      "Know what HTML, CSS, and JavaScript each do (concept only).",
      "Build an interactive browser app.",
      "Use a verification ritual every time.",
    ],
    [
      "Ship a quiz or similar interactive tool.",
      "Write bug reports as symptoms, not panic.",
    ],
    [
      p("Concepts (fill in):"),
      line("HTML is the ________ of a page."),
      line("CSS controls the ________."),
      line("JavaScript controls the ________ (what happens on click)."),
      sp(60),
      p("Verification ritual (always):"),
      n("Open in browser.", "numbers9"),
      n("Click every button.", "numbers9"),
      n("Try wrong inputs.", "numbers9"),
      n("Write: When I …, I expected …, but I saw …", "numbers9"),
      n("Paste to Claude and ask for root cause fix.", "numbers9"),
      sp(60),
      p("Lab:"),
      code([
        "Build a browser quiz app:",
        "- 5 multiple-choice questions on [YOUR TOPIC]",
        "- One question at a time",
        "- Score at the end + restart",
        "- Works on mobile width",
        "Then list manual test cases I should try.",
      ]),
    ],
    [
      b("Finish quiz app + write 5 manual tests you actually tried.", "bullets10"),
      b("Note one bug you found and how it was fixed.", "bullets10"),
    ],
    [
      code([
        "Bug report:",
        "Steps I took:",
        "Expected:",
        "Actual:",
        "Please find root cause, fix it, and explain simply.",
      ]),
    ],
    11
  ));

  // Module 5
  children.push(...moduleBlock(
    5, "Think Before You Build — PRD & Scope",
    [
      "Write a one-page PRD for your capstone.",
      "Cut scope to a real MVP (3–5 must features).",
    ],
    [
      "Turn PRD into a plan-mode prompt for Claude.",
      "List out-of-scope ideas in a backlog.",
    ],
    [
      p("Capstone PRD — fill this carefully:"),
      table(["Section", "Your answer"], [
        ["App name", ""],
        ["Problem (who struggles with what?)", ""],
        ["Primary user", ""],
        ["One-sentence job-to-be-done", ""],
        ["Must-have feature 1", ""],
        ["Must-have feature 2", ""],
        ["Must-have feature 3", ""],
        ["Must-have feature 4 (optional)", ""],
        ["Must-have feature 5 (optional)", ""],
        ["Out of scope for v1", ""],
        ["Screens / pages", ""],
        ["Data we store", ""],
        ["60-second success demo", ""],
      ], [3600, 6624]),
      sp(80),
      p("Dream features dump (then star only top 4):"),
      line("1."), line("2."), line("3."), line("4."), line("5."), line("6."),
    ],
    [
      b("Submit PRD for instructor feedback before Module 6.", "bullets12"),
      b("Draw a paper wireframe of your main 3 screens.", "bullets12"),
    ],
    [
      code([
        "/plan Read this PRD and propose a simple architecture for a beginner",
        "full-stack web app. List milestones M1–M4. Do not write code yet.",
        "Ask clarifying questions if anything is ambiguous.",
        "",
        "[PASTE YOUR PRD HERE]",
      ]),
    ],
    13
  ));

  // Module 6
  children.push(...moduleBlock(
    6, "Data & Backend Basics",
    [
      "Explain persistence and CRUD.",
      "Build an app that keeps data after refresh.",
      "Know why secrets (.env) must not be shared publicly.",
    ],
    [
      "Ship a notes or todo app with create/list/edit/delete.",
      "Explain where data is stored in plain English.",
    ],
    [
      p("CRUD means:"),
      table(["Letter", "Word", "Example for notes"], [
        ["C", "Create", "Add a new note"],
        ["R", "Read", "See list of notes"],
        ["U", "Update", "Edit a note"],
        ["D", "Delete", "Remove a note"],
      ], [1200, 2400, 6624]),
      sp(80),
      code([
        "Build a personal notes app:",
        "- List, create, edit, delete notes",
        "- Search by title",
        "- Data must persist after refresh",
        "Use the simplest approach that works for a beginner.",
        "Explain where data is stored in plain English.",
        "Add a README with how to run the app.",
        "After each feature, pause for my approval.",
      ]),
    ],
    [
      b("Finish notes/todo with persistence.", "bullets14"),
      b("Write 3 sentences: where data lives + how you tested save/reload.", "bullets14"),
    ],
    [
      code(["Never put passwords or API keys in public GitHub. Use .env and .gitignore."]),
    ],
    15
  ));

  // Module 7
  children.push(...moduleBlock(
    7, "Full-Stack MVP with Claude Code",
    [
      "Build your capstone MVP one milestone at a time.",
      "Create a short CLAUDE.md so Claude remembers project rules.",
    ],
    [
      "Core user flow works locally (create + list at minimum).",
      "You can start the app without the instructor typing for you.",
    ],
    [
      p("Milestone order (do not skip):"),
      n("Empty project runs (hello page / server).", "numbers16"),
      n("One data model (example: Task).", "numbers16"),
      n("Create + list only.", "numbers16"),
      n("Update + delete.", "numbers16"),
      n("Basic styling.", "numbers16"),
      n("Only later: login extras.", "numbers16"),
      sp(60),
      p("Session pattern:"),
      code([
        "/plan Build milestone 1 only: project scaffold + run instructions",
        "+ empty home page. Stop after plan for my approval.",
      ]),
      sp(40),
      code([
        "Implement the approved plan. Run the app and show exact start commands.",
        "Do not start milestone 2.",
      ]),
      sp(60),
      p("After /init, keep CLAUDE.md short. Example:"),
      code([
        "# Project",
        "Beginner full-stack app: [name]. Users: [who].",
        "# Commands",
        "- Install: ...",
        "- Run: ...",
        "# Rules",
        "- Prefer simple readable code",
        "- Never commit .env",
        "- After features: update README",
      ]),
    ],
    [
      b("Record a 3-minute unlisted video of your MVP happy path.", "bullets17"),
      b("README with run steps exists.", "bullets17"),
    ],
    [
      code(["what files changed? summarize in plain English before we commit"]),
    ],
    18
  ));

  // Module 8
  children.push(...moduleBlock(
    8, "Users, Login & Polish",
    [
      "Explain signup, login, logout, and “my data only.”",
      "Add auth or a serious polish pass to your app.",
    ],
    [
      "Test with two accounts (or document single-user scope honestly).",
      "Improve empty states, errors, and mobile readability.",
    ],
    [
      p("Auth words:"),
      table(["Term", "Plain meaning"], [
        ["Sign up", "Create an account"],
        ["Login", "Prove who you are"],
        ["Session", "App remembers you for a while"],
        ["Logout", "End this browser session"],
        ["Authorization", "What you are allowed to see/do"],
      ], [2800, 7424]),
      sp(60),
      code([
        "/plan Add authentication: sign up, login, logout,",
        "protect private pages, each user only sees their own records.",
        "Propose the simplest safe approach for our stack.",
        "No code until I approve.",
      ]),
      sp(60),
      p("Polish checklist — tick:"),
      b("Loading text while saving", "bullets19"),
      b("Empty state (“No items yet”)", "bullets19"),
      b("Form validation messages", "bullets19"),
      b("Readable on phone width", "bullets19"),
      b("Page titles make sense", "bullets19"),
    ],
    [
      b("Auth checkpoint OR polish checklist completed and signed off.", "bullets20"),
    ],
    [
      code(["Improve layout for mobile. Do not change business logic. Summarize what you changed."]),
    ],
    21
  ));

  // Module 9
  children.push(...moduleBlock(
    9, "Git & GitHub — Save Points",
    [
      "Explain commit, repo, push in plain language.",
      "Put your project on GitHub with a good README.",
    ],
    [
      "Create commits with clear messages.",
      "Confirm .env is not uploaded.",
    ],
    [
      p("Metaphors:"),
      b("Git = time machine / save points for your project folder", "bullets22"),
      b("Commit = labeled save", "bullets22"),
      b("GitHub = cloud backup + portfolio", "bullets22"),
      b("Push = upload saves to GitHub", "bullets22"),
      sp(60),
      code([
        "Show me what files changed.",
        "Commit with a clear message summarizing the feature.",
        "Help me create a GitHub repo and push (guide me through login if needed).",
        "Confirm .env is not tracked.",
      ]),
      sp(60),
      p("README must include:"),
      b("What the app does", "bullets23"),
      b("How to run locally", "bullets23"),
      b("Tech stack names (even if AI wrote the code)", "bullets23"),
      b("Your name", "bullets23"),
      b("Later: live URL", "bullets23"),
    ],
    [
      b("Public GitHub repo link: _______________________________", "bullets24"),
      b("README complete.", "bullets24"),
    ],
    [
      code(["Write a beginner-friendly README with overview, features, setup, and run steps."]),
    ],
    25
  ));

  // Module 10
  children.push(...moduleBlock(
    10, "Deploy, Test & Demo Day",
    [
      "Publish your app to a public URL.",
      "Run a manual test script.",
      "Present a 5-minute demo.",
    ],
    [
      "Someone else can open your link and complete the main flow.",
      "You can explain problem, flow, and 3 decisions you made.",
    ],
    [
      p("Deploy checklist:"),
      n("App runs cleanly on my laptop.", "numbers26"),
      n("Environment variables set on the host (not only laptop).", "numbers26"),
      n("Smoke test on phone + laptop.", "numbers26"),
      n("Two accounts tested if multi-user.", "numbers26"),
      n("README includes live URL.", "numbers26"),
      sp(60),
      p("My manual test script:"),
      line("1."), line("2."), line("3."), line("4."), line("5."), line("6."), line("7."),
      sp(60),
      p("Demo outline (5 minutes):"),
      table(["Time", "What I say / show"], [
        ["0:30", "Problem"],
        ["2–3 min", "Live walkthrough"],
        ["1 min", "What AI did vs what I decided"],
        ["0:30", "What I’d build next"],
      ], [2000, 8224]),
    ],
    [
      b("Live URL: ________________________________", "bullets27"),
      b("Demo rehearsed once out loud.", "bullets27"),
      b("Backup: screen recording if internet fails.", "bullets27"),
    ],
    [
      code(["Create a manual test checklist for a non-technical user of this app."]),
    ],
    28
  ));

  // Tools section
  children.push(
    h1("Tools You Will Use Every Day"),
    h2("Essential Claude Code commands"),
    table(["Type this", "When"], [
      ["/help", "I am lost"],
      ["/plan", "Before a big change"],
      ["/clear", "Starting a new task"],
      ["/compact", "Chat got very long"],
      ["/init", "New project — create CLAUDE.md"],
      ["/doctor", "Setup seems broken"],
      ["/login", "Need to sign in again"],
      ["/rewind", "I want to undo a bad path"],
    ], [2800, 7424]),
    sp(120),
    h2("3-minute stuck protocol"),
    n("Say out loud what you tried and what you expected.", "numbers29"),
    n("Copy the exact error text (if any).", "numbers29"),
    n("Paste to Claude using the bug report template.", "numbers29"),
    n("If still stuck after one careful try → ask a human (classmate/TA).", "numbers29"),
    sp(80),
    callout("Bug report template", [
      "Steps I took:",
      "Expected:",
      "Actual:",
      "Error text:",
      "Please find the root cause, fix it, and explain simply for a beginner.",
    ], C.softB, C.accent),
    sp(120),
    h2("Daily build checklist"),
    n("I know which folder I am in.", "numbers30"),
    n("I started Claude from that folder.", "numbers30"),
    n("My prompt has Purpose + Musts + Proof.", "numbers30"),
    n("Big work → plan mode first.", "numbers30"),
    n("I opened/ran the app myself.", "numbers30"),
    n("I tried at least one edge case.", "numbers30"),
    n("If it works, I saved a Git commit.", "numbers30"),
    n("I wrote one line: what I learned today.", "numbers30"),
    pb(),
  );

  // Prompt library
  children.push(
    h1("Prompt Pocket Library"),
    h3("Understand a project"),
    code(["what does this project do?", "explain the folder structure like I am new", "what should I run to start the app?"]),
    sp(80),
    h3("Plan"),
    code(["/plan [goal]. Ask clarifying questions. List risks. Do not edit files until I approve."]),
    sp(80),
    h3("Small implement"),
    code(["Implement only step 1 of the approved plan. Then stop and show me how to verify."]),
    sp(80),
    h3("UI polish"),
    code(["Improve spacing and mobile layout. Keep the same colors. Do not change logic. Summarize changes."]),
    sp(80),
    h3("Git"),
    code(["Summarize uncommitted changes in plain English.", "Commit with a descriptive message. Do not push yet."]),
    sp(80),
    h3("Docs"),
    code(["Write a beginner-friendly README with overview, features, setup, env vars, and run steps."]),
    pb(),
  );

  // Glossary
  children.push(
    h1("Beginner Glossary"),
    p("Revisit weekly. Put a ✓ when you can explain it without looking."),
    table(["✓", "Term", "Plain meaning"], [
      ["", "Terminal / CLI", "Text window to control the computer"],
      ["", "Claude Code", "AI that edits files and runs tools in your project"],
      ["", "Prompt", "Instructions you give the AI"],
      ["", "Bug", "Something that does not work as intended"],
      ["", "Frontend", "What the user sees and clicks"],
      ["", "Backend", "Hidden server logic"],
      ["", "Database", "Long-term stored data"],
      ["", "API", "Messages between systems"],
      ["", "Full-stack", "Frontend + backend + data"],
      ["", "Deploy", "Publish so others can use it online"],
      ["", "Localhost", "Your computer acting as server while building"],
      ["", "Repository", "Project tracked by Git"],
      ["", "Commit", "A labeled save point"],
      ["", "Push", "Upload commits to GitHub"],
      ["", "Environment variable", "Secret/config stored outside code"],
      ["", "CRUD", "Create, Read, Update, Delete"],
      ["", "MVP", "Smallest useful product version"],
      ["", "PRD", "Product requirements one-pager"],
      ["", "Authentication", "Proving who you are"],
      ["", "CLAUDE.md", "Project rules Claude reads every session"],
      ["", "Plan mode", "Claude plans without editing until you approve"],
      ["", "Context window", "How much conversation the AI can hold"],
    ], [600, 2800, 6824]),
    pb(),
  );

  // Capstone
  children.push(
    h1("Capstone Tracker"),
    p("Mission: design and ship a small full-stack web app that solves a real problem. Claude writes code. You own clarity, testing, and the demo."),
    h2("Minimum features"),
    b("Home/landing that explains the app"),
    b("Create / read / update / delete for main records"),
    b("Data persists"),
    b("Basic validation messages"),
    b("GitHub repo + README"),
    b("Live deployed URL"),
    b("Optional: user accounts"),
    h2("Timeline checkboxes"),
    table(["When", "Deliverable", "Done?"], [
      ["Week 5", "PRD approved", ""],
      ["Week 7", "MVP core flow local", ""],
      ["Week 8", "Auth/polish checkpoint", ""],
      ["Week 9", "GitHub complete", ""],
      ["Week 10", "Deployed + demo", ""],
    ], [2000, 5200, 3024]),
    sp(120),
    h2("Links (fill as you go)"),
    line("GitHub repo:"),
    line("Live URL:"),
    line("Demo video (backup):"),
    sp(120),
    h2("Three decisions I made (for demo)"),
    line("1."),
    line("2."),
    line("3."),
    sp(120),
    callout("Integrity", [
      "Using Claude Code is required and expected.",
      "Submitting an app you cannot explain is not success.",
      "In your demo, you must describe the problem, the flow, and decisions you made.",
    ], C.softY, C.warn),
    pb(),
  );

  // Homework log
  children.push(
    h1("Homework Log"),
    table(["Week", "Assignment", "Submitted?"], [
      ["1", "Custom Hello page + 3-sentence reflection + claude --version screenshot", ""],
      ["2", "Prompt journal (5 rewrites) + study timer/tip tool", ""],
      ["3", "4-page portfolio + glossary self-quiz", ""],
      ["4", "Interactive quiz + 5 manual tests", ""],
      ["5", "Capstone PRD v1", ""],
      ["6", "Notes/todo with persistence + data explanation", ""],
      ["7", "MVP local demo video (3 min)", ""],
      ["8", "Auth or polish checkpoint", ""],
      ["9", "Public GitHub + README", ""],
      ["10", "Live URL + demo day", ""],
    ], [1200, 7000, 2024]),
    pb(),
  );

  // Closing
  children.push(
    h1("You Can Build Software"),
    p("At the start of this course, the terminal may have looked scary. By the end, you will have directed an AI engineer, shipped an app, and shown it live."),
    p("Keep the habits: clear prompts, small milestones, plan before big edits, verify everything, save often, never publish secrets."),
    sp(120),
    callout("Your north star", [
      "Idea → PRD → Claude Code sessions → working app → public URL → clear demo.",
      "If you can do that, you are a builder.",
    ]),
    sp(200),
    p("— End of Student Workbook —", { align: AlignmentType.CENTER, bold: true, color: C.primary }),
  );

  const doc = new Document({
    styles: {
      default: { document: { run: { font: "Arial", size: 21 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 30, bold: true, font: "Arial", color: C.primary },
          paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 24, bold: true, font: "Arial", color: C.accent },
          paragraph: { spacing: { before: 220, after: 120 }, outlineLevel: 1 } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 21, bold: true, font: "Arial", color: C.dark },
          paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } },
      ],
    },
    numbering: { config: numberingConfig() },
    sections: [{
      properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: M, right: M, bottom: M, left: M } } },
      headers: {
        default: new Header({ children: [new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.primary, space: 4 } },
          children: [
            new TextRun({ text: "Student Workbook", font: "Arial", size: 15, bold: true, color: C.primary }),
            new TextRun({ text: "  ·  Vibe Coding with Claude Code", font: "Arial", size: 15, color: C.muted }),
          ],
        })] }),
      },
      footers: {
        default: new Footer({ children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.border, space: 6 } },
          children: [
            new TextRun({ text: "Page ", font: "Arial", size: 14, color: C.muted }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 14, color: C.muted }),
            new TextRun({ text: "  ·  Your notes matter — fill every blank you can", font: "Arial", size: 14, color: C.muted }),
          ],
        })] }),
      },
      children,
    }],
  });

  const out = "/home/shailesh/Downloads/Training/vibe-coding/Vibe_Coding_Student_Workbook.docx";
  fs.writeFileSync(out, await Packer.toBuffer(doc));
  console.log("Wrote", out);
}

main().catch((e) => { console.error(e); process.exit(1); });
