/**
 * Vibe Coding with Claude Code — Full Course Guide Generator
 * Complete instructor + student curriculum for non-technical learners
 */
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
        BorderStyle, WidthType, ShadingType, PageNumber, PageBreak,
        TableOfContents } = require('docx');
const fs = require('fs');

// ── Design tokens ──────────────────────────────────────────
const COLORS = {
  primary: "1B4F72",
  accent: "2874A6",
  dark: "1C2833",
  muted: "5D6D7E",
  lightBg: "EBF5FB",
  altBg: "F4F6F7",
  white: "FFFFFF",
  border: "D5D8DC",
  success: "196F3D",
  warn: "B7950B",
  danger: "922B21",
  softGreen: "E8F8F5",
  softYellow: "FEF9E7",
  softRed: "FDEDEC",
  softPurple: "F5EEF8",
};

const PAGE_W = 12240;
const PAGE_H = 15840;
const MARGIN = 1008; // 0.7"
const CONTENT_W = PAGE_W - MARGIN * 2; // 10224

const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: COLORS.border };
const borders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ── Helpers ────────────────────────────────────────────────
const p = (text, opts = {}) => new Paragraph({
  spacing: { after: opts.after ?? 140, before: opts.before ?? 0, line: opts.line ?? 276 },
  alignment: opts.align || AlignmentType.LEFT,
  ...opts.para,
  children: [new TextRun({
    text,
    font: "Arial",
    size: opts.size || 21,
    bold: opts.bold || false,
    italics: opts.italics || false,
    color: opts.color || COLORS.dark,
  })],
});

const multi = (runs, opts = {}) => new Paragraph({
  spacing: { after: opts.after ?? 140, before: opts.before ?? 0, line: 276 },
  alignment: opts.align || AlignmentType.LEFT,
  children: runs.map(r => new TextRun({
    text: r.text,
    font: "Arial",
    size: r.size || 21,
    bold: r.bold || false,
    italics: r.italics || false,
    color: r.color || COLORS.dark,
  })),
});

const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 360, after: 200 },
  children: [new TextRun({ text, font: "Arial", size: 32, bold: true, color: COLORS.primary })],
});

const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 280, after: 140 },
  children: [new TextRun({ text, font: "Arial", size: 26, bold: true, color: COLORS.accent })],
});

const h3 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  spacing: { before: 200, after: 100 },
  children: [new TextRun({ text, font: "Arial", size: 22, bold: true, color: COLORS.dark })],
});

const bullet = (text, ref = "bullets") => new Paragraph({
  numbering: { reference: ref, level: 0 },
  spacing: { after: 80, line: 276 },
  children: [new TextRun({ text, font: "Arial", size: 21, color: COLORS.dark })],
});

const bulletBold = (label, rest, ref = "bullets") => new Paragraph({
  numbering: { reference: ref, level: 0 },
  spacing: { after: 80, line: 276 },
  children: [
    new TextRun({ text: label, font: "Arial", size: 21, bold: true, color: COLORS.dark }),
    new TextRun({ text: rest, font: "Arial", size: 21, color: COLORS.dark }),
  ],
});

const num = (text, ref = "numbers") => new Paragraph({
  numbering: { reference: ref, level: 0 },
  spacing: { after: 80, line: 276 },
  children: [new TextRun({ text, font: "Arial", size: 21, color: COLORS.dark })],
});

const numBold = (label, rest, ref = "numbers") => new Paragraph({
  numbering: { reference: ref, level: 0 },
  spacing: { after: 80, line: 276 },
  children: [
    new TextRun({ text: label, font: "Arial", size: 21, bold: true, color: COLORS.dark }),
    new TextRun({ text: rest, font: "Arial", size: 21, color: COLORS.dark }),
  ],
});

const callout = (title, body, bg = COLORS.lightBg, titleColor = COLORS.primary) => {
  const rows = [
    new TableRow({
      children: [
        new TableCell({
          borders: noBorders,
          width: { size: CONTENT_W, type: WidthType.DXA },
          shading: { fill: bg, type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 60, left: 160, right: 160 },
          children: [
            new Paragraph({
              spacing: { after: 60 },
              children: [new TextRun({ text: title, font: "Arial", size: 20, bold: true, color: titleColor })],
            }),
          ],
        }),
      ],
    }),
  ];
  const lines = Array.isArray(body) ? body : [body];
  lines.forEach((line, i) => {
    rows.push(new TableRow({
      children: [
        new TableCell({
          borders: noBorders,
          width: { size: CONTENT_W, type: WidthType.DXA },
          shading: { fill: bg, type: ShadingType.CLEAR },
          margins: { top: 20, bottom: i === lines.length - 1 ? 100 : 20, left: 160, right: 160 },
          children: [
            new Paragraph({
              spacing: { after: 40 },
              children: [new TextRun({ text: line, font: "Arial", size: 19, color: COLORS.dark })],
            }),
          ],
        }),
      ],
    }));
  });
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows,
  });
};

const spacer = (after = 120) => new Paragraph({ spacing: { after }, children: [] });

const codeBlock = (lines) => {
  const arr = Array.isArray(lines) ? lines : lines.split("\n");
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: arr.map((line, i) => new TableRow({
      children: [
        new TableCell({
          borders: noBorders,
          width: { size: CONTENT_W, type: WidthType.DXA },
          shading: { fill: "2C3E50", type: ShadingType.CLEAR },
          margins: { top: i === 0 ? 80 : 20, bottom: i === arr.length - 1 ? 80 : 20, left: 140, right: 140 },
          children: [
            new Paragraph({
              spacing: { after: 0 },
              children: [new TextRun({ text: line || " ", font: "Consolas", size: 17, color: "ECF0F1" })],
            }),
          ],
        }),
      ],
    })),
  });
};

const cell = (text, opts = {}) => new TableCell({
  borders,
  width: { size: opts.w, type: WidthType.DXA },
  shading: { fill: opts.fill || COLORS.white, type: ShadingType.CLEAR },
  margins: { top: 60, bottom: 60, left: 80, right: 80 },
  verticalAlign: "center",
  children: [
    new Paragraph({
      children: [new TextRun({
        text,
        font: "Arial",
        size: opts.size || 18,
        bold: opts.bold || false,
        color: opts.color || COLORS.dark,
      })],
    }),
  ],
});

const simpleTable = (headers, rows, colWidths) => {
  const total = colWidths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({
        children: headers.map((h, i) => cell(h, {
          w: colWidths[i], fill: COLORS.primary, bold: true, color: COLORS.white, size: 18,
        })),
      }),
      ...rows.map((row, ri) => new TableRow({
        children: row.map((c, i) => cell(c, {
          w: colWidths[i],
          fill: ri % 2 === 0 ? COLORS.altBg : COLORS.white,
          size: 17,
        })),
      })),
    ],
  });
};

const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

// ── Numbering configs ──────────────────────────────────────
const numberingConfig = [
  { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets2", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets3", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets4", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets5", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets6", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets7", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets8", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets9", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets10", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets11", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets12", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets13", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets14", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets15", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets16", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets17", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets18", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets19", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets20", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets21", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets22", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets23", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets24", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "bullets25", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers2", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers3", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers4", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers5", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers6", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers7", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers8", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers9", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers10", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers11", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers12", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers13", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers14", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers15", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers16", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers17", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers18", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers19", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  { reference: "numbers20", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
];

// ── MODULE CONTENT BUILDERS ────────────────────────────────
function moduleHeader(num, title, duration, goals) {
  return [
    h1(`Module ${num}: ${title}`),
    multi([
      { text: "Duration: ", bold: true },
      { text: duration },
      { text: "  |  ", color: COLORS.muted },
      { text: "Audience: ", bold: true },
      { text: "Absolute beginners (non-technical)" },
    ]),
    h3("Learning Goals"),
    ...goals.map(g => bullet(g, `bullets-m${num}`)),
    spacer(80),
  ];
}

// Need unique bullet refs per module - I'll use shared bullets and accept continuous numbering isn't ideal - actually each reference is independent. I'll use bullets with different refs embedded in module builders.

function buildChildren() {
  const c = [];

  // ═══════════════ COVER ═══════════════
  c.push(
    spacer(600),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: "COMPLETE COURSE GUIDE", font: "Arial", size: 22, bold: true, color: COLORS.accent })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 160 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: COLORS.primary, space: 8 } },
      children: [new TextRun({ text: "Vibe Coding with Claude Code", font: "Arial", size: 48, bold: true, color: COLORS.primary })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 120 },
      children: [new TextRun({ text: "From Zero Tech Knowledge to Building Full-Stack Applications", font: "Arial", size: 26, color: COLORS.dark })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
      children: [new TextRun({ text: "A Step-by-Step Teaching Curriculum for Non-Technical Students", font: "Arial", size: 22, italics: true, color: COLORS.muted })],
    }),
    spacer(400),
    callout("What This Document Is", [
      "This is a complete instructor + student course guide. It includes: learning philosophy, non-tech gap analysis,",
      "12 progressive modules, lesson-by-lesson scripts, lab exercises, prompt libraries, assessments, glossaries,",
      "troubleshooting, and a full-stack capstone. Designed so you can teach it live or run it as a cohort.",
    ]),
    spacer(200),
    p("Document version: 1.0  |  Research-backed (Claude Code official docs, 2026)  |  Focus: Claude Code CLI", { size: 18, color: COLORS.muted, align: AlignmentType.CENTER }),
    p("Primary tool: Claude Code (Anthropic) terminal CLI  |  Secondary surfaces: Desktop app, VS Code extension optional", { size: 18, color: COLORS.muted, align: AlignmentType.CENTER }),
    pageBreak(),
  );

  // ═══════════════ TOC ═══════════════
  c.push(
    h1("Table of Contents"),
    new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-2" }),
    pageBreak(),
  );

  // ═══════════════ PART A: COURSE DESIGN ═══════════════
  c.push(
    h1("Part A — Course Design & Teaching Philosophy"),
    h2("1. Course Snapshot"),
    simpleTable(
      ["Attribute", "Detail"],
      [
        ["Course name", "Vibe Coding with Claude Code: Build Full-Stack Apps Without Being a Coder"],
        ["Target students", "Non-technical adults/students with zero programming background"],
        ["Duration", "8–12 weeks (recommended: 10 weeks × 2 sessions/week + labs)"],
        ["Session length", "90–120 minutes live + 2–4 hours practice between sessions"],
        ["Total effort", "~60–80 hours including projects"],
        ["Primary tool", "Claude Code CLI (claude command in the terminal)"],
        ["Accounts needed", "Claude Pro/Max (or Team), GitHub free, browser"],
        ["Hardware", "Laptop, 8GB+ RAM recommended, stable internet"],
        ["Outcome", "Each student ships a working full-stack web app they designed and directed"],
        ["Teaching style", "Slow, metaphor-first, demo-then-do, one concept at a time"],
      ],
      [2800, 7424]
    ),
    spacer(160),

    h2("2. What Is “Vibe Coding”? (Instructor Framing)"),
    p("Vibe coding is a way of building software by describing what you want in plain language while an AI agent writes, edits, runs, and debugs the code. You act as the director; Claude Code acts as a tireless junior builder who can touch files, run commands, and iterate."),
    p("Important distinction for students:"),
    bulletBold("Vibe coding (this course): ", "You describe outcomes; AI implements. You review results, steer, and verify."),
    bulletBold("Traditional coding: ", "You write every line of syntax yourself."),
    bulletBold("AI-assisted engineering: ", "You still own architecture deeply; AI speeds you up. We introduce healthy habits so students do not become “prompt-only” with zero understanding."),
    spacer(80),
    callout("Teaching Principle", [
      "Never tell beginners “you don’t need to understand anything.” That creates fragile builders.",
      "Instead: “You don’t need to memorize syntax. You DO need to understand the parts of an app,",
      "how to give clear instructions, how to verify work, and how to fix problems when AI gets stuck.”",
    ], COLORS.softYellow, COLORS.warn),
    spacer(120),

    h2("3. Why Claude Code CLI (Not Only ChatGPT)?"),
    bullet("Claude Code lives in the project folder — it can read many files, edit them, run terminal commands, and test."),
    bullet("It is agentic: it can plan multi-step work, not just paste one code snippet."),
    bullet("Official workflow is Explore → Plan → Implement → Commit (from Anthropic best practices)."),
    bullet("Features non-tech students will grow into: plan mode, CLAUDE.md memory, /init, permissions, skills, Git help."),
    bullet("Alternative surfaces (desktop app, VS Code extension, web) exist — mention them, but master CLI so students understand the real development loop."),
    spacer(80),
    callout("Account Reality Check", [
      "Claude Code requires a paid Claude plan (Pro, Max, Team, Enterprise) or API Console access — free Claude.ai alone is not enough.",
      "Budget this for students or provide lab accounts. Official docs: code.claude.com / docs.anthropic.com Claude Code setup.",
    ], COLORS.softRed, COLORS.danger),
    spacer(120),

    h2("4. Non-Technical Learner Gap Analysis"),
    p("Non-tech students fail AI coding courses when instructors skip foundational literacy. Design every week to close these gaps:"),
    spacer(80),
    simpleTable(
      ["Gap", "What Students Feel", "How This Course Closes It"],
      [
        ["File & folder mental model", "“Where did my project go?”", "Module 1: Desktop → folder → open in terminal, visual maps"],
        ["Terminal fear", "Black screen = danger", "Terminal as “text remote control”; copy-paste safe commands first"],
        ["Vocabulary wall", "Frontend/backend/API sound foreign", "Restaurant metaphor + glossary + spaced repetition"],
        ["Vague prompting", "“Make a nice app” → garbage results", "PROMPT formula + PRD templates + before/after examples"],
        ["No verification habit", "“Claude said it’s done”", "Always run it / open browser / click every button"],
        ["Error panic", "Red text = failure / shame", "Errors are clues; paste to Claude; calm protocol"],
        ["No product thinking", "Jump to code with no idea", "User, problem, screens, acceptance checks first"],
        ["Security naivety", "Paste API keys in public chat", "Secrets module: .env, never commit keys, phishing basics"],
        ["Git confusion", "“What is a commit?”", "Save points analogy + Claude does git with English"],
        ["Over-trust of AI", "Blind accept all changes", "Review diffs, plan mode, small steps, /rewind"],
        ["Context overload", "Long chats get worse", "One feature per session; /clear; /compact; CLAUDE.md"],
        ["Deployment mystery", "It works on my laptop only", "Deploy module: public URL, env vars, smoke test"],
      ],
      [2200, 3200, 4824]
    ),
    spacer(160),

    h2("5. Learning Outcomes (By End of Course)"),
    p("A successful graduate can:"),
    num("Open a terminal, navigate to a project folder, and start Claude Code confidently."),
    num("Explain in plain language what frontend, backend, database, API, and deployment mean."),
    num("Write a short Product Requirements Document (PRD) and turn it into build prompts."),
    num("Use plan mode for non-trivial features before allowing code changes."),
    num("Direct Claude to build UI, forms, data storage, and basic authentication for a web app."),
    num("Use Git + GitHub at a basic level (commit, push, explain what a repository is)."),
    num("Debug by reading errors, describing symptoms, and asking Claude to fix root causes."),
    num("Create a CLAUDE.md file so the AI remembers project rules."),
    num("Deploy an app to a public URL and share it."),
    num("Present a capstone full-stack application they directed end-to-end."),
    spacer(120),

    h2("6. Recommended Schedule"),
    simpleTable(
      ["Week", "Module Focus", "Student Delivers"],
      [
        ["0 (pre)", "Accounts, laptop check, mindset", "Accounts ready checklist"],
        ["1", "Digital literacy + terminal + first Claude session", "Hello webpage"],
        ["2", "Vibe coding craft + prompting + plan mode", "Prompt journal + mini site"],
        ["3", "How the web works (concepts) + static multi-page site", "Personal portfolio site"],
        ["4", "Interactive frontend (forms, JS concepts via AI)", "Interactive quiz or calculator app"],
        ["5", "Product thinking + PRD + wireframes", "PRD for capstone idea"],
        ["6", "Backend & data (concepts) + first API/data app", "Simple notes or todo with storage"],
        ["7", "Full-stack assembly (frontend + backend + DB)", "MVP of capstone (core flow)"],
        ["8", "Auth, polish, permissions, CLAUDE.md", "Login + protected feature"],
        ["9", "Git/GitHub + collaboration + code review habits", "Repo online + README"],
        ["10", "Deploy + testing + demo day", "Live URL + 5-min demo"],
      ],
      [1200, 4500, 4524]
    ),
    spacer(80),
    p("Compress to 8 weeks by combining weeks 4–5 and 8–9 for intensive cohorts. Expand to 12 weeks with extra lab days for weaker computer literacy groups.", { italics: true, color: COLORS.muted }),
    spacer(120),

    h2("7. Classroom Setup Checklist"),
    h3("Before Day 1 (Instructor)"),
    bullet("Confirm every student has a laptop (Windows 10+, macOS 13+, or Ubuntu 20.04+) with admin rights for installs."),
    bullet("Claude Pro/Max (or Team) seats arranged — free plan does not include Claude Code."),
    bullet("GitHub accounts created with 2FA encouraged."),
    bullet("Shared course folder: slides, prompt cheatsheet, glossary PDF, help desk channel (WhatsApp/Slack/Discord)."),
    bullet("Instructor machine: Claude Code installed, demo project pre-built for “when live demo fails.”"),
    bullet("Seating: pair strong computer users with complete beginners when possible."),
    h3("Room / Remote"),
    bullet("Projector or screen share for terminal demos (large font, 16–18pt minimum)."),
    bullet("Co-host/TA for 1:1 install help during Module 1 (critical!)."),
    bullet("Record sessions for students who get stuck on setup."),
    spacer(80),
    callout("Golden Classroom Rule", [
      "Never leave a student alone with a red error for more than 3 minutes in early modules.",
      "Early wins build trust; early shame kills the course. Normalize mistakes as normal engineering life.",
    ], COLORS.softGreen, COLORS.success),
    pageBreak(),
  );

  // ═══════════════ PART B: PRE-COURSE ═══════════════
  c.push(
    h1("Part B — Pre-Course Onboarding (Week 0)"),
    h2("Goals"),
    p("Remove setup friction before teaching concepts. If Module 1 is spent only on passwords and installs, learning energy dies."),
    h2("Student Pre-Work (Send 5–7 Days Ahead)"),
    num("Create email account dedicated or personal — write password in a password manager or sealed note (teach not to reuse passwords)."),
    num("Create Claude account and upgrade to Pro/Max as instructed by the school."),
    num("Create free GitHub account (github.com) — username should be professional."),
    num("Update OS if possible; free disk space ≥ 10 GB."),
    num("Install Chrome or Edge (latest)."),
    num("Watch a 5-minute “What is a terminal?” intro video provided by instructor (no coding yet)."),
    num("Fill a one-page survey: OS type, “I have used terminal before Y/N”, app idea interests."),
    h2("Instructor Verification Call / Lab Hour"),
    bullet("Confirm OS version meets Claude Code requirements (macOS 13+, Win 10 1809+, Ubuntu 20.04+)."),
    bullet("4 GB RAM minimum (8 GB better). Internet required."),
    bullet("For Windows: recommend Git for Windows install before class."),
    bullet("Print/PDF the “Do Not Panic” card (Appendix)."),
    pageBreak(),
  );

  // ═══════════════ MODULE 1 ═══════════════
  c.push(
    h1("Module 1 — Digital Ground Zero: Computer, Files, Terminal"),
    multi([{ text: "Duration: ", bold: true }, { text: "1–2 sessions (critical foundation)" }]),
    h2("Why This Module Exists"),
    p("Non-tech students often cannot complete AI coding courses because they cannot navigate folders, run a terminal, or understand where files live. Claude Code’s own docs include a Terminal Guide for new users — we expand that into pedagogy."),
    h2("Concepts (Metaphor First)"),
    bulletBold("Computer = workshop. ", "Files = materials. Folders = drawers. Terminal = a text remote control for the workshop."),
    bulletBold("Path = address. ", "Like /Users/you/Projects/my-app or C:\\Users\\you\\Projects\\my-app."),
    bulletBold("Command = short written instruction. ", "Example: ls (list), cd (change directory), mkdir (make folder)."),
    bulletBold("Browser vs Terminal. ", "Browser is for visiting websites; terminal is for controlling your computer with text."),
    h2("Lesson Flow"),
    h3("1.1 Files & Folders Lab (20 min)"),
    num("Create folder: Documents → vibe-coding-course → week1-hello."),
    num("Create a text file hello.txt with one sentence."),
    num("Rename, move, delete (to Trash), restore — practice without fear."),
    num("Show hidden extensions on Windows if needed (.txt vs .html matters later)."),
    h3("1.2 Open the Terminal (25 min)"),
    p("Platform-specific demo (instructor does first, students mirror):"),
    bulletBold("macOS: ", "Spotlight → Terminal. Or open folder in Finder → right-click → New Terminal at Folder."),
    bulletBold("Windows: ", "Win key → type PowerShell. Note: prompt shows PS C:\\ when in PowerShell."),
    bulletBold("Linux: ", "Ctrl+Alt+T or application menu → Terminal."),
    p("Safe first commands (students type or paste):"),
    codeBlock([
      "pwd          # where am I? (mac/linux)",
      "cd           # go home",
      "ls           # list files (mac/linux)   OR   dir   (Windows CMD)",
      "mkdir practice-folder",
      "cd practice-folder",
    ]),
    callout("Instructor Tip", [
      "Print a one-page “Terminal Cheat Sheet” with only 8 commands. Too many commands overwhelm.",
      "Ban scary commands (rm -rf, format, etc.) explicitly. Give a “never run this” list.",
    ]),
    spacer(100),
    h3("1.3 Install Claude Code (40 min — with TA support)"),
    p("Official recommended native install (2026 docs):"),
    p("macOS / Linux / WSL:", { bold: true }),
    codeBlock(["curl -fsSL https://claude.ai/install.sh | bash"]),
    p("Windows PowerShell:", { bold: true }),
    codeBlock(["irm https://claude.ai/install.ps1 | iex"]),
    p("Windows CMD (if not PowerShell):", { bold: true }),
    codeBlock(["curl -fsSL https://claude.ai/install.cmd -o install.cmd && install.cmd && del install.cmd"]),
    p("Also available: Homebrew (macOS), WinGet (Windows). Native install auto-updates."),
    p("Verify:"),
    codeBlock(["claude --version", "claude doctor"]),
    p("First login:"),
    codeBlock(["claude"]),
    p("Browser opens → log in with Claude subscription → return to terminal. Students should see the Claude Code welcome prompt."),
    h3("1.4 First Magic Moment (20 min)"),
    p("In an empty folder:"),
    codeBlock([
      "mkdir ~/vibe-coding-course/week1-hello",
      "cd ~/vibe-coding-course/week1-hello",
      "claude",
    ]),
    p("Then type in English:"),
    codeBlock(["make me a simple webpage that says Hello, my name is [Name] and I am learning vibe coding"]),
    p("Approve file creation when asked. Open the HTML file in a browser (double-click or ask Claude: “open this in my browser”)."),
    h2("Success Criteria"),
    bullet("Student can open terminal, cd into a folder, start claude, and generate a webpage."),
    bullet("Student can explain what a folder path is in their own words."),
    h2("Homework"),
    bullet("Customize the page: add favorite color background and a short bio paragraph by asking Claude in English."),
    bullet("Screenshot before/after; write 3 sentences: “What surprised me / what confused me / what I want next.”"),
    pageBreak(),
  );

  // ═══════════════ MODULE 2 ═══════════════
  c.push(
    h1("Module 2 — The Vibe Coding Craft: Directing Claude Code"),
    multi([{ text: "Duration: ", bold: true }, { text: "1–2 sessions" }]),
    h2("Core Idea"),
    p("Claude Code is not a magic genie. It is a powerful collaborator that does better work with clear goals, constraints, and verification. Anthropic’s best practices emphasize: manage context, explore then plan then code, give verification criteria, and be specific."),
    h2("The Official Loop (Teach Explicitly)"),
    numBold("Explore: ", "“What already exists? What files are here?”"),
    numBold("Plan: ", "Use plan mode (/plan or Shift+Tab) so Claude designs before editing."),
    numBold("Implement: ", "Allow edits; approve carefully."),
    numBold("Commit: ", "Save a checkpoint with Git (Module 9 deep dive; light version now)."),
    h2("Permission Modes (Keep Simple)"),
    bulletBold("Default: ", "Claude asks before risky actions — good for beginners."),
    bulletBold("Plan mode: ", "Read-only thinking — Claude cannot edit until you approve the plan. Use for anything bigger than a small tweak."),
    bulletBold("Auto-accept (later): ", "Only after students understand review habits."),
    h2("The PROMPT Formula for Beginners"),
    p("Teach this acronym on a poster:"),
    simpleTable(
      ["Letter", "Meaning", "Example"],
      [
        ["P", "Purpose — what outcome?", "Build a tip calculator webpage"],
        ["R", "Role / constraints", "Simple HTML/CSS/JS only, beginner-friendly code"],
        ["O", "Objects — screens, data, buttons", "Bill amount input, tip %, total display"],
        ["M", "Musts — acceptance checks", "Works on phone width; invalid input shows message"],
        ["P", "Proof — how we verify", "Open in browser; try 100 bill + 15% tip = 115"],
        ["T", "Tone / style", "Clean, large fonts, calm colors"],
      ],
      [1000, 3000, 6224]
    ),
    spacer(100),
    h3("Bad vs Good Prompt (Live Demo)"),
    p("Bad:", { bold: true }),
    codeBlock(["make a nice app"]),
    p("Good:", { bold: true }),
    codeBlock([
      "Create a single HTML file tip calculator.",
      "Inputs: bill amount (number), tip percent (5,10,15,20 buttons).",
      "Show tip amount and total.",
      "If bill is empty or negative, show a friendly error.",
      "Use large readable fonts. No frameworks.",
      "After building, tell me how to open it in the browser.",
      "Then wait for my feedback before more changes.",
    ]),
    h2("Session Hygiene (Prevent “Claude Got Dumb”)"),
    bulletBold("One feature per conversation when possible. ", "Use /clear when switching topics."),
    bulletBold("Context fills up. ", "Long chats degrade quality — official docs stress this. Teach /compact and fresh sessions."),
    bulletBold("Be specific: ", "Reference files with @filename when relevant."),
    bulletBold("Ask for evidence: ", "“Run it and show me the output” not “is it done?”"),
    h2("Essential Slash Commands (Beginner Set)"),
    simpleTable(
      ["Command", "When Students Use It"],
      [
        ["/help", "Lost — list of commands"],
        ["/init", "New project — create starter CLAUDE.md"],
        ["/plan", "Before big changes — plan mode"],
        ["/clear", "New task — reset conversation, keep project memory"],
        ["/compact", "Chat got long — summarize to free space"],
        ["/login", "Auth issues"],
        ["/doctor", "Something broken in setup"],
        ["/permissions", "Control what Claude may do"],
        ["/rewind", "Undo bad direction (checkpoint)"],
      ],
      [2800, 7424]
    ),
    h2("Lab: Prompt Upgrade Challenge"),
    num("Students write a bad one-line prompt for a “study timer” page."),
    num("Rewrite with PROMPT formula."),
    num("Build with Claude Code."),
    num("Peer review: partner tries to break the app (edge cases)."),
    num("Ask Claude to fix based on partner findings."),
    h2("Homework"),
    bullet("Prompt journal: 5 bad→good rewrites on any life problem turned into a tiny web tool."),
    bullet("Read glossary entries: frontend, backend, bug, deploy (preview for Module 3)."),
    pageBreak(),
  );

  // ═══════════════ MODULE 3 ═══════════════
  c.push(
    h1("Module 3 — How the Internet & Web Apps Work (No Syntax Torture)"),
    multi([{ text: "Duration: ", bold: true }, { text: "1–2 sessions" }]),
    h2("Goal"),
    p("Students gain a mental model of full-stack apps without memorizing code. They will later direct Claude correctly because they know the parts."),
    h2("The Restaurant Metaphor (Teach This Slide Deck)"),
    simpleTable(
      ["Restaurant", "Web App Part", "Student-Friendly Meaning"],
      [
        ["Dining room / menu", "Frontend", "What the user sees and clicks (pages, buttons, forms)"],
        ["Kitchen", "Backend", "Hidden logic that processes orders/requests"],
        ["Recipe book / pantry", "Database", "Stored information that persists"],
        ["Waiter taking order", "API / request", "Message between frontend and backend"],
        ["Restaurant address", "Domain / URL", "How people find your app on the internet"],
        ["Opening the restaurant", "Deployment", "Putting the app on a public server"],
      ],
      [2400, 2400, 5424]
    ),
    spacer(100),
    h2("Browser Developer Tools (Peek Only)"),
    bullet("Right-click → Inspect (don’t require mastery)."),
    bullet("Show Console = place error messages appear."),
    bullet("Show Network tab = requests flying between browser and servers."),
    bullet("Rule: “Red errors in Console are clues, not personal failure.”"),
    h2("What Is a Full-Stack Application?"),
    p("A full-stack app has:"),
    num("User interface (frontend)"),
    num("Server logic (backend)"),
    num("Stored data (database or file storage)"),
    num("Usually user accounts (auth) for real products"),
    num("Hosting so others can use it (deployment)"),
    p("Capstone target for this course: a small but real full-stack app (e.g., personal task manager, mini CRM for a shop, course tracker, expense log, community board)."),
    h2("Lab: Multi-Page Static Site with Claude"),
    p("Prompt starter:"),
    codeBlock([
      "Create a simple multi-page personal portfolio site with:",
      "- index.html (home)",
      "- about.html",
      "- projects.html",
      "- contact.html (form can be visual only for now)",
      "Shared CSS for consistent style.",
      "Mobile-friendly. Warm professional colors.",
      "Explain the folder structure to me like I'm new.",
    ]),
    h2("Conceptual Mini-Quiz (No Coding)"),
    bullet("If the menu looks wrong, is that frontend or backend?"),
    bullet("If saved data disappears after refresh, what might be missing?"),
    bullet("What does deploy mean in one sentence?"),
    pageBreak(),
  );

  // ═══════════════ MODULE 4 ═══════════════
  c.push(
    h1("Module 4 — Interactive Frontend: Making Pages Do Things"),
    multi([{ text: "Duration: ", bold: true }, { text: "1–2 sessions" }]),
    h2("Concepts Without Fear"),
    bulletBold("HTML: ", "Structure / skeleton (headings, buttons, inputs)."),
    bulletBold("CSS: ", "Look and layout (colors, spacing, responsive design)."),
    bulletBold("JavaScript: ", "Behavior (what happens when you click)."),
    p("Students do not write these by hand yet. They learn to name them so they can say: “The button is there but nothing happens when I click — please fix the JavaScript behavior.”"),
    h2("Lab Project: Study Quiz App"),
    p("Requirements students paste into Claude after planning:"),
    codeBlock([
      "Build a browser quiz app in one folder:",
      "- 5 multiple-choice questions on a topic I choose",
      "- Show one question at a time",
      "- Score at the end",
      "- Restart button",
      "- Works on mobile width",
      "After building: list test cases I should try manually.",
    ]),
    h2("Verification Ritual (Always)"),
    num("Open in browser."),
    num("Click every button."),
    num("Try wrong inputs (empty, spam click, back button)."),
    num("Write bugs as symptoms: “When I …, I expected …, but I saw …”"),
    num("Paste symptoms to Claude; ask for root cause fix + retest."),
    h2("UI Quality Bar for Non-Designers"),
    bullet("Readable font size (≥16px body)"),
    bullet("Contrast: dark text on light background (or intentional dark theme)"),
    bullet("Spacing: not cramped"),
    bullet("Primary button looks clickable"),
    bullet("Error messages in plain language"),
    p("Students may paste screenshots into Claude Code (images supported) and say: “Make it match this sketch / fix spacing.”"),
    pageBreak(),
  );

  // ═══════════════ MODULE 5 ═══════════════
  c.push(
    h1("Module 5 — Think Before You Build: PRD, Scope, Wireframes"),
    multi([{ text: "Duration: ", bold: true }, { text: "1 session + take-home" }]),
    h2("Why Non-Tech Builders Fail Without This"),
    p("AI will happily build the wrong product very quickly. Product thinking is the human superpower."),
    h2("One-Page PRD Template (Student Fills This)"),
    simpleTable(
      ["Section", "Prompt for Student"],
      [
        ["App name", "Working title"],
        ["Problem", "Who struggles with what today?"],
        ["User", "One primary user persona"],
        ["Core job", "In one sentence: the user can …"],
        ["Must-have features (MVP)", "Only 3–5 features for v1"],
        ["Out of scope", "What we will NOT build yet"],
        ["Screens", "List pages/screens"],
        ["Data", "What info do we store? (e.g., tasks: title, done, date)"],
        ["Success demo", "In 60 seconds, what will we click to prove it works?"],
        ["Non-goals", "No payments / no AI chat / no mobile native app (examples)"],
      ],
      [2800, 7424]
    ),
    spacer(100),
    h2("Wireframe Options (Pick One Path)"),
    bullet("Paper boxes + arrows (fastest for non-tech)"),
    bullet("Excalidraw / FigJam free boards"),
    bullet("Ask Claude: “Generate a text wireframe of screens and components”"),
    h2("Scope Control Game"),
    p("Students list 15 dream features → force-rank → keep top 4 for MVP. Teach the phrase: “That’s a v2 idea — capture it in backlog.”"),
    h2("Turn PRD into Claude Plan Prompt"),
    codeBlock([
      "/plan",
      "Read this PRD and propose an architecture for a beginner-friendly",
      "full-stack web app. Prefer simple stack suitable for learning",
      "(explain choices). List milestones M1–M4. Do not write code yet.",
      "Ask me clarifying questions if anything is ambiguous.",
      "",
      "[paste PRD]",
    ]),
    h2("Homework"),
    bullet("Complete PRD for capstone."),
    bullet("Instructor feedback before Module 6 (blocking review)."),
    pageBreak(),
  );

  // ═══════════════ MODULE 6 ═══════════════
  c.push(
    h1("Module 6 — Data & Backend Basics (Still Human Language)"),
    multi([{ text: "Duration: ", bold: true }, { text: "2 sessions" }]),
    h2("Concepts"),
    bulletBold("State: ", "What is true right now in the app (e.g., “3 tasks, 1 completed”)."),
    bulletBold("Persistence: ", "Data still there after refresh/restart."),
    bulletBold("CRUD: ", "Create, Read, Update, Delete — four basic data operations."),
    bulletBold("JSON: ", "A common text format for structured data (show a tiny example)."),
    bulletBold("Environment variables / secrets: ", "Passwords and API keys that must not be published."),
    bulletBold("Server: ", "A computer program that waits for requests and responds."),
    h2("Progressive Path for Beginners"),
    numBold("Level A: ", "LocalStorage in the browser (data stays on that computer/browser). Great first win."),
    numBold("Level B: ", "Simple backend with a file or SQLite database on the machine."),
    numBold("Level C: ", "Hosted database + deployed backend (Module 10)."),
    h2("Lab: Notes App with Persistence"),
    codeBlock([
      "Build a personal notes app:",
      "- List notes, create, edit, delete",
      "- Search by title",
      "- Data must persist after refresh",
      "Start with the simplest approach that works for a beginner.",
      "Explain where data is stored in plain English.",
      "Add a README with how to run the app.",
      "After each feature, pause for my approval.",
    ]),
    h2("Security Seed (Plant Early)"),
    bullet("Never commit passwords or API keys to GitHub."),
    bullet("Use .env files; ensure .gitignore includes .env."),
    bullet("Don’t paste secrets into public Discord/screenshots."),
    pageBreak(),
  );

  // ═══════════════ MODULE 7 ═══════════════
  c.push(
    h1("Module 7 — Full-Stack MVP with Claude Code"),
    multi([{ text: "Duration: ", bold: true }, { text: "2 sessions + lab" }]),
    h2("Goal"),
    p("Assemble frontend + backend + database into one working MVP for the student’s PRD (or a class default app if their PRD is not ready)."),
    h2("Recommended Beginner Stack (Instructor Chooses One Path)"),
    p("Keep ONE stack for the whole cohort to simplify support:"),
    simpleTable(
      ["Path", "Stack Idea", "Why for Non-Tech"],
      [
        ["A (Recommended)", "Next.js or simple Node + HTML, SQLite/Postgres", "One language ecosystem; lots of AI training data"],
        ["B", "Python FastAPI/Flask + simple HTML/JS", "Readable server code for instructors who prefer Python"],
        ["C", "All-in-one starter Claude scaffolds", "Fastest; less portable understanding"],
      ],
      [2000, 4000, 4224]
    ),
    spacer(80),
    callout("Stack Decision Rule", [
      "The best stack is the one the instructor can support live when things break.",
      "Do not let each student pick a different exotic stack in a beginner cohort.",
    ], COLORS.softYellow, COLORS.warn),
    spacer(100),
    h2("Milestone Build Order (Critical)"),
    p("Never build everything at once. Sequence:"),
    num("Empty project runs (“Hello server / Hello page”)."),
    num("One data model (e.g., Task)."),
    num("Create + list only."),
    num("Update + delete."),
    num("Basic styling polish."),
    num("Only then: auth / extras."),
    h2("Claude Session Pattern for MVP"),
    codeBlock([
      "/plan Build milestone 1 only: project scaffold + run instructions",
      "+ empty home page. Stop after plan for my approval.",
    ]),
    p("After approving plan:"),
    codeBlock([
      "Implement the approved plan. After changes, run the app and",
      "show me exact commands to start it. Do not start milestone 2.",
    ]),
    h2("CLAUDE.md Introduction"),
    p("Run /init in the project. Then refine a short CLAUDE.md:"),
    codeBlock([
      "# Project",
      "Beginner full-stack app: [name]. Users: [who].",
      "",
      "# Commands",
      "- Install: ...",
      "- Run dev: ...",
      "- Test: ...",
      "",
      "# Rules",
      "- Prefer simple, readable code",
      "- Explain changes in plain English before big edits",
      "- Never commit .env",
      "- After features: update README",
    ]),
    p("Official guidance: keep CLAUDE.md short — only rules that prevent repeated mistakes. Long files get ignored."),
    h2("Lab Checkpoint Rubric"),
    simpleTable(
      ["Check", "Pass Looks Like"],
      [
        ["Runs locally", "Student starts app without instructor typing for them"],
        ["Core CRUD", "Create/list/update/delete works for main entity"],
        ["Explains parts", "Can point to frontend vs data storage in their words"],
        ["README", "Another student could run it from written steps"],
      ],
      [2800, 7424]
    ),
    pageBreak(),
  );

  // ═══════════════ MODULE 8 ═══════════════
  c.push(
    h1("Module 8 — Users, Auth, and App Polish"),
    multi([{ text: "Duration: ", bold: true }, { text: "1–2 sessions" }]),
    h2("Auth in Plain Language"),
    bulletBold("Sign up: ", "Create account."),
    bulletBold("Login: ", "Prove who you are."),
    bulletBold("Session/token: ", "The app remembers you for a while."),
    bulletBold("Logout: ", "Forget this browser session."),
    bulletBold("Authorization: ", "What you are allowed to see (your tasks, not others’)."),
    h2("Beginner-Safe Auth Approach"),
    p("For learning, use a simple email/password auth pattern that Claude scaffolds, OR a hosted auth provider if instructor prefers less password security burden. Emphasize:"),
    bullet("Never store plain-text passwords (Claude should use proper hashing libraries)."),
    bullet("Users only access their own data (test with two accounts)."),
    h2("Lab: Add Auth to MVP"),
    codeBlock([
      "/plan Add authentication:",
      "- sign up, login, logout",
      "- protect private pages",
      "- each user only sees their own records",
      "List security concerns for beginners.",
      "Propose simplest safe approach for our stack.",
      "No code until I approve.",
    ]),
    h2("Polish Pass"),
    bullet("Loading states (“Saving…”)"),
    bullet("Empty states (“No tasks yet — create one”)"),
    bullet("Form validation messages"),
    bullet("Consistent buttons and spacing"),
    bullet("Favicon + page titles"),
    pageBreak(),
  );

  // ═══════════════ MODULE 9 ═══════════════
  c.push(
    h1("Module 9 — Git & GitHub: Save Points for Your Work"),
    multi([{ text: "Duration: ", bold: true }, { text: "1–2 sessions" }]),
    h2("Metaphors"),
    bulletBold("Git: ", "A time machine / save-point system for your project folder."),
    bulletBold("Commit: ", "A labeled save point with a message."),
    bulletBold("GitHub: ", "A cloud backup + portfolio + collaboration site for Git projects."),
    bulletBold("Repository (repo): ", "The project under version control."),
    bulletBold("Push: ", "Upload commits to GitHub."),
    bulletBold("Pull: ", "Download latest from GitHub."),
    h2("Minimal Git Workflow for This Course"),
    num("git init (or Claude creates repo)"),
    num("Make changes with Claude"),
    num("Review: “what files changed?”"),
    num("commit with message"),
    num("push to GitHub"),
    h2("Do Git Through Claude (English)"),
    codeBlock([
      "Show me what files changed.",
      "Commit these changes with a clear message summarizing the feature.",
      "Create a GitHub repo and push (guide me through auth if needed).",
    ]),
    h2("README as a Professional Habit"),
    bullet("What the app does"),
    bullet("Screenshots"),
    bullet("How to run locally"),
    bullet("Tech stack (even if AI wrote it — student should know names)"),
    bullet("Author name"),
    h2("Safety"),
    bullet("Check `git status` before push."),
    bullet("Ensure .env not tracked."),
    bullet("Public repos = public code — no secrets, no private client data."),
    pageBreak(),
  );

  // ═══════════════ MODULE 10 ═══════════════
  c.push(
    h1("Module 10 — Deploy, Test, and Demo Day"),
    multi([{ text: "Duration: ", bold: true }, { text: "2 sessions" }]),
    h2("What Deployment Means"),
    p("Moving the app from “only on my laptop” to “anyone with the link can use it.”"),
    h2("Beginner Hosting Options (Pick One for Cohort)"),
    bullet("Vercel / Netlify — great for many frontend and some full-stack setups"),
    bullet("Railway / Render / Fly.io — common for apps with servers/databases"),
    bullet("Instructor provides a single recommended path with screenshots"),
    h2("Deploy Checklist"),
    num("App runs locally cleanly."),
    num("Environment variables set on host (not only on laptop)."),
    num("Database provisioned if needed."),
    num("Smoke test on phone + laptop."),
    num("Create two user accounts; verify data isolation."),
    num("Custom error page or friendly failure message if possible."),
    num("README includes live URL."),
    h2("Manual Test Script (Students Write This)"),
    p("Example:"),
    codeBlock([
      "1. Open live URL",
      "2. Sign up with new email",
      "3. Create 2 items",
      "4. Edit 1 item",
      "5. Delete 1 item",
      "6. Logout / login again — data still there",
      "7. Try empty form submit — see validation",
    ]),
    h2("Demo Day Format (5 Minutes per Student)"),
    num("Problem (30s)"),
    num("Live walkthrough (2–3 min)"),
    num("What AI did vs what I decided (1 min)"),
    num("What I’d build next (30s)"),
    h2("Celebration & Portfolio"),
    bullet("LinkedIn post template provided"),
    bullet("GitHub repo pin"),
    bullet("Certificate of completion (optional)"),
    pageBreak(),
  );

  // ═══════════════ MODULE 11-12 advanced optional ═══════════════
  c.push(
    h1("Module 11 (Optional Advanced) — Skills, MCP, and Power Features"),
    p("Only after students ship an MVP. Prevent tool overload early."),
    h2("Topics"),
    bulletBold("Skills: ", "Reusable instruction packs (SKILL.md) Claude can load when relevant."),
    bulletBold("Custom commands: ", "Save repeated procedures."),
    bulletBold("MCP servers: ", "Connect Claude to external tools (databases, browsers, APIs) carefully."),
    bulletBold("Subagents / parallel tasks: ", "Advanced delegation — overview only."),
    bulletBold("Hooks: ", "Automation gates (e.g., run tests before stop) — demo by instructor."),
    h2("Student Exercise"),
    p("Create one skill: `/explain-like-beginner` that forces Claude to summarize any change in non-technical language and list manual test steps."),
    pageBreak(),

    h1("Module 12 (Optional) — From Vibe Coding to Responsible Building"),
    h2("Ethics & Limits"),
    bullet("AI can be wrong, insecure, or outdated — humans own the outcome."),
    bullet("Accessibility: alt text, keyboard basics, contrast."),
    bullet("Privacy: don’t collect data you don’t need."),
    bullet("Licensing: respect package licenses; don’t plagiarize proprietary code intentionally."),
    bullet("Academic honesty: if this is a school course, define what “your work” means when AI writes code (direction, testing, PRD, explanation)."),
    h2("Career Bridge"),
    bullet("Roles enabled: AI-augmented product builder, operations automator, founder MVP builder, marketing engineer."),
    bullet("Next learning: deeper JavaScript or Python fundamentals, databases, product management."),
    bullet("Honest message: vibe coding accelerates building; long-term careers still reward systems thinking and verification skill."),
    pageBreak(),
  );

  // ═══════════════ PART C: LESSON SCRIPTS ═══════════════
  c.push(
    h1("Part C — Instructor Facilitation Playbook"),
    h2("Session Structure Template (90–120 min)"),
    simpleTable(
      ["Minutes", "Block", "Purpose"],
      [
        ["0–10", "Warm-up + error hospital", "Fix broken homework; normalize struggle"],
        ["10–25", "Concept (metaphor + 1 diagram)", "Mental model only"],
        ["25–45", "Live demo (instructor)", "Students watch full loop once"],
        ["45–90", "Guided lab", "Students build; TAs roam"],
        ["90–100", "Share-out (2 students)", "Social proof"],
        ["100–110", "Recap + glossary words", "Spaced repetition"],
        ["110–120", "Homework brief", "Crystal clear next step"],
      ],
      [1500, 3000, 5724]
    ),
    spacer(120),
    h2("The 3-Minute Stuck Protocol"),
    num("Read the error out loud (or describe what doesn’t work)."),
    num("Copy exact error text."),
    num("Paste to Claude: “I did X. I expected Y. I got Z. [error]. Fix root cause and explain simply.”"),
    num("If still stuck → ask TA; do not randomly click for 20 minutes."),
    h2("Differentiation"),
    bulletBold("Struggling students: ", "Use class default project instead of custom capstone until Module 7 works."),
    bulletBold("Fast students: ", "Add stretch goals (export CSV, dark mode, tags) without leaving classmates behind."),
    h2("Assessment Model"),
    simpleTable(
      ["Type", "Weight (suggested)", "What Is Graded"],
      [
        ["Participation & labs", "20%", "Show-up, attempt, peer help"],
        ["Prompt journal", "10%", "Quality of rewrites"],
        ["PRD", "15%", "Clarity, scope control"],
        ["MVP checkpoints", "25%", "Runs + core features"],
        ["GitHub + README", "10%", "Repo hygiene"],
        ["Deployed demo", "20%", "Live URL + presentation"],
      ],
      [2800, 2400, 5024]
    ),
    spacer(100),
    h2("Capstone Rubric (100 points)"),
    simpleTable(
      ["Criterion", "Points", "Notes"],
      [
        ["Problem clarity & PRD", "15", "Who/what/why clear"],
        ["Working core user flow", "25", "Happy path works live"],
        ["Data persistence", "15", "Survives refresh/redeploy"],
        ["Auth or multi-user safety (if claimed)", "10", "Or justified single-user scope"],
        ["UX readability", "10", "Usable by a stranger"],
        ["Git history + README", "10", "Someone else can run"],
        ["Deployment", "10", "Public URL"],
        ["Demo & reflection", "5", "What they directed vs AI"],
      ],
      [4000, 1200, 5024]
    ),
    pageBreak(),
  );

  // ═══════════════ PART D: PROMPT LIBRARY ═══════════════
  c.push(
    h1("Part D — Ready-to-Teach Prompt Library"),
    h2("Setup & Orientation"),
    codeBlock([
      "what can you do in this folder?",
      "explain this project to me like I am completely new to coding",
      "what files matter most and why?",
    ]),
    h2("Planning"),
    codeBlock([
      "/plan [goal]. Ask clarifying questions. List risks.",
      "Do not edit files until I approve the plan.",
    ]),
    h2("Implementation"),
    codeBlock([
      "Implement only step 1 of the approved plan.",
      "Then stop and show me how to verify it.",
    ]),
    h2("Debugging"),
    codeBlock([
      "Bug report:",
      "Steps I took:",
      "Expected:",
      "Actual:",
      "Error text:",
      "Please find root cause, fix it, and give me a short beginner explanation.",
    ]),
    h2("UI"),
    codeBlock([
      "Improve layout for mobile. Increase spacing. Keep our color scheme.",
      "Do not change business logic. Show before/after summary.",
    ]),
    h2("Git"),
    codeBlock([
      "Summarize my uncommitted changes in plain English.",
      "Commit with a descriptive message. Do not push yet.",
    ]),
    h2("Documentation"),
    codeBlock([
      "Write a beginner-friendly README with: overview, features,",
      "setup steps, env vars needed, and how to run tests.",
    ]),
    h2("Verification"),
    codeBlock([
      "Create a manual test checklist for a non-technical user.",
      "Then run any automated tests that exist and report results.",
    ]),
    pageBreak(),
  );

  // ═══════════════ PART E: FULL CAPSTONE BRIEF ═══════════════
  c.push(
    h1("Part E — Capstone Project Brief (Student Handout)"),
    h2("Mission"),
    p("Design and ship a small full-stack web application that solves a real problem for a real user (including yourself). You will direct Claude Code for implementation. You are responsible for clarity, testing, and the final demo."),
    h2("Allowed Capstone Themes (Examples)"),
    bullet("Personal finance mini-tracker"),
    bullet("Habit / study streak tracker"),
    bullet("Local shop inventory for a family business"),
    bullet("Job application tracker"),
    bullet("Classroom resource bookmark manager"),
    bullet("Simple CRM for freelancers (clients + notes)"),
    bullet("Event RSVP list for a community group"),
    h2("Minimum Feature Set"),
    num("Landing or home screen explaining the app"),
    num("Create / read / update / delete for main records"),
    num("Persistent storage"),
    num("Basic validation and error messages"),
    num("README + GitHub repository"),
    num("Deployed live URL"),
    num("Optional but encouraged: user accounts"),
    h2("Timeline"),
    simpleTable(
      ["When", "Deliverable"],
      [
        ["Week 5", "PRD approved"],
        ["Week 7", "MVP core flow working locally"],
        ["Week 8", "Auth/polish checkpoint"],
        ["Week 9", "GitHub complete"],
        ["Week 10", "Deployed + demo"],
      ],
      [2800, 7424]
    ),
    h2("Academic Integrity Statement (Customize)"),
    p("Using Claude Code is required and expected. Submitting a project you cannot explain is not. In demos, you must describe the problem, user flow, and at least three decisions you made (scope cuts, design choices, bug fixes)."),
    pageBreak(),
  );

  // ═══════════════ PART F: GLOSSARY ═══════════════
  c.push(
    h1("Part F — Beginner Glossary (Teach & Revisit)"),
    simpleTable(
      ["Term", "Plain Meaning"],
      [
        ["Terminal / CLI", "Text window to control the computer with commands"],
        ["Claude Code", "Anthropic’s AI coding agent that can edit files and run tools"],
        ["Prompt", "Instructions you type to the AI"],
        ["Hallucination", "AI confidently invents wrong facts or code"],
        ["Bug", "Something that doesn’t work as intended"],
        ["Frontend", "User-facing part of an app"],
        ["Backend", "Server-side logic users don’t see directly"],
        ["Database", "Organized long-term storage for app data"],
        ["API", "Contract/messages between systems"],
        ["Full-stack", "Frontend + backend + data together"],
        ["Deploy", "Publish app so others can access it"],
        ["Repository", "Project tracked by Git"],
        ["Commit", "A saved snapshot in Git"],
        ["Branch", "A parallel line of work in Git"],
        ["PR / Pull Request", "Proposal to merge code changes (collaboration)"],
        ["Environment variable", "Config/secret stored outside code"],
        ["Localhost", "Your own computer acting as the server during development"],
        ["Port", "Numbered door on a computer where a server listens (e.g., 3000)"],
        ["Dependency / package", "Reusable code libraries your project installs"],
        ["Node.js / npm", "Common JavaScript runtime and package tool"],
        ["Framework", "Structured toolkit for building apps (e.g., Next.js)"],
        ["Authentication", "Proving who the user is"],
        ["Authorization", "What that user is allowed to do"],
        ["CRUD", "Create, Read, Update, Delete"],
        ["MVP", "Minimum Viable Product — smallest useful version"],
        ["PRD", "Product Requirements Document"],
        ["Refactor", "Improve code structure without changing behavior"],
        ["Regression", "Something that used to work breaks after a change"],
        ["CLAUDE.md", "Project instruction file Claude Code reads every session"],
        ["Plan mode", "Claude plans without editing until you approve"],
        ["Context window", "How much conversation/file content the AI can hold at once"],
        ["MCP", "Model Context Protocol — way to plug external tools into Claude"],
      ],
      [2800, 7424]
    ),
    pageBreak(),
  );

  // ═══════════════ PART G: TROUBLESHOOTING ═══════════════
  c.push(
    h1("Part G — Troubleshooting Field Guide"),
    h2("Install & Login"),
    simpleTable(
      ["Symptom", "Likely Cause", "Fix"],
      [
        ["command not found: claude", "PATH not set", "Add ~/.local/bin to PATH; reopen terminal; see official terminal guide"],
        ["irm not recognized", "Used CMD not PowerShell", "Open PowerShell or use CMD install command"],
        ["&& not valid separator", "Used PowerShell with CMD syntax", "Use PowerShell install line"],
        ["Login loop / no access", "Free plan only", "Need Pro/Max/Team/Console"],
        ["Region / 403 install", "Availability or network", "Check Anthropic supported countries; retry; alt install"],
      ],
      [2800, 3000, 4424]
    ),
    spacer(120),
    h2("During Building"),
    simpleTable(
      ["Symptom", "Likely Cause", "Fix"],
      [
        ["Claude ignores instructions", "CLAUDE.md too long / chat too long", "Shorten rules; /clear; new session; be more specific"],
        ["Works on laptop, fails online", "Env vars / DB not set on host", "Copy secrets to host dashboard; recheck logs"],
        ["Data mixes between users", "Missing ownership filters", "Test with 2 accounts; fix queries to filter by user id"],
        ["Port already in use", "Old server still running", "Stop old process or change port"],
        ["Endless broken fixes", "No plan; huge scope", "Revert/rewind; /plan smaller milestone"],
        ["Student panic at red text", "Emotional response", "3-minute stuck protocol; pair up"],
      ],
      [2800, 3000, 4424]
    ),
    spacer(120),
    h2("Windows Notes"),
    bullet("Git for Windows recommended so Claude can use Bash tooling."),
    bullet("Know whether you are in PowerShell (PS prompt) vs CMD."),
    bullet("WSL2 is optional but helpful for Linux-like workflows."),
    pageBreak(),
  );

  // ═══════════════ PART H: WEEKLY HOMEWORK SHEETS ═══════════════
  c.push(
    h1("Part H — Weekly Homework Sheets (Copy for LMS)"),
    h2("Week 1"),
    bullet("Customized Hello webpage + reflection (3 sentences)"),
    bullet("Screenshot of terminal showing claude --version"),
    h2("Week 2"),
    bullet("Prompt journal: 5 bad→good prompts"),
    bullet("Study timer or tip calculator built via PROMPT formula"),
    h2("Week 3"),
    bullet("4-page portfolio site"),
    bullet("Glossary quiz (self-check 10 terms)"),
    h2("Week 4"),
    bullet("Interactive quiz app + manual test notes"),
    h2("Week 5"),
    bullet("Capstone PRD v1 submitted for feedback"),
    h2("Week 6"),
    bullet("Notes/todo with persistence + explanation of where data lives"),
    h2("Week 7"),
    bullet("Capstone MVP local demo video (3 min unlisted)"),
    h2("Week 8"),
    bullet("Auth or polish checkpoint checklist signed off"),
    h2("Week 9"),
    bullet("Public GitHub repo + complete README"),
    h2("Week 10"),
    bullet("Live URL + demo day slides (3 slides max)"),
    pageBreak(),
  );

  // ═══════════════ PART I: SAMPLE 2-HOUR LESSON ═══════════════
  c.push(
    h1("Part I — Sample Detailed Lesson Plan (Module 2, Session A)"),
    h2("Title"),
    p("Directing Claude: From Vague Wishes to Verifiable Builds"),
    h2("Objectives"),
    bullet("Students rewrite vague prompts using PROMPT formula."),
    bullet("Students use plan mode once successfully."),
    bullet("Students verify an app with a written checklist."),
    h2("Materials"),
    bullet("Projector, large terminal font"),
    bullet("PROMPT poster"),
    bullet("Starter folder week2-prompts"),
    h2("Minute-by-Minute"),
    simpleTable(
      ["Time", "Activity"],
      [
        ["0:00", "Welcome; 2 students share homework screenshots"],
        ["0:10", "Story: “AI built the wrong app because I was vague”"],
        ["0:20", "PROMPT formula teach-back (students explain each letter)"],
        ["0:35", "Live demo: bad prompt failure → rewrite → success"],
        ["0:50", "Introduce /plan with a small feature"],
        ["1:05", "Lab: build tip calculator with partner tester role"],
        ["1:40", "Gallery walk: 3 apps on screen"],
        ["1:50", "Recap commands: /plan /clear /help"],
        ["2:00", "Homework assigned; exit ticket: one improved prompt"],
      ],
      [1500, 8724]
    ),
    h2("Exit Ticket Question"),
    p("“Write a good prompt for a grocery list webpage with 3 must-have behaviors.”"),
    h2("Common Failure Points"),
    bullet("Students skip verification — require checklist photo."),
    bullet("Students accept huge unplanned rewrites — enforce plan mode."),
    pageBreak(),
  );

  // ═══════════════ PART J: DEFAULT CLASS APP ═══════════════
  c.push(
    h1("Part J — Default Class App Spec (If Capstone Ideas Are Weak)"),
    h2("App Name"),
    p("“FocusBoard” — a personal task board for students."),
    h2("MVP Features"),
    bullet("Sign up / login"),
    bullet("Create tasks with title, notes, due date, status (todo/doing/done)"),
    bullet("Filter by status"),
    bullet("Edit and delete tasks"),
    bullet("Dashboard counts: total, done this week"),
    h2("Stretch"),
    bullet("Tags"),
    bullet("Dark mode"),
    bullet("Export tasks as CSV"),
    h2("Why This Works for Teaching"),
    p("CRUD is obvious, auth matters, UI is simple, demo is clear in under two minutes, and every student understands the domain."),
    pageBreak(),
  );

  // ═══════════════ PART K: PARENT/SPONSOR FAQ ═══════════════
  c.push(
    h1("Part K — FAQ for Students, Parents, and Sponsors"),
    h3("Do students become software engineers in 10 weeks?"),
    p("No. They become capable AI-directed builders of small real apps, with foundations to keep learning. That is already transformative for non-tech careers and founders."),
    h3("Is typing English enough forever?"),
    p("English (or natural language) is the interface, but judgment, product sense, verification, and basic technical vocabulary are the lasting skills."),
    h3("Is Claude Code free?"),
    p("Claude Code requires a paid Claude subscription tier or API access. Budget accordingly."),
    h3("What if AI writes insecure code?"),
    p("Treat security as a continuous review topic. Start simple, avoid handling payments/health data in v1, use checklist reviews, and teach secrets hygiene."),
    h3("Windows vs Mac?"),
    p("Both work. Windows students should know PowerShell vs CMD and benefit from Git for Windows."),
    h3("Can this be taught fully online?"),
    p("Yes, with strong screen-share norms, breakout install help, and recorded setup sessions."),
    pageBreak(),
  );

  // ═══════════════ PART L: RESEARCH NOTES ═══════════════
  c.push(
    h1("Part L — Research Basis & Tool Facts (Instructor Notes)"),
    h2("Claude Code (official facts to teach accurately)"),
    bullet("Surfaces: Terminal CLI, VS Code/Cursor extension, Desktop app, Web (claude.ai/code), JetBrains plugin."),
    bullet("Native install recommended; auto-updates on native installs."),
    bullet("macOS 13+, Windows 10 1809+, common Linux distros; 4GB+ RAM; internet required."),
    bullet("Auth: Pro/Max/Team/Enterprise or Console; free Claude.ai plan does not include Claude Code."),
    bullet("Best practices emphasize context management, verification loops, explore→plan→code, specific prompts, CLAUDE.md, skills for on-demand procedures."),
    bullet("Useful commands for curriculum: /init, /plan, /clear, /compact, /doctor, /permissions, /rewind, /help."),
    bullet("Official docs hubs: code.claude.com/docs and docs.anthropic.com Claude Code pages (overview, setup, quickstart, terminal guide, best practices, memory, skills, commands)."),
    h2("Pedagogy Research Themes (2025–2026 vibe coding education)"),
    bullet("Non-developer courses succeed when they include product framing (PRD/wireframe), not only tools."),
    bullet("Educators warn pure “accept all AI code” produces impressive demos without transferable skill — balance speed with explanation and verification."),
    bullet("Coursera/DeepLearning.AI-style curricula sequence: agent principles → PRD → build → debug → fundamentals overview."),
    bullet("Computer literacy gaps (terminal, files, accounts) are the #1 silent failure mode — hence Module 1 weight."),
    h2("Recommended Instructor Prep Sources"),
    bullet("Claude Code Quickstart + Terminal Guide + Best Practices (official)"),
    bullet("Anthropic supported countries list (for international cohorts)"),
    bullet("GitHub Skills “Introduction to GitHub” (optional pre-reading)"),
    bullet("One short product thinking article on MVP scope"),
    pageBreak(),
  );

  // ═══════════════ PART M: CHECKLISTS ═══════════════
  c.push(
    h1("Part M — Printable Checklists"),
    h2("Student Daily Build Checklist"),
    num("I know which folder I’m in."),
    num("I started Claude from that folder."),
    num("I stated Purpose + Musts + Proof in my prompt."),
    num("For big work, I used plan mode first."),
    num("I approved only changes I roughly understand."),
    num("I ran/opened the app myself."),
    num("I tried at least one edge case."),
    num("I committed a save point if it works."),
    num("I wrote one note: what I learned today."),
    spacer(120),
    h2("Instructor Module 1 Install Checklist"),
    num("OS version verified"),
    num("Install command matched shell (PowerShell vs bash)"),
    num("claude --version works"),
    num("claude doctor clean enough to proceed"),
    num("Login succeeds"),
    num("First HTML file generated and opened in browser"),
    num("Student can repeat start steps without reading chat history"),
    spacer(120),
    h2("Ship Checklist (Capstone)"),
    num("PRD linked in README"),
    num("Live URL works in private/incognito window"),
    num("Core CRUD verified"),
    num("No secrets in repo"),
    num("Two-device quick test (or phone + laptop)"),
    num("Demo script rehearsed once"),
    num("Backup: local recording if live internet fails during demo day"),
    pageBreak(),
  );

  // ═══════════════ PART N: 30-DAY ACCELERATED ═══════════════
  c.push(
    h1("Part N — Alternate Formats"),
    h2("Weekend Workshop (12 hours) — Taste of Vibe Coding"),
    bullet("Sat AM: Terminal + install + Hello page"),
    bullet("Sat PM: Prompt craft + interactive mini app"),
    bullet("Sun AM: PRD lite + data persistence app"),
    bullet("Sun PM: Deploy static or simple app + show-and-tell"),
    p("Outcome: confidence + one deployed simple app — not full-stack mastery."),
    h2("4-Week Intensive (Career Switch Sampler)"),
    simpleTable(
      ["Week", "Focus"],
      [
        ["1", "Literacy + Claude fluency + static sites"],
        ["2", "Interactive UI + PRD + persistence"],
        ["3", "Full-stack MVP (class app)"],
        ["4", "Auth lite + GitHub + deploy + demo"],
      ],
      [2000, 8224]
    ),
    h2("Corporate Non-Tech Cohort (Marketing/Ops)"),
    bullet("Emphasize internal tools, dashboards, form workflows."),
    bullet("Stricter data/privacy module."),
    bullet("Capstone tied to a real team pain point with manager sponsor."),
    pageBreak(),
  );

  // ═══════════════ CLOSING ═══════════════
  c.push(
    h1("Closing Note to the Instructor"),
    p("You are not teaching people to memorize JavaScript. You are teaching them to command a powerful AI engineer, to think in products, to verify reality, and to stay calm when systems break."),
    p("Go slower than feels necessary in Weeks 1–2. Speed comes after confidence. Protect early wins. Force small milestones. Make plan mode a habit. Celebrate deployed links like graduations — because for a non-technical student, a live URL is a identity shift: “I can build software.”"),
    spacer(120),
    callout("Course North Star Metric", [
      "Not “lines of code understood,” but:",
      "“Can a student with no prior tech background independently go from idea → PRD → Claude Code sessions → working full-stack app → public URL → clear demo?”",
      "If yes, the course succeeded.",
    ], COLORS.softGreen, COLORS.success),
    spacer(200),
    p("End of Course Guide — Vibe Coding with Claude Code (Full Curriculum for Non-Technical Students)", { bold: true, align: AlignmentType.CENTER, color: COLORS.primary }),
    p("Teach patiently. Verify always. Ship something real.", { italics: true, align: AlignmentType.CENTER, color: COLORS.muted }),
  );

  return c;
}

async function main() {
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: "Arial", size: 21 },
        },
      },
      paragraphStyles: [
        {
          id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 32, bold: true, font: "Arial", color: COLORS.primary },
          paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 },
        },
        {
          id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 26, bold: true, font: "Arial", color: COLORS.accent },
          paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 },
        },
        {
          id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 22, bold: true, font: "Arial", color: COLORS.dark },
          paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 },
        },
      ],
    },
    numbering: { config: numberingConfig },
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_W, height: PAGE_H },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: COLORS.primary, space: 4 } },
              spacing: { after: 80 },
              children: [
                new TextRun({ text: "Vibe Coding with Claude Code", font: "Arial", size: 16, color: COLORS.primary, bold: true }),
                new TextRun({ text: "  |  Full Course Guide for Non-Technical Students", font: "Arial", size: 16, color: COLORS.muted }),
              ],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              border: { top: { style: BorderStyle.SINGLE, size: 6, color: COLORS.border, space: 6 } },
              spacing: { before: 60 },
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ text: "Confidential teaching material  ·  Page ", font: "Arial", size: 14, color: COLORS.muted }),
                new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 14, color: COLORS.muted }),
                new TextRun({ text: " of ", font: "Arial", size: 14, color: COLORS.muted }),
                new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 14, color: COLORS.muted }),
              ],
            }),
          ],
        }),
      },
      children: buildChildren(),
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  const out = "/home/shailesh/Downloads/Training/vibe-coding/Vibe_Coding_Claude_Code_Full_Course_Guide.docx";
  fs.writeFileSync(out, buffer);
  console.log("Wrote:", out, "bytes:", buffer.length);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
