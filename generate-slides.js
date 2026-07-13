/**
 * Vibe Coding with Claude Code — Edition 2 Slides
 * Deep teaching deck aligned with Master Guide / Study Book / Workshops
 */
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");
const brand = JSON.parse(fs.readFileSync(path.join(__dirname, "brand.json"), "utf8"));

const P = {
  dark: "0B3D2E",
  mid: "146B4F",
  accent: "1FA97A",
  mint: "7DDEA8",
  cream: "F7F3EB",
  white: "FFFFFF",
  ink: "1A2421",
  muted: "5A6B63",
  card: "E8F5EF",
  softGold: "FBF3D8",
  coral: "C44B4B",
  softRed: "FDEDEC",
  codeBg: "1E2A32",
};

const bullets = (items) =>
  items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1 } }));

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = brand.name;
  pres.title = `${brand.course.title} — Module Slides`;

  const footer = (s) => {
    s.addText(`${brand.course.title} · ${brand.name}`, {
      x: 0.45, y: 5.28, w: 9.1, h: 0.22,
      fontSize: 10, fontFace: "Arial", color: P.muted, margin: 0,
    });
  };

  const content = (title) => {
    const s = pres.addSlide();
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: P.cream } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.68, fill: { color: P.dark } });
    s.addText(title, {
      x: 0.45, y: 0.14, w: 9.1, h: 0.42,
      fontSize: 20, fontFace: "Arial", bold: true, color: P.white, margin: 0,
    });
    return s;
  };

  const section = (kicker, title, tag) => {
    const s = pres.addSlide();
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: P.mid } });
    s.addShape(pres.shapes.OVAL, { x: 7.6, y: -1, w: 3.8, h: 3.8, fill: { color: P.dark } });
    s.addText(kicker, {
      x: 0.55, y: 1.55, w: 8.5, h: 0.35,
      fontSize: 13, fontFace: "Arial", bold: true, color: P.mint, margin: 0, charSpacing: 2,
    });
    s.addText(title, {
      x: 0.55, y: 2.05, w: 8.5, h: 1.2,
      fontSize: 30, fontFace: "Arial", bold: true, color: P.white, margin: 0,
    });
    if (tag) {
      s.addText(tag, {
        x: 0.55, y: 3.45, w: 8.5, h: 0.5,
        fontSize: 15, fontFace: "Arial", color: P.cream, margin: 0,
      });
    }
    return s;
  };

  const codeBox = (s, lines, x, y, w, h) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w, h, fill: { color: P.codeBg }, rectRadius: 0.08,
    });
    s.addText(lines.join("\n"), {
      x: x + 0.18, y: y + 0.15, w: w - 0.35, h: h - 0.25,
      fontSize: 12, fontFace: "Consolas", color: P.mint, margin: 0, valign: "top",
    });
  };

  const tableLike = (s, headers, rows) => {
    const colW = 9 / headers.length;
    headers.forEach((h, i) => {
      const x = 0.5 + i * colW;
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 0.95, w: colW - 0.05, h: 0.45, fill: { color: P.dark },
      });
      s.addText(h, {
        x: x + 0.08, y: 1.02, w: colW - 0.2, h: 0.32,
        fontSize: 11, bold: true, color: P.white, fontFace: "Arial", margin: 0,
      });
    });
    rows.forEach((row, ri) => {
      const y = 1.45 + ri * 0.52;
      row.forEach((cell, ci) => {
        const x = 0.5 + ci * colW;
        s.addShape(pres.shapes.RECTANGLE, {
          x, y, w: colW - 0.05, h: 0.48,
          fill: { color: ri % 2 === 0 ? P.white : "EEF2F0" },
        });
        s.addText(String(cell), {
          x: x + 0.08, y: y + 0.08, w: colW - 0.2, h: 0.35,
          fontSize: 11, color: P.ink, fontFace: "Arial", margin: 0,
        });
      });
    });
  };

  const cmdGrid = (s, items) => {
    items.forEach((it, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.4 + col * 4.8, y = 0.95 + row * 1.25;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y, w: 4.55, h: 1.1, fill: { color: P.white }, rectRadius: 0.08,
      });
      s.addText(it[0], {
        x: x + 0.18, y: y + 0.12, w: 4.2, h: 0.3,
        fontSize: 14, bold: true, color: P.mid, fontFace: "Consolas", margin: 0,
      });
      s.addText(it[1] + "  ·  " + it[2], {
        x: x + 0.18, y: y + 0.5, w: 4.2, h: 0.4,
        fontSize: 12, color: P.ink, fontFace: "Arial", margin: 0,
      });
    });
  };

  // ===== TITLE =====
  {
    const s = pres.addSlide();
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: P.dark } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.95, w: 10, h: 0.675, fill: { color: P.mid } });
    s.addText("EDITION 2  ·  EXPLANATION-FIRST", {
      x: 0.55, y: 1.35, w: 9, h: 0.35,
      fontSize: 13, bold: true, color: P.mint, fontFace: "Arial", margin: 0, charSpacing: 2,
    });
    s.addText("Vibe Coding with\nClaude Code", {
      x: 0.55, y: 1.85, w: 9, h: 1.6,
      fontSize: 38, bold: true, color: P.white, fontFace: "Arial", margin: 0,
    });
    s.addText("Features · Slash commands · HTML/CSS/JS/Python · Git · GitHub · Full-stack", {
      x: 0.55, y: 5.12, w: 9, h: 0.35,
      fontSize: 13, color: P.cream, fontFace: "Arial", margin: 0,
    });
    s.addText(`${brand.name}  ·  ${brand.orgs.provider_line}`, {
      x: 0.55, y: 4.75, w: 9, h: 0.3,
      fontSize: 11, color: P.mint, fontFace: "Arial", margin: 0,
    });
  }

  {
    const s = content("What this deck covers");
    const items = [
      "Claude Code features & how the agent works",
      "Slash commands — what / why / when",
      "CLAUDE.md, plan mode, permissions, skills, MCP",
      "HTML · CSS · JavaScript · Python foundations",
      "Git + GitHub complete beginner path",
      "Full-stack build path, deploy, demo",
    ];
    items.forEach((t, i) => {
      const y = 0.95 + i * 0.65;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.5, y, w: 9, h: 0.55, fill: { color: P.white }, rectRadius: 0.06,
      });
      s.addText(String(i + 1).padStart(2, "0"), {
        x: 0.7, y: y + 0.1, w: 0.7, h: 0.35, fontSize: 14, bold: true, color: P.accent, fontFace: "Arial", margin: 0,
      });
      s.addText(t, {
        x: 1.6, y: y + 0.1, w: 7.6, h: 0.35, fontSize: 15, color: P.ink, fontFace: "Arial", margin: 0,
      });
    });
    footer(s);
  }

  // ===== MINDSET =====
  section("PART A", "Mindset with mechanisms", "Direct · build · verify with understanding");
  {
    const s = content("Vibe coding — precise definition");
    s.addText("You describe outcomes in natural language.\nClaude Code edits files, runs tools, and iterates.\nYou decide, test, and steer when it is wrong.", {
      x: 0.5, y: 1.15, w: 9, h: 1.5, fontSize: 20, color: P.ink, fontFace: "Arial", margin: 0,
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.5, y: 3.0, w: 9, h: 1.6, fill: { color: P.softGold }, rectRadius: 0.08,
    });
    s.addText("Not magic: AI can be wrong. Clear instructions win.\nUnderstanding app parts makes you a better director.", {
      x: 0.75, y: 3.35, w: 8.5, h: 1.0, fontSize: 15, color: P.ink, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }
  {
    const s = content("Your role vs Claude Code");
    const cols = [
      { t: "You own", d: ["Problem & user", "What “done” means", "Approvals", "Real testing", "Scope decisions"] },
      { t: "Claude owns", d: ["Writing most code", "Multi-file edits", "Running commands", "Searching project", "Drafting plans"] },
      { t: "Together", d: ["Working software", "Git history", "Public demo", "Real learning", "Shipped URL"] },
    ];
    cols.forEach((c, i) => {
      const x = 0.4 + i * 3.15;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.0, w: 3.0, h: 3.8, fill: { color: i === 1 ? P.dark : P.white }, rectRadius: 0.1,
      });
      s.addText(c.t, {
        x: x + 0.15, y: 1.2, w: 2.7, h: 0.4, fontSize: 16, bold: true,
        color: i === 1 ? P.mint : P.mid, fontFace: "Arial", margin: 0,
      });
      s.addText(bullets(c.d), {
        x: x + 0.15, y: 1.8, w: 2.7, h: 2.7, fontSize: 13,
        color: i === 1 ? P.white : P.ink, fontFace: "Arial",
      });
    });
    footer(s);
  }
  {
    const s = content("Full-stack = all the layers");
    const layers = [
      ["Frontend", "What users see\nHTML · CSS · JS"],
      ["Backend", "Hidden logic\nPython / Node"],
      ["Database", "Lasting memory\nSQLite / Postgres"],
      ["Deploy", "Public access\nHost + URL"],
    ];
    layers.forEach((L, i) => {
      const x = 0.4 + i * 2.4;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.5, w: 2.25, h: 2.8, fill: { color: i % 2 ? P.dark : P.white }, rectRadius: 0.1,
      });
      s.addText(L[0], {
        x, y: 1.9, w: 2.25, h: 0.5, fontSize: 16, bold: true,
        color: i % 2 ? P.mint : P.mid, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(L[1], {
        x: x + 0.15, y: 2.6, w: 1.95, h: 1.3, fontSize: 13,
        color: i % 2 ? P.white : P.ink, fontFace: "Arial", align: "center", margin: 0,
      });
    });
    footer(s);
  }

  // ===== TERMINAL =====
  section("PART B", "Terminal foundations", "Every command means something");
  {
    const s = content("Essential commands");
    const rows = [
      ["pwd", "Print working directory — where am I?"],
      ["ls / dir", "List files in this folder"],
      ["cd name", "Change into a folder"],
      ["cd ..", "Go up one folder"],
      ["mkdir name", "Create a new folder"],
      ["claude", "Start Claude Code in this folder"],
    ];
    rows.forEach((r, i) => {
      const y = 0.9 + i * 0.65;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.5, y, w: 9, h: 0.55, fill: { color: P.white }, rectRadius: 0.05,
      });
      s.addText(r[0], {
        x: 0.7, y: y + 0.1, w: 2.3, h: 0.35, fontSize: 14, bold: true, color: P.mid, fontFace: "Consolas", margin: 0,
      });
      s.addText(r[1], {
        x: 3.2, y: y + 0.1, w: 6, h: 0.35, fontSize: 14, color: P.ink, fontFace: "Arial", margin: 0,
      });
    });
    footer(s);
  }
  {
    const s = content("PATH & “command not found”");
    s.addText("The shell finds programs using PATH folders.\nIf claude is installed but not on PATH → command not found.", {
      x: 0.5, y: 1.15, w: 9, h: 1.1, fontSize: 17, color: P.ink, fontFace: "Arial", margin: 0,
    });
    codeBox(s, [
      "# often after install — then open a NEW terminal",
      'export PATH="$HOME/.local/bin:$PATH"',
      "claude --version",
      "claude doctor",
    ], 0.5, 2.5, 9, 2.1);
    footer(s);
  }

  // ===== CLAUDE CORE =====
  section("PART C", "Claude Code — how it works", "Agent loop · context · install");
  {
    const s = content("Not just chat — an agent in your project");
    s.addText(bullets([
      "Reads and searches many project files",
      "Creates and edits files",
      "Runs shell commands (install, test, server)",
      "Follows CLAUDE.md + skills",
      "Can use MCP tools when configured",
      "Loops: try → observe → fix → continue",
    ]), { x: 0.5, y: 1.0, w: 9, h: 3.8, fontSize: 17, color: P.ink, fontFace: "Arial" });
    footer(s);
  }
  {
    const s = content("The agent loop");
    const steps = ["Read task\n+ context", "Pick tools\nedit / run", "Permission\nif needed", "Observe\nresults", "Continue\nor stop"];
    steps.forEach((t, i) => {
      const x = 0.3 + i * 1.95;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.8, w: 1.8, h: 1.8, fill: { color: i === 2 ? P.dark : P.white }, rectRadius: 0.1,
      });
      s.addText(String(i + 1), {
        x, y: 1.95, w: 1.8, h: 0.4, fontSize: 14, bold: true,
        color: i === 2 ? P.mint : P.accent, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(t, {
        x: x + 0.1, y: 2.5, w: 1.6, h: 0.9, fontSize: 12,
        color: i === 2 ? P.white : P.ink, fontFace: "Arial", align: "center", margin: 0,
      });
    });
    s.addText("It stops when work “looks done” — so YOU define proof (run app / tests).", {
      x: 0.5, y: 4.0, w: 9, h: 0.5, fontSize: 14, color: P.muted, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }
  {
    const s = content("Context window = scarce resource");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.5, y: 1.1, w: 9, h: 3.6, fill: { color: P.white }, rectRadius: 0.1,
    });
    s.addText(bullets([
      "Holds conversation + file reads + command output",
      "When full → quality drops (forgets, more mistakes)",
      "One major task per session when possible",
      "/clear for a new task · /compact to free space on same task",
      "Keep CLAUDE.md short · use @filename for exact files",
    ]), { x: 0.8, y: 1.4, w: 8.4, h: 3.0, fontSize: 16, color: P.ink, fontFace: "Arial" });
    footer(s);
  }
  {
    const s = content("Install (official)");
    s.addText("macOS / Linux / WSL", {
      x: 0.5, y: 0.95, w: 9, h: 0.3, fontSize: 13, bold: true, color: P.mid, fontFace: "Arial", margin: 0,
    });
    codeBox(s, ["curl -fsSL https://claude.ai/install.sh | bash"], 0.5, 1.3, 9, 0.7);
    s.addText("Windows PowerShell", {
      x: 0.5, y: 2.2, w: 9, h: 0.3, fontSize: 13, bold: true, color: P.mid, fontFace: "Arial", margin: 0,
    });
    codeBox(s, ["irm https://claude.ai/install.ps1 | iex"], 0.5, 2.55, 9, 0.7);
    s.addText("Needs Pro/Max/Team (or Console). Free Claude.ai alone is not enough.\nVerify: claude --version  ·  claude doctor  ·  cd project && claude", {
      x: 0.5, y: 3.55, w: 9, h: 1.0, fontSize: 14, color: P.ink, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }

  // ===== FEATURES =====
  section("PART D", "Features encyclopedia", "Plan · permissions · CLAUDE.md · skills · MCP");
  {
    const s = content("Permission modes (Shift+Tab)");
    const modes = [
      ["Manual", "Asks before many actions", "START HERE"],
      ["Plan", "Research only — no source edits", "Before big work"],
      ["acceptEdits", "Auto file edits", "After you trust review"],
      ["bypass", "Skips checks", "NEVER on class laptops"],
    ];
    modes.forEach((m, i) => {
      const y = 0.95 + i * 0.95;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.5, y, w: 9, h: 0.85, fill: { color: i === 3 ? P.softRed : P.white }, rectRadius: 0.08,
      });
      s.addText(m[0], {
        x: 0.75, y: y + 0.22, w: 2.2, h: 0.4, fontSize: 15, bold: true,
        color: i === 3 ? P.coral : P.mid, fontFace: "Arial", margin: 0,
      });
      s.addText(m[1], {
        x: 3.1, y: y + 0.22, w: 4.0, h: 0.4, fontSize: 14, color: P.ink, fontFace: "Arial", margin: 0,
      });
      s.addText(m[2], {
        x: 7.2, y: y + 0.22, w: 2.0, h: 0.4, fontSize: 12, bold: true, color: P.muted, fontFace: "Arial", margin: 0,
      });
    });
    footer(s);
  }
  {
    const s = content("Plan mode — highest leverage habit");
    s.addText("Claude explores and writes a plan without editing source until you approve.\nEnter: /plan   or   Shift+Tab until Plan.", {
      x: 0.5, y: 1.0, w: 9, h: 1.0, fontSize: 16, color: P.ink, fontFace: "Arial", margin: 0,
    });
    codeBox(s, [
      "/plan Add login with email/password.",
      "List security concerns. Ask clarifying questions.",
      "Do not write code until I approve.",
    ], 0.5, 2.2, 9, 1.8);
    s.addText("Wrong assumptions are free in a plan — expensive in messy code.", {
      x: 0.5, y: 4.3, w: 9, h: 0.4, fontSize: 14, italic: true, color: P.muted, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }
  {
    const s = content("CLAUDE.md = project memory");
    s.addText(bullets([
      "Special file Claude reads at session start",
      "Create starter with /init — then shorten it",
      "Put: run commands, rules, gotchas",
      "Don’t put: essays, huge docs Claude can infer from code",
      "If removing a line wouldn’t cause mistakes — delete it",
    ]), { x: 0.5, y: 1.0, w: 9, h: 3.6, fontSize: 16, color: P.ink, fontFace: "Arial" });
    footer(s);
  }
  {
    const s = content("Skills vs CLAUDE.md vs slash commands");
    tableLike(s, ["Tool", "Loaded when", "Best for"], [
      ["CLAUDE.md", "Every session", "Stable project rules"],
      ["Skill (SKILL.md)", "When relevant /name", "Procedures & checklists"],
      ["Slash /command", "When you type it", "Product controls"],
    ]);
    footer(s);
  }
  {
    const s = content("Other features to know by name");
    const feats = [
      ["@file", "Point Claude at exact files"],
      ["Images", "Paste screenshots / mockups"],
      ["Subagents", "Delegate focused side tasks"],
      ["MCP", "Plugins to external tools/data"],
      ["/rewind", "Undo a bad direction"],
      ["--continue", "Resume last session"],
    ];
    feats.forEach((f, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.45 + col * 4.75, y = 1.0 + row * 1.2;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y, w: 4.5, h: 1.05, fill: { color: P.white }, rectRadius: 0.08,
      });
      s.addText(f[0], {
        x: x + 0.2, y: y + 0.2, w: 4.1, h: 0.3, fontSize: 15, bold: true, color: P.mid, fontFace: "Arial", margin: 0,
      });
      s.addText(f[1], {
        x: x + 0.2, y: y + 0.55, w: 4.1, h: 0.3, fontSize: 13, color: P.ink, fontFace: "Arial", margin: 0,
      });
    });
    footer(s);
  }

  // ===== SLASH COMMANDS =====
  section("PART E", "Slash commands encyclopedia", "What · why · when");
  {
    const s = content("Essential — Week 1–2");
    cmdGrid(s, [
      ["/help", "Show help", "You’re lost"],
      ["/login", "Authenticate", "First run / auth errors"],
      ["/doctor", "Setup checkup", "Install feels broken"],
      ["/init", "Starter CLAUDE.md", "New project"],
      ["/plan", "Plan before edit", "Any real feature"],
      ["/clear", "Fresh conversation", "Switching tasks"],
    ]);
    footer(s);
  }
  {
    const s = content("Essential — daily workflow");
    cmdGrid(s, [
      ["/compact", "Summarize context", "Long same task"],
      ["/diff", "See file changes", "Before commit"],
      ["/rewind", "Roll back", "Wrong direction"],
      ["/resume", "Old sessions", "Next day continue"],
      ["/permissions", "Tool rules", "Too many prompts"],
      ["/config", "Settings", "Model / theme"],
    ]);
    footer(s);
  }
  {
    const s = content("Quality & advanced");
    cmdGrid(s, [
      ["/code-review", "Review the diff", "Before demo/merge"],
      ["/security-review", "Security check", "Before deploy"],
      ["/memory", "Edit lasting rules", "Claude forgets rules"],
      ["/context", "Context usage", "Quality dropped"],
      ["/mcp", "External tools", "After MCP setup"],
      ["/btw", "Side question", "Quick fact mid-task"],
    ]);
    footer(s);
  }
  {
    const s = content("Deep dive: /plan");
    s.addText("What: Enter plan mode (optionally with a task).\nWhy: Separates design from coding; wrong ideas are cheap.\nWhen: New features, refactors, multi-file work, unfamiliar code.", {
      x: 0.5, y: 1.0, w: 9, h: 1.4, fontSize: 15, color: P.ink, fontFace: "Arial", margin: 0,
    });
    codeBox(s, [
      "/plan Build a tip calculator as one HTML file.",
      "List UI, edge cases, and how I will test.",
      "Do not write code until I approve.",
    ], 0.5, 2.6, 9, 1.8);
    footer(s);
  }
  {
    const s = content("Deep dive: /clear vs /compact");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.4, y: 1.1, w: 4.4, h: 3.5, fill: { color: P.white }, rectRadius: 0.1,
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.2, y: 1.1, w: 4.4, h: 3.5, fill: { color: P.dark }, rectRadius: 0.1,
    });
    s.addText("/clear", {
      x: 0.65, y: 1.35, w: 4, h: 0.4, fontSize: 20, bold: true, color: P.mid, fontFace: "Arial", margin: 0,
    });
    s.addText("Empty conversation.\nOld chat still in /resume.\nUse when starting a\nDIFFERENT task.", {
      x: 0.65, y: 2.0, w: 3.9, h: 2.2, fontSize: 15, color: P.ink, fontFace: "Arial", margin: 0,
    });
    s.addText("/compact", {
      x: 5.45, y: 1.35, w: 4, h: 0.4, fontSize: 20, bold: true, color: P.mint, fontFace: "Arial", margin: 0,
    });
    s.addText("Summarizes chat to free\nspace but continues work.\nUse when SAME task,\ncontext is full.", {
      x: 5.45, y: 2.0, w: 3.9, h: 2.2, fontSize: 15, color: P.white, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }
  {
    const s = content("Official workflow to chant");
    const steps = ["1. Explore", "2. Plan", "3. Implement", "4. Commit"];
    steps.forEach((t, i) => {
      const x = 0.5 + i * 2.35;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 2.0, w: 2.2, h: 1.5, fill: { color: i % 2 ? P.dark : P.white }, rectRadius: 0.1,
      });
      s.addText(t, {
        x, y: 2.5, w: 2.2, h: 0.5, fontSize: 16, bold: true,
        color: i % 2 ? P.white : P.dark, fontFace: "Arial", align: "center", margin: 0,
      });
    });
    s.addText("Anthropic Claude Code best practices — teach as muscle memory.", {
      x: 0.5, y: 3.9, w: 9, h: 0.5, fontSize: 14, color: P.muted, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }

  // ===== PROMPTS =====
  section("PART F", "Prompt craft", "PROMPT formula · bug reports");
  {
    const s = content("Bad vs good");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.4, y: 1.1, w: 4.4, h: 3.5, fill: { color: P.softRed }, rectRadius: 0.1,
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.2, y: 1.1, w: 4.4, h: 3.5, fill: { color: P.card }, rectRadius: 0.1,
    });
    s.addText("Bad", {
      x: 0.65, y: 1.35, w: 4, h: 0.4, fontSize: 16, bold: true, color: P.coral, fontFace: "Arial", margin: 0,
    });
    s.addText("“make a nice app”\n“fix everything”\n“make UI modern”", {
      x: 0.65, y: 2.0, w: 3.9, h: 2.2, fontSize: 16, color: P.ink, fontFace: "Arial", margin: 0,
    });
    s.addText("Good", {
      x: 5.45, y: 1.35, w: 4, h: 0.4, fontSize: 16, bold: true, color: P.mid, fontFace: "Arial", margin: 0,
    });
    s.addText("Purpose + rules +\nobjects + musts +\nproof + style\n\nSymptom, expected,\nactual, error text", {
      x: 5.45, y: 2.0, w: 3.9, h: 2.2, fontSize: 15, color: P.ink, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }
  {
    const s = content("PROMPT formula");
    const letters = [
      ["P", "Purpose", "Outcome?"],
      ["R", "Rules", "Constraints?"],
      ["O", "Objects", "Screens/data?"],
      ["M", "Musts", "Acceptance?"],
      ["P", "Proof", "How verify?"],
      ["T", "Tone", "Look/feel?"],
    ];
    letters.forEach((L, i) => {
      const x = 0.4 + (i % 3) * 3.15;
      const y = 1.0 + Math.floor(i / 3) * 1.9;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y, w: 3.0, h: 1.7, fill: { color: P.white }, rectRadius: 0.08,
      });
      s.addShape(pres.shapes.OVAL, {
        x: x + 0.2, y: y + 0.45, w: 0.65, h: 0.65, fill: { color: P.dark },
      });
      s.addText(L[0], {
        x: x + 0.2, y: y + 0.55, w: 0.65, h: 0.45, fontSize: 16, bold: true, color: P.white, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(L[1], {
        x: x + 1.05, y: y + 0.4, w: 1.75, h: 0.4, fontSize: 15, bold: true, color: P.ink, fontFace: "Arial", margin: 0,
      });
      s.addText(L[2], {
        x: x + 1.05, y: y + 0.9, w: 1.75, h: 0.4, fontSize: 13, color: P.muted, fontFace: "Arial", margin: 0,
      });
    });
    footer(s);
  }
  {
    const s = content("Bug report template");
    codeBox(s, [
      "Bug report:",
      "Environment: OS, browser, how I start the app",
      "Steps I took:",
      "Expected:",
      "Actual:",
      "Error text:",
      "Please find ROOT CAUSE, fix, re-verify, explain simply.",
    ], 0.5, 1.1, 9, 3.5);
    footer(s);
  }

  // ===== WEB =====
  section("PART G", "How the web works", "Frontend · backend · database · API · deploy");
  {
    const s = content("Restaurant model (precise)");
    const items = [
      ["Dining room", "Frontend", "UI in browser"],
      ["Kitchen", "Backend", "Server logic"],
      ["Pantry", "Database", "Stored data"],
      ["Waiter", "API", "JSON messages"],
      ["Opening day", "Deploy", "Public URL"],
    ];
    items.forEach((it, i) => {
      const x = 0.3 + i * 1.9;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.4, w: 1.8, h: 3.0, fill: { color: i === 2 ? P.dark : P.white }, rectRadius: 0.08,
      });
      s.addText(it[0], {
        x: x + 0.08, y: 1.7, w: 1.64, h: 0.55, fontSize: 11,
        color: i === 2 ? P.mint : P.muted, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(it[1], {
        x: x + 0.08, y: 2.5, w: 1.64, h: 0.6, fontSize: 14, bold: true,
        color: i === 2 ? P.white : P.ink, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(it[2], {
        x: x + 0.08, y: 3.4, w: 1.64, h: 0.6, fontSize: 12,
        color: i === 2 ? P.cream : P.muted, fontFace: "Arial", align: "center", margin: 0,
      });
    });
    footer(s);
  }
  {
    const s = content("Opening a website — request flow");
    s.addText(bullets([
      "You enter a URL",
      "DNS finds the server address",
      "Browser requests the page (HTTPS)",
      "Server returns HTML (+ CSS/JS links)",
      "Browser renders structure, style, behavior",
      "Further clicks may call APIs for data",
    ]), { x: 0.5, y: 1.0, w: 9, h: 3.8, fontSize: 17, color: P.ink, fontFace: "Arial" });
    footer(s);
  }

  // ===== HTML =====
  section("PART H", "HTML — structure", "Tags · attributes · forms · semantics");
  {
    const s = content("HTML is markup (not full programming)");
    s.addText("HTML labels parts of a document so the browser knows\nwhat is a heading, paragraph, link, image, button, form…", {
      x: 0.5, y: 1.15, w: 9, h: 1.1, fontSize: 18, color: P.ink, fontFace: "Arial", margin: 0,
    });
    codeBox(s, [
      '<p class="intro">Hello world</p>',
      "",
      "opening tag + attribute + content + closing tag",
    ], 0.5, 2.5, 9, 1.9);
    footer(s);
  }
  {
    const s = content("Page skeleton you must recognize");
    codeBox(s, [
      "<!DOCTYPE html>",
      '<html lang="en">',
      "  <head> title, meta, CSS links </head>",
      "  <body> visible content + scripts </body>",
      "</html>",
    ], 0.5, 1.15, 9, 3.4);
    footer(s);
  }
  {
    const s = content("Essential tags");
    tableLike(s, ["Tag", "Purpose", "Remember"], [
      ["h1–h6", "Headings", "One main h1"],
      ["a", "Link", "needs href"],
      ["img", "Image", "src + alt"],
      ["form/input/label", "Data entry", "labels matter"],
      ["button", "Action", "not a div"],
      ["header/main/nav", "Semantics", "better than only div"],
    ]);
    footer(s);
  }

  // ===== CSS =====
  section("PART I", "CSS — presentation", "Selectors · box model · layout");
  {
    const s = content("CSS = how it looks");
    codeBox(s, [
      "p { color: #1a2421; font-size: 16px; }",
      ".card { padding: 16px; border-radius: 8px; }",
      "#logo { width: 48px; }",
    ], 0.5, 1.15, 9, 2.2);
    s.addText("element selector · .class selector · #id selector", {
      x: 0.5, y: 3.7, w: 9, h: 0.5, fontSize: 16, color: P.ink, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }
  {
    const s = content("Box model");
    const boxes = [
      ["Margin", "Outside space"],
      ["Border", "Edge line"],
      ["Padding", "Inside space"],
      ["Content", "Text/image"],
    ];
    boxes.forEach((b, i) => {
      const x = 0.5 + i * 2.35;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.8, w: 2.2, h: 2.0, fill: { color: i === 3 ? P.dark : P.white }, rectRadius: 0.1,
      });
      s.addText(b[0], {
        x, y: 2.2, w: 2.2, h: 0.45, fontSize: 16, bold: true,
        color: i === 3 ? P.mint : P.mid, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(b[1], {
        x, y: 2.8, w: 2.2, h: 0.45, fontSize: 13,
        color: i === 3 ? P.white : P.ink, fontFace: "Arial", align: "center", margin: 0,
      });
    });
    footer(s);
  }
  {
    const s = content("Layout tools you’ll hear");
    s.addText(bullets([
      "Flexbox — row/column alignment (nav bars, centering)",
      "Grid — two-dimensional sections / cards",
      "Media queries — different CSS under a max-width (mobile)",
      "Quality bar: 16px+ body, contrast, spacing, obvious primary button",
    ]), { x: 0.5, y: 1.1, w: 9, h: 3.5, fontSize: 17, color: P.ink, fontFace: "Arial" });
    footer(s);
  }

  // ===== JS =====
  section("PART J", "JavaScript — behavior", "State · DOM · events · fetch · npm");
  {
    const s = content("HTML · CSS · JS roles");
    const roles = [
      { t: "HTML", d: "Structure\n(nouns)" },
      { t: "CSS", d: "Look\n(adjectives)" },
      { t: "JavaScript", d: "Behavior\n(verbs)" },
    ];
    roles.forEach((r, i) => {
      const x = 0.7 + i * 3.1;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.5, w: 2.9, h: 2.8, fill: { color: i === 2 ? P.dark : P.white }, rectRadius: 0.1,
      });
      s.addText(r.t, {
        x, y: 2.0, w: 2.9, h: 0.5, fontSize: 20, bold: true,
        color: i === 2 ? P.mint : P.mid, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(r.d, {
        x: x + 0.2, y: 2.8, w: 2.5, h: 1.0, fontSize: 16,
        color: i === 2 ? P.white : P.ink, fontFace: "Arial", align: "center", margin: 0,
      });
    });
    footer(s);
  }
  {
    const s = content("State → render → events");
    s.addText("State = data that describes the app right now\n(score, current question, isRunning…)\nUI should be a projection of state.", {
      x: 0.5, y: 1.0, w: 9, h: 1.2, fontSize: 16, color: P.ink, fontFace: "Arial", margin: 0,
    });
    codeBox(s, [
      "let count = 0;",
      "btn.addEventListener(\"click\", () => { count += 1; render(); });",
      "function render() { out.textContent = String(count); }",
    ], 0.5, 2.4, 9, 2.0);
    footer(s);
  }
  {
    const s = content("Node.js & npm (recognition)");
    s.addText(bullets([
      "Node.js runs JavaScript outside the browser",
      "npm installs packages listed in package.json",
      "npm install — download dependencies",
      "npm run dev — often starts the dev server",
      "Claude may run these; you should know the words",
    ]), { x: 0.5, y: 1.0, w: 9, h: 3.6, fontSize: 17, color: P.ink, fontFace: "Arial" });
    footer(s);
  }

  // ===== PYTHON =====
  section("PART K", "Python — backend foundations", "Routes · venv · status codes");
  {
    const s = content("Why Python appears");
    s.addText("Readable language for backends and automation.\nFlask/FastAPI turn functions into web APIs.", {
      x: 0.5, y: 1.15, w: 9, h: 1.0, fontSize: 18, color: P.ink, fontFace: "Arial", margin: 0,
    });
    codeBox(s, [
      "def greet(name):",
      "    return f\"Hello, {name}!\"",
      "print(greet(\"Asha\"))",
    ], 0.5, 2.4, 9, 2.0);
    footer(s);
  }
  {
    const s = content("Request flow (full-stack)");
    const steps = ["Browser\nrequest", "Route\nmatches", "Python\nfunction", "DB read\n/write", "JSON\nresponse", "JS updates\nUI"];
    steps.forEach((t, i) => {
      const x = 0.25 + i * 1.6;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.9, w: 1.5, h: 1.8, fill: { color: i % 2 ? P.dark : P.white }, rectRadius: 0.08,
      });
      s.addText(t, {
        x: x + 0.05, y: 2.35, w: 1.4, h: 1.0, fontSize: 12, bold: true,
        color: i % 2 ? P.white : P.ink, fontFace: "Arial", align: "center", margin: 0,
      });
    });
    footer(s);
  }
  {
    const s = content("HTTP status codes");
    tableLike(s, ["Code", "Meaning", "Typical cause"], [
      ["200", "OK", "Successful read"],
      ["201", "Created", "Successful create"],
      ["400", "Bad request", "Invalid input"],
      ["401/403", "Auth problems", "Not logged in / not allowed"],
      ["404", "Not found", "Wrong path/id"],
      ["500", "Server error", "Bug on server"],
    ]);
    footer(s);
  }

  // ===== GIT =====
  section("PART L", "Git — version control", "Commits · branches · .gitignore");
  {
    const s = content("Problem Git solves");
    s.addText("Without Git: final2_really_final.zip chaos.\nWith Git: labeled snapshots, history, collaboration, recovery.", {
      x: 0.5, y: 1.5, w: 9, h: 1.3, fontSize: 20, color: P.ink, fontFace: "Arial", margin: 0,
    });
    s.addText("Git ≠ GitHub   ·   Git is the system   ·   GitHub hosts repos online", {
      x: 0.5, y: 3.3, w: 9, h: 0.6, fontSize: 16, bold: true, color: P.mid, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }
  {
    const s = content("Core vocabulary");
    tableLike(s, ["Term", "Plain meaning"], [
      ["Repository", "Project tracked by Git"],
      ["Commit", "Snapshot + message"],
      ["Stage (add)", "Mark files for next commit"],
      ["Branch", "Parallel line of work"],
      ["Remote", "Hosted copy (often GitHub)"],
      ["Push / Pull", "Upload / download commits"],
    ]);
    footer(s);
  }
  {
    const s = content("Daily Git loop");
    codeBox(s, [
      "git status",
      "git add .",
      "git commit -m \"Add login form validation\"",
      "git push",
    ], 0.5, 1.2, 9, 2.4);
    s.addText("Or ask Claude: “Show changes. Commit with a clear message. Do not push yet.”", {
      x: 0.5, y: 3.9, w: 9, h: 0.5, fontSize: 14, color: P.muted, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }
  {
    const s = content(".gitignore — security gate");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.8, y: 1.3, w: 8.4, h: 3.2, fill: { color: P.softRed }, rectRadius: 0.1,
    });
    s.addText(bullets([
      "Never commit .env or API keys",
      "Ignore node_modules/ and Python venv",
      "Check git status before every push",
      "If a secret leaked: revoke key + create new one",
    ]), { x: 1.2, y: 1.7, w: 7.6, h: 2.5, fontSize: 17, color: P.ink, fontFace: "Arial" });
    footer(s);
  }

  // ===== GITHUB =====
  section("PART M", "GitHub complete guide", "Repo · push · README · PR · portfolio");
  {
    const s = content("What GitHub adds");
    s.addText(bullets([
      "Cloud backup of your repository",
      "Portfolio of public work",
      "Issues for task tracking",
      "Pull Requests for review/merge",
      "Actions for automation (later)",
    ]), { x: 0.5, y: 1.1, w: 9, h: 3.5, fontSize: 18, color: P.ink, fontFace: "Arial" });
    footer(s);
  }
  {
    const s = content("First push path");
    const steps = [
      "Create GitHub account + 2FA",
      "Create empty repo",
      "git remote add origin …",
      "Authenticate (gh auth login)",
      "git push -u origin main",
      "Write a real README",
    ];
    steps.forEach((t, i) => {
      const y = 0.95 + i * 0.65;
      s.addShape(pres.shapes.OVAL, {
        x: 0.6, y: y + 0.05, w: 0.45, h: 0.45, fill: { color: P.dark },
      });
      s.addText(String(i + 1), {
        x: 0.6, y: y + 0.1, w: 0.45, h: 0.35, fontSize: 12, bold: true, color: P.white, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(t, {
        x: 1.3, y: y + 0.08, w: 8, h: 0.4, fontSize: 15, color: P.ink, fontFace: "Arial", margin: 0,
      });
    });
    footer(s);
  }
  {
    const s = content("README quality bar");
    s.addText(bullets([
      "What the app does (1–2 sentences)",
      "Screenshot",
      "Features list",
      "Tech stack",
      "Setup + run steps",
      "Env var NAMES (not secret values)",
      "Live demo URL",
      "Author name",
    ]), { x: 0.5, y: 0.95, w: 9, h: 3.9, fontSize: 16, color: P.ink, fontFace: "Arial" });
    footer(s);
  }
  {
    const s = content("Pull Request concept");
    s.addText("A PR proposes merging one branch’s changes into another\n(with discussion and review).", {
      x: 0.5, y: 1.2, w: 9, h: 1.1, fontSize: 18, color: P.ink, fontFace: "Arial", margin: 0,
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.5, y: 2.5, w: 9, h: 2.0, fill: { color: P.white }, rectRadius: 0.1,
    });
    s.addText("Even solo: use branches + PRs to practice professional workflow.\nClaude can draft PR text and run /code-review on the diff.", {
      x: 0.8, y: 2.95, w: 8.4, h: 1.2, fontSize: 16, color: P.ink, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }

  // ===== BUILD =====
  section("PART N", "Build path & full-stack", "Projects · milestones · vertical slices");
  {
    const s = content("Project ladder");
    const projects = [
      ["0", "Hello", "Install"],
      ["1", "Portfolio", "HTML/CSS"],
      ["2", "Quiz", "JS"],
      ["3", "Notes", "CRUD"],
      ["4", "API+UI", "Backend"],
      ["5", "Auth", "Users"],
      ["6", "Capstone", "Ship"],
    ];
    projects.forEach((pItem, i) => {
      const x = 0.25 + i * 1.38;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.5, w: 1.3, h: 2.9, fill: { color: i === 6 ? P.dark : P.white }, rectRadius: 0.08,
      });
      s.addText(pItem[0], {
        x, y: 1.75, w: 1.3, h: 0.4, fontSize: 18, bold: true,
        color: i === 6 ? P.mint : P.accent, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(pItem[1], {
        x: x + 0.05, y: 2.4, w: 1.2, h: 0.8, fontSize: 13, bold: true,
        color: i === 6 ? P.white : P.ink, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(pItem[2], {
        x: x + 0.05, y: 3.5, w: 1.2, h: 0.6, fontSize: 12,
        color: i === 6 ? P.cream : P.muted, fontFace: "Arial", align: "center", margin: 0,
      });
    });
    footer(s);
  }
  {
    const s = content("Milestone rule");
    const m = ["/plan milestone", "Approve plan", "Implement only that", "Verify yourself", "Commit", "Next milestone"];
    m.forEach((t, i) => {
      const y = 1.0 + i * 0.65;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 1.5, y, w: 7, h: 0.55, fill: { color: P.white }, rectRadius: 0.06,
      });
      s.addText(`${i + 1}.  ${t}`, {
        x: 1.8, y: y + 0.1, w: 6.5, h: 0.35, fontSize: 15, color: P.ink, fontFace: "Arial", margin: 0,
      });
    });
    footer(s);
  }

  // ===== DEPLOY =====
  section("PART O", "Deploy · security · demo", "Public URL · env vars · 5-min talk");
  {
    const s = content("Deploy means…");
    s.addText("Moving the app from “only on my laptop”\nto “anyone with the link can use it.”", {
      x: 0.5, y: 1.6, w: 9, h: 1.3, fontSize: 24, color: P.ink, fontFace: "Arial", margin: 0,
    });
    s.addText("Configure: build · start · env vars · database · URL", {
      x: 0.5, y: 3.3, w: 9, h: 0.6, fontSize: 15, color: P.muted, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }
  {
    const s = content("Why local works but production fails");
    s.addText(bullets([
      "Env vars set locally but not on the host",
      "Database URL still points to localhost",
      "Dev-only setup missing in production",
      "Wrong runtime version",
      "CORS / networking misconfig",
    ]), { x: 0.5, y: 1.1, w: 9, h: 3.5, fontSize: 17, color: P.ink, fontFace: "Arial" });
    footer(s);
  }
  {
    const s = content("Demo day — 5 minutes");
    const d = [
      ["0:30", "Problem & user"],
      ["2–3 min", "Live happy path"],
      ["1 min", "3 decisions YOU made"],
      ["0:30", "What’s next"],
    ];
    d.forEach((item, i) => {
      const y = 1.1 + i * 0.9;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 1.2, y, w: 7.6, h: 0.75, fill: { color: P.white }, rectRadius: 0.08,
      });
      s.addText(item[0], {
        x: 1.5, y: y + 0.18, w: 2.0, h: 0.4, fontSize: 15, bold: true, color: P.accent, fontFace: "Arial", margin: 0,
      });
      s.addText(item[1], {
        x: 3.8, y: y + 0.18, w: 4.7, h: 0.4, fontSize: 16, color: P.ink, fontFace: "Arial", margin: 0,
      });
    });
    footer(s);
  }

  // ===== REFERENCE =====
  section("REFERENCE", "Daily tools", "Stuck protocol · command strip · docs");
  {
    const s = content("3-minute stuck protocol");
    const steps = ["Say what you tried", "Copy exact error", "Bug template → Claude", "Still stuck? Ask human"];
    steps.forEach((t, i) => {
      const x = 0.4 + i * 2.4;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.7, w: 2.25, h: 2.3, fill: { color: i === 3 ? P.dark : P.white }, rectRadius: 0.1,
      });
      s.addText(String(i + 1), {
        x, y: 2.0, w: 2.25, h: 0.5, fontSize: 22, bold: true,
        color: i === 3 ? P.mint : P.accent, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(t, {
        x: x + 0.12, y: 2.7, w: 2.0, h: 1.0, fontSize: 13,
        color: i === 3 ? P.white : P.ink, fontFace: "Arial", align: "center", margin: 0,
      });
    });
    footer(s);
  }
  {
    const s = content("Command strip (memorize)");
    codeBox(s, [
      "/help  /login  /doctor  /init  /plan  /clear  /compact",
      "/diff  /rewind  /resume  /code-review  /security-review",
      "",
      "git status → add → commit → push",
      "claude --version   ·   claude doctor   ·   claude",
    ], 0.5, 1.2, 9, 3.3);
    footer(s);
  }
  {
    const s = content("Companion documents");
    const docs = [
      ["Master Teaching Guide", "Full explanations Parts 1–28"],
      ["Student Study Book", "Chapters + practice answers"],
      ["Volume 2 Workshops", "13 deep hands-on labs"],
      ["These slides (Ed. 2)", "Classroom projection deck"],
    ];
    docs.forEach((d, i) => {
      const y = 1.05 + i * 0.9;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.7, y, w: 8.6, h: 0.75, fill: { color: P.white }, rectRadius: 0.08,
      });
      s.addText(d[0], {
        x: 1.0, y: y + 0.18, w: 4.5, h: 0.4, fontSize: 15, bold: true, color: P.mid, fontFace: "Arial", margin: 0,
      });
      s.addText(d[1], {
        x: 5.5, y: y + 0.18, w: 3.5, h: 0.4, fontSize: 14, color: P.ink, fontFace: "Arial", margin: 0,
      });
    });
    footer(s);
  }

  // ===== CLOSE =====
  {
    const s = pres.addSlide();
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: P.dark } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.16, h: 5.625, fill: { color: P.accent } });
    s.addText("Explain. Direct. Verify. Ship.", {
      x: 0.65, y: 1.8, w: 8.7, h: 0.8, fontSize: 32, bold: true, color: P.white, fontFace: "Arial", margin: 0,
    });
    s.addText("Idea → PRD → Claude Code → app → GitHub → public URL → demo.", {
      x: 0.65, y: 2.9, w: 8.7, h: 0.7, fontSize: 16, color: P.mint, fontFace: "Arial", margin: 0,
    });
    s.addText("Edition 2 slides — matched to the deep teaching guides.", {
      x: 0.65, y: 4.1, w: 8.7, h: 0.4, fontSize: 14, color: P.cream, fontFace: "Arial", margin: 0,
    });
  }

  const out = path.join(
    "/home/shailesh/Downloads/Training/vibe-coding",
    "Vibe_Coding_Module_Slides.pptx"
  );
  await pres.writeFile({ fileName: out });
  console.log("Wrote", out);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
