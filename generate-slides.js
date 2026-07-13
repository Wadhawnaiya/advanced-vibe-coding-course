/**
 * Vibe Coding with Claude Code — Full Module Slide Deck
 * Student-facing classroom slides for every module
 */
const pptxgen = require("pptxgenjs");
const path = require("path");

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
  gold: "D4A017",
  softGold: "FBF3D8",
  coral: "C44B4B",
};

function bullets(items) {
  return items.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < items.length - 1 },
  }));
}

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Vibe Coding Course";
  pres.title = "Vibe Coding with Claude Code — Module Slides";
  pres.subject = "Student classroom slides for non-technical learners";

  // helpers
  const darkTitle = (title, subtitle) => {
    const s = pres.addSlide();
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: P.dark } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: P.accent } });
    s.addText(title, {
      x: 0.7, y: 1.8, w: 8.6, h: 1.2,
      fontSize: 36, fontFace: "Arial", bold: true, color: P.white, margin: 0,
    });
    if (subtitle) {
      s.addText(subtitle, {
        x: 0.7, y: 3.1, w: 8.6, h: 0.6,
        fontSize: 18, fontFace: "Arial", color: P.mint, margin: 0,
      });
    }
    return s;
  };

  const content = (title) => {
    const s = pres.addSlide();
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: P.cream } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: P.dark } });
    s.addText(title, {
      x: 0.5, y: 0.15, w: 9, h: 0.45,
      fontSize: 22, fontFace: "Arial", bold: true, color: P.white, margin: 0,
    });
    return s;
  };

  const section = (mod, title, tag) => {
    const s = pres.addSlide();
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: P.mid } });
    s.addShape(pres.shapes.OVAL, { x: 7.8, y: -0.8, w: 3.5, h: 3.5, fill: { color: P.dark } });
    s.addShape(pres.shapes.OVAL, { x: -1, y: 3.8, w: 3, h: 3, fill: { color: P.dark } });
    s.addText(mod, {
      x: 0.6, y: 1.6, w: 8, h: 0.4,
      fontSize: 14, fontFace: "Arial", bold: true, color: P.mint, margin: 0, charSpacing: 3,
    });
    s.addText(title, {
      x: 0.6, y: 2.1, w: 8.2, h: 1.2,
      fontSize: 32, fontFace: "Arial", bold: true, color: P.white, margin: 0,
    });
    if (tag) {
      s.addText(tag, {
        x: 0.6, y: 3.5, w: 8, h: 0.4,
        fontSize: 16, fontFace: "Arial", color: P.cream, margin: 0,
      });
    }
    return s;
  };

  const footer = (s, label) => {
    s.addText(label || "Vibe Coding · Claude Code", {
      x: 0.5, y: 5.25, w: 9, h: 0.25,
      fontSize: 10, fontFace: "Arial", color: P.muted, margin: 0,
    });
  };

  // ========== OPENING ==========
  {
    const s = pres.addSlide();
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: P.dark } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.9, w: 10, h: 0.725, fill: { color: P.mid } });
    s.addText("VIBE CODING", {
      x: 0.6, y: 1.5, w: 8.8, h: 0.4,
      fontSize: 14, fontFace: "Arial", bold: true, color: P.mint, margin: 0, charSpacing: 4,
    });
    s.addText("Build Full-Stack Apps\nwith Claude Code", {
      x: 0.6, y: 2.0, w: 8.8, h: 1.5,
      fontSize: 38, fontFace: "Arial", bold: true, color: P.white, margin: 0,
    });
    s.addText("A course for non-technical students  ·  You direct  ·  AI builds  ·  You verify", {
      x: 0.6, y: 5.1, w: 8.8, h: 0.35,
      fontSize: 14, fontFace: "Arial", color: P.cream, margin: 0,
    });
  }

  {
    const s = content("Today’s promise");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.5, y: 1.1, w: 4.3, h: 3.6, fill: { color: P.white }, rectRadius: 0.1,
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.2, y: 1.1, w: 4.3, h: 3.6, fill: { color: P.dark }, rectRadius: 0.1,
    });
    s.addText("You do NOT need", {
      x: 0.75, y: 1.35, w: 3.8, h: 0.4, fontSize: 16, bold: true, color: P.coral, fontFace: "Arial", margin: 0,
    });
    s.addText(bullets([
      "To memorize code syntax",
      "A computer science degree",
      "To be “technical already”",
      "To get everything right first try",
    ]), { x: 0.75, y: 1.9, w: 3.8, h: 2.4, fontSize: 15, color: P.ink, fontFace: "Arial" });
    s.addText("You WILL learn", {
      x: 5.45, y: 1.35, w: 3.8, h: 0.4, fontSize: 16, bold: true, color: P.mint, fontFace: "Arial", margin: 0,
    });
    s.addText(bullets([
      "To direct Claude Code clearly",
      "How apps are made (in plain words)",
      "To test and fix calmly",
      "To ship a real app online",
    ]), { x: 5.45, y: 1.9, w: 3.8, h: 2.4, fontSize: 15, color: P.white, fontFace: "Arial" });
    footer(s);
  }

  {
    const s = content("What is vibe coding?");
    s.addText("You describe what you want in plain language.\nClaude Code writes, edits, and runs the project.\nYou review, test, and steer.", {
      x: 0.5, y: 1.1, w: 9, h: 1.5, fontSize: 22, color: P.ink, fontFace: "Arial", margin: 0,
    });
    const roles = [
      { t: "You", d: "Director\nProduct thinker\nTester" },
      { t: "Claude Code", d: "Builder\nEditor\nCommand runner" },
      { t: "Together", d: "Working software\nthat real people\ncan use" },
    ];
    roles.forEach((r, i) => {
      const x = 0.5 + i * 3.1;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 2.9, w: 2.9, h: 1.9, fill: { color: i === 1 ? P.dark : P.white }, rectRadius: 0.1,
      });
      s.addText(r.t, {
        x: x + 0.15, y: 3.1, w: 2.6, h: 0.4,
        fontSize: 16, bold: true, color: i === 1 ? P.mint : P.mid, fontFace: "Arial", margin: 0,
      });
      s.addText(r.d, {
        x: x + 0.15, y: 3.55, w: 2.6, h: 1.1,
        fontSize: 14, color: i === 1 ? P.white : P.ink, fontFace: "Arial", margin: 0,
      });
    });
    footer(s);
  }

  {
    const s = content("Course map");
    const weeks = [
      ["0–1", "Setup + terminal"],
      ["2–3", "Prompts + web basics"],
      ["4–5", "UI + PRD"],
      ["6–7", "Data + full-stack"],
      ["8–9", "Auth + GitHub"],
      ["10", "Deploy + demo"],
    ];
    weeks.forEach((w, i) => {
      const x = 0.4 + (i % 3) * 3.15;
      const y = 1.05 + Math.floor(i / 3) * 1.9;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y, w: 3.0, h: 1.65, fill: { color: P.white }, rectRadius: 0.08,
      });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.12, h: 1.65, fill: { color: P.accent } });
      s.addText("Week " + w[0], {
        x: x + 0.3, y: y + 0.3, w: 2.5, h: 0.35, fontSize: 12, bold: true, color: P.accent, fontFace: "Arial", margin: 0,
      });
      s.addText(w[1], {
        x: x + 0.3, y: y + 0.75, w: 2.5, h: 0.6, fontSize: 16, bold: true, color: P.ink, fontFace: "Arial", margin: 0,
      });
    });
    footer(s);
  }

  // ========== WEEK 0 ==========
  section("WEEK 0", "Setup & Ready Check", "Accounts · Laptop · Mindset");
  {
    const s = content("Before class — checklist");
    s.addText(bullets([
      "Claude paid plan ready (Pro / Max / Team) — free alone is not enough for Claude Code",
      "GitHub free account created",
      "Laptop charged + internet works",
      "You know Windows vs Mac vs Linux",
      "Password notes stored safely (not on sticky notes on the webcam)",
    ]), { x: 0.5, y: 1.1, w: 9, h: 3.5, fontSize: 18, color: P.ink, fontFace: "Arial" });
    footer(s);
  }

  // ========== MODULE 1 ==========
  section("MODULE 1", "Digital Ground Zero", "Files · Terminal · First Claude page");
  {
    const s = content("Big idea");
    s.addText("Your computer is a workshop.\nFiles are materials.\nFolders are drawers.\nThe terminal is a text remote control.", {
      x: 0.5, y: 1.3, w: 9, h: 2.5, fontSize: 26, color: P.ink, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }
  {
    const s = content("Folder path = address");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.8, y: 1.5, w: 8.4, h: 2.2, fill: { color: P.dark }, rectRadius: 0.1,
    });
    s.addText("Documents  →  vibe-coding-course  →  week1-hello", {
      x: 1.1, y: 2.2, w: 7.8, h: 0.7, fontSize: 20, color: P.white, fontFace: "Consolas", align: "center", margin: 0,
    });
    s.addText("Always know: Where am I? What folder is open?", {
      x: 0.8, y: 4.0, w: 8.4, h: 0.5, fontSize: 16, color: P.muted, fontFace: "Arial", align: "center", margin: 0,
    });
    footer(s);
  }
  {
    const s = content("Safe first terminal commands");
    const cmds = [
      ["pwd / cd", "Where am I? / Go to folder"],
      ["ls  or  dir", "List files"],
      ["mkdir name", "Make a new folder"],
      ["claude", "Start Claude Code"],
    ];
    cmds.forEach((c, i) => {
      const y = 1.05 + i * 0.9;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.5, y, w: 9, h: 0.75, fill: { color: P.white }, rectRadius: 0.08,
      });
      s.addText(c[0], { x: 0.75, y: y + 0.18, w: 3.2, h: 0.4, fontSize: 16, bold: true, color: P.mid, fontFace: "Consolas", margin: 0 });
      s.addText(c[1], { x: 4.2, y: y + 0.18, w: 5, h: 0.4, fontSize: 16, color: P.ink, fontFace: "Arial", margin: 0 });
    });
    footer(s);
  }
  {
    const s = content("Install Claude Code");
    s.addText("macOS / Linux", { x: 0.5, y: 1.0, w: 9, h: 0.35, fontSize: 14, bold: true, color: P.mid, fontFace: "Arial", margin: 0 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.4, w: 9, h: 0.7, fill: { color: P.dark }, rectRadius: 0.06 });
    s.addText("curl -fsSL https://claude.ai/install.sh | bash", {
      x: 0.7, y: 1.55, w: 8.6, h: 0.4, fontSize: 16, color: P.mint, fontFace: "Consolas", margin: 0,
    });
    s.addText("Windows PowerShell", { x: 0.5, y: 2.4, w: 9, h: 0.35, fontSize: 14, bold: true, color: P.mid, fontFace: "Arial", margin: 0 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.8, w: 9, h: 0.7, fill: { color: P.dark }, rectRadius: 0.06 });
    s.addText("irm https://claude.ai/install.ps1 | iex", {
      x: 0.7, y: 2.95, w: 8.6, h: 0.4, fontSize: 16, color: P.mint, fontFace: "Consolas", margin: 0,
    });
    s.addText("Then:  claude --version   →   claude   →   log in when browser opens", {
      x: 0.5, y: 3.9, w: 9, h: 0.5, fontSize: 16, color: P.ink, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }
  {
    const s = content("Lab — first magic moment");
    s.addText(bullets([
      "Open folder week1-hello in the terminal",
      "Type: claude",
      "Say: make me a simple webpage that says Hello, my name is …",
      "Approve file creation",
      "Open the HTML file in your browser",
      "Smile. You just directed software.",
    ]), { x: 0.5, y: 1.1, w: 9, h: 3.6, fontSize: 18, color: P.ink, fontFace: "Arial" });
    footer(s);
  }
  {
    const s = content("Homework — Module 1");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 1, y: 1.3, w: 8, h: 3.2, fill: { color: P.white }, rectRadius: 0.1,
    });
    s.addText(bullets([
      "Customize your page (color + short bio)",
      "Screenshot before & after",
      "Write 3 sentences: surprised / confused / want next",
      "Screenshot: claude --version",
    ]), { x: 1.4, y: 1.7, w: 7.2, h: 2.5, fontSize: 18, color: P.ink, fontFace: "Arial" });
    footer(s);
  }

  // ========== MODULE 2 ==========
  section("MODULE 2", "Directing Claude", "Prompt craft · Plan mode · Verification");
  {
    const s = content("Vague wishes create random apps");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.5, y: 1.2, w: 4.3, h: 3.4, fill: { color: "FDEDEC" }, rectRadius: 0.1,
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.2, y: 1.2, w: 4.3, h: 3.4, fill: { color: P.card }, rectRadius: 0.1,
    });
    s.addText("Bad", { x: 0.8, y: 1.45, w: 3.7, h: 0.4, fontSize: 16, bold: true, color: P.coral, fontFace: "Arial", margin: 0 });
    s.addText("“make a nice app”", {
      x: 0.8, y: 2.2, w: 3.7, h: 1.5, fontSize: 22, color: P.ink, fontFace: "Arial", margin: 0,
    });
    s.addText("Good", { x: 5.5, y: 1.45, w: 3.7, h: 0.4, fontSize: 16, bold: true, color: P.mid, fontFace: "Arial", margin: 0 });
    s.addText("Purpose + rules + objects + musts + how to test + style", {
      x: 5.5, y: 2.2, w: 3.7, h: 1.8, fontSize: 18, color: P.ink, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }
  {
    const s = content("PROMPT formula");
    const letters = [
      ["P", "Purpose", "What outcome?"],
      ["R", "Rules", "Constraints"],
      ["O", "Objects", "Screens & data"],
      ["M", "Musts", "Acceptance checks"],
      ["P", "Proof", "How we verify"],
      ["T", "Tone", "Look & feel"],
    ];
    letters.forEach((L, i) => {
      const x = 0.4 + (i % 3) * 3.15;
      const y = 1.0 + Math.floor(i / 3) * 1.95;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y, w: 3.0, h: 1.75, fill: { color: P.white }, rectRadius: 0.08,
      });
      s.addShape(pres.shapes.OVAL, {
        x: x + 0.2, y: y + 0.35, w: 0.7, h: 0.7, fill: { color: P.dark },
      });
      s.addText(L[0], {
        x: x + 0.2, y: y + 0.48, w: 0.7, h: 0.45, fontSize: 18, bold: true, color: P.white, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(L[1], {
        x: x + 1.05, y: y + 0.35, w: 1.75, h: 0.4, fontSize: 16, bold: true, color: P.ink, fontFace: "Arial", margin: 0,
      });
      s.addText(L[2], {
        x: x + 1.05, y: y + 0.85, w: 1.75, h: 0.5, fontSize: 13, color: P.muted, fontFace: "Arial", margin: 0,
      });
    });
    footer(s);
  }
  {
    const s = content("Official loop");
    const steps = ["1. Explore", "2. Plan", "3. Implement", "4. Commit"];
    steps.forEach((t, i) => {
      const x = 0.45 + i * 2.4;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 2.0, w: 2.2, h: 1.5, fill: { color: i % 2 ? P.dark : P.white }, rectRadius: 0.1,
      });
      s.addText(t, {
        x, y: 2.5, w: 2.2, h: 0.5, fontSize: 16, bold: true,
        color: i % 2 ? P.white : P.dark, fontFace: "Arial", align: "center", margin: 0,
      });
      if (i < 3) {
        s.addText("→", {
          x: x + 2.05, y: 2.5, w: 0.4, h: 0.5, fontSize: 20, color: P.accent, fontFace: "Arial", margin: 0,
        });
      }
    });
    s.addText("Use /plan before big changes. Claude designs first — you approve — then code.", {
      x: 0.5, y: 4.0, w: 9, h: 0.5, fontSize: 15, color: P.muted, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }
  {
    const s = content("Commands to remember");
    const rows = [
      ["/help", "I’m lost"],
      ["/plan", "Think before edit"],
      ["/clear", "New task, fresh chat"],
      ["/compact", "Chat got too long"],
      ["/doctor", "Setup feels broken"],
    ];
    rows.forEach((r, i) => {
      const y = 1.0 + i * 0.7;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.2, y, w: 7.6, h: 0.6, fill: { color: P.white }, rectRadius: 0.06 });
      s.addText(r[0], { x: 1.5, y: y + 0.12, w: 2.5, h: 0.35, fontSize: 16, bold: true, color: P.mid, fontFace: "Consolas", margin: 0 });
      s.addText(r[1], { x: 4.2, y: y + 0.12, w: 4.3, h: 0.35, fontSize: 16, color: P.ink, fontFace: "Arial", margin: 0 });
    });
    footer(s);
  }
  {
    const s = content("Homework — Module 2");
    s.addText(bullets([
      "Prompt journal: 5 bad → good rewrites",
      "Build a tip calculator or study timer using PROMPT",
      "Partner test: try to break each other’s app",
    ]), { x: 0.5, y: 1.3, w: 9, h: 3, fontSize: 20, color: P.ink, fontFace: "Arial" });
    footer(s);
  }

  // ========== MODULE 3 ==========
  section("MODULE 3", "How Web Apps Work", "Frontend · Backend · Data · Deploy");
  {
    const s = content("Restaurant metaphor");
    const items = [
      ["Dining room", "Frontend", "What users see"],
      ["Kitchen", "Backend", "Hidden logic"],
      ["Pantry", "Database", "Stored info"],
      ["Waiter", "API", "Messages"],
      ["Opening day", "Deploy", "Public access"],
    ];
    items.forEach((it, i) => {
      const x = 0.3 + i * 1.9;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.4, w: 1.8, h: 3.0, fill: { color: i === 2 ? P.dark : P.white }, rectRadius: 0.08,
      });
      s.addText(it[0], {
        x: x + 0.1, y: 1.7, w: 1.6, h: 0.6, fontSize: 12, color: i === 2 ? P.mint : P.muted, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(it[1], {
        x: x + 0.1, y: 2.5, w: 1.6, h: 0.7, fontSize: 15, bold: true, color: i === 2 ? P.white : P.ink, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(it[2], {
        x: x + 0.1, y: 3.4, w: 1.6, h: 0.6, fontSize: 12, color: i === 2 ? P.cream : P.muted, fontFace: "Arial", align: "center", margin: 0,
      });
    });
    footer(s);
  }
  {
    const s = content("Full-stack = all the parts together");
    s.addText(bullets([
      "Frontend — screens and buttons",
      "Backend — rules and processing",
      "Database — memory that lasts",
      "Deploy — online so others can use it",
    ]), { x: 0.5, y: 1.2, w: 9, h: 3.2, fontSize: 22, color: P.ink, fontFace: "Arial" });
    footer(s);
  }
  {
    const s = content("Lab — portfolio site");
    s.addText(bullets([
      "Home · About · Projects · Contact pages",
      "Shared style (CSS)",
      "Mobile-friendly",
      "Claude explains the folder structure in plain English",
    ]), { x: 0.5, y: 1.2, w: 9, h: 3.2, fontSize: 20, color: P.ink, fontFace: "Arial" });
    footer(s);
  }

  // ========== MODULE 4 ==========
  section("MODULE 4", "Interactive Frontend", "HTML · CSS · JS (concepts) · Verification");
  {
    const s = content("Three layers of a page");
    const layers = [
      { t: "HTML", d: "Structure\n(skeleton)" },
      { t: "CSS", d: "Look & layout\n(style)" },
      { t: "JavaScript", d: "Behavior\n(clicks & logic)" },
    ];
    layers.forEach((L, i) => {
      const x = 0.7 + i * 3.1;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.5, w: 2.8, h: 2.8, fill: { color: i === 1 ? P.dark : P.white }, rectRadius: 0.1,
      });
      s.addText(L.t, {
        x, y: 2.0, w: 2.8, h: 0.5, fontSize: 22, bold: true,
        color: i === 1 ? P.mint : P.mid, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(L.d, {
        x: x + 0.2, y: 2.7, w: 2.4, h: 1.1, fontSize: 16,
        color: i === 1 ? P.white : P.ink, fontFace: "Arial", align: "center", margin: 0,
      });
    });
    footer(s);
  }
  {
    const s = content("Verification ritual");
    const steps = [
      "Open in browser",
      "Click every button",
      "Try wrong inputs",
      "Write: When I… Expected… Actual…",
      "Paste bug report to Claude",
    ];
    steps.forEach((t, i) => {
      s.addShape(pres.shapes.OVAL, {
        x: 0.6, y: 1.05 + i * 0.75, w: 0.5, h: 0.5, fill: { color: P.dark },
      });
      s.addText(String(i + 1), {
        x: 0.6, y: 1.12 + i * 0.75, w: 0.5, h: 0.4, fontSize: 14, bold: true, color: P.white, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(t, {
        x: 1.4, y: 1.1 + i * 0.75, w: 7.5, h: 0.45, fontSize: 18, color: P.ink, fontFace: "Arial", margin: 0,
      });
    });
    footer(s);
  }
  {
    const s = content("Lab — quiz app");
    s.addText(bullets([
      "5 questions on a topic you choose",
      "One question at a time",
      "Score + restart",
      "Works on phone width",
      "List manual tests and actually run them",
    ]), { x: 0.5, y: 1.2, w: 9, h: 3.5, fontSize: 20, color: P.ink, fontFace: "Arial" });
    footer(s);
  }

  // ========== MODULE 5 ==========
  section("MODULE 5", "Think Before You Build", "PRD · Scope · Wireframes");
  {
    const s = content("AI builds the wrong product very fast");
    s.addText("Your superpower is product thinking —\nnot typing syntax.", {
      x: 0.5, y: 1.6, w: 9, h: 1.5, fontSize: 28, color: P.ink, fontFace: "Arial", margin: 0,
    });
    s.addText("Write the plan first. Build second.", {
      x: 0.5, y: 3.4, w: 9, h: 0.6, fontSize: 20, color: P.mid, fontFace: "Arial", bold: true, margin: 0,
    });
    footer(s);
  }
  {
    const s = content("PRD one-pager sections");
    s.addText(bullets([
      "Problem & user",
      "One-sentence job-to-be-done",
      "3–5 must-have features only",
      "Out of scope (v2 parking lot)",
      "Screens list",
      "Data you store",
      "60-second success demo",
    ]), { x: 0.5, y: 1.1, w: 9, h: 3.8, fontSize: 18, color: P.ink, fontFace: "Arial" });
    footer(s);
  }
  {
    const s = content("Scope game");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.8, y: 1.4, w: 8.4, h: 2.8, fill: { color: P.white }, rectRadius: 0.1,
    });
    s.addText("List 15 dream features\n→ force-rank\n→ keep top 4 for MVP\n\nEverything else: “That’s a v2 idea.”", {
      x: 1.2, y: 1.7, w: 7.6, h: 2.3, fontSize: 20, color: P.ink, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }

  // ========== MODULE 6 ==========
  section("MODULE 6", "Data & Backend Basics", "Persistence · CRUD · Secrets");
  {
    const s = content("CRUD — four data superpowers");
    const c = [
      ["C", "Create", "Add new"],
      ["R", "Read", "View list"],
      ["U", "Update", "Edit"],
      ["D", "Delete", "Remove"],
    ];
    c.forEach((item, i) => {
      const x = 0.5 + i * 2.35;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.5, w: 2.2, h: 2.8, fill: { color: i % 2 ? P.dark : P.white }, rectRadius: 0.1,
      });
      s.addText(item[0], {
        x, y: 1.9, w: 2.2, h: 0.7, fontSize: 36, bold: true,
        color: i % 2 ? P.mint : P.mid, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(item[1], {
        x, y: 2.7, w: 2.2, h: 0.45, fontSize: 18, bold: true,
        color: i % 2 ? P.white : P.ink, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(item[2], {
        x, y: 3.3, w: 2.2, h: 0.4, fontSize: 14,
        color: i % 2 ? P.cream : P.muted, fontFace: "Arial", align: "center", margin: 0,
      });
    });
    footer(s);
  }
  {
    const s = content("Persistence");
    s.addText("If data disappears after refresh,\nsomething is not saving yet.", {
      x: 0.5, y: 1.5, w: 9, h: 1.3, fontSize: 26, color: P.ink, fontFace: "Arial", margin: 0,
    });
    s.addText("Your job: ask Claude where data lives — and test reload.", {
      x: 0.5, y: 3.2, w: 9, h: 0.6, fontSize: 18, color: P.mid, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }
  {
    const s = content("Secrets rule");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 1, y: 1.5, w: 8, h: 2.6, fill: { color: "FDEDEC" }, rectRadius: 0.1,
    });
    s.addText("Never put passwords or API keys\nin public GitHub or class chat screenshots.\n\nUse .env files. Keep them private.", {
      x: 1.4, y: 1.9, w: 7.2, h: 2.0, fontSize: 20, color: P.ink, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }

  // ========== MODULE 7 ==========
  section("MODULE 7", "Full-Stack MVP", "Milestones · CLAUDE.md · Working core flow");
  {
    const s = content("Build in milestones — never everything at once");
    const m = ["Scaffold runs", "One data model", "Create + list", "Update + delete", "Polish UI", "Auth later"];
    m.forEach((t, i) => {
      const y = 1.0 + i * 0.65;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y, w: 7, h: 0.55, fill: { color: P.white }, rectRadius: 0.06 });
      s.addShape(pres.shapes.OVAL, { x: 1.7, y: y + 0.08, w: 0.4, h: 0.4, fill: { color: P.dark } });
      s.addText(String(i + 1), { x: 1.7, y: y + 0.12, w: 0.4, h: 0.35, fontSize: 12, bold: true, color: P.white, fontFace: "Arial", align: "center", margin: 0 });
      s.addText(t, { x: 2.4, y: y + 0.1, w: 5.8, h: 0.35, fontSize: 16, color: P.ink, fontFace: "Arial", margin: 0 });
    });
    footer(s);
  }
  {
    const s = content("Session pattern");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 1.4, fill: { color: P.dark }, rectRadius: 0.08 });
    s.addText("/plan Build milestone 1 only… Stop for my approval.", {
      x: 0.8, y: 1.6, w: 8.4, h: 0.6, fontSize: 18, color: P.mint, fontFace: "Consolas", margin: 0,
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.9, w: 9, h: 1.4, fill: { color: P.white }, rectRadius: 0.08 });
    s.addText("Implement approved plan only. Show exact run commands. Do not start milestone 2.", {
      x: 0.8, y: 3.3, w: 8.4, h: 0.7, fontSize: 18, color: P.ink, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }
  {
    const s = content("CLAUDE.md = project memory");
    s.addText(bullets([
      "Short rules Claude reads every session",
      "How to install & run",
      "Never commit .env",
      "Keep it short — long files get ignored",
    ]), { x: 0.5, y: 1.2, w: 9, h: 3.2, fontSize: 20, color: P.ink, fontFace: "Arial" });
    footer(s);
  }

  // ========== MODULE 8 ==========
  section("MODULE 8", "Users, Login & Polish", "Auth basics · Empty states · Mobile");
  {
    const s = content("Auth in human words");
    const a = [
      ["Sign up", "Create account"],
      ["Login", "Prove identity"],
      ["Session", "Stay signed in"],
      ["Logout", "End session"],
      ["Authorization", "Your data only"],
    ];
    a.forEach((item, i) => {
      const y = 1.0 + i * 0.7;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1, y, w: 8, h: 0.6, fill: { color: P.white }, rectRadius: 0.06 });
      s.addText(item[0], { x: 1.3, y: y + 0.12, w: 3, h: 0.35, fontSize: 16, bold: true, color: P.mid, fontFace: "Arial", margin: 0 });
      s.addText(item[1], { x: 4.5, y: y + 0.12, w: 4.2, h: 0.35, fontSize: 16, color: P.ink, fontFace: "Arial", margin: 0 });
    });
    footer(s);
  }
  {
    const s = content("Test with two accounts");
    s.addText("If User A can see User B’s private data,\nthe app is not done — even if the UI looks pretty.", {
      x: 0.5, y: 1.8, w: 9, h: 1.8, fontSize: 24, color: P.ink, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }
  {
    const s = content("Polish checklist");
    s.addText(bullets([
      "Loading text while saving",
      "Empty state: “No items yet”",
      "Clear form errors",
      "Readable on phone",
      "Sensible page titles",
    ]), { x: 0.5, y: 1.2, w: 9, h: 3.5, fontSize: 20, color: P.ink, fontFace: "Arial" });
    footer(s);
  }

  // ========== MODULE 9 ==========
  section("MODULE 9", "Git & GitHub", "Save points · Backup · Portfolio");
  {
    const s = content("Git is a time machine for your project");
    const g = [
      ["Git", "Save points on your computer"],
      ["Commit", "A labeled snapshot"],
      ["GitHub", "Cloud backup + portfolio"],
      ["Push", "Upload your saves"],
    ];
    g.forEach((item, i) => {
      const x = 0.5 + (i % 2) * 4.7;
      const y = 1.1 + Math.floor(i / 2) * 1.8;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y, w: 4.4, h: 1.55, fill: { color: i === 0 || i === 3 ? P.dark : P.white }, rectRadius: 0.1,
      });
      s.addText(item[0], {
        x: x + 0.3, y: y + 0.3, w: 3.8, h: 0.4, fontSize: 18, bold: true,
        color: i === 0 || i === 3 ? P.mint : P.mid, fontFace: "Arial", margin: 0,
      });
      s.addText(item[1], {
        x: x + 0.3, y: y + 0.85, w: 3.8, h: 0.4, fontSize: 15,
        color: i === 0 || i === 3 ? P.white : P.ink, fontFace: "Arial", margin: 0,
      });
    });
    footer(s);
  }
  {
    const s = content("Ask Claude in English");
    s.addText(bullets([
      "Show me what files changed",
      "Commit with a clear message",
      "Help me push to GitHub",
      "Confirm .env is not tracked",
    ]), { x: 0.5, y: 1.2, w: 9, h: 3.5, fontSize: 22, color: P.ink, fontFace: "Arial" });
    footer(s);
  }
  {
    const s = content("README essentials");
    s.addText(bullets([
      "What the app does",
      "How to run it",
      "Tech stack names",
      "Your name",
      "Live URL (later)",
    ]), { x: 0.5, y: 1.2, w: 9, h: 3.5, fontSize: 22, color: P.ink, fontFace: "Arial" });
    footer(s);
  }

  // ========== MODULE 10 ==========
  section("MODULE 10", "Deploy, Test & Demo Day", "Public URL · Smoke tests · Presentation");
  {
    const s = content("Deploy = open the restaurant to the public");
    s.addText("Your app leaves your laptop\nand gets a link anyone can open.", {
      x: 0.5, y: 1.7, w: 9, h: 1.5, fontSize: 28, color: P.ink, fontFace: "Arial", margin: 0,
    });
    footer(s);
  }
  {
    const s = content("Ship checklist");
    s.addText(bullets([
      "Runs cleanly locally",
      "Env vars set on the host",
      "Works on phone + laptop",
      "Core CRUD verified",
      "No secrets in the repo",
      "README has live URL",
    ]), { x: 0.5, y: 1.1, w: 9, h: 3.8, fontSize: 20, color: P.ink, fontFace: "Arial" });
    footer(s);
  }
  {
    const s = content("Demo day — 5 minutes");
    const d = [
      ["0:30", "Problem"],
      ["2–3 min", "Live walkthrough"],
      ["1 min", "AI vs your decisions"],
      ["0:30", "What’s next"],
    ];
    d.forEach((item, i) => {
      const y = 1.1 + i * 0.9;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.2, y, w: 7.6, h: 0.75, fill: { color: P.white }, rectRadius: 0.08 });
      s.addText(item[0], { x: 1.5, y: y + 0.18, w: 2.2, h: 0.4, fontSize: 16, bold: true, color: P.accent, fontFace: "Arial", margin: 0 });
      s.addText(item[1], { x: 4.0, y: y + 0.18, w: 4.5, h: 0.4, fontSize: 18, color: P.ink, fontFace: "Arial", margin: 0 });
    });
    footer(s);
  }
  {
    const s = content("What we grade (simple view)");
    const items = [
      ["Problem clarity", "15"],
      ["Working core flow", "25"],
      ["Data persists", "15"],
      ["UX readable", "10"],
      ["GitHub + README", "10"],
      ["Deployed live", "10"],
      ["Auth / safety (if claimed)", "10"],
      ["Demo & reflection", "5"],
    ];
    items.forEach((it, i) => {
      const col = i < 4 ? 0 : 1;
      const row = i % 4;
      const x = 0.5 + col * 4.7;
      const y = 1.05 + row * 0.9;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 4.4, h: 0.75, fill: { color: P.white }, rectRadius: 0.06 });
      s.addText(it[0], { x: x + 0.25, y: y + 0.18, w: 3.2, h: 0.4, fontSize: 15, color: P.ink, fontFace: "Arial", margin: 0 });
      s.addText(it[1], { x: x + 3.4, y: y + 0.18, w: 0.7, h: 0.4, fontSize: 16, bold: true, color: P.mid, fontFace: "Arial", margin: 0 });
    });
    footer(s);
  }

  // ========== OPTIONAL ==========
  section("OPTIONAL", "Power Features & Responsible Building", "Skills · Ethics · Next steps");
  {
    const s = content("Later power tools (after you ship MVP)");
    s.addText(bullets([
      "Skills — reusable instruction packs",
      "MCP — connect external tools carefully",
      "Hooks — automatic checks",
      "Don’t learn these before Module 7 works",
    ]), { x: 0.5, y: 1.2, w: 9, h: 3.5, fontSize: 20, color: P.ink, fontFace: "Arial" });
    footer(s);
  }
  {
    const s = content("You still own the outcome");
    s.addText(bullets([
      "AI can be wrong or insecure",
      "Don’t collect data you don’t need",
      "Explain your app in demos — don’t just click",
      "Fundamentals still matter long-term",
    ]), { x: 0.5, y: 1.2, w: 9, h: 3.5, fontSize: 20, color: P.ink, fontFace: "Arial" });
    footer(s);
  }

  // ========== TOOLS ==========
  section("REFERENCE", "Everyday Tools", "Stuck protocol · Commands · Mindset");
  {
    const s = content("3-minute stuck protocol");
    const steps = [
      "Say what you tried and expected",
      "Copy exact error text",
      "Paste bug report to Claude",
      "Still stuck? Ask a human",
    ];
    steps.forEach((t, i) => {
      const x = 0.45 + i * 2.4;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.8, w: 2.25, h: 2.2, fill: { color: i === 3 ? P.dark : P.white }, rectRadius: 0.1,
      });
      s.addText(String(i + 1), {
        x, y: 2.1, w: 2.25, h: 0.5, fontSize: 24, bold: true,
        color: i === 3 ? P.mint : P.accent, fontFace: "Arial", align: "center", margin: 0,
      });
      s.addText(t, {
        x: x + 0.15, y: 2.7, w: 1.95, h: 1.0, fontSize: 13,
        color: i === 3 ? P.white : P.ink, fontFace: "Arial", align: "center", margin: 0,
      });
    });
    footer(s);
  }
  {
    const s = content("Daily builder checklist");
    s.addText(bullets([
      "I know which folder I’m in",
      "Claude started from that folder",
      "Prompt has Purpose + Musts + Proof",
      "Big change → /plan first",
      "I ran/opened the app myself",
      "I tried one edge case",
      "I committed if it works",
    ]), { x: 0.5, y: 1.05, w: 9, h: 3.9, fontSize: 18, color: P.ink, fontFace: "Arial" });
    footer(s);
  }

  // ========== CLOSING ==========
  {
    const s = pres.addSlide();
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: P.dark } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: P.accent } });
    s.addText("You can build software.", {
      x: 0.7, y: 1.8, w: 8.6, h: 0.8, fontSize: 36, bold: true, color: P.white, fontFace: "Arial", margin: 0,
    });
    s.addText("Idea → PRD → Claude Code → working app → public URL → clear demo.", {
      x: 0.7, y: 2.8, w: 8.6, h: 0.7, fontSize: 18, color: P.mint, fontFace: "Arial", margin: 0,
    });
    s.addText("Teach patiently. Verify always. Ship something real.", {
      x: 0.7, y: 4.0, w: 8.6, h: 0.5, fontSize: 16, color: P.cream, fontFace: "Arial", margin: 0,
    });
  }

  const out = path.join("/home/shailesh/Downloads/Training/vibe-coding", "Vibe_Coding_Module_Slides.pptx");
  await pres.writeFile({ fileName: out });
  console.log("Wrote", out);
  console.log("Slides:", pres._slides ? pres._slides.length : "(see file)");
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
