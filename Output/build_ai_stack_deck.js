const fs = require("fs");
const path = require("path");

const OUT = path.join("Output", "AI_Stack_Presentation_Deck.pptx");
const W = 13.333333;
const H = 7.5;
const EMU = 914400;
const SLIDE_CX = 12192000;
const SLIDE_CY = 6858000;
const colors = {
  bg: "0D1117",
  surface: "161B22",
  surface2: "21262D",
  panel: "0B1220",
  text: "E6EDF3",
  muted: "A7B2C3",
  dim: "6B7280",
  border: "334155",
  cyan: "38BDF8",
  teal: "2DD4BF",
  violet: "A78BFA",
  indigo: "818CF8",
  rose: "FB7185",
  amber: "FBBF24",
};
const font = "Aptos";
const fontDisplay = "Aptos Display";
const NS = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
};

function esc(s) {
  return String(s == null ? "" : s)
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
function trunc(s, n) {
  s = String(s || "").replace(/\s+/g, " ").trim();
  return s.length <= n ? s : s.slice(0, Math.max(0, n - 1)).replace(/\s+\S*$/, "") + "...";
}
function normArray(v) {
  if (!v) return [];
  return Array.isArray(v) ? v.filter(Boolean).map(String) : [String(v)];
}
function inch(v) {
  return Math.round(v * EMU);
}
function hex(v) {
  return String(v || "").replace("#", "").toUpperCase();
}
function exists(file) {
  try {
    return fs.existsSync(file);
  } catch {
    return false;
  }
}
function imgDim(file) {
  const b = fs.readFileSync(file);
  if (b.length > 24 && b[0] === 0x89 && b.slice(1, 4).toString() === "PNG") {
    return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
  }
  if (b[0] === 0xff && b[1] === 0xd8) {
    let i = 2;
    while (i < b.length) {
      if (b[i] !== 0xff) break;
      const marker = b[i + 1];
      const len = b.readUInt16BE(i + 2);
      if (marker >= 0xc0 && marker <= 0xc3) {
        return { h: b.readUInt16BE(i + 5), w: b.readUInt16BE(i + 7) };
      }
      i += 2 + len;
    }
  }
  return { w: 1600, h: 900 };
}
function fitContain(file, x, y, w, h) {
  const d = imgDim(file);
  const r = Math.min(w / d.w, h / d.h);
  const ww = d.w * r;
  const hh = d.h * r;
  return { x: x + (w - ww) / 2, y: y + (h - hh) / 2, w: ww, h: hh };
}

function crc32(buf) {
  let table = crc32.table;
  if (!table) {
    table = crc32.table = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[i] = c >>> 0;
    }
  }
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = table[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}
function dosTimeDate(date = new Date()) {
  return {
    time: (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2),
    date: ((date.getFullYear() - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
  };
}
function makeZip(entries) {
  const local = [];
  const central = [];
  let offset = 0;
  const td = dosTimeDate();
  for (const e of entries) {
    const name = Buffer.from(e.name, "utf8");
    const data = Buffer.isBuffer(e.data) ? e.data : Buffer.from(String(e.data), "utf8");
    const crc = crc32(data);
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0);
    lh.writeUInt16LE(20, 4);
    lh.writeUInt16LE(0x0800, 6);
    lh.writeUInt16LE(0, 8);
    lh.writeUInt16LE(td.time, 10);
    lh.writeUInt16LE(td.date, 12);
    lh.writeUInt32LE(crc, 14);
    lh.writeUInt32LE(data.length, 18);
    lh.writeUInt32LE(data.length, 22);
    lh.writeUInt16LE(name.length, 26);
    lh.writeUInt16LE(0, 28);
    local.push(lh, name, data);

    const cd = Buffer.alloc(46);
    cd.writeUInt32LE(0x02014b50, 0);
    cd.writeUInt16LE(20, 4);
    cd.writeUInt16LE(20, 6);
    cd.writeUInt16LE(0x0800, 8);
    cd.writeUInt16LE(0, 10);
    cd.writeUInt16LE(td.time, 12);
    cd.writeUInt16LE(td.date, 14);
    cd.writeUInt32LE(crc, 16);
    cd.writeUInt32LE(data.length, 20);
    cd.writeUInt32LE(data.length, 24);
    cd.writeUInt16LE(name.length, 28);
    cd.writeUInt16LE(0, 30);
    cd.writeUInt16LE(0, 32);
    cd.writeUInt16LE(0, 34);
    cd.writeUInt16LE(0, 36);
    cd.writeUInt32LE(0, 38);
    cd.writeUInt32LE(offset, 42);
    central.push(cd, name);
    offset += lh.length + name.length + data.length;
  }
  const centralSize = central.reduce((s, b) => s + b.length, 0);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralSize, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(0, 20);
  return Buffer.concat([...local, ...central, eocd]);
}

function extractData() {
  const html = fs.readFileSync("AI_Tools_Interactive.html", "utf8");
  const m = html.match(/<script>\s*window\.__AI_TOOLS_DATA__ = \(function \(\) \{[\s\S]*?\}\)\(\);\s*<\/script>/);
  if (!m) throw new Error("Could not find site data in HTML");
  const sandbox = {};
  Function("window", m[0].replace(/^<script>/, "").replace(/<\/script>$/, ""))(sandbox);
  return sandbox.__AI_TOOLS_DATA__;
}

let shapeId = 1;
let mediaCounter = 1;
const media = new Map();
const slides = [];
function nextId() {
  return shapeId++;
}
function newSlide() {
  shapeId = 2;
  const slide = { xml: [], rels: [], relCounter: 1, notes: "" };
  addRect(slide, 0, 0, 13.333333, 7.5, colors.bg, null, "rect");
  slides.push(slide);
  return slide;
}
function addRelImage(slide, file) {
  const key = path.resolve(file);
  let item = media.get(key);
  if (!item) {
    const extRaw = path.extname(file).toLowerCase().replace(".", "") || "png";
    const ext = extRaw === "jpeg" ? "jpg" : extRaw;
    const mediaName = `image${mediaCounter++}.${ext}`;
    item = {
      target: `../media/${mediaName}`,
      name: `ppt/media/${mediaName}`,
      data: fs.readFileSync(file),
      ext: ext === "jpg" ? "jpeg" : ext,
    };
    media.set(key, item);
  }
  const rid = `rId${slide.relCounter++}`;
  slide.rels.push(
    `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${item.target}"/>`
  );
  return rid;
}
function spPr(x, y, w, h, fill, line, prst = "rect") {
  const fillXml = fill ? `<a:solidFill><a:srgbClr val="${hex(fill)}"/></a:solidFill>` : "<a:noFill/>";
  const lineXml = line
    ? `<a:ln w="${Math.round((line.w || 1) * 12700)}"><a:solidFill><a:srgbClr val="${hex(line.color || colors.border)}"/></a:solidFill></a:ln>`
    : "<a:ln><a:noFill/></a:ln>";
  return `<p:spPr><a:xfrm><a:off x="${inch(x)}" y="${inch(y)}"/><a:ext cx="${inch(w)}" cy="${inch(h)}"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${fillXml}${lineXml}</p:spPr>`;
}
function addRect(slide, x, y, w, h, fill, line, prst = "roundRect") {
  const id = nextId();
  slide.xml.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>${spPr(x, y, w, h, fill, line, prst)}<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`);
}
function runXml(text, opts) {
  opts = opts || {};
  const sz = Math.round((opts.size || 14) * 100);
  const bold = opts.bold ? ' b="1"' : "";
  const italic = opts.italic ? ' i="1"' : "";
  return `<a:r><a:rPr lang="en-US" sz="${sz}"${bold}${italic}><a:solidFill><a:srgbClr val="${hex(opts.color || colors.text)}"/></a:solidFill><a:latin typeface="${esc(opts.display ? fontDisplay : font)}"/><a:ea typeface="${esc(font)}"/><a:cs typeface="${esc(font)}"/></a:rPr><a:t>${esc(text)}</a:t></a:r>`;
}
function paraXml(text, opts) {
  opts = opts || {};
  return `<a:p><a:pPr algn="${opts.align || "l"}"/>${runXml(text, opts)}<a:endParaRPr lang="en-US" sz="${Math.round((opts.size || 14) * 100)}"/></a:p>`;
}
function addText(slide, x, y, w, h, text, opts = {}) {
  const id = nextId();
  const lines = Array.isArray(text) ? text : String(text || "").split(/\n/);
  const paras = lines.map((line) => (typeof line === "object" ? paraXml(line.text, { ...opts, ...line }) : paraXml(line, opts))).join("");
  const bodyPr = `<a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="${opts.valign || "t"}"><a:normAutofit fontScale="70000" lnSpcReduction="12000"/></a:bodyPr>`;
  slide.xml.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>${spPr(x, y, w, h, null, null, "rect")}<p:txBody>${bodyPr}<a:lstStyle/>${paras}</p:txBody></p:sp>`);
}
function addTitle(slide, title, subtitle, accent = colors.cyan) {
  addText(slide, 0.55, 0.38, 9.4, 0.25, subtitle || "The AI Stack", { size: 9.5, color: accent, bold: true });
  addText(slide, 0.55, 0.68, 9.7, 0.68, title, { size: 34, color: colors.text, bold: true, display: true });
  addRect(slide, 0.55, 1.45, 1.15, 0.035, accent, null, "rect");
}
function addFooter(slide, label = "AI Stack | Tools for Advertising & Marketing") {
  addText(slide, 0.55, 7.08, 8, 0.18, label, { size: 8.5, color: colors.dim });
  addText(slide, 11.2, 7.08, 1.55, 0.18, "Kyle Zemba", { size: 8.5, color: colors.dim, align: "r" });
}
function setNotes(slide, text) {
  slide.notes = String(text || "").replace(/\r\n/g, "\n").trim();
}
function notesList(label, items) {
  const arr = normArray(items);
  if (!arr.length) return [label + ":", "- Not specified."];
  return [label + ":"].concat(arr.map((item) => "- " + item));
}
function toolNotes(cat, tool) {
  return [
    `Category: ${cat.label}`,
    `Tool: ${tool.name}`,
    `Company: ${tool.company || "Not specified"}`,
    `URL: ${tool.url || (tool.domain ? "https://" + tool.domain : "Not specified")}`,
    "",
    `Tagline: ${tool.tagline || "Not specified"}`,
    "",
    "What it is:",
    tool.whatItIs || "Not specified.",
    "",
    ...notesList("Best For", tool.bestFor),
    "",
    ...notesList("Pros", tool.pros),
    "",
    ...notesList("Cons", tool.cons),
    "",
    ...notesList("How to Use It", tool.howToUse),
    "",
    ...notesList("Highlight Features", tool.highlights),
    "",
    `Pricing: ${tool.pricing || "Check current plans before recommendation."}`,
  ].join("\n");
}
function addImage(slide, file, x, y, w, h) {
  if (!exists(file)) return;
  const rid = addRelImage(slide, file);
  const id = nextId();
  slide.xml.push(`<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(path.basename(file))}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${inch(x)}" y="${inch(y)}"/><a:ext cx="${inch(w)}" cy="${inch(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`);
}
function addContainedImage(slide, file, x, y, w, h) {
  if (!exists(file)) return;
  const r = fitContain(file, x, y, w, h);
  addImage(slide, file, r.x, r.y, r.w, r.h);
}
function addChip(slide, x, y, text, accent, w) {
  w = w || Math.min(2.2, 0.28 + String(text).length * 0.082);
  addRect(slide, x, y, w, 0.28, colors.surface2, { color: accent, w: 0.75 }, "roundRect");
  addText(slide, x + 0.08, y + 0.06, w - 0.16, 0.12, text, { size: 7.5, color: colors.text, bold: true, valign: "mid" });
  return w;
}
function panel(slide, x, y, w, h, title, body, accent, opts = {}) {
  addRect(slide, x, y, w, h, opts.fill || colors.surface, { color: opts.line || colors.border, w: 0.8 }, "roundRect");
  if (title) addText(slide, x + 0.18, y + 0.18, w - 0.36, 0.22, title.toUpperCase(), { size: 8.5, color: accent, bold: true });
  const content = Array.isArray(body) ? body : [body];
  addText(slide, x + 0.18, y + (title ? 0.5 : 0.18), w - 0.36, h - (title ? 0.62 : 0.28), content.filter(Boolean).join("\n"), { size: opts.size || 10.5, color: opts.color || colors.muted });
}
function addSectionDivider(label, subtitle, tools, accent) {
  const s = newSlide();
  addRect(s, 0, 0, W, H, colors.bg, null, "rect");
  addRect(s, 0, 0, 0.14, H, accent, null, "rect");
  addText(s, 0.62, 0.72, 3.2, 0.22, "SECTION", { size: 9, color: accent, bold: true });
  addText(s, 0.62, 1.16, 7.4, 0.85, label, { size: 42, color: colors.text, bold: true, display: true });
  addText(s, 0.64, 2.2, 6.8, 0.8, subtitle || "", { size: 15, color: colors.muted });
  addRect(s, 8.0, 0.72, 4.65, 5.85, colors.surface, { color: accent, w: 1.1 }, "roundRect");
  addText(s, 8.35, 1.05, 3.8, 0.28, `${tools.length} tools in this section`, { size: 12, color: accent, bold: true });
  let y = 1.52;
  tools.slice(0, 12).forEach((t, i) => {
    addText(s, 8.35, y, 3.75, 0.23, `${String(i + 1).padStart(2, "0")}  ${t.name}`, { size: 12.5, color: colors.text, bold: i === 0 });
    y += 0.38;
  });
  if (tools.length > 12) addText(s, 8.35, y, 3.75, 0.25, `+ ${tools.length - 12} more`, { size: 11, color: colors.dim });
  addFooter(s);
  setNotes(s, [
    `Section: ${label}`,
    "",
    subtitle || "",
    "",
    "Tools in this section:",
    ...tools.map((t, i) => `${i + 1}. ${t.name}${t.company ? " | " + t.company : ""}`),
  ].join("\n"));
}

const data = extractData();
const categorySubtitles = {
  llms: "Strategic thinking, synthesis, writing, and source interpretation.",
  imgvid: "Image, motion, model comparison, and visual territory exploration.",
  allinone: "Connected creative systems for proof-of-concept production.",
  spatial: "3D assets, materials, worlds, and immersive production tests.",
  audio: "Voice, music, sonic mood, and localization support.",
  copy: "Repeatable messaging, content variation, and brand-language systems.",
  research: "Source-backed discovery, category scans, and insight development.",
  workflow: "Production systems, coding support, automation, and agent workflows.",
  design: "Decks, layouts, prototypes, and visual storytelling.",
  marketing: "CRM, lifecycle, personalization, and measurable activation.",
};
const workflowSlides = [
  { step: "01", title: "Train Model", method: "ChatGPT 5.5 Thinking | Context Engineering", img: "Context/Example Workflow/Step 1 - AI Context Workflow.png", body: "The workflow starts by training the conversation with the concept, product, audience, creative intent, medium, constraints, and desired output. The model becomes more useful because the thinking is already framed.", tags: ["Concept definition", "Commercial medium", "Product constraints", "Creative direction"] },
  { step: "02", title: "Research", method: "ChatGPT Research + Tool Selection", img: "Context/Example Workflow/Step 2 - Higgsfield One Canvas.png", body: "The tool choice came after the project was framed. Higgsfield Canvas became the center because it could connect prompts, visual references, image generation, and video generation into one production map.", tags: ["Tool landscape", "Workflow examples", "Visual references", "Higgsfield selected"] },
  { step: "03", title: "Structure", method: "Workflow Architecture", img: "Context/Example Workflow/Step 3.png", body: "The production path was mapped before generating media: brief, mood board, concept, storyboard, shot table, keyframes, final prompt, and video generation.", tags: ["Brief", "Mood board", "Storyboard", "Shot table", "Keyframes", "Video prompt"] },
  { step: "04", title: "Build", method: "ChatGPT + Claude + Gemini", img: "Context/Example Workflow/Step 4 - Build Nodes.png", body: "Each node received a clear job. Thinking models helped define the brand, analyze mood, create the concept, extract the storyboard, format the shot table, and compile generation prompts.", tags: ["Prompt systems", "Node instructions", "Structured outputs", "Model collaboration"] },
  { step: "05", title: "Visualize", method: "Reference Images + Visual Anchors", img: "Context/Example Workflow/ChatGPT Image May 10, 2026, 01_48_18 PM.png", body: "Visual anchors helped lock the tone, characters, product, and cinematic language before video generation. The goal was consistency, not random output.", tags: ["Mood board", "Cleat reference", "Hero athlete", "Goalie reference", "Storyboard sheet"] },
  { step: "06", title: "Execute", method: "Higgsfield Canvas", img: "Context/Example Workflow/Step 6.png", body: "The final Canvas execution used the compiled prompt, storyboard sheet, character references, and product references as generation context.", tags: ["Canvas execution", "Node workflow", "Prompt + references", "Video generation"] },
  { step: "07", title: "Results", method: "Generated Commercial Output", img: "Context/Example Workflow/Results Video Poster 2s.jpg", body: "The result is a cinematic soccer commercial concept built from a structured AI workflow rather than a single prompt. The process is repeatable and explainable.", tags: ["Final video", "Repeatable workflow", "Commercial concept", "AI production pipeline"] },
];

function buildOpeningSlides() {
  let s = newSlide();
  addImage(s, "Context/Wireframe/AI Stack Hero Network.png", 0, 0, W, H);
  addRect(s, 0, 0, W, H, "050914", null, "rect");
  addRect(s, 0.55, 0.58, 0.12, 5.95, colors.cyan, null, "rect");
  addText(s, 0.9, 0.72, 7.6, 0.95, "The AI Stack", { size: 54, color: colors.text, bold: true, display: true });
  addText(s, 0.94, 1.68, 7.0, 0.45, "Tools for Advertising & Marketing", { size: 24, color: "93C5FD", bold: true });
  addText(s, 0.96, 2.46, 6.9, 0.9, "A presentation-ready map for understanding where AI belongs across strategy, creative development, production, workflow, and activation.", { size: 17, color: colors.muted });
  addText(s, 0.96, 6.56, 7.0, 0.25, "Built from the interactive AI Stack site | May 2026", { size: 10.5, color: colors.dim });

  s = newSlide();
  addTitle(s, "Not a tool directory. A workflow map.", "Opening frame", colors.cyan);
  addText(s, 0.6, 1.75, 6.2, 0.75, "The point is not to prove awareness of AI tools. The point is to show judgment: which tool belongs where, what it is good for, where it breaks down, and how it fits into advertising and marketing work.", { size: 18, color: colors.text });
  [
    ["01", "Strategy before output", "AI performs better when the brief, audience, channel, and standard are clear before generation starts.", colors.indigo],
    ["02", "Tools as a stack", "The strongest workflows move between thinking, image, video, audio, design, coding, and operations.", colors.teal],
    ["03", "Production reality", "The deck includes concepting tools, but also the systems that help teams scale, review, version, and ship creative work.", colors.amber],
  ].forEach((c, i) => {
    const x = 0.6 + i * 4.1;
    addRect(s, x, 3.28, 3.65, 2.25, colors.surface, { color: c[3], w: 1 }, "roundRect");
    addText(s, x + 0.24, 3.55, 0.5, 0.2, c[0], { size: 11, color: c[3], bold: true });
    addText(s, x + 0.24, 3.88, 3.1, 0.32, c[1], { size: 17, color: colors.text, bold: true });
    addText(s, x + 0.24, 4.34, 3.08, 0.86, c[2], { size: 11.5, color: colors.muted });
  });
  addFooter(s);

  s = newSlide();
  addTitle(s, "Better context creates better work.", "A note before the tools", colors.teal);
  addText(s, 0.6, 1.65, 5.1, 0.65, "AI is not a replacement for strategy. It is a multiplier for clear direction.", { size: 19, color: colors.text, bold: true });
  addText(s, 0.6, 2.45, 5.3, 0.9, "For advertising and marketing work, the advantage comes from the thinking before the output: the audience, the job to be done, the channel, the brand boundaries, and the creative standard you are trying to reach.", { size: 13.5, color: colors.muted });
  [
    ["Context first", "Give the tool the same setup you would give a strategist or creative partner."],
    ["Clarify before making", "Ask the model to restate the assignment and identify gaps before it produces work."],
    ["Connect the tools", "Strong workflows move across text, image, video, audio, and design."],
    ["Iterate with memory", "Reuse brand context, refine through feedback, and build from what the tool has learned."],
  ].forEach((c, i) => {
    const x = 0.6 + (i % 2) * 2.7;
    const y = 4.05 + Math.floor(i / 2) * 1.15;
    addRect(s, x, y, 2.45, 0.9, colors.surface, { color: colors.border, w: 0.7 }, "roundRect");
    addText(s, x + 0.15, y + 0.15, 2.1, 0.2, String(i + 1).padStart(2, "0"), { size: 8.5, color: colors.indigo, bold: true });
    addText(s, x + 0.15, y + 0.38, 2.1, 0.18, c[0], { size: 10.8, color: colors.text, bold: true });
    addText(s, x + 0.15, y + 0.6, 2.1, 0.22, c[1], { size: 8.2, color: colors.muted });
  });
  addRect(s, 6.45, 1.24, 6.25, 4.35, colors.surface, { color: colors.border, w: 0.8 }, "roundRect");
  addContainedImage(s, "Context/Wireframe/Better Context Better Answers - Edited.png", 6.62, 1.4, 5.9, 4.02);
  addFooter(s);

  s = newSlide();
  addTitle(s, "The full landscape, at a glance.", "Ecosystem map", colors.indigo);
  const grid = data.categories.concat([{ id: "watch", label: "Watch List / Emerging", accent: "#6B7280", tools: data.watchlist.map((w) => ({ name: w.name })) }]);
  const positions = [
    [0.55, 1.55, 3.85, 1.22], [4.55, 1.55, 8.15, 1.22],
    [0.55, 2.98, 3.85, 1.05], [4.55, 2.98, 3.85, 1.05], [8.55, 2.98, 4.15, 1.05],
    [0.55, 4.25, 3.85, 1.05], [4.55, 4.25, 3.85, 1.05], [8.55, 4.25, 4.15, 1.05],
    [0.55, 5.52, 3.85, 1.05], [4.55, 5.52, 3.85, 1.05], [8.55, 5.52, 4.15, 1.05],
  ];
  grid.forEach((cat, i) => {
    const [x, y, w, h] = positions[i];
    const accent = hex(cat.accent || "#6B7280");
    addRect(s, x, y, w, h, colors.surface, { color: accent, w: 0.75 }, "roundRect");
    addText(s, x + 0.14, y + 0.12, w - 0.3, 0.16, cat.label.toUpperCase(), { size: 7.4, color: accent, bold: true });
    addText(s, x + 0.14, y + 0.42, w - 0.26, h - 0.48, cat.tools.map((t) => t.name).slice(0, i === 1 ? 14 : 8).join("  |  "), { size: i === 1 ? 8.3 : 8.0, color: colors.text });
  });
  addFooter(s);
}

function buildWorkflowSlides() {
  let s = newSlide();
  addTitle(s, "Example AI workflow: from context to commercial.", "Workflow", colors.teal);
  addText(s, 0.6, 1.56, 6.9, 0.5, "A real proof-of-concept workflow for a Nike soccer cleat commercial. The important part is the sequence: strategy and structure before generation.", { size: 14.5, color: colors.muted });
  workflowSlides.forEach((wfl, i) => {
    const x = 0.7 + (i % 4) * 3.1;
    const y = 2.42 + Math.floor(i / 4) * 1.55;
    addRect(s, x, y, 2.7, 1.18, colors.surface, { color: i === 6 ? colors.amber : colors.teal, w: 0.8 }, "roundRect");
    addText(s, x + 0.16, y + 0.14, 0.5, 0.16, wfl.step, { size: 8.5, color: colors.teal, bold: true });
    addText(s, x + 0.16, y + 0.42, 2.25, 0.2, wfl.title, { size: 13.5, color: colors.text, bold: true });
    addText(s, x + 0.16, y + 0.72, 2.25, 0.25, wfl.method, { size: 8.5, color: colors.muted });
  });
  addFooter(s);

  workflowSlides.forEach((wfl) => {
    const slide = newSlide();
    addTitle(slide, `${wfl.step}. ${wfl.title}`, "Workflow detail", colors.teal);
    addText(slide, 0.62, 1.55, 4.8, 0.25, wfl.method, { size: 12.5, color: colors.teal, bold: true });
    addText(slide, 0.62, 2.05, 4.9, 1.05, wfl.body, { size: 14.2, color: colors.text });
    let cx = 0.62;
    let cy = 3.58;
    wfl.tags.forEach((tag) => {
      const ww = addChip(slide, cx, cy, tag, colors.teal);
      cx += ww + 0.12;
      if (cx > 4.9) {
        cx = 0.62;
        cy += 0.38;
      }
    });
    addRect(slide, 5.95, 1.32, 6.75, 4.9, colors.surface, { color: colors.border, w: 0.8 }, "roundRect");
    addContainedImage(slide, wfl.img, 6.12, 1.5, 6.4, 4.55);
    addFooter(slide);
    setNotes(slide, [
      `Workflow step ${wfl.step}: ${wfl.title}`,
      `Method: ${wfl.method}`,
      "",
      wfl.body,
      "",
      "Outputs and tags:",
      ...wfl.tags.map((tag) => "- " + tag),
    ].join("\n"));
  });
}

function toolSlide(cat, tool, idx) {
  const accent = hex(cat.accent || colors.cyan);
  const s = newSlide();
  addText(s, 0.55, 0.36, 5.5, 0.2, cat.label.toUpperCase(), { size: 8.5, color: accent, bold: true });
  addText(s, 0.55, 0.74, 6.3, 0.55, tool.name, { size: 31, color: colors.text, bold: true, display: true });
  addText(s, 0.58, 1.28, 4.4, 0.22, `${tool.company || ""}${tool.parentLabel ? " | " + tool.parentLabel : ""}`, { size: 10.5, color: colors.dim });
  addRect(s, 10.85, 0.48, 1.75, 0.42, accent, null, "roundRect");
  addText(s, 10.98, 0.61, 1.48, 0.12, `TOOL ${String(idx + 1).padStart(2, "0")}`, { size: 8, color: colors.bg, bold: true, align: "c" });
  addText(s, 0.62, 1.72, 6.5, 0.5, trunc(tool.tagline, 145), { size: 17, color: accent, bold: true });
  panel(s, 0.62, 2.48, 5.9, 1.35, "What it is", trunc(tool.whatItIs, 360), accent, { size: 10.6 });
  panel(s, 6.82, 1.58, 2.85, 1.5, "Best for", normArray(tool.bestFor).slice(0, 2).map((b) => "- " + trunc(b, 115)), accent, { size: 9.2 });
  panel(s, 9.9, 1.58, 2.8, 1.5, "How to use it", normArray(tool.howToUse).slice(0, 2).map((b) => "- " + trunc(b, 115)), accent, { size: 9.2 });
  panel(s, 0.62, 4.15, 3.8, 1.55, "Pros", normArray(tool.pros).slice(0, 3).map((b) => "- " + trunc(b, 95)), accent, { size: 8.7 });
  panel(s, 4.68, 4.15, 3.8, 1.55, "Watchouts", normArray(tool.cons).slice(0, 3).map((b) => "- " + trunc(b, 95)), accent, { size: 8.7 });
  const features = normArray(tool.highlights).slice(0, 3).map((b) => "- " + trunc(b, 92));
  panel(s, 8.74, 4.15, 3.95, 1.55, "Highlights", features.length ? features : ["- Best results come from a clear role in the workflow."], accent, { size: 8.7 });
  addRect(s, 0.62, 6.05, 12.08, 0.55, colors.panel, { color: colors.border, w: 0.6 }, "roundRect");
  addText(s, 0.84, 6.23, 7.6, 0.12, `Pricing: ${trunc(tool.pricing || "Check current plans before recommending.", 160)}`, { size: 8.6, color: colors.muted });
  addText(s, 8.65, 6.23, 3.85, 0.12, trunc(tool.url || ("https://" + tool.domain), 70), { size: 8.6, color: accent, align: "r" });
  addFooter(s, `AI Stack | ${cat.label}`);
  setNotes(s, toolNotes(cat, tool));
}

function buildToolSlides() {
  data.categories.forEach((cat) => {
    addSectionDivider(cat.label, cat.subtitle || categorySubtitles[cat.id], cat.tools, hex(cat.accent || colors.cyan));
    cat.tools.forEach((tool, idx) => toolSlide(cat, tool, idx));
  });

  addSectionDivider("Watch List / Emerging", "Signals, tools, and categories worth tracking as AI workflows mature.", data.watchlist, colors.dim);
  data.watchlist.forEach((item) => {
    const s = newSlide();
    addTitle(s, item.name, "Watch List / Emerging", colors.dim);
    addText(s, 0.62, 1.46, 4.3, 0.24, item.company || item.status || "Emerging", { size: 12, color: colors.dim, bold: true });
    addText(s, 0.62, 2.0, 5.9, 1.35, normArray(item.body).join("\n"), { size: 16, color: colors.text });
    panel(s, 7.1, 1.7, 4.8, 1.1, "Status", item.status || "Watch with caveats", colors.dim, { size: 14, color: colors.text });
    panel(s, 7.1, 3.15, 4.8, 1.35, "Why it matters", "Emerging tools are included to show where the landscape is moving, not because every item is ready for mainstream client workflow adoption.", colors.dim, { size: 11.5 });
    addRect(s, 7.1, 4.95, 4.8, 0.62, colors.surface, { color: colors.border, w: 0.8 }, "roundRect");
    addText(s, 7.35, 5.16, 4.25, 0.13, item.url || "", { size: 10, color: colors.muted, align: "c" });
    addFooter(s, "AI Stack | Watch List");
    setNotes(s, [
      `Watch List: ${item.name}`,
      `Company: ${item.company || "Not specified"}`,
      `Status: ${item.status || "Not specified"}`,
      `URL: ${item.url || "Not specified"}`,
      "",
      "Context:",
      ...normArray(item.body),
    ].join("\n"));
  });
}

function buildClosingSlide() {
  const s = newSlide();
  addRect(s, 0, 0, W, H, colors.bg, null, "rect");
  addText(s, 0.7, 1.3, 9.2, 0.9, "AI fluency is workflow fluency.", { size: 46, color: colors.text, bold: true, display: true });
  addText(s, 0.72, 2.45, 7.6, 0.75, "The strongest teams do not chase every new model. They understand the job, choose the right tool, structure the workflow, and keep human judgment in the loop.", { size: 19, color: colors.muted });
  addRect(s, 0.72, 4.35, 3.2, 0.055, colors.teal, null, "rect");
  addText(s, 0.72, 5.2, 5.0, 0.35, "The AI Stack | Tools for Advertising & Marketing", { size: 16, color: colors.teal, bold: true });
  addText(s, 0.72, 5.66, 4.2, 0.22, "Kyle Zemba | Link To VR | MBA candidate, Western New England University", { size: 10.5, color: colors.dim });
}

function slideXml(slide) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${slide.xml.join("")}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}
function notesRunXml(text, opts = {}) {
  const sz = Math.round((opts.size || 10) * 100);
  return `<a:r><a:rPr lang="en-US" sz="${sz}"><a:solidFill><a:srgbClr val="${hex(opts.color || "000000")}"/></a:solidFill><a:latin typeface="${esc(font)}"/><a:ea typeface="${esc(font)}"/><a:cs typeface="${esc(font)}"/></a:rPr><a:t>${esc(text)}</a:t></a:r>`;
}
function notesParaXml(text, opts = {}) {
  return `<a:p><a:pPr algn="l"/>${notesRunXml(text, opts)}<a:endParaRPr lang="en-US" sz="${Math.round((opts.size || 10) * 100)}"/></a:p>`;
}
function notesSlideXml(notes) {
  const lines = String(notes || "").split("\n");
  const paras = lines.map((line) => notesParaXml(line, { size: line.endsWith(":") || /^Category:|^Tool:|^Company:|^URL:|^Tagline:|^Pricing:|^Section:|^Workflow step|^Watch List:|^Status:/.test(line) ? 10.5 : 9.4 })).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notes xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${inch(0.5)}" y="${inch(0.45)}"/><a:ext cx="${inch(6.5)}" cy="${inch(9.05)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:normAutofit fontScale="80000" lnSpcReduction="12000"/></a:bodyPr><a:lstStyle/>${paras}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}
function relsXml(rels) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`;
}
function presentationXml() {
  const sldIds = slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join("");
  const notesMaster = slides.some((slide) => slide.notes) ? `<p:notesMasterIdLst><p:notesMasterId r:id="rId${slides.length + 2}"/></p:notesMasterIdLst>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>${notesMaster}<p:sldIdLst>${sldIds}</p:sldIdLst><p:sldSz cx="${SLIDE_CX}" cy="${SLIDE_CY}" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:defaultTextStyle></p:presentation>`;
}
function contentTypes() {
  const overrides = [
    '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
    '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>',
    '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
    '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
  ];
  slides.forEach((_, i) => overrides.push(`<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`));
  if (slides.some((slide) => slide.notes)) {
    overrides.push('<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>');
    slides.filter((slide) => slide.notes).forEach((_, i) => {
      overrides.push(`<Override PartName="/ppt/notesSlides/notesSlide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`);
    });
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="jpeg" ContentType="image/jpeg"/>${overrides.join("")}</Types>`;
}

function writeDeck() {
  const slideLayoutXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
  const slideMasterXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${colors.bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`;
  const notesMasterXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notesMaster xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr><a:defRPr sz="1000" lang="en-US"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="${esc(font)}"/></a:defRPr></a:lvl1pPr></p:notesStyle></p:notesMaster>`;
  const themeXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="${NS.a}" name="AI Stack"><a:themeElements><a:clrScheme name="AI Stack"><a:dk1><a:srgbClr val="${colors.bg}"/></a:dk1><a:lt1><a:srgbClr val="${colors.text}"/></a:lt1><a:dk2><a:srgbClr val="${colors.surface}"/></a:dk2><a:lt2><a:srgbClr val="${colors.muted}"/></a:lt2><a:accent1><a:srgbClr val="${colors.cyan}"/></a:accent1><a:accent2><a:srgbClr val="${colors.indigo}"/></a:accent2><a:accent3><a:srgbClr val="${colors.teal}"/></a:accent3><a:accent4><a:srgbClr val="${colors.amber}"/></a:accent4><a:accent5><a:srgbClr val="${colors.rose}"/></a:accent5><a:accent6><a:srgbClr val="${colors.violet}"/></a:accent6><a:hlink><a:srgbClr val="${colors.cyan}"/></a:hlink><a:folHlink><a:srgbClr val="${colors.violet}"/></a:folHlink></a:clrScheme><a:fontScheme name="AI Stack"><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="AI Stack"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
  const notesSlideItems = slides.map((slide, slideIndex) => ({ slide, slideIndex })).filter((item) => item.slide.notes).map((item, notesIndex) => ({ ...item, notesIndex }));
  notesSlideItems.forEach((item) => {
    const rid = `rId${item.slide.relCounter++}`;
    item.slide.rels.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide${item.notesIndex + 1}.xml"/>`);
  });
  const entries = [];
  entries.push({ name: "[Content_Types].xml", data: contentTypes() });
  entries.push({ name: "_rels/.rels", data: relsXml([
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>',
  ]) });
  const now = new Date().toISOString();
  entries.push({ name: "docProps/core.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>The AI Stack Presentation Deck</dc:title><dc:creator>Codex</dc:creator><cp:lastModifiedBy>Codex</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>` });
  entries.push({ name: "docProps/app.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Codex</Application><PresentationFormat>On-screen Show (16:9)</PresentationFormat><Slides>${slides.length}</Slides><Notes>${notesSlideItems.length}</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop></Properties>` });
  entries.push({ name: "ppt/presentation.xml", data: presentationXml() });
  const presRels = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'];
  slides.forEach((_, i) => presRels.push(`<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`));
  if (notesSlideItems.length) presRels.push(`<Relationship Id="rId${slides.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>`);
  entries.push({ name: "ppt/_rels/presentation.xml.rels", data: relsXml(presRels) });
  entries.push({ name: "ppt/slideMasters/slideMaster1.xml", data: slideMasterXml });
  entries.push({ name: "ppt/slideMasters/_rels/slideMaster1.xml.rels", data: relsXml([
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>',
  ]) });
  entries.push({ name: "ppt/slideLayouts/slideLayout1.xml", data: slideLayoutXml });
  entries.push({ name: "ppt/slideLayouts/_rels/slideLayout1.xml.rels", data: relsXml(['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>']) });
  entries.push({ name: "ppt/theme/theme1.xml", data: themeXml });
  if (notesSlideItems.length) {
    entries.push({ name: "ppt/notesMasters/notesMaster1.xml", data: notesMasterXml });
    entries.push({ name: "ppt/notesMasters/_rels/notesMaster1.xml.rels", data: relsXml(['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>']) });
  }
  slides.forEach((s, i) => {
    entries.push({ name: `ppt/slides/slide${i + 1}.xml`, data: slideXml(s) });
    if (s.rels.length) entries.push({ name: `ppt/slides/_rels/slide${i + 1}.xml.rels`, data: relsXml(s.rels) });
  });
  notesSlideItems.forEach((item) => {
    entries.push({ name: `ppt/notesSlides/notesSlide${item.notesIndex + 1}.xml`, data: notesSlideXml(item.slide.notes) });
    entries.push({ name: `ppt/notesSlides/_rels/notesSlide${item.notesIndex + 1}.xml.rels`, data: relsXml([
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide${item.slideIndex + 1}.xml"/>`,
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>',
    ]) });
  });
  for (const item of media.values()) entries.push({ name: item.name, data: item.data });
  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, makeZip(entries));
}

buildOpeningSlides();
buildWorkflowSlides();
buildToolSlides();
buildClosingSlide();
writeDeck();

console.log(JSON.stringify({
  output: OUT,
  slides: slides.length,
  tools: data.categories.reduce((n, c) => n + c.tools.length, 0),
  watchlist: data.watchlist.length,
  media: media.size,
  bytes: fs.statSync(OUT).size,
}, null, 2));
