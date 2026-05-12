const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

const OUT = path.join("Output", "AI_Stack_Source_Based_Presentation.pptx");
const W = 13.333333;
const H = 7.5;
const EMU = 914400;
const SLIDE_CX = 12192000;
const SLIDE_CY = 6858000;
const font = "Aptos";
const fontDisplay = "Aptos Display";

const colors = {
  bg: "0D1117",
  ink: "E6EDF3",
  muted: "A7B2C3",
  dim: "6B7280",
  panel: "161B22",
  panel2: "21262D",
  line: "334155",
  paper: "F8FAFC",
  paperInk: "111827",
  paperMuted: "475569",
  indigo: "818CF8",
  amber: "FBBF24",
  teal: "2DD4BF",
  blue: "60A5FA",
  rose: "FB7185",
  emerald: "34D399",
  violet: "A78BFA",
  slate: "94A3B8",
  orange: "FB923C",
  red: "F87171",
  gray: "6B7280",
};

const NS = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
};

const sourceFiles = {
  about: path.join("Context", "About Project.md"),
  page: path.join("Context", "Page Structure.md"),
  tone: path.join("Context", "Tone and Voice.md"),
  db: path.join("Context", "AI Tools Database.md"),
  philosophy: path.join("Context", "AI_Philosophy_Framework.docx"),
  mood: path.join("Context", "AI_Stack_Mood_Board.docx"),
  positioning: path.join("Context", "Tool_Positioning_Guide.docx"),
  brief: path.join("Output", "ChatGPT Prompting", "ChatGPT Project Brief \u2014 Updated.md"),
  ecosystemPrompt: path.join("Output", "ChatGPT Prompting", "GPT Image 2 \u2014 Ecosystem Map Prompt.md"),
  researchA: path.join("Research", "AI_Tools_Research_Brief___Agencies.md.docx"),
  researchB: path.join("Research", "AI_Tools_Research_Brief_Categories_4-7.docx"),
  researchC: path.join("Research", "AI_Tools_Research_Categories_8_9_10.docx"),
  rocketium: path.join("Research", "Rocketium_Codex_Handoff.docx"),
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

function decodeXml(s) {
  return String(s || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function cleanText(s) {
  return String(s || "")
    .replace(/\u00a0/g, " ")
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201c\u201d]/g, '"')
    .replace(/[\u2013\u2014]/g, " - ")
    .replace(/\u2026/g, "...")
    .replace(/\s+/g, " ")
    .replace(/\s+-\s+/g, " - ")
    .trim();
}

function trunc(s, n) {
  s = cleanText(s);
  if (s.length <= n) return s;
  return s.slice(0, Math.max(0, n - 1)).replace(/\s+\S*$/, "") + "...";
}

function stripStar(name) {
  return cleanText(name).replace(/\s*[★*]\s*$/g, "").trim();
}

function prettyToolName(name) {
  const raw = stripStar(name);
  const key = raw.toLowerCase();
  const known = {
    "claude": "Claude",
    "chatgpt": "ChatGPT",
    "gpt-5.5": "GPT-5.5",
    "gemini": "Gemini",
    "perplexity": "Perplexity",
    "grok": "Grok",
    "midjourney": "Midjourney",
    "flux (black forest labs)": "Flux (Black Forest Labs)",
    "dall-e 3": "DALL-E 3",
    "runway": "Runway",
    "sora": "Sora",
    "kling (kuaishou)": "Kling",
    "seedance (bytedance)": "Seedance",
    "luma dream machine": "Luma Dream Machine",
    "pika": "Pika",
    "veo 3 (google)": "Veo 3",
    "comfyui": "ComfyUI",
    "nano banana": "Nano Banana",
    "gpt image 2": "GPT Image 2",
    "higgsfield": "Higgsfield",
    "freepik": "Freepik",
    "artlist": "Artlist",
    "fal (with patina workflow)": "Fal / PATINA",
    "tripo 3d": "Tripo 3D",
    "world labs (marble)": "World Labs (Marble)",
    "google genie (project genie / genie 3)": "Google Genie",
    "elevenlabs": "ElevenLabs",
    "suno": "Suno",
    "udio": "Udio",
    "jasper": "Jasper",
    "copy.ai": "Copy.ai",
    "notebooklm (google)": "NotebookLM",
    "perplexity deep research": "Perplexity Deep Research",
    "chatgpt deep research": "ChatGPT Deep Research",
    "gemini deep research": "Gemini Deep Research",
    "github copilot": "GitHub Copilot",
    "cursor": "Cursor",
    "claude code": "Claude Code",
    "claude cowork": "Claude Cowork",
    "replit": "Replit",
    "canva ai (magic studio)": "Canva AI",
    "gamma": "Gamma",
    "figma (ai features)": "Figma",
    "salesforce einstein": "Salesforce Einstein",
    "hubspot breeze": "HubSpot Breeze",
    "heygen": "HeyGen",
    "rocketium": "Rocketium",
    "mythos (anthropic)": "Mythos",
    "ltx studio": "LTX Studio",
    "unbound (unboundcontent.ai)": "Unbound",
    "ai answer engine optimization (aeo)": "AI Answer Engine Optimization",
    "openclaw + nvidia nemoclaw": "OpenClaw + NVIDIA NemoClaw",
  };
  if (known[key]) return known[key];
  return raw.toLowerCase().replace(/\b\w/g, (c) => c.toUpperCase()).replace(/\bAi\b/g, "AI").replace(/\bCrm\b/g, "CRM").replace(/\b3d\b/gi, "3D");
}

function hex(v) {
  return String(v || "").replace("#", "").toUpperCase();
}

function inch(v) {
  return Math.round(v * EMU);
}

function normArray(v) {
  if (!v) return [];
  if (Array.isArray(v)) return v.map(cleanText).filter(Boolean);
  return [cleanText(v)].filter(Boolean);
}

function zipEntry(file, entryName) {
  const buf = fs.readFileSync(file);
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65558); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error("Could not find zip directory in " + file);
  const total = buf.readUInt16LE(eocd + 10);
  let off = buf.readUInt32LE(eocd + 16);
  for (let i = 0; i < total; i++) {
    if (buf.readUInt32LE(off) !== 0x02014b50) throw new Error("Bad central directory in " + file);
    const method = buf.readUInt16LE(off + 10);
    const compSize = buf.readUInt32LE(off + 20);
    const nameLen = buf.readUInt16LE(off + 28);
    const extraLen = buf.readUInt16LE(off + 30);
    const commentLen = buf.readUInt16LE(off + 32);
    const localOff = buf.readUInt32LE(off + 42);
    const name = buf.slice(off + 46, off + 46 + nameLen).toString("utf8");
    if (name === entryName) {
      const localNameLen = buf.readUInt16LE(localOff + 26);
      const localExtraLen = buf.readUInt16LE(localOff + 28);
      const dataStart = localOff + 30 + localNameLen + localExtraLen;
      const data = buf.slice(dataStart, dataStart + compSize);
      if (method === 0) return data;
      if (method === 8) return zlib.inflateRawSync(data);
      throw new Error("Unsupported zip compression method " + method);
    }
    off += 46 + nameLen + extraLen + commentLen;
  }
  throw new Error("Missing " + entryName + " in " + file);
}

function docxParagraphs(file) {
  const xml = zipEntry(file, "word/document.xml").toString("utf8");
  const paras = xml.match(/<w:p[\s\S]*?<\/w:p>/g) || [];
  return paras.map((p) => {
    const parts = [];
    p.replace(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g, (_, t) => {
      parts.push(decodeXml(t));
      return "";
    });
    p.replace(/<w:tab\/>/g, () => {
      parts.push(" ");
      return "";
    });
    return cleanText(parts.join(""));
  }).filter(Boolean);
}

function readMd(file) {
  return fs.readFileSync(file, "utf8").replace(/^\uFEFF/, "");
}

const FIELD_ALIASES = new Map([
  ["category", "categoryDetail"],
  ["url", "url"],
  ["tagline", "tagline"],
  ["what it is", "whatItIs"],
  ["best for", "bestFor"],
  ["pros", "pros"],
  ["cons", "cons"],
  ["how to use it", "howToUse"],
  ["how to use it:", "howToUse"],
  ["how to use it", "howToUse"],
  ["highlight features", "highlights"],
  ["pricing", "pricing"],
]);

function fieldKey(line) {
  const normalized = cleanText(line).replace(/:$/, "").toLowerCase();
  return FIELD_ALIASES.get(normalized) || null;
}

function categoryFromHeading(line) {
  const m = cleanText(line).match(/^CATEGORY\s+\d+:\s*(.+)$/i);
  return m ? titleCaseCategory(m[1]) : null;
}

function titleCaseCategory(s) {
  return cleanText(s)
    .replace(/\bAND\b/g, "and")
    .replace(/\bAI\b/g, "AI")
    .replace(/\bLLMS\b/g, "LLMs")
    .replace(/\bCRM\b/g, "CRM")
    .replace(/\b3D\b/g, "3D")
    .replace(/\s+\/\s+/g, " / ")
    .split(" ")
    .map((part) => {
      if (/^(AI|CRM|LLMs|3D|&|\/)$/i.test(part)) return part.replace(/^Llms$/i, "LLMs").replace(/^Crm$/i, "CRM").replace(/^Ai$/i, "AI");
      if (/^[A-Z]{2,}$/.test(part)) return part;
      return part.charAt(0).toUpperCase() + part.slice(1).toLowerCase();
    })
    .join(" ")
    .replace("And", "and");
}

function isIgnoredHeading(line) {
  return /^(AI TOOLS RESEARCH BRIEF|AI Tools Research Brief|Prepared|April\/May 2026|Research compiled)/i.test(cleanText(line));
}

function looksLikeToolStart(paras, i) {
  const line = cleanText(paras[i]);
  const next = cleanText(paras[i + 1] || "");
  if (!line || isIgnoredHeading(line)) return false;
  if (/^CATEGORY\s+\d+:/i.test(line) || /^WATCH LIST/i.test(line)) return false;
  if (/^(What it is|Best For|Pros|Cons|How to Use It|Highlight Features|Pricing|URL|Tagline|Category):?$/i.test(line)) return false;
  return /^Category:/i.test(next);
}

function parseResearchTools(paras, sourceName) {
  const tools = [];
  let sectionCategory = "";
  for (let i = 0; i < paras.length; i++) {
    const line = cleanText(paras[i]);
    const cat = categoryFromHeading(line);
    if (cat) {
      sectionCategory = cat;
      continue;
    }
    if (!looksLikeToolStart(paras, i)) continue;
    const tool = {
      name: prettyToolName(line),
      experienced: /[★*]/.test(line),
      sectionCategory,
      sourceName,
      categoryDetail: "",
      url: "",
      tagline: "",
      whatItIs: "",
      bestFor: [],
      pros: [],
      cons: [],
      howToUse: [],
      highlights: [],
      pricing: [],
    };
    let current = null;
    i++;
    for (; i < paras.length; i++) {
      const p = cleanText(paras[i]);
      if (categoryFromHeading(p) || /^WATCH LIST/i.test(p) || looksLikeToolStart(paras, i)) {
        i--;
        break;
      }
      const inline = p.match(/^([^:]{2,40}):\s*(.+)$/);
      if (inline && FIELD_ALIASES.has(inline[1].trim().toLowerCase())) {
        current = FIELD_ALIASES.get(inline[1].trim().toLowerCase());
        addField(tool, current, inline[2]);
        continue;
      }
      const fk = fieldKey(p);
      if (fk) {
        current = fk;
        continue;
      }
      if (current) addField(tool, current, p);
    }
    if (!tool.sectionCategory && tool.categoryDetail) tool.sectionCategory = tool.categoryDetail;
    tools.push(tool);
  }
  return tools;
}

function addField(tool, key, value) {
  value = cleanText(value);
  if (!value) return;
  if (["bestFor", "pros", "cons", "howToUse", "highlights", "pricing"].includes(key)) {
    tool[key].push(value);
  } else if (tool[key]) {
    tool[key] = cleanText(tool[key] + " " + value);
  } else {
    tool[key] = value;
  }
}

function parseWatchList(paras) {
  const start = paras.findIndex((p) => /^WATCH LIST/i.test(cleanText(p)));
  if (start < 0) return [];
  const out = [];
  let item = null;
  let field = null;
  for (let i = start + 1; i < paras.length; i++) {
    const p = cleanText(paras[i]);
    if (!p || /^Research compiled/i.test(p)) continue;
    const isEntry = /^[A-Z0-9 &+().\/-]{3,}$/.test(p) && !/^(What it is|Current status|Why watch it|Why it matters|What AEO actually involves|Key tools agencies are using|What agencies need to know)/i.test(p);
    if (isEntry) {
      if (item) out.push(item);
      item = { name: prettyToolName(p), fields: {}, body: [], status: "", sourceName: "Research Categories 8-10" };
      field = null;
      continue;
    }
    if (!item) continue;
    if (/^What it is$/i.test(p)) field = "whatItIs";
    else if (/^Current status$/i.test(p)) field = "status";
    else if (/^Why watch it/i.test(p)) field = "whyWatch";
    else if (/^Why it matters/i.test(p)) field = "whyMatters";
    else if (/^What AEO actually involves$/i.test(p)) field = "involves";
    else if (/^Key tools agencies are using$/i.test(p)) field = "tools";
    else if (/^What agencies need to know$/i.test(p)) field = "agencyTakeaway";
    else if (field) {
      if (!item.fields[field]) item.fields[field] = [];
      item.fields[field].push(p);
      if (field === "status" && !item.status) item.status = p;
      item.body.push(p);
    }
  }
  if (item) out.push(item);
  return out;
}

function parseRocketium(paras) {
  const getAfter = (label, stopLabels) => {
    const start = paras.findIndex((p) => cleanText(p).toLowerCase() === label.toLowerCase());
    if (start < 0) return [];
    const out = [];
    for (let i = start + 1; i < paras.length; i++) {
      const p = cleanText(paras[i]);
      if (stopLabels.some((s) => p.toLowerCase() === s.toLowerCase())) break;
      out.push(p);
    }
    return out;
  };
  const lineVal = (prefix) => {
    const line = paras.find((p) => cleanText(p).toLowerCase().startsWith(prefix.toLowerCase() + ":"));
    return line ? cleanText(line.slice(prefix.length + 1)) : "";
  };
  return {
    name: "Rocketium",
    experienced: false,
    sectionCategory: "Workflow & Productivity",
    sourceName: "Rocketium Codex Handoff",
    categoryDetail: lineVal("Recommended Category") || "Workflow & Productivity",
    url: "rocketium.ai",
    tagline: lineVal("Tagline") || "Creative production at campaign scale.",
    whatItIs: getAfter("What it is", ["Pros"]).join(" "),
    bestFor: [lineVal("Best For")].filter(Boolean),
    pros: getAfter("Pros", ["Cons"]),
    cons: getAfter("Cons", ["How to use it"]),
    howToUse: getAfter("How to use it", ["Highlight Features"]),
    highlights: getAfter("Highlight Features", ["Pricing"]),
    pricing: getAfter("Pricing", ["Site-Specific Opinionated Framing"]),
  };
}

function gpt55FromDatabase(md) {
  const m = md.match(/### GPT-5\.5[\s\S]*?(?=\n### |\n---|\n## Category|$)/);
  if (!m) return null;
  const block = m[0];
  const fields = {};
  block.replace(/- \*\*([^*]+):\*\*\s*([^\n]+)/g, (_, key, val) => {
    fields[key.trim().toLowerCase()] = cleanText(val).replace(/^\[TBD[^\]]*\]$/i, "");
    return "";
  });
  return {
    name: "GPT-5.5",
    experienced: false,
    sectionCategory: "LLMs / Thinking Tools",
    sourceName: "AI Tools Database",
    categoryDetail: "LLMs / Thinking Tools",
    url: fields.url || "chatgpt.com",
    tagline: fields.tagline || "OpenAI's newest flagship model, built for sustained multi-step knowledge work",
    whatItIs: fields["what it is"] || "",
    bestFor: [fields["best for"]].filter(Boolean),
    pros: [
      "Designed for sustained multi-step knowledge work, agentic coding, computer use, and larger-context tasks.",
      "Useful when strategy, research synthesis, and source material need to stay active across a long assignment.",
    ],
    cons: [fields.cons].filter(Boolean),
    howToUse: [
      "Reserve it for complex strategy documents, competitive analysis, and research synthesis where quality matters more than speed.",
      "Use the larger context window to keep source material, constraints, and final output expectations in one working session.",
    ],
    highlights: [fields["highlight features"]].filter(Boolean),
    pricing: [fields.pricing].filter(Boolean),
  };
}

function dedupeTools(tools) {
  const seen = new Set();
  const out = [];
  for (const tool of tools) {
    const key = (tool.sectionCategory + "|" + tool.name).toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(tool);
  }
  return out;
}

function buildData() {
  const about = readMd(sourceFiles.about);
  const brief = readMd(sourceFiles.brief);
  const db = readMd(sourceFiles.db);
  const philosophy = docxParagraphs(sourceFiles.philosophy);
  const mood = docxParagraphs(sourceFiles.mood);
  const positioning = docxParagraphs(sourceFiles.positioning);
  const researchA = docxParagraphs(sourceFiles.researchA);
  const researchB = docxParagraphs(sourceFiles.researchB);
  const researchC = docxParagraphs(sourceFiles.researchC);
  const rocketium = docxParagraphs(sourceFiles.rocketium);

  let tools = [
    ...parseResearchTools(researchA, "Research Brief: Categories 1-3"),
    ...parseResearchTools(researchB, "Research Brief: Categories 4-7"),
    ...parseResearchTools(researchC, "Research Brief: Categories 8-10"),
  ];

  if (!tools.some((t) => t.name === "GPT-5.5")) {
    const gpt = gpt55FromDatabase(db);
    if (gpt) tools.splice(Math.max(0, tools.findIndex((t) => t.name === "Gemini")), 0, gpt);
  }
  tools.push(parseRocketium(rocketium));
  tools = dedupeTools(tools);

  const cats = [
    { id: "llms", label: "LLMs / Thinking Tools", accent: colors.indigo, subtitle: "Reasoning, synthesis, writing, and source interpretation." },
    { id: "imagevideo", label: "Image & Video Generation", accent: colors.amber, subtitle: "Visual exploration, motion concepting, and production reference generation." },
    { id: "allinone", label: "All-in-One Generative AI", accent: colors.teal, subtitle: "Multi-modal production suites for moving from idea to proof of concept." },
    { id: "spatial", label: "3D & Spatial Generation", accent: colors.blue, subtitle: "Materials, meshes, environments, and interactive world exploration." },
    { id: "audio", label: "Audio & Voice", accent: colors.rose, subtitle: "Voiceover, music direction, licensing, and sonic experimentation." },
    { id: "copy", label: "Copy & Content", accent: colors.emerald, subtitle: "Brand voice systems, copy workflows, and repeatable content operations." },
    { id: "research", label: "Research & Insights", accent: colors.violet, subtitle: "Source-grounded discovery, owned-source synthesis, and research agents." },
    { id: "workflow", label: "Workflow & Productivity", accent: colors.slate, subtitle: "Coding, automation, desktop agents, and CreativeOps production flow." },
    { id: "design", label: "Presentation & Design", accent: colors.orange, subtitle: "Decks, prototypes, design systems, and presentation-first creative work." },
    { id: "marketing", label: "Marketing & CRM AI", accent: colors.red, subtitle: "CRM intelligence, lifecycle activation, avatar video, and campaign scaling." },
  ];

  const categoryMap = new Map(cats.map((c) => [c.label.toLowerCase(), c]));
  function matchCategory(tool) {
    const raw = cleanText(tool.sectionCategory || tool.categoryDetail).toLowerCase();
    if (raw.includes("llm") || ["claude", "chatgpt", "gpt-5.5", "gemini", "perplexity", "grok"].includes(tool.name.toLowerCase())) return "LLMs / Thinking Tools";
    if (raw.includes("image") || raw.includes("video")) return "Image & Video Generation";
    if (raw.includes("all-in-one")) return "All-in-One Generative AI";
    if (raw.includes("3d") || raw.includes("spatial")) return "3D & Spatial Generation";
    if (raw.includes("audio") || raw.includes("voice")) return "Audio & Voice";
    if (raw.includes("copy") || raw.includes("content")) return "Copy & Content";
    if (raw.includes("research") || raw.includes("insight")) return "Research & Insights";
    if (raw.includes("coding") || raw.includes("developer") || raw.includes("desktop") || raw.includes("workflow") || raw.includes("productivity") || raw.includes("creativeops")) return "Workflow & Productivity";
    if (raw.includes("presentation") || raw.includes("design") || raw.includes("ui/ux")) return "Presentation & Design";
    if (raw.includes("marketing") || raw.includes("crm") || raw.includes("avatar")) return "Marketing & CRM AI";
    return tool.sectionCategory || "Workflow & Productivity";
  }
  cats.forEach((cat) => (cat.tools = []));
  for (const tool of tools) {
    const label = matchCategory(tool);
    const cat = categoryMap.get(label.toLowerCase());
    if (cat) cat.tools.push(tool);
  }

  const watchlist = parseWatchList(researchC);
  return { about, brief, philosophy, mood, positioning, categories: cats, watchlist };
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

let shapeId = 2;
let mediaCounter = 1;
const media = new Map();
const slides = [];

function nextId() {
  return shapeId++;
}

function newSlide(bg = colors.bg) {
  shapeId = 2;
  const slide = { xml: [], rels: [], relCounter: 1, notes: "" };
  addRect(slide, 0, 0, W, H, bg, null, "rect");
  slides.push(slide);
  return slide;
}

function spPr(x, y, w, h, fill, line, prst = "rect") {
  const fillXml = fill ? `<a:solidFill><a:srgbClr val="${hex(fill)}"/></a:solidFill>` : "<a:noFill/>";
  const lineXml = line ? `<a:ln w="${Math.round((line.w || 1) * 12700)}"><a:solidFill><a:srgbClr val="${hex(line.color || colors.line)}"/></a:solidFill></a:ln>` : "<a:ln><a:noFill/></a:ln>";
  return `<p:spPr><a:xfrm><a:off x="${inch(x)}" y="${inch(y)}"/><a:ext cx="${inch(w)}" cy="${inch(h)}"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${fillXml}${lineXml}</p:spPr>`;
}

function addRect(slide, x, y, w, h, fill, line, prst = "rect") {
  const id = nextId();
  slide.xml.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>${spPr(x, y, w, h, fill, line, prst)}<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`);
}

function runXml(text, opts = {}) {
  const sz = Math.round((opts.size || 14) * 100);
  const bold = opts.bold ? ' b="1"' : "";
  return `<a:r><a:rPr lang="en-US" sz="${sz}"${bold}><a:solidFill><a:srgbClr val="${hex(opts.color || colors.ink)}"/></a:solidFill><a:latin typeface="${esc(opts.display ? fontDisplay : font)}"/><a:ea typeface="${esc(font)}"/><a:cs typeface="${esc(font)}"/></a:rPr><a:t>${esc(text)}</a:t></a:r>`;
}

function paraXml(text, opts = {}) {
  return `<a:p><a:pPr algn="${opts.align || "l"}"/>${runXml(text, opts)}<a:endParaRPr lang="en-US" sz="${Math.round((opts.size || 14) * 100)}"/></a:p>`;
}

function addText(slide, x, y, w, h, text, opts = {}) {
  const id = nextId();
  const lines = Array.isArray(text) ? text : String(text || "").split(/\n/);
  const paras = lines.map((line) => paraXml(typeof line === "object" ? line.text : line, typeof line === "object" ? { ...opts, ...line } : opts)).join("");
  const bodyPr = `<a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="${opts.valign || "t"}"><a:normAutofit fontScale="72000" lnSpcReduction="12000"/></a:bodyPr>`;
  slide.xml.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>${spPr(x, y, w, h, null, null, "rect")}<p:txBody>${bodyPr}<a:lstStyle/>${paras}</p:txBody></p:sp>`);
}

function setNotes(slide, text) {
  slide.notes = String(text || "")
    .replace(/\r\n/g, "\n")
    .split("\n")
    .map(cleanText)
    .join("\n")
    .trim();
}

function exists(file) {
  try { return fs.existsSync(file); } catch { return false; }
}

function imgDim(file) {
  const b = fs.readFileSync(file);
  if (b.length > 24 && b[0] === 0x89 && b.slice(1, 4).toString() === "PNG") return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
  if (b[0] === 0xff && b[1] === 0xd8) {
    let i = 2;
    while (i < b.length - 9) {
      if (b[i] !== 0xff) break;
      const marker = b[i + 1];
      const len = b.readUInt16BE(i + 2);
      if (marker >= 0xc0 && marker <= 0xc3) return { h: b.readUInt16BE(i + 5), w: b.readUInt16BE(i + 7) };
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

function addRelImage(slide, file) {
  const key = path.resolve(file);
  let item = media.get(key);
  if (!item) {
    const extRaw = path.extname(file).toLowerCase().replace(".", "") || "png";
    const ext = extRaw === "jpeg" ? "jpg" : extRaw;
    const mediaName = `image${mediaCounter++}.${ext}`;
    item = { target: `../media/${mediaName}`, name: `ppt/media/${mediaName}`, data: fs.readFileSync(file), ext: ext === "jpg" ? "jpeg" : ext };
    media.set(key, item);
  }
  const rid = `rId${slide.relCounter++}`;
  slide.rels.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${item.target}"/>`);
  return rid;
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

function addFooter(slide, label = "The AI Stack | Source-Based Presentation") {
  addText(slide, 0.55, 7.08, 8.0, 0.18, label, { size: 8.5, color: colors.dim });
  addText(slide, 11.1, 7.08, 1.6, 0.18, "Kyle Zemba", { size: 8.5, color: colors.dim, align: "r" });
}

function addTitle(slide, title, kicker, accent = colors.indigo) {
  addText(slide, 0.62, 0.44, 7.8, 0.22, cleanText(kicker || "The AI Stack").toUpperCase(), { size: 8.5, color: accent, bold: true });
  addText(slide, 0.62, 0.76, 10.2, 0.62, title, { size: 31, color: colors.ink, bold: true, display: true });
  addRect(slide, 0.62, 1.48, 1.05, 0.04, accent, null, "rect");
}

function panel(slide, x, y, w, h, title, body, accent, opts = {}) {
  addRect(slide, x, y, w, h, opts.fill || colors.panel, { color: opts.line || colors.line, w: 0.65 }, "roundRect");
  if (title) addText(slide, x + 0.18, y + 0.16, w - 0.36, 0.18, title.toUpperCase(), { size: 7.7, color: accent, bold: true });
  const text = Array.isArray(body) ? body.filter(Boolean).join("\n") : cleanText(body);
  addText(slide, x + 0.18, y + (title ? 0.45 : 0.18), w - 0.36, h - (title ? 0.55 : 0.25), text, { size: opts.size || 9.6, color: opts.color || colors.muted });
}

function toolNotes(cat, tool) {
  const lines = [
    `Source category: ${cat.label}`,
    `Source document: ${tool.sourceName || ""}`,
    `Tool: ${tool.name}`,
    `URL: ${tool.url || ""}`,
    "",
    `Tagline: ${tool.tagline || ""}`,
    "",
    "What it is:",
    tool.whatItIs || "",
    "",
    "Best For:",
    ...normArray(tool.bestFor).map((x) => "- " + x),
    "",
    "Pros:",
    ...normArray(tool.pros).map((x) => "- " + x),
    "",
    "Cons:",
    ...normArray(tool.cons).map((x) => "- " + x),
    "",
    "How to Use It:",
    ...normArray(tool.howToUse).map((x) => "- " + x),
    "",
    "Highlight Features:",
    ...normArray(tool.highlights).map((x) => "- " + x),
    "",
    "Pricing:",
    ...normArray(tool.pricing).map((x) => "- " + x),
  ];
  return lines.join("\n");
}

function addCover() {
  const s = newSlide();
  addRect(s, 0, 0, 0.18, H, colors.indigo, null, "rect");
  addText(s, 0.75, 0.82, 8.2, 0.25, "SOURCE-BASED POWERPOINT DECK", { size: 10, color: colors.indigo, bold: true });
  addText(s, 0.75, 1.27, 8.6, 0.9, "The AI Stack", { size: 54, color: colors.ink, bold: true, display: true });
  addText(s, 0.79, 2.24, 7.6, 0.42, "Tools for Advertising & Marketing", { size: 23, color: colors.muted, bold: true });
  addText(s, 0.8, 3.05, 6.6, 0.8, "A presentation-first version built from the project source files: strategy docs, research briefs, philosophy framework, mood board, and tool-positioning notes.", { size: 15.5, color: colors.muted });
  addRect(s, 8.45, 0.82, 3.7, 4.9, colors.panel, { color: colors.line, w: 0.7 }, "roundRect");
  addText(s, 8.8, 1.22, 2.9, 0.24, "DECK PRINCIPLE", { size: 9, color: colors.teal, bold: true });
  addText(s, 8.8, 1.78, 2.8, 0.68, "PowerPoint is for sequence, emphasis, and discussion.", { size: 21, color: colors.ink, bold: true, display: true });
  addText(s, 8.8, 3.0, 2.85, 1.22, "This version does not copy the website layout. It turns the same source thinking into a presenter-friendly story with one slide per tool and deeper detail in speaker notes.", { size: 11.5, color: colors.muted });
  addFooter(s, "Built from project folders/files only");
  setNotes(s, "This deck was generated from non-HTML project sources only. HTML files were intentionally excluded as source material.");
}

function addContextSlides(data) {
  let s = newSlide();
  addTitle(s, "Why This Project Exists", "Project Context", colors.teal);
  panel(s, 0.68, 1.85, 3.75, 3.1, "The audience", "Agency creative directors, strategists, production leaders, and marketing decision-makers who have seen surface-level AI decks and need a sharper point of view.", colors.teal, { size: 12 });
  panel(s, 4.8, 1.85, 3.75, 3.1, "The signal", "This is meant to demonstrate AI fluency, not tool awareness. The value is in knowing which tool belongs where, how to use it, and where it breaks.", colors.teal, { size: 12 });
  panel(s, 8.92, 1.85, 3.75, 3.1, "The proof", "Kyle's positioning is conceptual-before-executional: define the audience, purpose, and creative standard before opening the tool.", colors.teal, { size: 12 });
  addFooter(s);
  setNotes(s, "Source: About Project.md and ChatGPT Project Brief - Updated.md. The project originated as a portfolio asset to show strategic AI fluency for advertising and marketing work.");

  s = newSlide();
  addTitle(s, "AI Is Only As Good As The Thinking Before It", "AI Philosophy", colors.indigo);
  addText(s, 0.72, 1.84, 5.9, 1.0, "Most people start with the tool. That is the mistake.", { size: 24, color: colors.ink, bold: true, display: true });
  addText(s, 0.74, 3.02, 5.75, 1.02, "The difference between generic output and high-quality work is not just the model. It is the clarity of intent, context, references, audience, and standard you give it.", { size: 15, color: colors.muted });
  const items = [
    ["Context positioning", "Define project, references, tone, and audience before output."],
    ["Clarification first", "Ask the model to summarize understanding and identify gaps."],
    ["Multimodal thinking", "Use text, image, audio, and video tools together."],
    ["Memory and iteration", "Treat the interaction as an evolving working instance."],
    ["Strategic intent", "Define purpose, placement, and audience before production."],
  ];
  items.forEach((item, i) => {
    const y = 1.62 + i * 0.82;
    addText(s, 7.08, y, 0.48, 0.18, String(i + 1).padStart(2, "0"), { size: 8.5, color: colors.indigo, bold: true });
    addText(s, 7.62, y - 0.02, 3.9, 0.2, item[0], { size: 12.5, color: colors.ink, bold: true });
    addText(s, 7.62, y + 0.26, 4.15, 0.2, item[1], { size: 9.5, color: colors.muted });
  });
  addFooter(s);
  setNotes(s, data.philosophy.join("\n"));

  s = newSlide();
  addTitle(s, "How To Read The Stack", "Presentation Guide", colors.amber);
  addText(s, 0.72, 1.78, 5.8, 0.7, "Each tool slide answers a different question than the website does.", { size: 22, color: colors.ink, bold: true, display: true });
  const guide = [
    ["What it enables", "The practical role of the tool in agency work."],
    ["Where it fits", "The team, workflow, or project type that gets the most value."],
    ["How to use it", "Actionable guidance, not a generic feature tour."],
    ["Watchouts", "Limitations, pricing friction, policy issues, and quality gaps."],
  ];
  guide.forEach((g, i) => {
    const x = 0.78 + (i % 2) * 6.1;
    const y = 3.05 + Math.floor(i / 2) * 1.35;
    panel(s, x, y, 5.25, 1.0, g[0], g[1], colors.amber, { size: 11.5 });
  });
  addFooter(s);
  setNotes(s, "Source: Tool Card Template.md and Tone and Voice.md. Tool entries use the schema: name, company, tagline, what it is, best for, pros, cons, how to use it, highlight features, pricing, and URL.");

  s = newSlide();
  addTitle(s, "Design Direction For The Deck", "Visual Mood", colors.blue);
  panel(s, 0.72, 1.78, 3.5, 3.55, "Overall aesthetic", "Premium, modern, technical, restrained. The source mood board points toward a high-end agency deck, dark-mode dashboard, and product design system.", colors.blue, { size: 12 });
  panel(s, 4.58, 1.78, 3.5, 3.55, "Principles", "Clarity over decoration. Structured grid. Strong hierarchy. Flat category color accents. Consistent rhythm.", colors.blue, { size: 12 });
  panel(s, 8.44, 1.78, 3.5, 3.55, "Avoid", "Startup landing page energy, generic AI directory formatting, gradients, loud color, clutter, and over-designed UI.", colors.blue, { size: 12 });
  addFooter(s);
  setNotes(s, data.mood.join("\n"));
}

function addEcosystemSlide(data) {
  const s = newSlide();
  addTitle(s, "The AI Stack, Organized By Agency Workflow", "Ecosystem Map", colors.violet);
  const positions = [
    [0.62, 1.65, 3.75, 1.1], [4.58, 1.65, 8.05, 1.1],
    [0.62, 2.98, 3.75, 1.0], [4.58, 2.98, 3.75, 1.0], [8.54, 2.98, 4.09, 1.0],
    [0.62, 4.2, 3.75, 1.0], [4.58, 4.2, 3.75, 1.0], [8.54, 4.2, 4.09, 1.0],
    [0.62, 5.42, 3.75, 1.0], [4.58, 5.42, 3.75, 1.0],
  ];
  data.categories.forEach((cat, i) => {
    const [x, y, w, h] = positions[i];
    addRect(s, x, y, w, h, colors.panel, { color: cat.accent, w: 0.8 }, "roundRect");
    addText(s, x + 0.15, y + 0.13, w - 0.3, 0.16, cat.label.toUpperCase(), { size: 7.2, color: cat.accent, bold: true });
    addText(s, x + 0.15, y + 0.42, w - 0.3, 0.38, cat.tools.map((t) => t.name + (t.experienced ? " *" : "")).join("  |  "), { size: i === 1 ? 7.7 : 7.6, color: colors.ink });
  });
  addText(s, 8.78, 6.52, 3.8, 0.16, "* = direct hands-on experience noted in source files", { size: 8.5, color: colors.dim, align: "r" });
  addFooter(s);
  setNotes(s, data.categories.map((c) => `${c.label}: ${c.tools.map((t) => t.name).join(", ")}`).join("\n"));
}

function addWorkflowSlides() {
  const stages = [
    ["Research", "Claude", "Tool landscape mapping: model availability, pricing, and best use cases for the project."],
    ["Strategize", "ChatGPT", "Creative concept development, story route, and storyboard thinking."],
    ["Plan", "ChatGPT Pro + Claude Cowork", "Organized shot list with image/video prompts, timestamps, and model assignments per shot."],
    ["Deliver", "Higgsfield hub", "Higgsfield as the central hub, with Seedance 2.0, Nano Banana, Kling, and ElevenLabs feeding into generation and audio."],
  ];
  let s = newSlide();
  addTitle(s, "Example Workflow: Western Nerf Fight", "Proof Point", colors.teal);
  addText(s, 0.72, 1.7, 6.2, 0.62, "The source files describe a real AI-generated short film proof of concept built in under a day.", { size: 18, color: colors.ink, bold: true });
  stages.forEach((stage, i) => {
    const x = 0.78 + i * 3.05;
    panel(s, x, 3.0, 2.55, 1.7, stage[0], `${stage[1]}\n${stage[2]}`, colors.teal, { size: 8.8 });
  });
  addText(s, 0.78, 5.35, 8.2, 0.28, "Source video:", { size: 11, color: colors.teal, bold: true });
  addText(s, 1.8, 5.35, 8.8, 0.28, "Context/Video/Western Nerf Fight.mp4", { size: 11, color: colors.muted });
  addFooter(s);
  setNotes(s, stages.map((x) => `${x[0]} - ${x[1]}: ${x[2]}`).join("\n"));

  s = newSlide();
  addTitle(s, "Workflow Evidence From The Project Folder", "Source Visuals", colors.teal);
  const imgs = [
    "Context/Example Workflow/Step 1 - AI Context Workflow.png",
    "Context/Example Workflow/Step 2 - Higgsfield One Canvas.png",
    "Context/Example Workflow/Step 4 - Build Nodes.png",
    "Context/Example Workflow/Results Video Poster 2s.jpg",
  ];
  imgs.forEach((img, i) => {
    const x = 0.72 + (i % 2) * 6.08;
    const y = 1.65 + Math.floor(i / 2) * 2.55;
    addRect(s, x, y, 5.45, 2.18, colors.panel, { color: colors.line, w: 0.6 }, "roundRect");
    addContainedImage(s, img, x + 0.12, y + 0.12, 5.21, 1.94);
  });
  addFooter(s);
  setNotes(s, "Source images from Context/Example Workflow. These visuals support the workflow proof point without using the website HTML.");
}

function addSectionDivider(cat, index) {
  const s = newSlide();
  addRect(s, 0, 0, 0.16, H, cat.accent, null, "rect");
  addText(s, 0.72, 0.78, 1.4, 0.2, `SECTION ${String(index + 1).padStart(2, "0")}`, { size: 9, color: cat.accent, bold: true });
  addText(s, 0.72, 1.22, 7.9, 0.8, cat.label, { size: 42, color: colors.ink, bold: true, display: true });
  addText(s, 0.74, 2.26, 6.7, 0.48, cat.subtitle, { size: 15, color: colors.muted });
  addRect(s, 8.2, 0.9, 4.2, 5.18, colors.panel, { color: cat.accent, w: 0.9 }, "roundRect");
  addText(s, 8.52, 1.22, 3.2, 0.22, `${cat.tools.length} source cards`, { size: 12, color: cat.accent, bold: true });
  cat.tools.forEach((tool, i) => {
    if (i > 12) return;
    addText(s, 8.52, 1.68 + i * 0.32, 3.45, 0.18, `${String(i + 1).padStart(2, "0")}  ${tool.name}`, { size: 10.5, color: colors.ink });
  });
  addFooter(s, `The AI Stack | ${cat.label}`);
  setNotes(s, `${cat.label}\n${cat.subtitle}\n\nTools:\n${cat.tools.map((t) => "- " + t.name).join("\n")}`);
}

function addToolSlide(cat, tool, i) {
  const s = newSlide();
  addRect(s, 0, 0, 0.12, H, cat.accent, null, "rect");
  addText(s, 0.62, 0.42, 5.8, 0.18, cat.label.toUpperCase(), { size: 8.2, color: cat.accent, bold: true });
  addText(s, 0.62, 0.72, 6.85, 0.55, tool.name + (tool.experienced ? " *" : ""), { size: 30, color: colors.ink, bold: true, display: true });
  addText(s, 0.64, 1.26, 6.3, 0.22, trunc(tool.categoryDetail || tool.sourceName || "", 105), { size: 9.5, color: colors.dim });
  addRect(s, 10.78, 0.52, 1.62, 0.36, cat.accent, null, "roundRect");
  addText(s, 10.92, 0.63, 1.35, 0.1, `CARD ${String(i + 1).padStart(2, "0")}`, { size: 7.5, color: colors.bg, bold: true, align: "c" });
  addText(s, 0.66, 1.72, 6.9, 0.44, trunc(tool.tagline || "Source card", 150), { size: 15.2, color: cat.accent, bold: true });

  panel(s, 0.66, 2.45, 5.8, 1.22, "What it enables", trunc(tool.whatItIs, 330), cat.accent, { size: 9.8 });
  panel(s, 6.78, 1.68, 2.78, 1.38, "Best for", normArray(tool.bestFor).slice(0, 2).map((x) => "- " + trunc(x, 110)), cat.accent, { size: 8.6 });
  panel(s, 9.82, 1.68, 2.78, 1.38, "Use it when", normArray(tool.howToUse).slice(0, 2).map((x) => "- " + trunc(x, 110)), cat.accent, { size: 8.6 });
  panel(s, 0.66, 4.0, 3.75, 1.45, "Strengths", normArray(tool.pros).slice(0, 3).map((x) => "- " + trunc(x, 90)), cat.accent, { size: 8.3 });
  panel(s, 4.68, 4.0, 3.75, 1.45, "Watchouts", normArray(tool.cons).slice(0, 3).map((x) => "- " + trunc(x, 90)), cat.accent, { size: 8.3 });
  panel(s, 8.7, 4.0, 3.9, 1.45, "Highlights", normArray(tool.highlights).slice(0, 3).map((x) => "- " + trunc(x, 88)), cat.accent, { size: 8.3 });
  addRect(s, 0.66, 5.88, 11.94, 0.56, colors.panel2, { color: colors.line, w: 0.55 }, "roundRect");
  addText(s, 0.86, 6.08, 7.45, 0.12, `Pricing: ${trunc(normArray(tool.pricing).join(" | ") || "Verify current pricing before recommendation.", 145)}`, { size: 8.2, color: colors.muted });
  addText(s, 8.56, 6.08, 3.8, 0.12, trunc(tool.url || "", 74), { size: 8.2, color: cat.accent, align: "r" });
  addFooter(s, `The AI Stack | ${cat.label}`);
  setNotes(s, toolNotes(cat, tool));
}

function addWatchlist(data) {
  const s = newSlide();
  addRect(s, 0, 0, 0.16, H, colors.gray, null, "rect");
  addText(s, 0.72, 0.78, 1.4, 0.2, "FINAL SECTION", { size: 9, color: colors.gray, bold: true });
  addText(s, 0.72, 1.22, 7.9, 0.8, "Watch List / Emerging", { size: 40, color: colors.ink, bold: true, display: true });
  addText(s, 0.74, 2.22, 6.95, 0.5, "Forward-looking tools, categories, and signals from the source research. These are not all ready for mainstream agency workflow adoption.", { size: 15, color: colors.muted });
  data.watchlist.forEach((item, i) => {
    const x = 0.78 + (i % 3) * 4.08;
    const y = 3.32 + Math.floor(i / 3) * 1.35;
    panel(s, x, y, 3.55, 0.98, item.name, trunc((item.fields.whatItIs || item.body || []).join(" "), 135), colors.gray, { size: 8.4 });
  });
  addFooter(s, "The AI Stack | Watch List");
  setNotes(s, data.watchlist.map((w) => `${w.name}\n${Object.entries(w.fields).map(([k, v]) => `${k}: ${v.join(" ")}`).join("\n")}`).join("\n\n"));

  data.watchlist.forEach((item) => {
    const slide = newSlide();
    addTitle(slide, item.name, "Watch List / Emerging", colors.gray);
    panel(slide, 0.72, 1.72, 5.7, 1.55, "What it is", trunc((item.fields.whatItIs || []).join(" "), 360), colors.gray, { size: 10.2 });
    panel(slide, 6.82, 1.72, 5.15, 1.15, "Current status", trunc((item.fields.status || []).join(" "), 260), colors.gray, { size: 9.5 });
    panel(slide, 0.72, 3.72, 11.25, 1.75, "Why it matters", trunc((item.fields.whyWatch || item.fields.whyMatters || item.fields.agencyTakeaway || []).join(" "), 560), colors.gray, { size: 10.2 });
    addFooter(slide, "The AI Stack | Watch List");
    setNotes(slide, `${item.name}\n\n${Object.entries(item.fields).map(([k, v]) => `${k}:\n${v.map((x) => "- " + x).join("\n")}`).join("\n\n")}`);
  });
}

function addClose(data) {
  const s = newSlide();
  addText(s, 0.78, 1.18, 9.2, 0.84, "AI fluency is workflow fluency.", { size: 45, color: colors.ink, bold: true, display: true });
  addText(s, 0.8, 2.36, 7.2, 0.75, "The strongest teams do not chase every new model. They understand the job, choose the right tool, structure the workflow, and keep human judgment in the loop.", { size: 18, color: colors.muted });
  addRect(s, 0.8, 4.2, 3.0, 0.05, colors.teal, null, "rect");
  addText(s, 0.8, 5.05, 5.0, 0.34, "The AI Stack | Tools for Advertising & Marketing", { size: 16, color: colors.teal, bold: true });
  addText(s, 0.8, 5.48, 4.6, 0.22, "Source-based PowerPoint deck generated from project folders/files", { size: 10.5, color: colors.dim });
  addFooter(s);
  setNotes(s, "Close the deck by returning to the strategic frame from the source files: the project demonstrates fluency through workflow judgment, not tool collection.");
}

function slideXml(slide) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${slide.xml.join("")}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

function notesRunXml(text, opts = {}) {
  const sz = Math.round((opts.size || 9.4) * 100);
  return `<a:r><a:rPr lang="en-US" sz="${sz}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="${esc(font)}"/><a:ea typeface="${esc(font)}"/><a:cs typeface="${esc(font)}"/></a:rPr><a:t>${esc(text)}</a:t></a:r>`;
}

function notesParaXml(text) {
  return `<a:p><a:pPr algn="l"/>${notesRunXml(text)}<a:endParaRPr lang="en-US" sz="940"/></a:p>`;
}

function notesSlideXml(notes) {
  const paras = String(notes || "").split(/\n/).map(notesParaXml).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notes xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${inch(0.45)}" y="${inch(0.42)}"/><a:ext cx="${inch(6.8)}" cy="${inch(9.2)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:normAutofit fontScale="80000" lnSpcReduction="12000"/></a:bodyPr><a:lstStyle/>${paras}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}

function relsXml(rels) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`;
}

function presentationXml(notesCount) {
  const sldIds = slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join("");
  const notesMaster = notesCount ? `<p:notesMasterIdLst><p:notesMasterId r:id="rId${slides.length + 2}"/></p:notesMasterIdLst>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>${notesMaster}<p:sldIdLst>${sldIds}</p:sldIdLst><p:sldSz cx="${SLIDE_CX}" cy="${SLIDE_CY}" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:defaultTextStyle></p:presentation>`;
}

function contentTypes(notesItems) {
  const overrides = [
    '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
    '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>',
    '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
    '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
  ];
  slides.forEach((_, i) => overrides.push(`<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`));
  if (notesItems.length) {
    overrides.push('<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>');
    notesItems.forEach((_, i) => overrides.push(`<Override PartName="/ppt/notesSlides/notesSlide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`));
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="jpeg" ContentType="image/jpeg"/>${overrides.join("")}</Types>`;
}

function writeDeck() {
  const notesItems = slides.map((slide, slideIndex) => ({ slide, slideIndex })).filter((item) => item.slide.notes).map((item, notesIndex) => ({ ...item, notesIndex }));
  notesItems.forEach((item) => {
    const rid = `rId${item.slide.relCounter++}`;
    item.slide.rels.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide${item.notesIndex + 1}.xml"/>`);
  });
  const slideLayoutXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
  const slideMasterXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${colors.bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`;
  const notesMasterXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notesMaster xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr><a:defRPr sz="1000" lang="en-US"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="${esc(font)}"/></a:defRPr></a:lvl1pPr></p:notesStyle></p:notesMaster>`;
  const themeXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="${NS.a}" name="AI Stack Source Deck"><a:themeElements><a:clrScheme name="AI Stack"><a:dk1><a:srgbClr val="${colors.bg}"/></a:dk1><a:lt1><a:srgbClr val="${colors.ink}"/></a:lt1><a:dk2><a:srgbClr val="${colors.panel}"/></a:dk2><a:lt2><a:srgbClr val="${colors.muted}"/></a:lt2><a:accent1><a:srgbClr val="${colors.indigo}"/></a:accent1><a:accent2><a:srgbClr val="${colors.amber}"/></a:accent2><a:accent3><a:srgbClr val="${colors.teal}"/></a:accent3><a:accent4><a:srgbClr val="${colors.blue}"/></a:accent4><a:accent5><a:srgbClr val="${colors.rose}"/></a:accent5><a:accent6><a:srgbClr val="${colors.violet}"/></a:accent6><a:hlink><a:srgbClr val="${colors.teal}"/></a:hlink><a:folHlink><a:srgbClr val="${colors.violet}"/></a:folHlink></a:clrScheme><a:fontScheme name="AI Stack"><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="AI Stack"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
  const entries = [];
  entries.push({ name: "[Content_Types].xml", data: contentTypes(notesItems) });
  entries.push({ name: "_rels/.rels", data: relsXml([
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>',
  ]) });
  const now = new Date().toISOString();
  entries.push({ name: "docProps/core.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>The AI Stack Source-Based Presentation</dc:title><dc:creator>Codex</dc:creator><cp:lastModifiedBy>Codex</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>` });
  entries.push({ name: "docProps/app.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Codex</Application><PresentationFormat>On-screen Show (16:9)</PresentationFormat><Slides>${slides.length}</Slides><Notes>${notesItems.length}</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop></Properties>` });
  entries.push({ name: "ppt/presentation.xml", data: presentationXml(notesItems.length) });
  const presRels = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'];
  slides.forEach((_, i) => presRels.push(`<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`));
  if (notesItems.length) presRels.push(`<Relationship Id="rId${slides.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>`);
  entries.push({ name: "ppt/_rels/presentation.xml.rels", data: relsXml(presRels) });
  entries.push({ name: "ppt/slideMasters/slideMaster1.xml", data: slideMasterXml });
  entries.push({ name: "ppt/slideMasters/_rels/slideMaster1.xml.rels", data: relsXml([
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>',
  ]) });
  entries.push({ name: "ppt/slideLayouts/slideLayout1.xml", data: slideLayoutXml });
  entries.push({ name: "ppt/slideLayouts/_rels/slideLayout1.xml.rels", data: relsXml(['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>']) });
  entries.push({ name: "ppt/theme/theme1.xml", data: themeXml });
  if (notesItems.length) {
    entries.push({ name: "ppt/notesMasters/notesMaster1.xml", data: notesMasterXml });
    entries.push({ name: "ppt/notesMasters/_rels/notesMaster1.xml.rels", data: relsXml(['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>']) });
  }
  slides.forEach((s, i) => {
    entries.push({ name: `ppt/slides/slide${i + 1}.xml`, data: slideXml(s) });
    if (s.rels.length) entries.push({ name: `ppt/slides/_rels/slide${i + 1}.xml.rels`, data: relsXml(s.rels) });
  });
  notesItems.forEach((item) => {
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

function build() {
  const data = buildData();
  addCover();
  addContextSlides(data);
  addEcosystemSlide(data);
  addWorkflowSlides();
  data.categories.forEach((cat, i) => {
    addSectionDivider(cat, i);
    cat.tools.forEach((tool, idx) => addToolSlide(cat, tool, idx));
  });
  addWatchlist(data);
  addClose(data);
  writeDeck();
  console.log(JSON.stringify({
    output: OUT,
    slides: slides.length,
    notes: slides.filter((s) => s.notes).length,
    tools: data.categories.reduce((n, c) => n + c.tools.length, 0),
    watchlist: data.watchlist.length,
    categories: data.categories.map((c) => ({ label: c.label, tools: c.tools.length })),
    media: media.size,
    bytes: fs.statSync(OUT).size,
  }, null, 2));
}

build();
