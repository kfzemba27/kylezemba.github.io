const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

const INPUT = path.join("Output", "AI_Tools_Presentation.pptx");
const OUTPUT = path.join("Output", "AI_Tools_Presentation_with_Logos.pptx");
const EMU = 914400;
const NS = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
};

function inch(v) {
  return Math.round(v * EMU);
}

function esc(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function hex(v) {
  return String(v || "").replace("#", "").toUpperCase();
}

function readZipEntries(file) {
  const buf = fs.readFileSync(file);
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65558); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error("Could not find zip directory");
  const total = buf.readUInt16LE(eocd + 10);
  let off = buf.readUInt32LE(eocd + 16);
  const entries = [];
  for (let i = 0; i < total; i++) {
    if (buf.readUInt32LE(off) !== 0x02014b50) throw new Error("Bad zip central directory");
    const method = buf.readUInt16LE(off + 10);
    const compSize = buf.readUInt32LE(off + 20);
    const nameLen = buf.readUInt16LE(off + 28);
    const extraLen = buf.readUInt16LE(off + 30);
    const commentLen = buf.readUInt16LE(off + 32);
    const localOff = buf.readUInt32LE(off + 42);
    const name = buf.slice(off + 46, off + 46 + nameLen).toString("utf8");
    const localNameLen = buf.readUInt16LE(localOff + 26);
    const localExtraLen = buf.readUInt16LE(localOff + 28);
    const dataStart = localOff + 30 + localNameLen + localExtraLen;
    const comp = buf.slice(dataStart, dataStart + compSize);
    let data;
    if (name.endsWith("/")) data = Buffer.alloc(0);
    else if (method === 0) data = comp;
    else if (method === 8) data = zlib.inflateRawSync(comp);
    else throw new Error("Unsupported compression method " + method + " for " + name);
    entries.push({ name, data });
    off += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
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

const slides = [
  [8, "Claude", "CL", "C15F3C"],
  [9, "ChatGPT", "CG", "10A37F"],
  [10, "Gemini", "G", "4285F4"],
  [11, "Perplexity", "PX", "20B8CD"],
  [12, "Grok", "GX", "111827"],
  [13, "GPT-5.5", "G5", "10A37F"],
  [15, "Midjourney", "MJ", "111827"],
  [16, "Flux", "FL", "111827"],
  [17, "DALL-E", "DE", "10A37F"],
  [18, "GPT Image 2", "GI", "10A37F"],
  [19, "Runway", "RW", "111827"],
  [20, "Sora", "SO", "10A37F"],
  [21, "Kling", "KL", "6D28D9"],
  [22, "Seedance", "SD", "0F172A"],
  [23, "Luma", "LU", "111827"],
  [24, "Pika", "PK", "F59E0B"],
  [25, "Nano Banana", "NB", "FACC15"],
  [26, "ComfyUI", "CU", "F97316"],
  [28, "Higgsfield", "HF", "111827"],
  [29, "Freepik", "FP", "1273EB"],
  [30, "Artlist", "AL", "111827"],
  [32, "Fal + PATINA", "FA", "111827"],
  [33, "Tripo 3D", "T3", "2563EB"],
  [34, "World Labs", "WL", "111827"],
  [35, "Google Genie", "GG", "4285F4"],
  [37, "ElevenLabs", "11", "111827"],
  [38, "Suno", "SU", "7C3AED"],
  [39, "Udio", "UD", "EC4899"],
  [41, "Jasper", "JA", "7C3AED"],
  [42, "Copy.ai", "CA", "111827"],
  [44, "NotebookLM", "NL", "4285F4"],
  [45, "Perplexity", "PX", "20B8CD"],
  [46, "ChatGPT Deep Research", "DR", "10A37F"],
  [47, "Gemini Deep Research", "DR", "4285F4"],
  [49, "GitHub Copilot", "GH", "24292F"],
  [50, "Cursor", "CU", "111827"],
  [51, "Claude Code", "CC", "C15F3C"],
  [52, "Claude Cowork", "CW", "C15F3C"],
  [53, "Replit", "RP", "F26207"],
  [55, "Canva AI", "CV", "00C4CC"],
  [56, "Gamma", "GA", "111827"],
  [57, "Figma", "FG", "A259FF"],
  [59, "Salesforce Einstein", "SF", "00A1E0"],
  [60, "HubSpot Breeze", "HB", "FF5C35"],
  [61, "HeyGen", "HG", "111827"],
  [63, "Mythos", "MY", "C15F3C"],
  [64, "LTX Studio", "LX", "111827"],
  [65, "AI Answer Engine Optimization", "AE", "6B7280"],
  [66, "OpenClaw + NVIDIA NemoClaw", "OC", "76B900"],
];

function maxShapeId(xml) {
  let max = 1;
  xml.replace(/<p:cNvPr\b[^>]*\bid="(\d+)"/g, (_, id) => {
    max = Math.max(max, Number(id));
    return "";
  });
  return max;
}

function badgeXml(id1, id2, label, fill, textColor = "FFFFFF") {
  const x = inch(0.45);
  const y = inch(0.27);
  const w = inch(0.46);
  const h = inch(0.46);
  const line = fill === "111827" || fill === "24292F" ? "94A3B8" : fill;
  return `
<p:sp><p:nvSpPr><p:cNvPr id="${id1}" name="Tool Logo Badge"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${hex(fill)}"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="${hex(line)}"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="${id2}" name="Tool Logo Initials"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y + inch(0.02)}"/><a:ext cx="${w}" cy="${h - inch(0.04)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="950" b="1"><a:solidFill><a:srgbClr val="${hex(textColor)}"/></a:solidFill><a:latin typeface="Aptos Display"/><a:ea typeface="Aptos Display"/><a:cs typeface="Aptos Display"/></a:rPr><a:t>${esc(label)}</a:t></a:r><a:endParaRPr lang="en-US" sz="950"/></a:p></p:txBody></p:sp>`;
}

function shiftTitle(xml, title) {
  const titleEsc = esc(title);
  const re = new RegExp(`(<p:sp>(?:(?!<\\/p:sp>)[\\s\\S])*?<a:t>${titleEsc}<\\/a:t>(?:(?!<\\/p:sp>)[\\s\\S])*?<\\/p:sp>)`);
  return xml.replace(re, (sp) => {
    return sp
      .replace(/<a:off x="(\d+)" y="(\d+)"\/>/, `<a:off x="${inch(1.02)}" y="$2"/>`)
      .replace(/<a:ext cx="(\d+)" cy="(\d+)"\/>/, `<a:ext cx="${inch(6.56)}" cy="$2"/>`);
  });
}

function addBadge(xml, title, label, fill) {
  if (xml.includes("Tool Logo Badge")) return xml;
  const id = maxShapeId(xml) + 1;
  const textColor = fill === "FACC15" || fill === "00C4CC" || fill === "76B900" ? "111827" : "FFFFFF";
  let out = shiftTitle(xml, title);
  return out.replace("</p:spTree>", `${badgeXml(id, id + 1, label, fill, textColor)}</p:spTree>`);
}

const entries = readZipEntries(INPUT);
let edited = 0;
for (const item of slides) {
  const [num, title, label, fill] = item;
  const name = `ppt/slides/slide${num}.xml`;
  const entry = entries.find((e) => e.name === name);
  if (!entry) continue;
  const before = entry.data.toString("utf8");
  const after = addBadge(before, title, label, fill);
  if (after !== before) {
    entry.data = Buffer.from(after, "utf8");
    edited++;
  }
}

fs.writeFileSync(OUTPUT, makeZip(entries));
console.log(JSON.stringify({ input: INPUT, output: OUTPUT, edited, bytes: fs.statSync(OUTPUT).size }, null, 2));
