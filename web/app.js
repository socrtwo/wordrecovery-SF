'use strict';

// ============================================================================
// ImmortalInflate — never-throws DEFLATE decoder.
// Ported verbatim from socrtwo/Universal-File-Repair-Tool.
// ============================================================================
const ImmortalInflate = (function () {
  class BitStream {
    constructor(u8) { this.buf = u8; this.pos = 0; this.bit = 0; this.len = u8.length; }
    read(n) {
      let v = 0;
      for (let i = 0; i < n; i++) {
        if (this.pos >= this.len) return -1;
        v |= ((this.buf[this.pos] >>> this.bit) & 1) << i;
        this.bit++;
        if (this.bit === 8) { this.bit = 0; this.pos++; }
      }
      return v;
    }
    align() { if (this.bit !== 0) { this.bit = 0; this.pos++; } }
  }
  const FIXED_LIT = new Uint8Array(288);
  for (let i = 0; i < 144; i++) FIXED_LIT[i] = 8;
  for (let i = 144; i < 256; i++) FIXED_LIT[i] = 9;
  for (let i = 256; i < 280; i++) FIXED_LIT[i] = 7;
  for (let i = 280; i < 288; i++) FIXED_LIT[i] = 8;
  const FIXED_DIST = new Uint8Array(32); for (let i = 0; i < 32; i++) FIXED_DIST[i] = 5;
  const CLEN_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
  const LEN_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
  const LEN_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
  const DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
  const DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];

  function buildTree(lengths) {
    const counts = new Int32Array(16), nextCode = new Int32Array(16);
    let maxLen = 0;
    for (let i = 0; i < lengths.length; i++) { counts[lengths[i]]++; if (lengths[i] > maxLen) maxLen = lengths[i]; }
    if (maxLen === 0) return null;
    let code = 0; counts[0] = 0;
    for (let i = 1; i <= 15; i++) { code = (code + counts[i - 1]) << 1; nextCode[i] = code; }
    const map = {};
    for (let i = 0; i < lengths.length; i++) {
      const len = lengths[i];
      if (len !== 0) { map[(len << 16) | nextCode[len]] = i; nextCode[len]++; }
    }
    return { map, maxLen };
  }
  function decodeSym(s, t) {
    let c = 0;
    for (let l = 1; l <= t.maxLen; l++) {
      const b = s.read(1); if (b === -1) return -1;
      c = (c << 1) | b;
      const k = (l << 16) | c;
      if (t.map[k] !== undefined) return t.map[k];
    }
    return -2;
  }
  return function inflate(u8) {
    const s = new BitStream(u8);
    const out = [];
    let bfinal = 0;
    let corrupted = false;
    try {
      while (!bfinal) {
        bfinal = s.read(1); const btype = s.read(2);
        if (bfinal === -1 || btype === -1) { corrupted = true; break; }
        if (btype === 0) {
          s.align(); const len = s.read(16); s.read(16);
          if (len === -1) { corrupted = true; break; }
          for (let i = 0; i < len; i++) out.push(s.buf[s.pos++] || 0);
        } else if (btype === 1 || btype === 2) {
          let lt, dt;
          if (btype === 1) { lt = buildTree(FIXED_LIT); dt = buildTree(FIXED_DIST); }
          else {
            const hl = s.read(5) + 257, hd = s.read(5) + 1, hc = s.read(4) + 4;
            if (hl < 257) { corrupted = true; break; }
            const cl = new Uint8Array(19);
            for (let i = 0; i < hc; i++) cl[CLEN_ORDER[i]] = s.read(3);
            const ct = buildTree(cl); if (!ct) { corrupted = true; break; }
            const unpack = (count) => {
              const r = [];
              while (r.length < count) {
                const sy = decodeSym(s, ct);
                if (sy < 0 || sy > 18) return null;
                if (sy < 16) r.push(sy);
                else if (sy === 16) { let c = 3 + s.read(2), p = r[r.length - 1]; while (c--) r.push(p); }
                else if (sy === 17) { let z = 3 + s.read(3); while (z--) r.push(0); }
                else if (sy === 18) { let z = 11 + s.read(7); while (z--) r.push(0); }
              }
              return new Uint8Array(r);
            };
            const ll = unpack(hl), dl = unpack(hd);
            if (!ll || !dl) { corrupted = true; break; }
            lt = buildTree(ll); dt = buildTree(dl);
          }
          if (!lt || !dt) { corrupted = true; break; }
          while (true) {
            const sym = decodeSym(s, lt);
            if (sym === -1 || sym === -2) { corrupted = true; break; }
            if (sym === 256) break;
            if (sym < 256) out.push(sym);
            else {
              const lc = sym - 257; if (lc > 28) { corrupted = true; break; }
              const len = LEN_BASE[lc] + s.read(LEN_EXTRA[lc]);
              const dc = decodeSym(s, dt); if (dc < 0) { corrupted = true; break; }
              const dist = DIST_BASE[dc] + s.read(DIST_EXTRA[dc]);
              if (dist > out.length) { corrupted = true; bfinal = 1; break; }
              let ptr = out.length - dist;
              for (let i = 0; i < len; i++) out.push(out[ptr++]);
            }
          }
        } else { corrupted = true; break; }
      }
    } catch (e) { corrupted = true; }
    return { data: new Uint8Array(out), isCorrupt: corrupted };
  };
})();

// ============================================================================
// Byte-level ZIP scanner — ported from Universal-File-Repair-Tool repairOffice.
// Walks PK\x03\x04 local file headers. Tries inflate at shifts 0..47.
// Survives corrupt central directory and partially-bad deflate streams.
// ============================================================================
function rateQuality(u8, name) {
  let valid = 0;
  const limit = Math.min(u8.length, 500);
  for (let i = 0; i < limit; i++) {
    if ((u8[i] >= 32 && u8[i] <= 126) || u8[i] === 10 || u8[i] === 13) valid++;
  }
  if (name && (name.endsWith('.xml') || name.endsWith('.rels'))) {
    const peek = Math.min(u8.length, 100);
    for (let i = 0; i < peek; i++) if (u8[i] === 60) { valid += 500; break; }
  }
  return valid;
}

function scanZipBytes(u8) {
  const view = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  let offset = 0;
  const recovered = {};
  const decoder = new TextDecoder('utf-8', { fatal: false });
  while (offset < u8.length - 30) {
    if (u8[offset] !== 0x50 || u8[offset + 1] !== 0x4b ||
        u8[offset + 2] !== 0x03 || u8[offset + 3] !== 0x04) {
      offset++; continue;
    }
    try {
      const meth = view.getUint16(offset + 8, true);
      const nl = view.getUint16(offset + 26, true);
      const el = view.getUint16(offset + 28, true);
      if (nl === 0 || nl > 512) { offset++; continue; }
      const name = decoder.decode(u8.subarray(offset + 30, offset + 30 + nl));
      if (name.endsWith('/')) { offset += 30 + nl + el; continue; }
      const dStart = offset + 30 + nl + el;
      let next = u8.length;
      for (let k = dStart; k < u8.length - 4; k++) {
        if (u8[k] === 0x50 && u8[k + 1] === 0x4b &&
            (u8[k + 2] === 0x01 || u8[k + 2] === 0x03 || u8[k + 2] === 0x05)) {
          next = k; break;
        }
      }
      const rawChunk = u8.subarray(dStart, next);
      let finalData = null, isCorrupt = false;
      if (meth === 0) {
        finalData = rawChunk; isCorrupt = false;
      } else if (meth === 8) {
        let bestRes = { data: new Uint8Array(0), isCorrupt: true }, bestScore = 0;
        for (let shift = 0; shift < 48 && shift < rawChunk.length; shift++) {
          const res = ImmortalInflate(rawChunk.subarray(shift));
          if (res.data.length > 0) {
            const score = rateQuality(res.data, name) + (res.isCorrupt ? 0 : 100);
            if (score > bestScore) { bestScore = score; bestRes = res; if (score > 1000) break; }
          }
        }
        finalData = bestRes.data; isCorrupt = bestRes.isCorrupt;
      }
      if (finalData && finalData.length > 0) {
        const prev = recovered[name];
        if (!prev || prev.data.length < finalData.length) {
          recovered[name] = { data: finalData, isCorrupt };
        }
      }
      offset = next;
    } catch (e) { offset++; }
  }
  return recovered;
}

// ============================================================================
// XML repair from the original SourceForge tool's InvalidTags / ValidTags.
// ============================================================================
const InvalidTags = {
  omathWps: '<mc:AlternateContent><mc:Choice Requires="wps"><m:oMath>',
  omathWpg: '<mc:AlternateContent><mc:Choice Requires="wpg"><m:oMath>',
  omathWpi: '<mc:AlternateContent><mc:Choice Requires="wpi"><m:oMath>',
  omathWpc: '<mc:AlternateContent><mc:Choice Requires="wpc"><m:oMath>',
  vshape:   '</w:txbxContent></w:pict></mc:Fallback></mc:AlternateContent>',
  mcChoiceRe: /(<\/mc:Choice>)(<(.).*?(\/>|>))/g,
  fallbackRe: /(<mc:Fallback><w:pict\/>)(<(.).*?(\/>|>))/g,
};
const ValidTags = {
  omathWps: '<m:oMath><mc:AlternateContent><mc:Choice Requires="wps">',
  omathWpg: '<m:oMath><mc:AlternateContent><mc:Choice Requires="wpg">',
  omathWpi: '<m:oMath><mc:AlternateContent><mc:Choice Requires="wpi">',
  omathWpc: '<m:oMath><mc:AlternateContent><mc:Choice Requires="wpc">',
  vshape:   '</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent>',
  mcChoice3: '</mc:Choice></mc:AlternateContent>',
  mcChoice4: '</mc:Choice></mc:AlternateContent></w:r>',
  omitFallback: '</mc:AlternateContent></w:r>',
};

function applyDocumentXmlTagFixes(xml) {
  let s = xml;
  let changes = 0;
  const replaceAll = (needle, repl) => {
    if (s.indexOf(needle) === -1) return;
    let n = 0; let out = '';
    let i = 0;
    while (true) {
      const j = s.indexOf(needle, i);
      if (j === -1) { out += s.slice(i); break; }
      out += s.slice(i, j) + repl;
      i = j + needle.length;
      n++;
    }
    s = out; changes += n;
  };
  replaceAll(InvalidTags.omathWps, ValidTags.omathWps);
  replaceAll(InvalidTags.omathWpg, ValidTags.omathWpg);
  replaceAll(InvalidTags.omathWpi, ValidTags.omathWpi);
  replaceAll(InvalidTags.omathWpc, ValidTags.omathWpc);
  replaceAll(InvalidTags.vshape,   ValidTags.vshape);
  s = s.replace(InvalidTags.mcChoiceRe, (m, g1, g2) => { changes++; return ValidTags.mcChoice3 + g2; });
  s = s.replace(InvalidTags.fallbackRe, (m, g1, g2) => { changes++; return ValidTags.omitFallback + g2; });
  return { xml: s, changes };
}

function healXMLStrict(xml) {
  xml = xml.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
  const validStart = xml.search(/<\?xml|<[a-zA-Z0-9]+:/);
  if (validStart > -1) xml = xml.substring(validStart);
  else { const f = xml.indexOf('<'); if (f > -1) xml = xml.substring(f); }
  let out = '', inTag = false, inQuote = false, quoteChar = null;
  for (let i = 0; i < xml.length; i++) {
    const c = xml[i], next = xml[i + 1] || '';
    if (!inTag) {
      if (c === '<') {
        if (/[a-zA-Z0-9_:\/?!]/.test(next)) { inTag = true; out += c; }
        else { out += '&lt;'; }
      } else if (c === '>') { out += '&gt;'; }
      else { out += c; }
    } else {
      if (inQuote) { if (c === quoteChar) { inQuote = false; quoteChar = null; } out += c; }
      else {
        if (c === '"' || c === "'") { inQuote = true; quoteChar = c; out += c; }
        else if (c === '>') { inTag = false; out += c; }
        else if (c === '<') { out += '>'; inTag = false; i--; }
        else { out += c; }
      }
    }
  }
  if (inQuote) out += quoteChar;
  if (inTag) out += '>';
  const stack = [], re = /<\/?([a-zA-Z0-9_:\-]+)[^>]*(\/?)>/g;
  let m;
  while ((m = re.exec(out)) !== null) {
    const tag = m[1], full = m[0];
    if (full.startsWith('</')) { if (stack.length && stack[stack.length - 1] === tag) stack.pop(); }
    else if (!full.endsWith('/>') && m[2] !== '/' && !full.startsWith('<?')) { stack.push(tag); }
  }
  while (stack.length) out += `</${stack.pop()}>`;
  return out;
}

// ============================================================================
// Text extractors
// ============================================================================
function extractTextFromWordXml(xml) {
  if (!xml) return '';
  const out = [];
  const tagRe = /<w:(t|tab|br|p|cr)(\s[^>]*)?(\/>|>)/g;
  let m;
  let pendingClose = -1;
  while ((m = tagRe.exec(xml)) !== null) {
    const tag = m[1];
    const selfClose = m[3] === '/>';
    if (tag === 't' && !selfClose) {
      const start = tagRe.lastIndex;
      const end = xml.indexOf('</w:t>', start);
      if (end === -1) break;
      out.push(xml.slice(start, end)
        .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
        .replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'"));
      tagRe.lastIndex = end + 6;
    } else if (tag === 'tab') {
      out.push('\t');
    } else if (tag === 'br' || tag === 'cr') {
      out.push('\n');
    } else if (tag === 'p') {
      if (!selfClose) {
        out.push('\n');
      }
    }
  }
  return out.join('').replace(/\n{3,}/g, '\n\n').trim();
}

function extractTextFromOdtXml(xml) {
  if (!xml) return '';
  let s = xml.replace(/<text:line-break\s*\/>/g, '\n')
    .replace(/<text:tab\s*\/>/g, '\t')
    .replace(/<\/text:p>/g, '\n')
    .replace(/<\/text:h>/g, '\n');
  s = s.replace(/<[^>]+>/g, '');
  s = s.replace(/&lt;/g, '<').replace(/&gt;/g, '>')
       .replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
  return s.replace(/\n{3,}/g, '\n\n').trim();
}

function extractTextFromRtf(bytes) {
  let s = '';
  if (typeof bytes === 'string') s = bytes;
  else s = new TextDecoder('latin1').decode(bytes);
  if (!s.startsWith('{\\rtf')) return '';
  s = s.replace(/\\u(-?\d+)\??/g, (m, n) => {
    let code = parseInt(n, 10);
    if (code < 0) code += 65536;
    try { return String.fromCharCode(code); } catch (e) { return ''; }
  });
  s = s.replace(/\\'([0-9a-fA-F]{2})/g, (m, h) => String.fromCharCode(parseInt(h, 16)));
  s = s.replace(/\\par[d ]?/g, '\n').replace(/\\line ?/g, '\n').replace(/\\tab ?/g, '\t');
  s = s.replace(/\\[a-zA-Z]+-?\d* ?/g, '');
  s = s.replace(/\\\*/g, '').replace(/[{}]/g, '');
  s = s.replace(/\\\\/g, '\\').replace(/\\~/g, ' ');
  return s.replace(/\n{3,}/g, '\n\n').trim();
}

function stringsScan(u8) {
  const minLen = 4;
  const ascii = [];
  let buf = '';
  for (let i = 0; i < u8.length; i++) {
    const b = u8[i];
    if ((b >= 32 && b <= 126) || b === 9) buf += String.fromCharCode(b);
    else { if (buf.length >= minLen) ascii.push(buf); buf = ''; }
  }
  if (buf.length >= minLen) ascii.push(buf);
  const utf16 = [];
  buf = '';
  for (let i = 0; i + 1 < u8.length; i += 2) {
    const lo = u8[i], hi = u8[i + 1];
    if (hi === 0 && ((lo >= 32 && lo <= 126) || lo === 9 || lo === 10 || lo === 13)) {
      buf += String.fromCharCode(lo);
    } else {
      if (buf.length >= minLen) utf16.push(buf);
      buf = '';
    }
  }
  if (buf.length >= minLen) utf16.push(buf);
  const all = [...ascii, ...utf16];
  const seen = new Set();
  const unique = [];
  for (const s of all) {
    const key = s.replace(/\s+/g, ' ').trim();
    if (key.length >= minLen && !seen.has(key)) { seen.add(key); unique.push(key); }
  }
  return unique.join('\n');
}

// ============================================================================
// Method runners
// ============================================================================
async function method1_StandardParse(u8, ext) {
  if (typeof JSZip === 'undefined') throw new Error('JSZip not loaded');
  const zip = await JSZip.loadAsync(u8);
  const files = {};
  const wantedDocx = ['word/document.xml', 'word/header1.xml', 'word/header2.xml', 'word/header3.xml',
    'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml',
    'word/footnotes.xml', 'word/endnotes.xml', 'word/comments.xml'];
  const wantedOdt = ['content.xml'];
  const wanted = ext === 'odt' ? wantedOdt : wantedDocx;
  let mainXml = '';
  for (const name of wanted) {
    const f = zip.file(name);
    if (f) {
      const txt = await f.async('text');
      files[name] = txt;
      if (name === 'word/document.xml' || name === 'content.xml') mainXml = txt;
    }
  }
  const text = ext === 'odt' ? extractTextFromOdtXml(mainXml) : extractTextFromWordXml(mainXml);
  return { text, files, mainXml };
}

function method2_ByteLevelRecovery(u8) {
  const recovered = scanZipBytes(u8);
  const named = {};
  let mainXml = '';
  let hasDocument = false;
  for (const [name, info] of Object.entries(recovered)) {
    let data = info.data;
    if (name.endsWith('.xml') || name.endsWith('.rels')) {
      let s = new TextDecoder('utf-8', { fatal: false }).decode(data);
      s = healXMLStrict(s);
      data = new TextEncoder().encode(s);
    }
    named[name] = { data, isCorrupt: info.isCorrupt };
    if (name === 'word/document.xml') {
      mainXml = new TextDecoder('utf-8', { fatal: false }).decode(data);
      hasDocument = true;
    }
  }
  const text = hasDocument ? extractTextFromWordXml(mainXml) : '';
  return { text, files: named, mainXml, hasDocument, count: Object.keys(named).length };
}

function method3_XmlTagFixes(mainXml) {
  if (!mainXml) return { text: '', xml: '', changes: 0 };
  const fixed = applyDocumentXmlTagFixes(mainXml);
  const text = extractTextFromWordXml(fixed.xml);
  return { text, xml: fixed.xml, changes: fixed.changes };
}

function method4_Rtf(u8) {
  const text = extractTextFromRtf(u8);
  return { text };
}

function method5_Strings(u8) {
  const text = stringsScan(u8);
  return { text };
}

// ============================================================================
// Repackage as a fresh DOCX
// ============================================================================
const DUMMY_DOCX = {
  TYPES: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '</Types>',
  RELS: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>',
  DOC_RELS: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>',
  STYLES: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
};

function buildBlankDocumentXml(text) {
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const paras = text.split(/\n/).map(line => {
    if (!line) return '<w:p/>';
    const runs = line.split('\t').map(seg =>
      seg ? `<w:r><w:t xml:space="preserve">${esc(seg)}</w:t></w:r>` : '<w:r><w:tab/></w:r>'
    ).join('<w:r><w:tab/></w:r>');
    return `<w:p>${runs}</w:p>`;
  }).join('');
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${paras}<w:sectPr/></w:body></w:document>`;
}

async function buildRepairedDocx(method2Files, fixedDocXml, recoveredText) {
  const zip = new JSZip();
  if (method2Files && Object.keys(method2Files).length > 0) {
    for (const [name, info] of Object.entries(method2Files)) {
      if (name === 'word/document.xml') continue;
      zip.file(name, info.data);
    }
    if (!method2Files['[Content_Types].xml']) zip.file('[Content_Types].xml', DUMMY_DOCX.TYPES);
    if (!method2Files['_rels/.rels']) zip.file('_rels/.rels', DUMMY_DOCX.RELS);
    if (!method2Files['word/_rels/document.xml.rels']) zip.file('word/_rels/document.xml.rels', DUMMY_DOCX.DOC_RELS);
    if (!method2Files['word/styles.xml']) zip.file('word/styles.xml', DUMMY_DOCX.STYLES);
    const docXml = fixedDocXml || buildBlankDocumentXml(recoveredText || '');
    zip.file('word/document.xml', docXml);
  } else {
    zip.file('[Content_Types].xml', DUMMY_DOCX.TYPES);
    zip.file('_rels/.rels', DUMMY_DOCX.RELS);
    zip.file('word/_rels/document.xml.rels', DUMMY_DOCX.DOC_RELS);
    zip.file('word/styles.xml', DUMMY_DOCX.STYLES);
    zip.file('word/document.xml', fixedDocXml || buildBlankDocumentXml(recoveredText || ''));
  }
  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

// ============================================================================
// Main pipeline
// ============================================================================
async function recoverFile(file) {
  const buf = await file.arrayBuffer();
  const u8 = new Uint8Array(buf);
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  const results = {
    methods: [],
    raws: [],
    bestText: '',
    bestMethod: '',
    method2Files: null,
    method3Xml: '',
  };

  let m1 = null;
  try {
    m1 = await method1_StandardParse(u8, ext);
    results.methods.push({
      name: 'Standard parse (JSZip)',
      desc: 'Treat the file as a normal Office package and read text from document.xml.',
      status: m1.text ? 'success' : 'partial',
      detail: m1.text ? `${m1.text.length.toLocaleString()} chars from ${Object.keys(m1.files).length} XML parts` : 'Parsed but no text found',
    });
    results.raws.push({ name: 'Standard parse', body: m1.text });
  } catch (e) {
    results.methods.push({
      name: 'Standard parse (JSZip)',
      desc: 'Treat the file as a normal Office package.',
      status: 'failed',
      detail: e.message || String(e),
    });
    results.raws.push({ name: 'Standard parse', body: '' });
  }

  let m2 = null;
  try {
    m2 = method2_ByteLevelRecovery(u8);
    results.method2Files = m2.files;
    results.methods.push({
      name: 'Byte-level ZIP recovery',
      desc: 'Walks PK\\x03\\x04 local file headers and inflates each entry with shift sweep 0..47. Recovers files even when the central directory is corrupt or individual deflate streams are damaged.',
      status: m2.hasDocument && m2.text ? 'success' : (m2.count > 0 ? 'partial' : 'failed'),
      detail: `Recovered ${m2.count} entries${m2.hasDocument ? ', including word/document.xml' : ''}.`,
    });
    results.raws.push({ name: 'Byte-level ZIP recovery', body: m2.text || '(no document.xml recovered)' });
  } catch (e) {
    results.methods.push({
      name: 'Byte-level ZIP recovery',
      desc: 'Walks PK\\x03\\x04 local file headers and inflates each entry.',
      status: 'failed',
      detail: e.message || String(e),
    });
    results.raws.push({ name: 'Byte-level ZIP recovery', body: '' });
  }

  let m3 = null;
  const seedXml = (m2 && m2.mainXml) || (m1 && m1.mainXml) || '';
  if (seedXml) {
    m3 = method3_XmlTagFixes(seedXml);
    results.method3Xml = m3.xml;
    results.methods.push({
      name: 'DOCX XML tag fixes',
      desc: 'Applies the original SourceForge tool’s InvalidTags → ValidTags substitutions to word/document.xml (mc:AlternateContent / oMath / Fallback / vshape).',
      status: m3.text ? (m3.changes > 0 ? 'success' : 'partial') : 'failed',
      detail: m3.changes > 0 ? `${m3.changes} tag substitution(s) applied; ${m3.text.length.toLocaleString()} chars` : 'No invalid tag patterns matched',
    });
    results.raws.push({ name: 'DOCX XML tag fixes', body: m3.text });
  } else {
    results.methods.push({
      name: 'DOCX XML tag fixes',
      desc: 'Applies the original tool’s InvalidTags → ValidTags substitutions.',
      status: 'skipped',
      detail: 'No document.xml available to fix.',
    });
    results.raws.push({ name: 'DOCX XML tag fixes', body: '' });
  }

  let m4 = null;
  try {
    m4 = method4_Rtf(u8);
    if (m4.text) {
      results.methods.push({
        name: 'RTF extraction',
        desc: 'Strips control words and unescapes hex/Unicode runs.',
        status: 'success',
        detail: `${m4.text.length.toLocaleString()} chars extracted`,
      });
      results.raws.push({ name: 'RTF extraction', body: m4.text });
    } else {
      results.methods.push({
        name: 'RTF extraction',
        desc: 'Strips control words and unescapes hex/Unicode runs.',
        status: 'skipped',
        detail: 'File does not start with {\\rtf.',
      });
      results.raws.push({ name: 'RTF extraction', body: '' });
    }
  } catch (e) {
    results.methods.push({ name: 'RTF extraction', desc: '', status: 'failed', detail: e.message });
    results.raws.push({ name: 'RTF extraction', body: '' });
  }

  let m5 = null;
  try {
    m5 = method5_Strings(u8);
    results.methods.push({
      name: 'Strings scan (last resort)',
      desc: 'Scans raw bytes for runs of printable ASCII and UTF-16LE — works on .doc, renamed temp files, or any blob.',
      status: m5.text ? 'partial' : 'failed',
      detail: m5.text ? `${m5.text.split('\n').length.toLocaleString()} unique runs (${m5.text.length.toLocaleString()} chars total)` : 'No printable runs found',
    });
    results.raws.push({ name: 'Strings scan', body: m5.text });
  } catch (e) {
    results.methods.push({ name: 'Strings scan (last resort)', desc: '', status: 'failed', detail: e.message });
    results.raws.push({ name: 'Strings scan', body: '' });
  }

  const candidates = [
    { name: 'Method 1 (standard parse)', text: m1 && m1.text || '', score: (m1 && m1.text || '').length * 2 },
    { name: 'Method 3 (XML tag fixes)',  text: m3 && m3.text || '', score: (m3 && m3.text || '').length * 1.5 },
    { name: 'Method 2 (byte-level recovery)', text: m2 && m2.text || '', score: (m2 && m2.text || '').length * 1.4 },
    { name: 'Method 4 (RTF extraction)',  text: m4 && m4.text || '', score: (m4 && m4.text || '').length },
    { name: 'Method 5 (strings scan)',   text: m5 && m5.text || '', score: (m5 && m5.text || '').length * 0.5 },
  ];
  candidates.sort((a, b) => b.score - a.score);
  results.bestText = candidates[0].text;
  results.bestMethod = candidates[0].name;
  return results;
}

// ============================================================================
// Manual recovery instructions (per OS) — replicating the original tool's links
// ============================================================================
const MANUAL = {
  windows: `
    <h3>1. AutoRecover files</h3>
    <p>Word saves AutoRecover snapshots as <code>.asd</code> files. Open File Explorer and paste these into the address bar:</p>
    <ul>
      <li><code>%AppData%\\Microsoft\\Word\\</code></li>
      <li><code>%LocalAppData%\\Microsoft\\Office\\UnsavedFiles\\</code></li>
      <li><code>%UserProfile%\\AppData\\Local\\Microsoft\\Office\\UnsavedFiles\\</code></li>
    </ul>
    <h3>2. Temporary files</h3>
    <ul>
      <li><code>%Temp%</code> — search for <code>~WRD*.tmp</code>, <code>~WRL*.tmp</code>, <code>~$*.docx</code>.</li>
      <li>In Word: <strong>File → Info → Manage Document → Recover Unsaved Documents</strong>.</li>
    </ul>
    <h3>3. Previous Versions / Shadow Copies</h3>
    <p>Right-click the file (or its folder) → <strong>Properties → Previous Versions</strong>. Requires System Protection / File History to be enabled.</p>
    <p>For older Windows, try <a href="http://www.shadowexplorer.com/" target="_blank" rel="noopener">ShadowExplorer</a>.</p>
    <h3>4. Open and Repair</h3>
    <ol>
      <li>In Word: <strong>File → Open → Browse</strong>.</li>
      <li>Select the file, click the arrow next to the <strong>Open</strong> button, choose <strong>Open and Repair</strong>.</li>
    </ol>
    <h3>5. Recover Text from Any File</h3>
    <p>In Word: <strong>File → Options → General</strong>, enable <em>Confirm file format conversion on open</em>. Then File → Open, change the type filter to <em>Recover Text from Any File (*.*)</em>.</p>
    <h3>6. Microsoft’s official guidance</h3>
    <ul>
      <li><a href="https://support.microsoft.com/en-us/office/recover-an-earlier-version-of-an-office-file-851b6ea8-1d12-43e2-b3f2-1095bb16cee2" target="_blank" rel="noopener">Recover an earlier version of an Office file</a></li>
      <li><a href="https://support.microsoft.com/en-us/office/recover-files-in-office-365-61f5be25-4cdc-4d70-9d12-5a4d4070ddc6" target="_blank" rel="noopener">Recover files in Office</a></li>
      <li><a href="https://support.microsoft.com/en-us/topic/how-to-troubleshoot-damaged-documents-in-word-b3b1986b-0c45-4ec2-baef-1e9a222badf3" target="_blank" rel="noopener">Troubleshoot damaged documents in Word (KB)</a></li>
    </ul>
    <h3>7. Deleted-file recovery utilities</h3>
    <ul>
      <li><a href="https://www.ccleaner.com/recuva" target="_blank" rel="noopener">Recuva</a></li>
      <li><a href="https://www.cgsecurity.org/wiki/PhotoRec" target="_blank" rel="noopener">PhotoRec</a></li>
      <li><a href="https://www.voidtools.com/" target="_blank" rel="noopener">Everything (find files by name)</a></li>
    </ul>
  `,
  macos: `
    <h3>1. AutoRecovery files</h3>
    <p>Word for Mac stores AutoRecovery copies under your Library folder:</p>
    <ul>
      <li><code>~/Library/Containers/com.microsoft.Word/Data/Library/Preferences/AutoRecovery/</code></li>
      <li><code>~/Library/Application Support/Microsoft/Office/Office 2011 AutoRecovery/</code> (Office 2011)</li>
    </ul>
    <p>In Finder: <strong>Go → Go to Folder…</strong> and paste the path. Hidden Library: hold <strong>Option</strong> while clicking the Go menu.</p>
    <h3>2. Temporary files</h3>
    <ul>
      <li><code>/private/var/folders/</code> — Word’s scratch directory.</li>
      <li><code>/tmp/</code></li>
    </ul>
    <h3>3. Time Machine</h3>
    <p>Open the file’s containing folder, then click the Time Machine icon to browse snapshots.</p>
    <h3>4. Open and Repair (Word for Mac 2016+)</h3>
    <ol>
      <li>Open Word, then <strong>File → Open</strong>.</li>
      <li>Select the file, click the arrow next to the <strong>Open</strong> button, choose <strong>Open and Repair</strong>.</li>
    </ol>
    <h3>5. Try LibreOffice</h3>
    <p><a href="https://www.libreoffice.org/" target="_blank" rel="noopener">LibreOffice</a> is often more permissive than Word and may open files Word can’t.</p>
  `,
  linux: `
    <h3>1. Try LibreOffice</h3>
    <p>Open the file in <a href="https://www.libreoffice.org/" target="_blank" rel="noopener">LibreOffice Writer</a> — usually more tolerant of damaged DOCX/DOC than Microsoft Word.</p>
    <h3>2. Repair the ZIP container manually</h3>
    <p>A <code>.docx</code> is a ZIP archive. Use <code>zip</code>’s self-repair:</p>
    <pre><code>cp damaged.docx broken.zip
zip -FF broken.zip --out fixed.zip
mv fixed.zip recovered.docx</code></pre>
    <h3>3. Extract text directly</h3>
    <pre><code>unzip -p damaged.docx word/document.xml | sed 's/&lt;[^&gt;]*&gt;//g'
# or
pandoc damaged.docx -o recovered.txt
# or
antiword old.doc &gt; recovered.txt</code></pre>
    <h3>4. Strings on .doc binaries</h3>
    <pre><code>strings -n 4 damaged.doc | less
# UTF-16 (Word’s default)
strings -n 4 -e l damaged.doc</code></pre>
    <h3>5. Deleted-file recovery</h3>
    <ul>
      <li><a href="https://www.cgsecurity.org/wiki/PhotoRec" target="_blank" rel="noopener">PhotoRec</a></li>
      <li><a href="https://github.com/sleuthkit/sleuthkit" target="_blank" rel="noopener">The Sleuth Kit</a></li>
    </ul>
  `,
  ios: `
    <h3>1. iCloud version history</h3>
    <p>Open the document in Word for iOS, then <strong>… menu → Restore → Browse Version History</strong>.</p>
    <h3>2. Files app → Recently Deleted</h3>
    <p>Open the <strong>Files</strong> app, tap <strong>Browse → Recently Deleted</strong>. Files stay there for ~30 days.</p>
    <h3>3. Email yourself the corrupted file</h3>
    <p>Then open it on a desktop machine where you can run this Web app or Word’s Open and Repair.</p>
    <h3>4. Online viewers</h3>
    <ul>
      <li><a href="https://office.live.com/start/Word.aspx" target="_blank" rel="noopener">Microsoft Word Online</a> — sometimes opens what the desktop refuses.</li>
      <li><a href="https://docs.google.com/" target="_blank" rel="noopener">Google Docs</a></li>
    </ul>
  `,
  android: `
    <h3>1. Word for Android version history</h3>
    <p>Open the document, tap the overflow (⋮) menu → <strong>History → See version history</strong>.</p>
    <h3>2. Google Drive trash</h3>
    <p>If the file lived in Google Drive, check <strong>Drive → Trash</strong>. Files stay there for 30 days.</p>
    <h3>3. Local temp / cache</h3>
    <p>Use a file manager to browse <code>/sdcard/Android/data/com.microsoft.office.word/cache/</code> (some Word versions store recovery copies there).</p>
    <h3>4. Online viewers</h3>
    <ul>
      <li><a href="https://office.live.com/start/Word.aspx" target="_blank" rel="noopener">Word Online</a></li>
      <li><a href="https://docs.google.com/" target="_blank" rel="noopener">Google Docs</a></li>
    </ul>
  `,
};

// ============================================================================
// UI wiring
// ============================================================================
const $ = (sel) => document.querySelector(sel);
const state = { lastResults: null, currentFileName: '', currentFileBuf: null };

function setProgress(pct, label) {
  const bar = $('#progressBar'); if (bar) bar.style.width = `${Math.max(0, Math.min(100, pct))}%`;
  const txt = $('#progressText'); if (txt) txt.textContent = label || '';
}

function fmtBytes(n) {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1024 / 1024).toFixed(2)} MB`;
}

async function handleFile(file) {
  state.currentFileName = file.name;
  state.currentFileBuf = await file.arrayBuffer();
  $('#status').hidden = false;
  $('#filename').textContent = file.name;
  $('#filemeta').textContent = `${fmtBytes(file.size)} · ${file.type || 'unknown type'}`;
  setProgress(10, 'Reading file…');

  setProgress(30, 'Running recovery methods…');
  let results;
  try {
    results = await recoverFile(file);
  } catch (e) {
    setProgress(0, `Error: ${e.message || e}`);
    return;
  }
  state.lastResults = results;
  setProgress(100, `Done. Best: ${results.bestMethod}.`);

  $('#results').hidden = false;
  renderMethods(results.methods);
  renderRaws(results.raws);
  $('#recoveredText').value = results.bestText || '';
  const words = (results.bestText || '').trim().split(/\s+/).filter(Boolean).length;
  $('#recoveredStats').textContent = results.bestText
    ? `${results.bestText.length.toLocaleString()} chars · ${words.toLocaleString()} words · ${results.bestMethod}`
    : 'No text recovered. See Methods used and Manual recovery tabs.';
  showOsInstructions(detectOs());
}

function renderMethods(methods) {
  const ul = $('#methodsList'); ul.innerHTML = '';
  for (const m of methods) {
    const li = document.createElement('li');
    li.innerHTML = `<span class="method-badge ${m.status}">${m.status}</span>` +
      `<div class="method-info"><h4>${m.name}</h4>` +
      `<p>${m.desc}</p>` +
      (m.detail ? `<p class="muted">${m.detail}</p>` : '') +
      `</div>`;
    ul.appendChild(li);
  }
}

function renderRaws(raws) {
  const div = $('#rawList'); div.innerHTML = '';
  for (const r of raws) {
    const det = document.createElement('details');
    const sum = document.createElement('summary');
    sum.textContent = `${r.name} (${(r.body || '').length.toLocaleString()} chars)`;
    det.appendChild(sum);
    const pre = document.createElement('pre');
    pre.textContent = r.body || '(empty)';
    det.appendChild(pre);
    div.appendChild(det);
  }
}

function detectOs() {
  const ua = (navigator.userAgent || '').toLowerCase();
  if (/iphone|ipad|ipod/.test(ua)) return 'ios';
  if (/android/.test(ua)) return 'android';
  if (/mac/.test(ua)) return 'macos';
  if (/linux/.test(ua)) return 'linux';
  return 'windows';
}

function showOsInstructions(os) {
  const c = $('#manualContent'); if (!c) return;
  c.innerHTML = MANUAL[os] || MANUAL.windows;
  document.querySelectorAll('.ostab').forEach(b => {
    b.classList.toggle('active', b.dataset.os === os);
  });
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

function baseName() {
  return (state.currentFileName || 'document').replace(/\.[^.]+$/, '') || 'document';
}

if (typeof document !== 'undefined') document.addEventListener('DOMContentLoaded', () => {
  const drop = $('#drop');
  const file = $('#file');
  if (drop) {
    ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, (e) => {
      e.preventDefault(); drop.classList.add('dragover');
    }));
    ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, (e) => {
      e.preventDefault(); drop.classList.remove('dragover');
    }));
    drop.addEventListener('drop', (e) => {
      const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) handleFile(f);
    });
  }
  if (file) file.addEventListener('change', (e) => {
    const f = e.target.files && e.target.files[0]; if (f) handleFile(f);
  });

  document.querySelectorAll('.tab').forEach(t => t.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(x => x.classList.toggle('active', x === t));
    document.querySelectorAll('.panel').forEach(p => p.classList.toggle('active', p.dataset.panel === t.dataset.tab));
  }));
  document.querySelectorAll('.ostab').forEach(t => t.addEventListener('click', () => showOsInstructions(t.dataset.os)));

  const copyBtn = $('#copyBtn');
  if (copyBtn) copyBtn.addEventListener('click', async () => {
    const txt = $('#recoveredText').value;
    try {
      await navigator.clipboard.writeText(txt);
      copyBtn.textContent = 'Copied!';
      setTimeout(() => copyBtn.textContent = 'Copy', 1500);
    } catch (e) {
      $('#recoveredText').select(); document.execCommand('copy');
    }
  });

  const txtBtn = $('#downloadTxtBtn');
  if (txtBtn) txtBtn.addEventListener('click', () => {
    const txt = $('#recoveredText').value;
    downloadBlob(new Blob([txt], { type: 'text/plain;charset=utf-8' }), `${baseName()}_recovered.txt`);
  });

  const docxBtn = $('#downloadDocxBtn');
  if (docxBtn) docxBtn.addEventListener('click', async () => {
    const r = state.lastResults;
    const txt = $('#recoveredText').value;
    docxBtn.textContent = 'Building…'; docxBtn.disabled = true;
    try {
      const blob = await buildRepairedDocx(
        r ? r.method2Files : null,
        r ? r.method3Xml : '',
        txt
      );
      downloadBlob(blob, `${baseName()}_repaired.docx`);
    } catch (e) {
      alert('Could not build DOCX: ' + (e.message || e));
    } finally {
      docxBtn.textContent = 'Download .docx'; docxBtn.disabled = false;
    }
  });
});

// ============================================================================
// Node smoke-test export
// ============================================================================
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    ImmortalInflate, scanZipBytes, rateQuality, healXMLStrict,
    applyDocumentXmlTagFixes, extractTextFromWordXml, extractTextFromOdtXml,
    extractTextFromRtf, stringsScan,
    method1_StandardParse, method2_ByteLevelRecovery, method3_XmlTagFixes,
    method4_Rtf, method5_Strings, recoverFile, buildRepairedDocx,
    InvalidTags, ValidTags,
  };
}
