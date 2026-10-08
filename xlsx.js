// Minimal dependency-free .xlsx reader/writer (zip + SpreadsheetML).
// Runs in the browser and in Node 18+ (uses the built-in Compression/DecompressionStream).
(function () {
  const enc = new TextEncoder(), dec = new TextDecoder();

  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(buf) {
    let c = 0xffffffff;
    for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 255] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  async function pipe(u8, stream) {
    return new Uint8Array(await new Response(new Blob([u8]).stream().pipeThrough(stream)).arrayBuffer());
  }
  const inflateRaw = (u8) => pipe(u8, new DecompressionStream('deflate-raw'));
  const deflateRaw = (u8) => pipe(u8, new CompressionStream('deflate-raw'));

  const concat = (parts) => {
    const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
    let o = 0;
    for (const p of parts) { out.set(p, o); o += p.length; }
    return out;
  };

  // ---------- zip ----------

  async function unzip(u8) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let e = u8.length - 22;
    while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
    if (e < 0) throw new Error('That file is not a valid .xlsx workbook');
    const count = dv.getUint16(e + 10, true);
    let p = dv.getUint32(e + 16, true);
    const files = {};
    for (let i = 0; i < count; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('Corrupt .xlsx file');
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
      const off = dv.getUint32(p + 42, true);
      const name = dec.decode(u8.subarray(p + 46, p + 46 + nl));
      const start = off + 30 + dv.getUint16(off + 26, true) + dv.getUint16(off + 28, true);
      const data = u8.subarray(start, start + csize);
      files[name] = method === 0 ? data : await inflateRaw(data);
      p += 46 + nl + el + cl;
    }
    return files;
  }

  async function zip(entries) {
    const parts = [], central = [];
    let off = 0;
    for (const { name, data } of entries) {
      const nb = enc.encode(name), comp = await deflateRaw(data), crc = crc32(data);
      const lh = new Uint8Array(30), l = new DataView(lh.buffer);
      l.setUint32(0, 0x04034b50, true); l.setUint16(4, 20, true); l.setUint16(6, 0x0800, true); l.setUint16(8, 8, true);
      l.setUint16(12, 0x21, true); l.setUint32(14, crc, true); l.setUint32(18, comp.length, true);
      l.setUint32(22, data.length, true); l.setUint16(26, nb.length, true);
      parts.push(lh, nb, comp);
      const ch = new Uint8Array(46), c = new DataView(ch.buffer);
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
      c.setUint16(10, 8, true); c.setUint16(14, 0x21, true); c.setUint32(16, crc, true); c.setUint32(20, comp.length, true);
      c.setUint32(24, data.length, true); c.setUint16(28, nb.length, true); c.setUint32(42, off, true);
      central.push(ch, nb);
      off += 30 + nb.length + comp.length;
    }
    const cd = concat(central);
    const end = new Uint8Array(22), d = new DataView(end.buffer);
    d.setUint32(0, 0x06054b50, true); d.setUint16(8, entries.length, true); d.setUint16(10, entries.length, true);
    d.setUint32(12, cd.length, true); d.setUint32(16, off, true);
    return concat([...parts, cd, end]);
  }

  // ---------- XML helpers ----------

  const unescapeXml = (s) =>
    s.replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|quot|apos|amp);/g, (_, m) => {
      if (m[0] === '#') return String.fromCodePoint(m[1] === 'x' ? parseInt(m.slice(2), 16) : parseInt(m.slice(1), 10));
      return { lt: '<', gt: '>', quot: '"', apos: "'", amp: '&' }[m];
    });
  const escapeXml = (s) =>
    String(s).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const attrs = (s) => {
    const o = {};
    for (const m of s.matchAll(/([\w:]+)="([^"]*)"/g)) o[m[1]] = unescapeXml(m[2]);
    return o;
  };
  const textOf = (xml) => [...xml.matchAll(/<t\b[^>]*?(?:\/>|>([\s\S]*?)<\/t>)/g)].map((m) => unescapeXml(m[1] || '')).join('');

  function colIndex(ref) {
    let n = 0;
    for (const ch of ref.replace(/[^A-Z]/gi, '').toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
    return n - 1;
  }
  function colName(i) {
    let s = '';
    for (i++; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
    return s;
  }

  // ---------- read ----------

  // Returns [{ name, rows: string[][] }]
  async function readXlsx(u8) {
    const files = await unzip(u8);
    const get = (n) => (files[n] ? dec.decode(files[n]) : null);
    const wb = get('xl/workbook.xml');
    if (!wb) throw new Error('That file is not a valid .xlsx workbook');

    const sst = [];
    const sstXml = get('xl/sharedStrings.xml');
    if (sstXml) for (const m of sstXml.matchAll(/<si\b[^>]*?(?:\/>|>([\s\S]*?)<\/si>)/g)) sst.push(textOf(m[1] || ''));

    const rels = {};
    for (const m of (get('xl/_rels/workbook.xml.rels') || '').matchAll(/<Relationship\b([^>]*?)\/?>/g)) {
      const a = attrs(m[1]);
      rels[a.Id] = a.Target;
    }

    const sheets = [];
    for (const m of wb.matchAll(/<sheet\b([^>]*?)\/?>/g)) {
      const a = attrs(m[1]);
      let target = rels[a['r:id']];
      if (!target) continue;
      target = target.startsWith('/') ? target.slice(1) : 'xl/' + target;
      const xml = get(target);
      if (!xml) continue;
      const rows = [];
      for (const rm of xml.matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
        const row = [];
        for (const cm of (rm[2] || '').matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
          const ca = attrs(cm[1]);
          const inner = cm[2] || '';
          const v = inner.match(/<v>([\s\S]*?)<\/v>/);
          let val = '';
          if (ca.t === 'inlineStr') val = textOf(inner);
          else if (ca.t === 's') val = v ? sst[Number(v[1])] ?? '' : '';
          else if (v) val = unescapeXml(v[1]);
          row[colIndex(ca.r || 'A1')] = val;
        }
        rows.push(Array.from(row, (x) => x ?? ''));
      }
      sheets.push({ name: a.name, rows });
    }
    return sheets;
  }

  // ---------- write ----------

  function sheetXml(rows) {
    const out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData>'];
    rows.forEach((row, r) => {
      out.push(`<row r="${r + 1}">`);
      row.forEach((v, c) => {
        if (v === '' || v == null) return;
        const ref = colName(c) + (r + 1);
        if (typeof v === 'number') out.push(`<c r="${ref}"><v>${v}</v></c>`);
        else out.push(`<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(String(v).slice(0, 32767))}</t></is></c>`);
      });
      out.push('</row>');
    });
    out.push('</sheetData></worksheet>');
    return out.join('');
  }

  // sheets: [{ name, rows: (string|number)[][] }]  ->  Uint8Array
  async function writeXlsx(sheets) {
    const E = (n, s) => ({ name: n, data: enc.encode(s) });
    const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    return zip([
      E('[Content_Types].xml', `${head}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`),
      E('_rels/.rels', `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
      E('xl/workbook.xml', `${head}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheets>${sheets.map((s, i) => `<sheet name="${escapeXml(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`),
      E('xl/_rels/workbook.xml.rels', `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length + 1}" Type="${R}/styles" Target="styles.xml"/></Relationships>`),
      E('xl/styles.xml', `${head}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`),
      ...sheets.map((s, i) => E(`xl/worksheets/sheet${i + 1}.xml`, sheetXml(s.rows))),
    ]);
  }

  const api = { readXlsx, writeXlsx };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else globalThis.XlsxLite = api;
})();
