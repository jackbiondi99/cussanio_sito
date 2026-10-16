/* =====================================================================
   Lettura e scrittura di file Excel (.xlsx) nel browser, senza librerie
   ===================================================================== */

/* ---------------------------- LETTURA ---------------------------- */
async function unzip(buf){
  const dv = new DataView(buf), dec = new TextDecoder();
  let eocd = -1;
  for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 65557); i--)
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('Il file non sembra un Excel valido (.xlsx).');
  const count = dv.getUint16(eocd + 10, true), cdOff = dv.getUint32(eocd + 16, true);
  const files = {};
  let p = cdOff;
  for (let k = 0; k < count; k++){
    const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
    const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
    const lo = dv.getUint32(p + 42, true);
    files[dec.decode(new Uint8Array(buf, p + 46, nl))] = { method, csize, lo };
    p += 46 + nl + el + cl;
  }
  return async name => {
    const f = files[name]; if (!f) return null;
    const start = f.lo + 30 + dv.getUint16(f.lo + 26, true) + dv.getUint16(f.lo + 28, true);
    const data = new Uint8Array(buf, start, f.csize);
    if (f.method === 0) return dec.decode(data);
    const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return dec.decode(await new Response(stream).arrayBuffer());
  };
}
const parseXml = s => new DOMParser().parseFromString(s, 'application/xml');
const tags = (n, t) => Array.from(n.getElementsByTagNameNS('*', t));
const attrLocal = (el, name) => { for (const a of el.attributes) if (a.localName === name) return a.value; return null; };
const colIndex = ref => { let n = 0; for (const ch of ref.replace(/[0-9]/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };

function sheetToObjects(doc, shared){
  const grid = [];
  for (const r of tags(doc, 'row')){
    const arr = []; let auto = 0;
    for (const c of tags(r, 'c')){
      const ref = c.getAttribute('r'), t = c.getAttribute('t'), v = tags(c, 'v')[0];
      const idx = ref ? colIndex(ref) : auto; auto = idx + 1;
      let val = null;
      if (t === 's') val = v ? shared[+v.textContent] : null;
      else if (t === 'inlineStr') val = tags(c, 't').map(x => x.textContent).join('');
      else if (t === 'str') val = v ? v.textContent : null;
      else if (t === 'b') val = v ? v.textContent === '1' : null;
      else if (t === 'e') val = null;
      else val = v && v.textContent !== '' ? parseFloat(v.textContent) : null;
      if (typeof val === 'string') { val = val.trim(); if (val === '') val = null; }
      arr[idx] = val;
    }
    grid[(+r.getAttribute('r') || grid.length + 1) - 1] = arr;
  }
  const hi = grid.findIndex(r => r && r.some(x => x != null));
  if (hi < 0) return [];
  const head = Array.from(grid[hi], h => h == null ? null : String(h).trim());
  const out = [];
  for (let i = hi + 1; i < grid.length; i++){
    const r = grid[i]; if (!r || !r.some(x => x != null)) continue;
    const o = {}; head.forEach((h, j) => { if (h) o[h] = r[j] ?? null; }); out.push(o);
  }
  return out;
}

export async function readWorkbook(buf){
  const get = await unzip(buf);
  const wb = parseXml(await get('xl/workbook.xml'));
  const rels = parseXml(await get('xl/_rels/workbook.xml.rels'));
  const relMap = {};
  tags(rels, 'Relationship').forEach(r => relMap[r.getAttribute('Id')] = r.getAttribute('Target'));
  const sst = await get('xl/sharedStrings.xml');
  const shared = sst ? tags(parseXml(sst), 'si').map(si =>
      tags(si, 't').filter(t => t.parentNode.localName !== 'rPh').map(t => t.textContent).join('')) : [];
  const out = {};
  for (const s of tags(wb, 'sheet')){
    let target = relMap[attrLocal(s, 'id')]; if (!target) continue;
    target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
    const txt = await get(target);
    if (txt) out[s.getAttribute('name')] = sheetToObjects(parseXml(txt), shared);
  }
  return out;
}

/* ---------------------------- SCRITTURA ---------------------------- */
const CRC_TABLE = (() => { const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++){ let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
  return t; })();
const crc32 = b => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC_TABLE[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };

function zipStore(files){            // files: [{name, text}]
  const enc = new TextEncoder(), parts = [], central = [];
  let off = 0;
  for (const f of files){
    const name = enc.encode(f.name), data = enc.encode(f.text), crc = crc32(data);
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
    lh.setUint16(8, 0, true); lh.setUint16(10, 0, true); lh.setUint16(12, 0x21, true);
    lh.setUint32(14, crc, true); lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true);
    lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
    parts.push(new Uint8Array(lh.buffer), name, data);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
    ch.setUint16(10, 0, true); ch.setUint16(12, 0, true); ch.setUint16(14, 0x21, true);
    ch.setUint32(16, crc, true); ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true);
    ch.setUint16(28, name.length, true); ch.setUint32(42, off, true);
    central.push(new Uint8Array(ch.buffer), name);
    off += 30 + name.length + data.length;
  }
  const cdSize = central.reduce((s, p) => s + p.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
  end.setUint32(12, cdSize, true); end.setUint32(16, off, true);
  return new Blob([...parts, ...central, new Uint8Array(end.buffer)],
    { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

const x = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]))
                        .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
const colL = i => { let s = ''; i++; while (i){ const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = (i - m - 1) / 26; } return s; };
const isoToSerial = iso => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || ''); return m ? Date.UTC(+m[1], m[2] - 1, +m[3]) / 864e5 + 25569 : null; };

/* sheets: [{ name, table?, columns:[{key, width?, type?:'date'|'num'|'text', list?:string[]|string }], rows?:[{}], text?:[[...]] }] */
export function writeWorkbook(sheets){
  const files = [], ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
        rns = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  let tableN = 0;
  const ctOverrides = [];

  sheets.forEach((sh, si) => {
    const idx = si + 1, rowsXml = [];
    let extra = '', rels = '', cols = '', dvs = '', ref = 'A1';
    if (sh.text){                                   // foglio di sole istruzioni
      sh.text.forEach((line, r) => {
        const st = line[1] || 0;
        rowsXml.push(`<row r="${r + 1}"><c r="A${r + 1}" t="inlineStr" s="${st}"><is><t xml:space="preserve">${x(line[0] ?? '')}</t></is></c></row>`);
      });
      cols = `<cols><col min="1" max="1" width="110" customWidth="1"/></cols>`;
    } else {
      const C = sh.columns, rows = sh.rows || [];
      rowsXml.push(`<row r="1">${C.map((c, j) => `<c r="${colL(j)}1" t="inlineStr" s="1"><is><t>${x(c.key)}</t></is></c>`).join('')}</row>`);
      rows.forEach((row, i) => {
        const r = i + 2;
        const cells = C.map((c, j) => {
          let v = row[c.key]; const a = `${colL(j)}${r}`;
          if (v == null || v === '') return '';
          if (c.type === 'date'){ const n = isoToSerial(v); return n == null ? `<c r="${a}" t="inlineStr"><is><t>${x(v)}</t></is></c>` : `<c r="${a}" s="2"><v>${n}</v></c>`; }
          if (typeof v === 'number' && isFinite(v)) return `<c r="${a}"><v>${v}</v></c>`;
          return `<c r="${a}" t="inlineStr"><is><t xml:space="preserve">${x(v)}</t></is></c>`;
        }).join('');
        rowsXml.push(`<row r="${r}">${cells}</row>`);
      });
      const last = Math.max(2, rows.length + 1);
      ref = `A1:${colL(C.length - 1)}${last}`;
      cols = `<cols>${C.map((c, j) => `<col min="${j + 1}" max="${j + 1}" width="${c.width || 14}" customWidth="1"${c.type === 'date' ? ' style="2"' : ''}/>`).join('')}</cols>`;
      const lists = C.map((c, j) => [c, j]).filter(([c]) => c.list);
      if (lists.length){
        dvs = `<dataValidations count="${lists.length}">${lists.map(([c, j]) => {
          const f = Array.isArray(c.list) ? `"${c.list.join(',')}"` : c.list;
          return `<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="${colL(j)}2:${colL(j)}2000"><formula1>${x(f)}</formula1></dataValidation>`;
        }).join('')}</dataValidations>`;
      }
      if (sh.table){
        tableN++;
        files.push({ name: `xl/tables/table${tableN}.xml`, text:
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table ${ns} id="${tableN}" name="${sh.table}" displayName="${sh.table}" ref="${ref}" totalsRowShown="0">` +
          `<autoFilter ref="${ref}"/><tableColumns count="${C.length}">${C.map((c, j) => `<tableColumn id="${j + 1}" name="${x(c.key)}"/>`).join('')}</tableColumns>` +
          `<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>` });
        ctOverrides.push(`<Override PartName="/xl/tables/table${tableN}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`);
        extra = `<tableParts count="1"><tablePart r:id="rId1"/></tableParts>`;
        rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table${tableN}.xml"/></Relationships>`;
      }
    }
    const pane = sh.text ? '' : `<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>`;
    files.push({ name: `xl/worksheets/sheet${idx}.xml`, text:
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet ${ns} ${rns}>` +
      `<dimension ref="${ref}"/><sheetViews><sheetView workbookViewId="0"${si === 0 ? ' tabSelected="1"' : ''}>${pane}</sheetView></sheetViews>` +
      `<sheetFormatPr defaultRowHeight="15"/>${cols}<sheetData>${rowsXml.join('')}</sheetData>${dvs}${extra}</worksheet>` });
    if (rels) files.push({ name: `xl/worksheets/_rels/sheet${idx}.xml.rels`, text: rels });
    ctOverrides.push(`<Override PartName="/xl/worksheets/sheet${idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
  });

  files.unshift(
    { name: '[Content_Types].xml', text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${ctOverrides.join('')}</Types>` },
    { name: '_rels/.rels', text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
    { name: 'xl/workbook.xml', text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook ${ns} ${rns}><sheets>${sheets.map((s, i) => `<sheet name="${x(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>` },
    { name: 'xl/_rels/workbook.xml.rels', text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` },
    { name: 'xl/styles.xml', text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet ${ns}>` +
      `<numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/></numFmts>` +
      `<fonts count="3"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font><font><b/><sz val="13"/><color rgb="FF1F3864"/><name val="Arial"/></font></fonts>` +
      `<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill></fills>` +
      `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
      `<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>` +
      `<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
      `<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>` +
      `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>` }
  );
  return zipStore(files);
}
