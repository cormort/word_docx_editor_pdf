/* docx-fidelity.js — 保版面模式的 .docx → HTML
 * 目的：mammoth 只保留「內容」，這一支補上「頁面層級」在瀏覽器裡做得到的部分：
 *   紙張大小與邊界（sectPr）、頁首／頁尾（列印時每頁重複）、分頁符號、
 *   字型與字級、粗斜底線、顏色與螢光、表格欄寬與合併儲存格、清單層級與樣式。
 * 誠實的限制：Word 的換行位置（沒有排版引擎）、圖片浮動文繞圖、欄位自動更新做不到；
 *   頁尾的 PAGE 欄位會顯示 Word 上次存檔時算出的數字。
 * 自帶 ZIP 讀取（DecompressionStream），不依賴任何外部套件。
 */
(function () {
  'use strict';

  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const TWIP_MM = 25.4 / 1440;          // twips → mm
  const EMU_PT = 12700;                 // EMU → pt
  const esc = (t) => String(t == null ? '' : t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const tags = (el, name) => el ? [...el.children].filter((c) => c.localName === name) : [];
  const tag = (el, name) => tags(el, name)[0] || null;
  const val = (el, name) => { if (!el) return null; return el.getAttributeNS(W, name) || el.getAttribute('w:' + name) || el.getAttribute(name); };
  const rval = (el, name) => { if (!el) return null; return el.getAttributeNS(R, name) || el.getAttribute('r:' + name) || el.getAttribute(name); };
  const num = (v, d) => { const n = parseInt(v, 10); return isFinite(n) ? n : d; };

  // ---------- ZIP ----------
  async function unzip(buffer) {
    const u8 = new Uint8Array(buffer);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 66000); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('不是有效的 ZIP／docx 檔');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const out = new Map();
    for (let i = 0; i < count; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const usize = dv.getUint32(p + 24, true);
      const nlen = dv.getUint16(p + 28, true);
      const elen = dv.getUint16(p + 30, true);
      const clen = dv.getUint16(p + 32, true);
      const lho = dv.getUint32(p + 42, true);
      const name = new TextDecoder().decode(u8.subarray(p + 46, p + 46 + nlen));
      const lnlen = dv.getUint16(lho + 26, true);
      const lelen = dv.getUint16(lho + 28, true);
      const start = lho + 30 + lnlen + lelen;
      out.set(name, { method, raw: u8.subarray(start, start + (csize || usize)) });
      p += 46 + nlen + elen + clen;
    }
    return out;
  }
  async function entryBytes(zip, name) {
    const e = zip.get(name);
    if (!e) return null;
    if (e.method === 0) return e.raw;
    const stream = new Blob([e.raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }
  async function entryText(zip, name) {
    const b = await entryBytes(zip, name);
    return b ? new TextDecoder().decode(b) : null;
  }
  const xml = (text) => new DOMParser().parseFromString(text || '<x/>', 'application/xml');

  // ---------- styles.xml ----------
  function parseStyles(text) {
    const map = new Map();
    if (!text) return { map, defR: null, defP: null };
    const doc = xml(text);
    for (const st of doc.getElementsByTagName('*')) {
      if (st.localName !== 'style') continue;
      const id = val(st, 'styleId');
      if (!id) continue;
      map.set(id, { name: val(tag(st, 'name'), 'val') || '', rPr: tag(st, 'rPr'), pPr: tag(st, 'pPr'), type: val(st, 'type') });
    }
    const dd = [...doc.getElementsByTagName('*')].find((e) => e.localName === 'docDefaults');
    return {
      map,
      defR: dd ? tag(tag(dd, 'rPrDefault'), 'rPr') : null,
      defP: dd ? tag(tag(dd, 'pPrDefault'), 'pPr') : null,
    };
  }

  // 樣式名稱 → HTML 標籤（與 index.html 的 mammoth styleMap 一致，含中文樣式名）
  const STYLE_TAG = [
    [/^(title|標題)$/i, ['h1', 'title']],
    [/^(subtitle|副標題)$/i, ['p', 'subtitle']],
    [/^(heading|標題)\s*([1-9])$/i, null],       // 動態判斷
    [/^(quote|intense quote|引言|引文)$/i, ['blockquote', 'quote']],
    [/^(caption|圖表標題|圖說)$/i, ['p', 'caption']],
  ];
  function styleToTag(name) {
    if (!name) return null;
    const n = String(name).trim();
    const h = /^(heading|標題)\s*([1-9])$/i.exec(n);
    if (h) return ['h' + Math.min(4, +h[2]), 'h' + Math.min(4, +h[2])];
    for (const [re, out] of STYLE_TAG) if (re.test(n)) return out;
    return null;
  }

  // ---------- numbering.xml ----------
  function parseNumbering(text) {
    const abs = new Map(), nums = new Map();
    if (!text) return { abs, nums };
    const doc = xml(text);
    for (const a of doc.getElementsByTagName('*')) {
      if (a.localName !== 'abstractNum') continue;
      const lv = [];
      for (const l of tags(a, 'lvl')) {
        const i = num(val(l, 'ilvl'), 0);
        const pPr = tag(l, 'pPr'), ind = pPr ? tag(pPr, 'ind') : null;
        lv[i] = {
          fmt: val(tag(l, 'numFmt'), 'val') || 'bullet',
          text: val(tag(l, 'lvlText'), 'val') || '',
          left: num(val(ind, 'left'), 720 * (i + 1)),
          hanging: num(val(ind, 'hanging'), 360),
        };
      }
      abs.set(val(a, 'abstractNumId'), lv);
    }
    for (const n of doc.getElementsByTagName('*')) {
      if (n.localName === 'num') nums.set(val(n, 'numId'), val(tag(n, 'abstractNumId'), 'val'));
    }
    return { abs, nums };
  }
  const LIST_STYLE = { bullet: 'disc', decimal: 'decimal', lowerLetter: 'lower-alpha', upperLetter: 'upper-alpha', lowerRoman: 'lower-roman', upperRoman: 'upper-roman', none: 'none' };

  // ---------- rels ----------
  function parseRels(text) {
    const map = new Map();
    if (!text) return map;
    const doc = xml(text);
    for (const r of doc.getElementsByTagName('*')) {
      if (r.localName !== 'Relationship') continue;
      map.set(r.getAttribute('Id'), { type: r.getAttribute('Type') || '', target: r.getAttribute('Target') || '', mode: r.getAttribute('TargetMode') || '' });
    }
    return map;
  }

  // ---------- 執行（run）→ HTML ----------
  function runStyle(rPr, ctx) {
    const st = { css: [], b: false, i: false, u: false };
    if (!rPr) return st;
    if (tag(rPr, 'b') && val(tag(rPr, 'b'), 'val') !== '0' && val(tag(rPr, 'b'), 'val') !== 'false') st.b = true;
    if (tag(rPr, 'i') && val(tag(rPr, 'i'), 'val') !== '0' && val(tag(rPr, 'i'), 'val') !== 'false') st.i = true;
    if (tag(rPr, 'u') && (val(tag(rPr, 'u'), 'val') || 'single') !== 'none') st.u = true;
    const color = val(tag(rPr, 'color'), 'val');
    if (color && /^[0-9a-f]{6}$/i.test(color) && color.toLowerCase() !== '000000' && color.toLowerCase() !== 'auto') st.css.push('color:#' + color);
    const hl = val(tag(rPr, 'highlight'), 'val');
    const HL = { yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000000' };
    const shd = tag(rPr, 'shd');
    const fill = shd ? val(shd, 'fill') : null;
    if (hl && HL[hl]) st.css.push('background-color:' + HL[hl]);
    else if (fill && /^[0-9a-f]{6}$/i.test(fill) && fill.toLowerCase() !== 'auto' && fill.toUpperCase() !== 'FFFFFF') st.css.push('background-color:#' + fill);
    const fonts = tag(rPr, 'rFonts');
    if (fonts) {
      const ea = val(fonts, 'eastAsia'), as = val(fonts, 'ascii') || val(fonts, 'hAnsi');
      const fam = [ea, as].filter(Boolean);
      if (fam.length) st.css.push('font-family:' + fam.map((f) => (/[\s]/.test(f) ? '"' + f + '"' : f)).join(','));
    }
    const sz = num(val(tag(rPr, 'sz'), 'val'), 0);
    if (sz) st.css.push('font-size:' + (sz / 2) + 'pt');
    const va = val(tag(rPr, 'vertAlign'), 'val');
    if (va === 'superscript') st.css.push('vertical-align:super;font-size:.75em');
    if (va === 'subscript') st.css.push('vertical-align:sub;font-size:.75em');
    if (tag(rPr, 'rStyle')) {
      const s = ctx.styles.map.get(val(tag(rPr, 'rStyle'), 'styleId'));
      if (s) {
        const nm = (s.name || '').toLowerCase();
        if (nm === 'strong') st.b = true;
        if (nm === 'emphasis') st.i = true;
      }
    }
    return st;
  }
  function wrapRun(text, st) {
    let out = esc(text).replace(/\n/g, '<br>');
    if (!out) return '';
    if (st.u) out = '<u>' + out + '</u>';
    if (st.i) out = '<i>' + out + '</i>';
    if (st.b) out = '<b>' + out + '</b>';
    if (st.css.length) out = '<span style="' + st.css.join(';') + '">' + out + '</span>';
    return out;
  }

  // ---------- 段落 ----------
  function paraAlign(pPr) {
    const jc = val(tag(pPr, 'jc'), 'val');
    if (jc === 'center') return 'center';
    if (jc === 'right' || jc === 'end') return 'right';
    if (jc === 'both' || jc === 'distribute') return 'justify';
    return '';
  }
  function paraIndent(pPr) {
    const ind = pPr ? tag(pPr, 'ind') : null;
    if (!ind) return '';
    const left = num(val(ind, 'left'), 0) || num(val(ind, 'start'), 0);
    return left > 0 ? 'margin-left:' + (left * TWIP_MM).toFixed(1) + 'mm' : '';
  }
  function docxParagraph(p, ctx) {
    const pPr = tag(p, 'pPr');
    const styleId = val(tag(pPr, 'pStyle'), 'val');
    const style = styleId ? ctx.styles.map.get(styleId) : null;
    const styleName = style ? style.name : '';
    // 段落樣式自帶的字型／大小（例如 Heading 1 的粗體與字級）
    let base = { css: [], b: false, i: false, u: false };
    if (style && style.rPr) base = runStyle(style.rPr, ctx);
    let mapped = styleToTag(styleName) || (styleId && /^Heading([1-9])$/i.test(styleId) ? ['h' + Math.min(4, +styleId.replace(/\D/g, '')), ''] : null);
    const htmlTag = mapped ? mapped[0] : 'p';
    const cls = mapped && mapped[1] ? ' class="' + mapped[1] + '"' : '';
    const styles = [];
    const align = paraAlign(pPr);
    if (align) styles.push('text-align:' + align);
    const indent = paraIndent(pPr);
    if (indent) styles.push(indent);
    const shd = pPr ? tag(pPr, 'shd') : null;
    const fill = shd ? val(shd, 'fill') : null;
    if (fill && /^[0-9a-f]{6}$/i.test(fill) && fill.toUpperCase() !== 'FFFFFF') styles.push('background-color:#' + fill);
    if (base.css.length) styles.push(...base.css);

    let inner = '';
    const parts = [];
    const emit = (html) => { if (html) parts.push(html); };
    const walk = (node, rPr) => {
      for (const ch of node.children) {
        if (ch.localName === 'r') {
          const st = runStyle(tag(ch, 'rPr') || rPr || (style ? style.rPr : null) || ctx.styles.defR, ctx);
          if (base.b) st.b = true;
          if (base.i) st.i = true;
          if (base.u) st.u = true;
          for (const c of ch.children) {
            if (c.localName === 't') emit(wrapRun(c.textContent, st));
            else if (c.localName === 'tab') emit('<span style="display:inline-block;width:2em"></span>');
            else if (c.localName === 'br') {
              const type = val(c, 'type');
              if (type === 'page') ctx.pageBreaks.push(parts.length);
              emit(type === 'page' ? '</p><div class="pagebreak" contenteditable="false"></div><p>' : '<br>');
            } else if (c.localName === 'drawing' || c.localName === 'pict') emit(docxImage(c, ctx));
            else if (c.localName === 'sym') emit(esc(String.fromCharCode(num(val(c, 'char'), 0))));
            else if (c.localName === 'fldSimple' || c.localName === 'instrText') emit(esc(c.textContent));
          }
        } else if (ch.localName === 'hyperlink') {
          const rel = ctx.rels.get(rval(ch, 'id'));
          const inner2 = [];
          const sub = { children: ch.children };
          const collect = (n) => { for (const c of n.children) { if (c.localName === 'r') {
              const st = runStyle(tag(c, 'rPr') || (style ? style.rPr : null) || ctx.styles.defR, ctx);
              for (const cc of c.children) if (cc.localName === 't') inner2.push(wrapRun(cc.textContent, st));
            } else if (c.localName === 'hyperlink' || c.localName === 'smartTag') collect(c); } };
          collect(ch);
          const txt = inner2.join('');
          if (rel && /^https?:/i.test(rel.target)) emit('<a href="' + esc(rel.target) + '">' + txt + '</a>');
          else emit(txt);
        } else if (ch.localName === 'fldSimple') {
          // 頁碼之類的欄位：Word 會把上次算出的數字存在 run 裡，直接沿用
          walk(ch, rPr);
        } else if (ch.localName === 'smartTag' || ch.localName === 'sdt' || ch.localName === 'ins') {
          walk(ch, rPr);
        } else if (ch.localName === 'del') { /* 追蹤修訂的刪除內容：不輸出 */ }
      }
    };
    walk(p, null);
    inner = parts.join('');
    if (!inner.trim()) inner = '';
    return { html: '<' + htmlTag + cls + (styles.length ? ' style="' + styles.join(';') + '"' : '') + '>' + inner + '</' + htmlTag + '>',
      pPr, numId: val(tag(pPr, 'numPr') ? tag(tag(pPr, 'numPr'), 'numId') : null, 'val'),
      ilvl: num(val(tag(pPr, 'numPr') ? tag(tag(pPr, 'numPr'), 'ilvl') : null, 'val'), 0),
      empty: !inner.trim() };
  }

  // ---------- 圖片 ----------
  function docxImage(node, ctx) {
    const blip = [...node.getElementsByTagName('*')].find((e) => e.localName === 'blip');
    const embed = blip ? (blip.getAttributeNS(R, 'embed') || blip.getAttribute('r:embed')) : null;
    if (!embed) return '';
    const rel = ctx.rels.get(embed);
    if (!rel) return '';
    const path = 'word/' + rel.target.replace(/^\/?word\//, '').replace(/^\.\//, '');
    const bytes = ctx.media.get(path) || ctx.media.get(rel.target) || ctx.media.get('word/' + rel.target);
    if (!bytes) return '';
    const ext = (path.split('.').pop() || 'png').toLowerCase();
    const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : ext === 'gif' ? 'image/gif' : ext === 'webp' ? 'image/webp' : 'image/png';
    let b64 = '';
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) b64 += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    const dataUrl = 'data:' + mime + ';base64,' + btoa(b64);
    const ext2 = [...node.getElementsByTagName('*')].find((e) => e.localName === 'extent');
    let style = 'max-width:100%';
    if (ext2) {
      const cx = num(ext2.getAttribute('cx'), 0);
      if (cx) style += ';width:' + (cx / EMU_PT).toFixed(1) + 'pt';
    }
    return '<img src="' + dataUrl + '" style="' + style + '" alt="">';
  }

  // ---------- 表格 ----------
  function docxTable(tbl, ctx) {
    const grid = tag(tbl, 'tblGrid');
    const cols = grid ? tags(grid, 'gridCol').map((c) => num(val(c, 'w'), 0)) : [];
    const totalW = cols.reduce((a, b) => a + b, 0) || 0;
    const rows = tags(tbl, 'tr');
    const matrix = [];
    rows.forEach((tr, r) => {
      matrix[r] = matrix[r] || [];
      let c = 0;
      for (const tc of tags(tr, 'tc')) {
        while (matrix[r][c]) c++;
        const tcPr = tag(tc, 'tcPr');
        const span = num(val(tag(tcPr, 'gridSpan'), 'val'), 1);
        const vm = tag(tcPr, 'vMerge');
        const restart = vm && (val(vm, 'val') || 'continue') === 'restart';
        const cont = vm && !restart;
        for (let i = 0; i < 1; i++) for (let j = 0; j < span; j++) {
          matrix[r][c + j] = { tc, start: j === 0, cont, restart };
        }
        c += span;
      }
    });
    // 第二輪：矩陣建好之後才算縱向合併的 rowspan（往下數續格）
    const rowspans = new Map();
    rows.forEach((tr, r) => {
      const width = matrix[r] ? matrix[r].length : 0;
      for (let c = 0; c < width; c++) {
        const m = matrix[r][c];
        if (!m || !m.restart || !m.start) continue;
        let count = 1, rr = r + 1;
        while (matrix[rr] && matrix[rr][c] && matrix[rr][c].cont) { count++; rr++; }
        rowspans.set(r + ':' + c, count);
      }
    });
    let body = '';
    rows.forEach((tr, r) => {
      let cells = '';
      const width = matrix[r] ? matrix[r].length : 0;
      for (let c = 0; c < width; c++) {
        const m = matrix[r][c];
        if (!m) { cells += '<td></td>'; continue; }
        if (!m.start) continue;                       // 橫向被 gridSpan 蓋掉
        if (m.cont) continue;                          // 縱向續格：由上面那格的 rowspan 代表，不輸出
        const tcPr = tag(m.tc, 'tcPr');
        const span = num(val(tag(tcPr, 'gridSpan'), 'val'), 1);
        const rs = rowspans.get(r + ':' + c) || 1;
        const attrs = (span > 1 ? ' colspan="' + span + '"' : '') + (rs > 1 ? ' rowspan="' + rs + '"' : '');
        const isHead = val(tcPr && tag(tcPr, 'tcW'), 'w') && false;
        let inner = '';
        for (const p of tags(m.tc, 'p')) inner += docxParagraph(p, ctx).html;
        if (!inner) inner = '<p><br></p>';
        cells += '<td' + attrs + '>' + inner + '</td>';
      }
      body += '<tr>' + cells + '</tr>';
    });
    let colgroup = '';
    if (cols.length && totalW) colgroup = '<colgroup>' + cols.map((w) => '<col style="width:' + ((w / totalW) * 100).toFixed(2) + '%">').join('') + '</colgroup>';
    return '<table>' + colgroup + '<tbody>' + body + '</tbody></table>';
  }

  // ---------- 頁首／頁尾 ----------
  function convertEdge(text, ctx) {
    if (!text) return null;
    const doc = xml(text);
    const root = [...doc.getElementsByTagName('*')].find((e) => e.localName === 'hdr' || e.localName === 'ftr');
    if (!root) return null;
    let out = '';
    for (const p of tags(root, 'p')) {
      const { html } = docxParagraph(p, ctx);
      out += html;
    }
    return /[^\s]/.test(out.replace(/<[^>]+>/g, '')) ? out : null;
  }

  // ---------- 主流程 ----------
  async function convert(arrayBuffer) {
    const zip = await unzip(arrayBuffer);
    const documentXml = await entryText(zip, 'word/document.xml');
    if (!documentXml) throw new Error('docx 裡找不到 word/document.xml');
    const ctx = {
      styles: parseStyles(await entryText(zip, 'word/styles.xml')),
      numbering: parseNumbering(await entryText(zip, 'word/numbering.xml')),
      rels: parseRels(await entryText(zip, 'word/_rels/document.xml.rels')),
      media: new Map(),
      pageBreaks: [],
      warnings: [],
    };
    for (const [name, entry] of zip) {
      if (/^word\/media\//i.test(name)) ctx.media.set(name, entry.method === 0 ? entry.raw : await entryBytes(zip, name));
    }
    const doc = xml(documentXml);
    const body = [...doc.getElementsByTagName('*')].find((e) => e.localName === 'body');
    if (!body) throw new Error('docx 內容格式不認識');

    const html = [];
    let listCtx = null;                                  // {numId, ilvl, tag}
    let listCount = 0;
    const closeList = (level) => { while (listCtx && listCtx.ilvl >= level) { html.push(listCtx.tag === 'ol' ? '</ol>' : '</ul>'); listCtx = listCtx.parent || null; } };
    const openList = (numId, ilvl) => {
      const absId = ctx.numbering.nums.get(numId);
      const lv = absId != null ? (ctx.numbering.abs.get(absId) || [])[ilvl] : null;
      const fmt = lv ? lv.fmt : 'bullet';
      const tagName = fmt === 'bullet' || fmt === 'none' ? 'ul' : 'ol';
      const style = 'list-style-type:' + (LIST_STYLE[fmt] || 'disc') + (lv ? ';padding-left:' + Math.max(6, (lv.left * TWIP_MM)).toFixed(1) + 'mm' : '');
      html.push('<' + tagName + ' style="' + style + '">');
      listCtx = { numId, ilvl, tag: tagName, parent: listCtx };
      listCount++;
    };

    for (const node of body.children) {
      if (node.localName === 'p') {
        const r = docxParagraph(node, ctx);
        if (r.numId != null) {
          const ilvl = r.ilvl;
          while (listCtx && listCtx.ilvl > ilvl) closeList(listCtx.ilvl);
          if (!listCtx || listCtx.numId !== r.numId || listCtx.ilvl < ilvl) openList(r.numId, ilvl);
          html.push('<li' + (r.html.match(/style="([^"]*)"/) ? ' style="' + r.html.match(/style="([^"]*)"/)[1] + '"' : '') + '>' + r.html.replace(/^<[^>]+>|<\/[^>]+>$/g, '') + '</li>');
        } else {
          closeList(0);
          html.push(r.html || '<p><br></p>');
        }
      } else if (node.localName === 'tbl') {
        closeList(0);
        html.push(docxTable(node, ctx));
      } else if (node.localName === 'sectPr') {
        ctx.sectPr = node;
      }
    }
    closeList(0);

    // 紙張與邊界
    const sect = ctx.sectPr || [...doc.getElementsByTagName('*')].find((e) => e.localName === 'sectPr');
    let page = null, header = null, footer = null;
    if (sect) {
      const sz = tag(sect, 'pgSz');
      const mar = tag(sect, 'pgMar');
      if (sz) {
        const wmm = num(val(sz, 'w'), 11906) * TWIP_MM, hmm = num(val(sz, 'h'), 16838) * TWIP_MM;
        page = {
          w: +wmm.toFixed(1), h: +hmm.toFixed(1),
          top: mar ? +(num(val(mar, 'top'), 1134) * TWIP_MM).toFixed(1) : 20,
          right: mar ? +(num(val(mar, 'right'), 1134) * TWIP_MM).toFixed(1) : 20,
          bottom: mar ? +(num(val(mar, 'bottom'), 1134) * TWIP_MM).toFixed(1) : 20,
          left: mar ? +(num(val(mar, 'left'), 1134) * TWIP_MM).toFixed(1) : 20,
        };
      }
      for (const ref of [...sect.children]) {
        const type = val(ref, 'type');
        if (ref.localName !== 'headerReference' && ref.localName !== 'footerReference') continue;
        const rel = ctx.rels.get(rval(ref, 'id'));
        if (!rel) continue;
        const part = 'word/' + rel.target.replace(/^\/?word\//, '');
        const text = await entryText(zip, part);
        if (ref.localName === 'headerReference' && (type === 'default' || !header)) header = await convertEdge(text, ctx);
        if (ref.localName === 'footerReference' && (type === 'default' || !footer)) footer = await convertEdge(text, ctx);
      }
      if (header && /\bPAGE\b/.test(header)) ctx.warnings.push('頁尾的頁碼欄位會顯示 Word 上次存檔的數字');
    }
    return { html: html.join('\n'), page, header, footer, warnings: ctx.warnings, lists: listCount, breaks: (html.join('').match(/class="pagebreak"/g) || []).length };
  }

  window.docxFidelity = { convert, unzip };
})();
