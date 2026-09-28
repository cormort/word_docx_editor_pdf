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
      map.set(id, { name: val(tag(st, 'name'), 'val') || '', basedOn: val(tag(st, 'basedOn'), 'val'), rPr: tag(st, 'rPr'), pPr: tag(st, 'pPr'), type: val(st, 'type') });
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
      // 字型名稱用單引號：這段 CSS 會放進 style="…"，雙引號會提早結束屬性
      if (fam.length) st.css.push('font-family:' + fam.map((f) => "'" + f.replace(/['";<>&]/g, '') + "'").join(','));
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
  // 段落屬性要沿著樣式鏈找（docDefaults → basedOn … → 樣式 → 段落本身），後者蓋前者
  function pLayers(pPr, styleId, ctx) {
    const out = [];
    for (let id = styleId, n = 0; id && n < 10; n++) { const s = ctx.styles.map.get(id); if (!s) break; if (s.pPr) out.unshift(s.pPr); id = s.basedOn; }
    if (ctx.styles.defP) out.unshift(ctx.styles.defP);
    if (pPr) out.push(pPr);
    return out;
  }
  function pAttr(layers, name, attr) {
    for (let i = layers.length - 1; i >= 0; i--) { const v = val(tag(layers[i], name), attr); if (v != null && v !== '') return v; }
    return null;
  }
  function paraBox(layers) {
    const css = [];
    const tw = (v) => (v / 20).toFixed(1) + 'pt';
    const left = num(pAttr(layers, 'ind', 'left'), num(pAttr(layers, 'ind', 'start'), 0));
    const right = num(pAttr(layers, 'ind', 'right'), num(pAttr(layers, 'ind', 'end'), 0));
    // firstLine 與 hanging 互斥：取最上層有設定的那一個
    const indL = layers.slice().reverse().find((l) => { const i = tag(l, 'ind'); return i && (val(i, 'firstLine') != null || val(i, 'hanging') != null); });
    const first = indL ? num(val(tag(indL, 'ind'), 'firstLine'), 0) : 0, hang = indL ? num(val(tag(indL, 'ind'), 'hanging'), 0) : 0;
    if (left) css.push('margin-left:' + tw(left));
    if (right) css.push('margin-right:' + tw(right));
    if (hang) css.push('text-indent:-' + tw(hang));
    else if (first) css.push('text-indent:' + tw(first));
    const before = pAttr(layers, 'spacing', 'before'), after = pAttr(layers, 'spacing', 'after');
    css.push('margin-top:' + tw(num(before, 0)), 'margin-bottom:' + tw(num(after, 0)));
    const line = num(pAttr(layers, 'spacing', 'line'), 0), rule = pAttr(layers, 'spacing', 'lineRule') || 'auto';
    if (line) css.push('line-height:' + (rule === 'auto' ? (line / 240 * 1.3).toFixed(2) : tw(line)));  // ponytail: auto 行距 ×1.3 近似 Word 的單行高度
    return css;
  }
  function paraAlign(pPr) {
    const jc = val(tag(pPr, 'jc'), 'val');
    if (jc === 'center') return 'center';
    if (jc === 'right' || jc === 'end') return 'right';
    if (jc === 'both' || jc === 'distribute') return 'justify';
    return '';
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
    const layers = pLayers(pPr, styleId, ctx);
    const align = paraAlign(layers.slice().reverse().find((l) => tag(l, 'jc')));
    if (align) styles.push('text-align:' + align);
    styles.push(...paraBox(layers));
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
            } else if (c.localName === 'drawing' || c.localName === 'pict' || c.localName === 'AlternateContent') emit(docxObject(c, ctx));
            else if (c.localName === 'sym') { const code = parseInt(val(c, 'char'), 16); if (code > 0) emit(esc(String.fromCharCode(code))); }  // w:char 是十六進位
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
    const outer = ctx.floats;
    ctx.floats = [];
    walk(p, null);
    // 同一段落錨定的多個浮動物件（圖、圖說文字方塊）：照 Word 的垂直位移由上而下排，整組放在段落後面；
    // 匯入後 index.html 的 placeFloats 再依實際排版把整組放到「錨點段落往下 data-voff pt」的位置
    let after = '';
    if (ctx.floats.length) {
      const fl = ctx.floats.sort((a, b) => a.v - b.v);
      const nw = fl.every((f) => f.noWrap);                  // 全是「不繞排」：Word 裡不會把文字往下擠
      after = '<div data-fgroup data-voff="' + fl[0].v.toFixed(1) + '"' + (nw ? ' data-nowrap' : '') + '>' + fl.map((f) => f.html).join('') + '</div>';
    }
    ctx.floats = outer;
    inner = parts.join('');
    // 段落裡有分頁符號：浮動物件屬於分頁前那一頁，群組放在分頁符號之前
    const pb = inner.indexOf('</p><div class="pagebreak"');
    if (after && pb >= 0) { inner = inner.slice(0, pb + 4) + after + inner.slice(pb + 4); after = ''; }
    if (!inner.trim()) inner = '';
    return { html: '<' + htmlTag + cls + (styles.length ? ' style="' + styles.join(';') + '"' : '') + '>' + inner + '</' + htmlTag + '>', after,
      pPr, numId: val(tag(pPr, 'numPr') ? tag(tag(pPr, 'numPr'), 'numId') : null, 'val'),
      ilvl: num(val(tag(pPr, 'numPr') ? tag(tag(pPr, 'numPr'), 'ilvl') : null, 'val'), 0),
      empty: !inner.trim() };
  }

  // 圖片大小照 docx 的外框（寬高都照，比例跟 Word 一樣）；浮動圖片依水平對齊／位移擺放
  const FLOAT_MAX_PT = 200;   // ponytail: 比這寬的浮動圖改成獨立一行，免得滿版編輯區裡文字繞到旁邊
  function imgBox(node, ctx, isBox) {
    const all = [...node.getElementsByTagName('*')];
    const css = isBox ? ['max-width:100%'] : ['max-width:100%', 'height:auto'];
    let wPt = 0, hPt = 0;
    const ext = all.find((e) => e.localName === 'extent');
    if (ext) { wPt = num(ext.getAttribute('cx'), 0) / EMU_PT; hPt = num(ext.getAttribute('cy'), 0) / EMU_PT; }
    else {                                                        // VML：尺寸寫在 v:shape 的 style 裡
      const sh = all.find((e) => /width:/.test(e.getAttribute('style') || ''));
      const len = (k) => { const m = new RegExp('(?:^|;)\\s*' + k + ':\\s*([\\d.]+)(pt|in|px|mm|cm)?').exec(sh ? sh.getAttribute('style') : ''); if (!m) return 0;
        return +m[1] * ({ in: 72, px: 0.75, mm: 72 / 25.4, cm: 72 / 2.54 }[m[2]] || 1); };
      wPt = len('width'); hPt = len('height');
    }
    if (wPt) css.push('width:' + wPt.toFixed(1) + 'pt');
    if (wPt && hPt && !isBox) css.push('aspect-ratio:' + wPt.toFixed(2) + '/' + hPt.toFixed(2));
    const anchor = all.find((e) => e.localName === 'anchor');
    let block = null;
    if (anchor) {
      const posH = tag(anchor, 'positionH'), wrap = [...anchor.children].find((e) => /^wrap/.test(e.localName));
      const align = posH && tag(posH, 'align') ? tag(posH, 'align').textContent.trim() : '';
      const off = posH && tag(posH, 'posOffset') ? num(tag(posH, 'posOffset').textContent, 0) / EMU_PT : 0;
      // 垂直：只處理相對段落／行的位移（相對頁面的在沒有分頁引擎下無從對應）
      const posV = tag(anchor, 'positionV'), relV = posV ? posV.getAttribute('relativeFrom') : '';
      const voff = posV && tag(posV, 'posOffset') && /^(paragraph|line)$/.test(relV) ? num(tag(posV, 'posOffset').textContent, 0) / EMU_PT : 0;
      // 文字方塊「不繞排」＝浮在文字上：小方塊（標籤、註記）用零高度外框＋絕對定位疊上去；
      // 大方塊通常是整頁版面框，沒有分頁引擎時疊上去會互相覆蓋，改成照寬度與水平對齊留在文字流裡
      const noWrap = !wrap || wrap.localName === 'wrapNone';
      if (isBox && noWrap && wPt && wPt <= FLOAT_MAX_PT && hPt <= FLOAT_MAX_PT) {
        const base = posH && posH.getAttribute('relativeFrom') === 'page' ? -(ctx.marL || 0) : 0;
        const pos = ['position:absolute', 'top:' + voff.toFixed(1) + 'pt'];
        if (align === 'center') pos.push('left:50%', 'transform:translateX(-50%)');
        else if (align === 'right' || align === 'outside') pos.push('right:0');
        else pos.push('left:' + (base + off).toFixed(1) + 'pt');
        return { css: css.concat(pos).join(';'), abs: true };
      }
      const floaty = wrap && /^wrap(Square|Tight|Through)$/.test(wrap.localName) && wPt && wPt <= FLOAT_MAX_PT && align !== 'center';
      if (floaty) css.push('float:' + (align === 'right' || align === 'outside' ? 'right' : 'left'), 'margin:' + Math.max(0, voff).toFixed(1) + 'pt 6pt 4pt ' + (off > 0 && !align ? off.toFixed(1) : '0') + 'pt');
      else {
        // 獨立成行的浮動物件：交給段落收集，依垂直位移排序後放在段落後面（見 docxParagraph）
        block = { v: voff, noWrap: isBox && noWrap };
        css.push('display:block');
        if (align === 'center') css.push('margin-left:auto', 'margin-right:auto');
        else if (align === 'right' || align === 'outside') css.push('margin-left:auto');
        else if (off > 0 && !align) css.push('margin-left:' + off.toFixed(1) + 'pt');
      }
    }
    return { css: css.join(';'), block };
  }

  // ---------- 圖片／文字方塊 ----------
  // 浮動定位做不到：文字方塊內容改成段落間的區塊，圖片照原寬度放在原處
  function docxObject(node, ctx) {
    if (node.localName === 'AlternateContent') {
      // Choice（DrawingML）與 Fallback（VML）是同一物件的兩種寫法，只取一個
      const branch = tag(node, 'Choice') || tag(node, 'Fallback');
      return branch ? [...branch.children].map((c) => docxObject(c, ctx)).join('') : '';
    }
    const all = [...node.getElementsByTagName('*')];
    const inTxbx = (e) => { for (let x = e.parentNode; x && x !== node; x = x.parentNode) if (x.localName === 'txbxContent') return true; return false; };
    let out = '';
    for (const e of all) {
      if ((e.localName !== 'blip' && e.localName !== 'imagedata') || inTxbx(e)) continue;
      const img = docxImage(e, node, ctx), box = imgBox(node, ctx, false);
      if (img && box.block && ctx.floats) ctx.floats.push({ v: box.block.v, html: '<p>' + img + '</p>' });
      else out += img;
    }
    const boxes = all.filter((e) => e.localName === 'txbxContent' && !inTxbx(e));
    if (boxes.length) {
      let inner = '';
      for (const b of boxes) for (const ch of b.children) {
        if (ch.localName === 'p') { const r = docxParagraph(ch, ctx); inner += r.html + r.after; }
        else if (ch.localName === 'tbl') inner += docxTable(ch, ctx);
      }
      if (inner.replace(/<[^>]+>/g, '').trim()) {
        const b = imgBox(node, ctx, true);
        const div = '<div class="docx-txbx" style="' + b.css + '">' + inner + '</div>';
        if (b.block && ctx.floats) ctx.floats.push({ v: b.block.v, html: div, noWrap: b.block.noWrap });
        else out += '</p>' + (b.abs ? '<div class="docx-txbx-anchor" style="position:relative;height:0">' + div + '</div>' : div) + '<p>';
      }
    }
    return out;
  }
  function docxImage(blip, node, ctx) {
    const embed = blip.getAttributeNS(R, 'embed') || blip.getAttribute('r:embed') || blip.getAttributeNS(R, 'id') || blip.getAttribute('r:id');
    if (!embed) return '';
    const rel = ctx.rels.get(embed);
    if (!rel) return '';
    const path = 'word/' + rel.target.replace(/^\/?word\//, '').replace(/^\.\//, '');
    const ext = (path.split('.').pop() || 'png').toLowerCase().replace(/^emz$/, 'emf').replace(/^wmz$/, 'wmf');
    const bytes = ctx.media.get(path) || ctx.media.get(rel.target) || ctx.media.get('word/' + rel.target);
    let dataUrl = '';
    if (ext === 'emf' && bytes) {
      try { dataUrl = emfToPng(bytes); } catch (e) { dataUrl = ''; }
    }
    if (ext === 'wmf' && bytes) {
      try { dataUrl = wmfToPng(bytes); } catch (e) { dataUrl = ''; }
    }
    if (/^(emf|wmf|emz|wmz)$/.test(ext) && !dataUrl) {
      ctx.vectorImages = (ctx.vectorImages || 0) + 1;
      return '<span class="docx-noimg">〔' + ext.toUpperCase() + ' 向量圖：無法轉換，請在 Word 另存成 PNG 後重新插入〕</span>';
    }
    if (!bytes) return '';
    if (!dataUrl) {
      const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : ext === 'gif' ? 'image/gif' : ext === 'webp' ? 'image/webp' : 'image/png';
      let b64 = '';
      const chunk = 0x8000;
      for (let i = 0; i < bytes.length; i += chunk) b64 += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
      dataUrl = 'data:' + mime + ';base64,' + btoa(b64);
    }
    return '<img src="' + dataUrl + '" style="' + imgBox(node, ctx, false).css + '" alt="">';
  }


  // DIB（BITMAPINFO + 像素）→ canvas，EMF 與 WMF 共用；bits 省略時像素緊接在調色盤後面
  function dibCanvas(u8, h, bits) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const i32 = (o) => dv.getInt32(o, true), u32 = (o) => dv.getUint32(o, true);
    const hdr = u32(h);
    if (hdr < 40) return null;                                 // 舊式 BITMAPCOREHEADER：罕見，略過
    const w = i32(h + 4), hh = i32(h + 8), bpp = dv.getUint16(h + 14, true), comp = u32(h + 16);
    const H = Math.abs(hh), up = hh > 0;
    if (comp === 4 || comp === 5 || !w || !H) return null;     // 內嵌 JPEG/PNG：少見，略過
    let nPal = u32(h + 32); if (!nPal && bpp <= 8) nPal = 1 << bpp;
    const pal = h + hdr + (comp === 3 && hdr === 40 ? 12 : 0);
    if (bits == null) bits = pal + (bpp <= 8 ? nPal * 4 : 0);
    const c = document.createElement('canvas'); c.width = w; c.height = H;
    const id = c.getContext('2d').createImageData(w, H), d = id.data;
    const stride = ((w * bpp + 31) >> 5) << 2;
    if (bits + stride * H > u8.length) return null;
    for (let y = 0; y < H; y++) {
      const row = bits + (up ? H - 1 - y : y) * stride;
      for (let x = 0; x < w; x++) {
        const o = (y * w + x) * 4; let b, gg, r;
        if (bpp === 32) { b = u8[row + x * 4]; gg = u8[row + x * 4 + 1]; r = u8[row + x * 4 + 2]; }
        else if (bpp === 24) { b = u8[row + x * 3]; gg = u8[row + x * 3 + 1]; r = u8[row + x * 3 + 2]; }
        else if (bpp === 16) { const v = dv.getUint16(row + x * 2, true); r = ((v >> 10) & 31) << 3; gg = ((v >> 5) & 31) << 3; b = (v & 31) << 3; }
        else { const per = 8 / bpp, byte = u8[row + ((x / per) | 0)], idx = (byte >> (8 - bpp * ((x % per) + 1))) & ((1 << bpp) - 1), q = pal + idx * 4; b = u8[q]; gg = u8[q + 1]; r = u8[q + 2]; }
        d[o] = r; d[o + 1] = gg; d[o + 2] = b; d[o + 3] = 255;
      }
    }
    c.getContext('2d').putImageData(id, 0, 0);
    return c;
  }

  // 網底筆刷（HS_*）→ 8×8 的重複圖樣
  function hatchCanvas(color, style, bg) {
    const c = document.createElement('canvas'); c.width = c.height = 8;
    const g = c.getContext('2d');
    if (bg) { g.fillStyle = bg; g.fillRect(0, 0, 8, 8); }
    g.strokeStyle = color; g.lineWidth = 1; g.beginPath();
    if (style === 0 || style === 4) { g.moveTo(0, 3.5); g.lineTo(8, 3.5); }
    if (style === 1 || style === 4) { g.moveTo(3.5, 0); g.lineTo(3.5, 8); }
    if (style === 2 || style === 5) { g.moveTo(0, 0); g.lineTo(8, 8); }
    if (style === 3 || style === 5) { g.moveTo(8, 0); g.lineTo(0, 8); }
    g.stroke();
    return c;
  }
  // 1 位元的 Bitmap16（WMF 的 CREATEPATTERNBRUSH）→ canvas
  function mono16(u8, o) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const w = dv.getInt16(o + 2, true), h = dv.getInt16(o + 4, true), wb = dv.getInt16(o + 6, true), bpp = u8[o + 9];
    if (bpp !== 1 || w <= 0 || h <= 0) return null;
    const bits = o + 28;
    const c = document.createElement('canvas'); c.width = w; c.height = h;
    const id = c.getContext('2d').createImageData(w, h);
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
      const v = (u8[bits + y * wb + (x >> 3)] >> (7 - (x & 7))) & 1 ? 255 : 0, q = (y * w + x) * 4;
      id.data[q] = id.data[q + 1] = id.data[q + 2] = v; id.data[q + 3] = 255;
    }
    c.getContext('2d').putImageData(id, 0, 0);
    return c;
  }

  // ---------- WMF → PNG ----------
  // 16 位元的舊格式：參數多半倒著放（y 在 x 前、下右上左），物件表用「最小的空位」編號。
  // ponytail: 同 EMF，只做常見紀錄；區域（region）裁切與點陣筆刷略過。
  const CHARSET = { 128: 'shift_jis', 129: 'euc-kr', 134: 'gbk', 136: 'big5', 161: 'windows-1253', 162: 'windows-1254', 177: 'windows-1255', 178: 'windows-1256', 186: 'windows-1257', 204: 'windows-1251', 222: 'windows-874', 238: 'windows-1250' };
  function wmfToPng(u8) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const i16 = (o) => dv.getInt16(o, true), u16 = (o) => dv.getUint16(o, true), u32 = (o) => dv.getUint32(o, true);
    const rgb = (o) => 'rgb(' + u8[o] + ',' + u8[o + 1] + ',' + u8[o + 2] + ')';
    let p = 0, box = null, inch = 1440;
    if (u32(0) === 0x9AC6CDD7) { box = [i16(6), i16(8), i16(10), i16(12)]; inch = u16(14) || 1440; p = 22; }
    if (u16(p) > 2 || u16(p + 2) !== 9) throw new Error('not wmf');
    p += u16(p + 2) * 2;
    const first = p;
    // 沒有 placeable 標頭就用最後一次設定的視窗範圍當畫布
    if (!box) {
      let org = [0, 0], ext = null;
      for (let q = first; q + 6 <= u8.length;) { const sz = u32(q) * 2, f = u16(q + 4); if (sz < 6) break;
        if (f === 0x020B) org = [i16(q + 8), i16(q + 6)]; if (f === 0x020C) ext = [i16(q + 8), i16(q + 6)]; if (!f) break; q += sz; }
      if (!ext) throw new Error('wmf no extent');
      box = [org[0], org[1], org[0] + ext[0], org[1] + ext[1]];
      inch = 1440;
    }
    const lw = Math.abs(box[2] - box[0]), lh = Math.abs(box[3] - box[1]);
    if (!lw || !lh) throw new Error('empty wmf');
    let pw = lw / inch * 96, ph = lh / inch * 96;
    const k = Math.min(EMF_MAX / Math.max(pw, ph), Math.max(pw, ph) < 600 ? 2 : 1);
    pw *= k; ph *= k;
    const cv = document.createElement('canvas');
    cv.width = Math.max(1, Math.round(pw)); cv.height = Math.max(1, Math.round(ph));
    const g = cv.getContext('2d');
    const objs = [];
    const add = (o) => { let i = 0; while (objs[i]) i++; objs[i] = o; };
    let st = { pen: { t: 'p', c: '#000', w: 0 }, brush: { t: 'b', c: '#fff' }, font: null, textColor: '#000', bkColor: '#fff', bkMode: 2, align: 0, fill: 'evenodd',
      wOrg: [box[0], box[1]], wExt: [box[2] - box[0], box[3] - box[1]], cur: [0, 0], clip: null };
    const stack = [];
    const copy = (x) => Object.assign({}, x, { wOrg: x.wOrg.slice(), wExt: x.wExt.slice(), cur: x.cur.slice() });
    // 視窗（邏輯座標）直接對應整張畫布；視埠設定在 placeable 圖裡沒有意義
    const devM = () => new DOMMatrix().scale(cv.width / (st.wExt[0] || 1), cv.height / (st.wExt[1] || 1)).translate(-st.wOrg[0], -st.wOrg[1]);
    const withClip = (fn) => {
      g.save();
      if (st.clip) { g.setTransform(1, 0, 0, 1, 0, 0); g.beginPath(); for (const r of st.clip) g.rect(r[0], r[1], r[2] - r[0], r[3] - r[1]); g.clip(); }
      g.setTransform(devM()); fn(); g.restore();
    };
    const toDev = (l, t, r, b) => { const m = devM(), a = m.transformPoint(new DOMPoint(l, t)), c = m.transformPoint(new DOMPoint(r, b));
      return [Math.min(a.x, c.x), Math.min(a.y, c.y), Math.max(a.x, c.x), Math.max(a.y, c.y)]; };
    const andClip = (rects) => {
      if (!st.clip) { st.clip = rects; return; }
      const out = [];
      for (const a of st.clip) for (const b of rects) { const r = [Math.max(a[0], b[0]), Math.max(a[1], b[1]), Math.min(a[2], b[2]), Math.min(a[3], b[3])]; if (r[2] > r[0] && r[3] > r[1]) out.push(r); }
      st.clip = out;
    };
    const pat = (img) => (img ? g.createPattern(img, 'repeat') : '#c0c0c0');
    const fillOf = (c) => { if (c && typeof c !== 'string') c.setTransform(devM().inverse()); return c; };  // 圖樣以畫布像素重複
    const paint = (q, doFill) => withClip(() => {
      if (doFill && st.brush && st.brush.c) { g.fillStyle = fillOf(st.brush.c); g.fill(q, st.fill); }
      if (st.pen && st.pen.c) { const m = devM(); g.strokeStyle = st.pen.c; g.lineWidth = st.pen.w > 0 ? st.pen.w : 1 / (Math.abs(m.a) || 1); g.stroke(q); }
    });
    const ptsXY = (o, n) => { const a = []; for (let i = 0; i < n; i++) a.push([i16(o + i * 4), i16(o + i * 4 + 2)]); return a; };
    const poly = (a, closed) => { const q = new Path2D(); a.forEach((pt, i) => (i ? q.lineTo(pt[0], pt[1]) : q.moveTo(pt[0], pt[1]))); if (closed) q.closePath(); return q; };
    const text = (x, y, bytes, dx) => {
      const f = st.font || { h: -12, weight: 400, italic: 0, esc: 0, name: 'sans-serif', cs: 0 };
      let txt; try { txt = new TextDecoder(CHARSET[f.cs] || 'big5').decode(bytes); } catch (e) { txt = new TextDecoder('big5').decode(bytes); }
      withClip(() => {
        g.font = (f.italic ? 'italic ' : '') + (f.weight >= 600 ? 'bold ' : '') + (Math.abs(f.h) || 12) + "px '" + f.name + "', 'Microsoft JhengHei', 'PingFang TC', sans-serif";
        g.fillStyle = st.textColor;
        const va = st.align & 24; g.textBaseline = va === 24 ? 'alphabetic' : va === 8 ? 'bottom' : 'top';
        const ha = st.align & 6; g.textAlign = ha === 6 ? 'center' : ha === 2 ? 'right' : 'left';
        g.translate(x, y); if (f.esc) g.rotate(-f.esc / 10 * Math.PI / 180);
        if (dx && ha === 0) { let cx = 0, bi = 0; for (const ch of txt) { g.fillText(ch, cx, 0); const nb = ch.charCodeAt(0) > 127 && CHARSET[f.cs] !== 'windows-1252' ? 2 : 1; for (let j = 0; j < nb && bi < dx.length; j++) cx += dx[bi++]; } }
        else g.fillText(txt, 0, 0);
      });
    };
    const blit = (img, sx, sy, sw, sh, dx, dy, dw, dh) => withClip(() => { g.imageSmoothingQuality = 'high'; g.drawImage(img, sx, img.height - sy - sh, sw, sh, dx, dy, dw, dh); });
    while (p + 6 <= u8.length) {
      const size = u32(p) * 2, f = u16(p + 4), o = p + 6;
      if (size < 6 || p + size > u8.length || f === 0) break;
      try { switch (f) {
        case 0x020B: st.wOrg = [i16(o + 2), i16(o)]; break;
        case 0x020C: st.wExt = [i16(o + 2), i16(o)]; break;
        case 0x0106: st.fill = u16(o) === 2 ? 'nonzero' : 'evenodd'; break;
        case 0x012E: st.align = u16(o); break;
        case 0x0209: st.textColor = rgb(o); break;
        case 0x0201: st.bkColor = rgb(o); break;
        case 0x0102: st.bkMode = u16(o); break;
        case 0x001E: stack.push(copy(st)); break;
        case 0x0127: { const n = i16(o); const idx = n < 0 ? stack.length + n : n - 1; if (stack[idx]) { st = stack[idx]; stack.length = idx; } break; }
        case 0x02FA: add({ t: 'p', c: (u16(o) & 15) === 5 ? null : rgb(o + 6), w: i16(o + 2) }); break;
        case 0x02FC: { const bs = u16(o);
          add({ t: 'b', c: bs === 1 ? null : bs === 2 ? pat(hatchCanvas(rgb(o + 2), u16(o + 6), st.bkMode === 1 ? null : st.bkColor)) : rgb(o + 2) }); break; }
        case 0x02FB: { let e = o + 18; while (e < p + size && u8[e]) e++;
          add({ t: 'f', h: i16(o), esc: i16(o + 4), weight: i16(o + 8), italic: u8[o + 10], cs: u8[o + 13],
            name: new TextDecoder(CHARSET[u8[o + 13]] || 'big5').decode(u8.subarray(o + 18, e)) }); break; }
        case 0x0142: add({ t: 'b', c: pat(dibCanvas(u8, o + 4, null)) }); break;   // DIB 圖樣筆刷
        case 0x01F9: add({ t: 'b', c: pat(mono16(u8, o)) }); break;               // 單色圖樣筆刷
        case 0x00F7: add({ t: 'x' }); break;                                      // 調色盤：佔位
        case 0x06FF: {                                                            // 區域：掃描線 → 矩形
          const n = u16(o + 12), rects = []; let q = o + 22;
          for (let i = 0; i < n && q + 6 <= p + size; i++) {
            const cnt = u16(q), top = i16(q + 2), bot = i16(q + 4);
            for (let j = 0; j < cnt / 2; j++) rects.push([i16(q + 6 + j * 4), top, i16(q + 8 + j * 4), bot]);
            q += 8 + cnt * 2;
          }
          add({ t: 'r', rects }); break; }
        case 0x012C: { const ob = objs[u16(o)]; st.clip = ob && ob.t === 'r' ? ob.rects.map((r) => toDev(r[0], r[1], r[2], r[3])) : null; break; }
        case 0x012D: { const ob = objs[u16(o)]; if (ob) { if (ob.t === 'p') st.pen = ob; else if (ob.t === 'b') st.brush = ob; else if (ob.t === 'f') st.font = ob; } break; }
        case 0x01F0: objs[u16(o)] = undefined; break;
        case 0x0416: andClip([toDev(i16(o + 6), i16(o + 4), i16(o + 2), i16(o))]); break;
        case 0x0214: st.cur = [i16(o + 2), i16(o)]; break;
        case 0x0213: { const q = new Path2D(); q.moveTo(st.cur[0], st.cur[1]); st.cur = [i16(o + 2), i16(o)]; q.lineTo(st.cur[0], st.cur[1]); paint(q, false); break; }
        case 0x0324: paint(poly(ptsXY(o + 2, u16(o)), true), true); break;
        case 0x0325: paint(poly(ptsXY(o + 2, u16(o)), false), false); break;
        case 0x0538: { const n = u16(o); let q = o + 2 + n * 2; const all = new Path2D();
          for (let i = 0; i < n; i++) { const c = u16(o + 2 + i * 2); all.addPath(poly(ptsXY(q, c), true)); q += c * 4; }
          paint(all, true); break; }
        case 0x041B: case 0x061C: { const j = f === 0x061C ? 4 : 0, b = i16(o + j), r = i16(o + j + 2), t = i16(o + j + 4), l = i16(o + j + 6);
          const q = new Path2D(); q.rect(l, t, r - l, b - t); paint(q, true); break; }
        case 0x0418: { const b = i16(o), r = i16(o + 2), t = i16(o + 4), l = i16(o + 6); const q = new Path2D();
          q.ellipse((l + r) / 2, (t + b) / 2, Math.abs(r - l) / 2, Math.abs(b - t) / 2, 0, 0, Math.PI * 2); paint(q, true); break; }
        case 0x041D: { const rop = u32(o), h = i16(o + 4), w = i16(o + 6), y = i16(o + 8), x = i16(o + 10);
          const c = rop === 0x00FF0062 ? '#fff' : rop === 0x00000042 ? '#000' : rop === 0x00F00021 ? st.brush && st.brush.c : null;
          if (c) withClip(() => { g.fillStyle = fillOf(c); g.fillRect(x, y, w, h); }); break; }
        case 0x0521: { const n = u16(o), s0 = o + 2, pad = (n + 1) & ~1; text(i16(s0 + pad + 2), i16(s0 + pad), u8.subarray(s0, s0 + n)); break; }
        case 0x0A32: { const y = i16(o), x = i16(o + 2), n = u16(o + 4), opt = u16(o + 6); const s0 = o + 8 + (opt & 6 ? 8 : 0);
          const pad = (n + 1) & ~1, dxs = [];
          for (let i = 0; i < n && s0 + pad + i * 2 + 2 <= p + size; i++) dxs.push(i16(s0 + pad + i * 2));
          text(x, y, u8.subarray(s0, s0 + n), dxs.length === n ? dxs : null); break; }
        case 0x0B41: case 0x0F43: case 0x0940: {
          // DIBSTRETCHBLT／STRETCHDIB／DIBBITBLT：參數倒著放，最後接 DIB
          const sd = f === 0x0F43 ? 2 : 0, q = o + 4 + sd;
          let sh, sw, sy, sx, dh, dw, dy, dx, bmi;
          if (f === 0x0940) { sy = i16(q); sx = i16(q + 2); dh = sh = i16(q + 4); dw = sw = i16(q + 6); dy = i16(q + 8); dx = i16(q + 10); bmi = q + 12; }
          else { sh = i16(q); sw = i16(q + 2); sy = i16(q + 4); sx = i16(q + 6); dh = i16(q + 8); dw = i16(q + 10); dy = i16(q + 12); dx = i16(q + 14); bmi = q + 16; }
          if (bmi >= p + size) break;                          // 沒帶點陣圖的變體
          const img = dibCanvas(u8, bmi, null);
          if (img) blit(img, sx, sy, sw, sh, dx, dy, dw, dh); break; }
      } } catch (e) { /* 單一紀錄壞掉就跳過 */ }
      p += size;
    }
    return cv.toDataURL('image/png');
  }

  // ---------- EMF → PNG ----------
  // 瀏覽器不支援 EMF：自己把 GDI 紀錄畫到 canvas 再轉 PNG。EMF+ 註解略過（Word 存的是雙格式，GDI 那份就夠）。
  // ponytail: 只實作 Word 圖表／截圖常見的紀錄；遇到沒實作的紀錄就跳過，畫面可能缺東西但不會壞掉。
  const EMF_MAX = 2400;                      // 輸出最長邊（px）
  function emfToPng(u8) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    if (dv.getUint32(0, true) !== 1 || dv.getUint32(40, true) !== 0x464D4520) throw new Error('not emf');
    const bL = dv.getInt32(8, true), bT = dv.getInt32(12, true), bR = dv.getInt32(16, true), bB = dv.getInt32(20, true);
    const bw = bR - bL + 1, bh = bB - bT + 1;
    if (bw <= 1 || bh <= 1) throw new Error('empty emf');
    const k = Math.min(1, EMF_MAX / Math.max(bw, bh)) * (Math.max(bw, bh) < 600 ? 2 : 1);
    const cv = document.createElement('canvas');
    cv.width = Math.max(1, Math.round(bw * k)); cv.height = Math.max(1, Math.round(bh * k));
    const g = cv.getContext('2d');
    const i32 = (o) => dv.getInt32(o, true), u32 = (o) => dv.getUint32(o, true), i16 = (o) => dv.getInt16(o, true), f32 = (o) => dv.getFloat32(o, true);
    const rgb = (o) => 'rgb(' + u8[o] + ',' + u8[o + 1] + ',' + u8[o + 2] + ')';
    const STOCK = { 0: { t: 'b', c: '#fff' }, 1: { t: 'b', c: '#c0c0c0' }, 2: { t: 'b', c: '#808080' }, 3: { t: 'b', c: '#404040' }, 4: { t: 'b', c: '#000' }, 5: { t: 'b', c: null },
      6: { t: 'p', c: '#fff', w: 0 }, 7: { t: 'p', c: '#000', w: 0 }, 8: { t: 'p', c: null, w: 0 } };
    const objs = [];
    let st = { pen: STOCK[7], brush: STOCK[0], font: null, textColor: '#000', bkColor: '#fff', bkMode: 2, align: 0, fill: 'evenodd',
      world: new DOMMatrix(), wOrg: [0, 0], wExt: [1, 1], vOrg: [0, 0], vExt: [1, 1], mapMode: 1, cur: [0, 0], clip: null };
    const stack = [];
    const copy = (s) => Object.assign({}, s, { world: DOMMatrix.fromMatrix(s.world), wOrg: s.wOrg.slice(), wExt: s.wExt.slice(), vOrg: s.vOrg.slice(), vExt: s.vExt.slice(), cur: s.cur.slice() });
    const devM = () => {                     // 邏輯座標 → canvas 像素
      const aniso = st.mapMode === 7 || st.mapMode === 8;
      const sx = aniso ? st.vExt[0] / (st.wExt[0] || 1) : 1, sy = aniso ? st.vExt[1] / (st.wExt[1] || 1) : 1;
      return new DOMMatrix().scale(k).translate(-bL, -bT).translate(st.vOrg[0], st.vOrg[1]).scale(sx, sy).translate(-st.wOrg[0], -st.wOrg[1]).multiply(st.world);
    };
    let path = null;                          // BEGINPATH 之後累積的路徑
    const withClip = (fn) => {
      g.save();
      if (st.clip) { g.setTransform(1, 0, 0, 1, 0, 0); g.beginPath(); for (const r of st.clip) g.rect(r[0], r[1], r[2] - r[0], r[3] - r[1]); g.clip(); }
      if (st.clipPath) { g.setTransform(st.clipPath.m); g.clip(st.clipPath.p, st.clipPath.rule); }
      g.setTransform(devM()); fn(); g.restore();
    };
    const pat = (img) => (img ? g.createPattern(img, 'repeat') : '#c0c0c0');
    const fillOf = (c) => { if (c && typeof c !== 'string') c.setTransform(devM().inverse()); return c; };  // 圖樣以畫布像素重複
    const lineW = () => { const m = devM(); const sc = Math.hypot(m.a, m.b) || 1; return st.pen.w > 0 ? st.pen.w : 1 / sc; };
    const paint = (p, doFill, doStroke) => withClip(() => {
      if (doFill && st.brush && st.brush.c) { g.fillStyle = fillOf(st.brush.c); g.fill(p, st.fill); }
      if (doStroke && st.pen && st.pen.c) { g.strokeStyle = st.pen.c; g.lineWidth = lineW(); g.lineJoin = 'round'; g.stroke(p); }
    });
    const shape = (p, closed) => { if (path) { path.addPath(p); return; } paint(p, closed, true); };
    const pts16 = (o, n) => { const a = []; for (let i = 0; i < n; i++) a.push([i16(o + i * 4), i16(o + i * 4 + 2)]); return a; };
    const poly = (a, closed, bez) => {
      const p = new Path2D(); if (!a.length) return p;
      p.moveTo(a[0][0], a[0][1]);
      if (bez) for (let i = 1; i + 2 < a.length; i += 3) p.bezierCurveTo(a[i][0], a[i][1], a[i + 1][0], a[i + 1][1], a[i + 2][0], a[i + 2][1]);
      else for (let i = 1; i < a.length; i++) p.lineTo(a[i][0], a[i][1]);
      if (closed) p.closePath();
      return p;
    };
    const devRect = (l, t, r, b) => { const m = devM(); const p1 = m.transformPoint(new DOMPoint(l, t)), p2 = m.transformPoint(new DOMPoint(r, b));
      return [Math.min(p1.x, p2.x), Math.min(p1.y, p2.y), Math.max(p1.x, p2.x), Math.max(p1.y, p2.y)]; };
    const andClip = (rects) => {
      if (!st.clip) { st.clip = rects; return; }
      const out = [];
      for (const a of st.clip) for (const b of rects) { const r = [Math.max(a[0], b[0]), Math.max(a[1], b[1]), Math.min(a[2], b[2]), Math.min(a[3], b[3])]; if (r[2] > r[0] && r[3] > r[1]) out.push(r); }
      st.clip = out;
    };
    const dib = (rec, offBmi, offBits) => dibCanvas(u8, rec + offBmi, rec + offBits);
    const blit = (img, sx, sy, sw, sh, dx, dy, dw, dh) => withClip(() => {
      g.imageSmoothingQuality = 'high';
      // DIB 的來源座標是由下往上；drawImage 用的是由上往下，這裡已經在 dib() 翻好了
      g.drawImage(img, sx, img.height - sy - sh, sw, sh, dx, dy, dw, dh);
    });
    const dec = new TextDecoder('utf-16le');
    let p = 0;
    while (p + 8 <= u8.length) {
      const t = u32(p), size = u32(p + 4);
      if (size < 8 || p + size > u8.length) break;
      const o = p + 8;
      try { switch (t) {
        case 9: st.wExt = [i32(o), i32(o + 4)]; break;
        case 10: st.wOrg = [i32(o), i32(o + 4)]; break;
        case 11: st.vExt = [i32(o), i32(o + 4)]; break;
        case 12: st.vOrg = [i32(o), i32(o + 4)]; break;
        case 17: st.mapMode = u32(o); break;
        case 18: st.bkMode = u32(o); break;
        case 19: st.fill = u32(o) === 2 ? 'nonzero' : 'evenodd'; break;
        case 22: st.align = u32(o); break;
        case 24: st.textColor = rgb(o); break;
        case 25: st.bkColor = rgb(o); break;
        case 27: st.cur = [i32(o), i32(o + 4)]; if (path) path.moveTo(st.cur[0], st.cur[1]); break;
        case 54: { const q = new Path2D(); q.moveTo(st.cur[0], st.cur[1]); st.cur = [i32(o), i32(o + 4)]; q.lineTo(st.cur[0], st.cur[1]);
          if (path) path.lineTo(st.cur[0], st.cur[1]); else paint(q, false, true); break; }
        case 30: andClip([devRect(i32(o), i32(o + 4), i32(o + 8), i32(o + 12))]); break;
        case 33: stack.push(copy(st)); break;
        case 34: { const n = i32(o); const idx = n < 0 ? stack.length + n : n - 1; if (stack[idx]) { st = stack[idx]; stack.length = idx; } break; }
        case 35: st.world = new DOMMatrix([f32(o), f32(o + 4), f32(o + 8), f32(o + 12), f32(o + 16), f32(o + 20)]); break;
        case 36: { const x = new DOMMatrix([f32(o), f32(o + 4), f32(o + 8), f32(o + 12), f32(o + 16), f32(o + 20)]), mode = u32(o + 24);
          st.world = mode === 1 ? new DOMMatrix() : mode === 2 ? st.world.multiply(x) : mode === 3 ? x.multiply(st.world) : x; break; }
        case 37: { const h = u32(o); const ob = h & 0x80000000 ? STOCK[h & 0x7fffffff] : objs[h];
          if (ob) { if (ob.t === 'p') st.pen = ob; else if (ob.t === 'b') st.brush = ob; else if (ob.t === 'f') st.font = ob; } break; }
        case 38: objs[u32(o)] = { t: 'p', c: (u32(o + 4) & 15) === 5 ? null : rgb(o + 16), w: i32(o + 8) }; break;
        case 39: { const bs = u32(o + 4);
          objs[u32(o)] = { t: 'b', c: bs === 1 ? null : bs === 2 ? pat(hatchCanvas(rgb(o + 8), u32(o + 12), st.bkMode === 1 ? null : st.bkColor)) : rgb(o + 8) }; break; }
        case 93: case 94: objs[u32(o)] = { t: 'b', c: pat(dibCanvas(u8, p + u32(o + 8), p + u32(o + 16))) }; break;   // 單色／DIB 圖樣筆刷
        case 40: delete objs[u32(o)]; break;
        case 95: objs[u32(o)] = { t: 'p', c: (u32(o + 20) & 15) === 5 || u32(o + 28) === 1 ? null : rgb(o + 32), w: u32(o + 24) }; break;
        case 82: { const lf = o + 4, name = dec.decode(u8.subarray(lf + 28, lf + 92)).replace(/\0.*$/s, '');
          objs[u32(o)] = { t: 'f', h: i32(lf), weight: i32(lf + 16), italic: u8[lf + 20], esc: i32(lf + 8), name }; break; }
        case 28: break;                                        // SETMETARGN：沿用目前裁切
        case 75: { const cb = u32(o), mode = u32(o + 4);
          if (mode === 5 && !cb) { st.clip = null; st.clipPath = null; break; }
          const rd = o + 8, n = u32(rd + 8), rects = [];
          for (let i = 0; i < n; i++) { const q = rd + 32 + i * 16; rects.push([(i32(q) - bL) * k, (i32(q + 4) - bT) * k, (i32(q + 8) - bL) * k, (i32(q + 12) - bT) * k]); }
          if (mode === 5) st.clip = rects; else if (mode === 1) andClip(rects); break; }
        case 59: path = new Path2D(); break;
        case 67: if (path) { st.clipPath = { p: path, m: devM(), rule: st.fill }; path = null; } break;   // SELECTCLIPPATH
        case 60: break;
        case 61: if (path) path.closePath(); break;
        case 62: case 63: case 64: if (path) { const q = path; path = null; paint(q, t !== 64, t !== 62); } break;
        case 85: case 86: case 87: case 88: case 89: {
          const a = pts16(o + 20, u32(o + 16));
          if (t === 88 || t === 89) a.unshift(st.cur.slice());
          if (path && (t === 88 || t === 89)) { const q = poly(a, false, t === 88); path.addPath(q); }
          else shape(poly(a, t === 86, t === 85 || t === 88), t === 86);
          if (a.length) st.cur = a[a.length - 1].slice(); break; }
        case 90: case 91: {
          const n = u32(o + 16); let q = o + 24 + n * 4; const all = new Path2D();
          for (let i = 0; i < n; i++) { const c = u32(o + 24 + i * 4); all.addPath(poly(pts16(q, c), t === 91)); q += c * 4; }
          shape(all, t === 91); break; }
        case 43: { const q = new Path2D(); q.rect(i32(o), i32(o + 4), i32(o + 8) - i32(o), i32(o + 12) - i32(o + 4)); shape(q, true); break; }
        case 42: { const l = i32(o), tt = i32(o + 4), r = i32(o + 8), b = i32(o + 12); const q = new Path2D();
          q.ellipse((l + r) / 2, (tt + b) / 2, Math.abs(r - l) / 2, Math.abs(b - tt) / 2, 0, 0, Math.PI * 2); shape(q, true); break; }
        case 84: {                                             // EXTTEXTOUTW
          const e = o + 28, x = i32(e), y = i32(e + 4), n = u32(e + 8), offS = u32(e + 12), offDx = u32(e + 36);
          if (!n) break;
          const txt = dec.decode(u8.subarray(p + offS, p + offS + n * 2));
          const f = st.font || { h: -12, weight: 400, italic: 0, esc: 0, name: 'sans-serif' };
          withClip(() => {
            const px = Math.abs(f.h) || 12;
            g.font = (f.italic ? 'italic ' : '') + (f.weight >= 600 ? 'bold ' : '') + px + "px '" + f.name + "', 'Microsoft JhengHei', 'PingFang TC', sans-serif";
            g.fillStyle = st.textColor;
            const va = st.align & 24; g.textBaseline = va === 24 ? 'alphabetic' : va === 8 ? 'bottom' : 'top';
            const ha = st.align & 6;
            g.translate(x, y); if (f.esc) g.rotate(-f.esc / 10 * Math.PI / 180);
            if (offDx) {                                       // 逐字位置照 Word 算好的間距擺
              let total = 0; const dx = []; for (let i = 0; i < n; i++) { dx.push(i32(p + offDx + i * 4)); total += dx[i]; }
              let cx = ha === 6 ? -total / 2 : ha === 2 ? -total : 0; g.textAlign = 'left';
              for (let i = 0; i < n; i++) { g.fillText(txt[i], cx, 0); cx += dx[i]; }
            } else { g.textAlign = ha === 6 ? 'center' : ha === 2 ? 'right' : 'left'; g.fillText(txt, 0, 0); }
          });
          break; }
        case 76: {                                             // BITBLT
          const dx = i32(o + 16), dy = i32(o + 20), dw = i32(o + 24), dh = i32(o + 28), rop = u32(o + 32);
          const offBmi = u32(o + 68), cbBmi = u32(o + 72), offBits = u32(o + 76), cbBits = u32(o + 80);
          if (!cbBmi) {
            const c = rop === 0x00FF0062 ? '#fff' : rop === 0x00000042 ? '#000' : rop === 0x00F00021 ? st.brush && st.brush.c : null;  // 其他 ROP 需要讀目的地，略過
            if (c) withClip(() => { g.fillStyle = fillOf(c); g.fillRect(dx, dy, dw, dh); });
          } else { const img = dib(p, offBmi, offBits, cbBits); if (img) blit(img, i32(o + 36), i32(o + 40), dw, dh, dx, dy, dw, dh); }
          break; }
        case 81: {                                             // STRETCHDIBITS
          const dx = i32(o + 16), dy = i32(o + 20), sx = i32(o + 24), sy = i32(o + 28), sw = i32(o + 32), sh = i32(o + 36);
          const img = dib(p, u32(o + 40), u32(o + 48), u32(o + 52));
          if (img) blit(img, sx, sy, sw, sh, dx, dy, i32(o + 64), i32(o + 68));
          break; }
        case 14: p = u8.length; break;                         // EOF
      } } catch (e) { /* 單一紀錄壞掉就跳過 */ }
      p += size;
    }
    return cv.toDataURL('image/png');
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
    let body = '', head = '', inHead = true;
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
        for (const p of tags(m.tc, 'p')) { const r = docxParagraph(p, ctx); inner += r.html + r.after; }
        if (!inner) inner = '<p><br></p>';
        // 儲存格垂直對齊：Word 預設靠上（CSS 已預設），置中／靠下才另外標
        const va = { center: 'middle', bottom: 'bottom' }[val(tag(tcPr, 'vAlign'), 'val')];
        // 儲存格底色：w:shd 的 fill
        const fill = val(tag(tcPr, 'shd'), 'fill'),
          bg = fill && /^[0-9a-f]{6}$/i.test(fill) && fill.toUpperCase() !== 'FFFFFF' ? 'background-color:#' + fill : '';
        // 儲存格自己的框線（tcBorders）：nil／none → hidden（蓋過鄰格），其餘照樣式、粗細（1/8 pt）、顏色
        const tb = tag(tcPr, 'tcBorders'), bd = [];
        const BS = { single: 'solid', dashed: 'dashed', dotted: 'dotted', double: 'double', thick: 'solid' };
        for (const [side, names] of [['top', ['top']], ['right', ['right', 'end']], ['bottom', ['bottom']], ['left', ['left', 'start']]]) {
          const e = names.map((n) => tag(tb, n)).find(Boolean);
          if (!e) continue;
          const v = val(e, 'val');
          if (v === 'nil' || v === 'none') { bd.push('border-' + side + ':hidden'); continue; }
          const col = val(e, 'color');
          bd.push('border-' + side + ':' + Math.max(0.25, num(val(e, 'sz'), 4) / 8) + 'pt ' + (BS[v] || 'solid') + ' ' + (col && /^[0-9a-f]{6}$/i.test(col) ? '#' + col : '#000'));
        }
        const tdCss = [va ? 'vertical-align:' + va : '', bg, ...bd].filter(Boolean).join(';');
        cells += '<td' + attrs + (tdCss ? ' style="' + tdCss + '"' : '') + '>' + inner + '</td>';
      }
      // 列高：Word 的 trHeight（twips；exact／atLeast 都當最小高度）
      const trH = num(val(tag(tag(tr, 'trPr'), 'trHeight'), 'val'), 0);
      const row = '<tr' + (trH ? ' style="height:' + (trH / 20).toFixed(1) + 'pt"' : '') + '>' + cells + '</tr>';
      // 開頭連續標了 tblHeader 的列＝跨頁重複的標題列 → <thead>
      inHead = inHead && !!tag(tag(tr, 'trPr'), 'tblHeader') && val(tag(tag(tr, 'trPr'), 'tblHeader'), 'val') !== '0';
      if (inHead) head += row; else body += row;
    });
    let colgroup = '';
    if (cols.length && totalW) colgroup = '<colgroup>' + cols.map((w) => '<col style="width:' + ((w / totalW) * 100).toFixed(2) + '%">').join('') + '</colgroup>';
    return '<table>' + colgroup + (head ? '<thead>' + head + '</thead>' : '') + '<tbody>' + body + '</tbody></table>';
  }

  // ---------- 頁首／頁尾 ----------
  function convertEdge(text, ctx) {
    if (!text) return null;
    const doc = xml(text);
    const root = [...doc.getElementsByTagName('*')].find((e) => e.localName === 'hdr' || e.localName === 'ftr');
    if (!root) return null;
    let out = '';
    for (const p of tags(root, 'p')) {
      const { html, after } = docxParagraph(p, ctx);
      out += html + after;
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
      if (!/^word\/media\//i.test(name)) continue;
      let b = entry.method === 0 ? entry.raw : await entryBytes(zip, name);
      // EMZ／WMZ 就是 gzip 過的 EMF／WMF：先解開，後面當一般向量圖處理
      if (/\.(emz|wmz)$/i.test(name)) {
        try { b = new Uint8Array(await new Response(new Blob([b]).stream().pipeThrough(new DecompressionStream('gzip'))).arrayBuffer()); } catch (e) { continue; }
      }
      ctx.media.set(name, b);
    }
    const doc = xml(documentXml);
    const mar0 = [...doc.getElementsByTagName('*')].find((e) => e.localName === 'pgMar');
    ctx.marL = mar0 ? num(val(mar0, 'left'), 1134) / 20 : 0;       // 相對「頁面」的水平位移要扣掉左邊界（pt）
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
          html.push('<li' + (r.html.match(/style="([^"]*)"/) ? ' style="' + r.html.match(/style="([^"]*)"/)[1] + '"' : '') + '>' + r.html.replace(/^<[^>]+>|<\/[^>]+>$/g, '') + r.after + '</li>');
        } else {
          closeList(0);
          html.push((r.html || '<p><br></p>') + r.after);
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
      if (ctx.vectorImages) ctx.warnings.push(ctx.vectorImages + ' 張 EMF/WMF 向量圖無法轉換，已用文字標示位置');
      if (header && /\bPAGE\b/.test(header)) ctx.warnings.push('頁尾的頁碼欄位會顯示 Word 上次存檔的數字');
    }
    return { html: html.join('\n'), page, header, footer, warnings: ctx.warnings, lists: listCount, breaks: (html.join('').match(/class="pagebreak"/g) || []).length };
  }

  window.docxFidelity = { convert, unzip };
})();
