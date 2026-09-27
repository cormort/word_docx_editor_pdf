/* pdf-export.js — 編輯器內容 → 文字可選取、可搜尋的 PDF
 * 做法移植自 pdfviewer_v2 的 docx.js：
 *   1. 依紙張與邊界把內容排進固定大小的頁（分頁符號強制換頁；表格、清單可跨頁拆開）
 *   2. 每頁交給瀏覽器自己的排版引擎畫成圖（SVG foreignObject），不行才改用 html2canvas
 *   3. 用 pdf-lib 把圖放進 PDF，再疊一層看不見的文字（render mode 3），位置取自瀏覽器實際排版
 * 文字層用 Identity-H 的 CID 字型而且不內嵌字型檔：看不見的字不需要字形，
 * 所以不必下載中文字型，PDF 閱讀器仍能選取、複製、搜尋每個字。
 */
(function () {
  'use strict';

  const PT = 0.75; // px → pt
  const MM = 96 / 25.4; // mm → px
  const FONT = 'WdeText';

  // ---------- 分頁 ----------
  // 從表格尾端一列一列往下一頁搬，直到這頁放得下；第一列全是 th 時當標題列，每頁重複。
  // 回傳下一頁要接著放的表格，連一列都放不下就回傳 null。
  function splitTable(table, fits) {
    const rows = [...table.querySelectorAll(':scope > tr, :scope > * > tr')];
    const head = rows.length && [...rows[0].cells].every(c => c.tagName === 'TH') ? 1 : 0;
    if (rows.length - head < 2) return null;
    let kept = rows.length;
    while (kept > head && !fits()) rows[--kept].remove();
    if (kept === head) {
      for (const row of rows.slice(head)) (rows[0].parentNode || table).appendChild(row);
      return null;
    }
    // 跨頁的縱向合併儲存格：這頁到斷點為止，剩下的列數交給下一頁同一欄的格子
    const columnOf = cell => {
      let col = 0;
      for (const c of cell.parentNode.cells) {
        if (c === cell) return col;
        col += c.colSpan;
      }
    };
    rows.slice(0, kept).forEach((row, i) => {
      for (const cell of row.cells) {
        const over = i + cell.rowSpan - kept;
        if (cell.rowSpan < 2 || over <= 0) continue;
        cell.rowSpan -= over;
        const col = columnOf(cell);
        const filler = document.createElement(cell.tagName);
        filler.rowSpan = over;
        const next = rows[kept],
          after = [...next.cells].find(c => columnOf(c) >= col);
        next.insertBefore(filler, after || null);
      }
    });
    const rest = table.cloneNode(false);
    for (const cg of table.querySelectorAll(':scope > colgroup')) rest.appendChild(cg.cloneNode(true));
    const body = document.createElement('tbody');
    rest.appendChild(body);
    if (head) body.appendChild(rows[0].cloneNode(true));
    for (const row of rows.slice(kept)) body.appendChild(row);
    return rest;
  }
  // 清單同理，一個項目一個項目搬；編號清單在下一頁接著編號
  function splitList(list, fits) {
    const items = [...list.children].filter(li => li.tagName === 'LI');
    if (items.length < 2) return null;
    let kept = items.length;
    while (kept > 0 && !fits()) items[--kept].remove();
    if (!kept) {
      for (const li of items) list.appendChild(li);
      return null;
    }
    const rest = list.cloneNode(false);
    if (list.tagName === 'OL') rest.start = (list.start || 1) + kept;
    for (const li of items.slice(kept)) rest.appendChild(li);
    return rest;
  }
  const hasContent = body => body.textContent.trim() !== '' || !!body.querySelector('img,table,hr');

  function paginate(source, o) {
    const W = o.pageW * MM,
      H = o.pageH * MM;
    const [T, R, B, L] = o.margins.map(v => v * MM);
    const contentH = H - T - B;
    const host = document.createElement('div');
    host.style.cssText = 'position:fixed;left:-30000px;top:0;pointer-events:none';
    document.body.appendChild(host);
    const pages = [];
    let page, body;
    const newPage = () => {
      page = document.createElement('div');
      page.className = 'paper pdfpage';
      page.style.cssText = 'position:relative;width:' + W + 'px;height:' + H + 'px;overflow:hidden;background:#fff';
      body = document.createElement('div');
      body.className = 'pdfbody';
      body.style.cssText =
        'position:absolute;display:flow-root;left:' + L + 'px;top:' + T + 'px;width:' + (W - L - R) + 'px';
      page.appendChild(body);
      host.appendChild(page);
      pages.push(page);
    };
    const fits = () => body.offsetHeight <= contentH + 1;
    newPage();
    const blocks = [...source.childNodes].map(n => n.cloneNode(true));
    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      if (block.nodeType === 1 && block.classList.contains('pagebreak')) {
        if (body.childNodes.length) newPage();
        continue;
      }
      body.appendChild(block);
      if (fits()) continue;
      const rest =
        block.tagName === 'TABLE' ? splitTable(block, fits) : /^(UL|OL)$/.test(block.tagName) ? splitList(block, fits) : null;
      if (rest) {
        blocks.splice(i + 1, 0, rest);
        newPage();
        continue;
      }
      if (body.childNodes.length > 1) {
        body.removeChild(block);
        newPage();
        i--; // 放到新的一頁
        continue;
      }
      // 單一區塊就比一頁高（例如很長的圖片）：這一頁拉長，不切開內容
      page.style.height = body.offsetHeight + T + B + 'px';
    }
    const out = pages.filter(p => hasContent(p.firstChild));
    pages.filter(p => !out.includes(p)).forEach(p => p.remove());

    // 頁首、頁尾、頁碼放在邊界裡：也是真的文字，一樣可以搜尋
    const edge = (p, html, css) => {
      const d = document.createElement('div');
      d.style.cssText =
        'position:absolute;left:' + L + 'px;width:' + (W - L - R) + 'px;font-size:10pt;line-height:1.4;color:#333;' + css;
      d.innerHTML = html;
      p.appendChild(d);
    };
    out.forEach((p, i) => {
      const h = p.offsetHeight;
      if (o.header) edge(p, o.header, 'bottom:' + (h - T + 2 * MM) + 'px');
      if (o.footer) edge(p, o.footer, 'top:' + (h - B + 2 * MM) + 'px');
      if (o.pageNum)
        edge(p, '第 ' + (i + 1) + ' 頁 / 共 ' + out.length + ' 頁', 'bottom:' + 4 * MM + 'px;text-align:center;font-size:9pt;color:#666');
    });
    return { host, pages: out };
  }

  // ---------- 繪製 ----------
  // 頁面連同本頁的樣式放進 SVG <foreignObject> 再畫到 canvas：用的是瀏覽器真正的排版，
  // 比 html2canvas 用 JS 重新實作 CSS 快很多也準。Safari 會把這種 canvas 標成不可讀，那就改用 html2canvas。
  async function paintNative(page, styles, scale) {
    const w = page.offsetWidth,
      h = page.offsetHeight;
    const xml = new XMLSerializer();
    const body = styles.map(s => xml.serializeToString(s)).join('') + xml.serializeToString(page);
    const svg =
      '<svg xmlns="http://www.w3.org/2000/svg" width="' + w + '" height="' + h + '">' +
      '<foreignObject width="100%" height="100%"><div xmlns="http://www.w3.org/1999/xhtml">' + body +
      '</div></foreignObject></svg>';
    const img = new Image();
    img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
    await img.decode();
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(w * scale);
    canvas.height = Math.round(h * scale);
    const g = canvas.getContext('2d');
    g.fillStyle = '#fff';
    g.fillRect(0, 0, canvas.width, canvas.height);
    g.drawImage(img, 0, 0, canvas.width, canvas.height);
    return canvas;
  }
  function jpegBytes(canvas) {
    return new Promise((ok, fail) =>
      canvas.toBlob(
        b => (b ? b.arrayBuffer().then(buf => ok(new Uint8Array(buf)), fail) : fail(new Error('canvas 匯出失敗'))),
        'image/jpeg',
        0.9,
      ),
    );
  }

  // ---------- 隱形文字層 ----------
  const hex4 = n => n.toString(16).padStart(4, '0').toUpperCase();
  // 同一行的字合成一個 TJ：第一個字用 Tm 定位，後面每個字用 TJ 位移對到瀏覽器排的位置，
  // 閱讀器才會把整行當成一段文字（逐字定位會讓「預算」拆成兩段，搜尋就找不到）
  function textOps(page, widths) {
    const base = page.getBoundingClientRect();
    const height = page.offsetHeight;
    const range = document.createRange();
    const walker = document.createTreeWalker(page, NodeFilter.SHOW_TEXT);
    const glyphs = [];
    let node;
    while ((node = walker.nextNode())) {
      const size = parseFloat(getComputedStyle(node.parentElement).fontSize) || 16;
      const text = node.nodeValue;
      for (let i = 0; i < text.length; i++) {
        const code = text.charCodeAt(i);
        if ((code >= 0xd800 && code <= 0xdfff) || code < 32) continue; // 超出 2 位元組 CID／控制字元
        range.setStart(node, i);
        range.setEnd(node, i + 1);
        const r = range.getClientRects()[0];
        if (!r || !r.width) continue; // 被摺疊掉的空白
        if (!(code in widths)) widths[code] = Math.round((r.width / size) * 1000);
        glyphs.push({
          code,
          size: size * PT,
          x: (r.left - base.left) * PT,
          y: (height - (r.bottom - base.top) + size * 0.22) * PT,
        });
      }
    }
    const ops = ['BT 3 Tr'];
    let run = null;
    const flush = () => {
      if (run) ops.push('[' + run.parts.join(' ') + '] TJ');
      run = null;
    };
    for (const g of glyphs) {
      const sameLine = run && g.size === run.size && Math.abs(g.y - run.y) < g.size * 0.5 && g.x >= run.penX - g.size;
      if (!sameLine) {
        flush();
        ops.push('/' + FONT + ' ' + g.size.toFixed(2) + ' Tf 1 0 0 1 ' + g.x.toFixed(2) + ' ' + g.y.toFixed(2) + ' Tm');
        run = { size: g.size, y: g.y, penX: g.x, parts: [] };
      } else {
        const adjust = Math.round(((run.penX - g.x) * 1000) / g.size); // TJ 的數字是千分之一字級
        if (adjust) run.parts.push(adjust);
      }
      run.parts.push('<' + hex4(g.code) + '>');
      run.penX = g.x + (widths[g.code] * g.size) / 1000;
    }
    flush();
    ops.push('ET');
    return ops.join('\n');
  }
  function toUnicodeCMap() {
    const ranges = [];
    for (let hi = 0; hi < 256; hi++) {
      const h = hi.toString(16).padStart(2, '0').toUpperCase();
      ranges.push('<' + h + '00> <' + h + 'FF> <' + h + '00>');
    }
    const blocks = [];
    for (let i = 0; i < ranges.length; i += 100) {
      const part = ranges.slice(i, i + 100);
      blocks.push(part.length + ' beginbfrange\n' + part.join('\n') + '\nendbfrange');
    }
    return (
      '/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n' +
      '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n' +
      '/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n' +
      blocks.join('\n') +
      '\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend'
    );
  }
  function registerFont(ctx, fontRef, widths, PDFString) {
    const w = [];
    Object.keys(widths)
      .map(Number)
      .sort((a, b) => a - b)
      .forEach(c => w.push(c, [widths[c]]));
    const descriptor = ctx.register(
      ctx.obj({
        Type: 'FontDescriptor', FontName: FONT, Flags: 4, FontBBox: [0, -200, 1000, 900],
        ItalicAngle: 0, Ascent: 880, Descent: -120, CapHeight: 700, StemV: 80,
      }),
    );
    const cidFont = ctx.register(
      ctx.obj({
        Type: 'Font', Subtype: 'CIDFontType2', BaseFont: FONT,
        CIDSystemInfo: { Registry: PDFString.of('Adobe'), Ordering: PDFString.of('Identity'), Supplement: 0 },
        FontDescriptor: descriptor, DW: 1000, W: w, CIDToGIDMap: 'Identity',
      }),
    );
    const toUnicode = ctx.register(ctx.flateStream(toUnicodeCMap()));
    ctx.assign(
      fontRef,
      ctx.obj({ Type: 'Font', Subtype: 'Type0', BaseFont: FONT, Encoding: 'Identity-H', DescendantFonts: [cidFont], ToUnicode: toUnicode }),
    );
  }

  // o: { source, pageW, pageH, margins:[上,右,下,左]（mm）, header, footer, pageNum, scale,
  //      html2canvas: () => Promise<fn>, title, onProgress }
  async function build(o) {
    const { host, pages } = paginate(o.source, o);
    try {
      if (!pages.length) throw new Error('文件沒有內容');
      await Promise.all([...host.querySelectorAll('img')].map(im => im.decode().catch(() => {})));
      const { PDFDocument, PDFName, PDFString } = await import('./vendor/pdf-lib.esm.min.js');
      const pdf = await PDFDocument.create();
      if (o.title) pdf.setTitle(o.title);
      pdf.setCreator('Word 文件簡易編輯器');
      const ctx = pdf.context,
        fontRef = ctx.nextRef(),
        widths = {};
      const styles = [...document.querySelectorAll('style')];
      let native = true;
      for (const [i, page] of pages.entries()) {
        if (o.onProgress) o.onProgress(i + 1, pages.length);
        let canvas, bytes;
        if (native) {
          try {
            canvas = await paintNative(page, styles, o.scale);
            bytes = await jpegBytes(canvas); // canvas 被標成不可讀時這裡會丟例外
          } catch (e) {
            native = false;
          }
        }
        if (!native) {
          const h2c = await o.html2canvas();
          canvas = await h2c(page, { scale: o.scale, backgroundColor: '#fff', logging: false, useCORS: true });
          bytes = await jpegBytes(canvas);
        }
        const w = page.offsetWidth * PT,
          h = page.offsetHeight * PT;
        const jpg = await pdf.embedJpg(bytes);
        canvas.width = canvas.height = 0; // 手機 canvas 記憶體有限，用完就放掉
        const out = pdf.addPage([w, h]);
        out.drawImage(jpg, { x: 0, y: 0, width: w, height: h });
        out.node.setFontDictionary(PDFName.of(FONT), fontRef);
        out.node.addContentStream(ctx.register(ctx.flateStream(textOps(page, widths))));
      }
      registerFont(ctx, fontRef, widths, PDFString);
      return { bytes: await pdf.save(), pages: pages.length, native };
    } finally {
      host.remove();
    }
  }

  window.pdfExport = { build, paginate };
})();
