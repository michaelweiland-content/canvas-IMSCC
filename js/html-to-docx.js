/*
 * html-to-docx.js
 * Converts the small subset of HTML that Canvas uses in question text
 * (paragraphs, bold/italic/underline, super/subscript, lists, tables,
 * links, images, equations) into docx.js Paragraph / Table objects.
 */
(function (global) {
  'use strict';

  const CQE = (global.CQE = global.CQE || {});
  const D = () => global.docx;

  const BLOCK_TAGS = new Set([
    'p', 'div', 'section', 'article', 'header', 'footer', 'main', 'aside', 'center',
    'address', 'figure', 'figcaption', 'pre', 'blockquote', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'dl', 'dt', 'dd'
  ]);
  const SKIP_TAGS = new Set(['script', 'style', 'noscript', 'iframe', 'object', 'embed', 'audio', 'video', 'source', 'track', 'input', 'button', 'select', 'textarea']);
  const MAX_IMAGE_WIDTH_PX = 560; // ~5.8 in at 96 dpi
  const TWIPS_PER_INDENT = 360;

  function runProps(fmt) {
    const p = {};
    if (fmt.bold) p.bold = true;
    if (fmt.italics) p.italics = true;
    if (fmt.underline) p.underline = {};
    if (fmt.strike) p.strike = true;
    if (fmt.superScript) p.superScript = true;
    if (fmt.subScript) p.subScript = true;
    if (fmt.color) p.color = fmt.color;
    if (fmt.size) p.size = fmt.size;
    if (fmt.font) p.font = fmt.font;
    if (fmt.highlight) p.highlight = fmt.highlight;
    return p;
  }

  function styleFmt(el, fmt) {
    const st = (el.getAttribute && el.getAttribute('style')) || '';
    if (!st) return fmt;
    const f = Object.assign({}, fmt);
    if (/font-weight\s*:\s*(bold|[6-9]00)/i.test(st)) f.bold = true;
    if (/font-style\s*:\s*italic/i.test(st)) f.italics = true;
    if (/text-decoration[^;]*underline/i.test(st)) f.underline = true;
    if (/text-decoration[^;]*line-through/i.test(st)) f.strike = true;
    if (/vertical-align\s*:\s*super/i.test(st)) f.superScript = true;
    if (/vertical-align\s*:\s*sub/i.test(st)) f.subScript = true;
    const c = st.match(/(?:^|;)\s*color\s*:\s*#([0-9a-f]{6}|[0-9a-f]{3})\b/i);
    if (c && !fmt.keepColor) {
      let hex = c[1];
      if (hex.length === 3) hex = hex.split('').map((x) => x + x).join('');
      // ignore near-white text that would vanish on paper
      if (!/^f[0-9a-f]f[0-9a-f]f[0-9a-f]$/i.test(hex)) f.color = hex.toUpperCase();
    }
    return f;
  }

  /**
   * ctx.getImage(src) -> Promise<{data:Uint8Array,type:'png'|'jpg'|'gif'|'bmp',width,height}|null>
   */
  class HtmlConverter {
    constructor(ctx) {
      this.ctx = ctx || {};
    }

    /**
     * @param {string} html
     * @param {object} opts
     *   indent: left indent in twips
     *   hanging: hanging indent in twips (first line)
     *   tabStops: docx tab stops for the first paragraph
     *   prefixRuns: runs placed at the start of the first paragraph
     *   run: base run formatting ({bold, italics, color, size})
     *   paragraph: extra paragraph props for every paragraph (spacing, shading, border, keepNext)
     *   emptyText: text used when html is empty (only if prefixRuns)
     */
    async convert(html, opts) {
      opts = opts || {};
      const st = {
        blocks: [],
        runs: null,
        hasContent: false,
        props: null,
        pendingPrefix: (opts.prefixRuns || []).slice(),
        prefixUsed: false,
        opts
      };
      const doc = new DOMParser().parseFromString('<body>' + (html || '') + '</body>', 'text/html');
      await this.walkChildren(doc.body, st, Object.assign({}, opts.run || {}), { indent: opts.indent || 0, listDepth: 0 });
      this.flush(st);
      if (!st.prefixUsed && opts.prefixRuns && opts.prefixRuns.length) {
        // Empty content but we still want the prefix (e.g. "A." with blank answer)
        this.ensure(st, { indent: opts.indent || 0 });
        if (opts.emptyText) st.runs.push(new (D().TextRun)(Object.assign({ text: opts.emptyText }, runProps(opts.run || {}))));
        st.hasContent = true;
        this.flush(st);
      }
      return st.blocks;
    }

    ensure(st, layout) {
      if (st.runs) return;
      st.runs = [];
      st.hasContent = false;
      const first = !st.prefixUsed && st.pendingPrefix.length > 0;
      if (first) {
        st.runs.push(...st.pendingPrefix);
        st.pendingPrefix = [];
        st.prefixUsed = true;
      }
      const o = st.opts;
      const indent = { left: layout.indent || 0 };
      if (layout.hanging) indent.hanging = layout.hanging;
      else if (first && o.hanging) indent.hanging = o.hanging;
      st.props = Object.assign(
        { spacing: { after: 80 } },
        o.paragraph || {},
        { indent },
        first && o.tabStops ? { tabStops: o.tabStops } : layout.tabStops ? { tabStops: layout.tabStops } : {},
        layout.extra || {}
      );
      if (!first && o.hanging && !layout.hanging) {
        // continuation paragraphs of a hanging-indented item line up with its text
        st.props.indent = { left: layout.indent || 0 };
      }
    }

    flush(st) {
      if (!st.runs) return;
      if (st.hasContent) {
        st.blocks.push(new (D().Paragraph)(Object.assign({}, st.props, { children: st.runs })));
      } else if (st.runs.length && st.prefixUsed && st.blocks.length === 0) {
        // keep prefix-only paragraph so it isn't lost
        st.blocks.push(new (D().Paragraph)(Object.assign({}, st.props, { children: st.runs })));
      }
      st.runs = null;
      st.hasContent = false;
    }

    pushText(st, text, fmt, layout) {
      if (!text) return;
      if (!st.runs || !st.hasContent) {
        text = text.replace(/^[ \t\n\r]+/, '');
        if (!text) return;
      }
      this.ensure(st, layout);
      st.runs.push(new (D().TextRun)(Object.assign({ text }, runProps(fmt))));
      st.hasContent = true;
    }

    async walkChildren(node, st, fmt, layout) {
      for (const child of Array.from(node.childNodes)) {
        await this.walk(child, st, fmt, layout);
      }
    }

    async walk(node, st, fmt, layout) {
      if (node.nodeType === 3) {
        const t = node.nodeValue.replace(/[ \t\n\r\f]+/g, ' ');
        this.pushText(st, t, fmt, layout);
        return;
      }
      if (node.nodeType !== 1) return;
      const tag = node.localName;
      if (SKIP_TAGS.has(tag)) {
        if (tag === 'iframe' || tag === 'video' || tag === 'audio' || tag === 'object' || tag === 'embed') {
          const src = node.getAttribute('src') || node.getAttribute('data') || '';
          this.pushText(st, '[Embedded media' + (src ? ': ' + src : '') + ']', Object.assign({}, fmt, { italics: true, color: '7F7F7F' }), layout);
        }
        return;
      }
      fmt = styleFmt(node, fmt);

      switch (tag) {
        case 'br':
          this.ensure(st, layout);
          st.runs.push(new (D().TextRun)({ text: '', break: 1 }));
          st.hasContent = true;
          return;
        case 'b':
        case 'strong':
          return this.walkChildren(node, st, Object.assign({}, fmt, { bold: true }), layout);
        case 'i':
        case 'em':
        case 'cite':
        case 'var':
          return this.walkChildren(node, st, Object.assign({}, fmt, { italics: true }), layout);
        case 'u':
        case 'ins':
          return this.walkChildren(node, st, Object.assign({}, fmt, { underline: true }), layout);
        case 's':
        case 'strike':
        case 'del':
          return this.walkChildren(node, st, Object.assign({}, fmt, { strike: true }), layout);
        case 'sup':
          return this.walkChildren(node, st, Object.assign({}, fmt, { superScript: true }), layout);
        case 'sub':
          return this.walkChildren(node, st, Object.assign({}, fmt, { subScript: true }), layout);
        case 'code':
        case 'tt':
        case 'kbd':
        case 'samp':
          return this.walkChildren(node, st, Object.assign({}, fmt, { font: 'Consolas' }), layout);
        case 'mark':
          return this.walkChildren(node, st, Object.assign({}, fmt, { highlight: 'yellow' }), layout);
        case 'a':
          return this.walkLink(node, st, fmt, layout);
        case 'img':
          return this.walkImage(node, st, fmt, layout);
        case 'math':
          this.pushText(st, ' ' + (node.getAttribute('alttext') || node.textContent || '').trim() + ' ', Object.assign({}, fmt, { font: 'Cambria Math' }), layout);
          return;
        case 'ul':
        case 'ol':
          return this.walkList(node, st, fmt, layout);
        case 'li':
          // li outside a list: treat as a bullet line
          this.flush(st);
          await this.walkListItem(node, st, fmt, layout, '•');
          return;
        case 'table':
          return this.walkTable(node, st, fmt, layout);
        case 'hr':
          this.flush(st);
          st.blocks.push(
            new (D().Paragraph)({
              children: [],
              border: { bottom: { style: D().BorderStyle.SINGLE, size: 6, color: 'BFBFBF', space: 1 } },
              spacing: { after: 120 }
            })
          );
          return;
        default:
          break;
      }

      if (BLOCK_TAGS.has(tag)) {
        this.flush(st);
        let f = fmt;
        let l = layout;
        if (/^h[1-6]$/.test(tag)) f = Object.assign({}, fmt, { bold: true, size: tag === 'h1' ? 30 : tag === 'h2' ? 28 : tag === 'h3' ? 26 : fmt.size });
        if (tag === 'blockquote' || tag === 'dd') l = Object.assign({}, layout, { indent: (layout.indent || 0) + 2 * TWIPS_PER_INDENT });
        if (tag === 'pre') f = Object.assign({}, fmt, { font: 'Consolas' });
        if (tag === 'pre') {
          // keep line breaks inside <pre>
          const lines = (node.textContent || '').split('\n');
          lines.forEach((line, i) => {
            this.ensure(st, l);
            if (i > 0) st.runs.push(new (D().TextRun)({ text: '', break: 1 }));
            st.runs.push(new (D().TextRun)(Object.assign({ text: line }, runProps(f))));
            st.hasContent = true;
          });
          this.flush(st);
          return;
        }
        const align = this.alignOf(node);
        if (align) l = Object.assign({}, l, { extra: Object.assign({}, l.extra || {}, { alignment: align }) });
        await this.walkChildren(node, st, f, l);
        this.flush(st);
        return;
      }

      // Unknown inline element (span, font, label, etc.)
      await this.walkChildren(node, st, fmt, layout);
    }

    alignOf(node) {
      const st = (node.getAttribute('style') || '') + ';align:' + (node.getAttribute('align') || '');
      const A = D().AlignmentType;
      if (/text-align\s*:\s*center|align:center/i.test(st)) return A.CENTER;
      if (/text-align\s*:\s*right|align:right/i.test(st)) return A.RIGHT;
      return null;
    }

    async walkLink(node, st, fmt, layout) {
      const href = node.getAttribute('href') || '';
      const linkFmt = Object.assign({}, fmt, { color: '0563C1', underline: true });
      if (/^https?:\/\//i.test(href) && !node.querySelector('img')) {
        const text = (node.textContent || '').replace(/\s+/g, ' ').trim() || href;
        this.ensure(st, layout);
        st.runs.push(
          new (D().ExternalHyperlink)({
            link: href,
            children: [new (D().TextRun)(Object.assign({ text }, runProps(linkFmt)))]
          })
        );
        st.hasContent = true;
        return;
      }
      await this.walkChildren(node, st, /^https?:/i.test(href) ? linkFmt : fmt, layout);
    }

    async walkImage(node, st, fmt, layout) {
      const src = node.getAttribute('src') || '';
      const alt = (node.getAttribute('alt') || node.getAttribute('title') || '').trim();
      const eq = node.getAttribute('data-equation-content');
      const isEquation = eq || /equation_image/.test(node.className || '') || /\/equation_images\//.test(src);
      if (isEquation) {
        let latex = eq || alt;
        if (!latex) {
          const m = src.match(/equation_images\/([^?]+)/);
          if (m) {
            try {
              latex = decodeURIComponent(decodeURIComponent(m[1]));
            } catch (e) {
              latex = m[1];
            }
          }
        }
        this.pushText(st, ' ' + latex + ' ', Object.assign({}, fmt, { font: 'Cambria Math', italics: true }), layout);
        return;
      }

      let img = null;
      try {
        img = this.ctx.getImage ? await this.ctx.getImage(src) : null;
      } catch (e) {
        img = null;
      }
      if (!img) {
        const label = alt ? '[Image: ' + alt + ']' : '[Image not included in export' + (src && !/^data:/.test(src) ? ': ' + shortSrc(src) : '') + ']';
        this.pushText(st, label, Object.assign({}, fmt, { italics: true, color: '7F7F7F' }), layout);
        return;
      }

      let w = parseFloat(node.getAttribute('width')) || 0;
      let h = parseFloat(node.getAttribute('height')) || 0;
      const sw = (node.getAttribute('style') || '').match(/width\s*:\s*([\d.]+)px/i);
      const sh = (node.getAttribute('style') || '').match(/height\s*:\s*([\d.]+)px/i);
      if (!w && sw) w = parseFloat(sw[1]);
      if (!h && sh) h = parseFloat(sh[1]);
      if (w && !h && img.width) h = (w * img.height) / img.width;
      if (h && !w && img.height) w = (h * img.width) / img.height;
      if (!w || !h) {
        w = img.width || 300;
        h = img.height || 200;
      }
      const maxW = MAX_IMAGE_WIDTH_PX - Math.round(((layout.indent || 0) / 1440) * 96);
      if (w > maxW) {
        h = (h * maxW) / w;
        w = maxW;
      }
      this.ensure(st, layout);
      st.runs.push(
        new (D().ImageRun)({
          type: img.type,
          data: img.data,
          transformation: { width: Math.round(w), height: Math.round(h) },
          altText: { name: (alt || 'image').slice(0, 100), description: alt || 'Question image', title: (alt || 'image').slice(0, 100) }
        })
      );
      st.hasContent = true;
    }

    async walkList(node, st, fmt, layout) {
      this.flush(st);
      const ordered = node.localName === 'ol';
      let n = parseInt(node.getAttribute('start'), 10);
      if (isNaN(n)) n = 1;
      const type = (node.getAttribute('type') || '').trim();
      const depth = (layout.listDepth || 0) + 1;
      const bullets = ['•', '◦', '▪'];
      for (const li of Array.from(node.children)) {
        if (li.localName !== 'li') {
          await this.walk(li, st, fmt, layout);
          continue;
        }
        const marker = ordered ? listMarker(n++, type) : bullets[(depth - 1) % bullets.length];
        await this.walkListItem(li, st, fmt, Object.assign({}, layout, { listDepth: depth }), marker);
      }
    }

    async walkListItem(li, st, fmt, layout, marker) {
      this.flush(st);
      const base = layout.indent || 0;
      const textIndent = base + TWIPS_PER_INDENT * 2;
      const itemLayout = Object.assign({}, layout, {
        indent: textIndent,
        hanging: TWIPS_PER_INDENT,
        tabStops: [{ type: D().TabStopType.LEFT, position: textIndent }]
      });
      this.ensure(st, itemLayout);
      st.runs.push(new (D().TextRun)(Object.assign({ text: marker }, runProps(fmt))));
      st.runs.push(new (D().TextRun)({ children: [new (D().Tab)()] }));
      const childLayout = Object.assign({}, layout, { indent: textIndent, hanging: 0, tabStops: null });
      // Walk children; first paragraph continues the bullet line
      for (const child of Array.from(li.childNodes)) {
        const isBlock = child.nodeType === 1 && (BLOCK_TAGS.has(child.localName) || child.localName === 'ul' || child.localName === 'ol' || child.localName === 'table');
        if (isBlock && st.runs && !st.hasContent && (child.localName === 'p' || child.localName === 'div')) {
          // <li><p>text</p></li>: keep text on the bullet line
          await this.walkChildren(child, st, styleFmt(child, fmt), childLayout);
          this.flush(st);
          continue;
        }
        await this.walk(child, st, fmt, childLayout);
      }
      if (st.runs && !st.hasContent) st.hasContent = true; // keep empty bullet
      this.flush(st);
    }

    async walkTable(node, st, fmt, layout) {
      this.flush(st);
      const D_ = D();
      const rowsEl = [];
      for (const part of Array.from(node.children)) {
        if (part.localName === 'tr') rowsEl.push(part);
        else if (/^(thead|tbody|tfoot)$/.test(part.localName)) {
          for (const tr of Array.from(part.children)) if (tr.localName === 'tr') rowsEl.push(tr);
        } else if (part.localName === 'caption') {
          await this.walk(part, st, Object.assign({}, fmt, { bold: true }), layout);
          this.flush(st);
        }
      }
      if (!rowsEl.length) return;
      let cols = 0;
      for (const tr of rowsEl) {
        let c = 0;
        for (const td of Array.from(tr.children)) if (/^t[dh]$/.test(td.localName)) c += parseInt(td.getAttribute('colspan'), 10) || 1;
        cols = Math.max(cols, c);
      }
      cols = Math.max(cols, 1);
      const total = 9360 - (layout.indent || 0);
      const colW = Math.floor(total / cols);
      const sub = new HtmlConverter(this.ctx);
      const rows = [];
      for (const tr of rowsEl) {
        const cells = [];
        for (const td of Array.from(tr.children)) {
          if (!/^t[dh]$/.test(td.localName)) continue;
          const span = parseInt(td.getAttribute('colspan'), 10) || 1;
          const cellFmt = td.localName === 'th' ? Object.assign({}, fmt, { bold: true }) : fmt;
          let children = await sub.convert(td.innerHTML, { run: cellFmt, paragraph: { spacing: { after: 40 } } });
          if (!children.length) children = [new D_.Paragraph({ children: [] })];
          cells.push(
            new D_.TableCell({
              children,
              columnSpan: span > 1 ? span : undefined,
              width: { size: colW * span, type: D_.WidthType.DXA },
              margins: { top: 60, bottom: 60, left: 100, right: 100 },
              shading: td.localName === 'th' ? { type: D_.ShadingType.CLEAR, color: 'auto', fill: 'F2F2F2' } : undefined
            })
          );
        }
        if (cells.length) rows.push(new D_.TableRow({ children: cells }));
      }
      if (!rows.length) return;
      st.blocks.push(
        new D_.Table({
          rows,
          width: { size: total, type: D_.WidthType.DXA },
          columnWidths: new Array(cols).fill(colW),
          indent: layout.indent ? { size: layout.indent, type: D_.WidthType.DXA } : undefined
        })
      );
      // spacing paragraph after a table
      st.blocks.push(new D_.Paragraph({ children: [], spacing: { after: 60 } }));
    }
  }

  function shortSrc(src) {
    const s = src.replace(/^\$IMS[-_]CC[-_]FILEBASE\$\//i, '').split('?')[0];
    try {
      return decodeURIComponent(s).slice(-80);
    } catch (e) {
      return s.slice(-80);
    }
  }

  function listMarker(n, type) {
    if (type === 'a') return toAlpha(n).toLowerCase() + '.';
    if (type === 'A') return toAlpha(n) + '.';
    if (type === 'i') return toRoman(n).toLowerCase() + '.';
    if (type === 'I') return toRoman(n) + '.';
    return n + '.';
  }
  function toAlpha(n) {
    let s = '';
    while (n > 0) {
      n--;
      s = String.fromCharCode(65 + (n % 26)) + s;
      n = Math.floor(n / 26);
    }
    return s;
  }
  function toRoman(n) {
    const map = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
    let s = '';
    for (const [v, r] of map) while (n >= v) {
      s += r;
      n -= v;
    }
    return s;
  }

  CQE.HtmlConverter = HtmlConverter;
  CQE.toAlpha = toAlpha;
})(typeof window !== 'undefined' ? window : globalThis);
