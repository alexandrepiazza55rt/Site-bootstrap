// Renderização WordprocessingML -> HTML (unidades em pt), fiel ao layout do Word.
import { NS, isW, wKid, wKids, wAttr, wNum, escapeHtml as esc } from './xml.js';
import { parseParagraph } from './segments.js';
import {
  paraBaseProps, runBaseProps, runProps, applyPPr, applyRPr, readBorder, readShd, cssFont, lineFactor, isMono,
} from './styles.js';

const pt = (twips) => +(twips / 20).toFixed(2) + 'pt';
const emuPt = (emu) => +(emu / 12700).toFixed(2);

const HIGHLIGHT = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000',
  darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
};

/** Propriedades da seção (último w:sectPr do corpo). */
export function sectionProps(doc) {
  const body = doc.getElementsByTagNameNS(NS.w, 'body')[0];
  const sect = body ? wKid(body, 'sectPr') : null;
  const pgSz = wKid(sect, 'pgSz');
  const pgMar = wKid(sect, 'pgMar');
  const s = {
    w: wNum(pgSz, 'w', 11906), h: wNum(pgSz, 'h', 16838),
    top: wNum(pgMar, 'top', 1417), right: wNum(pgMar, 'right', 1701), bottom: wNum(pgMar, 'bottom', 1417),
    left: wNum(pgMar, 'left', 1701), header: wNum(pgMar, 'header', 708), footer: wNum(pgMar, 'footer', 708),
    headers: {}, footers: {}, titlePg: !!wKid(sect, 'titlePg'),
  };
  if (wAttr(pgSz, 'orient') === 'landscape' && s.w < s.h) [s.w, s.h] = [s.h, s.w];
  s.top = Math.abs(s.top); s.bottom = Math.abs(s.bottom);
  for (const ref of wKids(sect, 'headerReference')) s.headers[wAttr(ref, 'type') || 'default'] = ref.getAttributeNS(NS.r, 'id');
  for (const ref of wKids(sect, 'footerReference')) s.footers[wAttr(ref, 'type') || 'default'] = ref.getAttributeNS(NS.r, 'id');
  s.contentW = s.w - s.left - s.right;
  return s;
}

/**
 * Renderizador. opts:
 *  - resolve(id) -> {text, empty, label}   valor de um campo {{id}}
 *  - mode: 'screen' | 'print'
 *  - marks: Map(node -> 'edited'|'new')
 */
export class Renderer {
  constructor(ctx, opts = {}) {
    this.ctx = ctx;
    this.styles = ctx.styles;
    this.numbering = ctx.numbering;
    this.opts = opts;
    this.nodes = [];
  }

  nodeIndex(node) {
    this.nodes.push(node);
    return this.nodes.length - 1;
  }

  /** Renderiza o corpo e devolve elementos de bloco (para paginação). */
  renderBody(doc, section) {
    this.section = section || sectionProps(doc);
    const body = doc.getElementsByTagNameNS(NS.w, 'body')[0];
    const env = { counter: this.numbering.createCounter(), rels: this.ctx.mainRels, prev: null, contentW: this.section.contentW };
    const html = [];
    this.renderChildren(body, env, html, true);
    return htmlToBlocks(html.join(''));
  }

  /** Renderiza cabeçalho/rodapé (HTML string). */
  renderPart(part, section) {
    this.section = section;
    const env = { counter: this.numbering.createCounter(), rels: part.rels, prev: null, inHeader: true, contentW: section.contentW };
    const html = [];
    this.renderChildren(part.doc.documentElement, env, html, false);
    return html.join('');
  }

  renderChildren(container, env, html, top) {
    for (let c = container.firstElementChild; c; c = c.nextElementSibling) {
      if (isW(c, 'p')) this.renderParagraph(c, env, html, top);
      else if (isW(c, 'tbl')) this.renderTable(c, env, html, top);
      else if (isW(c, 'sdt')) { const sc = wKid(c, 'sdtContent'); if (sc) this.renderChildren(sc, env, html, top); }
    }
  }

  // ------------------------------------------------------- parágrafo
  paraProps(p, env) {
    const pPr = wKid(p, 'pPr');
    const pp = paraBaseProps(this.styles, pPr, env.tableStyleId);
    const np = pPr && wKid(pPr, 'numPr');
    let numId = pp.numId, ilvl = pp.ilvl;
    if (np) {
      const d = wAttr(wKid(np, 'numId'), 'val');
      const l = wAttr(wKid(np, 'ilvl'), 'val');
      if (d !== null) numId = d;
      if (l !== null) ilvl = parseInt(l, 10) || 0;
    }
    let lvl = null;
    if (numId && numId !== '0') {
      lvl = this.numbering.level(numId, ilvl);
      if (lvl) applyPPr(pp, lvl.pPr);
    }
    applyPPr(pp, pPr);
    pp.numId = numId; pp.ilvl = ilvl; pp.lvl = lvl;
    // run "marca de parágrafo"
    const base = runBaseProps(this.styles, pp.styleId, env.tableStyleId);
    pp.runBase = base;
    pp.mark = runProps(this.styles, base, pPr && wKid(pPr, 'rPr'));
    return pp;
  }

  renderParagraph(p, env, html, top) {
    const unit = this.opts.listUnits && this.opts.listUnits.get(p);
    if (!unit) return this.renderParagraphOnce(p, env, html, top);
    for (let i = 0; i < unit.count; i++) {
      env.itemIdx = i;
      this.renderParagraphOnce(p, env, html, top, i);
    }
    env.itemIdx = undefined;
  }

  renderParagraphOnce(p, env, html, top, itemIdx) {
    const pp = this.paraProps(p, env);
    const segs = parseParagraph(p);
    const xi = this.nodeIndex(p);

    // numeração
    let numHtml = '';
    if (pp.lvl) {
      const r = env.counter.next(pp.numId, pp.ilvl);
      if (r && r.text !== '') {
        const rp = { ...pp.mark };
        applyRPr(rp, r.lvl.rPr, this.styles);
        const st = runCss(rp, pp);
        numHtml = `<span class="w-num" style="${st}">${esc(r.text)}</span>`;
        if (r.lvl.suff === 'tab') numHtml += `<span class="w-tab w-numtab" style="${st}"></span>`;
        else if (r.lvl.suff === 'space') numHtml += `<span class="w-sfx" style="${st};width:${isMono(rp.font) ? 0.6 : 0.28}em"></span>`;
      }
    }

    // espaçamento contextual ("não adicionar espaço entre parágrafos do mesmo estilo")
    let before = pp.beforeAuto ? 280 : pp.before;
    const after = pp.afterAuto ? 280 : pp.after;
    const sameStyle = env.prev && env.prev.styleId === pp.styleId;
    if (pp.contextualSpacing && sameStyle) before = 0;
    if (sameStyle && env.prev.contextualSpacing && env.prevHtml >= 0 && html[env.prevHtml]) {
      html[env.prevHtml] = html[env.prevHtml].replace(/padding-bottom:[\d.]+pt/, 'padding-bottom:0pt');
    }

    const parts = [[]];
    const fieldState = { depth: 0, instr: '', kind: null, phase: null };
    segs.forEach((s, i) => {
      if (s.k === 'br' && s.type === 'page') { parts.push([]); return; }
      parts[parts.length - 1].push(this.renderSeg(s, i, pp, env, fieldState));
    });

    const mark = pp.mark;
    const outerStyle = [`padding-top:${pt(before)}`, `padding-bottom:${pt(after)}`];
    const inner = innerCss(pp, mark);
    const cls = ['w-p'];
    if (pp.keepNext) cls.push('keep-next');
    if (pp.keepLines) cls.push('keep-lines');
    if (pp.lvl) cls.push('w-lvl' + pp.ilvl);
    const mk = this.opts.marks && this.opts.marks.get(p);
    if (mk) cls.push('chg-' + mk);
    const tabsAttr = pp.tabs.length ? ` data-tabs="${pp.tabs.map((t) => `${t.pos},${t.val},${t.leader}`).join(';')}"` : '';
    const dataAttrs = `data-xi="${xi}"${itemIdx !== undefined ? ` data-item="${itemIdx}"` : ''}${pp.widowControl ? ' data-widow="1"' : ''}${pp.pageBreakBefore ? ' data-pbb="1"' : ''}`;
    const innerAttrs = `class="w-pi" style="${inner}" data-il="${pp.indLeft}" data-fl="${pp.firstLine}" data-jc="${pp.jc}" data-dt="${this.ctx.defaultTab}"${tabsAttr}`;

    parts.forEach((partRuns, k) => {
      let content = partRuns.join('');
      const isEmpty = !content.replace(/<span class="w-anchor[\s\S]*?<\/span>/g, '').replace(/<[^>]+>/g, '');
      const hasVisual = /w-tab|<img|w-shape|w-sfx/.test(content);
      if (isEmpty && !hasVisual) content += `<span class="w-empty" style="${runCss(mark, pp)}">​</span>`;
      const extra = k > 0 ? ' data-pbb="1"' : '';
      const ostyle = outerStyle.slice();
      if (parts.length > 1) {
        if (k > 0) ostyle[0] = 'padding-top:0';
        if (k < parts.length - 1) ostyle[1] = 'padding-bottom:0';
      }
      html.push(`<div class="${cls.join(' ')}${top ? ' w-block' : ''}" ${dataAttrs}${extra} style="${ostyle.join(';')}"><div ${innerAttrs}>${k === 0 ? numHtml : ''}${content}</div></div>`);
    });
    env.prev = pp;
    env.prevHtml = html.length - 1;
  }

  renderSeg(s, i, pp, env, fs) {
    const rp = s.rPr !== undefined ? runProps(this.styles, pp.runBase, s.rPr) : pp.mark;
    if (rp.vanish) return '';
    if (s.k === 't') {
      if (fs.phase === 'instr') return '';
      const extra = fs.phase === 'result' && fs.kind ? ` data-pagefield="${fs.kind}"` : '';
      return `<span class="w-r" data-s="${i}" data-o="0"${extra} style="${runCss(rp, pp)}">${esc(s.text)}</span>`;
    }
    if (s.k === 'f') {
      const res = this.opts.resolve ? this.opts.resolve(s.id, env.itemIdx) : { text: '{{' + s.id + '}}', empty: false };
      const screen = this.opts.mode !== 'print';
      let text = res.text;
      let cls = 'w-r w-fld';
      if (res.empty) {
        cls += ' is-empty';
        if (screen) text = res.label || s.id;
      }
      if (res.computed) cls += ' is-computed';
      const body = esc(text || '').replace(/\n/g, '<br>');
      if (!screen) return `<span class="w-r" style="${runCss(rp, pp)}">${body}</span>`;
      return `<span class="${cls}" data-s="${i}" data-f="${esc(s.id)}" style="${runCss(rp, pp)}">${body || '​'}</span>`;
    }
    if (s.k === 'tab') {
      return `<span class="w-tab" data-s="${i}" style="${runCss(rp, pp)}${rp.u ? ';border-bottom:1px solid currentColor' : ''}"></span>`;
    }
    if (s.k === 'br') return '<br>';
    if (s.k === 'raw') return this.renderRaw(s.node, rp, pp, env, fs, i);
    return '';
  }

  renderRaw(node, rp, pp, env, fs, i) {
    if (!isW(node)) {
      // m:oMath e afins: apenas o texto
      const t = node.textContent || '';
      return t ? `<span class="w-r" style="${runCss(rp, pp)}">${esc(t)}</span>` : '';
    }
    const n = node.localName;
    if (n === 'r') {
      let out = '';
      for (let c = node.firstElementChild; c; c = c.nextElementSibling) {
        if (isW(c, 'fldChar')) {
          const t = wAttr(c, 'fldCharType');
          if (t === 'begin') { fs.depth++; fs.phase = 'instr'; fs.instr = ''; fs.kind = null; }
          else if (t === 'separate') {
            fs.phase = 'result';
            fs.kind = /\bNUMPAGES\b/.test(fs.instr) ? 'numpages' : /\bPAGE\b/.test(fs.instr) ? 'page' : null;
          } else if (t === 'end') { fs.depth = Math.max(0, fs.depth - 1); fs.phase = fs.depth ? 'result' : null; fs.kind = null; }
        } else if (isW(c, 'instrText')) fs.instr += c.textContent;
        else if (isW(c, 'drawing')) out += this.renderDrawing(c, env, i);
        else if (isW(c, 'pict')) out += this.renderVml(c, env, i);
        else if (c.namespaceURI === NS.mc && c.localName === 'AlternateContent') {
          const choice = c.getElementsByTagNameNS(NS.mc, 'Choice')[0];
          const fb = c.getElementsByTagNameNS(NS.mc, 'Fallback')[0];
          const d = choice && choice.getElementsByTagNameNS(NS.w, 'drawing')[0];
          if (d) out += this.renderDrawing(d, env, i);
          else if (fb) { const pict = fb.getElementsByTagNameNS(NS.w, 'pict')[0]; if (pict) out += this.renderVml(pict, env, i); }
        } else if (isW(c, 'sym')) {
          const code = parseInt(wAttr(c, 'char') || '0', 16);
          const ch = code >= 0xf000 ? '•' : String.fromCharCode(code);
          out += `<span class="w-r" style="${runCss(rp, pp)}">${esc(ch)}</span>`;
        } else if (isW(c, 'ptab')) out += `<span class="w-tab" style="${runCss(rp, pp)}"></span>`;
        else if (isW(c, 't')) out += `<span class="w-r" style="${runCss(rp, pp)}">${esc(c.textContent)}</span>`;
      }
      return out;
    }
    if (n === 'fldSimple') {
      const instr = wAttr(node, 'instr') || '';
      const kind = /\bNUMPAGES\b/.test(instr) ? 'numpages' : /\bPAGE\b/.test(instr) ? 'page' : null;
      const t = node.textContent || '';
      return `<span class="w-r"${kind ? ` data-pagefield="${kind}"` : ''} style="${runCss(rp, pp)}">${esc(t)}</span>`;
    }
    return '';
  }

  // ---------------------------------------------------------- imagens
  renderDrawing(d, env, segIndex) {
    const inline = d.getElementsByTagNameNS(NS.wp, 'inline')[0];
    const anchor = d.getElementsByTagNameNS(NS.wp, 'anchor')[0];
    const box = inline || anchor;
    if (!box) return '';
    const ext = box.getElementsByTagNameNS(NS.wp, 'extent')[0];
    const w = emuPt(+(ext && ext.getAttribute('cx')) || 0);
    const h = emuPt(+(ext && ext.getAttribute('cy')) || 0);
    const blip = box.getElementsByTagNameNS(NS.a, 'blip')[0];
    let content = '';
    if (blip) {
      const url = this.ctx.mediaUrl(env.rels, blip.getAttributeNS(NS.r, 'embed') || blip.getAttributeNS(NS.r, 'link'));
      if (url) content = `<img src="${url}" alt="" draggable="false" style="width:${w}pt;height:${h}pt">`;
    } else {
      // forma simples (linha/retângulo preenchido)
      const fill = box.getElementsByTagNameNS(NS.a, 'solidFill')[0];
      const clr = fill && fill.getElementsByTagNameNS(NS.a, 'srgbClr')[0];
      const color = clr ? '#' + clr.getAttribute('val') : null;
      const txbx = box.getElementsByTagNameNS(NS.wps, 'txbx')[0];
      if (txbx) {
        const tc = txbx.getElementsByTagNameNS(NS.w, 'txbxContent')[0];
        const inner = [];
        if (tc) this.renderChildren(tc, { ...env, prev: null }, inner, false);
        content = `<div class="w-txbx" style="width:${w}pt;min-height:${h}pt;${color ? 'background:' + color : ''}">${inner.join('')}</div>`;
      } else if (color) content = `<span class="w-shape" style="display:block;width:${w}pt;height:${Math.max(h, 0.5)}pt;background:${color}"></span>`;
      else return '';
    }
    if (inline) return `<span class="w-img" data-s="${segIndex}" style="display:inline-block;vertical-align:bottom;width:${w}pt;height:${h}pt">${content}</span>`;

    // flutuante
    const behind = anchor.getAttribute('behindDoc') === '1';
    const ph = anchor.getElementsByTagNameNS(NS.wp, 'positionH')[0];
    const pv = anchor.getElementsByTagNameNS(NS.wp, 'positionV')[0];
    const sec = this.section;
    const hRel = ph ? ph.getAttribute('relativeFrom') : 'column';
    const vRel = pv ? pv.getAttribute('relativeFrom') : 'paragraph';
    const hOff = posOf(ph);
    const vOff = posOf(pv);
    const hAlign = alignOf(ph);
    const contentW = sec.contentW / 20;
    let left = 0;
    if (hAlign) {
      const area = hRel === 'page' ? sec.w / 20 : contentW;
      const base = hRel === 'page' ? -sec.left / 20 : 0;
      left = base + (hAlign === 'center' ? (area - w) / 2 : hAlign === 'right' ? area - w : 0);
    } else left = (hRel === 'page' ? hOff - sec.left / 20 : hOff);
    let style = `position:absolute;left:${left.toFixed(2)}pt;width:${w}pt;height:${h}pt;z-index:${behind ? 0 : 3}`;
    let extra = '';
    if (vRel === 'page' || vRel === 'margin' || vRel === 'topMargin') {
      const abs = vRel === 'page' ? vOff : vOff + sec.top / 20;
      extra = ` data-abs-v="${abs}"`;
      style += ';top:0';
    } else style += `;top:${vOff}pt`;
    return `<span class="w-anchor${behind ? ' behind' : ''}"${extra} style="${style}">${content}</span>`;
  }

  renderVml(pict, env, segIndex) {
    const shape = pict.getElementsByTagNameNS(NS.v, 'shape')[0];
    const imd = pict.getElementsByTagNameNS(NS.v, 'imagedata')[0];
    if (!shape || !imd) return '';
    const st = shape.getAttribute('style') || '';
    const get = (k) => { const m = new RegExp('(?:^|;)\\s*' + k + '\\s*:\\s*([\\d.]+)(pt|in|px)?').exec(st); if (!m) return 0; const v = parseFloat(m[1]); return m[2] === 'in' ? v * 72 : m[2] === 'px' ? v * 0.75 : v; };
    const w = get('width'), h = get('height');
    const url = this.ctx.mediaUrl(env.rels, imd.getAttributeNS(NS.r, 'id'));
    if (!url) return '';
    return `<span class="w-img" data-s="${segIndex}" style="display:inline-block;width:${w}pt;height:${h}pt"><img src="${url}" alt="" style="width:${w}pt;height:${h}pt"></span>`;
  }

  // ----------------------------------------------------------- tabela
  renderTable(tbl, env, html, top) {
    const xi = this.nodeIndex(tbl);
    const tblPr = wKid(tbl, 'tblPr');
    const styleId = wAttr(wKid(tblPr, 'tblStyle'), 'val') || this.styles.defaultTable;
    const chain = this.styles.chain(styleId);
    const borders = {};
    const mar = { top: 0, left: 108, bottom: 0, right: 108 };
    let ind = 0, jc = null;
    const applyTblPr = (pr) => {
      if (!pr) return;
      const tb = wKid(pr, 'tblBorders');
      if (tb) for (const b of wKids(tb)) borders[normSide(b.localName)] = readBorder(b);
      const cm = wKid(pr, 'tblCellMar');
      if (cm) for (const m of wKids(cm)) mar[normSide(m.localName)] = wNum(m, 'w', 0);
      const ti = wKid(pr, 'tblInd'); if (ti) ind = wNum(ti, 'w', 0);
      const j = wKid(pr, 'jc'); if (j) jc = wAttr(j, 'val');
    };
    for (const s of chain) applyTblPr(s.tblPr);
    applyTblPr(tblPr);

    const grid = wKids(wKid(tbl, 'tblGrid'), 'gridCol').map((g) => wNum(g, 'w', 0));
    const rows = [];
    const rowItem = [];
    for (const tr of wKids(tbl, 'tr')) {
      const unit = this.opts.listUnits && this.opts.listUnits.get(tr);
      const n = unit ? unit.count : 1;
      for (let i = 0; i < n; i++) { rows.push(tr); rowItem.push(unit ? i : undefined); }
    }
    // matriz de vMerge
    const layout = [];
    rows.forEach((tr, ri) => {
      let col = wNum(wKid(wKid(tr, 'trPr'), 'gridBefore'), 'val', 0);
      const cells = [];
      for (const tc of wKids(tr, 'tc')) {
        const tcPr = wKid(tc, 'tcPr');
        const span = wNum(wKid(tcPr, 'gridSpan'), 'val', 1) || 1;
        const vm = wKid(tcPr, 'vMerge');
        const vMerge = vm ? (wAttr(vm, 'val') === 'restart' ? 'restart' : 'continue') : null;
        cells.push({ tc, tcPr, col, span, vMerge, rowspan: 1 });
        col += span;
      }
      layout.push(cells);
      void ri;
    });
    for (let r = 0; r < layout.length; r++) {
      for (const cell of layout[r]) {
        if (cell.vMerge !== 'restart') continue;
        for (let k = r + 1; k < layout.length; k++) {
          const below = layout[k].find((c) => c.col === cell.col);
          if (below && below.vMerge === 'continue') { cell.rowspan++; below.skip = true; } else break;
        }
      }
    }
    const nCols = Math.max(grid.length, ...layout.map((cells) => cells.reduce((a, c) => Math.max(a, c.col + c.span), 0)));
    const totalW = grid.reduce((a, b) => a + b, 0);
    const tStyle = [`width:${pt(totalW)}`];
    if (jc === 'center') tStyle.push('margin-left:auto;margin-right:auto');
    else if (jc === 'right' || jc === 'end') tStyle.push('margin-left:auto');
    else tStyle.push(`margin-left:${pt(ind - (mar.left || 0) * 0)}`);
    const out = [`<div class="w-tblwrap${top ? ' w-block' : ''}" data-xi="${xi}"><table class="w-tbl" style="${tStyle.join(';')}"><colgroup>`];
    for (let c = 0; c < nCols; c++) out.push(`<col style="width:${pt(grid[c] || 0)}">`);
    out.push('</colgroup><tbody>');
    const cellEnv = { ...env, tableStyleId: styleId, prev: null };
    rows.forEach((tr, ri) => {
      const trPr = wKid(tr, 'trPr');
      const trH = wKid(trPr, 'trHeight');
      const hVal = wNum(trH, 'val', 0);
      const header = !!wKid(trPr, 'tblHeader') && onOffEl(wKid(trPr, 'tblHeader'));
      const cant = !!wKid(trPr, 'cantSplit');
      const rxi = this.nodeIndex(tr);
      cellEnv.itemIdx = rowItem[ri];
      out.push(`<tr data-xi="${rxi}"${rowItem[ri] !== undefined ? ` data-item="${rowItem[ri]}"` : ''}${header ? ' data-header="1"' : ''}${cant ? ' data-cant="1"' : ''}${hVal ? ` style="height:${pt(hVal)}"` : ''}>`);
      for (const cell of layout[ri]) {
        if (cell.skip) continue;
        const { tcPr } = cell;
        const cs = [];
        const tcb = wKid(tcPr, 'tcBorders');
        const own = {};
        if (tcb) for (const b of wKids(tcb)) own[normSide(b.localName)] = readBorder(b);
        const lastRow = ri + cell.rowspan - 1 === rows.length - 1;
        const edge = {
          top: 'top' in own ? own.top : (ri === 0 ? borders.top : borders.insideH),
          bottom: 'bottom' in own ? own.bottom : (lastRow ? borders.bottom : borders.insideH),
          left: 'left' in own ? own.left : (cell.col === 0 ? borders.left : borders.insideV),
          right: 'right' in own ? own.right : (cell.col + cell.span >= nCols ? borders.right : borders.insideV),
        };
        for (const side of ['top', 'right', 'bottom', 'left']) cs.push(`border-${side}:${borderCss(edge[side])}`);
        const cm = { ...mar };
        const tcm = wKid(tcPr, 'tcMar');
        if (tcm) for (const m of wKids(tcm)) cm[normSide(m.localName)] = wNum(m, 'w', 0);
        cs.push(`padding:${pt(cm.top || 0)} ${pt(cm.right || 0)} ${pt(cm.bottom || 0)} ${pt(cm.left || 0)}`);
        const va = wAttr(wKid(tcPr, 'vAlign'), 'val');
        cs.push(`vertical-align:${va === 'center' ? 'middle' : va === 'bottom' ? 'bottom' : 'top'}`);
        const shdEl = wKid(tcPr, 'shd');
        const shd = shdEl ? readShd(shdEl) : null;
        if (shd) cs.push(`background:${shd}`);
        const attrs = `${cell.span > 1 ? ` colspan="${cell.span}"` : ''}${cell.rowspan > 1 ? ` rowspan="${cell.rowspan}"` : ''}`;
        const inner = [];
        this.renderChildren(cell.tc, cellEnv, inner, false);
        cellEnv.prev = null;
        out.push(`<td${attrs} style="${cs.join(';')}">${inner.join('')}</td>`);
      }
      out.push('</tr>');
    });
    out.push('</tbody></table></div>');
    html.push(out.join(''));
    env.prev = null;
  }
}

function onOffEl(el) {
  const v = wAttr(el, 'val');
  return v === null || !(v === '0' || v === 'false' || v === 'off');
}

function normSide(s) {
  return s === 'start' ? 'left' : s === 'end' ? 'right' : s;
}

function borderCss(b) {
  if (!b) return 'none';
  const w = Math.max(0.5, b.sz / 8);
  const style = b.val === 'double' ? 'double' : /dot/.test(b.val) ? 'dotted' : /dash/.test(b.val) ? 'dashed' : 'solid';
  return `${b.val === 'double' ? Math.max(2.25, w * 3) : w}pt ${style} ${b.color}`;
}

function posOf(el) {
  if (!el) return 0;
  const o = el.getElementsByTagNameNS(NS.wp, 'posOffset')[0];
  return o ? emuPt(+o.textContent || 0) : 0;
}

function alignOf(el) {
  if (!el) return null;
  const a = el.getElementsByTagNameNS(NS.wp, 'align')[0];
  return a ? a.textContent : null;
}

function innerCss(pp, mark) {
  const st = [];
  if (pp.indLeft < 0) st.push(`margin-left:${pt(pp.indLeft)}`); else st.push(`padding-left:${pt(pp.indLeft)}`);
  if (pp.indRight < 0) st.push(`margin-right:${pt(pp.indRight)}`); else st.push(`padding-right:${pt(pp.indRight)}`);
  if (pp.firstLine) st.push(`text-indent:${pt(pp.firstLine)}`);
  const jc = pp.jc;
  st.push(`text-align:${jc === 'both' || jc === 'distribute' || jc === 'thaiDistribute' || jc === 'lowKashida' || jc === 'mediumKashida' || jc === 'highKashida' ? 'justify' : jc === 'center' ? 'center' : jc === 'right' ? 'right' : 'left'}`);
  if (jc === 'distribute') st.push('text-align-last:justify');
  st.push(lineCss(pp, mark));
  st.push(`font-family:${cssFont(mark.font)}`, `font-size:${mark.sz / 2}pt`);
  const b = pp.borders || {};
  for (const side of ['top', 'bottom', 'left', 'right']) {
    if (b[side]) {
      st.push(`border-${side}:${borderCss(b[side])}`);
      st.push(`padding-${side === 'left' || side === 'right' ? side : side}:${b[side].space}pt`);
    }
  }
  if (b.left && pp.indLeft >= 0) st.push(`padding-left:${pt(pp.indLeft)}`);
  if (pp.shd) st.push(`background:${pp.shd}`);
  return st.join(';');
}

function lineCss(pp, rp) {
  if (pp.lineRule === 'exact') return `line-height:${pt(pp.line)}`;
  if (pp.lineRule === 'atLeast') {
    const single = (rp.sz / 2) * lineFactor(rp.font);
    return `line-height:${Math.max(single, pp.line / 20).toFixed(2)}pt`;
  }
  return `line-height:${((pp.line / 240) * lineFactor(rp.font)).toFixed(4)}`;
}

export function runCss(rp, pp) {
  const st = [`font-family:${cssFont(rp.font)}`, `font-size:${rp.sz / 2}pt`];
  if (pp && (pp.lineRule === 'auto' || !pp.lineRule)) st.push(`line-height:${((pp.line / 240) * lineFactor(rp.font)).toFixed(4)}`);
  st.push(`font-weight:${rp.b ? 700 : 400}`, `font-style:${rp.i ? 'italic' : 'normal'}`);
  const deco = [];
  if (rp.u) deco.push('underline');
  if (rp.strike || rp.dstrike) deco.push('line-through');
  if (deco.length) {
    st.push(`text-decoration-line:${deco.join(' ')}`);
    if (rp.u === 'double' || rp.dstrike) st.push('text-decoration-style:double');
    else if (rp.u && /dot/.test(rp.u)) st.push('text-decoration-style:dotted');
    else if (rp.u && /dash/.test(rp.u)) st.push('text-decoration-style:dashed');
    else if (rp.u && /wav/.test(rp.u)) st.push('text-decoration-style:wavy');
    if (rp.u === 'thick') st.push('text-decoration-thickness:2px');
  } else st.push('text-decoration-line:none');
  if (rp.color) st.push(`color:${rp.color}`);
  const bg = rp.highlight ? HIGHLIGHT[rp.highlight] : rp.shd;
  if (bg) st.push(`background-color:${bg}`);
  if (rp.caps) st.push('text-transform:uppercase');
  if (rp.smallCaps) st.push('font-variant:small-caps');
  if (rp.vertAlign === 'superscript') st.push('vertical-align:super;font-size:' + (rp.sz / 2) * 0.65 + 'pt');
  if (rp.vertAlign === 'subscript') st.push('vertical-align:sub;font-size:' + (rp.sz / 2) * 0.65 + 'pt');
  if (rp.spacing) st.push(`letter-spacing:${pt(rp.spacing)}`);
  return st.join(';');
}

export function htmlToBlocks(html) {
  const tpl = document.createElement('template');
  tpl.innerHTML = html;
  return Array.from(tpl.content.children);
}

// ------------------------------------------------------ tabulações
/** Calcula a largura das tabulações (precisa estar no DOM, sem zoom). */
export function layoutTabs(root) {
  const inners = root.querySelectorAll('.w-pi');
  for (const inner of inners) {
    const tabs = inner.querySelectorAll(':scope .w-tab');
    if (!tabs.length) continue;
    const outer = inner.parentElement;
    const originX = outer.getBoundingClientRect().left;
    const scale = outer.getBoundingClientRect().width / (outer.offsetWidth || 1) || 1;
    const twipPerPx = 15 / scale;
    const il = +inner.dataset.il || 0;
    const fl = +inner.dataset.fl || 0;
    const dt = +inner.dataset.dt || 720;
    const stops = (inner.dataset.tabs || '').split(';').filter(Boolean).map((s) => {
      const [pos, val, leader] = s.split(',');
      return { pos: +pos, val, leader };
    }).filter((t) => t.val !== 'bar' && t.val !== 'clear');
    const rightEdge = (inner.getBoundingClientRect().right - originX) * twipPerPx - (parseFloat(getComputedStyle(inner).paddingRight) || 0) * 15;
    for (const tab of tabs) {
      tab.style.width = '0px';
      const x = (tab.getBoundingClientRect().left - originX) * twipPerPx;
      let stop = null;
      for (const s of stops) if (s.pos > x + 2) { stop = s; break; }
      if (fl < 0 && il > x + 2 && (!stop || il < stop.pos)) stop = { pos: il, val: 'left', leader: 'none' };
      if (!stop) {
        const pos = (Math.floor((x + 2) / dt) + 1) * dt;
        stop = { pos: Math.min(pos, Math.max(x, rightEdge)), val: 'left', leader: 'none' };
      }
      let w = stop.pos - x;
      if (stop.val === 'right' || stop.val === 'end' || stop.val === 'center' || stop.val === 'decimal') {
        const follow = followingWidth(tab) * twipPerPx;
        w = stop.val === 'center' ? w - follow / 2 : w - follow;
      }
      w = Math.max(0, w);
      tab.style.width = (w / twipPerPx).toFixed(2) + 'px';
      if (stop.leader && stop.leader !== 'none') {
        tab.style.borderBottom = stop.leader === 'dot' ? '1.5px dotted currentColor' : stop.leader === 'hyphen' ? '1px dashed currentColor' : '1px solid currentColor';
        tab.classList.add('lead');
      }
    }
  }
}

function followingWidth(tab) {
  let w = 0;
  for (let n = tab.nextSibling; n; n = n.nextSibling) {
    if (n.nodeType === 1 && n.classList.contains('w-tab')) break;
    if (n.nodeType === 1) w += n.getBoundingClientRect().width;
  }
  return w;
}

/** Ajusta posições verticais de flutuantes relativos à página. */
export function fixPageAnchors(container, containerTopPt) {
  for (const el of container.querySelectorAll('[data-abs-v]')) {
    const abs = parseFloat(el.dataset.absV);
    const host = el.offsetParent;
    if (!host) continue;
    const hostTop = host.getBoundingClientRect().top - container.getBoundingClientRect().top;
    const scale = container.getBoundingClientRect().height / (container.offsetHeight || 1) || 1;
    const hostTopPt = (hostTop / scale) * 0.75;
    el.style.top = (abs - containerTopPt - hostTopPt).toFixed(2) + 'pt';
  }
}
