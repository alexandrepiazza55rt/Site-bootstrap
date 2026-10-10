// Geração dos arquivos finais: Word (.docx idêntico ao modelo) e PDF (impressão vetorial do navegador).
import { DocxPackage } from './docx/pkg.js';
import { buildFinalXml, coreProps, findListUnits } from './docx/fill.js';
import { parseXml } from './docx/xml.js';
import { buildPages } from './ui/editor/DocCanvas.js';
import { fixPageAnchors } from './docx/render.js';

export async function buildDocx({ pkgBuf, docXml, resolver, titulo, autor }) {
  const pkg = await DocxPackage.load(pkgBuf);
  const main = pkg.mainPath();
  const finalXml = resolver ? buildFinalXml(docXml, resolver) : docXml;
  const overrides = { [main]: finalXml };
  if (pkg.has('docProps/core.xml')) overrides['docProps/core.xml'] = coreProps(pkg.get('docProps/core.xml'), { titulo, autor });
  return pkg.toBlob(overrides);
}

const STD_SIZES = [
  [595.28, 841.89, 'A4'], [612, 792, 'letter'], [612, 1008, 'legal'], [419.53, 595.28, 'A5'],
];

/** Abre a janela de impressão já com as páginas prontas (o usuário escolhe "Salvar como PDF"). */
export async function printDocument({ ctx, docXml, resolver, titulo }) {
  const doc = parseXml(docXml);
  const listUnits = findListUnits(doc, resolver);
  const resolve = (tok, itemIdx) => {
    const m = /^([A-Za-z0-9_]+)(?:@(\d+))?$/.exec(tok);
    const c = m && resolver.byId.get(m[1]);
    if (!m || !c) return { text: '', empty: true };
    const idx = m[2] !== undefined ? +m[2] : resolver.isList(c) ? itemIdx : undefined;
    return resolver.get(m[1], idx);
  };
  const out = buildPages({ ctx, doc, resolve, listUnits, mode: 'print' });
  let root = document.getElementById('print-root');
  if (!root) {
    root = document.createElement('div');
    root.id = 'print-root';
    document.body.appendChild(root);
  }
  root.replaceChildren(...out.pages);
  let w = out.section.w / 20, h = out.section.h / 20;
  // encaixa em tamanhos padrão de papel (evita página em branco extra por diferença mínima)
  for (const [sw, sh, name] of STD_SIZES) {
    const fits = (a, b, c, d) => Math.abs(a - c) < 4 && Math.abs(b - d) < 4;
    if (fits(w, h, sw, sh) || fits(w, h, sh, sw)) {
      const land = w > h;
      w = land ? sh : sw; h = land ? sw : sh;
      root.dataset.paper = name;
      break;
    }
  }
  for (const p of out.pages) {
    p.style.width = w + 'pt';
    p.style.height = h + 'pt';
    const body = p.querySelector('.page-body');
    if (body) fixPageAnchors(body, out.bodyTop);
  }
  let st = document.getElementById('print-page-style');
  if (!st) {
    st = document.createElement('style');
    st.id = 'print-page-style';
    document.head.appendChild(st);
  }
  st.textContent = `@page { size: ${w}pt ${h}pt; margin: 0; }`;
  const oldTitle = document.title;
  document.title = titulo || oldTitle;
  await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
  // aguarda imagens
  await Promise.all(Array.from(root.querySelectorAll('img')).map((img) => (img.complete ? null : new Promise((r) => { img.onload = r; img.onerror = r; }))));
  const cleanup = () => {
    document.title = oldTitle;
    root.replaceChildren();
    window.removeEventListener('afterprint', cleanup);
  };
  window.addEventListener('afterprint', cleanup);
  window.print();
  return out.pages.length;
}
