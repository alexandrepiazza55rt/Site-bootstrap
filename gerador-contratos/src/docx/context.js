// Contexto de renderização de um pacote: estilos, numeração, cabeçalhos/rodapés e mídias.
import { DocxPackage, mimeOf } from './pkg.js';
import { Styles } from './styles.js';
import { Numbering } from './numbering.js';
import { wKid, wNum, parseXml } from './xml.js';

export class DocContext {
  static async fromBuffer(buf) {
    const pkg = await DocxPackage.load(buf);
    return new DocContext(pkg);
  }

  constructor(pkg) {
    this.pkg = pkg;
    this.main = pkg.mainPath();
    this.mainRels = pkg.rels(this.main);
    const rel = (t) => pkg.relatedPath(this.main, t);
    const themePath = rel('theme');
    this.styles = new Styles(pkg.xml(rel('styles')), themePath ? pkg.xml(themePath) : null);
    this.numbering = new Numbering(pkg.xml(rel('numbering')), this.styles);
    const settings = pkg.xml(rel('settings'));
    this.defaultTab = 720;
    if (settings) {
      const dt = wKid(settings.documentElement, 'defaultTabStop');
      if (dt) this.defaultTab = wNum(dt, 'val', 720) || 720;
    }
    this.urls = new Map();
    this.partCache = new Map();
  }

  /** Documento principal (string) original do pacote. */
  mainXml() {
    return this.pkg.get(this.main);
  }

  /** Parte relacionada (cabeçalho/rodapé) por rId do documento principal. */
  part(rId) {
    if (this.partCache.has(rId)) return this.partCache.get(rId);
    const r = this.mainRels.get(rId);
    let out = null;
    if (r && this.pkg.has(r.target)) {
      out = { path: r.target, doc: parseXml(this.pkg.get(r.target)), rels: this.pkg.rels(r.target) };
    }
    this.partCache.set(rId, out);
    return out;
  }

  /** URL (blob:) de uma imagem referenciada em uma parte. */
  mediaUrl(rels, rId) {
    const r = rels && rels.get(rId);
    if (!r || r.external) return r && r.external ? r.target : null;
    if (this.urls.has(r.target)) return this.urls.get(r.target);
    const data = this.pkg.get(r.target);
    if (!(data instanceof Uint8Array)) return null;
    const url = URL.createObjectURL(new Blob([data], { type: mimeOf(r.target) }));
    this.urls.set(r.target, url);
    return url;
  }

  dispose() {
    for (const u of this.urls.values()) URL.revokeObjectURL(u);
    this.urls.clear();
  }
}
