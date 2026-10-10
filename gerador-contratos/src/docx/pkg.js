// Leitura/escrita do pacote .docx (ZIP) e dos relacionamentos.
import JSZip from 'jszip';
import { parseXml, NS } from './xml.js';

const TEXT_EXT = /\.(xml|rels)$/i;

export class DocxPackage {
  constructor() {
    this.files = new Map(); // caminho -> string (xml) | Uint8Array
  }

  static async load(data) {
    const zip = await JSZip.loadAsync(data);
    const pkg = new DocxPackage();
    const jobs = [];
    zip.forEach((path, entry) => {
      if (entry.dir) return;
      if (TEXT_EXT.test(path)) jobs.push(entry.async('string').then((s) => pkg.files.set(path, s)));
      else jobs.push(entry.async('uint8array').then((b) => pkg.files.set(path, b)));
    });
    await Promise.all(jobs);
    if (!pkg.mainPath()) throw new Error('Arquivo não parece ser um documento do Word (.docx).');
    return pkg;
  }

  has(path) { return this.files.has(path); }
  get(path) { return this.files.get(path); }
  set(path, data) { this.files.set(path, data); }

  xml(path) {
    const s = this.files.get(path);
    return typeof s === 'string' ? parseXml(s) : null;
  }

  /** Caminho do documento principal (normalmente word/document.xml). */
  mainPath() {
    const rels = this.files.get('_rels/.rels');
    if (typeof rels === 'string') {
      const m = /<Relationship[^>]*Type="[^"]*\/officeDocument"[^>]*Target="([^"]+)"/.exec(rels) ||
        /<Relationship[^>]*Target="([^"]+)"[^>]*Type="[^"]*\/officeDocument"/.exec(rels);
      if (m) return m[1].replace(/^\//, '');
    }
    return this.files.has('word/document.xml') ? 'word/document.xml' : null;
  }

  /** Relacionamentos de uma parte: Map id -> {type, target(caminho resolvido), external} */
  rels(partPath) {
    const dir = partPath.includes('/') ? partPath.slice(0, partPath.lastIndexOf('/')) : '';
    const name = partPath.slice(partPath.lastIndexOf('/') + 1);
    const relPath = (dir ? dir + '/' : '') + '_rels/' + name + '.rels';
    const out = new Map();
    const s = this.files.get(relPath);
    if (typeof s !== 'string') return out;
    const doc = parseXml(s);
    for (const r of Array.from(doc.getElementsByTagNameNS(NS.rel, 'Relationship'))) {
      const external = r.getAttribute('TargetMode') === 'External';
      let target = r.getAttribute('Target') || '';
      if (!external) target = resolvePath(dir, target);
      out.set(r.getAttribute('Id'), { type: r.getAttribute('Type') || '', target, external });
    }
    return out;
  }

  /** Parte relacionada por tipo (ex.: 'styles', 'numbering'). */
  relatedPath(partPath, typeSuffix) {
    for (const r of this.rels(partPath).values()) if (r.type.endsWith('/' + typeSuffix)) return r.target;
    return null;
  }

  async toBlob(overrides = {}) {
    const zip = new JSZip();
    // [Content_Types].xml primeiro, por convenção
    const order = Array.from(this.files.keys()).sort((a, b) => (a === '[Content_Types].xml' ? -1 : b === '[Content_Types].xml' ? 1 : 0));
    for (const path of order) {
      const data = Object.prototype.hasOwnProperty.call(overrides, path) ? overrides[path] : this.files.get(path);
      zip.file(path, data);
    }
    return zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });
  }

  async toArrayBuffer(overrides = {}) {
    const blob = await this.toBlob(overrides);
    return blob.arrayBuffer();
  }
}

export function resolvePath(baseDir, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (baseDir ? baseDir.split('/') : []).concat(target.split('/'));
  const out = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p && p !== '.') out.push(p);
  }
  return out.join('/');
}

const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/emf', wmf: 'image/wmf' };

export function mimeOf(path) {
  const ext = path.slice(path.lastIndexOf('.') + 1).toLowerCase();
  return MIME[ext] || 'application/octet-stream';
}
