// Build do ContratoFácil: empacota src/ em app.js (IIFE) — funciona até abrindo o index.html direto do disco.
// Uso: npm run build   |   npm run watch
import * as esbuild from 'esbuild';
import { readFileSync, writeFileSync } from 'node:fs';

const watch = process.argv.includes('--watch');
const entries = process.env.ENTRY ? process.env.ENTRY.split(',') : ['src/main.js'];

const options = {
  entryPoints: entries,
  bundle: true,
  format: 'iife',
  target: ['chrome100', 'firefox100', 'safari15', 'edge100'],
  outdir: process.env.OUTDIR || '.',
  entryNames: process.env.ENTRY ? '[name]' : 'app',
  minify: !watch && !process.env.NOMINIFY,
  sourcemap: watch ? 'inline' : false,
  loader: { '.docx': 'binary' },
  jsx: 'automatic',
  legalComments: 'none',
  define: { __VERSION__: JSON.stringify(JSON.parse(readFileSync('package.json', 'utf8')).version) },
  logLevel: 'info',
};

if (watch) {
  const ctx = await esbuild.context(options);
  await ctx.watch();
  console.log('Observando alterações…');
} else {
  await esbuild.build(options);
  if (!process.env.ENTRY) {
    // versão do service worker (força atualização do cache a cada build)
    const sw = readFileSync('sw.src.js', 'utf8').replace('__BUILD__', Date.now().toString(36));
    writeFileSync('sw.js', sw);
    // versão "arquivo único": tudo embutido num só .html (abre com dois cliques, sem internet)
    let html = readFileSync('index.html', 'utf8');
    const css = readFileSync('app.css', 'utf8');
    const js = readFileSync('app.js', 'utf8').replace(/<\/script/g, '<\\/script');
    html = html.replace('<link rel="stylesheet" href="app.css">', () => '<style>' + css + '</style>')
      .replace('<script src="app.js"></script>', () => '<script>' + js + '</script>')
      .replace('<link rel="manifest" href="manifest.webmanifest">', '');
    writeFileSync('ContratoFacil.html', html);
  }
}
