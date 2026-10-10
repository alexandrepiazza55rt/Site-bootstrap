// Ponto de entrada do ContratoFácil.
import { render } from 'preact';
import { html } from './ui/kit.js';
import { App } from './ui/App.js';
import './styles/doc.css';
import './styles/app.css';

render(html`<${App} />`, document.getElementById('app'));

// Funciona offline e pode ser instalado como aplicativo (quando servido via https).
if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost' || location.hostname === '127.0.0.1')) {
  window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(() => {}));
}
