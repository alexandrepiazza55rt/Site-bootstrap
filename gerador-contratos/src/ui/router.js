// Roteamento por hash (#/contratos, #/contrato/:id ...) — funciona em qualquer hospedagem estática.
import { createStore } from './kit.js';

const current = () => (location.hash.replace(/^#/, '') || '/');
export const routeStore = createStore({ path: current() });
window.addEventListener('hashchange', () => routeStore.set({ path: current() }));

export function navigate(path) {
  if (current() === path) routeStore.set({ path });
  else location.hash = path;
}

/** "/contrato/:id" x "/contrato/abc" -> {id: 'abc'} | null */
export function match(pattern, path) {
  const a = pattern.split('/').filter(Boolean);
  const b = path.split('?')[0].split('/').filter(Boolean);
  if (a.length !== b.length) return null;
  const params = {};
  for (let i = 0; i < a.length; i++) {
    if (a[i].startsWith(':')) params[a[i].slice(1)] = decodeURIComponent(b[i]);
    else if (a[i] !== b[i]) return null;
  }
  return params;
}
