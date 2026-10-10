// Estado global leve (configurações do app).
import { createStore } from './kit.js';

export const appConfig = createStore({});

export function applyTheme(tema) {
  const dark = tema === 'escuro' || (tema !== 'claro' && window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches);
  document.documentElement.dataset.theme = dark ? 'dark' : 'light';
  try { localStorage.setItem('cf-tema', tema || 'auto'); } catch { /* modo privado */ }
}
