// Componentes básicos de interface (Preact + htm, sem etapa de JSX).
import { h, Fragment } from 'preact';
import { useState, useEffect, useRef, useCallback } from 'preact/hooks';
import htm from 'htm';
import { Icon } from './icons.js';

export const html = htm.bind(h);
export { Fragment, Icon };

export const cx = (...a) => a.filter(Boolean).join(' ');

// ----------------------------------------------------------- store simples
export function createStore(initial) {
  let state = initial;
  const subs = new Set();
  return {
    get: () => state,
    set: (patch) => { state = typeof patch === 'function' ? patch(state) : { ...state, ...patch }; subs.forEach((f) => f(state)); },
    sub: (f) => { subs.add(f); return () => subs.delete(f); },
  };
}

export function useStore(store) {
  const [s, setS] = useState(store.get());
  useEffect(() => store.sub(setS), [store]);
  return s;
}

// ----------------------------------------------------------- botões
export function Button({ variant = 'default', size, icon, iconRight, children, class: cls, loading, ...rest }) {
  return html`<button type="button" class=${cx('btn', 'btn-' + variant, size && 'btn-' + size, loading && 'is-loading', cls)} disabled=${rest.disabled || loading} ...${rest}>
    ${loading ? html`<span class="spinner" />` : icon && html`<${Icon} name=${icon} size=${size === 'sm' ? 15 : 17} />`}
    ${children && html`<span>${children}</span>`}
    ${iconRight && html`<${Icon} name=${iconRight} size=${15} />`}
  </button>`;
}

export function IconButton({ icon, title, active, size = 18, class: cls, ...rest }) {
  return html`<button type="button" class=${cx('icon-btn', active && 'is-active', cls)} title=${title} aria-label=${title} ...${rest}>
    <${Icon} name=${icon} size=${size} />
  </button>`;
}

export function Badge({ tone = 'gray', children, icon }) {
  return html`<span class=${'badge badge-' + tone}>${icon && html`<${Icon} name=${icon} size=${12} />`}${children}</span>`;
}

export function Spinner({ size = 18 }) {
  return html`<span class="spinner" style=${{ width: size, height: size }} />`;
}

export function Empty({ icon = 'folder-open', title, children, action }) {
  return html`<div class="empty">
    <div class="empty-icon"><${Icon} name=${icon} size=${28} /></div>
    <h3>${title}</h3>
    ${children && html`<p>${children}</p>`}
    ${action}
  </div>`;
}

// ----------------------------------------------------------- toasts
const toastStore = createStore({ list: [] });
let toastSeq = 0;

export function toast(msg, opts = {}) {
  const id = ++toastSeq;
  const t = { id, msg, type: opts.type || 'info', action: opts.action, title: opts.title };
  toastStore.set((s) => ({ list: [...s.list.slice(-3), t] }));
  const ms = opts.duration || (opts.action ? 6500 : opts.type === 'error' ? 6000 : 3200);
  setTimeout(() => dismissToast(id), ms);
  return id;
}
toast.ok = (m, o) => toast(m, { ...o, type: 'ok' });
toast.err = (m, o) => toast(m, { ...o, type: 'error' });
toast.warn = (m, o) => toast(m, { ...o, type: 'warn' });

export function dismissToast(id) {
  toastStore.set((s) => ({ list: s.list.filter((t) => t.id !== id) }));
}

const TOAST_ICON = { ok: 'check-circle', error: 'alert-circle', warn: 'alert', info: 'info' };

export function ToastHost() {
  const { list } = useStore(toastStore);
  return html`<div class="toasts" role="status" aria-live="polite">
    ${list.map((t) => html`<div class=${'toast toast-' + t.type} key=${t.id}>
      <${Icon} name=${TOAST_ICON[t.type]} size=${18} />
      <div class="toast-body">${t.title && html`<strong>${t.title}</strong>`}<span>${t.msg}</span></div>
      ${t.action && html`<button class="toast-action" onClick=${() => { t.action.onClick(); dismissToast(t.id); }}>${t.action.label}</button>`}
      <button class="toast-x" aria-label="Fechar" onClick=${() => dismissToast(t.id)}><${Icon} name="x" size=${14} /></button>
    </div>`)}
  </div>`;
}

// ----------------------------------------------------------- modais
const modalStore = createStore({ stack: [] });
let modalSeq = 0;

/** Abre um modal. render({close}) -> vnode. Retorna Promise com o valor passado a close(). */
export function openModal(render, opts = {}) {
  return new Promise((resolve) => {
    const id = ++modalSeq;
    const close = (v) => {
      modalStore.set((s) => ({ stack: s.stack.filter((m) => m.id !== id) }));
      resolve(v);
    };
    modalStore.set((s) => ({ stack: [...s.stack, { id, render, close, opts }] }));
  });
}

export function ModalHost() {
  const { stack } = useStore(modalStore);
  useEffect(() => {
    if (!stack.length) return undefined;
    const top = stack[stack.length - 1];
    const onKey = (e) => { if (e.key === 'Escape' && !top.opts.persistent) { e.stopPropagation(); top.close(undefined); } };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, [stack]);
  return html`${stack.map((m) => html`<div class="modal-backdrop" key=${m.id} onMouseDown=${(e) => { if (e.target === e.currentTarget && !m.opts.persistent) m.close(undefined); }}>
    <div class=${cx('modal', m.opts.size && 'modal-' + m.opts.size)} role="dialog" aria-modal="true">${m.render({ close: m.close })}</div>
  </div>`)}`;
}

export function ModalHead({ title, sub, icon, close }) {
  return html`<div class="modal-head">
    ${icon && html`<div class="modal-icon"><${Icon} name=${icon} size=${20} /></div>`}
    <div class="modal-titles"><h2>${title}</h2>${sub && html`<p>${sub}</p>`}</div>
    ${close && html`<${IconButton} icon="x" title="Fechar" onClick=${() => close(undefined)} />`}
  </div>`;
}

export function confirmDialog({ title, message, ok = 'Confirmar', cancel = 'Cancelar', danger, icon }) {
  return openModal(({ close }) => html`
    <${ModalHead} title=${title} icon=${icon || (danger ? 'alert' : 'help')} close=${close} />
    <div class="modal-body">${typeof message === 'string' ? html`<p class="muted">${message}</p>` : message}</div>
    <div class="modal-foot">
      <${Button} variant="ghost" onClick=${() => close(false)}>${cancel}</${Button}>
      <${Button} variant=${danger ? 'danger' : 'primary'} onClick=${() => close(true)} autoFocus>${ok}</${Button}>
    </div>`, { size: 'sm' });
}

export function promptDialog({ title, label, value = '', ok = 'Salvar', placeholder, icon, multiline }) {
  return openModal(({ close }) => {
    const Inner = () => {
      const [v, setV] = useState(value);
      const ref = useRef();
      useEffect(() => { ref.current && ref.current.focus(); ref.current && ref.current.select && ref.current.select(); }, []);
      const submit = (e) => { e && e.preventDefault(); if (v.trim()) close(v.trim()); };
      return html`<form onSubmit=${submit}>
        <${ModalHead} title=${title} icon=${icon || 'pencil'} close=${close} />
        <div class="modal-body">
          <label class="field"><span class="field-label">${label}</span>
            ${multiline
              ? html`<textarea ref=${ref} class="input" rows="4" value=${v} placeholder=${placeholder} onInput=${(e) => setV(e.currentTarget.value)} />`
              : html`<input ref=${ref} class="input" value=${v} placeholder=${placeholder} onInput=${(e) => setV(e.currentTarget.value)} />`}
          </label>
        </div>
        <div class="modal-foot">
          <${Button} variant="ghost" onClick=${() => close(undefined)}>Cancelar</${Button}>
          <${Button} variant="primary" type="submit" disabled=${!v.trim()}>${ok}</${Button}>
        </div>
      </form>`;
    };
    return html`<${Inner} />`;
  }, { size: 'sm' });
}

// ----------------------------------------------------------- menu suspenso
export function Menu({ trigger, items, align = 'left', class: cls }) {
  const [open, setOpen] = useState(false);
  const ref = useRef();
  useEffect(() => {
    if (!open) return undefined;
    const onDoc = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    const onKey = (e) => { if (e.key === 'Escape') setOpen(false); };
    document.addEventListener('mousedown', onDoc);
    document.addEventListener('keydown', onKey);
    return () => { document.removeEventListener('mousedown', onDoc); document.removeEventListener('keydown', onKey); };
  }, [open]);
  return html`<div class=${cx('menu-wrap', cls)} ref=${ref}>
    ${trigger({ open, toggle: () => setOpen(!open) })}
    ${open && html`<div class=${'menu menu-' + align} role="menu">
      ${items.filter(Boolean).map((it, i) => it === '-' ? html`<div class="menu-sep" key=${i} />`
        : it.header ? html`<div class="menu-header" key=${i}>${it.header}</div>`
        : html`<button key=${i} role="menuitem" class=${cx('menu-item', it.danger && 'danger')} disabled=${it.disabled}
            onClick=${() => { setOpen(false); it.onClick && it.onClick(); }}>
            ${it.icon && html`<${Icon} name=${it.icon} size=${16} />`}
            <span class="menu-label">${it.label}${it.hint && html`<small>${it.hint}</small>`}</span>
            ${it.kbd && html`<kbd>${it.kbd}</kbd>`}
          </button>`)}
    </div>`}
  </div>`;
}

// ----------------------------------------------------------- utilidades
export function useDebounced(fn, ms) {
  const t = useRef();
  const f = useRef(fn);
  f.current = fn;
  const call = useCallback((...a) => { clearTimeout(t.current); t.current = setTimeout(() => f.current(...a), ms); }, [ms]);
  call.flush = (...a) => { clearTimeout(t.current); f.current(...a); };
  call.cancel = () => clearTimeout(t.current);
  return call;
}

export function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

export function pickFile(accept) {
  return new Promise((resolve) => {
    const inp = document.createElement('input');
    inp.type = 'file';
    inp.accept = accept;
    inp.onchange = () => resolve(inp.files && inp.files[0] ? inp.files[0] : null);
    inp.click();
  });
}

const RTF = new Intl.RelativeTimeFormat('pt-BR', { numeric: 'auto' });
export function relTime(iso) {
  if (!iso) return '';
  const diff = (Date.parse(iso) - Date.now()) / 1000;
  const abs = Math.abs(diff);
  if (abs < 45) return 'agora mesmo';
  if (abs < 3600) return RTF.format(Math.round(diff / 60), 'minute');
  if (abs < 86400) return RTF.format(Math.round(diff / 3600), 'hour');
  if (abs < 86400 * 30) return RTF.format(Math.round(diff / 86400), 'day');
  return new Date(iso).toLocaleDateString('pt-BR');
}

export const brl = (n) => (Number.isFinite(n) ? n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }) : '—');
export const dataBR = (iso) => (iso ? iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4) : '—');
