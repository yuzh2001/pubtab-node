import { createApp } from 'vue';

import App from './App.vue';
import './style.css';

export function mountApp(container: Element | string = '#app') {
  const target = typeof container === 'string' ? document.querySelector(container) : container;
  if (!target) {
    throw new Error(`Mount target not found: ${String(container)}`);
  }

  const app = createApp(App);
  return app.mount(target);
}

if (typeof document !== 'undefined') {
  const existing = document.querySelector('#app');
  if (existing) {
    mountApp(existing);
  }
}
