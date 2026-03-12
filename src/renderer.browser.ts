import { createRenderer } from './render-core.js';
import { getTheme } from './themes.browser.js';

export const { render } = createRenderer(getTheme);
