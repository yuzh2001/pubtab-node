import { describe, expect, it } from 'vitest';

import { THEME_PRESETS } from '../src/generated/theme-presets.js';

describe('generated theme presets', () => {
  it('包含 three_line 且内容与现有主题约定一致', () => {
    expect(THEME_PRESETS.three_line.name).toBe('three_line');
    expect(THEME_PRESETS.three_line.packages).toContain('booktabs');
    expect(THEME_PRESETS.three_line.caption_position).toBe('top');
  });
});
