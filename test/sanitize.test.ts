import { describe, expect, it } from 'vitest';
import { htmlToText, truncate } from '../src/util/sanitize.js';

describe('sanitize', () => {
  it('skips img/style/script and hides duplicate hrefs', () => {
    const t = htmlToText('<style>p{}</style><script>x()</script><p>Hi <img src="a.png"> <a href="https://a.com">https://a.com</a></p>');
    expect(t).toBe('Hi https://a.com');
  });
  it('collapses blank lines and strips zero-width chars', () => {
    expect(htmlToText('<p>a​</p><br><br><br><br><p>b</p>')).not.toMatch(/\n{3,}|​/);
  });
  it('truncates with a marker', () => {
    expect(truncate('abcdef', 4)).toBe('abcd\n...[truncated 2 chars]');
    expect(truncate('abc', 4)).toBe('abc');
  });
});
