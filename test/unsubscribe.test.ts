import { describe, expect, it } from 'vitest';
import { parseMailto, parseUnsubscribeHeaders } from '../src/util/unsubscribe.js';

describe('parseUnsubscribeHeaders', () => {
  it('detects one-click https and mailto targets', () => {
    const info = parseUnsubscribeHeaders([
      { name: 'List-Unsubscribe', value: '<mailto:leave@list.example?subject=stop>, <https://list.example/u/abc>' },
      { name: 'list-unsubscribe-post', value: 'List-Unsubscribe=One-Click' },
    ]);
    expect(info.oneClick).toBe(true);
    expect(info.https).toEqual(['https://list.example/u/abc']);
    expect(info.mailto).toEqual([{ address: 'leave@list.example', subject: 'stop', body: 'unsubscribe' }]);
  });

  it('is not one-click without the Post header', () => {
    const info = parseUnsubscribeHeaders([{ name: 'List-Unsubscribe', value: '<https://list.example/u>' }]);
    expect(info.oneClick).toBe(false);
  });

  it('ignores plain http links and missing headers', () => {
    expect(parseUnsubscribeHeaders([{ name: 'List-Unsubscribe', value: '<http://x.example/u>' }]).https).toEqual([]);
    expect(parseUnsubscribeHeaders([])).toEqual({ https: [], mailto: [], oneClick: false });
  });
});

describe('parseMailto', () => {
  it('rejects non-mailto and invalid addresses', () => {
    expect(parseMailto('https://x.example')).toBeNull();
    expect(parseMailto('mailto:nobody')).toBeNull();
  });
});
