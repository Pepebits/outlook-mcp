import os from 'node:os';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { READ_ONLY_SCOPES, READ_WRITE_SCOPES, parseConfig } from '../src/config.js';

describe('parseConfig', () => {
  it('only accepts https Microsoft Graph hosts as the base URL', () => {
    expect(parseConfig({}).graphBaseUrl).toBe('https://graph.microsoft.com/v1.0');
    expect(parseConfig({ OUTLOOK_GRAPH_BASE_URL: 'https://graph.microsoft.us/v1.0/' }).graphBaseUrl).toBe('https://graph.microsoft.us/v1.0');
    for (const bad of ['https://evil.example/v1.0', 'http://graph.microsoft.com/v1.0', 'https://graph.microsoft.com.evil.example/v1.0']) {
      expect(() => parseConfig({ OUTLOOK_GRAPH_BASE_URL: bad })).toThrow(/Microsoft Graph host/);
    }
  });

  it('narrows the default scopes in read-only mode; explicit OUTLOOK_SCOPES wins', () => {
    expect(parseConfig({}).scopes.join(' ')).toBe(READ_WRITE_SCOPES);
    expect(parseConfig({ OUTLOOK_READ_ONLY: 'true' }).scopes.join(' ')).toBe(READ_ONLY_SCOPES);
    expect(parseConfig({ OUTLOOK_READ_ONLY: 'true', OUTLOOK_SCOPES: 'User.Read Mail.ReadWrite' }).scopes).toEqual(['User.Read', 'Mail.ReadWrite']);
  });

  it('defaults attachment roots to the download dir, Documents and Desktop', () => {
    const home = os.homedir();
    expect(parseConfig({ OUTLOOK_DOWNLOAD_DIR: '/tmp/dl' }).attachDirs).toEqual([path.resolve('/tmp/dl'), path.join(home, 'Documents'), path.join(home, 'Desktop')]);
    expect(parseConfig({ OUTLOOK_ATTACH_DIRS: `~/a${path.delimiter}/b` }).attachDirs).toEqual([path.join(home, 'a'), path.resolve('/b')]);
  });
});
