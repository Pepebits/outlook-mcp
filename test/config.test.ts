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

describe('default user directories', () => {
  const plat = (platform: NodeJS.Platform, files: Record<string, string> = {}) => ({ platform, home: '/home/ana', readFile: (p: string) => files[p] });

  it('uses localized XDG user dirs on Linux and ignores dirs set to $HOME', () => {
    const dirs = `# comment\nXDG_DOWNLOAD_DIR="$HOME/Descargas"\nXDG_DOCUMENTS_DIR="$HOME/Documentos"\nXDG_DESKTOP_DIR="$HOME"\n`;
    const cfg = parseConfig({}, plat('linux', { '/home/ana/.config/user-dirs.dirs': dirs }));
    expect(cfg.downloadDir).toBe(path.resolve('/home/ana/Descargas'));
    expect(cfg.attachDirs).toEqual([path.resolve('/home/ana/Descargas'), path.resolve('/home/ana/Documentos'), path.resolve('/home/ana/Desktop')]);
  });

  it('honours XDG_CONFIG_HOME and falls back to English names without the file', () => {
    const cfg = parseConfig({ XDG_CONFIG_HOME: '/cfg' }, plat('linux', { '/cfg/user-dirs.dirs': 'XDG_DOCUMENTS_DIR="/data/docs"' }));
    expect(cfg.attachDirs).toContain(path.resolve('/data/docs'));
    expect(parseConfig({}, plat('linux')).downloadDir).toBe(path.resolve('/home/ana/Downloads'));
  });

  it('adds OneDrive-redirected Documents and Desktop on Windows', () => {
    const cfg = parseConfig({ OneDrive: '/od' }, plat('win32'));
    expect(cfg.attachDirs).toEqual(expect.arrayContaining([path.resolve('/od/Documents'), path.resolve('/od/Desktop'), path.resolve('/home/ana/Documents')]));
  });
});
