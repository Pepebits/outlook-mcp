import dns from 'node:dns/promises';
import net from 'node:net';

const blocked = new net.BlockList();
for (const [addr, prefix] of [
  ['0.0.0.0', 8], ['10.0.0.0', 8], ['100.64.0.0', 10], ['127.0.0.0', 8], ['169.254.0.0', 16], ['172.16.0.0', 12],
  ['192.0.0.0', 24], ['192.168.0.0', 16], ['198.18.0.0', 15], ['224.0.0.0', 4], ['240.0.0.0', 4],
] as const) blocked.addSubnet(addr, prefix, 'ipv4');
for (const [addr, prefix] of [['::', 128], ['::1', 128], ['fc00::', 7], ['fe80::', 10], ['ff00::', 8], ['64:ff9b::', 96]] as const) {
  blocked.addSubnet(addr, prefix, 'ipv6');
}

/** True for loopback, private, link-local, CGNAT, multicast and other non-public addresses. */
export function isPrivateAddress(ip: string): boolean {
  const mapped = /^::ffff:(\d+\.\d+\.\d+\.\d+)$/i.exec(ip);
  if (mapped) return blocked.check(mapped[1], 'ipv4');
  const family = net.isIP(ip);
  if (family === 4) return blocked.check(ip, 'ipv4');
  if (family === 6) return blocked.check(ip, 'ipv6');
  return true;
}

/** Indirection so tests can replace DNS resolution. */
export const resolver = {
  lookup: async (host: string): Promise<string[]> => (await dns.lookup(host, { all: true })).map((a) => a.address),
};

/**
 * Checks that `url` is https and that its host only resolves to public addresses, so a
 * sender-controlled URL cannot make this machine POST to the local network.
 */
export async function assertPublicHttpsUrl(url: string): Promise<URL> {
  let u: URL;
  try {
    u = new URL(url);
  } catch {
    throw new Error('not a valid URL');
  }
  if (u.protocol !== 'https:') throw new Error('only https links are used');
  const host = u.hostname.replace(/^\[|\]$/g, '').toLowerCase();
  if (host === 'localhost' || host.endsWith('.localhost') || host.endsWith('.local') || host.endsWith('.internal')) {
    throw new Error(`${host} is a local host name`);
  }
  const addresses = net.isIP(host) ? [host] : await resolver.lookup(host);
  if (!addresses.length || addresses.some(isPrivateAddress)) throw new Error(`${host} points to a private or local address`);
  return u;
}
