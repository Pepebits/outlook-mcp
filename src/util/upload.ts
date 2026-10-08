import fs from 'node:fs/promises';

/** Chunk size for upload sessions: 12 x 320 KiB (Graph requires a multiple of 320 KiB). */
export const UPLOAD_CHUNK_BYTES = 12 * 320 * 1024;
const MAX_CHUNK_ATTEMPTS = 4;
const RETRY_STATUS = new Set([429, 500, 502, 503, 504]);

export interface UploadOptions {
  chunkBytes?: number;
  sleep?: (ms: number) => Promise<void>;
}

const defaultSleep = (ms: number) => new Promise<void>((r) => setTimeout(r, ms));

/**
 * Streams a local file to a pre-authenticated Graph upload session URL, one chunk at a time
 * (the file is never fully loaded in memory). The uploadUrl already carries its own
 * authorization, so no Authorization header is sent. Retries a chunk on 429/5xx.
 */
export async function uploadFileInChunks(uploadUrl: string, filePath: string, size: number, opts: UploadOptions = {}): Promise<void> {
  const chunkBytes = opts.chunkBytes ?? UPLOAD_CHUNK_BYTES;
  const sleep = opts.sleep ?? defaultSleep;
  const handle = await fs.open(filePath, 'r');
  try {
    for (let start = 0; start < size; start += chunkBytes) {
      const length = Math.min(chunkBytes, size - start);
      const chunk = Buffer.alloc(length);
      let filled = 0;
      while (filled < length) {
        const { bytesRead } = await handle.read(chunk, filled, length - filled, start + filled);
        if (bytesRead === 0) throw new Error('The file changed while it was being uploaded.');
        filled += bytesRead;
      }
      const headers = {
        'Content-Length': String(length),
        'Content-Range': `bytes ${start}-${start + length - 1}/${size}`,
        'Content-Type': 'application/octet-stream',
      };
      for (let attempt = 1; ; attempt++) {
        const res = await fetch(uploadUrl, { method: 'PUT', headers, body: chunk });
        if (res.ok) break;
        if (RETRY_STATUS.has(res.status) && attempt < MAX_CHUNK_ATTEMPTS) {
          const ra = Number(res.headers.get('Retry-After'));
          await sleep(Number.isFinite(ra) && ra > 0 ? ra * 1000 : 500 * 2 ** (attempt - 1));
          continue;
        }
        throw new Error(`Upload failed with HTTP ${res.status} at bytes ${start}-${start + length - 1}.`);
      }
    }
  } finally {
    await handle.close();
  }
}
