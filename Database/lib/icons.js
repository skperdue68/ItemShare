// Item icons: fetched once from UESP's PNG mirror, then stored in MariaDB.
//
// The addon records the game's icon path, which is a DDS texture browsers cannot
// draw. UESP mirrors the same tree as PNG, so /esoui/art/icons/x.dds becomes
// https://esoicons.uesp.net/esoui/art/icons/x.png. Each distinct path is fetched
// at most once: afterwards the site serves the stored bytes from /icons/<key>.png
// and never contacts UESP again.

import { createHash } from 'node:crypto';

const ICON_BASE = 'https://esoicons.uesp.net';
const MAX_ICON_BYTES = 512 * 1024;

/** Stable cache key for an icon path. */
export function iconCacheKey(iconPath) {
  const path = String(iconPath || '').trim();
  return path ? createHash('sha1').update(path, 'utf8').digest('hex') : '';
}

/** UESP PNG URL for a game icon path, or '' when the path is unusable. */
export function iconSourceUrl(iconPath) {
  const path = String(iconPath || '').trim();
  if (!path.startsWith('/')) {
    return '';
  }
  return `${ICON_BASE}${path.replace(/\.dds$/i, '.png')}`;
}

/**
 * Fetches and stores only the icons that are not in the database yet.
 *
 * UESP is a last resort, not a per-upload step: the cache table is checked first
 * (in chunks), and every path already present costs nothing but one row in that
 * lookup. Best effort: a failure leaves the icon missing and is retried by a
 * later upload, never by page views.
 *
 * @returns {Promise<{fetched: number, skipped: number, failed: number}>}
 */
export async function cacheItemIcons(connection, iconPaths, { limit = 60, fetcher = fetch } = {}) {
  const paths = [
    ...new Set(
      (iconPaths || [])
        .map((path) => String(path || '').trim())
        .filter((path) => path.startsWith('/'))
    ),
  ];

  if (paths.length === 0) {
    return { fetched: 0, skipped: 0, failed: 0 };
  }

  const entries = paths.map((path) => ({ key: iconCacheKey(path), path }));

  // Database first: anything already stored is skipped below and never requested
  // from UESP again. Chunked so a large inventory does not build one huge IN().
  const cached = new Set();
  const chunkSize = 250;

  for (let start = 0; start < entries.length; start += chunkSize) {
    const chunk = entries.slice(start, start + chunkSize);
    const [rows] = await connection.execute(
      `SELECT icon_key FROM itemshare_item_icons WHERE icon_key IN (${chunk.map(() => '?').join(', ')})`,
      chunk.map((entry) => entry.key)
    );
    for (const row of rows) {
      cached.add(row.icon_key);
    }
  }

  let fetched = 0;
  let failed = 0;

  for (const entry of entries) {
    if (cached.has(entry.key)) {
      continue;
    }
    if (fetched + failed >= limit) {
      break;
    }

    const url = iconSourceUrl(entry.path);
    if (!url) {
      continue;
    }

    try {
      const response = await fetcher(url, { signal: AbortSignal.timeout(8000) });

      if (!response.ok) {
        failed += 1;
        continue;
      }

      const contentType = String(response.headers.get('content-type') || '').split(';')[0].trim();

      if (!contentType.startsWith('image/')) {
        failed += 1;
        continue;
      }

      const image = Buffer.from(await response.arrayBuffer());

      if (image.length === 0 || image.length > MAX_ICON_BYTES) {
        failed += 1;
        continue;
      }

      await connection.execute(
        `INSERT INTO itemshare_item_icons (icon_key, source_path, content_type, image_data)
         VALUES (?, ?, ?, ?)
         ON DUPLICATE KEY UPDATE icon_key = icon_key`,
        [entry.key, entry.path, contentType, image]
      );
      fetched += 1;
    } catch {
      failed += 1;
    }
  }

  return { fetched, skipped: entries.length - fetched - failed, failed };
}

export async function getCachedIcon(connection, iconKey) {
  const key = String(iconKey || '');

  if (!/^[a-f0-9]{40}$/.test(key)) {
    return null;
  }

  const [rows] = await connection.execute(
    'SELECT content_type, image_data FROM itemshare_item_icons WHERE icon_key = ?',
    [key]
  );

  return rows[0] || null;
}
