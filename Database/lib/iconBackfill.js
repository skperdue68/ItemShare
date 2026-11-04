// Repairs icons that an upload could not resolve.
//
// Icon lookup is best effort: UESP can be down, refuse this host, change its page,
// or the upload can hit its per-request budget. Nothing about those cases changes
// later, so without a retry the item would stay a placeholder for ever - the only
// trigger for a lookup used to be somebody uploading a file.
//
// This pass closes that gap: it asks the database which item ids are still blank,
// looks those up on UESP, writes the paths back and caches the PNGs. It is bounded
// per run and safe to call on a timer: a second pass before the first finishes
// would find nothing new to do, and a failed lookup simply stays blank for the next
// run.

import { cacheItemIcons, resolveIconPathByItemId } from './icons.js';
import { fillIconsByItemId, getItemIdsMissingIcons } from './store.js';

const NO_ICONS = { fetched: 0, skipped: 0, failed: 0 };

/**
 * @returns {Promise<{itemIds: number, resolved: number, rows: number, icons: object}>}
 *   `itemIds` looked up, `resolved` found a path, `rows` item rows updated.
 */
export async function backfillItemIcons(connection, { fetcher = fetch, limit = 20 } = {}) {
  const itemIds = await getItemIdsMissingIcons(connection, limit);

  if (itemIds.length === 0) {
    return { itemIds: 0, resolved: 0, rows: 0, icons: NO_ICONS };
  }

  const paths = new Map();

  for (const itemId of itemIds) {
    const icon = await resolveIconPathByItemId(itemId, { fetcher });

    if (icon) {
      paths.set(itemId, icon);
    }
  }

  const rows = await fillIconsByItemId(connection, paths);
  const icons = paths.size
    ? await cacheItemIcons(connection, [...paths.values()], { fetcher })
    : NO_ICONS;

  return { itemIds: itemIds.length, resolved: paths.size, rows, icons };
}
