// Database operations shared by the CLI (ingest.js) and the website (server.js).
// An upload is a full replace: the uploading account's guild membership and item
// list are both rewritten from the file, inside one transaction.

import { createHash } from 'node:crypto';

import { iconCacheKey } from './icons.js';

const ITEM_COLUMNS = [
  'account_name',
  'character_name',
  'item_key',
  'item_key_hash',
  'item_link',
  'item_name',
  'item_id',
  'icon',
  'icon_key',
  'item_type',
  'item_type_name',
  'quality',
  'quality_name',
  'trait',
  'trait_type',
  'apparel_slot',
  'equip_type',
  'apparel_weight',
  'armor_type',
  'weapon_type',
  'weapon_handedness',
  'furniture_data_id',
  'location_key',
  'location_label',
  'item_count',
  'first_dumped_at',
  'last_dumped_at',
  'upload_id',
];

const INSERT_ITEM_SQL = `
INSERT INTO itemshare_shared_items (${ITEM_COLUMNS.join(', ')})
VALUES (${ITEM_COLUMNS.map(() => '?').join(', ')})`;

export function hashItemKey(itemKey) {
  return createHash('sha1').update(itemKey, 'utf8').digest('hex');
}

function itemValues(item, uploadId) {
  return [
    item.accountName,
    item.characterName,
    item.itemKey,
    hashItemKey(item.itemKey),
    item.itemLink,
    item.itemName,
    item.itemId || 0,
    item.icon || '',
    iconCacheKey(item.icon),
    item.itemType,
    item.itemTypeName,
    item.quality,
    item.qualityName,
    item.trait,
    item.traitType,
    item.apparelSlot,
    item.equipType,
    item.apparelWeight,
    item.armorType,
    item.weaponType,
    item.weaponHandedness,
    item.furnitureDataId,
    item.locationKey,
    item.locationLabel,
    item.itemCount,
    item.firstDumpedAt,
    item.lastDumpedAt,
    uploadId,
  ];
}

async function upsertGuilds(connection, guilds) {
  for (const [guildId, guildName] of guilds) {
    await connection.execute(
      `INSERT INTO itemshare_guilds (guild_id, guild_name) VALUES (?, ?)
       ON DUPLICATE KEY UPDATE guild_name = IF(VALUES(guild_name) = '', guild_name, VALUES(guild_name))`,
      [guildId, guildName]
    );
  }
}

async function replaceAccountGuilds(connection, accountName, guilds) {
  await connection.execute('DELETE FROM itemshare_account_guilds WHERE account_name = ?', [accountName]);

  for (const guildId of guilds.keys()) {
    await connection.execute(
      'INSERT IGNORE INTO itemshare_account_guilds (account_name, guild_id) VALUES (?, ?)',
      [accountName, guildId]
    );
  }
}

/**
 * Replaces the shared items and guild membership of one ESO account. Runs inside
 * the caller's transaction so every account in a file is replaced atomically.
 */
async function replaceOneAccount(connection, { accountName, items, guilds, sourceFile, userId }) {
  const totalQuantity = items.reduce((sum, item) => sum + (item.itemCount || 0), 0);

  await upsertGuilds(connection, guilds);
  await replaceAccountGuilds(connection, accountName, guilds);

  const [upload] = await connection.execute(
    `INSERT INTO itemshare_uploads
       (account_name, character_name, source_file, item_count, total_quantity, guild_count, user_id)
     VALUES (?, ?, ?, ?, ?, ?, ?)`,
    [accountName, '*', sourceFile || '', items.length, totalQuantity, guilds.size, userId || null]
  );

  const uploadId = upload.insertId;

  await connection.execute('DELETE FROM itemshare_shared_items WHERE account_name = ?', [accountName]);

  for (const item of items) {
    const [inserted] = await connection.execute(INSERT_ITEM_SQL, itemValues(item, uploadId));
    for (const guildId of item.guildIds) {
      await connection.execute(
        'INSERT IGNORE INTO itemshare_shared_item_guilds (item_id, guild_id) VALUES (?, ?)',
        [inserted.insertId, guildId]
      );
    }
  }

  return { accountName, uploadId, itemCount: items.length, totalQuantity, guildCount: guilds.size };
}

/**
 * Full replace for every ESO account in one saved variables file. Either all of
 * the accounts are updated or, on any failure, none of them are.
 * @param {Array<{accountName: string, items: object[], guilds: Map<number, string>}>} accounts
 * @returns {Promise<Array<{accountName: string, uploadId: number, itemCount: number, totalQuantity: number, guildCount: number}>>}
 */
export async function replaceAccounts(connection, accounts, { sourceFile, userId } = {}) {
  await connection.beginTransaction();

  try {
    const results = [];
    for (const account of accounts) {
      results.push(await replaceOneAccount(connection, { ...account, sourceFile, userId }));
    }
    await connection.commit();
    return results;
  } catch (error) {
    await connection.rollback();
    throw error;
  }
}

/**
 * Resolves the person behind a provider sign-in.
 *
 * 1. a known (provider, provider_user_id) identity wins;
 * 2. otherwise a verified email that already belongs to a person joins that
 *    person, so Discord and Google logins of the same human share their ESO
 *    accounts instead of creating a second, empty account;
 * 3. otherwise a new person is created.
 *
 * @returns {Promise<{userId: number, displayName: string, email: string, emailVerified: boolean, matched: 'identity'|'email'|'created'}>}
 */
export async function signInWithProvider(
  connection,
  { provider, providerUserId, username, email, emailVerified }
) {
  const providerKey = String(provider);
  const subject = String(providerUserId);
  const address = email || '';
  const verified = Boolean(emailVerified);
  const name = username || '';

  const [identityRows] = await connection.execute(
    'SELECT user_id FROM itemshare_user_identities WHERE provider = ? AND provider_user_id = ?',
    [providerKey, subject]
  );

  let userId = identityRows[0] ? Number(identityRows[0].user_id) : null;
  let matched = 'identity';

  if (userId === null && address && verified) {
    const [byEmail] = await connection.execute(
      `SELECT id FROM itemshare_users
        WHERE email = ? AND email_verified = 1
        ORDER BY id ASC
        LIMIT 1`,
      [address]
    );

    if (byEmail[0]) {
      userId = Number(byEmail[0].id);
      matched = 'email';
    }
  }

  if (userId === null) {
    const [inserted] = await connection.execute(
      'INSERT INTO itemshare_users (display_name, email, email_verified) VALUES (?, ?, ?)',
      [name, address, verified ? 1 : 0]
    );
    userId = Number(inserted.insertId);
    matched = 'created';
  } else {
    await connection.execute(
      `UPDATE itemshare_users
          SET display_name = IF(? = '', display_name, ?),
              email = IF(? = '', email, ?),
              email_verified = IF(? = 1, 1, email_verified)
        WHERE id = ?`,
      [name, name, address, address, verified ? 1 : 0, userId]
    );
  }

  await connection.execute(
    `INSERT INTO itemshare_user_identities
       (provider, provider_user_id, user_id, provider_username, provider_email, email_verified)
     VALUES (?, ?, ?, ?, ?, ?)
     ON DUPLICATE KEY UPDATE
       user_id = VALUES(user_id),
       provider_username = VALUES(provider_username),
       provider_email = VALUES(provider_email),
       email_verified = VALUES(email_verified)`,
    [providerKey, subject, userId, name, address, verified ? 1 : 0]
  );

  return { userId, displayName: name, email: address, emailVerified: verified, matched };
}

/** Email for an ESO account's owner, for request notifications. */
export async function getOwnerNotificationEmail(connection, accountName) {
  const [rows] = await connection.execute(
    `SELECT u.email, u.email_verified
       FROM itemshare_accounts a
       JOIN itemshare_users u ON u.id = a.user_id
      WHERE a.eso_account_name = ?`,
    [accountName]
  );

  const row = rows[0];
  if (!row || !row.email) {
    return null;
  }
  return { email: row.email, verified: Number(row.email_verified) === 1 };
}

export async function getAccountSummary(connection, accountName) {
  const [[items]] = await connection.execute(
    `SELECT COUNT(*) AS itemCount, COALESCE(SUM(item_count), 0) AS totalQuantity
       FROM itemshare_shared_items WHERE account_name = ?`,
    [accountName]
  );

  const [[guilds]] = await connection.execute(
    `SELECT COUNT(*) AS guildCount FROM itemshare_account_guilds WHERE account_name = ?`,
    [accountName]
  );

  const [[lastUpload]] = await connection.execute(
    `SELECT uploaded_at, item_count, total_quantity, source_file
       FROM itemshare_uploads WHERE account_name = ?
      ORDER BY id DESC LIMIT 1`,
    [accountName]
  );

  const [linked] = await connection.execute(
    `SELECT a.user_id, u.display_name
       FROM itemshare_accounts a
       LEFT JOIN itemshare_users u ON u.id = a.user_id
      WHERE a.eso_account_name = ?`,
    [accountName]
  );

  return {
    itemCount: Number(items?.itemCount || 0),
    totalQuantity: Number(items?.totalQuantity || 0),
    guildCount: Number(guilds?.guildCount || 0),
    lastUpload: lastUpload || null,
    link: linked[0] || null,
  };
}

/** Every ESO account linked to one person, oldest link first. */
export async function getViewerAccounts(connection, userId) {
  const [rows] = await connection.execute(
    `SELECT a.eso_account_name, a.user_id, a.linked_at, u.display_name
       FROM itemshare_accounts a
       LEFT JOIN itemshare_users u ON u.id = a.user_id
      WHERE a.user_id = ?
      ORDER BY a.linked_at ASC, a.eso_account_name ASC`,
    [Number(userId)]
  );
  return rows;
}

function viewerAccountNames(accountNames) {
  return [...new Set((accountNames || []).map((name) => String(name || '').trim()).filter(Boolean))];
}

function inClause(values) {
  return values.map(() => '?').join(', ');
}

/** Guilds any of the viewer's ESO accounts belongs to. */
export async function getViewerGuilds(connection, accountNames) {
  const names = viewerAccountNames(accountNames);
  if (names.length === 0) {
    return [];
  }

  const [rows] = await connection.execute(
    `SELECT ag.guild_id, COALESCE(g.guild_name, '') AS guild_name
       FROM itemshare_account_guilds ag
       LEFT JOIN itemshare_guilds g ON g.guild_id = ag.guild_id
      WHERE ag.account_name IN (${inClause(names)})
      GROUP BY ag.guild_id, g.guild_name
      ORDER BY guild_name ASC, ag.guild_id ASC`,
    names
  );
  return rows;
}

const VISIBLE_ITEMS_FROM = `
       FROM itemshare_shared_items i
       JOIN itemshare_shared_item_guilds ig ON ig.item_id = i.id
       JOIN itemshare_account_guilds viewer
            ON viewer.guild_id = ig.guild_id
           AND viewer.account_name IN`;

/**
 * Item types and owner accounts present in what the viewer can see, used to fill
 * the filter controls. Not affected by the active filters.
 */
export async function getViewerFacets(connection, accountNames) {
  const names = viewerAccountNames(accountNames);
  if (names.length === 0) {
    return { types: [], owners: [] };
  }

  const [types] = await connection.execute(
    `SELECT i.item_type_name AS name, COUNT(*) AS total
     ${VISIBLE_ITEMS_FROM} (${inClause(names)})
      WHERE i.item_type_name <> ''
      GROUP BY i.item_type_name
      ORDER BY i.item_type_name ASC`,
    names
  );

  const [owners] = await connection.execute(
    `SELECT DISTINCT i.account_name AS name
     ${VISIBLE_ITEMS_FROM} (${inClause(names)})
      ORDER BY i.account_name ASC`,
    names
  );

  return {
    types: types.map((row) => ({ name: row.name, total: Number(row.total) || 0 })),
    owners: owners.map((row) => row.name),
  };
}

/**
 * Shared items offered to any guild the viewer belongs to, grouped by item type
 * and alphabetical inside each type. Ownership comes from the item's own account,
 * visibility from the guild overlap.
 */
export async function getSharedItemsForViewer(
  connection,
  accountNames,
  { type = '', owner = '', search = '', limit = 500 } = {}
) {
  const names = viewerAccountNames(accountNames);
  if (names.length === 0) {
    return [];
  }

  const safeLimit = Math.max(1, Math.min(2000, Number(limit) || 500));
  const filters = [];
  const params = [...names];

  if (type) {
    filters.push('i.item_type_name = ?');
    params.push(type);
  }
  if (owner) {
    filters.push('i.account_name = ?');
    params.push(owner);
  }
  if (search) {
    filters.push('i.item_name LIKE ?');
    params.push(`%${search.replace(/[\\%_]/g, (character) => `\\${character}`)}%`);
  }

  const where = filters.length ? `WHERE ${filters.join(' AND ')}` : '';

  const [rows] = await connection.execute(
    `SELECT i.id,
            i.item_name,
            i.item_id,
            i.icon_key,
            i.item_count,
            i.quality,
            i.quality_name,
            i.item_type_name,
            i.trait,
            i.armor_type,
            i.weapon_type,
            i.apparel_slot,
            i.equip_type,
            i.location_label,
            i.last_dumped_at,
            i.account_name,
            i.character_name,
            (icon.icon_key IS NOT NULL) AS has_icon,
            GROUP_CONCAT(DISTINCT g.guild_name ORDER BY g.guild_name SEPARATOR ', ') AS guild_names
       FROM itemshare_shared_items i
       JOIN itemshare_shared_item_guilds ig ON ig.item_id = i.id
       JOIN itemshare_account_guilds viewer
            ON viewer.guild_id = ig.guild_id
           AND viewer.account_name IN (${inClause(names)})
       LEFT JOIN itemshare_guilds g ON g.guild_id = ig.guild_id
       LEFT JOIN itemshare_item_icons icon ON icon.icon_key = i.icon_key
       ${where}
      GROUP BY i.id, i.item_name, i.item_id, i.icon_key, i.item_count, i.quality,
               i.quality_name, i.item_type_name, i.trait, i.armor_type, i.weapon_type,
               i.apparel_slot, i.equip_type, i.location_label, i.last_dumped_at,
               i.account_name, i.character_name, icon.icon_key
      ORDER BY i.item_type_name ASC, i.item_name ASC, i.account_name ASC
      LIMIT ${safeLimit}`,
    params
  );

  return rows;
}

/**
 * Links an ESO account to the signed-in person.
 * @returns {Promise<'linked'|'already-linked'|'conflict'>}
 */
export async function linkAccountToUser(connection, { accountName, userId }) {
  const [rows] = await connection.execute(
    'SELECT user_id FROM itemshare_accounts WHERE eso_account_name = ?',
    [accountName]
  );

  const existing = rows[0];

  if (existing && existing.user_id && Number(existing.user_id) !== Number(userId)) {
    return 'conflict';
  }

  if (existing) {
    await connection.execute(
      'UPDATE itemshare_accounts SET user_id = ? WHERE eso_account_name = ?',
      [userId, accountName]
    );
    return 'already-linked';
  }

  await connection.execute(
    'INSERT INTO itemshare_accounts (eso_account_name, user_id) VALUES (?, ?)',
    [accountName, userId]
  );
  return 'linked';
}
