// Database operations shared by the CLI (ingest.js) and the website (server.js).
// An upload is a full replace: the uploading account's guild membership and item
// list are both rewritten from the file, inside one transaction.

import { createHash } from 'node:crypto';

const ITEM_COLUMNS = [
  'account_name',
  'character_name',
  'item_key',
  'item_key_hash',
  'item_link',
  'item_name',
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
async function replaceOneAccount(connection, { accountName, items, guilds, sourceFile, discordUserId }) {
  const totalQuantity = items.reduce((sum, item) => sum + (item.itemCount || 0), 0);

  await upsertGuilds(connection, guilds);
  await replaceAccountGuilds(connection, accountName, guilds);

  const [upload] = await connection.execute(
    `INSERT INTO itemshare_uploads
       (account_name, character_name, source_file, item_count, total_quantity, guild_count, discord_user_id)
     VALUES (?, ?, ?, ?, ?, ?, ?)`,
    [accountName, '*', sourceFile || '', items.length, totalQuantity, guilds.size, discordUserId || '']
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
export async function replaceAccounts(connection, accounts, { sourceFile, discordUserId } = {}) {
  await connection.beginTransaction();

  try {
    const results = [];
    for (const account of accounts) {
      results.push(await replaceOneAccount(connection, { ...account, sourceFile, discordUserId }));
    }
    await connection.commit();
    return results;
  } catch (error) {
    await connection.rollback();
    throw error;
  }
}

/**
 * Records or refreshes the Discord profile captured at sign-in. The email is
 * stored so a later item request can notify the owner without asking again.
 */
export async function upsertDiscordUser(connection, { id, username, email, emailVerified }) {
  await connection.execute(
    `INSERT INTO itemshare_discord_users
       (discord_user_id, discord_username, discord_email, email_verified)
     VALUES (?, ?, ?, ?)
     ON DUPLICATE KEY UPDATE
       discord_username = VALUES(discord_username),
       discord_email = IF(VALUES(discord_email) = '', discord_email, VALUES(discord_email)),
       email_verified = VALUES(email_verified)`,
    [String(id), username || '', email || '', emailVerified ? 1 : 0]
  );
}

/** Email for an ESO account's linked Discord user, for request notifications. */
export async function getOwnerNotificationEmail(connection, accountName) {
  const [rows] = await connection.execute(
    `SELECT u.discord_email, u.email_verified
       FROM itemshare_accounts a
       JOIN itemshare_discord_users u ON u.discord_user_id = a.discord_user_id
      WHERE a.eso_account_name = ?`,
    [accountName]
  );

  const row = rows[0];
  if (!row || !row.discord_email) {
    return null;
  }
  return { email: row.discord_email, verified: Number(row.email_verified) === 1 };
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
    `SELECT discord_user_id, discord_username FROM itemshare_accounts WHERE eso_account_name = ?`,
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

/** Every ESO account linked to one Discord login, oldest link first. */
export async function getViewerAccounts(connection, discordUserId) {
  const [rows] = await connection.execute(
    `SELECT eso_account_name, discord_user_id, discord_username, linked_at
       FROM itemshare_accounts
      WHERE discord_user_id = ?
      ORDER BY linked_at ASC, eso_account_name ASC`,
    [String(discordUserId)]
  );
  return rows;
}

function viewerAccountNames(accountNames) {
  return [...new Set((accountNames || []).map((name) => String(name || '').trim()).filter(Boolean))];
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
      WHERE ag.account_name IN (${names.map(() => '?').join(', ')})
      GROUP BY ag.guild_id, g.guild_name
      ORDER BY guild_name ASC, ag.guild_id ASC`,
    names
  );
  return rows;
}

/**
 * Shared items offered to any guild the viewer belongs to, alphabetical.
 * Ownership comes from the item's own account, visibility from the guild overlap.
 */
export async function getSharedItemsForViewer(connection, accountNames, limit = 500) {
  const names = viewerAccountNames(accountNames);
  if (names.length === 0) {
    return [];
  }

  const safeLimit = Math.max(1, Math.min(2000, Number(limit) || 500));

  const [rows] = await connection.execute(
    `SELECT i.id,
            i.item_name,
            i.item_count,
            i.quality,
            i.quality_name,
            i.item_type_name,
            i.trait,
            i.account_name,
            i.character_name,
            GROUP_CONCAT(DISTINCT g.guild_name ORDER BY g.guild_name SEPARATOR ', ') AS guild_names
       FROM itemshare_shared_items i
       JOIN itemshare_shared_item_guilds ig ON ig.item_id = i.id
       JOIN itemshare_account_guilds viewer
            ON viewer.guild_id = ig.guild_id
           AND viewer.account_name IN (${names.map(() => '?').join(', ')})
       LEFT JOIN itemshare_guilds g ON g.guild_id = ig.guild_id
      GROUP BY i.id, i.item_name, i.item_count, i.quality, i.quality_name,
               i.item_type_name, i.trait, i.account_name, i.character_name
      ORDER BY i.item_name ASC, i.account_name ASC
      LIMIT ${safeLimit}`,
    names
  );

  return rows;
}

/**
 * Links a Discord user to an ESO account on first upload.
 * @returns {Promise<'linked'|'already-linked'|'conflict'>}
 */
export async function linkDiscordAccount(connection, { accountName, discordUserId, discordUsername }) {
  const [rows] = await connection.execute(
    `SELECT discord_user_id FROM itemshare_accounts WHERE eso_account_name = ?`,
    [accountName]
  );

  const existing = rows[0];

  if (existing && existing.discord_user_id !== discordUserId) {
    return 'conflict';
  }

  if (existing) {
    await connection.execute(
      `UPDATE itemshare_accounts SET discord_username = ? WHERE eso_account_name = ?`,
      [discordUsername || '', accountName]
    );
    return 'already-linked';
  }

  await connection.execute(
    `INSERT INTO itemshare_accounts (eso_account_name, discord_user_id, discord_username) VALUES (?, ?, ?)`,
    [accountName, discordUserId, discordUsername || '']
  );
  return 'linked';
}
