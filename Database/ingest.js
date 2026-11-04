#!/usr/bin/env node
// Stage one importer: reads ItemShareSavedVars.lua and stores its contents in
// the ItemShare MariaDB database. This is the command line equivalent of the
// website upload: each run fully replaces the items and guilds of every ESO
// account found in the file, inside one transaction.
//
//   node ingest.js <path to ItemShareSavedVars.lua> [--dry-run]

import { readFile } from 'node:fs/promises';
import path from 'node:path';
import 'dotenv/config';

import { extractShareData } from './lib/savedVars.js';
import { openDatabase, readConfig } from './lib/database.js';
import { cacheItemIcons } from './lib/icons.js';
import { replaceAccounts } from './lib/store.js';

function groupByAccount(items) {
  const groups = new Map();
  for (const item of items) {
    if (!groups.has(item.accountName)) {
      groups.set(item.accountName, []);
    }
    groups.get(item.accountName).push(item);
  }
  return groups;
}
async function main() {
  const args = process.argv.slice(2);
  const dryRun = args.includes('--dry-run');
  const filePath = args.find((arg) => !arg.startsWith('--'));

  if (!filePath) {
    console.error('Usage: node ingest.js <ItemShareSavedVars.lua> [--dry-run]');
    process.exitCode = 1;
    return;
  }

  const absolutePath = path.resolve(filePath);
  const fileText = await readFile(absolutePath, 'utf8');
  const { guilds, accountGuilds, items, alwaysShare, addonVersions, accounts } =
    extractShareData(fileText);
  const groups = groupByAccount(items);
  const alwaysGroups = groupByAccount(alwaysShare);

  const totalQuantity = items.reduce((sum, item) => sum + item.itemCount, 0);
  const withGuild = items.filter((item) => item.guildIds.length > 0).length;

  console.log(`Read ${absolutePath}`);
  console.log(`  accounts: ${accounts.join(', ') || '(none)'}`);
  console.log(`  shared items: ${items.length} (${withGuild} linked to a guild), total quantity ${totalQuantity}`);
  console.log(`  always shared: ${alwaysShare.length} (not counted as inventory)`);
  console.log(
    `  addon version: ${[...new Set([...addonVersions.values()])].join(', ') || 'not recorded (addon older than 1.8.1)'}`
  );
  console.log(`  guilds: ${guilds.size}`);
  for (const [guildId, guildName] of guilds) {
    console.log(`    ${guildId} ${guildName || '(name not captured yet)'}`);
  }
  for (const accountName of accounts) {
    const memberOf = accountGuilds.get(accountName);
    const accountItems = groups.get(accountName) || [];
    console.log(
      `  ${accountName}: ${accountItems.length} items, guilds ${memberOf && memberOf.size ? [...memberOf.keys()].join(', ') : '(none)'}`
    );
    const alwaysItems = alwaysGroups.get(accountName) || [];
    if (alwaysItems.length) {
      console.log(`    always shared: ${alwaysItems.map((entry) => entry.itemName).join(', ')}`);
    }
  }

  if (dryRun) {
    console.log('Dry run: no database connection opened, nothing written.');
    return;
  }

  const config = readConfig();
  console.log(`Writing to MariaDB database "${config.database}" at ${config.host}:${config.port} as ${config.user}`);

  const { connection, migrations } = await openDatabase(config);

  try {
    const results = await replaceAccounts(
      connection,
      accounts.map((accountName) => ({
        accountName,
        items: groups.get(accountName) || [],
        alwaysShare: alwaysGroups.get(accountName) || [],
        addonVersion: addonVersions.get(accountName) || '',
        guilds: accountGuilds.get(accountName) || new Map(),
      })),
      { sourceFile: path.basename(absolutePath) }
    );

    for (const result of results) {
      console.log(
        `  ${result.accountName}: replaced with ${result.itemCount} items, total quantity ${result.totalQuantity}, ${result.guildCount} guilds, ${result.alwaysShareCount} always shared`
      );
    }
    const iconResult = await cacheItemIcons(connection, items.map((item) => item.icon));
    console.log(
      `  icons: fetched ${iconResult.fetched}, already stored ${iconResult.skipped}, failed ${iconResult.failed}`
    );
    console.log(`Applied migrations: ${migrations.join(', ')}`);
  } finally {
    await connection.end();
  }
}

main().catch((error) => {
  console.error(`ItemShare ingest failed: ${error.message}`);
  process.exitCode = 1;
});
