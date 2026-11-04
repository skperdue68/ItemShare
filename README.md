# ItemShare database and upload site

This folder turns `ItemShareSavedVars.lua` into rows in MariaDB, either from the
website or from the command line. The Google Sheets upload in `../GoogleSheets`
is untouched and keeps working exactly as before.

ItemShare always uses its own database (`itemshare` by default). It never reads
`MARIADB_DATABASE`, so it cannot write into the GuildSync schema even though both
live on the same MariaDB server.

## Setup

```bash
cd Database
npm install
cp .env.example .env      # MariaDB credentials, session secret, Discord app
npm start                 # http://localhost:3000
```

## Startup and migrations

On startup the site:

1. connects to MariaDB using the settings below (the database itself is **not**
   selected yet, so it can be created);
2. runs `CREATE DATABASE IF NOT EXISTS` for `ITEMSHARE_DATABASE`;
3. applies every `.sql` file in `migrations/` that has not been applied before.

Applied files are recorded in `itemshare_schema_migrations`, so each one runs
exactly once. To add a migration later, drop a new file in `migrations/` with a
sortable name (`004_add_requests.sql`) and restart: it will be applied on the
next boot and never again. Recorded migrations are not re-run, so they no longer
need to be idempotent, though `CREATE TABLE IF NOT EXISTS` is still a good habit.

If the database is unreachable at startup the site still boots, logs the reason,
and shows the error on the page instead of crashing. Connections made per request
connect directly to the existing database rather than re-running migrations.

## Settings

All of these live in `Database/.env` (see `.env.example`); every one has a
default, so the site runs without a `.env` on a local root/no-password server.

| Variable | Default | Purpose |
| --- | --- | --- |
| `MARIADB_HOST` | `127.0.0.1` | MariaDB host |
| `MARIADB_PORT` | `3306` | MariaDB port |
| `MARIADB_USER` | `root` | Account used for connections |
| `MARIADB_PASSWORD` | empty | Password for that account |
| `ITEMSHARE_DATABASE` | `itemshare` | **ItemShare's own database**, created if missing. `MARIADB_DATABASE` is deliberately ignored so the GuildSync schema can never be touched |
| `PORT` | `3000` | Web port |
| `SESSION_SECRET` | random per boot | Signs the session cookie; set it so logins survive a restart |
| `DISCORD_CLIENT_ID` | empty | Discord application ID |
| `DISCORD_CLIENT_SECRET` | empty | Discord application secret |
| `DISCORD_REDIRECT_URI` | `http://localhost:$PORT/auth/discord/callback` | Must match the Discord app's redirect |

The MariaDB account needs CREATE privileges on first run (to create the database
and `itemshare_*` tables). After that, ordinary read/write is enough.

### Discord sign-in (optional)

1. Create an application at <https://discord.com/developers/applications>.
2. Under **OAuth2**, add the redirect URL `http://localhost:3000/auth/discord/callback`
   (change the host/port if you run the site elsewhere).
3. Copy the Client ID and Client Secret into `.env` as `DISCORD_CLIENT_ID` and
   `DISCORD_CLIENT_SECRET`, and set `DISCORD_REDIRECT_URI` to the same URL.

Without those variables the site still accepts uploads, but nothing is linked to a
Discord user and uploads are anonymous.

## How an upload behaves

- The ESO account name is read from the file itself, so there is nothing to type.
  A file containing more than one account is rejected.
- Every upload **fully replaces** the items stored for that account: the previous
  rows are deleted and rewritten in one transaction, so un-sharing items in game
  and re-uploading removes them here too.
- Quantity is stored per item in `item_count`, and the upload total in
  `itemshare_uploads.total_quantity`.
- The first upload from a signed-in user links that Discord account to the ESO
  account in `itemshare_accounts`. If the ESO account is already linked to a
  different Discord user the upload is refused rather than silently reassigned.
- Guild membership is replaced in full as well (`itemshare_account_guilds`), so
  leaving a guild in game removes that guild's visibility after the next upload.

## Seeing shared items

Any Discord account can sign in; it only establishes identity. Once an ESO
account is linked, the page lists every shared item offered to any guild that
account belongs to, alphabetically, with the owner, quantity, item type, trait
and the guilds it is visible through. Items the viewer owns are tagged `You`.

The list refreshes automatically after an upload.

Visibility joins three tables: the item's own guilds
(`itemshare_shared_item_guilds`), the viewer's guilds
(`itemshare_account_guilds`), and the guild names (`itemshare_guilds`). Nothing
is readable without a guild overlap, so an account in no guilds sees nothing.

## Command line equivalent

```bash
node ingest.js ../SampleData/Sample1/ItemShareSavedVars.lua           # import
node ingest.js ../SampleData/Sample1/ItemShareSavedVars.lua --dry-run  # parse only
```

## Tables

| Table | Purpose |
| --- | --- |
| `itemshare_uploads` | One row per import: account, file, item count, total quantity, Discord user |
| `itemshare_accounts` | ESO account → Discord user link |
| `itemshare_guilds` | Guild ID → name, the stable identity for a guild |
| `itemshare_shared_items` | One row per shared item per character |
| `itemshare_shared_item_guilds` | Which guilds each shared item is offered to |

Guild IDs come from the addon (1.3.0+; `/ishare guilds` shows what was captured).
Names are display only, because names are neither unique nor stable.

## Not done yet (next step)

Requesting items is out of scope. The schema is shaped for it: a request would
reference `itemshare_shared_items.id` and reuse the same guild overlap check as
the item list.

Item icons are not shown. ESO icon art is keyed by icon name rather than the item
ID in the link, so quality colour, type and trait are displayed instead of
guessing at a texture.
