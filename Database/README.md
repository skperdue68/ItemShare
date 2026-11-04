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
| `DISCORD_REDIRECT_URI` | `$BASE_URL/auth/discord/callback` | Must match the Discord app's redirect |
| `GOOGLE_CLIENT_ID` | empty | Google OAuth client ID |
| `GOOGLE_CLIENT_SECRET` | empty | Google OAuth client secret |
| `GOOGLE_REDIRECT_URI` | `$BASE_URL/auth/google/callback` | Must match the Google client's redirect URI |
| `BASE_URL` | `http://localhost:$PORT` | Public address used to build the default redirect URIs |

The MariaDB account needs CREATE privileges on first run (to create the database
and `itemshare_*` tables). After that, ordinary read/write is enough.

### Sign-in providers (optional, either or both)

**Discord**

1. Create an application at <https://discord.com/developers/applications>.
2. Under **OAuth2**, add the redirect URL `http://localhost:3000/auth/discord/callback`
   (change the host/port if you run the site elsewhere).
3. Copy the Client ID and Client Secret into `.env` as `DISCORD_CLIENT_ID` and
   `DISCORD_CLIENT_SECRET`, and set `DISCORD_REDIRECT_URI` to the same URL.

**Google**

1. Create an OAuth client at <https://console.cloud.google.com/apis/credentials>
   (type *Web application*).
2. Add `http://localhost:3000/auth/google/callback` under **Authorised redirect
   URIs**.
3. Copy the client ID and secret into `.env` as `GOOGLE_CLIENT_ID` and
   `GOOGLE_CLIENT_SECRET`, and set `GOOGLE_REDIRECT_URI` to the same URL.

Google returns the person's name, email address and whether that address is
verified, so it supplies everything the site stores for Discord. Both providers
give a stable numeric/opaque subject id, which is what identifies the login.

Each provider is independent: configuring one shows only that button, configuring
both shows two. With neither, the site still accepts uploads, but nothing is
linked to a person and uploads are anonymous.

**One person, either button.** Identities live in `itemshare_user_identities` and
point at a row in `itemshare_users`. When someone signs in with a provider for the
first time and that provider reports a *verified* email already known to another
identity, the new identity joins the existing person instead of creating a second
empty one. That way the ESO accounts linked through Discord stay visible after
signing in with Google. Unverified addresses are never used to merge people.

## How an upload behaves

- The ESO account names are read from the file itself, so there is nothing to type.
  Every account in the file is imported: all of them are linked to the same Discord
  user, and each one keeps its own item list and guild membership. Either every
  account is replaced or, if anything fails, none of them are.
- Item rows keep the account they came from (`itemshare_shared_items.account_name`),
  so a member who plays several ESO accounts can see which account holds each item.
- ItemShare 1.4.0 also saves the item id and the game's icon path for each entry,
  which is what feeds the thumbnails and the UESP links. Files from older versions
  import fine; those rows simply carry no icon until they are uploaded again.
- Every upload **fully replaces** the items stored for that account: the previous
  rows are deleted and rewritten in one transaction, so un-sharing items in game
  and re-uploading removes them here too.
- Quantity is stored per item in `item_count`, and the upload total in
  `itemshare_uploads.total_quantity`.
- The first upload from a signed-in person links their identity to the ESO
  account in `itemshare_accounts`. If the ESO account is already linked to a
  different person the upload is refused rather than silently reassigned.
- Guild membership is replaced in full as well (`itemshare_account_guilds`), so
  leaving a guild in game removes that guild's visibility after the next upload.

## Seeing shared items

Any Discord or Google account can sign in; it only establishes identity. Once one or more
ESO accounts are linked to it, the page lists every shared item offered to any
guild any of those accounts belongs to, alphabetically, with the owning account,
character, quantity, item type, trait and the guilds it is visible through. Items
held by any of the viewer's own accounts are tagged `You`, and the summary card
shows each linked account with its own totals and guilds.

Items are grouped by type (Weapon, Apparel, Furnishing, …) and alphabetical inside
each group. A filter bar above the table narrows the list by item type, owner
account and item name. Filters are ordinary query parameters
(`/?type=Weapon&owner=@Name&q=sword`), so a filtered view can be bookmarked.

The list refreshes automatically after an upload.

Visibility joins three tables: the item's own guilds
(`itemshare_shared_item_guilds`), the viewer's guilds
(`itemshare_account_guilds`), and the guild names (`itemshare_guilds`). Nothing
is readable without a guild overlap, so an account in no guilds sees nothing.

## Item icons

ItemShare 1.4.0 records the game's icon path for every shared item
(`/esoui/art/icons/…dds`). Browsers cannot draw DDS textures, so after an upload
the site fetches each *new* icon once from UESP's PNG mirror (`esoicons.uesp.net`)
and stores the bytes in `itemshare_item_icons`. Pages then serve them from
`/icons/<key>.png`, so browsing never contacts UESP and a later upload of the same
item fetches nothing at all. Fetching is best effort: a failure is logged, that
row simply shows no thumbnail, and it is retried later.

When the client records no icon path at all - a share written before 1.5.0, or a
client without `GetItemIcon` - the item id is used instead: one request to
`esoitem.uesp.net/itemLink.php?itemid=...` yields the same icon path, which then
travels the ordinary cache route, so nothing downstream knows the difference.

Both steps run at upload and again on a timer, because a blank icon never repairs
itself otherwise. `ICON_BACKFILL_INTERVAL_MS` (default 30 minutes) sets how often
the retry pass runs and `ICON_BACKFILL_BATCH` (default 20) how many item ids one
pass may look up; `ICON_BACKFILL_INTERVAL_MS=0` turns the pass off. A pass looks
only at ids that have no path anywhere, so an item UESP is still refusing simply
stays queued for the next one.

Item names also link to `esoitem.uesp.net/itemLink.php?itemid=…` when the id is
known. That link is a convenience only: type, quality, trait and icon all come
from the addon, not from UESP.

## Command line equivalent

```bash
node ingest.js ../SampleData/Sample1/ItemShareSavedVars.lua           # import
node ingest.js ../SampleData/Sample1/ItemShareSavedVars.lua --dry-run  # parse only
```

The command line importer handles multi-account files exactly like the website:
every account is replaced inside one transaction, each with its own guilds.

## Tables

| Table | Purpose |
| --- | --- |
| `itemshare_uploads` | One row per import: account, file, item count, total quantity, person |
| `itemshare_accounts` | ESO account → person link; several accounts may belong to one person |
| `itemshare_users` | One row per person: display name, email, verified flag |
| `itemshare_user_identities` | `(provider, provider_user_id)` → person, for Discord and Google |
| `itemshare_discord_users` | Legacy Discord profile table, superseded by `itemshare_users` |
| `itemshare_item_icons` | Cached PNG bytes per icon path, fetched once from UESP |
| `itemshare_guilds` | Guild ID → name, the stable identity for a guild |
| `itemshare_shared_items` | One row per shared item per character |
| `itemshare_shared_item_guilds` | Which guilds each shared item is offered to |

Guild IDs come from the addon (1.3.0+; `/ishare guilds` shows what was captured).
Names are display only, because names are neither unique nor stable.

## Email addresses

Discord sign-in requests the `identify` and `email` scopes; Google sign-in
requests `openid email profile`. Either way the address lands in
`itemshare_users.email` together with the provider's own verified flag, and is
refreshed at every sign-in, so it stays correct if someone changes it.

The address is personal data and is only ever read server-side. It is shown back
to its own owner on the page and is never included in item lists or any other
member's view. `getOwnerNotificationEmail(connection, accountName)` is the single
server-side helper for finding an owner's address when a request needs to notify
them.

Sending mail is not wired up yet, because the request flow does not exist. When it
lands it will need mail credentials in `.env` (SMTP or an API key); nothing about
the capture side needs to change.

## Not done yet (next step)

Requesting items is out of scope. The schema is shaped for it: a request would
reference `itemshare_shared_items.id` and reuse the same guild overlap check as
the item list.

Item icons are not shown. ESO icon art is keyed by icon name rather than the item
ID in the link, so quality colour, type and trait are displayed instead of
guessing at a texture.

## What is captured from the game

Addon 1.5.0 records everything the running client will answer for a shared item:
name, item id, icon path, type, quality, trait, apparel slot, equip type, armour
and weapon type, handedness, furniture data id, location, **set name,
enchantment, crafting style, required level and champion points, and flavour
text**.

Every getter is probed through `pcall` and its results are type-checked, so a
field the current API does not expose is simply absent instead of raising an
error while an item is being added to the list. That also means the same build
keeps working when ZeniMax renames or reshapes one of these functions.

Each probe's raw scalar results are stored with the item and end up in the
`eso_data` JSON column, keyed by getter name:

```json
{"GetItemLinkSetInfo": [true, "Briarheart", 5, 3, 5], "GetItemLinkRequiredLevel": [50]}
```

So detail the site does not render yet is already in the database and can be
surfaced later without asking anyone to upload again. UESP enrichment is then
only used for what the client left empty — the icon PNG mirror is the existing
example.

## Troubleshooting startup

The server prints the `.env` file it actually read, so a setting that "is in there" but has no effect is usually a file in the wrong place — a `.env` in the repo root instead of `Database/`, or Windows Notepad saving it as `.env.txt`.

**`SESSION_SECRET is not set ... using an ephemeral secret`** — no `Database/.env`, or the key is missing from it. Every restart invalidates existing sessions, so sign-in is lost. Make a secret and paste the printed line into `Database/.env`:

```bash
node -e "console.log('SESSION_SECRET=' + require('crypto').randomBytes(32).toString('hex'))"
```

**`Server requests authentication using unknown plugin auth_gssapi_client`** — the MariaDB account is configured for Kerberos/GSSAPI, which `mysql2` cannot use. On the MariaDB host, as an administrator, give it an ordinary password login:

```sql
SELECT user, host, plugin FROM mysql.user WHERE user = 'itemshare';
ALTER USER 'itemshare'@'%' IDENTIFIED VIA mysql_native_password USING PASSWORD('the-password-from-.env');
FLUSH PRIVILEGES;
```

Leave the password inside `PASSWORD(...)` as the plain text one from `Database/.env`; MariaDB hashes it. If `Access denied for user ...` comes back instead, the plugin is fine and the password or the host grant is wrong — check that the account may connect from this machine's address, not just `localhost`.

**`Sign-in providers: none`** — `DISCORD_CLIENT_ID`/`DISCORD_CLIENT_SECRET` (or the `GOOGLE_*` pair) are empty. Discord also needs `DISCORD_REDIRECT_URI` registered under OAuth2 -> Redirects in the Discord application, and the same URL must be the one the browser is sent to.
