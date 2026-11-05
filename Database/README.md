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
- An upload **fills in, then prunes**. Every item in the file is written where it
  already stands - same row, same id, same first-seen date - and only afterwards are
  the rows the file no longer mentions deleted. Nothing blinks out of the list
  between one upload and the next, and anything pointing at a row (a guild link, an
  open request) keeps pointing at it. Un-sharing an item in game and re-uploading
  still removes it here. The mark the prune uses is `upload_id`: everything the file
  mentions is stamped with the new upload, so a row still holding an older id was
  not in this file. All of it happens in one transaction.
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
guild any of those accounts belongs to, with the owning account, character,
quantity, item type and trait.

The list is split into tabs - *Request* (other people's items), *Offering* (your own)
and *Your requests* - so no row needs a `You` tag: the tab already says whose it is.

Rows are grouped by type (Weapon, Apparel, Furnishing, …) and alphabetical inside
each group by default. The **Item** and **Added** headings sort the list, and clicking
one again reverses it; sorting by either drops the type headings, so the order you
asked for is the order you get.

`Added` is the day that item first entered its owner's shared list - the day they
offered it. It is deliberately not the day of their latest upload, so re-uploading a
long-held item does not make it look new.

`Location` is where the owner's copy was the last time they saved - `Bank`,
`Inventory (Character Name)`, `Furnishing Vault (...)`. It is a column rather than a
sentence in the detail line because it is something people scan for, and it follows
the item: an upload rewrites it, so moving something in game and saving again updates
it here without anyone touching the site.

A filter bar above the table narrows the list by item type, owner account, how
recently it was added (the last day, 7 days, 30 days) and item name. Filters are
ordinary query parameters (`/?type=Weapon&owner=@Name&added=7&q=sword&sort=added&dir=desc`),
so a filtered, sorted view can be bookmarked.

The list refreshes automatically after an upload, and a request being cancelled
reaches the person who asked the same way.

Visibility joins three tables: the item's own guilds
(`itemshare_shared_item_guilds`), the viewer's guilds
(`itemshare_account_guilds`), and the guild names (`itemshare_guilds`). Nothing
is readable without a guild overlap, so an account in no guilds sees nothing.

## Asking for something

Every row in *Request* has a **Request** button. It opens a small dialog offering one,
the whole stack, or any number in between. The largest number offered is what is
still unspoken for rather than the raw `item_count`, so two people can never be
promised the same five. A standing offer has no count behind it, so it always means one copy, and it can
only be asked for while the owner has one in a bag right now. When they do not, the
row says *Not in a bag right now* instead of offering a button that cannot be
honoured, and the server refuses the same request either way.

When only one copy can be given - a stack of one, or the last copy of a bigger stack
- there is nothing to choose, so the button sends the request for that one straight
away rather than opening a dialog whose only answer is "one". Leaving the page with
it unsent still raises the usual warning, and it can be cancelled from *Your
requests*.

A request is a **reservation, not a delivery**, and it reserves from the moment it is
picked rather than from the moment it is sent:

- The list has two quantity columns. **Free** is what nobody is holding, **Held** is
  what a request is standing on, and they always add up to what the owner reported.
  Once nothing is free the button becomes *All held*.
- Asking puts the item in a **basket** with a five minute clock (`REQUEST_HOLD_MINUTES`).
  Everything picked shares one timer, and adding or removing anything starts it again.
  *Your requests* shows the countdown next to a **Send requests** button; until it is
  pressed nobody has been asked, but the copies are held, so two people cannot be
  offered the same one.
- When the clock runs out the basket is cancelled with the reason shown, and the
  items are free again. That happens **server-side on a timer**
  (`REQUEST_SWEEP_INTERVAL_MS`, default 10 s), so it also happens to someone who
  requested an item and closed the tab: nothing is held because a browser forgot to
  tell anyone. Every query also treats a lapsed reservation as free, so the
  arithmetic is right even between two sweeps.
- Each item in the basket can be removed before sending, which also restarts the
  clock for the rest. A removed row stays where it was, marked **Cancelled**, for
  about half a minute, then fades out and the rows below it close up. It is not added
  to the history, because it was never sent.
- After every upload the owner's requests are settled oldest first, because they were
  promised first. One whose item has gone, or whose turn the stack no longer reaches,
  is cancelled with a sentence the requester can read: "Ghost Bane now has 2 of
  Briarheart Dagger, which does not cover the 3 you asked for." A row that merely
  moved (the item changed character, or was re-shared) is re-pointed by
  `item_key_hash` instead of being cancelled.
- *Your requests* holds the basket and, below it, **Already sent**: a history of what
  you actually asked for, newest first and oldest last. Each row carries the item's
  picture and trait, and anything that closed says why. When a live update brings back
  a cancellation the page also says so in a toast, once: opening the page does not
  replay old ones.
- A requester can cancel their own waiting request. The owner sees "2 asked for" on
  the row.
- Leaving the page with a basket still holding items raises a styled warning
  ("You have unsubmitted item requests") before following a link, with the browser's
  own unload prompt as the fallback for closing the tab or posting a form.

## Making it look busy: simulated people

`npm run simulate` (in `Database/`) drives three made-up people through the same store
code the site uses, so a list with one real owner has traffic on it: they ask for items
in *your* list, leave some in their baskets, cancel others, and re-save their files so
items move between bags, the bank and the furnishing vault and stack counts change.
Every round ends by asking the running site to announce what happened.

```
npm run simulate                        one round
npm run simulate -- --rounds 12 --every 4000
npm run simulate -- --only requests     just asking and cancelling
npm run simulate -- --only items        just re-saving files
npm run simulate -- --reset             remove the simulated people and their data
```

Set `SIMULATE_TOKEN` in `Database/.env` and restart the site to have open pages refresh
as it runs; without it the data still changes and you see it on the next refresh. The
three people are identified by a `simulator` identity, so they are the same people on
every run, and `--reset` takes all of them out again.

Their activity appears in the log as a sentence rather than a revision number:

```
Live: @DemoTrader asked @YourAccount for 2 x Soul Gem; @QuietSmith saved again: 2 items, now in Bank / Craft Bag - 1 open page told to refresh
Requests: 3 unsent requests timed out, so the items are free again
```

"1 open page told to refresh" is the delivery detail: how many browsers were connected
over the socket at that moment, each of which refetches its own list. It says nothing
about how many people use the site - closing the page drops it to zero.

Everything here is recorded server-side; the browser never decides how many are
left. Handing an item over in game is not recorded yet, so nothing moves from
*Waiting* to *Handed over* and nobody is emailed - see *Not done yet*.

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

## Staying signed in

Signing in is meant to be a one-off. Sessions are stored in MariaDB
(`itemshare_sessions`) rather than in process memory, so a restart, a crash or a
redeploy no longer signs everyone out while their cookie still says otherwise, and
the cookie lasts `SESSION_TTL_DAYS` (default 30) with `rolling` renewal, so a person
who visits now and then never sees a login screen.

The avatar shows the picture from the sign-in provider when there is one: Google's
`picture` claim, or Discord's avatar hash turned into a CDN URL (animated avatars
keep their GIF form). Only https URLs are stored, and if a picture stops loading the
initials underneath show instead, so the header never depends on a third party.

The only way out is the avatar menu in the header: it opens a panel with the name and
provider and holds **Log out**. That posts to
`/logout`, which deletes the session row and clears the cookie. Without JavaScript
the form still submits.

Expired rows are pruned at startup and twice a day. If MariaDB is unreachable the
session layer reports "no session" instead of failing the request, so the page still
renders its database-error card rather than a blank screen.

## Command line equivalent

```bash
node ingest.js ../SampleData/Sample1/ItemShareSavedVars.lua           # import
node ingest.js ../SampleData/Sample1/ItemShareSavedVars.lua --dry-run  # parse only
node backfillIcons.js --dry-run                                       # what lacks an icon
node backfillIcons.js --limit 20 --passes 5                            # retry missing icons
node backfillIcons.js --id 223698                                      # does UESP know this item?
```

`backfillIcons.js` reads `Database/.env` by absolute path, exactly like the site,
so it works from any directory. It runs the same pass the timer runs, with the
same per-pass ceiling, and `--id` asks UESP about one item without opening a
database at all - the quickest way to tell a refused lookup from a page that no
longer embeds the image.

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
| `itemshare_always_share` | Standing offers: items the owner will hand out whenever they have one |
| `itemshare_always_share_guilds` | Which guilds each standing offer is shown in |
| `itemshare_requests` | Who asked for what, how many, and how it ended |

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

Sending mail is not wired up yet: requests are recorded, but nobody is emailed
about them. When that lands it will need mail credentials in `.env` (SMTP or an API
key); nothing about the capture side needs to change.

## Not done yet (next step)

**Handing a request over.** Requests are recorded, reserved and cancelled correctly,
but nothing yet marks one as delivered, and an owner is not told that somebody asked:
no mail, and no in-page notice of their own. `itemshare_requests.status`
(`pending` / `cancelled` / `fulfilled` / `declined`) and `closed_at` are in place for
that step.

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

## Live updates

The item list updates itself when somebody else uploads. Socket.IO runs on the same
HTTP server as the site, under `/socket.io`, so there is no second port and no extra
proxy rule.

The socket deliberately carries **no item data**: it only says "something you can see
has changed". Each open page then re-fetches `/items` — the same query the page was
built from, with the viewer's own session cookie — and swaps just that region into
the page. Visibility therefore stays in exactly one place (`lib/store.js`), and a
mistake in the push path cannot show another guild's items.

Uploads broadcast after the rows are committed, and again if new icons were fetched.
The icon backfill announces its own repairs, so a picture that was missing earlier
appears without a reload.

### Staying where you were

A redraw that jumps to the top of the table is worse than a stale row, so every row
is marked with its stable `item_key_hash` (`data-item-key="inv:…"` / `"offer:…"`).
Before the swap the page notes which row crosses a line about a third of the way
down the viewport and how far it sits from that line. After the swap it finds the
same row and scrolls by the difference, so the reader keeps their exact place.

If that item was removed, the row that moved up into its slot is used instead: the
reader lands one item down the list rather than at the top. If the change adds or
removes an item type or an owner, the filter controls would be out of date, so the
page reloads and restores the same row from session storage.

### One process

The broadcast is in-process. Running several Node processes against the same
database needs a shared Socket.IO adapter (`@socket.io/mysql`), because an upload
handled by one process would otherwise never reach the sockets held by another.

## Using the site

- **Filters apply as you change them.** Type, "Offered by" and the item name box
  update the list immediately; there is no Apply button, and "Clear filters" is a
  button instead of a link. The address bar keeps up, so a reload or a live update
  uses the same filters. With scripts off, the form's own submit button appears and
  everything still works.
- **Uploads start on their own.** Dropping `ItemShareSavedVars.lua` on the upload
  card, or picking it with the file chooser, uploads it straight away. A toast says
  the file has been received and is being processed, then summarises what landed.
- **Three tabs.** *Request* is what other people in your guilds are sharing,
  *Offering* is your own items and your own standing offers, and *Your requests* is
  what you asked others for and what became of it. That split is why no row carries a
  "You" tag - which tab it is in already says whose it is.
- **Sort by clicking a heading.** *Item* and *Added* sort the table, and clicking the
  same heading again reverses it. *Added* also has its own filter (day / 7 days /
  30 days) for finding what is new.
- **Ask for what you need.** Requestable rows have a Request button: one, the whole
  stack, or any number up to what is still unclaimed. What you ask for is held for
  you, shown under *Your requests*, and cancelled with a reason if the owner's file
  stops containing the item.
- **Nothing is listed before your first upload.** Items are shared through guilds,
  so until your account and its guilds are known there is nothing anyone could
  legitimately show you.
- Guild names are not repeated on every row: a row is there *because* of a shared
  guild. The guild list at the top shows names only, no ids.

## Previewing it with a fake guildmate

A saved variables file cannot act as a second person - every account in an uploaded
file is linked to whoever uploaded it. To see the Request tab before someone else
really uploads, seed a demo account instead:

```bash
npm run seed:demo                      # into every guild you are in
npm run seed:demo -- --account @Name   # only that account's guilds
npm run seed:demo -- --guild 761817    # only that guild
npm run seed:demo -- --remove          # take it all out again
```

It writes `@DemoTrader` with nine items and three standing offers into guilds that
already exist, and never links that account to your login, so its items appear under
Request exactly like a real guildmate's would. The items are deliberately spread
across Weapon, Apparel, Jewelry, Consumable, Material, Furnishing and Style, with
two of them in a furnishing vault, so the location line and the type grouping have
something to show.
