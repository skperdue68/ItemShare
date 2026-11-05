#!/usr/bin/env node
// ItemShare upload site: signs in with Discord or Google (both optional),
// accepts ItemShareSavedVars.lua and replaces every shared item for every ESO
// account in the file. Styling follows the GuildSync palette.
//
//   npm start            (reads Database/.env)

import { existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import express from 'express';
import session from 'express-session';
import multer from 'multer';
import dotenv from 'dotenv';

import { extractShareData } from './lib/savedVars.js';
import { describeConnectionFailure, openDatabase, readConfig } from './lib/database.js';
import { cacheItemIcons, getCachedIcon, resolveMissingIconPaths } from './lib/icons.js';
import { backfillItemIcons } from './lib/iconBackfill.js';
import {
  fillKnownIcons,
  getAccountSummary,
  getAlwaysShareForViewer,
  getSharedItemsForViewer,
  getViewerAccounts,
  getViewerFacets,
  getViewerGuilds,
  linkAccountToUser,
  replaceAccounts,
  signInWithProvider,
} from './lib/store.js';
import { MariaDbSessionStore } from './lib/sessionStore.js';
import { renderMessagePage, renderPage } from './lib/views.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Read Database/.env by absolute path. The `dotenv/config` shorthand resolves
// .env against the shell's working directory, so starting the server as
// `node Database/server.js` from the repo root found no settings at all and the
// site came up with no session secret and no sign-in providers.
const envPath = path.join(__dirname, '.env');
const hasEnvFile = existsSync(envPath);
dotenv.config({ path: envPath });

const PORT = Number(process.env.PORT || 3000);
const BASE_URL = process.env.BASE_URL || `http://localhost:${PORT}`;

// Both providers speak OAuth2 authorization code, so one flow serves both. Only
// the endpoints, scopes and profile fields differ.
const PROVIDERS = {
  discord: {
    label: 'Discord',
    clientId: process.env.DISCORD_CLIENT_ID || '',
    clientSecret: process.env.DISCORD_CLIENT_SECRET || '',
    redirectUri: process.env.DISCORD_REDIRECT_URI || `${BASE_URL}/auth/discord/callback`,
    authorizeUrl: 'https://discord.com/oauth2/authorize',
    tokenUrl: 'https://discord.com/api/oauth2/token',
    profileUrl: 'https://discord.com/api/users/@me',
    scope: 'identify email',
    readProfile(profile) {
      return {
        providerUserId: String(profile.id),
        username: profile.global_name || profile.username || String(profile.id),
        email: typeof profile.email === 'string' ? profile.email : '',
        emailVerified: profile.verified === true,
        avatarUrl: discordAvatarUrl(profile),
      };
    },
  },
  google: {
    label: 'Google',
    clientId: process.env.GOOGLE_CLIENT_ID || '',
    clientSecret: process.env.GOOGLE_CLIENT_SECRET || '',
    redirectUri: process.env.GOOGLE_REDIRECT_URI || `${BASE_URL}/auth/google/callback`,
    authorizeUrl: 'https://accounts.google.com/o/oauth2/v2/auth',
    tokenUrl: 'https://oauth2.googleapis.com/token',
    profileUrl: 'https://openidconnect.googleapis.com/v1/userinfo',
    scope: 'openid email profile',
    readProfile(profile) {
      return {
        providerUserId: String(profile.sub),
        username: profile.name || profile.email || String(profile.sub),
        email: typeof profile.email === 'string' ? profile.email : '',
        emailVerified: profile.email_verified === true,
        avatarUrl: googleAvatarUrl(profile),
      };
    },
  },
};

for (const [key, provider] of Object.entries(PROVIDERS)) {
  provider.key = key;
  provider.configured = Boolean(provider.clientId && provider.clientSecret);
}

const enabledProviders = Object.values(PROVIDERS)
  .filter((provider) => provider.configured)
  .map(({ key, label }) => ({ key, label }));

if (!process.env.SESSION_SECRET) {
  console.warn(
    `SESSION_SECRET is not set${hasEnvFile ? ` in ${envPath}` : `, and there is no .env at ${envPath}`}; ` +
      'using an ephemeral secret, so everyone is signed out on every restart. Make one with\n' +
      '  node -e "console.log(\'SESSION_SECRET=\' + require(\'crypto\').randomBytes(32).toString(\'hex\'))"\n' +
      'and paste the line into Database/.env.'
  );
}
const SESSION_SECRET = process.env.SESSION_SECRET || `itemshare-${Math.random().toString(36).slice(2)}`;

// Staying signed in is the default, not a trick: sessions are kept in MariaDB so a
// restart no longer signs everyone out, and the cookie lasts a month rather than
// half a day. `rolling` renews it on each visit, so an active member never sees a
// login screen for no reason. Log out from the avatar menu is the only way out
// short of clearing cookies, and SESSION_TTL_DAYS=0 is not a thing - one day is
// the floor.
const SESSION_TTL_DAYS = Math.max(Number(process.env.SESSION_TTL_DAYS || 30) || 30, 1);
const SESSION_TTL_MS = SESSION_TTL_DAYS * 24 * 60 * 60 * 1000;
const sessionStore = new MariaDbSessionStore({ withConnection: withDatabase, ttlMs: SESSION_TTL_MS });

const app = express();
app.use(express.urlencoded({ extended: false }));
app.use(
  session({
    secret: SESSION_SECRET,
    store: sessionStore,
    resave: false,
    saveUninitialized: false,
    rolling: true,
    cookie: { httpOnly: true, sameSite: 'lax', secure: false, maxAge: SESSION_TTL_MS },
  })
);
app.use(express.static(path.join(__dirname, 'public')));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 12 * 1024 * 1024, files: 1 },
});

function readParam(value) {
  return typeof value === 'string' ? value.trim().slice(0, 120) : '';
}

// Every page the site can end up on uses the same shell, so a cancelled sign-in or
// a wrong address looks like ItemShare rather than a line of unstyled text.
function sendPage(req, res, status, options) {
  res
    .status(status)
    .type('html')
    .send(
      renderMessagePage({
        ...options,
        user: req.session ? req.session.user || null : null,
        providers: enabledProviders,
      })
    );
}

// Migrations run once at startup; request connections just use the database.
async function prepareDatabase() {
  const config = readConfig();
  const { connection, migrations } = await openDatabase(config);
  await connection.end();
  console.log(`Database ready: ${config.database} on ${config.host}:${config.port}`);
  console.log(`Migrations: ${migrations.length ? migrations.join(', ') : 'already up to date'}`);
}

// A single short lived connection per request keeps the app simple: the traffic
// here is one person uploading one file at a time.
async function withDatabase(handler) {
  const config = readConfig();
  const { connection } = await openDatabase(config, { migrate: false });
  try {
    return await handler(connection, config);
  } finally {
    await connection.end();
  }
}

app.get('/', async (req, res) => {
  const user = req.session.user || null;
  const filters = {
    type: readParam(req.query.type),
    owner: readParam(req.query.owner),
    search: readParam(req.query.q),
  };

  const base = {
    user,
    providers: enabledProviders,
    accounts: [],
    sharedItems: [],
    alwaysShare: [],
    facets: { types: [], owners: [] },
    filters,
    dbError: null,
  };

  try {
    await withDatabase(async (connection) => {
      // One person can own several ESO accounts; every account linked to them
      // contributes guilds, and the items stay labelled with their account.
      let accountNames = [];

      if (user) {
        const linked = await getViewerAccounts(connection, user.id);
        accountNames = linked.map((row) => row.eso_account_name);
      }

      if (accountNames.length === 0 && Array.isArray(req.session.accountNames)) {
        accountNames = req.session.accountNames.slice();
      }

      if (accountNames.length === 0) {
        return;
      }

      for (const accountName of accountNames) {
        base.accounts.push({
          accountName,
          ...(await getAccountSummary(connection, accountName)),
          guilds: await getViewerGuilds(connection, [accountName]),
        });
      }

      base.facets = await getViewerFacets(connection, accountNames);
      base.sharedItems = await getSharedItemsForViewer(connection, accountNames, filters);
      base.alwaysShare = await getAlwaysShareForViewer(connection, accountNames, filters);
    });
  } catch (error) {
    base.dbError = error.message;
  }

  res.type('html').send(renderPage(base));
});

// Icons were fetched from UESP once, at upload time, and live in MariaDB now.
// Long lived caching headers keep repeat views off the database too.
app.get('/icons/:file', async (req, res) => {
  const match = String(req.params.file || '').match(/^([a-f0-9]{40})\.png$/);

  if (!match) {
    res.status(404).end();
    return;
  }

  try {
    const icon = await withDatabase((connection) => getCachedIcon(connection, match[1]));

    if (!icon) {
      res.status(404).end();
      return;
    }

    res.set('Content-Type', icon.content_type || 'image/png');
    res.set('Cache-Control', 'public, max-age=31536000, immutable');
    res.send(icon.image_data);
  } catch (error) {
    console.warn(`Could not serve icon ${match[1]}: ${error.message}`);
    res.status(500).end();
  }
});

app.get('/auth/:provider', (req, res) => {
  const provider = PROVIDERS[req.params.provider];

  if (!provider || !provider.configured) {
    sendPage(req, res, 503, {
      heading: 'That sign-in option is not set up',
      message: 'This site has no client configured for that provider yet, so there is nothing to sign in with.',
      details: 'Set the provider client id and secret in Database/.env and restart the site.',
    });
    return;
  }

  const params = new URLSearchParams({
    client_id: provider.clientId,
    redirect_uri: provider.redirectUri,
    response_type: 'code',
    scope: provider.scope,
    state: req.sessionID,
  });

  if (provider.key === 'google') {
    params.set('prompt', 'select_account');
  }

  res.redirect(`${provider.authorizeUrl}?${params.toString()}`);
});

app.get('/auth/:provider/callback', async (req, res) => {
  const provider = PROVIDERS[req.params.provider];
  const { code, error } = req.query;

  if (!provider || !provider.configured) {
    sendPage(req, res, 503, {
      heading: 'That sign-in option is not set up',
      message: 'This site has no client configured for that provider yet, so there is nothing to sign in with.',
      details: 'Set the provider client id and secret in Database/.env and restart the site.',
    });
    return;
  }
  if (error || !code) {
    sendPage(req, res, 400, {
      heading: `${provider.label} sign-in did not finish`,
      message: error
        ? 'The provider sent us back without permission, or the prompt was cancelled.'
        : 'The provider sent us back without a code, so there was nothing to verify.',
      details: 'Nothing was changed. You can try again from the header.',
      tone: 'error',
    });
    return;
  }

  try {
    const tokenResponse = await fetch(provider.tokenUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        Accept: 'application/json',
      },
      body: new URLSearchParams({
        client_id: provider.clientId,
        client_secret: provider.clientSecret,
        grant_type: 'authorization_code',
        code: String(code),
        redirect_uri: provider.redirectUri,
      }),
    });

    const token = await tokenResponse.json().catch(() => ({}));

    if (!tokenResponse.ok || !token.access_token) {
      throw new Error(`${provider.label} token exchange failed (${tokenResponse.status})`);
    }

    const userResponse = await fetch(provider.profileUrl, {
      headers: { Authorization: `Bearer ${token.access_token}` },
    });

    if (!userResponse.ok) {
      throw new Error(`${provider.label} profile lookup failed (${userResponse.status})`);
    }

    const profile = await userResponse.json();
    const identity = provider.readProfile(profile);

    const signedIn = await withDatabase((connection) =>
      signInWithProvider(connection, { provider: provider.key, ...identity })
    );

    req.session.user = {
      id: signedIn.userId,
      provider: provider.key,
      providerLabel: provider.label,
      username: identity.username,
      email: identity.email,
      emailVerified: identity.emailVerified,
      // Discarded by the store if it is not https, so whatever is left is safe to
      // put in an img src.
      avatarUrl: signedIn.avatarUrl || identity.avatarUrl || '',
    };
    req.session.accountNames = null;

    res.redirect('/');
  } catch (caught) {
    sendPage(req, res, 502, {
      heading: 'Sign-in could not be completed',
      message: 'The provider accepted the request, but the account could not be looked up or saved.',
      details: caught.message,
      tone: 'error',
    });
  }
});

// Log out is the only way to end a session that is not cookie expiry, so it has to
// actually delete the row as well as the cookie: leaving either behind means the
// next request walks straight back in.
app.post('/logout', (req, res) => {
  const finish = () => {
    res.clearCookie('connect.sid');
    res.redirect('/');
  };

  if (req.session) {
    req.session.destroy(() => finish());
    return;
  }

  finish();
});

app.post('/upload', upload.single('savedvars'), async (req, res) => {
  const user = req.session.user || null;

  if (enabledProviders.length > 0 && !user) {
    res.status(401).json({ error: 'Sign in with Discord or Google before uploading.' });
    return;
  }
  if (!req.file) {
    res.status(400).json({ error: 'Choose your ItemShareSavedVars.lua file first.' });
    return;
  }

  let parsed;
  try {
    parsed = extractShareData(req.file.buffer.toString('utf8'));
  } catch (error) {
    res.status(400).json({ error: `Could not read that file: ${error.message}` });
    return;
  }

  if (parsed.accounts.length === 0) {
    res.status(400).json({ error: 'No ESO account was found in that file. Upload ItemShareSavedVars.lua, not another file.' });
    return;
  }

  // Icons are resolved before the replace and outside its transaction: the
  // database may already know the path from an earlier upload, and anything it
  // does not is looked up once by item id on UESP. Holding the replace open
  // across those requests would be wrong, and an icon is never worth failing an
  // upload over.
  const iconEntries = [...parsed.items, ...parsed.alwaysShare];

  try {
    await withDatabase((connection) => fillKnownIcons(connection, iconEntries));
  } catch (error) {
    console.warn(`Could not reuse stored icon paths: ${error.message}`);
  }

  let iconLookup = { resolved: 0, attempted: 0, missing: 0 };

  try {
    iconLookup = await resolveMissingIconPaths(iconEntries);

    if (iconLookup.attempted) {
      console.log(
        `Icons: resolved ${iconLookup.resolved} of ${iconLookup.missing} missing paths on UESP`
      );
    }
  } catch (error) {
    console.warn(`Could not look up icon paths on UESP: ${error.message}`);
  }

  try {
    const outcome = await withDatabase(async (connection) => {
      if (user) {
        // Every ESO account in the file belongs to the signed-in person. An
        // account claimed by somebody else stops the whole upload.
        for (const accountName of parsed.accounts) {
          const link = await linkAccountToUser(connection, { accountName, userId: user.id });

          if (link === 'conflict') {
            const conflict = new Error(
              `The ESO account ${accountName} is already linked to a different person.`
            );
            conflict.statusCode = 409;
            throw conflict;
          }
        }
      }

      return replaceAccounts(
        connection,
        parsed.accounts.map((accountName) => ({
          accountName,
          items: parsed.items.filter((item) => item.accountName === accountName),
          alwaysShare: parsed.alwaysShare.filter((entry) => entry.accountName === accountName),
          addonVersion: parsed.addonVersions.get(accountName) || '',
          guilds: parsed.accountGuilds.get(accountName) || new Map(),
        })),
        { sourceFile: req.file.originalname, userId: user ? user.id : null }
      );
    });

    // One fetch per new icon, ever: afterwards /icons serves the stored bytes.
    let icons = { fetched: 0, skipped: 0, failed: 0 };

    try {
      icons = await withDatabase((connection) =>
        cacheItemIcons(connection, parsed.items.map((item) => item.icon))
      );

      if (icons.fetched || icons.failed) {
        console.log(`Icons: fetched ${icons.fetched}, already stored ${icons.skipped}, failed ${icons.failed}`);
      }
    } catch (error) {
      console.warn(`Could not cache item icons: ${error.message}`);
    }

    req.session.accountNames = parsed.accounts.slice();
    res.json({
      accounts: outcome,
      accountCount: outcome.length,
      itemCount: outcome.reduce((sum, entry) => sum + entry.itemCount, 0),
      alwaysShareCount: outcome.reduce((sum, entry) => sum + entry.alwaysShareCount, 0),
      totalQuantity: outcome.reduce((sum, entry) => sum + entry.totalQuantity, 0),
      guildCount: parsed.guilds.size,
      iconLookup,
      icons,
      link: user ? 'linked' : 'skipped',
    });
  } catch (error) {
    res.status(error.statusCode || 500).json({ error: error.message });
  }
});

// Icons that an upload could not resolve are retried here, because nothing else
// ever asks again: a blank path stays blank for ever otherwise. A bounded batch
// runs shortly after boot and then on an interval; 0 turns the pass off entirely
// for anyone who would rather not have the server calling out on its own.
const ICON_BACKFILL_INTERVAL_MS = Number(process.env.ICON_BACKFILL_INTERVAL_MS ?? 1000 * 60 * 30);
const ICON_BACKFILL_BATCH = Number(process.env.ICON_BACKFILL_BATCH || 20);

let iconBackfillRunning = false;

async function backfillIcons(reason) {
  if (iconBackfillRunning) {
    return;
  }

  iconBackfillRunning = true;

  try {
    const result = await withDatabase((connection) =>
      backfillItemIcons(connection, { limit: ICON_BACKFILL_BATCH })
    );

    if (result.itemIds) {
      console.log(
        `Icons: ${reason} backfill looked up ${result.itemIds}, resolved ${result.resolved}, ` +
          `filled ${result.rows} rows, fetched ${result.icons.fetched} images`
      );
    }
  } catch (error) {
    console.warn(`Could not backfill item icons: ${error.message}`);
  } finally {
    iconBackfillRunning = false;
  }
}

// Everything that is not a page we serve, including a mistyped address, still gets
// the site's own layout. Icons are images, so they get an empty 404 instead.
app.use((req, res) => {
  if (req.path.startsWith('/icons/')) {
    res.status(404).end();
    return;
  }

  if (req.path === '/favicon.ico') {
    res.status(204).end();
    return;
  }

  sendPage(req, res, 404, {
    heading: 'Page not found',
    message: 'That address is not part of ItemShare. The shared item list is on the home page.',
    details: `No route for ${req.method} ${req.path}`,
  });
});

// Uploads are answered in JSON because the page reads them with fetch; anything
// else that throws gets a styled page. Multer's own errors land here too, which is
// how an oversized file stops being an Express stack trace.
app.use((error, req, res, next) => {
  if (res.headersSent) {
    next(error);
    return;
  }

  const status = error.statusCode || (error.code === 'LIMIT_FILE_SIZE' ? 413 : 500);
  const message = error.message || 'Something went wrong.';
  const wantsJson =
    req.path === '/upload' ||
    req.xhr ||
    String(req.get('accept') || '').includes('application/json');

  console.warn(`${req.method} ${req.path} failed (${status}): ${message}`);

  if (wantsJson) {
    res.status(status).json({ error: message });
    return;
  }

  sendPage(req, res, status, {
    heading: status === 413 ? 'That file is too large' : 'Something went wrong',
    message,
    details:
      status === 413 ? 'ItemShareSavedVars.lua should be well under 12 MB, so check you picked the right file.' : '',
    tone: 'error',
  });
});

await prepareDatabase().catch((error) => {
  let config = {};
  try {
    config = readConfig();
  } catch {
    // A malformed ITEMSHARE_DATABASE is reported by prepareDatabase itself.
  }
  console.warn(describeConnectionFailure(error, config));
});

// Expired rows are invisible to every query, so this is housekeeping rather than
// correctness: once at startup, then twice a day.
sessionStore
  .prune()
  .then((removed) => {
    if (removed) {
      console.log(`Sessions: removed ${removed} expired row${removed === 1 ? '' : 's'}`);
    }
  })
  .catch((error) => console.warn(`Could not prune expired sessions: ${error.message}`));

setInterval(() => {
  sessionStore.prune().catch(() => {});
}, 1000 * 60 * 60 * 12);

app.listen(PORT, () => {
  console.log(`ItemShare upload site listening on http://localhost:${PORT}`);
  console.log(
    enabledProviders.length
      ? `Sign-in providers: ${enabledProviders.map((provider) => provider.label).join(', ')}`
      : 'Sign-in providers: none (set DISCORD_* or GOOGLE_* env vars)'
  );
  console.log(`Env file: ${hasEnvFile ? envPath : `${envPath} not found - copy .env.example to .env`}`);
  console.log(`Sign-in stays active for ${SESSION_TTL_DAYS} days (SESSION_TTL_DAYS), stored in MariaDB`);
  console.log(`Database target: ${readConfig().database}`);

  if (ICON_BACKFILL_INTERVAL_MS > 0) {
    console.log(
      `Icon backfill: up to ${ICON_BACKFILL_BATCH} missing items every ${Math.round(
        ICON_BACKFILL_INTERVAL_MS / 60000
      )} minutes (ICON_BACKFILL_INTERVAL_MS=0 disables)`
    );

    // Not unref'd: the server is the process, so the pass should live as long as
    // it does. The first run waits a minute so a restart loop cannot hammer UESP.
    setTimeout(() => backfillIcons('startup'), 60 * 1000);
    setInterval(() => backfillIcons('scheduled'), ICON_BACKFILL_INTERVAL_MS);
  } else {
    console.log('Icon backfill: disabled (ICON_BACKFILL_INTERVAL_MS=0)');
  }
});
