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
import { cacheItemIcons, getCachedIcon } from './lib/icons.js';
import {
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
import { renderPage } from './lib/views.js';

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

const app = express();
app.use(express.urlencoded({ extended: false }));
app.use(
  session({
    secret: SESSION_SECRET,
    resave: false,
    saveUninitialized: false,
    cookie: { httpOnly: true, sameSite: 'lax', secure: false, maxAge: 1000 * 60 * 60 * 12 },
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
    res.status(503).send('That sign-in provider is not configured on this server.');
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
    res.status(503).send('That sign-in provider is not configured on this server.');
    return;
  }
  if (error || !code) {
    res.status(400).send(`${provider.label} sign-in was cancelled or failed: ${error || 'no code returned'}`);
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
    };
    req.session.accountNames = null;

    res.redirect('/');
  } catch (caught) {
    res.status(502).send(`Could not complete sign-in: ${caught.message}`);
  }
});

app.post('/logout', (req, res) => {
  req.session.destroy(() => res.redirect('/'));
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
      icons,
      link: user ? 'linked' : 'skipped',
    });
  } catch (error) {
    res.status(error.statusCode || 500).json({ error: error.message });
  }
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

app.listen(PORT, () => {
  console.log(`ItemShare upload site listening on http://localhost:${PORT}`);
  console.log(
    enabledProviders.length
      ? `Sign-in providers: ${enabledProviders.map((provider) => provider.label).join(', ')}`
      : 'Sign-in providers: none (set DISCORD_* or GOOGLE_* env vars)'
  );
  console.log(`Env file: ${hasEnvFile ? envPath : `${envPath} not found - copy .env.example to .env`}`);
  console.log(`Database target: ${readConfig().database}`);
});
