#!/usr/bin/env node
// ItemShare upload site: signs in with Discord (optional), accepts
// ItemShareSavedVars.lua and replaces every shared item for the ESO account in
// the file. Styling follows the GuildSync palette.
//
//   npm start            (reads Database/.env)

import { fileURLToPath } from 'node:url';
import path from 'node:path';
import express from 'express';
import session from 'express-session';
import multer from 'multer';
import 'dotenv/config';

import { extractShareData } from './lib/savedVars.js';
import { openDatabase, readConfig } from './lib/database.js';
import {
  getAccountSummary,
  getSharedItemsForViewer,
  getViewerAccounts,
  getViewerGuilds,
  linkDiscordAccount,
  replaceAccounts,
  upsertDiscordUser,
} from './lib/store.js';
import { renderPage } from './lib/views.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const PORT = Number(process.env.PORT || 3000);
const DISCORD_CLIENT_ID = process.env.DISCORD_CLIENT_ID || '';
const DISCORD_CLIENT_SECRET = process.env.DISCORD_CLIENT_SECRET || '';
const DISCORD_REDIRECT_URI = process.env.DISCORD_REDIRECT_URI || `http://localhost:${PORT}/auth/discord/callback`;
const discordConfigured = Boolean(DISCORD_CLIENT_ID && DISCORD_CLIENT_SECRET);

if (!process.env.SESSION_SECRET) {
  console.warn('SESSION_SECRET is not set; using an ephemeral secret. Sessions end when the server restarts.');
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
  const base = { user, discordConfigured, accounts: [], sharedItems: [], dbError: null };

  try {
    await withDatabase(async (connection) => {
      // One Discord login can own several ESO accounts; every account linked to
      // it contributes guilds, and the items stay labelled with their account.
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

      base.sharedItems = await getSharedItemsForViewer(connection, accountNames);
    });
  } catch (error) {
    base.dbError = error.message;
  }

  res.type('html').send(renderPage(base));
});

app.get('/auth/discord', (req, res) => {
  if (!discordConfigured) {
    res.status(503).send('Discord sign-in is not configured on this server.');
    return;
  }

  const params = new URLSearchParams({
    client_id: DISCORD_CLIENT_ID,
    redirect_uri: DISCORD_REDIRECT_URI,
    response_type: 'code',
    scope: 'identify email',
    state: req.sessionID,
  });

  res.redirect(`https://discord.com/oauth2/authorize?${params.toString()}`);
});

app.get('/auth/discord/callback', async (req, res) => {
  const { code, error } = req.query;

  if (!discordConfigured) {
    res.status(503).send('Discord sign-in is not configured on this server.');
    return;
  }
  if (error || !code) {
    res.status(400).send(`Discord sign-in was cancelled or failed: ${error || 'no code returned'}`);
    return;
  }

  try {
    const tokenResponse = await fetch('https://discord.com/api/oauth2/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: DISCORD_CLIENT_ID,
        client_secret: DISCORD_CLIENT_SECRET,
        grant_type: 'authorization_code',
        code: String(code),
        redirect_uri: DISCORD_REDIRECT_URI,
      }),
    });

    if (!tokenResponse.ok) {
      throw new Error(`Discord token exchange failed (${tokenResponse.status})`);
    }

    const token = await tokenResponse.json();
    const userResponse = await fetch('https://discord.com/api/users/@me', {
      headers: { Authorization: `Bearer ${token.access_token}` },
    });

    if (!userResponse.ok) {
      throw new Error(`Discord user lookup failed (${userResponse.status})`);
    }

    const profile = await userResponse.json();
    req.session.user = {
      id: String(profile.id),
      username: profile.global_name || profile.username || String(profile.id),
      email: typeof profile.email === 'string' ? profile.email : '',
      emailVerified: profile.verified === true,
    };
    req.session.accountNames = null;

    // Best effort: a database problem must not block sign-in. The profile is
    // stored again on upload, so nothing is lost while the database is down.
    try {
      await withDatabase((connection) =>
        upsertDiscordUser(connection, {
          id: req.session.user.id,
          username: req.session.user.username,
          email: req.session.user.email,
          emailVerified: req.session.user.emailVerified,
        })
      );
    } catch (error) {
      console.warn(`Could not store the Discord profile: ${error.message}`);
    }

    res.redirect('/');
  } catch (caught) {
    res.status(502).send(`Could not complete Discord sign-in: ${caught.message}`);
  }
});

app.post('/logout', (req, res) => {
  req.session.destroy(() => res.redirect('/'));
});

app.post('/upload', upload.single('savedvars'), async (req, res) => {
  const user = req.session.user || null;

  if (discordConfigured && !user) {
    res.status(401).json({ error: 'Sign in with Discord before uploading.' });
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
        // Every ESO account in the file belongs to this Discord login. An
        // account claimed by somebody else stops the whole upload.
        for (const accountName of parsed.accounts) {
          const link = await linkDiscordAccount(connection, {
            accountName,
            discordUserId: user.id,
            discordUsername: user.username,
          });

          if (link === 'conflict') {
            const conflict = new Error(
              `The ESO account ${accountName} is already linked to a different Discord user.`
            );
            conflict.statusCode = 409;
            throw conflict;
          }
        }

        await upsertDiscordUser(connection, {
          id: user.id,
          username: user.username,
          email: user.email,
          emailVerified: user.emailVerified,
        });
      }

      return replaceAccounts(
        connection,
        parsed.accounts.map((accountName) => ({
          accountName,
          items: parsed.items.filter((item) => item.accountName === accountName),
          guilds: parsed.accountGuilds.get(accountName) || new Map(),
        })),
        { sourceFile: req.file.originalname, discordUserId: user ? user.id : '' }
      );
    });

    req.session.accountNames = parsed.accounts.slice();
    res.json({
      accounts: outcome,
      accountCount: outcome.length,
      itemCount: outcome.reduce((sum, entry) => sum + entry.itemCount, 0),
      totalQuantity: outcome.reduce((sum, entry) => sum + entry.totalQuantity, 0),
      guildCount: parsed.guilds.size,
      link: user ? 'linked' : 'skipped',
    });
  } catch (error) {
    res.status(error.statusCode || 500).json({ error: error.message });
  }
});

await prepareDatabase().catch((error) => {
  console.warn(`Database not ready at startup: ${error.message}`);
  console.warn('The site stays up and shows the error; fix the MariaDB settings and restart to apply migrations.');
});

app.listen(PORT, () => {
  console.log(`ItemShare upload site listening on http://localhost:${PORT}`);
  console.log(`Discord sign-in: ${discordConfigured ? 'enabled' : 'disabled (set DISCORD_* env vars)'}`);
  console.log(`Database target: ${readConfig().database}`);
});
