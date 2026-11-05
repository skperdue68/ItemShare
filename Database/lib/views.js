// Server rendered pages. Styling lives in public/styles.css and follows the
// GuildSync palette (navy gradient, translucent panels, blue accents).

import { createHash } from 'node:crypto';

export function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ESO quality colours: normal, fine, superior, epic, legendary, mythic.
const QUALITY_COLORS = ['#e5e7eb', '#7cb342', '#3b82f6', '#a855f7', '#fbbf24', '#f97316'];

function qualityColor(quality) {
  const index = Number(quality);
  return QUALITY_COLORS[Number.isFinite(index) ? Math.max(0, Math.min(5, index)) : 0];
}

function seenDate(value) {
  const seconds = Number(value);
  if (!Number.isFinite(seconds) || seconds <= 0) {
    return '';
  }
  const date = new Date(seconds * 1000);
  return Number.isNaN(date.getTime()) ? '' : date.toISOString().slice(0, 10);
}

// Two initials, the way an avatar reads at a glance. Discord names are one word
// often enough that the fallback matters more than the fancy case.
function initialsFor(username) {
  const words = String(username || '')
    .trim()
    .split(/[\s._-]+/)
    .filter(Boolean);

  if (words.length === 0) {
    return '?';
  }

  if (words.length === 1) {
    return words[0].slice(0, 2).toUpperCase();
  }

  return `${words[0][0]}${words[1][0]}`.toUpperCase();
}

// Signed in: a circular avatar that opens a small menu, with Log out inside it.
// The logout is a real form post, so it still works with scripts blocked.
function authArea(user, providers) {
  if (user) {
    const name = escapeHtml(user.username || 'signed in');
    const provider = escapeHtml(user.providerLabel || user.provider || 'sign-in');
    const initials = escapeHtml(initialsFor(user.username));

    // The picture sits on top of the initials, so if it fails to load the initials are
    // already there. app.js removes a broken image; without scripts the alt text is
    // empty, so nothing worse than a blank circle appears.
    const avatar = user.avatarUrl
      ? `<span class="avatar avatar-with-photo" aria-hidden="true">
            <span class="avatar-initials">${initials}</span>
            <img class="avatar-photo" src="${escapeHtml(user.avatarUrl)}" alt="" loading="lazy" referrerpolicy="no-referrer">
          </span>`
      : `<span class="avatar" aria-hidden="true">${initials}</span>`;

    return `
      <div class="account-menu">
        <button class="avatar-button" type="button" id="accountMenuButton"
                aria-haspopup="true" aria-expanded="false" aria-controls="accountMenuPanel"
                title="${name}">
          ${avatar}
          <span class="avatar-caret" aria-hidden="true"></span>
          <span class="visually-hidden">Account menu for ${name}</span>
        </button>
        <div class="account-menu-panel" id="accountMenuPanel" role="menu"
             aria-labelledby="accountMenuButton" hidden>
          <div class="account-menu-head">
            <span class="account-menu-name">${name}</span>
            <span class="account-menu-note">signed in via ${provider}</span>
          </div>
          <form method="post" action="/logout">
            <button class="account-menu-item" type="submit" role="menuitem">Log out</button>
          </form>
        </div>
      </div>`;
  }

  if (!providers || providers.length === 0) {
    return '';
  }

  const buttons = providers
    .map(
      (provider) => `
      <a class="discord-button auth-button-${escapeHtml(provider.key)}" href="/auth/${escapeHtml(provider.key)}">
        <span class="discord-mark">${provider.key === 'google' ? 'G' : '&#9670;'}</span> Sign in with ${escapeHtml(provider.label)}
      </a>`
    )
    .join('');

  return `<div class="auth-buttons">${buttons}</div>`;
}

function statusCard({ dbError, providers, user }) {
  if (dbError) {
    return `
      <section class="card card-error">
        <h2>Database unavailable</h2>
        <p class="muted">${escapeHtml(dbError)}</p>
        <p class="muted">Check MARIADB_HOST, MARIADB_USER and MARIADB_PASSWORD in <code>Database/.env</code>, then reload.</p>
      </section>`;
  }

  if (!providers || providers.length === 0) {
    return `
      <section class="card card-notice">
        <h2>Sign-in is not configured</h2>
        <p class="muted">
          Uploads work, but they will not be linked to a person. Set
          <code>DISCORD_CLIENT_ID</code> / <code>DISCORD_CLIENT_SECRET</code> or
          <code>GOOGLE_CLIENT_ID</code> / <code>GOOGLE_CLIENT_SECRET</code> in
          <code>Database/.env</code> to enable sign-in.
        </p>
      </section>`;
  }

  if (!user) {
    return `
      <section class="card card-notice">
        <h2>Sign in to see and share items</h2>
        <p class="muted">
          Sign in with Discord or Google; it is only used to connect your uploads to
          your ESO accounts. Once an account is linked, you will see every shared item
          from the guilds you belong to.
        </p>
      </section>`;
  }

  return '';
}

function guildChips(guilds) {
  if (!guilds || guilds.length === 0) {
    return `<p class="muted">No guilds recorded yet. Upload a file from ItemShare 1.3.0 or newer to record your guild membership.</p>`;
  }

  const chips = guilds
    .map((guild) => `<span class="chip">${escapeHtml(guild.guild_name || 'Unknown guild')}</span>`)
    .join('');

  return `<div class="chips">${chips}</div>`;
}

function emailLine(user) {
  if (!user) {
    return '';
  }

  const email = user.email ? escapeHtml(user.email) : 'not shared with this site';
  const verified = user.email ? (user.emailVerified ? 'verified' : 'unverified') : '';
  return `<div><span class="meta-label">Sign-in email</span><span class="meta-value">${email}${verified ? ` <span class="meta-note">(${verified})</span>` : ''}</span></div>`;
}

function accountsCard(accounts, user) {
  if (!accounts || accounts.length === 0) {
    return `
      <section class="card">
        <h2>Your ESO accounts</h2>
        <p class="muted">
          No ESO account is linked to your Discord login yet. Upload your
          <code>ItemShareSavedVars.lua</code> file below: every account in that file is
          linked to this login, and each account keeps its own items and guilds.
        </p>
        <p class="muted">
          Nothing is listed until you do. Items are shared through guilds, so until your
          ESO account and its guilds are known there is nobody to share with you.
        </p>
        <div class="meta-row">${emailLine(user)}</div>
      </section>`;
  }

  const totalItems = accounts.reduce((sum, account) => sum + Number(account.itemCount || 0), 0);
  const totalQuantity = accounts.reduce((sum, account) => sum + Number(account.totalQuantity || 0), 0);

  const blocks = accounts
    .map((account) => {
      const uploadVersion = account.lastUpload && account.lastUpload.addon_version
        ? ` &middot; addon ${escapeHtml(account.lastUpload.addon_version)}`
        : '';
      const lastUpload = account.lastUpload
        ? `${escapeHtml(account.lastUpload.uploaded_at)} &middot; ${escapeHtml(account.lastUpload.source_file || 'upload')}${uploadVersion}`
        : 'never';

      const linked = account.link && account.link.user_id
        ? escapeHtml(account.link.display_name || `user #${account.link.user_id}`)
        : 'not linked yet';

      return `
      <div class="account-block">
        <div class="meta-row">
          <div><span class="meta-label">ESO account</span><span class="meta-value">${escapeHtml(account.accountName)}</span></div>
          <div><span class="meta-label">Linked by</span><span class="meta-value">${linked}</span></div>
          <div><span class="meta-label">Last upload</span><span class="meta-value">${lastUpload}</span></div>
          ${emailLine(user)}
        </div>
        <div class="stat-grid">
          <div class="stat"><span class="stat-value">${account.itemCount}</span><span class="stat-label">Items</span></div>
          <div class="stat"><span class="stat-value">${account.totalQuantity}</span><span class="stat-label">Quantity</span></div>
          <div class="stat"><span class="stat-value">${account.guildCount}</span><span class="stat-label">Guilds</span></div>
        </div>
        <div class="subsection">
          <h3>Guilds</h3>
          ${guildChips(account.guilds)}
        </div>
      </div>`;
    })
    .join('');

  const heading = accounts.length === 1 ? 'Your ESO account' : `Your ESO accounts (${accounts.length})`;
  const note = accounts.length > 1
    ? `<p class="muted">
        All ${accounts.length} accounts are tied to this Discord login. Items stay listed
        under the account that holds them, so you can tell which account has what.
      </p>
      <div class="stat-grid">
        <div class="stat"><span class="stat-value">${totalItems}</span><span class="stat-label">Items, all accounts</span></div>
        <div class="stat"><span class="stat-value">${totalQuantity}</span><span class="stat-label">Quantity, all accounts</span></div>
      </div>`
    : '';

  return `
    <section class="card">
      <h2>${heading}</h2>
      ${note}
      ${blocks}
    </section>`;
}

function filterBar(facets, filters) {
  const types = facets?.types || [];
  const owners = facets?.owners || [];
  const active = filters || { type: '', owner: '', search: '' };

  if (types.length === 0 && owners.length === 0) {
    return '';
  }

  const typeOptions = types
    .map(
      (entry) => `
            <option value="${escapeHtml(entry.name)}"${active.type === entry.name ? ' selected' : ''}>${escapeHtml(entry.name)} (${entry.total})</option>`
    )
    .join('');

  const ownerOptions = owners
    .map(
      (owner) => `
            <option value="${escapeHtml(owner)}"${active.owner === owner ? ' selected' : ''}>${escapeHtml(owner)}</option>`
    )
    .join('');

  // Filters apply as soon as they change (app.js), so there is no button to forget
  // to press. The submit button exists for the one case that needs it: scripts off.
  return `
    <form class="filters" method="get" action="/" data-filters>
      <label class="filter">
        <span class="filter-label">Type</span>
        <select name="type" data-filter>
          <option value="">All types</option>${typeOptions}
        </select>
      </label>
      <label class="filter">
        <span class="filter-label">Offered by</span>
        <select name="owner" data-filter>
          <option value="">Anyone</option>${ownerOptions}
        </select>
      </label>
      <label class="filter filter-search">
        <span class="filter-label">Item name</span>
        <input type="search" name="q" data-filter value="${escapeHtml(active.search)}" placeholder="sword, ring, motif...">
      </label>
      <noscript><button class="primary-button filter-button" type="submit">Filter</button></noscript>
      <button class="filter-clear" type="button" data-clear-filters>Clear filters</button>
    </form>`;
}

function itemIcon(item) {
  return Number(item.has_icon) === 1
    ? `<img class="item-icon" src="/icons/${escapeHtml(item.icon_key)}.png" alt="" width="24" height="24" loading="lazy">`
    : `<span class="item-icon item-icon-empty" aria-hidden="true"></span>`;
}

function itemName(item) {
  const name = escapeHtml(item.item_name);

  return Number(item.item_id) > 0
    ? `<a class="item-name-link" href="https://esoitem.uesp.net/itemLink.php?itemid=${Number(item.item_id)}" target="_blank" rel="noopener noreferrer">${name}</a>`
    : name;
}

// A detail earns its place only when it says something. The addon stores a "0" for
// every getter that answered with a number, and "0" is not a label, so those are
// dropped rather than printed next to the set name.
function metaLine(parts) {
  return parts
    .filter((part) => part !== null && part !== undefined && !['', '0'].includes(String(part).trim()))
    .map((part) => escapeHtml(part))
    .join(' &middot; ');
}

function ownerCell(item) {
  return `<span class="owner">${escapeHtml(item.account_name)}</span>
                <span class="owner-character">${escapeHtml(item.character_name || '')}</span>`;
}

// Guild names are deliberately absent: every row on this page is already there
// because the viewer and the owner share a guild.
function inventoryRows(items, { mine }) {
  return items
    .map((item) => {
      const unattributed = !item.guild_names;
      const details = metaLine([
        item.set_name,
        item.style_name,
        item.enchant_name ? `Enchant: ${item.enchant_name}` : '',
        Number(item.required_cp) > 0 ? `CP ${Number(item.required_cp)}` : '',
        Number(item.required_level) > 0 ? `Level ${Number(item.required_level)}` : '',
        item.location_label ? `owner had it in ${item.location_label}` : '',
        seenDate(item.last_dumped_at) ? `last seen ${seenDate(item.last_dumped_at)}` : '',
      ]);

      return `
        <tr data-item-key="inv:${escapeHtml(item.item_key_hash || item.id)}">
          <td>
            <div class="item-cell">
              ${itemIcon(item)}
              <div class="item-cell-text">
                <span class="item-name" style="color:${qualityColor(item.quality)}">${itemName(item)}</span>
                ${mine && unattributed ? '<span class="own-tag">Unattributed</span>' : ''}
                ${details ? `<span class="item-meta">${details}</span>` : ''}
              </div>
            </div>
          </td>
          <td class="num">${Number(item.item_count) || 0}</td>
          <td>${escapeHtml(item.item_type_name || '')}</td>
          <td>${escapeHtml(item.trait || '')}</td>
          <td>${ownerCell(item)}</td>
        </tr>`;
    })
    .join('');
}

function typeGroups(items) {
  const groups = new Map();

  for (const item of items) {
    const name = item.item_type_name || 'Other';
    if (!groups.has(name)) {
      groups.set(name, []);
    }
    groups.get(name).push(item);
  }

  return [...groups.entries()];
}

function inventoryTable(items, options) {
  const body = typeGroups(items)
    .map(
      ([typeName, rows]) => `
        <tr class="type-row">
          <td colspan="5">${escapeHtml(typeName)} <span class="type-count">${rows.length}</span></td>
        </tr>
        ${inventoryRows(rows, options)}`
    )
    .join('');

  return `
    <div class="table-shell">
      <table>
        <thead>
          <tr><th>Item</th><th class="num">Qty</th><th>Type</th><th>Trait</th><th>Offered by</th></tr>
        </thead>
        <tbody>${body}</tbody>
      </table>
    </div>`;
}

function offerRows(items) {
  return items
    .map((item) => {
      const details = metaLine([
        item.item_type_name,
        item.set_name,
        item.trait,
        Number(item.in_inventory) === 1 ? 'has one in a bag right now' : '',
        Number(item.first_added_at) > 0 ? `offered since ${seenDate(item.first_added_at)}` : '',
      ]);

      return `
        <tr data-item-key="offer:${escapeHtml(item.item_key_hash || item.id)}">
          <td>
            <div class="item-cell">
              ${itemIcon(item)}
              <div class="item-cell-text">
                <span class="item-name" style="color:${qualityColor(item.quality)}">${itemName(item)}</span>
                ${details ? `<span class="item-meta">${details}</span>` : ''}
              </div>
            </div>
          </td>
          <td>${ownerCell(item)}</td>
        </tr>`;
    })
    .join('');
}

function offerTable(items) {
  return `
    <div class="table-shell">
      <table>
        <thead><tr><th>Item</th><th>Offered by</th></tr></thead>
        <tbody>${offerRows(items)}</tbody>
      </table>
    </div>`;
}

/**
 * Stable identity for one rendered list: the accounts, item types and owners in it.
 * The page embeds it and a live refresh compares it, so a change that adds or
 * removes a filter option rebuilds the page instead of leaving the controls stale.
 */
export function listRegionSignature({ accounts = [], facets = {} } = {}) {
  const payload = JSON.stringify({
    accounts: (accounts || []).map((account) => account.accountName),
    types: ((facets && facets.types) || []).map((entry) => entry.name),
    owners: (facets && facets.owners) || [],
  });

  return createHash('sha1').update(payload).digest('hex').slice(0, 12);
}

/**
 * The two lists, inside one element the browser can replace without touching the
 * filters, the upload card or the reader's place in either list.
 *
 * Request is what other people share and you can ask for; Offering is your own. The
 * split is why no row needs a "You" tag: which tab it is in already says whose it is.
 */
export function renderListRegion({ accounts, sharedItems, alwaysShare, facets, filters }) {
  if (!accounts || accounts.length === 0) {
    return '';
  }

  const myAccounts = new Set(accounts.map((account) => account.accountName));
  const requestable = (sharedItems || []).filter((item) => !myAccounts.has(item.account_name));
  const mine = (sharedItems || []).filter((item) => myAccounts.has(item.account_name));
  const othersOffers = (alwaysShare || []).filter((item) => !myAccounts.has(item.account_name));
  const myOffers = (alwaysShare || []).filter((item) => myAccounts.has(item.account_name));

  const activeFilters = Boolean(filters?.type || filters?.owner || filters?.search);
  const requestCount = requestable.length + othersOffers.length;
  const offerCount = mine.length + myOffers.length;

  const empty = (fallback) =>
    `<p class="muted">${
      activeFilters ? 'No items match those filters. Clear them to see everything.' : fallback
    }</p>`;

  const requestBody =
    requestCount === 0
      ? empty('Nothing is being shared by other members of your guilds right now. Whatever they share appears here on its own.')
      : `${requestable.length > 0 ? inventoryTable(requestable, { mine: false }) : ''}${
          othersOffers.length > 0
            ? `<h3 class="subhead">Standing offers</h3>
               <p class="muted">They will hand these out whenever they have one, even when it is not in a bag today.</p>
               ${offerTable(othersOffers)}`
            : ''
        }`;

  const offerBody =
    offerCount === 0
      ? empty('You are not offering anything yet. Upload your file and the items you share in game appear here.')
      : `${mine.length > 0 ? inventoryTable(mine, { mine: true }) : ''}${
          myOffers.length > 0
            ? `<h3 class="subhead">Your standing offers</h3>
               ${offerTable(myOffers)}`
            : ''
        }`;

  return `
    <div class="item-lists" id="itemLists" data-signature="${escapeHtml(listRegionSignature({ accounts, facets }))}">
      <section class="card">
        <div class="list-head">
          <div class="tabs" role="tablist" aria-label="Item lists">
            <button class="tab" type="button" role="tab" id="tabRequest" aria-controls="panelRequest"
                    aria-selected="true" data-tab="request">
              Request <span class="tab-count">${requestCount}</span>
            </button>
            <button class="tab" type="button" role="tab" id="tabOffer" aria-controls="panelOffer"
                    aria-selected="false" tabindex="-1" data-tab="offer">
              Offering <span class="tab-count">${offerCount}</span>
            </button>
          </div>
          <span class="live-pill" data-live="wait" hidden>Live</span>
        </div>

        <div class="list-panel" id="panelRequest" role="tabpanel" aria-labelledby="tabRequest" data-panel="request">
          ${
            requestCount > 0
              ? `<p class="muted">${requestable.length} item${requestable.length === 1 ? '' : 's'} shared by other members of your guilds. Quantities are what each owner last reported.</p>`
              : ''
          }
          ${requestBody}
        </div>

        <div class="list-panel" id="panelOffer" role="tabpanel" aria-labelledby="tabOffer" data-panel="offer" hidden>
          ${offerBody}
        </div>
      </section>
    </div>`;
}

// One shell for every page. Without this a failed sign-in, a wrong address or a
// cancelled prompt dropped the browser onto a bare line of text with no styling and
// no way back, which reads as a broken site rather than a message.
function pageShell({ user = null, providers = [], content, title = 'ItemShare' }) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml(title)}</title>
  <link rel="stylesheet" href="/styles.css">
</head>
<body>
  <header class="top-bar">
    <div class="brand">
      <div class="brand-mark">IS</div>
      <div class="brand-text">
        <span class="brand-title">ItemShare</span>
        <span class="brand-subtitle">Shared items across your guilds</span>
      </div>
    </div>
    ${authArea(user, providers)}
  </header>

  <main class="surface">
${content}
  </main>

  <footer class="footer">
    <span>ItemShare &middot; uploads replace your guilds and items in full</span>
  </footer>

  <div class="toast" id="toast" role="status" aria-live="polite" hidden></div>

  <script src="/socket.io/socket.io.js" defer></script>
  <script src="/app.js" defer></script>
</body>
</html>`;
}

/**
 * A full page for something that is not the item list: a cancelled sign-in, a
 * provider that is not configured, a wrong address, an upload that failed.
 */
export function renderMessagePage({
  title,
  heading,
  message,
  details = '',
  tone = 'notice',
  user = null,
  providers = [],
}) {
  const cardClass = tone === 'error' ? 'card card-error' : 'card card-notice';

  const content = `
    <section class="${cardClass}">
      <h2>${escapeHtml(heading)}</h2>
      <p class="muted">${escapeHtml(message)}</p>
      ${details ? `<p class="muted">${escapeHtml(details)}</p>` : ''}
      <div class="actions">
        <a class="primary-button" href="/">Back to ItemShare</a>
      </div>
    </section>`;

  return pageShell({ user, providers, content, title: title || `${heading} \u00b7 ItemShare` });
}

function uploadCard() {
  return `
    <section class="card">
      <h2>Share your file</h2>
      <p class="muted">
        Drop <code>ItemShareSavedVars.lua</code> here or click to choose it - it uploads as
        soon as it is picked, and replaces your guilds and items in full. The file lives in
        <code>Documents\\Elder Scrolls Online\\live\\SavedVariables\\</code>.
      </p>
      <label class="drop-zone" id="dropZone" for="savedvars">
        <span class="drop-title">Drop the file here, or click to choose it</span>
        <span class="drop-file" id="fileName">No file selected</span>
        <input id="savedvars" name="savedvars" type="file" accept=".lua,.txt,text/plain">
      </label>
    </section>`;
}

export function renderPage({ user, providers, accounts, sharedItems, alwaysShare, facets, filters, dbError }) {
  const canUpload = !dbError && (!providers || providers.length === 0 || user);
  // Nothing is listed before the first upload: an account with no guilds shares no
  // guild with anybody, so it has nothing it could legitimately be shown.
  const initialized = Boolean(accounts && accounts.length > 0);
  const bar = initialized ? filterBar(facets, filters) : '';

  const content = `
    ${statusCard({ dbError, providers, user })}

    ${canUpload ? uploadCard() : ''}

    ${accountsCard(accounts, user)}

    ${bar ? `<section class="card card-filters">${bar}</section>` : ''}
    ${renderListRegion({ accounts, sharedItems, alwaysShare, facets, filters })}`;

  return pageShell({ user, providers, content });
}
