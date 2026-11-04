// Server rendered pages. Styling lives in public/styles.css and follows the
// GuildSync palette (navy gradient, translucent panels, blue accents).

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

function authArea(user, providers) {
  if (user) {
    return `
      <form method="post" action="/logout" class="user-area">
        <span class="user-name">
          ${escapeHtml(user.username)}
          <span class="provider-note">via ${escapeHtml(user.providerLabel || user.provider || 'sign-in')}</span>
        </span>
        <button class="discord-button discord-button-secondary" type="submit">Sign out</button>
      </form>`;
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
    .map(
      (guild) => `
        <span class="chip">
          ${escapeHtml(guild.guild_name || 'Unknown guild')}
          <span class="chip-id">${escapeHtml(guild.guild_id)}</span>
        </span>`
    )
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
          No ESO account is linked to your Discord user yet. Upload your
          <code>ItemShareSavedVars.lua</code> file below: every account in that file is
          linked to this Discord login, and each account keeps its own items and guilds.
        </p>
        <div class="meta-row">${emailLine(user)}</div>
      </section>`;
  }

  const totalItems = accounts.reduce((sum, account) => sum + Number(account.itemCount || 0), 0);
  const totalQuantity = accounts.reduce((sum, account) => sum + Number(account.totalQuantity || 0), 0);

  const blocks = accounts
    .map((account) => {
      const lastUpload = account.lastUpload
        ? `${escapeHtml(account.lastUpload.uploaded_at)} &middot; ${escapeHtml(account.lastUpload.source_file || 'upload')}`
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

  if (types.length === 0 && owners.length === 0) {
    return '';
  }

  const typeOptions = types
    .map(
      (entry) => `
            <option value="${escapeHtml(entry.name)}"${filters.type === entry.name ? ' selected' : ''}>${escapeHtml(entry.name)} (${entry.total})</option>`
    )
    .join('');

  const ownerOptions = owners
    .map(
      (owner) => `
            <option value="${escapeHtml(owner)}"${filters.owner === owner ? ' selected' : ''}>${escapeHtml(owner)}</option>`
    )
    .join('');

  return `
    <form class="filters" method="get" action="/">
      <label class="filter">
        <span class="filter-label">Type</span>
        <select name="type">
          <option value="">All types</option>${typeOptions}
        </select>
      </label>
      <label class="filter">
        <span class="filter-label">Owner</span>
        <select name="owner">
          <option value="">Anyone</option>${ownerOptions}
        </select>
      </label>
      <label class="filter filter-search">
        <span class="filter-label">Item name</span>
        <input type="search" name="q" value="${escapeHtml(filters.search)}" placeholder="sword, ring, motif...">
      </label>
      <button class="primary-button filter-button" type="submit">Filter</button>
      <a class="filter-clear" href="/">Clear</a>
    </form>`;
}

function sharedItemsCard({ accounts, sharedItems, facets, filters }) {
  if (!accounts || accounts.length === 0) {
    return '';
  }

  const heading = `
    <section class="card">
      <h2>Shared items in your guilds</h2>`;

  const activeFilters = Boolean(filters?.type || filters?.owner || filters?.search);
  const bar = filterBar(facets, filters || { type: '', owner: '', search: '' });

  if (!sharedItems || sharedItems.length === 0) {
    return `${heading}
      ${bar}
      <p class="muted">
        ${
          activeFilters
            ? 'No items match those filters. Clear them to see everything your guilds share.'
            : 'Nothing is shared with your guilds yet. If you expected items here, check that both you and the owner uploaded a file from ItemShare 1.3.0+ while in the same guild.'
        }
      </p>
    </section>`;
  }

  const myAccounts = new Set(accounts.map((account) => account.accountName));

  const rowsFor = (items) =>
    items
      .map((item) => {
        const isMine = myAccounts.has(item.account_name);
        const icon =
          Number(item.has_icon) === 1
            ? `<img class="item-icon" src="/icons/${escapeHtml(item.icon_key)}.png" alt="" width="24" height="24" loading="lazy">`
            : `<span class="item-icon item-icon-empty" aria-hidden="true"></span>`;

        const name = escapeHtml(item.item_name);
        const linkedName =
          Number(item.item_id) > 0
            ? `<a class="item-name-link" href="https://esoitem.uesp.net/itemLink.php?itemid=${Number(item.item_id)}" target="_blank" rel="noopener noreferrer">${name}</a>`
            : name;

        const details = [
          item.set_name,
          item.quality_name,
          item.weapon_type || item.armor_type || item.apparel_slot || item.equip_type,
          item.trait,
          item.enchant_name ? `Enchant: ${item.enchant_name}` : '',
          Number(item.required_cp) > 0 ? `CP ${Number(item.required_cp)}` : '',
          Number(item.required_level) > 0 ? `Level ${Number(item.required_level)}` : '',
          item.style_name,
          item.location_label ? `in ${item.location_label}` : '',
          seenDate(item.last_dumped_at) ? `seen ${seenDate(item.last_dumped_at)}` : '',
        ]
          .filter(Boolean)
          .map((part) => escapeHtml(part))
          .join(' &middot; ');

        return `
        <tr>
          <td>
            <div class="item-cell">
              ${icon}
              <div class="item-cell-text">
                <span class="item-name" style="color:${qualityColor(item.quality)}">${linkedName}</span>
                ${isMine ? '<span class="own-tag">You</span>' : ''}
                ${details ? `<span class="item-meta">${details}</span>` : ''}
              </div>
            </div>
          </td>
          <td class="num">${Number(item.item_count) || 0}</td>
          <td>${escapeHtml(item.item_type_name || '')}</td>
          <td>${escapeHtml(item.trait || '')}</td>
          <td>${escapeHtml(item.guild_names || '')}</td>
          <td>
            <span class="owner">${escapeHtml(item.account_name)}</span>
            <span class="owner-character">${escapeHtml(item.character_name || '')}</span>
          </td>
        </tr>`;
      })
      .join('');

  const groups = new Map();
  for (const item of sharedItems) {
    const typeName = item.item_type_name || 'Other';
    if (!groups.has(typeName)) {
      groups.set(typeName, []);
    }
    groups.get(typeName).push(item);
  }

  const body = [...groups.entries()]
    .map(
      ([typeName, items]) => `
        <tr class="type-row">
          <td colspan="6">${escapeHtml(typeName)} <span class="type-count">${items.length}</span></td>
        </tr>
        ${rowsFor(items)}`
    )
    .join('');

  return `${heading}
    <p class="muted">
      ${sharedItems.length} item${sharedItems.length === 1 ? '' : 's'} across ${groups.size}
      type${groups.size === 1 ? '' : 's'}, grouped by type and alphabetical inside each group.
      Quantities are what each owner last reported.
    </p>
    ${bar}
    <div class="table-shell">
      <table>
        <thead>
          <tr><th>Item</th><th class="num">Qty</th><th>Type</th><th>Trait</th><th>Guilds</th><th>Owner</th></tr>
        </thead>
        <tbody>${body}</tbody>
      </table>
    </div>
  </section>`;
}

export function renderPage({ user, providers, accounts, sharedItems, facets, filters, dbError }) {
  const canUpload = !dbError && (!providers || providers.length === 0 || user);

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>ItemShare</title>
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
    ${statusCard({ dbError, providers, user })}

    ${
      canUpload
        ? `<section class="card">
      <h2>Upload ItemShareSavedVars.lua</h2>
      <p class="muted">
        Select the file from
        <code>Documents\\Elder Scrolls Online\\live\\SavedVariables\\ItemShareSavedVars.lua</code>.
        Your guilds and items are replaced in full, and the list below refreshes.
      </p>
      <label class="drop-zone" id="dropZone" for="savedvars">
        <span class="drop-title">Drop the file here or click to browse</span>
        <span class="drop-file" id="fileName">No file selected</span>
        <input id="savedvars" name="savedvars" type="file" accept=".lua,.txt,text/plain">
      </label>
      <div class="actions">
        <button id="uploadButton" class="primary-button" type="button">Upload and replace</button>
      </div>
      <div id="status" class="status" role="status"></div>
    </section>`
        : ''
    }

    ${accountsCard(accounts, user)}
    ${sharedItemsCard({ accounts, sharedItems, facets, filters })}
  </main>

  <footer class="footer">
    <span>ItemShare &middot; uploads replace your guilds and items in full</span>
  </footer>

  <script src="/app.js"></script>
</body>
</html>`;
}
