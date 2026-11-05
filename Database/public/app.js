const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('savedvars');
const fileName = document.getElementById('fileName');
const uploadButton = document.getElementById('uploadButton');
const status = document.getElementById('status');

if (dropZone && fileInput) {
  fileInput.addEventListener('change', () => {
    const file = fileInput.files && fileInput.files[0];
    fileName.textContent = file ? `${file.name} (${Math.round(file.size / 1024)} KB)` : 'No file selected';
  });

  ['dragenter', 'dragover'].forEach((eventName) => {
    dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropZone.classList.add('dragging');
    });
  });

  ['dragleave', 'drop'].forEach((eventName) => {
    dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropZone.classList.remove('dragging');
    });
  });

  dropZone.addEventListener('drop', (event) => {
    const dropped = event.dataTransfer && event.dataTransfer.files;
    if (dropped && dropped.length > 0) {
      fileInput.files = dropped;
      fileInput.dispatchEvent(new Event('change'));
    }
  });
}

function setStatus(message, kind) {
  status.textContent = message || '';
  status.className = `status ${kind || ''}`;
}

if (uploadButton) {
  uploadButton.addEventListener('click', async () => {
    const file = fileInput.files && fileInput.files[0];
    if (!file) {
      setStatus('Choose your ItemShareSavedVars.lua file first.', 'err');
      return;
    }

    uploadButton.disabled = true;
    setStatus('Uploading and replacing every account in the file...', '');

    try {
      const body = new FormData();
      body.append('savedvars', file);

      const response = await fetch('/upload', { method: 'POST', body });
      const result = await response.json().catch(() => ({}));

      if (!response.ok) {
        throw new Error(result.error || `Upload failed with status ${response.status}`);
      }

      const summaryLines = (result.accounts || []).map(
        (account) =>
          `${account.accountName}: ${account.itemCount} items, quantity ${account.totalQuantity}, ${account.guildCount} guilds`
      );

      const fetchedIcons = result.icons ? Number(result.icons.fetched) || 0 : 0;

      setStatus(
        [
          `Saved ${summaryLines.length} ESO account${summaryLines.length === 1 ? '' : 's'}.`,
          ...summaryLines,
          fetchedIcons > 0 ? `Cached ${fetchedIcons} item icon${fetchedIcons === 1 ? '' : 's'} for local display.` : '',
          result.link === 'linked' ? 'Linked to your Discord user.' : '',
        ]
          .filter(Boolean)
          .join('\n'),
        'ok'
      );

      window.setTimeout(() => window.location.reload(), 1400);
    } catch (error) {
      setStatus(error.message, 'err');
      uploadButton.disabled = false;
    }
  });
}

// Account menu on the circular avatar. Log out sits inside the menu instead of
// taking up header space next to the name. The form posts without this script, so
// the menu only has to handle opening, closing and keyboard dismissal.
const accountMenuButton = document.getElementById('accountMenuButton');
const accountMenuPanel = document.getElementById('accountMenuPanel');

if (accountMenuButton && accountMenuPanel) {
  const setMenuOpen = (open) => {
    accountMenuPanel.hidden = !open;
    accountMenuButton.setAttribute('aria-expanded', String(open));

    if (open) {
      const firstItem = accountMenuPanel.querySelector('button, a');
      if (firstItem) {
        firstItem.focus();
      }
    }
  };

  accountMenuButton.addEventListener('click', (event) => {
    event.stopPropagation();
    setMenuOpen(accountMenuPanel.hidden);
  });

  document.addEventListener('click', (event) => {
    if (!accountMenuPanel.hidden && !accountMenuPanel.contains(event.target)) {
      setMenuOpen(false);
    }
  });

  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !accountMenuPanel.hidden) {
      setMenuOpen(false);
      accountMenuButton.focus();
    }
  });
}

// A profile picture can stop loading (revoked URL, offline, provider outage). Drop
// the image so the initials underneath show rather than a broken-image marker.
document.querySelectorAll('.avatar-photo').forEach((photo) => {
  photo.addEventListener('error', () => photo.remove());
});

// Live item lists.
//
// socket.io only ever says "something you can see has changed"; the list itself
// comes back over the ordinary signed-in request, so the server still decides what
// may be on screen and this file never has to know the visibility rules.
//
// Redrawing a long table under someone's eyes and dumping them back at the top is
// worse than a stale row, so every row carries data-item-key. The row crossing a
// line a third of the way down the viewport is remembered before the swap and the
// page is scrolled by the difference afterwards. If that item is gone, the row that
// moved up into its place is used, so the reader drops one item down the list
// instead of losing their spot entirely.

const listRegion = document.getElementById('itemLists');

if (listRegion) {
  const ANCHOR_LINE_RATIO = 0.35;
  const ANCHOR_LINE_MAX = 280;
  const CHANGE_DELAY_MS = 120;
  const RELOAD_ANCHOR_KEY = 'itemshare:scroll-anchor';

  let liveState = 'wait';
  let refreshing = false;
  let queued = false;
  let timer = null;
  let everConnected = false;

  const anchorLine = () => Math.min(window.innerHeight * ANCHOR_LINE_RATIO, ANCHOR_LINE_MAX);

  function setLiveState(state) {
    liveState = state;
    const labels = { live: 'Live', wait: 'Reconnecting', off: 'Not live' };

    listRegion.querySelectorAll('.live-pill').forEach((pill) => {
      pill.hidden = false;
      pill.dataset.live = state;
      pill.textContent = labels[state] || state;
    });
  }

  // The row the reader is looking at: the first one crossing the anchor line. Null
  // when the list is off screen, because then nothing on it has to hold its place.
  function captureAnchor() {
    const rows = Array.from(listRegion.querySelectorAll('tr[data-item-key]'));

    if (rows.length === 0) {
      return null;
    }

    const line = anchorLine();
    let index = rows.length - 1;

    for (let i = 0; i < rows.length; i += 1) {
      if (rows[i].getBoundingClientRect().bottom > line) {
        index = i;
        break;
      }
    }

    const rect = rows[index].getBoundingClientRect();

    if (rect.bottom <= 0 || rect.top >= window.innerHeight) {
      return null;
    }

    return { key: rows[index].dataset.itemKey, index, offset: rect.top - line };
  }

  function restoreAnchor(anchor) {
    if (!anchor) {
      return;
    }

    const rows = Array.from(listRegion.querySelectorAll('tr[data-item-key]'));

    if (rows.length === 0) {
      return;
    }

    const match = rows.find((row) => row.dataset.itemKey === anchor.key);
    // Removed item: the row that moved up into its slot is the next one down.
    const target = match || rows[Math.min(anchor.index, rows.length - 1)];
    const delta = target.getBoundingClientRect().top - anchorLine() - anchor.offset;

    if (Math.abs(delta) >= 1) {
      window.scrollBy(0, delta);
    }
  }

  function rememberAcrossReload(anchor) {
    try {
      window.sessionStorage.setItem(RELOAD_ANCHOR_KEY, JSON.stringify(anchor));
    } catch (error) {
      // Private mode or storage turned off: a rebuild simply starts at the top.
    }
  }

  function applyRegion(html) {
    const template = document.createElement('template');
    template.innerHTML = html;

    const next = template.content.querySelector('#itemLists');

    if (!next) {
      // The viewer has no list any more, so the page around it has to go too.
      window.location.reload();
      return;
    }

    const anchor = captureAnchor();

    if (next.getAttribute('data-signature') !== listRegion.getAttribute('data-signature')) {
      // A type or an owner appeared or vanished, so the filter controls are stale.
      rememberAcrossReload(anchor);
      window.location.reload();
      return;
    }

    listRegion.innerHTML = next.innerHTML;
    restoreAnchor(anchor);
    setLiveState(liveState);
  }

  async function refresh() {
    if (refreshing) {
      queued = true;
      return;
    }

    refreshing = true;

    try {
      const response = await fetch(`/items${window.location.search}`, {
        credentials: 'same-origin',
        headers: { Accept: 'text/html' },
        cache: 'no-store',
      });

      if (response.ok) {
        applyRegion(await response.text());
      }
    } catch (error) {
      // Offline, or the server is restarting: the next reconnect refreshes.
    } finally {
      refreshing = false;

      if (queued) {
        queued = false;
        scheduleRefresh();
      }
    }
  }

  function scheduleRefresh(delay = CHANGE_DELAY_MS) {
    if (timer) {
      return;
    }

    timer = window.setTimeout(() => {
      timer = null;
      refresh();
    }, delay);
  }

  const socket = typeof window.io === 'function' ? window.io({ withCredentials: true }) : null;

  if (socket) {
    socket.on('connect', () => {
      setLiveState('live');

      // A reconnect may have missed changes while the tab was asleep or the server
      // was restarting, so the list is fetched once the link is back.
      if (everConnected) {
        scheduleRefresh(0);
      }
      everConnected = true;
    });

    socket.on('disconnect', () => setLiveState('wait'));
    socket.on('connect_error', () => setLiveState('wait'));
    socket.on('items:changed', () => scheduleRefresh());
  } else {
    setLiveState('off');
  }

  // A rebuild for a changed filter list happens in a new page; put the reader back
  // on the row they were on rather than at the top.
  if (typeof window.requestAnimationFrame === 'function') {
    window.requestAnimationFrame(() => {
      let anchor = null;

      try {
        const raw = window.sessionStorage.getItem(RELOAD_ANCHOR_KEY);

        if (raw) {
          window.sessionStorage.removeItem(RELOAD_ANCHOR_KEY);
          anchor = JSON.parse(raw);
        }
      } catch (error) {
        anchor = null;
      }

      if (anchor) {
        restoreAnchor(anchor);
      }
    });
  }
}
