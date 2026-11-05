// ItemShare page behaviour: uploads, the account menu and the live item lists.
//
// Everything here is an enhancement. The logout form posts without scripts and the
// filter form still submits on its own, so a failed script costs convenience only.

const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('savedvars');
const fileName = document.getElementById('fileName');
const toast = document.getElementById('toast');

// Filled in by the list code once it can refresh without a page load; until then a
// change has to be shown the only other way there is.
let refreshListsSoon = null;

function refreshAfterChange() {
  if (refreshListsSoon) {
    refreshListsSoon(0);
    return;
  }

  window.location.reload();
}

// ---------------------------------------------------------------------- toast
// One short lived message, replaced by whatever the next step has to say. The
// upload is the only thing here that takes long enough to need reassuring.
let toastHideTimer = null;

function showToast(message, kind = 'info', { sticky = false } = {}) {
  if (!toast) {
    return;
  }

  toast.textContent = message;
  toast.className = `toast ${kind}`;
  toast.hidden = false;
  window.requestAnimationFrame(() => toast.classList.add('visible'));

  window.clearTimeout(toastHideTimer);

  if (sticky) {
    return;
  }

  toastHideTimer = window.setTimeout(() => {
    toast.classList.remove('visible');
    toastHideTimer = window.setTimeout(() => {
      toast.hidden = true;
    }, 220);
  }, kind === 'err' ? 7000 : 4500);
}

// --------------------------------------------------------------------- upload
// Dropping a file or picking one both upload straight away: there is no button to
// press and nothing to confirm first.
async function uploadFile(file) {
  if (!file) {
    return;
  }

  if (fileName) {
    fileName.textContent = `${file.name} (${Math.round(file.size / 1024)} KB)`;
  }
  if (dropZone) {
    dropZone.classList.add('busy');
  }

  showToast(`Received ${file.name} - uploading and processing...`, 'info', { sticky: true });

  try {
    const body = new FormData();
    body.append('savedvars', file);

    const response = await fetch('/upload', { method: 'POST', body });
    const result = await response.json().catch(() => ({}));

    if (!response.ok) {
      throw new Error(result.error || `Upload failed with status ${response.status}`);
    }

    const accounts = result.accounts || [];
    const icons = result.iconLookup && result.iconLookup.resolved
      ? `Found ${result.iconLookup.resolved} icon${result.iconLookup.resolved === 1 ? '' : 's'} that were missing.`
      : '';

    showToast(
      [
        `${file.name} has been received.`,
        ...accounts.map(
          (account) =>
            `${account.accountName}: ${account.itemCount} item${account.itemCount === 1 ? '' : 's'}, quantity ${account.totalQuantity}`
        ),
        icons,
      ]
        .filter(Boolean)
        .join('\n'),
      'ok'
    );

    refreshAfterChange();
  } catch (error) {
    showToast(error.message, 'err');
  } finally {
    if (dropZone) {
      dropZone.classList.remove('busy');
    }
    // Clearing the input means the same file can be chosen again straight away.
    if (fileInput) {
      fileInput.value = '';
    }
    if (fileName) {
      fileName.textContent = 'No file selected';
    }
  }
}

if (dropZone && fileInput) {
  fileInput.addEventListener('change', () => {
    const file = fileInput.files && fileInput.files[0];
    if (file) {
      uploadFile(file);
    }
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
      uploadFile(dropped[0]);
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
  const SEARCH_DELAY_MS = 250;
  const RELOAD_ANCHOR_KEY = 'itemshare:scroll-anchor';
  const TAB_KEY = 'itemshare:tab';

  let liveState = 'wait';
  let refreshing = false;
  let queued = false;
  let timer = null;
  let everConnected = false;
  let activeTab = 'request';

  try {
    const stored = window.sessionStorage.getItem(TAB_KEY);
    if (stored === 'request' || stored === 'offer') {
      activeTab = stored;
    }
  } catch (error) {
    // No storage: the request tab is the default, which is also the useful one.
  }

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

  // ---------------------------------------------------------------------- tabs
  // Request and Offering are the same data cut two ways, so the tab is a display
  // choice the server does not need to know about. It survives a live refresh (the
  // region is re-rendered from the Request tab) and a page rebuild.
  function setTab(name, { focus = false } = {}) {
    const tabs = Array.from(listRegion.querySelectorAll('[data-tab]'));
    const panels = Array.from(listRegion.querySelectorAll('[data-panel]'));

    if (!tabs.some((tab) => tab.dataset.tab === name)) {
      return;
    }

    activeTab = name;

    tabs.forEach((tab) => {
      const selected = tab.dataset.tab === name;
      tab.setAttribute('aria-selected', String(selected));
      tab.tabIndex = selected ? 0 : -1;

      if (selected && focus) {
        tab.focus();
      }
    });

    panels.forEach((panel) => {
      panel.hidden = panel.dataset.panel !== name;
    });

    try {
      window.sessionStorage.setItem(TAB_KEY, name);
    } catch (error) {
      // Tab choice is a convenience; failing to remember it changes nothing.
    }
  }

  listRegion.addEventListener('click', (event) => {
    const closest = (selector) => (event.target.closest ? event.target.closest(selector) : null);
    const tab = closest('[data-tab]');

    if (tab) {
      event.preventDefault();
      setTab(tab.dataset.tab, { focus: true });
      return;
    }

    // A column heading is also the sort control: clicking sorts by it, clicking
    // again reverses. The list is fetched instead of reloading the page.
    const sortLink = closest('[data-sort]');

    if (sortLink) {
      event.preventDefault();
      sortState = {
        field: sortLink.dataset.sort,
        dir: sortLink.dataset.dir === 'desc' ? 'desc' : 'asc',
      };

      const params = new URLSearchParams(window.location.search);
      params.set('sort', sortState.field);

      if (sortState.dir === 'desc') {
        params.set('dir', 'desc');
      } else {
        params.delete('dir');
      }

      if (window.history && window.history.replaceState) {
        window.history.replaceState(null, '', `/?${params.toString()}`);
      }

      scheduleRefresh(0);
      return;
    }

    const requestButton = closest('[data-request]');

    if (requestButton) {
      event.preventDefault();
      openRequestDialog(requestButton);
      return;
    }

    const cancelButton = closest('[data-cancel-request]');

    if (cancelButton) {
      event.preventDefault();
      cancelOwnRequest(cancelButton);
    }
  });

  listRegion.addEventListener('keydown', (event) => {
    if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') {
      return;
    }

    const tabs = Array.from(listRegion.querySelectorAll('[data-tab]'));
    const index = tabs.findIndex((tab) => tab.dataset.tab === activeTab);

    if (index === -1) {
      return;
    }

    event.preventDefault();
    const step = event.key === 'ArrowRight' ? 1 : tabs.length - 1;
    setTab(tabs[(index + step) % tabs.length].dataset.tab, { focus: true });
  });

  setTab(activeTab);

  // ------------------------------------------------------------------- filters
  // Filters apply as soon as they change - there is no button to forget - and the
  // address bar keeps up, so a reload or a live refresh uses the same filters.
  const filterForm = document.querySelector('[data-filters]');

  // The sort order is not a filter, so it is not in the form; it is kept here and
  // re-applied with whatever the form says, which is why changing a filter never
  // silently resets the order the reader chose.
  let sortState = (() => {
    const params = new URLSearchParams(window.location.search);
    const sort = params.get('sort');

    return {
      field: sort === 'name' || sort === 'added' ? sort : 'type',
      dir: params.get('dir') === 'desc' ? 'desc' : 'asc',
    };
  })();

  function applyFilters() {
    if (!filterForm) {
      return;
    }

    const params = new URLSearchParams();

    filterForm.querySelectorAll('[data-filter]').forEach((field) => {
      const value = String(field.value || '').trim();
      if (value) {
        params.set(field.name, value);
      }
    });

    if (sortState.field !== 'type') {
      params.set('sort', sortState.field);

      if (sortState.dir === 'desc') {
        params.set('dir', 'desc');
      }
    }

    const query = params.toString();

    if (window.history && window.history.replaceState) {
      window.history.replaceState(null, '', query ? `/?${query}` : '/');
    }

    scheduleRefresh(0);
  }

  if (filterForm) {
    let searchTimer = null;

    filterForm.addEventListener('change', (event) => {
      if (event.target.matches('[data-filter]')) {
        applyFilters();
      }
    });

    // Typing filters as it is typed, just slowly enough not to fetch on every key.
    filterForm.addEventListener('input', (event) => {
      if (!event.target.matches('input[type="search"][data-filter]')) {
        return;
      }

      window.clearTimeout(searchTimer);
      searchTimer = window.setTimeout(applyFilters, SEARCH_DELAY_MS);
    });

    filterForm.addEventListener('submit', (event) => {
      event.preventDefault();
      applyFilters();
    });

    const clearButton = filterForm.querySelector('[data-clear-filters]');

    if (clearButton) {
      clearButton.addEventListener('click', () => {
        filterForm.querySelectorAll('[data-filter]').forEach((field) => {
          field.value = '';
        });
        applyFilters();
      });
    }
  }

  // -------------------------------------------------------- keeping the place
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
    setTab(activeTab);
    noteClosedRequests();
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

  refreshListsSoon = (delay = 0) => scheduleRefresh(delay);

  // ---------------------------------------------------------------- requests
  // The button carries how many are still unspoken for, so the dialog cannot offer
  // a number that is not there. One, the whole stack, or anything in between - the
  // last two only when the owner has more than one.
  let dialog = null;
  const knownClosed = new Set();
  let requestsPrimed = false;

  function escapeText(value) {
    const node = document.createElement('span');
    node.textContent = String(value || '');
    return node.innerHTML;
  }

  function closeRequestDialog() {
    if (dialog) {
      dialog.remove();
      dialog = null;
    }

    document.removeEventListener('keydown', onDialogKey);
  }

  function onDialogKey(event) {
    if (event.key === 'Escape') {
      closeRequestDialog();
    }
  }

  function openRequestDialog(button) {
    const kind = button.dataset.kind === 'offer' ? 'offer' : 'inventory';
    const max = Math.max(1, Number(button.dataset.max) || 1);
    const stack = Math.max(max, Number(button.dataset.stack) || max);
    const name = button.dataset.name || 'that item';
    const choices = kind === 'inventory' && max > 1;

    closeRequestDialog();

    const backdrop = document.createElement('div');
    backdrop.className = 'modal-backdrop';
    backdrop.innerHTML = `
      <div class="modal" role="dialog" aria-modal="true" aria-labelledby="requestTitle">
        <h2 id="requestTitle">Ask for ${escapeText(name)}</h2>
        <p class="muted">${
          kind === 'offer'
            ? 'This is a standing offer, so it means one copy.'
            : `${max} still unspoken for${stack > max ? ` of ${stack}` : ''}.`
        }</p>
        <form class="request-form" data-request-form>
          ${
            choices
              ? `<label class="request-option"><input type="radio" name="how" value="one" checked> Just one</label>
                 <label class="request-option"><input type="radio" name="how" value="stack"> The whole stack (${max})</label>
                 <label class="request-option"><input type="radio" name="how" value="number">
                   <input class="request-number" type="number" name="quantity" min="1" max="${max}" value="${max}" aria-label="How many"> of them
                 </label>`
              : '<label class="request-option"><input type="radio" name="how" value="one" checked> One</label>'
          }
          <div class="modal-actions">
            <button class="filter-clear" type="button" data-close-dialog>Cancel</button>
            <button class="primary-button" type="submit">Send request</button>
          </div>
        </form>
      </div>`;

    document.body.appendChild(backdrop);
    dialog = backdrop;
    document.addEventListener('keydown', onDialogKey);

    backdrop.addEventListener('click', (event) => {
      if (event.target === backdrop || (event.target.closest && event.target.closest('[data-close-dialog]'))) {
        closeRequestDialog();
      }
    });

    const numberInput = backdrop.querySelector('.request-number');

    if (numberInput) {
      backdrop.querySelectorAll('input[name="how"]').forEach((radio) => {
        radio.addEventListener('change', () => {
          if (radio.checked && radio.value === 'number') {
            numberInput.focus();
            numberInput.select();
          }
        });
      });
    }

    const form = backdrop.querySelector('[data-request-form]');

    form.addEventListener('submit', (event) => {
      event.preventDefault();

      const how = form.querySelector('input[name="how"]:checked');
      let quantity = 1;

      if (choices && how && how.value === 'stack') {
        quantity = max;
      } else if (choices && how && how.value === 'number') {
        quantity = Math.min(max, Math.max(1, Number(numberInput.value) || 1));
      }

      closeRequestDialog();
      sendRequest({ kind, id: button.dataset.id, quantity, name });
    });

    const first = backdrop.querySelector('input[name="how"]');

    if (first) {
      first.focus();
    }
  }

  async function sendRequest({ kind, id, quantity, name }) {
    try {
      const response = await fetch('/requests', {
        method: 'POST',
        credentials: 'same-origin',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({ kind, id, quantity }),
      });
      const result = await response.json().catch(() => ({}));

      if (!response.ok) {
        throw new Error(result.error || `That request failed with status ${response.status}`);
      }

      showToast(
        quantity > 1 ? `Asked for ${quantity} of ${name}.` : `Asked for ${name}.`,
        'ok'
      );
      refreshAfterChange();
    } catch (error) {
      showToast(error.message, 'err');
    }
  }

  async function cancelOwnRequest(button) {
    try {
      const response = await fetch(`/requests/${encodeURIComponent(button.dataset.id)}/cancel`, {
        method: 'POST',
        credentials: 'same-origin',
      });
      const result = await response.json().catch(() => ({}));

      if (!response.ok) {
        throw new Error(result.error || 'That request could not be cancelled.');
      }

      showToast('Request cancelled.', 'ok');
      refreshAfterChange();
    } catch (error) {
      showToast(error.message, 'err');
    }
  }

  // A request can be closed by someone else's upload, which arrives as an ordinary
  // list refresh. The first pass only records what is already closed, so opening the
  // page does not announce history; after that, anything newly cancelled is news.
  function noteClosedRequests() {
    const closed = Array.from(
      listRegion.querySelectorAll('[data-my-request][data-status="cancelled"]')
    );
    const fresh = [];

    for (const row of closed) {
      const id = row.dataset.requestId;

      if (!id || knownClosed.has(id)) {
        continue;
      }

      knownClosed.add(id);

      if (requestsPrimed) {
        fresh.push(row);
      }
    }

    if (!requestsPrimed) {
      requestsPrimed = true;
      return;
    }

    if (fresh.length > 0) {
      showToast(
        fresh
          .map(
            (row) =>
              `Your request for ${row.dataset.requestName || 'an item'} was cancelled: ${
                row.dataset.requestReason || 'the owner no longer has it.'
              }`
          )
          .join('\n'),
        'err'
      );
    }
  }

  noteClosedRequests();

  // ------------------------------------------------------------------- socket
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

// ------------------------------------------------------- unsent request basket
// A request holds its item before it is sent, so the clock matters twice: the person
// needs to see how long they have left, and walking away would strand a reservation
// nobody knows about. Both live here.

function basketHeldItems() {
  const basket = document.querySelector('[data-basket]');
  return basket ? Number(basket.dataset.items) || 0 : 0;
}

// The server sends the seconds left as it renders the page; from then on this counts
// down against the clock, so the number moves without one request per second. A live
// refresh replaces the element, which restarts it from whatever the server now says.
function tickCountdown() {
  const basket = document.querySelector('[data-basket]');
  const label = basket ? basket.querySelector('[data-countdown]') : null;

  if (!basket || !label) {
    return;
  }

  if (!basket.dataset.deadline) {
    basket.dataset.deadline = String(Date.now() + (Number(basket.dataset.seconds) || 0) * 1000);
  }

  const remaining = Math.max(0, Math.round((Number(basket.dataset.deadline) - Date.now()) / 1000));
  const minutes = Math.floor(remaining / 60);

  label.textContent = `${minutes}:${String(remaining % 60).padStart(2, '0')}`;
  label.classList.toggle('countdown-urgent', remaining <= 30);

  // Zero is when the server has already let the items go; saying so is better than
  // leaving a basket on screen that can no longer be sent.
  if (remaining === 0 && !basket.dataset.lapsed) {
    basket.dataset.lapsed = '1';
    showToast('The time to send those requests ran out, so the items went back to being available.', 'err');
    refreshAfterChange();
  }
}

window.setInterval(tickCountdown, 1000);
tickCountdown();

async function sendBasket(button) {
  button.disabled = true;

  try {
    const response = await fetch('/requests/submit', {
      method: 'POST',
      credentials: 'same-origin',
    });
    const result = await response.json().catch(() => ({}));

    if (!response.ok) {
      throw new Error(result.error || 'Those requests could not be sent.');
    }

    showToast(
      result.sent === 1 ? 'Request sent.' : `${result.sent} requests sent.`,
      'ok'
    );
    refreshAfterChange();
  } catch (error) {
    showToast(error.message, 'err');
  } finally {
    button.disabled = false;
  }
}

document.addEventListener('click', (event) => {
  const submit = event.target.closest ? event.target.closest('[data-submit-requests]') : null;

  if (submit) {
    event.preventDefault();
    sendBasket(submit);
  }
});

// ------------------------------------------------------------ leaving the page
// The browser's own beforeunload dialog cannot be styled, so the prompt here is a
// real one and beforeunload is only the last resort for a closed tab or a form post.
let leavePrompt = null;

function closeLeavePrompt() {
  if (leavePrompt) {
    leavePrompt.remove();
    leavePrompt = null;
  }

  document.removeEventListener('keydown', onLeavePromptKey);
}

function onLeavePromptKey(event) {
  if (event.key === 'Escape') {
    closeLeavePrompt();
  }
}

function openLeavePrompt(href) {
  const held = basketHeldItems();

  closeLeavePrompt();

  const backdrop = document.createElement('div');
  backdrop.className = 'modal-backdrop';
  backdrop.innerHTML = `
    <div class="modal" role="dialog" aria-modal="true" aria-labelledby="leaveTitle">
      <h2 id="leaveTitle">You have unsubmitted item requests</h2>
      <p class="muted" data-leave-sentence></p>
      <p class="muted">
        They are held for a few more minutes. If you leave without sending them, the items
        go back to being available to everyone else - which is what will happen anyway when
        the clock runs out.
      </p>
      <div class="leave-actions">
        <button class="filter-clear" type="button" data-stay>Stay on this page</button>
        <button class="primary-button" type="button" data-leave-anyway>Leave anyway</button>
      </div>
    </div>`;

  // Written as text rather than markup: it is a count, but the habit is worth keeping.
  backdrop.querySelector('[data-leave-sentence]').textContent =
    held === 1 ? '1 item is waiting to be sent.' : `${held} items are waiting to be sent.`;

  document.body.appendChild(backdrop);
  leavePrompt = backdrop;
  document.addEventListener('keydown', onLeavePromptKey);

  backdrop.addEventListener('click', (event) => {
    if (event.target === backdrop || (event.target.closest && event.target.closest('[data-stay]'))) {
      closeLeavePrompt();
      return;
    }

    if (event.target.closest && event.target.closest('[data-leave-anyway]')) {
      closeLeavePrompt();
      window.location.assign(href);
    }
  });

  const stay = backdrop.querySelector('[data-stay]');

  if (stay) {
    stay.focus();
  }
}

document.addEventListener('click', (event) => {
  if (event.defaultPrevented || event.button !== 0) {
    return;
  }
  if (event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) {
    return;
  }
  if (basketHeldItems() === 0) {
    return;
  }

  const link = event.target.closest ? event.target.closest('a[href]') : null;

  if (!link || link.target === '_blank' || link.hasAttribute('download')) {
    return;
  }

  const href = link.getAttribute('href') || '';

  if (!href || href.startsWith('#') || href.startsWith('javascript:')) {
    return;
  }

  event.preventDefault();
  openLeavePrompt(href);
});

window.addEventListener('beforeunload', (event) => {
  if (basketHeldItems() === 0) {
    return undefined;
  }

  event.preventDefault();
  event.returnValue = 'You have unsubmitted item requests.';
  return event.returnValue;
});

// ------------------------------------------------------------- ESO accounts
// Which accounts are linked is reference material: it lives behind the profile menu,
// in the modal the page already rendered, rather than above the item list.
const accountsModal = document.getElementById('accountsModal');

function closeAccountsModal() {
  if (!accountsModal || accountsModal.hidden) {
    return;
  }

  accountsModal.hidden = true;

  const opener = document.querySelector('[data-open-accounts]');

  if (opener) {
    opener.focus();
  }
}

function openAccountsModal() {
  if (!accountsModal) {
    return;
  }

  accountsModal.hidden = false;

  const close = accountsModal.querySelector('[data-close-accounts]');

  if (close) {
    close.focus();
  }
}

document.addEventListener('click', (event) => {
  const closest = event.target.closest ? event.target.closest.bind(event.target) : () => null;

  if (!accountsModal) {
    return;
  }

  if (closest('[data-open-accounts]')) {
    event.preventDefault();
    openAccountsModal();
    return;
  }

  if (closest('[data-close-accounts]')) {
    event.preventDefault();
    closeAccountsModal();
    return;
  }

  // The backdrop is the modal's own parent, so a click on it is a click outside.
  if (!accountsModal.hidden && event.target === accountsModal) {
    closeAccountsModal();
  }
});

document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && accountsModal && !accountsModal.hidden) {
    closeAccountsModal();
  }
});
