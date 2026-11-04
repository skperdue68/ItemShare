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
