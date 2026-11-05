// Request emails.
//
// The site sends these through Resend's HTTP API (`POST /emails`), so nothing has to
// be installed and the key stays in Database/.env:
//
//   RESEND_API_KEY=re_...
//   EMAIL_FROM=ItemShare <requests@your-domain.example>
//
// With either missing the site still works and requests still send - the server log
// says the email was skipped, which is an honest state rather than a silent failure.
//
// The message is built here rather than in views.js because email clients are not
// browsers: the layout is a table with inline styles, and there is a plain-text copy
// for clients that refuse HTML.

const ENDPOINT = 'https://api.resend.com/emails';

// Quality colours that survive a white mail body: the game's own white would vanish
// there, so the ordinary tier is a dark grey instead.
const QUALITY_COLOURS = {
  1: '#4a4a4a',
  2: '#2e8b3a',
  3: '#2f6fed',
  4: '#7a3fd1',
  5: '#a87400',
};

function qualityColour(quality) {
  return QUALITY_COLOURS[Number(quality) || 0] || '#3c4557';
}

/**
 * What tells two copies of the same name apart: its type, its trait and, when it has
 * one, its enchantment. Two Briarheart swords with different traits are the same
 * name, so this is the line the owner actually needs to pick the right one.
 */
function itemDetailText(item) {
  return [item.type, item.trait, item.enchant ? `Enchant: ${item.enchant}` : '']
    .map((part) => String(part || '').trim())
    .filter(Boolean)
    .join(' \u00b7 ');
}

function escape(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

export function emailConfig(env = process.env) {
  const apiKey = String(env.RESEND_API_KEY || '').trim();
  const from = String(env.EMAIL_FROM || '').trim();
  const siteUrl = String(env.BASE_URL || '')
    .trim()
    .replace(/\/+$/, '');

  return { apiKey, from, siteUrl, configured: Boolean(apiKey && from) };
}

/**
 * One request email: what was asked for, how many of each, and the ESO name to mail
 * it to. `items` is a plain list, because the caller has already grouped the requests
 * by the person who owns them.
 */
export function renderRequestEmail({
  requester = '',
  requesterAccount = '',
  ownerAccount = '',
  items = [],
  siteUrl = '',
  sentAt = new Date(),
}) {
  const who = requesterAccount || requester || 'Someone';
  const total = items.reduce((sum, item) => sum + (Number(item.quantity) || 1), 0);
  const subject = `ESO Item Request from ${who}${items.length > 1 ? ` (${items.length} items)` : ''}`;
  const when = sentAt instanceof Date ? sentAt.toUTCString() : String(sentAt);

  const rows = items
    .map((item) => {
      const details = itemDetailText(item);

      return `
              <tr>
                <td style="padding:8px 0;border-bottom:1px solid #e7e9ef;font-size:14px">
                  <span style="color:${qualityColour(
                    item.quality
                  )};font-weight:600">${escape(item.name)}</span>${
                    details
                      ? `<div style="margin-top:2px;font-size:12px;color:#6b7280">${escape(
                          details
                        )}</div>`
                      : ''
                  }
                </td>
                <td style="padding:8px 0;border-bottom:1px solid #e7e9ef;font-size:14px;text-align:right;font-variant-numeric:tabular-nums;vertical-align:top">${
                  Number(item.quantity) || 1
                }</td>
              </tr>`;
    })
    .join('');

  const html = `<!DOCTYPE html>
<html lang="en">
<body style="margin:0;padding:0;background:#f4f5f7;color:#1f2430;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f4f5f7;padding:24px 12px">
    <tr>
      <td align="center">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:560px;background:#ffffff;border-radius:10px;border:1px solid #e3e6ee">
          <tr>
            <td style="padding:16px 22px;background:#1b2440;border-radius:10px 10px 0 0;color:#eef2ff;font-size:12px;letter-spacing:.12em;text-transform:uppercase">
              ESO Item Request
            </td>
          </tr>
          <tr>
            <td style="padding:22px">
              <h1 style="margin:0 0 10px;font-size:19px;line-height:1.3">${escape(
                who
              )} asked you for ${items.length === 1 ? 'an item' : `${items.length} items`}</h1>
              <p style="margin:0 0 18px;font-size:14px;line-height:1.55;color:#3c4557">
                These are shared from <strong>${escape(
                  ownerAccount
                )}</strong> on ItemShare, and someone has asked for them. Mail ${
                  total === 1 ? 'it' : 'them'
                } to <strong>${escape(who)}</strong> in game when you are next online.
              </p>
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:14px">
                <tr>
                  <th align="left" style="padding:0 0 6px;font-size:12px;letter-spacing:.06em;text-transform:uppercase;color:#7a8395">Item</th>
                  <th align="right" style="padding:0 0 6px;font-size:12px;letter-spacing:.06em;text-transform:uppercase;color:#7a8395">Qty</th>
                </tr>${rows}
              </table>
              <p style="margin:18px 0 0;font-size:13px;color:#5b6478">
                Sent from ItemShare on ${escape(when)}.${
                  siteUrl
                    ? ` <a href="${escape(siteUrl)}/" style="color:#2f6fed">Open your list</a> to cancel or change what you share.`
                    : ''
                }
              </p>
            </td>
          </tr>
        </table>
        <p style="margin:14px 0 0;font-size:12px;color:#8a92a3">
          You are getting this because your ESO account ${escape(
            ownerAccount
          )} shares these items on ItemShare.
        </p>
      </td>
    </tr>
  </table>
</body>
</html>`;

  const text = [
    'ESO Item Request',
    '',
    `${who} asked you for these items from ${ownerAccount} on ItemShare:`,
    '',
    ...items.map((item) => {
      const details = itemDetailText(item);
      return `  ${Number(item.quantity) || 1} x ${item.name}${details ? ` - ${details}` : ''}`;
    }),
    '',
    `Mail ${total === 1 ? 'it' : 'them'} to ${who} in game when you are next online.`,
    `Sent ${when}.`,
    siteUrl ? `${siteUrl}/` : '',
  ]
    .filter((line) => line !== '')
    .join('\n');

  return { subject, html, text };
}

/**
 * Sends one message. Returns `{ ok }` or `{ ok: false, error }` rather than throwing:
 * a mail provider having a bad day must never turn a sent request into an error page.
 */
export async function sendEmail({ to, subject, html, text }, config = emailConfig()) {
  if (!config.configured) {
    return { ok: false, error: 'email is not configured' };
  }

  if (!to) {
    return { ok: false, error: 'no address to send to' };
  }

  try {
    const response = await fetch(ENDPOINT, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${config.apiKey}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ from: config.from, to: [to], subject, html, text }),
    });

    if (!response.ok) {
      const detail = await response.text().catch(() => '');
      return { ok: false, error: `${response.status} ${detail.slice(0, 200)}`.trim() };
    }

    const result = await response.json().catch(() => ({}));
    return { ok: true, id: result.id || '' };
  } catch (error) {
    return { ok: false, error: error.message };
  }
}
