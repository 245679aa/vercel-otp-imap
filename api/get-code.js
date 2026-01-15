import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';

async function getAccessToken(clientId, refreshToken) {
  const params = new URLSearchParams();
  params.set('client_id', clientId);
  params.set('grant_type', 'refresh_token');
  params.set('refresh_token', refreshToken);

  const response = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params
  });

  const data = await response.json();
  if (!response.ok) {
    throw new Error(data.error_description || 'Auth Failed');
  }
  return data.access_token;
}

function extractOtp(text) {
  if (!text) return null;
  const cleanText = text.replace(/<[^>]+>/g, ' ');
  const match = cleanText.match(/\b(\d{6})\b/);
  return match ? match[1] : null;
}

async function listSendcodeMails(email, accessToken, maxPerBox = 100) {
  const client = new ImapFlow({
    host: 'outlook.office365.com',
    port: 993,
    secure: true,
    auth: {
      user: email,
      accessToken: accessToken,
      method: 'XOAUTH2'
    },
    logger: false
  });

  const results = [];
  try {
    await client.connect();
    for (const box of ['INBOX', 'Junk']) {
      try {
        await client.mailboxOpen(box);
        const searchResult = await client.search({ subject: '验证码' });
        const uids = Array.from(searchResult).slice(-maxPerBox).reverse();

        for (const uid of uids) {
          const fetchRes = await client.fetchOne(uid, { source: true, envelope: true });
          if (!fetchRes || !fetchRes.source) continue;

          const parsed = await simpleParser(fetchRes.source);
          const code = extractOtp(parsed.text || parsed.html);
          if (!code) continue;

          results.push({
            subject: fetchRes.envelope.subject || '',
            from: parsed.from?.text || '',
            sentAt: parsed.date ? parsed.date.toISOString() : null,
            code: code
          });
        }
      } catch (e) {
        // Skip box if inaccessible
      }
    }
    return results.sort((a, b) => (b.sentAt || '').localeCompare(a.sentAt || ''));
  } finally {
    await client.logout();
  }
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method Not Allowed' });
  }

  const { email, client_id, refresh_token, maxPerBox = 50 } = req.body || {};
  if (!email || !client_id || !refresh_token) {
    return res.status(400).json({ error: 'Missing parameters' });
  }

  try {
    const token = await getAccessToken(client_id, refresh_token);
    const data = await listSendcodeMails(email, token, Number(maxPerBox));
    res.status(200).json({ ok: true, data });
  } catch (err) {
    res.status(500).json({ 
      ok: false, 
      error: err.message || 'Internal Server Error'
    });
  }
}
