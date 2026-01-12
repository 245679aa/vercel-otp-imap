import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';

export const config = {
  api: { bodyParser: { sizeLimit: '1mb' } }
};

async function getAccessToken(client_id, refresh_token) {
  const form = new URLSearchParams();
  form.set('client_id', client_id);
  form.set('grant_type', 'refresh_token');
  form.set('refresh_token', refresh_token);

  const resp = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: form
  });

  const txt = await resp.text();
  let json = {};
  try { json = JSON.parse(txt); } catch { /* ignore */ }

  if (!resp.ok) {
    const err = new Error('获取 access_token 失败');
    err.details = Object.keys(json || {}).length ? json : { raw: txt };
    throw err;
  }

  if (!json.access_token) {
    const err = new Error('响应里没有 access_token');
    err.details = json;
    throw err;
  }

  return json.access_token;
}

function toIsoOrNull(d) {
  try {
    if (!d) return null;
    const dt = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dt.getTime())) return null;
    return dt.toISOString();
  } catch {
    return null;
  }
}

async function listSendcodeMails(email, accessToken, { maxPerBox = 200 } = {}) {
  const client = new ImapFlow({
    host: 'outlook.office365.com',
    port: 993,
    secure: true,
    auth: {
      user: email,
      accessToken,
      method: 'XOAUTH2'
    },
    logger: false
  });

  const mailboxes = ['Junk', 'INBOX'];
  const results = [];

  try {
    await client.connect();

    for (const box of mailboxes) {
      try {
        await client.mailboxOpen(box);

        // 拉最近 maxPerBox 封，避免全量太慢
        const uidsAll = await client.search({ all: true });
        const uids = uidsAll.slice(-maxPerBox).reverse(); // 新到旧

        for (const uid of uids) {
          const msg = await client.fetchOne(uid, { source: true });
          if (!msg?.source) continue;

          const parsed = await simpleParser(msg.source);

          const fromLower = (parsed.from?.text || '').toLowerCase();
          if (!fromLower.includes('sendcode@alphasmtp.verifycode.link')) continue;

          // 发送时间：优先 parsed.date；其次尝试 header date
          const sentAt = toIsoOrNull(parsed.date) || toIsoOrNull(parsed.headers?.get?.('date'));

          results.push({
            mailbox: box,
            sentAt, // ISO string
            from: parsed.from?.text || '',
            subject: parsed.subject || ''
            // 你如果还想加正文片段：可以加 snippet
            // snippet: (parsed.text || '').slice(0, 200)
          });
        }
      } catch {
        // 单个文件夹失败不影响整体
      }
    }

    // 按时间倒序（最新在前），无时间的排后面
    results.sort((a, b) => {
      const aa = a.sentAt || '';
      const bb = b.sentAt || '';
      return bb.localeCompare(aa);
    });

    return results;
  } finally {
    try { await client.logout(); } catch {}
  }
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ ok: false, error: 'Method Not Allowed' });
    return;
  }

  const { email, client_id, refresh_token, maxPerBox = 200 } = req.body || {};
  if (!email || !client_id || !refresh_token) {
    res.status(400).json({ ok: false, error: '缺少参数：email / client_id / refresh_token' });
    return;
  }

  const max = Math.max(10, Math.min(2000, Number(maxPerBox) || 200));

  try {
    const accessToken = await getAccessToken(client_id, refresh_token);
    const mails = await listSendcodeMails(email, accessToken, { maxPerBox: max });

    res.status(200).json({
      ok: true,
      count: mails.length,
      mails
    });
  } catch (e) {
    res.status(500).json({
      ok: false,
      error: e?.message || String(e),
      details: e?.details || null
    });
  }
}
