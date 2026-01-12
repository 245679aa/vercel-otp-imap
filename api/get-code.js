import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';

export const config = {
  api: { bodyParser: { sizeLimit: '1mb' } }
};

function sleep(ms){ return new Promise(r => setTimeout(r, ms)); }

async function getAccessToken(client_id, refresh_token){
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
    err.details = json && Object.keys(json).length ? json : { raw: txt };
    throw err;
  }

  if (!json.access_token) {
    const err = new Error('响应里没有 access_token');
    err.details = json;
    throw err;
  }
  return json.access_token;
}

function buildXoauth2Token(email, accessToken){
  // IMAP XOAUTH2: base64("user=<email>\x01auth=Bearer <token>\x01\x01")
  const authStr = `user=${email}\x01auth=Bearer ${accessToken}\x01\x01`;
  return Buffer.from(authStr, 'utf8').toString('base64');
}

function extractOtp(text){
  if (!text) return null;

  const candidates = [
    text,
    // 简单去 HTML tag
    text.replace(/<[^>]+>/g, ' ')
  ];

  const patterns = [
    /letter-spacing:\s*25px[^>]*>\s*(\d{6})/i,
    /验证码(?:为|是)?\s*[:：]?\s*(\d{6})/,
    /\b(\d{6})\b/
  ];

  for (const t of candidates) {
    for (const re of patterns) {
      const m = t.match(re);
      if (m) return m[1];
    }
  }
  return null;
}

async function findCodeViaImap(email, accessToken, { timeoutSec, intervalSec }){
  const client = new ImapFlow({
    host: 'outlook.office365.com',
    port: 993,
    secure: true,
    auth: {
      user: email,
      accessToken,      // ImapFlow 会自动用 XOAUTH2（等价）
      method: 'XOAUTH2'
    },
    logger: false
  });

  const deadline = Date.now() + timeoutSec * 1000;

  try {
    await client.connect();

    const mailboxes = ['Junk', 'INBOX'];

    while (Date.now() < deadline) {
      for (const box of mailboxes) {
        try {
          await client.mailboxOpen(box);

          // 拉最近 50 封（够用了），从新到旧
          // 如果你要更彻底，可以扩大范围，但会更慢
          const searchUids = await client.search({ all: true });
          const uids = searchUids.slice(-50).reverse();

          for (const uid of uids) {
            const msg = await client.fetchOne(uid, { source: true, envelope: true });
            if (!msg?.source) continue;

            const parsed = await simpleParser(msg.source);
            const from = (parsed.from?.text || '').toLowerCase();

            if (!from.includes('sendcode@alphasmtp.verifycode.link')) continue;

            const text = (parsed.text || '') + '\n' + (parsed.html || '');
            const code = extractOtp(text);
            if (code) {
              return { code, mailbox: box, from: parsed.from?.text || '' };
            }
          }
        } catch (e) {
          // 某些账号/语言环境文件夹名可能不同；这里不中断整体轮询
        }
      }

      await sleep(intervalSec * 1000);
    }

    return null;
  } finally {
    try { await client.logout(); } catch {}
  }
}

export default async function handler(req, res){
  if (req.method !== 'POST') {
    res.status(405).json({ ok: false, error: 'Method Not Allowed' });
    return;
  }

  const { email, client_id, refresh_token, timeoutSec = 300, intervalSec = 5 } = req.body || {};
  if (!email || !client_id || !refresh_token) {
    res.status(400).json({ ok: false, error: '缺少参数：email / client_id / refresh_token' });
    return;
  }

  try {
    const accessToken = await getAccessToken(client_id, refresh_token);

    // 这里额外构造一下 XOAUTH2 字符串也行（用于你调试/比对）
    const xoauth2 = buildXoauth2Token(email, accessToken);

    const found = await findCodeViaImap(email, accessToken, {
      timeoutSec: Math.max(10, Math.min(1800, Number(timeoutSec) || 300)),
      intervalSec: Math.max(2, Math.min(60, Number(intervalSec) || 5))
    });

    if (!found) {
      res.status(200).json({ ok: false, error: '超时未找到验证码邮件' });
      return;
    }

    res.status(200).json({
      ok: true,
      code: found.code,
      mailbox: found.mailbox,
      from: found.from,
      // xoauth2 仅调试用，建议你用完就删掉这行，避免泄漏
      debug: { xoauth2_preview: xoauth2.slice(0, 20) + '...' }
    });
  } catch (e) {
    res.status(500).json({
      ok: false,
      error: e?.message || String(e),
      details: e?.details || null
    });
  }
}
