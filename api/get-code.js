import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';

/**
 * 获取 Microsoft OAuth2 Access Token
 */
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
    const err = new 错误('获取 access_token 失败');
    err.details = Object.keys(json || {}).length ? json : { raw: txt };
    throw err;
  }

  if (!json.access_token) {
    const err = new 错误('响应里没有 access_token');
    err.details = json;
    throw err;
  }

  return json.access_token;
}

/**
 * 从文本中提取 6 位数字验证码
 */
function extractOtp(text) {
  if (!text) return null;

  const candidates = [
    text,
    text.替换(/<[^>]+>/g, ' ') // 过滤 HTML 标签
  ];

  const patterns = [
    /验证码(?:为|是)?\s*[:：]?\s*(\d{6})/, // 匹配“验证码是123456”
    /\b(\d{6})\b/                           // 匹配任意独立的6位数字
  ];

  for (const t of candidates) {
    for (const re of patterns) {
      const m = t.match(re);
      if (m) return m[1];
    }
  }
  return null;
}

/**
 * 格式化日期
 */
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

/**
 * 核心逻辑：获取标题包含“验证码”的邮件
 */
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

  const mailboxes = ['INBOX', 'Junk'];
  const results = [];

  try {
    await client.connect();

    for (const box of mailboxes) {
      try {
        await client.mailboxOpen(box);

        // --- 修正点：将 Set 转换为 Array 以后再进行 slice 和 reverse ---
        const searchResult = await client.search({ subject: '验证码' });
        const uidsAll = Array.from(searchResult); 
        
        // 截取最近的邮件
        const uids = uidsAll.slice(-maxPerBox).reverse();

        for (const uid of uids) {
          const msg = await client.fetchOne(uid, { source: true, envelope: true });
          if (!msg?.source) continue;

          const parsed = await simpleParser(msg.source);
          const code = extractOtp(parsed.text || parsed.html);
          if (!code) continue;

          const sentAt = toIsoOrNull(parsed.date) || toIsoOrNull(parsed.headers?.get?.('date'));

          results.push({
            subject: msg.envelope?.subject || '',
            from: parsed.from?.text || '',
            sentAt,
            code    
          });
        }
      } catch (boxErr) {
        console.error(`Folder ${box} error:`, boxErr.message);
      }
    }

    results.sort((a, b) => (b.sentAt || '').localeCompare(a.sentAt || ''));
    return results;
  } finally {
    try { await client.logout(); } catch {}
  }
}

/**
 * API 接口入口
 */
export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ ok: false, error: 'Method Not Allowed' });
    return;
  }

  const { email, client_id, refresh_token, maxPerBox = 100 } = req.body || {};
  
  if (!email || !client_id || !refresh_token) {
    res.status(400).json({ ok: false, error: '缺少必要参数' });
    return;
  }

  // 限制单次查询数量，防止超时
  const max = Math.max(1, Math.min(500, Number(maxPerBox) || 100));

  try {
    const accessToken = await getAccessToken(client_id, refresh_token);
    const mails = await listSendcodeMails(email, accessToken, { maxPerBox: max });

    res。status(200).json({
      ok: true,
      count: mails.length,
      data: mails
    });
  } catch (e) {
    res.status(500).json({
      ok: false,
      error: e?.message || String(e),
      details: e?.details || null
    });
  }
}
