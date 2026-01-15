import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';

// 获取 access token
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

// 提取验证码
function extractOtp(text) {
  if (!text) return null;

  const candidates = [
    text,
    text.replace(/<[^>]+>/g, ' ') // 去掉 HTML 标签
  ];

  const patterns = [
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

// 提取邮件的发送时间
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

// 标题包含 TOFAI（不区分大小写）
function subjectHasTOFAI(subject) {
  if (!subject) return false;
  return subject.toLowerCase().includes('tofai');
}

// 列出符合条件的邮件（发送时间、验证码、标题）
async function listTOFAIMails(email, accessToken, { maxPerBox = 200 } = {}) {
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

        const uidsAll = await client.search({ all: true });
        const uids = uidsAll.slice(-maxPerBox).reverse(); // 最近 maxPerBox 封

        for (const uid of uids) {
          const msg = await client.fetchOne(uid, { source: true });
          if (!msg?.source) continue;

          const parsed = await simpleParser(msg.source);

          // ✅ 过滤：标题包含 TOFAI
          const subject = parsed.subject || '';
          if (!subjectHasTOFAI(subject)) continue;

          // ✅ 提取验证码（正文 text/html）
          const code = extractOtp(parsed.text || parsed.html);
          if (!code) continue;

          // ✅ 发送时间
          const sentAt =
            toIsoOrNull(parsed.date) ||
            toIsoOrNull(parsed.headers?.get?.('date'));

          results.push({
            sentAt,
            code,
            subject
          });
        }
      } catch {
        // 文件夹读取失败时跳过
      }
    }

    // 最新在前
    results.sort((a, b) => String(b.sentAt || '').localeCompare(String(a.sentAt || '')));

    return results;
  } finally {
    try { await client.logout(); } catch {}
  }
}

// API 请求处理
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
    const mails = await listTOFAIMails(email, accessToken, { maxPerBox: max });

    // ✅ 直接返回数组（前端 renderTable(data) 能用）
    res.status(200).json(mails);
  } catch (e) {
    res.status(500).json({
      ok: false,
      error: e?.message || String(e),
      details: e?.details || null
    });
  }
}
