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
  // 匹配 4-8 位纯数字验证码，适应更多场景
  const match = text.替换(/<[^>]+>/g, ' ').match(/\b(\d{4,8})\b/);
  return match ? match[1] : null;
}

async function listSendcodeMails(email, accessToken, maxPerBox = 50) {
  const client = new ImapFlow({
    host: 'outlook.office365.com',
    port: 993,
    secure: true,
    auth: {
      user: email,
      accessToken: accessToken,
      method: 'XOAUTH2'
    },
    logger: false // 如果还是获取不到，可以改为 console 查看底层通讯
  });

  const results = [];
  try {
    await client.connect();

    // 遍历常见的文件夹名称
    const targetBoxes = ['INBOX', 'Junk', 'Archive'];
    
    for (const boxName of targetBoxes) {
      try {
        let mailbox = await client.mailboxOpen(boxName);
        if (!mailbox) continue;

        // 优化搜索：搜索标题包含“码”或“Code”的邮件，扩大搜索面
        // 注意：部分服务器对中文搜索支持有限，这里使用 OR 条件
        const searchResult = await client.search({
          or: [
            { subject: '验证码' },
            { subject: 'code' },
            { subject: 'verification' }
          ]
        });

        const uids = Array.from(searchResult).slice(-maxPerBox).reverse();

        for (const uid of uids) {
          const fetchRes = await client.fetchOne(uid, { source: true, envelope: true });
          if (!fetchRes || !fetchRes.source) continue;

          const parsed = await simpleParser(fetchRes.source);
          const code = extractOtp(parsed.text || parsed.html);
          
          if (code) {
            results.push({
              subject: fetchRes.envelope.subject || '',
              from: parsed.from?.text || '',
              sentAt: parsed.date ? parsed.date.toISOString() : null,
              code: code
            });
          }
        }
      } catch (e) {
        // 忽略单个文件夹打开失败的情况
      }
    }
    
    // 按时间排序
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
    
    res.status(200).json({ 
      ok: true, 
      count: data.length,
      data: data 
    });
  } catch (err) {
    res.status(500).json({ 
      ok: false, 
      error: err.message 
    });
  }
}
