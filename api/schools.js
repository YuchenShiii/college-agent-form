import { google } from 'googleapis';

function getAuth() {
  return new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_CLIENT_EMAIL,
      private_key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

async function writeLog(sheets, spreadsheetId, { operator, operatorType, studentEmail, action, detail }) {
  const ts = new Date().toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai' });
  try {
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: 'logs!A:F',
      valueInputOption: 'RAW',
      requestBody: { values: [[ts, operator, operatorType, studentEmail, action, detail]] },
    });
  } catch(e) { /* logs tab 不存在时静默失败 */ }
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  const spreadsheetId = process.env.SHEET_ID;

  if (req.method === 'GET') {
    try {
      const { email } = req.query;
      const resp = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: 'schools!A:I',
      });
      const rows = resp.data.values || [];
      if (rows.length <= 1) return res.json({ success: true, schools: [] });
      const headers = rows[0];
      const data = rows.slice(1)
        .map(row => {
          const obj = {};
          headers.forEach((h, j) => obj[h] = (row[j] || '').trim());
          return obj;
        })
        .filter(r => !email || r['学生邮箱'] === email.trim());
      return res.json({ success: true, schools: data });
    } catch (err) {
      return res.status(500).json({ success: false, error: err.message });
    }
  }

  if (req.method === 'POST') {
    try {
      const { studentEmail, studentName, schools, operator } = req.body;
      if (!studentEmail || !schools.length) {
        return res.status(400).json({ success: false, error: '缺少必要字段' });
      }

      const ts = new Date().toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai' });

      // 读取 schools tab 全部数据
      const existing = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: 'schools!A:I',
      });
      const allRows = existing.data.values || [];
      const headers = allRows.length > 0
        ? allRows[0]
        : ['学生邮箱','中文姓名','学校英文名','学校中文名','类别','申请方式','状态','录入时间','录入人'];

      // 过滤掉该学生旧记录
      const otherRows = allRows.slice(1).filter(row => (row[0]||'').trim() !== studentEmail.trim());

      // 新记录
      const newRows = schools.map(s => [
        studentEmail.trim(),
        studentName || '',
        s.school || '',
        s.schoolCN || '',
        s.category || '',
        s.appType || '',
        '待确认',
        ts,
        operator || '顾问',
      ]);

      const finalRows = [headers, ...otherRows, ...newRows];

      // 写回 schools tab
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: 'schools!A1',
        valueInputOption: 'RAW',
        requestBody: { values: finalRows },
      });

      // 清掉多余的旧行（精确到 schools tab）
      const oldDataLen = allRows.length;
      const newDataLen = finalRows.length;
      if (oldDataLen > newDataLen) {
        await sheets.spreadsheets.values.clear({
          spreadsheetId,
          range: `schools!A${newDataLen + 1}:I${oldDataLen}`,
        });
      }

      await writeLog(sheets, spreadsheetId, {
        operator: operator || '顾问',
        operatorType: '顾问',
        studentEmail,
        action: '保存选校',
        detail: schools.map(s => `${s.schoolCN||s.school}(${s.category}/${s.appType})`).join('，'),
      });

      return res.json({ success: true });
    } catch (err) {
      return res.status(500).json({ success: false, error: err.message });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
}
