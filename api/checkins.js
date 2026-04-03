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
      const resp = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'checkins!A:H' });
      const rows = resp.data.values || [];
      if (rows.length <= 1) return res.json({ success: true, checkins: [] });
      const headers = rows[0];
      const data = rows.slice(1)
        .map(row => { const obj = {}; headers.forEach((h,j) => obj[h]=row[j]||''); return obj; })
        .filter(r => !email || r['学生邮箱'] === email);
      return res.json({ success: true, checkins: data });
    } catch (err) {
      return res.status(500).json({ success: false, error: err.message });
    }
  }

  if (req.method === 'POST') {
    try {
      const { studentEmail, school, ddlType, ddlDate, checked, operator, operatorType } = req.body;
      if (!studentEmail || !ddlType) return res.status(400).json({ success: false, error: '缺少必要字段' });

      const ts = new Date().toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai' });
      const status = checked ? '已完成' : '未完成';

      const existing = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'checkins!A:H' });
      const rows = existing.data.values || [];
      const headers = rows[0] || [];
      const emailIdx = headers.indexOf('学生邮箱');
      const schoolIdx = headers.indexOf('学校英文名');
      const typeIdx = headers.indexOf('DDL类型');
      const dateIdx = headers.indexOf('DDL日期');

      let existingRowIdx = -1;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][emailIdx]===studentEmail && rows[i][schoolIdx]===school &&
            rows[i][typeIdx]===ddlType && rows[i][dateIdx]===ddlDate) {
          existingRowIdx = i + 1; break;
        }
      }

      if (existingRowIdx > 0) {
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `checkins!E${existingRowIdx}:H${existingRowIdx}`,
          valueInputOption: 'RAW',
          requestBody: { values: [[status, operator||'顾问', operatorType||'顾问', ts]] },
        });
      } else {
        await sheets.spreadsheets.values.append({
          spreadsheetId, range: 'checkins!A:H', valueInputOption: 'RAW',
          requestBody: { values: [[studentEmail, school||'', ddlType, ddlDate||'', status, operator||'顾问', operatorType||'顾问', ts]] },
        });
      }

      await sheets.spreadsheets.values.append({
        spreadsheetId, range: 'logs!A:F', valueInputOption: 'RAW',
        requestBody: { values: [[ts, operator||'顾问', operatorType||'顾问', studentEmail,
          checked?'勾选完成':'取消勾选', `${school} · ${ddlType} · ${ddlDate}`]] },
      });

      return res.json({ success: true });
    } catch (err) {
      return res.status(500).json({ success: false, error: err.message });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
}
