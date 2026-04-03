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
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: 'logs!A:F',
    valueInputOption: 'RAW',
    requestBody: { values: [[ts, operator, operatorType, studentEmail, action, detail]] },
  });
}

async function getSheetId(sheets, spreadsheetId, sheetName) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = meta.data.sheets.find(s => s.properties.title === sheetName);
  return sheet ? sheet.properties.sheetId : 0;
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  const spreadsheetId = process.env.SHEET_ID;

  // GET — 读取某学生的选校列表
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
          headers.forEach((h, j) => obj[h] = row[j] || '');
          return obj;
        })
        .filter(r => !email || r['学生邮箱'] === email);
      return res.json({ success: true, schools: data });
    } catch (err) {
      return res.status(500).json({ success: false, error: err.message });
    }
  }

  // POST — 录入选校（先清空该学生旧记录，再写新记录）
  if (req.method === 'POST') {
    try {
      const { studentEmail, studentName, schools, operator } = req.body;
      if (!studentEmail || !schools || !schools.length) {
        return res.status(400).json({ success: false, error: '缺少必要字段' });
      }

      const sheetId = await getSheetId(sheets, spreadsheetId, 'schools');

      // 读取现有数据，找出该学生所有行（倒序，从后往前删，避免行号偏移）
      const existing = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: 'schools!A:A',
      });
      const existRows = existing.data.values || [];
      // 收集需要删除的行索引（0-indexed），跳过表头(index 0)，倒序排列
      const toDelete = existRows
        .map((r, i) => (i > 0 && r[0] === studentEmail) ? i : null)
        .filter(i => i !== null)
        .sort((a, b) => b - a); // 倒序，从最后一行开始删

      // 逐行删除（倒序保证行号不偏移）
      for (const rowIdx of toDelete) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: {
            requests: [{
              deleteDimension: {
                range: {
                  sheetId,
                  dimension: 'ROWS',
                  startIndex: rowIdx,
                  endIndex: rowIdx + 1,
                },
              },
            }],
          },
        });
      }

      // 写入新选校
      const ts = new Date().toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai' });
      const rows = schools.map(s => [
        studentEmail,
        studentName || '',
        s.school || '',
        s.schoolCN || '',
        s.category || '',
        s.appType || '',
        '待确认',
        ts,
        operator || '顾问',
      ]);
      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: 'schools!A:I',
        valueInputOption: 'RAW',
        requestBody: { values: rows },
      });

      // 写日志
      await writeLog(sheets, spreadsheetId, {
        operator: operator || '顾问',
        operatorType: '顾问',
        studentEmail,
        action: '录入选校',
        detail: schools.map(s => `${s.schoolCN||s.school}(${s.category}/${s.appType})`).join('，'),
      });

      return res.json({ success: true });
    } catch (err) {
      return res.status(500).json({ success: false, error: err.message });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
}
