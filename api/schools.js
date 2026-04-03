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
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  const spreadsheetId = process.env.SHEET_ID;

  if (req.method === 'GET') {
    try {
      const { email } = req.query;
      const resp = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'schools!A:I' });
      const rows = resp.data.values || [];
      if (rows.length <= 1) return res.json({ success: true, schools: [] });
      const headers = rows[0];
      const data = rows.slice(1)
        .map((row, i) => {
          const obj = {};
          headers.forEach((h, j) => obj[h] = row[j] || '');
          obj._rowIndex = i + 2;
          return obj;
        })
        .filter(r => !email || r['学生邮箱'] === email);
      return res.json({ success: true, schools: data });
    } catch (err) {
      return res.status(500).json({ success: false, error: err.message });
    }
  }

  if (req.method === 'POST') {
    try {
      const { studentEmail, studentName, schools, operator } = req.body;
      if (!studentEmail || !schools || !schools.length) {
        return res.status(400).json({ success: false, error: '缺少必要字段' });
      }

      const existing = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'schools!A:A' });
      const existRows = existing.data.values || [];
      const deleteIndexes = existRows.map((r, i) => r[0] === studentEmail ? i + 1 : null).filter(Boolean).reverse();
      for (const rowIdx of deleteIndexes) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: {
            requests: [{
              deleteDimension: {
                range: { sheetId: await getSheetId(sheets, spreadsheetId, 'schools'), dimension: 'ROWS', startIndex: rowIdx - 1, endIndex: rowIdx }
              }
            }]
          }
        });
      }

      const ts = new Date().toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai' });
      const rows = schools.map(s => [
        studentEmail, studentName || '', s.school || '', s.schoolCN || '',
        s.category || '', s.appType || '', '待确认', ts, operator || '顾问'
      ]);
      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: 'schools!A:I',
        valueInputOption: 'RAW',
        requestBody: { values: rows },
      });

      await writeLog(sheets, spreadsheetId, {
        operator: operator || '顾问',
        operatorType: '顾问',
        studentEmail,
        action: '录入选校',
        detail: schools.map(s => `${s.schoolCN}(${s.category}/${s.appType})`).join('，'),
      });

      return res.json({ success: true });
    } catch (err) {
      return res.status(500).json({ success: false, error: err.message });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
}
