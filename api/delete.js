import { google } from 'googleapis';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { rowIndex } = req.body;
    if (rowIndex === undefined || rowIndex === null) {
      return res.status(400).json({ success: false, error: '缺少 rowIndex' });
    }
    const auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: process.env.GOOGLE_CLIENT_EMAIL,
        private_key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const sheets = google.sheets({ version: 'v4', auth });
    const meta = await sheets.spreadsheets.get({ spreadsheetId: process.env.SHEET_ID });
    const sheet = meta.data.sheets.find(s => s.properties.title === 'Sheet1');
    const sheetId = sheet ? sheet.properties.sheetId : 0;
    const startRowIndex = parseInt(rowIndex) + 1; // +1 跳过表头
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: process.env.SHEET_ID,
      requestBody: {
        requests: [{ deleteDimension: { range: { sheetId, dimension: 'ROWS', startIndex: startRowIndex, endIndex: startRowIndex + 1 } } }]
      }
    });
    res.status(200).json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
}
