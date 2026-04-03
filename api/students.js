import { google } from 'googleapis';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: process.env.GOOGLE_CLIENT_EMAIL,
        private_key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    });

    const sheets = google.sheets({ version: 'v4', auth });
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.SHEET_ID,
      range: 'Sheet1',
    });

    const rows = response.data.values || [];
    if (rows.length === 0) {
      return res.status(200).json({ success: true, count: 0, students: [] });
    }

    const headers = rows[0];
    const students = rows.slice(1)
      .filter(row => row.length > 1 && (row[1] || '').trim())
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[h.trim()] = (row[i] || '').trim(); });
        return obj;
      });

    res.status(200).json({ success: true, count: students.length, students });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
}
