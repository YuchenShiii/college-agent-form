import { google } from 'googleapis';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: process.env.GOOGLE_CLIENT_EMAIL,
        private_key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const sheets = google.sheets({ version: 'v4', auth });
    const body = req.body;

    const row = [
      new Date().toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai' }),
      body.name || '',
      body.wechat || '',
      body.school || '',
      body.grade || '',
      body.enrollYear || '',
      body.gpa || '',
      body.gpaScale || '',
      body.sat || '',
      body.satMath || '',
      body.act || '',
      body.toefl || '',
      body.ap || '',
      body.major || '',
      body.region || '',
      body.budget || '',
      body.scholarship || '',
      body.tendency || '',
      body.activity1 || '',
      body.activity2 || '',
      body.activity3 || '',
      body.awards || '',
      body.proudest || '',
      body.whyMajor || '',
      body.unique || '',
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId: process.env.SHEET_ID,
      range: 'Sheet1!A:Y',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [row] },
    });

    res.status(200).json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
}
