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
    const b = req.body;

    const row = [
      new Date().toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai' }),
      b.nameCN, b.nameEN, b.dob, b.gender, b.nationality, b.passport, b.usStatus, b.nativeLang,
      b.wechat, b.phone, b.email, b.addr1, b.city, b.state, b.zip, b.country,
      b.school, b.schoolType, b.schoolCity, b.grade, b.gradYear,
      b.gpa, b.gpaScale, b.classRank, b.curriculum,
      ...[1,2,3,4,5,6,7,8].flatMap(n => [b[`ap${n}_name`], b[`ap${n}_score`]]),
      ...[1,2,3,4,5,6].flatMap(n => [b[`ib${n}_level`], b[`ib${n}_name`], b[`ib${n}_score`]]),
      b.sat, b.satMath, b.satEbrw, b.satDate, b.satRetake,
      b.act, b.actSub,
      b.toefl, b.toeflSub, b.ielts, b.duolingo, b.otherTests,
      ...[1,2,3,4,5,6,7,8,9,10].flatMap(n => [
        b[`act${n}_type`], b[`act${n}_role`], b[`act${n}_name`],
        b[`act${n}_desc`], b[`act${n}_hrs`], b[`act${n}_period`],
      ]),
      ...[1,2,3,4,5].flatMap(n => [b[`honor${n}_name`], b[`honor${n}_level`], b[`honor${n}_grade`]]),
      b.major1, b.major2, b.major3, b.enrollTerm, b.appRound, b.region, b.schoolSize,
      b.budget, b.aidNeeded, b.tendency, b.targetSchools, b.appNotes,
      b.fatherName, b.fatherNameEN, b.fatherJob, b.fatherCompany, b.fatherEdu, b.fatherPhone,
      b.motherName, b.motherNameEN, b.motherJob, b.motherCompany, b.motherEdu, b.motherPhone,
      b.siblings, b.familyAlumni, b.legacyDetail, b.familyIncome,
      b.proudest, b.lifeEvent, b.whyMajor, b.unique, b.recommenders, b.otherInfo,
    ].map(v => v || '');

    await sheets.spreadsheets.values.append({
      spreadsheetId: process.env.SHEET_ID,
      range: 'Sheet1!A1',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [row] },
    });

    res.status(200).json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
}
