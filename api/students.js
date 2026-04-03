import { google } from 'googleapis';

const HEADERS = [
  '时间戳','中文姓名','英文姓名','出生日期','性别','国籍','护照号码','美国身份','母语',
  '微信号','手机号','邮箱','街道地址','城市','省份','邮编','国家',
  '高中名称','高中类型','高中城市','年级','毕业年份',
  'GPA','GPA满分制','Class Rank','课程体系',
  'AP1课程','AP1分数','AP2课程','AP2分数','AP3课程','AP3分数','AP4课程','AP4分数',
  'AP5课程','AP5分数','AP6课程','AP6分数','AP7课程','AP7分数','AP8课程','AP8分数',
  'IB1级别','IB1课程','IB1分数','IB2级别','IB2课程','IB2分数',
  'IB3级别','IB3课程','IB3分数','IB4级别','IB4课程','IB4分数',
  'IB5级别','IB5课程','IB5分数','IB6级别','IB6课程','IB6分数',
  'SAT总分','SAT数学','SAT EBRW','SAT考试日期','是否再考SAT',
  'ACT总分','ACT各科','托福总分','托福各项','雅思','DuoLingo','其他标化',
  '活动1类别','活动1职位','活动1名称','活动1描述','活动1每周时长','活动1时间段',
  '活动2类别','活动2职位','活动2名称','活动2描述','活动2每周时长','活动2时间段',
  '活动3类别','活动3职位','活动3名称','活动3描述','活动3每周时长','活动3时间段',
  '活动4类别','活动4职位','活动4名称','活动4描述','活动4每周时长','活动4时间段',
  '活动5类别','活动5职位','活动5名称','活动5描述','活动5每周时长','活动5时间段',
  '活动6类别','活动6职位','活动6名称','活动6描述','活动6每周时长','活动6时间段',
  '活动7类别','活动7职位','活动7名称','活动7描述','活动7每周时长','活动7时间段',
  '活动8类别','活动8职位','活动8名称','活动8描述','活动8每周时长','活动8时间段',
  '活动9类别','活动9职位','活动9名称','活动9描述','活动9每周时长','活动9时间段',
  '活动10类别','活动10职位','活动10名称','活动10描述','活动10每周时长','活动10时间段',
  '奖项1名称','奖项1级别','奖项1年级','奖项2名称','奖项2级别','奖项2年级',
  '奖项3名称','奖项3级别','奖项3年级','奖项4名称','奖项4级别','奖项4年级',
  '奖项5名称','奖项5级别','奖项5年级',
  '意向专业1','意向专业2','意向专业3','入学学期','申请轮次','地区偏好','学校规模',
  '年费预算','财援需求','录取倾向','意向学校','申请备注',
  '父亲姓名','父亲英文名','父亲职位','父亲公司','父亲学历','父亲手机',
  '母亲姓名','母亲英文名','母亲职位','母亲公司','母亲学历','母亲手机',
  '兄弟姐妹','家庭Legacy','Legacy详情','家庭年收入',
  '最骄傲的事','改变你的经历','为什么选这个专业','与众不同之处','推荐信老师','其他备注',
];

function rowToObj(row) {
  const obj = {};
  HEADERS.forEach((h, i) => { obj[h] = row[i] || ''; });
  return obj;
}

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
      range: 'Sheet1!A2:Z',
    });

    const rows = response.data.values || [];
    const students = rows
      .filter(row => row.length > 1 && row[1])
      .map(rowToObj);

    res.status(200).json({ success: true, count: students.length, students });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
}
