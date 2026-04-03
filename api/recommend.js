import OpenAI from 'openai';
import {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, HeadingLevel, AlignmentType, WidthType, BorderStyle,
  ShadingType, TableLayoutType
} from 'docx';

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

function buildPrompt(s) {
  const v = (k) => (s[k] || '').trim();
  const scores = [
    v('GPA') ? `GPA: ${v('GPA')}/${v('GPA满分制') || '4.0'}` : null,
    v('SAT总分') ? `SAT: ${v('SAT总分')}（数学${v('SAT数学')} / EBRW ${v('SAT EBRW')}）` : null,
    v('ACT总分') ? `ACT: ${v('ACT总分')}` : null,
    v('托福总分') ? `托福: ${v('托福总分')}` : null,
    v('雅思') ? `雅思: ${v('雅思')}` : null,
    v('是否再考SAT') ? `SAT再考计划: ${v('是否再考SAT')}` : null,
  ].filter(Boolean).join('\n');

  const aps = [];
  for (let i = 1; i <= 8; i++) {
    const n = v(`AP${i}课程`), sc = v(`AP${i}分数`);
    if (n) aps.push(`${n}${sc ? ' (' + sc + ')' : ''}`);
  }
  const acts = [];
  for (let i = 1; i <= 10; i++) {
    const name = v(`活动${i}名称`), role = v(`活动${i}职位`), desc = v(`活动${i}描述`);
    if (name) acts.push(`- ${name}${role ? '（' + role + '）' : ''}${desc ? '：' + desc : ''}`);
  }
  const honors = [];
  for (let i = 1; i <= 5; i++) {
    const n = v(`奖项${i}名称`), lv = v(`奖项${i}级别`);
    if (n) honors.push(`- ${n}${lv ? '（' + lv + '）' : ''}`);
  }

  return `你是一位专业的美本申请顾问，请根据以下学生档案，生成一份专业的定校建议报告。

===== 学生档案 =====
姓名：${v('中文姓名')}（${v('英文姓名')}）
学校：${v('高中名称')}（${v('高中城市')}，${v('高中类型')}）
年级：${v('年级')}，毕业年份：${v('毕业年份')}
课程体系：${v('课程体系')}

【成绩】
${scores || '未填写'}

【AP课程】
${aps.length ? aps.join('\n') : '未填写'}

【课外活动】
${acts.length ? acts.join('\n') : '未填写'}

【荣誉奖项】
${honors.length ? honors.join('\n') : '无'}

【申请意向】
意向专业：${[v('意向专业1'), v('意向专业2'), v('意向专业3')].filter(Boolean).join(' / ') || '未填写'}
申请轮次：${v('申请轮次') || '未指定'}
录取倾向：${v('录取倾向') || '未指定'}
入学学期：${v('入学学期') || '未指定'}
地区偏好：${v('地区偏好') || '不限'}
学校规模：${v('学校规模') || '不限'}
年费预算：${v('年费预算') || '未指定'}
财援需求：${v('财援需求') || '无'}
意向学校：${v('意向学校') || '无'}

【文书素材线索】
最骄傲的事：${v('最骄傲的事') || '未填写'}
改变你的经历：${v('改变你的经历') || '未填写'}
为什么选这个专业：${v('为什么选这个专业') || '未填写'}
与众不同之处：${v('与众不同之处') || '未填写'}

===== 请生成以下内容 =====
请严格按照以下结构输出，使用 JSON 格式返回：

{
  "studentSummary": "2-3句话的学生整体评估",
  "strategy": "申请策略说明（3-5句话）",
  "reaches": [
    { "school": "学校英文名", "schoolCN": "学校中文名", "rank": "USNews排名如#5", "reason": "推荐理由1-2句", "tips": "申请建议1句" }
  ],
  "matches": [...同上格式，4-5所],
  "safeties": [...同上格式，3-4所],
  "edSuggestion": "ED首选学校英文名",
  "edSuggestionCN": "ED首选学校中文名",
  "edReason": "ED建议理由2-3句",
  "essayHints": "文书主题建议3-5句",
  "nextSteps": ["行动1", "行动2", "行动3"]
}

冲刺校2-3所，匹配校4-5所，保底校3-4所。请基于学生真实数据给出专业判断。`;
}

async function callGPT(prompt) {
  const completion = await openai.chat.completions.create({
    model: 'gpt-4o',
    messages: [{ role: 'user', content: prompt }],
    response_format: { type: 'json_object' },
    temperature: 0.7,
  });
  return JSON.parse(completion.choices[0].message.content);
}

function makeCell(text, options = {}) {
  const { bold = false, color = '000000', bg = null, width = null, center = false } = options;
  return new TableCell({
    children: [new Paragraph({
      children: [new TextRun({ text: String(text), bold, color, size: 20 })],
      alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
    })],
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    ...(bg ? { shading: { type: ShadingType.CLEAR, fill: bg } } : {}),
    ...(width ? { width: { size: width, type: WidthType.DXA } } : {}),
  });
}

function schoolTable(schools, label) {
  if (!schools || schools.length === 0) return [];
  const headerBg = label === '冲刺' ? '1251b5' : label === '匹配' ? '0694a2' : '059669';
  return [
    new Paragraph({ text: '', spacing: { before: 300 } }),
    new Paragraph({
      children: [new TextRun({ text: `▌ ${label}院校`, bold: true, size: 26, color: headerBg })],
    }),
    new Paragraph({ text: '', spacing: { before: 100 } }),
    new Table({
      layout: TableLayoutType.FIXED,
      width: { size: 9000, type: WidthType.DXA },
      rows: [
        new TableRow({
          children: [
            makeCell('学校', { bold: true, bg: headerBg, color: 'ffffff', center: true, width: 2200 }),
            makeCell('排名', { bold: true, bg: headerBg, color: 'ffffff', center: true, width: 800 }),
            makeCell('推荐理由', { bold: true, bg: headerBg, color: 'ffffff', width: 3500 }),
            makeCell('申请建议', { bold: true, bg: headerBg, color: 'ffffff', width: 2500 }),
          ],
          tableHeader: true,
        }),
        ...schools.map((s, i) => new TableRow({
          children: [
            makeCell(`${s.schoolCN || s.school}\n${s.school}`, { bg: i % 2 === 0 ? 'f8fafc' : 'ffffff', width: 2200 }),
            makeCell(s.rank || '-', { center: true, bg: i % 2 === 0 ? 'f8fafc' : 'ffffff', width: 800 }),
            makeCell(s.reason || '', { bg: i % 2 === 0 ? 'f8fafc' : 'ffffff', width: 3500 }),
            makeCell(s.tips || '', { bg: i % 2 === 0 ? 'f8fafc' : 'ffffff', width: 2500 }),
          ],
        })),
      ],
    }),
  ];
}

async function buildDocx(s, rec) {
  const nameCN = (s['中文姓名'] || '').trim();
  const nameEN = (s['英文姓名'] || '').trim();
  const today = new Date().toLocaleDateString('zh-CN', { timeZone: 'Asia/Shanghai' });

  const doc = new Document({
    styles: {
      default: { document: { run: { font: 'Microsoft YaHei', size: 22 } } },
    },
    sections: [{
      properties: { page: { margin: { top: 1000, bottom: 1000, left: 1200, right: 1200 } } },
      children: [
        new Paragraph({
          children: [new TextRun({ text: '美本申请定校建议报告', bold: true, size: 52, color: '1251b5' })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [new TextRun({ text: `${nameCN}（${nameEN}）· ${today}`, size: 22, color: '718096' })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
        }),
        new Paragraph({
          border: { bottom: { color: '1251b5', size: 6, value: BorderStyle.SINGLE } },
          spacing: { after: 400 },
          children: [],
        }),
        new Paragraph({
          children: [new TextRun({ text: '一、学生综合评估', bold: true, size: 28, color: '1251b5' })],
          spacing: { before: 300, after: 150 },
        }),
        new Paragraph({
          children: [new TextRun({ text: rec.studentSummary || '', size: 22 })],
          spacing: { after: 300 },
        }),
        new Paragraph({
          children: [new TextRun({ text: '二、申请策略规划', bold: true, size: 28, color: '1251b5' })],
          spacing: { before: 300, after: 150 },
        }),
        new Paragraph({
          children: [new TextRun({ text: rec.strategy || '', size: 22 })],
          spacing: { after: 300 },
        }),
        new Paragraph({
          children: [new TextRun({ text: '三、选校清单', bold: true, size: 28, color: '1251b5' })],
          spacing: { before: 300, after: 150 },
        }),
        ...schoolTable(rec.reaches, '冲刺'),
        ...schoolTable(rec.matches, '匹配'),
        ...schoolTable(rec.safeties, '保底'),
        new Paragraph({
          children: [new TextRun({ text: '四、ED 首选建议', bold: true, size: 28, color: '1251b5' })],
          spacing: { before: 500, after: 150 },
        }),
        new Paragraph({
          children: [
            new TextRun({ text: '推荐 ED 院校：', bold: true, size: 22 }),
            new TextRun({ text: `${rec.edSuggestionCN || ''}（${rec.edSuggestion || ''}）`, size: 22, color: '1251b5', bold: true }),
          ],
          spacing: { after: 150 },
        }),
        new Paragraph({
          children: [new TextRun({ text: rec.edReason || '', size: 22 })],
          spacing: { after: 300 },
        }),
        new Paragraph({
          children: [new TextRun({ text: '五、文书主题建议', bold: true, size: 28, color: '1251b5' })],
          spacing: { before: 300, after: 150 },
        }),
        new Paragraph({
          children: [new TextRun({ text: rec.essayHints || '', size: 22 })],
          spacing: { after: 300 },
        }),
        new Paragraph({
          children: [new TextRun({ text: '六、近期行动清单', bold: true, size: 28, color: '1251b5' })],
          spacing: { before: 300, after: 150 },
        }),
        ...(rec.nextSteps || []).map((step, i) =>
          new Paragraph({
            children: [new TextRun({ text: `${i + 1}. ${step}`, size: 22 })],
            spacing: { after: 100 },
          })
        ),
        new Paragraph({ children: [], spacing: { before: 600 } }),
        new Paragraph({
          border: { top: { color: 'e2e8f0', size: 4, value: BorderStyle.SINGLE } },
          children: [new TextRun({ text: '本报告由 AdmitHere 顾问系统生成，仅供参考，请结合顾问意见使用。', size: 18, color: '718096', italics: true })],
          alignment: AlignmentType.CENTER,
          spacing: { before: 200 },
        }),
      ],
    }],
  });

  return await Packer.toBuffer(doc);
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const student = req.body;
    if (!student || !student['中文姓名']) {
      return res.status(400).json({ success: false, error: '缺少学生数据' });
    }
    const prompt = buildPrompt(student);
    const rec = await callGPT(prompt);
    const buffer = await buildDocx(student, rec);
    const nameCN = (student['中文姓名'] || 'student').trim();
    const filename = encodeURIComponent(`${nameCN}_定校建议.docx`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
}
