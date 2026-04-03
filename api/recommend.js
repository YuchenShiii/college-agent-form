import OpenAI from 'openai';
import {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, AlignmentType, WidthType, BorderStyle,
  ShadingType, TableLayoutType, HeadingLevel
} from 'docx';

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

// ── Prompt ──────────────────────────────────────────────────────────────────
function buildPrompt(s) {
  const v = k => (s[k] || '').trim();

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
    if (n) aps.push(`${n}${sc ? ` (${sc})` : ''}`);
  }
  const acts = [];
  for (let i = 1; i <= 10; i++) {
    const name = v(`活动${i}名称`), role = v(`活动${i}职位`), desc = v(`活动${i}描述`);
    if (name) acts.push(`- ${name}${role ? `（${role}）` : ''}${desc ? `：${desc}` : ''}`);
  }
  const honors = [];
  for (let i = 1; i <= 5; i++) {
    const n = v(`奖项${i}名称`), lv = v(`奖项${i}级别`);
    if (n) honors.push(`- ${n}${lv ? `（${lv}）` : ''}`);
  }

  return `你是一位顶尖的美本申请顾问，请根据以下学生档案生成完整的申请分析报告。

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

===== 请生成以下内容（JSON格式）=====

{
  "studentSummary": "2-3句学生整体评估，包括优势和薄弱点",
  "strategy": "申请策略3-5句，包括ED/EA建议、轮次规划",
  "reaches": [
    { "school": "英文名", "schoolCN": "中文名", "rank": "#XX", "reason": "推荐理由1-2句", "tips": "申请建议1句", "fitNote": "该校与学生的契合点" }
  ],
  "matches": [...同上，4-5所],
  "safeties": [...同上，3-4所],
  "edSuggestion": "ED首选英文名",
  "edSuggestionCN": "ED首选中文名",
  "edReason": "ED理由2-3句",
  "essayMainTheme": "主文书核心主题建议（2-3句，点出学生最有力的故事角度）",
  "essayIdeas": [
    { "angle": "故事角度标题", "description": "这个角度怎么写、为什么有力（2-3句）", "prompt": "给学生的引导问题（帮助学生想得更深）" }
  ],
  "supplementalHints": "补充文书方向建议（2-3句，针对Why School类文书）",
  "nextSteps": ["近期行动1", "近期行动2", "近期行动3", "近期行动4"]
}

冲刺2-3所，匹配4-5所，保底3-4所。essayIdeas提供3-4个不同角度。请基于学生真实数据给出专业判断，文书建议要具体有启发性。`;
}

// ── docx 工具函数 ────────────────────────────────────────────────────────────
function cell(text, opts = {}) {
  const { bold = false, color = '1a202c', bg = null, width = null, center = false, italic = false, size = 20 } = opts;
  return new TableCell({
    children: [new Paragraph({
      children: [new TextRun({ text: String(text || ''), bold, color, size, italics: italic })],
      alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
    })],
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    ...(bg ? { shading: { type: ShadingType.CLEAR, fill: bg } } : {}),
    ...(width ? { width: { size: width, type: WidthType.DXA } } : {}),
  });
}

function heading(text, color = '1251b5') {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 28, color })],
    spacing: { before: 400, after: 160 },
  });
}

function body(text) {
  return new Paragraph({
    children: [new TextRun({ text: String(text || ''), size: 22 })],
    spacing: { after: 160 },
  });
}

function divider(color = 'e2e8f0') {
  return new Paragraph({
    border: { bottom: { color, size: 4, value: BorderStyle.SINGLE } },
    spacing: { after: 320 },
    children: [],
  });
}

function schoolTable(schools, label, showFitNote = false) {
  if (!schools || !schools.length) return [];
  const colors = { '冲刺': '1251b5', '匹配': '0694a2', '保底': '059669' };
  const hBg = colors[label] || '1251b5';

  const headerCells = [
    cell('学校', { bold: true, bg: hBg, color: 'ffffff', center: true, width: 2000 }),
    cell('排名', { bold: true, bg: hBg, color: 'ffffff', center: true, width: 700 }),
    cell('推荐理由', { bold: true, bg: hBg, color: 'ffffff', width: 3200 }),
    cell('申请建议', { bold: true, bg: hBg, color: 'ffffff', width: 2100 }),
  ];
  if (showFitNote) headerCells.push(cell('契合点 / 顾问备注', { bold: true, bg: hBg, color: 'ffffff', width: 2000 }));

  return [
    new Paragraph({ children: [new TextRun({ text: `▌ ${label}院校`, bold: true, size: 26, color: hBg })], spacing: { before: 280, after: 100 } }),
    new Table({
      layout: TableLayoutType.FIXED,
      width: { size: showFitNote ? 10000 : 8000, type: WidthType.DXA },
      rows: [
        new TableRow({ children: headerCells, tableHeader: true }),
        ...schools.map((s, i) => {
          const rowBg = i % 2 === 0 ? 'f8fafc' : 'ffffff';
          const dataCells = [
            cell(`${s.schoolCN || ''}\n${s.school || ''}`, { bg: rowBg, width: 2000 }),
            cell(s.rank || '-', { center: true, bg: rowBg, width: 700 }),
            cell(s.reason || '', { bg: rowBg, width: 3200 }),
            cell(s.tips || '', { bg: rowBg, width: 2100 }),
          ];
          if (showFitNote) dataCells.push(cell(s.fitNote || '', { bg: rowBg, width: 2000, italic: true, color: '718096' }));
          return new TableRow({ children: dataCells });
        }),
      ],
    }),
    new Paragraph({ children: [], spacing: { after: 200 } }),
  ];
}

// ── 学生版 docx ──────────────────────────────────────────────────────────────
async function buildStudentDocx(s, rec) {
  const nameCN = (s['中文姓名'] || '').trim();
  const nameEN = (s['英文姓名'] || '').trim();
  const today = new Date().toLocaleDateString('zh-CN', { timeZone: 'Asia/Shanghai' });

  const doc = new Document({
    styles: { default: { document: { run: { font: 'Microsoft YaHei', size: 22 } } } },
    sections: [{
      properties: { page: { margin: { top: 1000, bottom: 1000, left: 1200, right: 1200 } } },
      children: [
        // 封面
        new Paragraph({ children: [new TextRun({ text: '🎓 美本申请定校建议', bold: true, size: 56, color: '1251b5' })], alignment: AlignmentType.CENTER, spacing: { after: 160 } }),
        new Paragraph({ children: [new TextRun({ text: `${nameCN}（${nameEN}）`, bold: true, size: 32, color: '1a202c' })], alignment: AlignmentType.CENTER, spacing: { after: 100 } }),
        new Paragraph({ children: [new TextRun({ text: `生成日期：${today}  ·  AdmitHere 升学顾问`, size: 20, color: '718096' })], alignment: AlignmentType.CENTER, spacing: { after: 500 } }),
        divider('1251b5'),

        // 综合评估
        heading('一、综合评估'),
        body(rec.studentSummary),

        // 申请策略
        heading('二、申请策略'),
        body(rec.strategy),
        new Paragraph({ children: [], spacing: { after: 100 } }),

        // 选校清单
        heading('三、选校清单'),
        ...schoolTable(rec.reaches, '冲刺'),
        ...schoolTable(rec.matches, '匹配'),
        ...schoolTable(rec.safeties, '保底'),

        // ED 建议
        heading('四、ED 首选建议'),
        new Paragraph({
          children: [
            new TextRun({ text: '推荐 ED 院校：', bold: true, size: 22 }),
            new TextRun({ text: `${rec.edSuggestionCN || ''}（${rec.edSuggestion || ''}）`, bold: true, size: 22, color: '1251b5' }),
          ],
          spacing: { after: 120 },
        }),
        body(rec.edReason),

        // 文书 Brainstorm
        heading('五、文书 Brainstorm'),
        new Paragraph({ children: [new TextRun({ text: '📌 主文书核心方向', bold: true, size: 24, color: '4a5568' })], spacing: { after: 120 } }),
        body(rec.essayMainTheme),
        new Paragraph({ children: [], spacing: { after: 160 } }),
        new Paragraph({ children: [new TextRun({ text: '💡 故事角度探索', bold: true, size: 24, color: '4a5568' })], spacing: { after: 120 } }),

        ...(rec.essayIdeas || []).flatMap((idea, i) => [
          new Paragraph({ children: [new TextRun({ text: `角度 ${i + 1}：${idea.angle || ''}`, bold: true, size: 22, color: '1251b5' })], spacing: { before: 240, after: 80 } }),
          body(idea.description),
          new Paragraph({
            children: [
              new TextRun({ text: '🤔 思考一下：', bold: true, size: 20, color: '0694a2' }),
              new TextRun({ text: `  ${idea.prompt || ''}`, size: 20, color: '4a5568', italics: true }),
            ],
            spacing: { after: 160 },
          }),
        ]),

        new Paragraph({ children: [new TextRun({ text: '📝 补充文书方向（Why School）', bold: true, size: 24, color: '4a5568' })], spacing: { before: 240, after: 120 } }),
        body(rec.supplementalHints),

        // 近期行动
        heading('六、近期行动清单'),
        ...(rec.nextSteps || []).map((step, i) =>
          new Paragraph({ children: [new TextRun({ text: `${i + 1}.  ${step}`, size: 22 })], spacing: { after: 100 } })
        ),

        // 页脚
        new Paragraph({ children: [], spacing: { before: 600 } }),
        divider(),
        new Paragraph({
          children: [new TextRun({ text: '本报告由 AdmitHere 顾问系统生成，仅供参考，请结合顾问意见使用。', size: 18, color: '718096', italics: true })],
          alignment: AlignmentType.CENTER,
        }),
      ],
    }],
  });

  return Packer.toBuffer(doc);
}

// ── 顾问版 docx ──────────────────────────────────────────────────────────────
async function buildAdvisorDocx(s, rec) {
  const nameCN = (s['中文姓名'] || '').trim();
  const nameEN = (s['英文姓名'] || '').trim();
  const today = new Date().toLocaleDateString('zh-CN', { timeZone: 'Asia/Shanghai' });

  const noteBox = (label) => new Table({
    layout: TableLayoutType.FIXED,
    width: { size: 9000, type: WidthType.DXA },
    rows: [new TableRow({ children: [
      cell(`📝 ${label}`, { bg: 'fef3c7', bold: true, color: '92400e', width: 2000 }),
      cell('', { bg: 'fffbeb', width: 7000 }),
    ]})],
  });

  const doc = new Document({
    styles: { default: { document: { run: { font: 'Microsoft YaHei', size: 22 } } } },
    sections: [{
      properties: { page: { margin: { top: 1000, bottom: 1000, left: 1200, right: 1200 } } },
      children: [
        // 封面
        new Paragraph({ children: [new TextRun({ text: '📋 定校建议 · 顾问工作版', bold: true, size: 52, color: '1251b5' })], alignment: AlignmentType.CENTER, spacing: { after: 160 } }),
        new Paragraph({ children: [new TextRun({ text: `${nameCN}（${nameEN}）`, bold: true, size: 32 })], alignment: AlignmentType.CENTER, spacing: { after: 100 } }),
        new Paragraph({ children: [new TextRun({ text: `生成日期：${today}  ·  内部使用，请勿直接转发学生`, size: 20, color: 'dc2626' })], alignment: AlignmentType.CENTER, spacing: { after: 500 } }),
        divider('1251b5'),

        // AI 评估
        heading('一、AI 综合评估'),
        body(rec.studentSummary),
        new Paragraph({ children: [], spacing: { after: 120 } }),
        noteBox('顾问补充评估'),
        new Paragraph({ children: [], spacing: { after: 300 } }),

        // 申请策略
        heading('二、申请策略'),
        body(rec.strategy),
        new Paragraph({ children: [], spacing: { after: 120 } }),
        noteBox('策略调整备注'),
        new Paragraph({ children: [], spacing: { after: 300 } }),

        // 选校清单（含契合点和顾问备注列）
        heading('三、选校清单审核'),
        new Paragraph({ children: [new TextRun({ text: '※ 灰斜体为 AI 生成的契合点/备注，请在此列填写确认意见或替换学校', size: 18, color: '718096', italics: true })], spacing: { after: 160 } }),
        ...schoolTable(rec.reaches, '冲刺', true),
        ...schoolTable(rec.matches, '匹配', true),
        ...schoolTable(rec.safeties, '保底', true),

        new Paragraph({ children: [new TextRun({ text: '最终确认选校名单', bold: true, size: 24, color: '1251b5' })], spacing: { before: 320, after: 120 } }),
        new Table({
          layout: TableLayoutType.FIXED,
          width: { size: 9000, type: WidthType.DXA },
          rows: [
            new TableRow({ children: [
              cell('类别', { bold: true, bg: '1251b5', color: 'ffffff', center: true, width: 1200 }),
              cell('学校名称', { bold: true, bg: '1251b5', color: 'ffffff', width: 3000 }),
              cell('ED/EA/RD', { bold: true, bg: '1251b5', color: 'ffffff', center: true, width: 1500 }),
              cell('备注', { bold: true, bg: '1251b5', color: 'ffffff', width: 3300 }),
            ], tableHeader: true }),
            ...['冲刺1', '冲刺2', '冲刺3', '匹配1', '匹配2', '匹配3', '匹配4', '匹配5', '保底1', '保底2', '保底3'].map((label, i) =>
              new TableRow({ children: [
                cell(label, { bg: i % 2 === 0 ? 'f8fafc' : 'ffffff', center: true, width: 1200 }),
                cell('', { bg: i % 2 === 0 ? 'f8fafc' : 'ffffff', width: 3000 }),
                cell('', { bg: i % 2 === 0 ? 'f8fafc' : 'ffffff', center: true, width: 1500 }),
                cell('', { bg: i % 2 === 0 ? 'f8fafc' : 'ffffff', width: 3300 }),
              ]})
            ),
          ],
        }),
        new Paragraph({ children: [], spacing: { after: 300 } }),

        // ED 建议
        heading('四、ED 建议'),
        new Paragraph({
          children: [
            new TextRun({ text: 'AI 推荐 ED：', bold: true, size: 22 }),
            new TextRun({ text: `${rec.edSuggestionCN || ''}（${rec.edSuggestion || ''}）`, bold: true, size: 22, color: '1251b5' }),
          ],
          spacing: { after: 120 },
        }),
        body(rec.edReason),
        new Paragraph({ children: [], spacing: { after: 120 } }),
        noteBox('顾问 ED 最终决定'),
        new Paragraph({ children: [], spacing: { after: 300 } }),

        // 文书 Brainstorm + 审核
        heading('五、文书 Brainstorm 审核'),
        new Paragraph({ children: [new TextRun({ text: '📌 主文书方向（AI 建议）', bold: true, size: 24, color: '4a5568' })], spacing: { after: 120 } }),
        body(rec.essayMainTheme),
        new Paragraph({ children: [], spacing: { after: 120 } }),
        noteBox('顾问对主文书方向的调整意见'),
        new Paragraph({ children: [], spacing: { after: 240 } }),

        new Paragraph({ children: [new TextRun({ text: '💡 故事角度（AI 建议）', bold: true, size: 24, color: '4a5568' })], spacing: { after: 120 } }),
        ...(rec.essayIdeas || []).flatMap((idea, i) => [
          new Paragraph({ children: [new TextRun({ text: `角度 ${i + 1}：${idea.angle || ''}`, bold: true, size: 22, color: '1251b5' })], spacing: { before: 240, after: 80 } }),
          body(idea.description),
          new Paragraph({
            children: [
              new TextRun({ text: '引导问题：', bold: true, size: 20, color: '0694a2' }),
              new TextRun({ text: `  ${idea.prompt || ''}`, size: 20, italics: true, color: '4a5568' }),
            ],
            spacing: { after: 100 },
          }),
          noteBox(`角度 ${i + 1} 顾问评语`),
          new Paragraph({ children: [], spacing: { after: 160 } }),
        ]),

        new Paragraph({ children: [new TextRun({ text: '📝 补充文书方向', bold: true, size: 24, color: '4a5568' })], spacing: { before: 200, after: 120 } }),
        body(rec.supplementalHints),
        new Paragraph({ children: [], spacing: { after: 120 } }),
        noteBox('顾问补充文书备注'),
        new Paragraph({ children: [], spacing: { after: 300 } }),

        // 近期行动
        heading('六、近期行动清单'),
        ...(rec.nextSteps || []).map((step, i) =>
          new Paragraph({ children: [new TextRun({ text: `${i + 1}.  ${step}`, size: 22 })], spacing: { after: 100 } })
        ),
        new Paragraph({ children: [], spacing: { after: 120 } }),
        noteBox('顾问补充行动项'),
        new Paragraph({ children: [], spacing: { after: 400 } }),

        // 页脚
        divider(),
        new Paragraph({
          children: [new TextRun({ text: `内部工作文件  ·  ${nameCN}  ·  AdmitHere  ·  ${today}`, size: 18, color: '718096', italics: true })],
          alignment: AlignmentType.CENTER,
        }),
      ],
    }],
  });

  return Packer.toBuffer(doc);
}

// ── Handler ──────────────────────────────────────────────────────────────────
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

    // 调用 GPT
    const completion = await openai.chat.completions.create({
      model: 'gpt-4o',
      messages: [{ role: 'user', content: buildPrompt(student) }],
      response_format: { type: 'json_object' },
      temperature: 0.7,
    });
    const rec = JSON.parse(completion.choices[0].message.content);

    // 生成两个 docx
    const [studentBuf, advisorBuf] = await Promise.all([
      buildStudentDocx(student, rec),
      buildAdvisorDocx(student, rec),
    ]);

    const nameCN = (student['中文姓名'] || 'student').trim();

    // 返回 JSON 包含两个 base64 文件
    res.status(200).json({
      success: true,
      studentDoc: Buffer.from(studentBuf).toString('base64'),
      advisorDoc: Buffer.from(advisorBuf).toString('base64'),
      filename: nameCN,
    });

  } catch (err) {
    console.error('recommend error:', err);
    res.status(500).json({ success: false, error: err.message });
  }
}
