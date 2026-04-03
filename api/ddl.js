import OpenAI from 'openai';

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

function buildPrompt(schools, studentName) {
  const list = schools.map(s => `${s.schoolCN || s.school}（${s.school}）- ${s.appType || 'RD'}`).join('\n');
  return `你是美本申请顾问，学生 ${studentName} 的最终选校清单如下：

${list}

请为每所学校生成详细的 DDL 时间线（2025-2026 申请季），JSON 格式返回：

{
  "schools": [
    {
      "school": "学校英文名",
      "schoolCN": "学校中文名",
      "appType": "ED/EA/RD",
      "deadlines": [
        { "type": "申请截止", "date": "YYYY-MM-DD", "note": "备注（可选）" },
        { "type": "文书准备建议", "date": "YYYY-MM-DD", "note": "提前至少2周完成" },
        { "type": "推荐信请求", "date": "YYYY-MM-DD", "note": "提前1个月联系老师" },
        { "type": "CSS Profile", "date": "YYYY-MM-DD", "note": "（如需财援）" },
        { "type": "FAFSA", "date": "YYYY-MM-DD", "note": "（如需财援）" },
        { "type": "成绩单提交", "date": "YYYY-MM-DD" }
      ]
    }
  ],
  "summary": "整体时间规划建议（2-3句，提醒关键节点）"
}

请基于真实的申请截止日期（ED 通常 11/1 或 11/15，EA 11/1，RD 1/1 或 1/15），倒推准备时间。文书建议提前2-3周，推荐信提前1个月。`;
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { schools, studentName } = req.body;
    if (!schools || !schools.length) {
      return res.status(400).json({ success: false, error: '缺少选校数据' });
    }

    const completion = await openai.chat.completions.create({
      model: 'gpt-4o',
      messages: [{ role: 'user', content: buildPrompt(schools, studentName || '学生') }],
      response_format: { type: 'json_object' },
      temperature: 0.3,
    });

    const result = JSON.parse(completion.choices[0].message.content);
    return res.json({ success: true, data: result });

  } catch (err) {
    console.error('ddl error:', err);
    return res.status(500).json({ success: false, error: err.message });
  }
}
