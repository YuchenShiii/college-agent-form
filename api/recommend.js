import OpenAI from 'openai';

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const student = req.body;
    const nameCN = (student['中文姓名'] || '').trim();
    if (!nameCN) return res.status(400).json({ success: false, error: '缺少学生姓名' });

    const completion = await openai.chat.completions.create({
      model: 'gpt-4o',
      messages: [{ role: 'user', content: `请用一句话介绍美本申请顾问的工作` }],
      max_tokens: 100,
    });

    return res.status(200).json({
      success: true,
      test: completion.choices[0].message.content,
    });

  } catch (err) {
    console.error('recommend error:', err);
    return res.status(500).json({ success: false, error: err.message });
  }
}
