import OpenAI from 'openai';

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

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
  for (let i = 1; i <= 8; i++) { const n = v(`AP${i}课程`), sc = v(`AP${i}分数`); if (n) aps.push(`${n}${sc ? ` (${sc})` : ''}`); }
  const acts = [];
  for (let i = 1; i <= 10; i++) { const name = v(`活动${i}名称`), role = v(`活动${i}职位`), desc = v(`活动${i}描述`); if (name) acts.push(`- ${name}${role ? `（${role}）` : ''}${desc ? `：${desc}` : ''}`); }
  const honors = [];
  for (let i = 1; i <= 5; i++) { const n = v(`奖项${i}名称`), lv = v(`奖项${i}级别`); if (n) honors.push(`- ${n}${lv ? `（${lv}）` : ''}`); }

  return `你是一位顶尖的美本申请顾问，请根据以下学生档案生成完整的申请分析报告。

===== 学生档案 =====
姓名：${v('中文姓名')}（${v('英文姓名')}）
学校：${v('高中名称')}（${v('高中城市')}，${v('高中类型')}）
年级：${v('年级')}，毕业年份：${v('毕业年份')}

【成绩】${scores || '未填写'}
【AP课程】${aps.join(', ') || '未填写'}
【活动】${acts.join(' | ') || '未填写'}
【奖项】${honors.join(' | ') || '无'}
【意向专业】${[v('意向专业1'), v('意向专业2'), v('意向专业3')].filter(Boolean).join(' / ') || '未填写'}
【申请偏好】轮次:${v('申请轮次')} 倾向:${v('录取倾向')} 地区:${v('地区偏好')} 规模:${v('学校规模')} 预算:${v('年费预算')} 财援:${v('财援需求')}
【文书素材】最骄傲:${v('最骄傲的事')} | 改变经历:${v('改变你的经历')} | 为何选专业:${v('为什么选这个专业')} | 与众不同:${v('与众不同之处')}

===== 请生成JSON =====
{
  "studentSummary": "2-3句整体评估",
  "strategy": "3-5句申请策略",
  "reaches": [{"school":"英文名","schoolCN":"中文名","rank":"#XX","reason":"理由","tips":"建议","fitNote":"契合点"}],
  "matches": [...同上4-5所],
  "safeties": [...同上3-4所],
  "edSuggestion": "ED英文名",
  "edSuggestionCN": "ED中文名",
  "edReason": "ED理由2-3句",
  "essayMainTheme": "主文书方向2-3句",
  "essayIdeas": [{"angle":"角度标题","description":"描述2-3句","prompt":"引导问题"}],
  "supplementalHints": "补充文书建议2-3句",
  "nextSteps": ["行动1","行动2","行动3","行动4"]
}
冲刺2-3所，匹配4-5所，保底3-4所，essayIdeas给3-4个角度。`;
}

function buildStudentHTML(s, rec) {
  const nameCN = (s['中文姓名'] || '').trim();
  const nameEN = (s['英文姓名'] || '').trim();
  const today = new Date().toLocaleDateString('zh-CN', { timeZone: 'Asia/Shanghai' });
  const colors = { '冲刺': '#1251b5', '匹配': '#0694a2', '保底': '#059669' };

  const schoolSection = (schools, label) => {
    if (!schools || !schools.length) return '';
    const bg = colors[label] || '#1251b5';
    return `
    <h3 style="color:${bg};margin-top:20px">▌ ${label}院校</h3>
    <table>
      <tr style="background:${bg};color:#fff">
        <th style="width:20%">学校</th><th style="width:8%;text-align:center">排名</th>
        <th style="width:37%">推荐理由</th><th style="width:35%">申请建议</th>
      </tr>
      ${schools.map((sc, i) => `
      <tr style="background:${i % 2 === 0 ? '#f8fafc' : '#fff'}">
        <td><strong>${sc.schoolCN || ''}</strong><br><small style="color:#718096">${sc.school || ''}</small></td>
        <td style="text-align:center">${sc.rank || '-'}</td>
        <td>${sc.reason || ''}</td>
        <td>${sc.tips || ''}</td>
      </tr>`).join('')}
    </table>`;
  };

  return `<html><head><meta charset="UTF-8">
  <style>
    body{font-family:"Microsoft YaHei",Arial,sans-serif;font-size:12pt;color:#1a202c;margin:40px}
    h1{color:#1251b5;font-size:22pt;text-align:center;margin-bottom:4px}
    h2{color:#1251b5;font-size:14pt;border-left:4px solid #1251b5;padding-left:10px;margin-top:28px}
    p{line-height:1.8;margin:6px 0}
    table{border-collapse:collapse;width:100%;margin:8px 0;font-size:11pt}
    th,td{border:1px solid #e2e8f0;padding:8px 10px;vertical-align:top}
    .subtitle{text-align:center;color:#718096;font-size:11pt;margin-bottom:30px}
    .idea-box{background:#f0f4f9;border-left:3px solid #1251b5;padding:10px 14px;margin:10px 0;border-radius:0 6px 6px 0}
    .footer{text-align:center;color:#718096;font-size:10pt;margin-top:40px;border-top:1px solid #e2e8f0;padding-top:12px}
  </style></head><body>
  <h1>🎓 美本申请定校建议</h1>
  <div class="subtitle">${nameCN}（${nameEN}）· ${today} · AdmitHere 升学顾问</div>

  <h2>一、综合评估</h2><p>${rec.studentSummary || ''}</p>

  <h2>二、申请策略</h2><p>${rec.strategy || ''}</p>

  <h2>三、选校清单</h2>
  ${schoolSection(rec.reaches, '冲刺')}
  ${schoolSection(rec.matches, '匹配')}
  ${schoolSection(rec.safeties, '保底')}

  <h2>四、ED 首选建议</h2>
  <p><strong>推荐 ED 院校：</strong><span style="color:#1251b5;font-weight:bold">${rec.edSuggestionCN || ''}（${rec.edSuggestion || ''}）</span></p>
  <p>${rec.edReason || ''}</p>

  <h2>五、文书 Brainstorm</h2>
  <p><strong>📌 主文书核心方向</strong></p><p>${rec.essayMainTheme || ''}</p>
  <p><strong>💡 故事角度探索</strong></p>
  ${(rec.essayIdeas || []).map((idea, i) => `
  <div class="idea-box">
    <p><strong>角度 ${i + 1}：${idea.angle || ''}</strong></p>
    <p>${idea.description || ''}</p>
    <p style="color:#0694a2;font-style:italic">🤔 思考一下：${idea.prompt || ''}</p>
  </div>`).join('')}
  <p><strong>📝 补充文书方向（Why School）</strong></p><p>${rec.supplementalHints || ''}</p>

  <h2>六、近期行动清单</h2>
  <ol>${(rec.nextSteps || []).map(st => `<li>${st}</li>`).join('')}</ol>

  <div class="footer">本报告由 AdmitHere 顾问系统生成，仅供参考，请结合顾问意见使用。</div>
  </body></html>`;
}

function buildAdvisorHTML(s, rec) {
  const nameCN = (s['中文姓名'] || '').trim();
  const nameEN = (s['英文姓名'] || '').trim();
  const today = new Date().toLocaleDateString('zh-CN', { timeZone: 'Asia/Shanghai' });
  const colors = { '冲刺': '#1251b5', '匹配': '#0694a2', '保底': '#059669' };

  const schoolSection = (schools, label) => {
    if (!schools || !schools.length) return '';
    const bg = colors[label] || '#1251b5';
    return `
    <h3 style="color:${bg};margin-top:20px">▌ ${label}院校</h3>
    <table>
      <tr style="background:${bg};color:#fff">
        <th style="width:17%">学校</th><th style="width:7%;text-align:center">排名</th>
        <th style="width:28%">推荐理由</th><th style="width:22%">申请建议</th>
        <th style="width:26%">契合点 / 顾问备注</th>
      </tr>
      ${schools.map((sc, i) => `
      <tr style="background:${i % 2 === 0 ? '#f8fafc' : '#fff'}">
        <td><strong>${sc.schoolCN || ''}</strong><br><small style="color:#718096">${sc.school || ''}</small></td>
        <td style="text-align:center">${sc.rank || '-'}</td>
        <td>${sc.reason || ''}</td>
        <td>${sc.tips || ''}</td>
        <td style="color:#718096;font-style:italic">${sc.fitNote || ''}</td>
      </tr>`).join('')}
    </table>`;
  };

  const noteBox = label => `
  <table style="margin:8px 0;width:100%"><tr>
    <td style="background:#fef3c7;color:#92400e;font-weight:bold;width:22%;padding:8px 10px">📝 ${label}</td>
    <td style="background:#fffbeb;padding:8px 10px">&nbsp;</td>
  </tr></table>`;

  return `<html><head><meta charset="UTF-8">
  <style>
    body{font-family:"Microsoft YaHei",Arial,sans-serif;font-size:12pt;color:#1a202c;margin:40px}
    h1{color:#1251b5;font-size:20pt;text-align:center;margin-bottom:4px}
    h2{color:#1251b5;font-size:14pt;border-left:4px solid #1251b5;padding-left:10px;margin-top:28px}
    p{line-height:1.8;margin:6px 0}
    table{border-collapse:collapse;width:100%;margin:8px 0;font-size:11pt}
    th,td{border:1px solid #e2e8f0;padding:8px 10px;vertical-align:top}
    .subtitle{text-align:center;color:#718096;font-size:11pt;margin-bottom:10px}
    .warning{text-align:center;color:#dc2626;font-size:11pt;margin-bottom:30px}
    .idea-box{background:#f0f4f9;border-left:3px solid #1251b5;padding:10px 14px;margin:10px 0}
    .footer{text-align:center;color:#718096;font-size:10pt;margin-top:40px;border-top:1px solid #e2e8f0;padding-top:12px}
  </style></head><body>
  <h1>📋 定校建议 · 顾问工作版</h1>
  <div class="subtitle">${nameCN}（${nameEN}）· ${today}</div>
  <div class="warning">⚠️ 内部使用，请勿直接转发学生</div>

  <h2>一、AI 综合评估</h2><p>${rec.studentSummary || ''}</p>
  ${noteBox('顾问补充评估')}

  <h2>二、申请策略</h2><p>${rec.strategy || ''}</p>
  ${noteBox('策略调整备注')}

  <h2>三、选校清单审核</h2>
  <p style="color:#718096;font-style:italic;font-size:11pt">※ 灰斜体为 AI 生成的契合点，请在此列填写确认意见或替换学校</p>
  ${schoolSection(rec.reaches, '冲刺')}
  ${schoolSection(rec.matches, '匹配')}
  ${schoolSection(rec.safeties, '保底')}

  <h3 style="color:#1251b5;margin-top:20px">最终确认选校名单</h3>
  <table>
    <tr style="background:#1251b5;color:#fff">
      <th style="width:12%">类别</th><th style="width:30%">学校名称</th>
      <th style="width:18%;text-align:center">ED/EA/RD</th><th>备注</th>
    </tr>
    ${['冲刺1','冲刺2','冲刺3','匹配1','匹配2','匹配3','匹配4','匹配5','保底1','保底2','保底3'].map((l,i) => `
    <tr style="background:${i%2===0?'#f8fafc':'#fff'}">
      <td style="text-align:center">${l}</td><td></td><td></td><td></td>
    </tr>`).join('')}
  </table>

  <h2>四、ED 建议</h2>
  <p><strong>AI 推荐 ED：</strong><span style="color:#1251b5;font-weight:bold">${rec.edSuggestionCN || ''}（${rec.edSuggestion || ''}）</span></p>
  <p>${rec.edReason || ''}</p>
  ${noteBox('顾问 ED 最终决定')}

  <h2>五、文书 Brainstorm 审核</h2>
  <p><strong>📌 主文书方向（AI 建议）</strong></p><p>${rec.essayMainTheme || ''}</p>
  ${noteBox('顾问对主文书方向的调整意见')}
  <p><strong>💡 故事角度（AI 建议）</strong></p>
  ${(rec.essayIdeas || []).map((idea, i) => `
  <div class="idea-box">
    <p><strong>角度 ${i+1}：${idea.angle || ''}</strong></p>
    <p>${idea.description || ''}</p>
    <p style="color:#0694a2;font-style:italic">引导问题：${idea.prompt || ''}</p>
  </div>
  ${noteBox(`角度 ${i+1} 顾问评语`)}`).join('')}
  <p><strong>📝 补充文书方向</strong></p><p>${rec.supplementalHints || ''}</p>
  ${noteBox('顾问补充文书备注')}

  <h2>六、近期行动清单</h2>
  <ol>${(rec.nextSteps || []).map(st => `<li>${st}</li>`).join('')}</ol>
  ${noteBox('顾问补充行动项')}

  <div class="footer">内部工作文件 · ${nameCN} · AdmitHere · ${today}</div>
  </body></html>`;
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

    const completion = await openai.chat.completions.create({
      model: 'gpt-4o',
      messages: [{ role: 'user', content: buildPrompt(student) }],
      response_format: { type: 'json_object' },
      temperature: 0.7,
    });
    const rec = JSON.parse(completion.choices[0].message.content);

    const nameCN = (student['中文姓名'] || 'student').trim();
    res.status(200).json({
      success: true,
      studentDoc: Buffer.from(buildStudentHTML(student, rec)).toString('base64'),
      advisorDoc: Buffer.from(buildAdvisorHTML(student, rec)).toString('base64'),
      filename: nameCN,
    });

  } catch (err) {
    console.error('recommend error:', err);
    res.status(500).json({ success: false, error: err.message });
  }
}
