/**
 * BAFS 狀元工廠 — Google Sheet 後端（Apps Script Web App）
 *
 * 用法：
 * 1. 打開 script.google.com → 新建專案 → 貿全部貼入呢個檔案
 * 2. 掛一個 Google Sheet 做「綁定試算表」（Project Settings → Google Sheets Spreadsheet）
 *    或喺 CONFIG.SPREADSHEET_ID 填入現有 Sheet 的 ID
 * 3. 部署 → 新部署 → 類型：網頁應用程式 → 以「我」身份執行 → 誰可存取：任何人
 * 4. 把取得嘅 /exec URL 貼返去 index.html 嘅 GOOGLE_SHEET_URL
 *
 * 佢會自動維持：
 *   - Raw            ：一列一題嘅原始答題紀錄（資料來源）
 *   - 每個用家一個分頁：該學生每次交卷、課題正確率、最差題目
 *   - 全班總覽        ：學生對比 + 全班課題強弱 + 全班最差題目 Top 20
 */

const CONFIG = {
  SPREADSHEET_ID: '', // 留空 = 用綁定嘅試算表；否則填 Sheet ID
  SHARED_TOKEN: '',   // 設咗就要求前端帶 token（擋爬蟲/濫用，唔係真認證）
  RAW_SHEET: 'Raw',
  OVERVIEW_SHEET: '全班總覽',
  MC_SHEET: 'MC 答案分佈',
  RATE_LIMIT_SECONDS: 3,   // 同一暱稱最短間隔（秒）
  MAX_QUESTIONS: 200,      // 單次交卷最多題數
  MAX_STR: 60,             // 一般字串長度上限
  MAX_NAME: 40,            // 暱稱上限
  MAX_QTEXT: 300,          // 題目文字上限（截斷，控制 Raw 表體積）
  TAB_NAME_LIMIT: 40       // 分頁名稱上限（Google 係 100，留返啲空間）
};

const RAW_HEADERS = [
  'ts', 'submissionId', 'nickname', 'year', 'mode', 'score', 'total', 'accuracy',
  'qid', 'qnum', 'topic', 'userAns', 'correctAns', 'isCorrect', 'qtext'
];
const COL = { ts: 0, sub: 1, nick: 2, year: 3, mode: 4, score: 5, total: 6, acc: 7, qid: 8, qnum: 9, topic: 10, userAns: 11, correctAns: 12, ok: 13, qtext: 14 };
/* ==== 課題分類對照（2012–2025，共414題；同 index (4).html 題目庫同步）==== */
const TOPIC_LABELS = [
  "1(a) 營商環境",
  "1(a) 營商環境：企業擁有權類型",
  "1(a) 營商環境：企業擁有權類型（公營公司）",
  "1(a) 營商環境：商業道德與社會責任",
  "1(a) 營商環境：影響商業決定的因素",
  "1(a) 營商環境：香港的營商環境",
  "1(a) 營商環境：香港的營商環境（CEPA）",
  "1(a) 營商環境：香港的營商環境（WTO）",
  "1(a) 營商環境：香港的營商環境（全球一體化）",
  "1(a) 營商環境：香港的營商環境（國際貿易組織 APEC）",
  "1(c) 基礎會計：會計假設、原則及慣例",
  "1(c) 基礎會計：會計假設、原則及慣例（企業個體觀念）",
  "1(c) 基礎會計：會計假設與原則",
  "1(c) 基礎會計：會計假設與原則（歷史成本原則）",
  "1(c) 基礎會計：會計循環",
  "1(c) 基礎會計：會計循環與期末調整",
  "1(c) 基礎會計：會計循環與複式記帳",
  "1(c) 基礎會計：會計循環與試算表",
  "1(c) 基礎會計：會計循環與財務報表",
  "1(c) 基礎會計：會計循環（原始分錄帳簿與過帳）",
  "1(c) 基礎會計：會計循環（原始分錄帳簿）",
  "1(c) 基礎會計：會計循環（損益表與財務比率）",
  "1(c) 基礎會計：會計循環（損益表）",
  "1(c) 基礎會計：會計循環（會計等式）",
  "1(c) 基礎會計：會計循環（複式記帳法）",
  "1(c) 基礎會計：會計循環（試算表）",
  "1(c) 基礎會計：會計循環（財務報表）",
  "1(c) 基礎會計：會計的目的及角色",
  "1(c) 基礎會計：會計的目的及角色（會計循環）",
  "1(c) 基礎會計：財務報表的用途",
  "1(d) 基礎個人理財：以證券交易作投資",
  "1(d) 基礎個人理財：個人理財策劃",
  "1(d) 基礎個人理財：消費者信貸",
  "1(d) 基礎個人理財：策劃個人理財及投資",
  "1(d) 基礎個人理財：金錢的時間值",
  "3(a)【選修】管理導論：中小型企業",
  "3(a)【選修】管理導論：中小型企業與企業家精神",
  "3(a)【選修】管理導論：中小型企業（企業家精神）",
  "3(a)【選修】管理導論：主要商業功能",
  "3(a)【選修】管理導論：主要商業功能 / 管理功能",
  "3(a)【選修】管理導論：主要商業功能（人力資源管理）",
  "3(a)【選修】管理導論：主要商業功能（營運管理）",
  "3(a)【選修】管理導論：主要商業功能（財務管理）",
  "3(a)【選修】管理導論：主要商業功能（資訊管理/內部溝通）",
  "3(a)【選修】管理導論：主要商業功能（資訊管理）",
  "3(a)【選修】管理導論：有效管理的原則（MBO）",
  "3(a)【選修】管理導論：有效管理的原則（權力與責任）",
  "3(a)【選修】管理導論：有效管理的原則（目標管理 MBO）",
  "3(a)【選修】管理導論：管理功能",
  "3(a)【選修】管理導論：管理功能（SMART目標原則）",
  "3(a)【選修】管理導論：管理功能（控制 vs 組織）",
  "3(a)【選修】管理導論：管理功能（計劃）",
  "3(a)【選修】管理導論：管理功能（領導）",
  "3(b)【選修】財務管理：財務分析",
  "3(b)【選修】財務管理：財務分析（營運資金/流動比率）",
  "3(c)【選修】人力資源管理：內部溝通",
  "3(c)【選修】人力資源管理：發展優質人力（內部溝通過程）",
  "3(c)【選修】人力資源管理：發展優質人力（內部溝通）"
];
const TOPIC_ROWS = {
  '2012': [39,22,24,35,33,50,4,35,34,8,6,1,3,28,40,23,23,30,45,27,32,54,35,33,29,34,11,19,37,56],
  '2013': [1,33,14,51,1,33,3,46,24,15,35,5,38,22,34,5,33,17,22,31,57,5,37,32,33,53,53,16,44,27],
  '2014': [1,33,3,33,10,35,33,32,53,25,22,30,34,28,4,13,49,22,53,29,47,36,55,44,38,26,33,32,24,9],
  '2015': [1,3,52,20,30,48,25,32,44,26,35,5,42,26,29,5,34,22,43,33,41,7,23,48,24,8,2,33,21,22],
  '2016': [1,48,48,33,12,1,48,48,32,48,25,22,26,34,3,1,5,24,34,35,24,48,23,48,33,29,30,1,22,22],
  '2017': [5,23,30,48,28,5,33,38,24,34,1,53,34,35,29,4,30,35,23,35,53,0,32,48,12,30,48,24,53,1],
  '2018': [5,25,30,24,48,12,3,24,38,26,48,24,32,3,34,48,23,1,34,22,53,30,27,1,33,48,48,1,29,33],
  '2019': [1,12,48,24,5,23,38,34,3,53,1,26,35,26,30,22,3,33,24,33,35,5,28,30,29,32,48,23,34,48],
  '2020': [1,48,24,30,23,48,48,34,53,5,29,48,28,22,30,1,48,22,33,12,38,3,53,5,33,48,1,30,48,22],
  '2021': [5,48,27,34,48,22,4,48,23,32,22,1,25,32,29,3,48,12,3,33,35,53,33,5,48,24,33,38,26,34],
  '2022': [5,35,27,22,34,1,48,25,30,24,4,23,30,24,48,23,32,48,53,53,1,38,48,33,1,33,33,5,22,25],
  '2023': [5,48,27,34,28,1,22,48,32,23,1,30,48,53,33,5,12,35,33,22,1,35,26,34,53,30,3,18,33,48],
  '2024': [5,48,29,33,29,48,25,32,5,23,48,12,5,32,24,30,48,23,4,30,53,35,24,38,34,1,12,35,34,26],
  '2025': [5,5,5,5,1,1,1,3,34,32,32,33,33,33,30,30,27,28,26,29,22,22,22,23],
};
const TOPIC_BY_QID = {};
Object.keys(TOPIC_ROWS).forEach(y => {
  TOPIC_ROWS[y].forEach((c, i) => {
    TOPIC_BY_QID[y + String(i + 1).padStart(2, '0')] = TOPIC_LABELS[c];
  });
});

/* ============================ 入口 ============================ */

function doGet() {
  return json_({ ok: true, app: 'BAFS MC', hint: 'POST JSON to this endpoint' });
}

function doPost(e) {
  try {
    const body = safeParse_(e && e.postData && e.postData.contents);
    if (!body) return json_({ ok: false, err: 'bad_json' });

    if (CONFIG.SHARED_TOKEN && body.token !== CONFIG.SHARED_TOKEN) {
      return json_({ ok: false, err: 'bad_token' });
    }

    const problems = validate_(body);
    if (problems.length) return json_({ ok: false, err: 'invalid', detail: problems });

    const nickname = str_(body.nickname, CONFIG.MAX_NAME);
    const submissionId = str_(body.submissionId, CONFIG.MAX_STR);

    // 限速（CacheService，3 秒一擊）
    const cache = CacheService.getScriptCache();
    const rlKey = 'rl:' + hash_(nickname);
    if (cache.get(rlKey)) return json_({ ok: false, err: 'too_many_requests' });
    cache.put(rlKey, '1', CONFIG.RATE_LIMIT_SECONDS);

    const ss = openSheet_();
    const raw = getRaw_(ss);

    // 去重：同一 submissionId 只入一次（防重複交卷/重試）
    if (submissionId && raw.seen[submissionId]) {
      return json_({ ok: true, duplicate: true });
    }

    const rows = buildRows_(body, nickname, submissionId);
    raw.sheet.getRange(raw.sheet.getLastRow() + 1, 1, rows.length, RAW_HEADERS.length).setValues(rows);
    const allRows = raw.values.concat(rows);

    updateUserTab_(ss, nickname, allRows);
    updateOverview_(ss, allRows);
    updateMCStats_(ss, allRows);

    return json_({ ok: true, nickname: nickname, submissionId: submissionId });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, err: 'server_error', detail: String(err && err.message || err) });
  }
}

/* ============================ 驗證 ============================ */

function validate_(b) {
  const p = [];
  if (typeof b.nickname !== 'string' || !b.nickname.trim()) p.push('nickname');
  if (typeof b.year !== 'string') p.push('year');
  if (b.mode !== 'practice' && b.mode !== 'exam') p.push('mode');
  if (!isNum_(b.score, 0, 1000)) p.push('score');
  if (!isNum_(b.total, 0, 1000)) p.push('total');
  if (!isNum_(b.accuracy, 0, 100)) p.push('accuracy');
  if (b.answered !== undefined && !isNum_(b.answered, 0, 1000)) p.push('answered');
  if (!Array.isArray(b.questions)) p.push('questions');
  else if (b.questions.length > CONFIG.MAX_QUESTIONS) p.push('questions_too_long');
  else {
    b.questions.forEach((q, i) => {
      if (!q || typeof q !== 'object') { p.push('questions[' + i + ']'); return; }
      if (q.correctAns !== 'A' && q.correctAns !== 'B' && q.correctAns !== 'C' && q.correctAns !== 'D') p.push('questions[' + i + '].correctAns');
      if (q.userAns !== '' && q.userAns !== 'A' && q.userAns !== 'B' && q.userAns !== 'C' && q.userAns !== 'D') p.push('questions[' + i + '].userAns');
      if (q.isCorrect !== 0 && q.isCorrect !== 1) p.push('questions[' + i + '].isCorrect');
      if (q.topic !== undefined && typeof q.topic !== 'string') p.push('questions[' + i + '].topic');
      if (q.qtext !== undefined && typeof q.qtext !== 'string') p.push('questions[' + i + '].qtext');
    });
  }
  return p;
}

function isNum_(v, min, max) {
  return typeof v === 'number' && isFinite(v) && v >= min && v <= max;
}

/** 有冇作答？未作答（userAns 空白）唔計對錯，亦唔計入分母 */
function isAns_(r) {
  const v = r[COL.userAns];
  return v !== undefined && v !== null && String(v) !== '';
}

/* ============================ 寫入 ============================ */

function buildRows_(b, nickname, submissionId) {
  const ts = new Date().toISOString();
  const nicknameSafe = guard_(str_(nickname, CONFIG.MAX_NAME));
  const year = guard_(str_(b.year, CONFIG.MAX_STR));
  const mode = b.mode;
  const score = b.score, total = b.total, accuracy = b.accuracy;

  return b.questions.map((q, i) => [
    ts,
    submissionId,
    nicknameSafe,
    year,
    mode,
    score,
    total,
    accuracy,
    guard_(str_(q.qid, 20)),
    i + 1,
    guard_(str_(q.topic, CONFIG.MAX_STR)),
    q.userAns || '',
    q.correctAns,
    q.isCorrect,
    guard_(q.qtext === undefined ? '' : q.qtext, CONFIG.MAX_QTEXT)
  ]);
}

/* ============================ 每個用家一個分頁 ============================ */

function updateUserTab_(ss, nickname, allRows) {
  const sheet = ensureUserSheet_(ss, nickname);
  const mine = allRows.filter(r => r[2] === guard_(nickname) || r[2] === nickname);

  // Block 1：每次交卷
  const bySub = {};
  mine.forEach(r => {
    const key = r[1] || (r[0] + r[5]);
    if (!bySub[key]) bySub[key] = { ts: r[0], year: r[3], mode: r[4], score: r[5], total: r[6], acc: r[7], a: 0, s: 0 };
    if (isAns_(r)) bySub[key].a += 1; else bySub[key].s += 1;
  });
  const sessions = Object.keys(bySub).map(k => bySub[k])
    .sort((a, b) => (a.ts < b.ts ? 1 : -1))
    .map(s => [s.ts, s.year, s.mode === 'exam' ? '模擬考' : '練習', s.score + '/' + s.total, s.a, s.s, s.acc + '%']);

  // Block 2：課題正確率（分母只計已作答，未作答單獨一欄）
  const topics = {};
  mine.forEach(r => {
    const t = r[COL.topic] || '未分類';
    topics[t] = topics[t] || { a: 0, c: 0, s: 0 };
    if (isAns_(r)) {
      topics[t].a += 1;
      if (r[COL.ok] === 1) topics[t].c += 1;
    } else {
      topics[t].s += 1;
    }
  });
  const topicRows = Object.keys(topics).map(t => {
    const v = topics[t];
    return [t, v.a, v.s, v.c, v.a - v.c, v.a ? Math.round((v.c / v.a) * 100) + '%' : '0%'];
  }).sort((a, b) => pctNum_(a[5]) - pctNum_(b[5])); // 由弱到強

  // Block 3：最差題目（題目文字 + 累計錯誤次數；未作答唔計錯）
  const qs = {};
  mine.forEach(r => {
    const id = String(r[COL.qid]);
    qs[id] = qs[id] || { topic: r[COL.topic], a: 0, w: 0, s: 0 };
    if (isAns_(r)) {
      qs[id].a += 1;
      if (r[COL.ok] !== 1) qs[id].w += 1;
    } else {
      qs[id].s += 1;
    }
  });
  const text = qtextMap_(mine);
  const weakRows = Object.keys(qs)
    .filter(id => qs[id].w > 0)
    .map(id => [
      id,
      text[id] || '(題目未同步)',
      qs[id].topic,
      qs[id].w,                                        // 該題累計錯咗幾多次
      qs[id].a,                                        // 已作答次數（分母）
      qs[id].a ? Math.round((qs[id].w / qs[id].a) * 100) + '%' : '0%'
    ])
    .sort((a, b) => (b[3] - a[3]) || (pctNum_(b[5]) - pctNum_(a[5])))
    .slice(0, 50);

  sheet.clear();
  let row = 1;
  row = writeBlock_(sheet, row, guard_(nickname) + ' — 成績分析', [], []);
  row += 1;
  row = writeBlock_(sheet, row, '① 每次交卷（新→舊）', ['時間', '年份', '模式', '分數', '已作答', '未作答', '正確率'], sessions);
  row += 1;
  row = writeBlock_(sheet, row, '② 各課題正確率（由弱到強，分母只計已作答）',
    ['課題', '已作答', '未作答', '答對', '答錯', '正確率'], topicRows);
  row += 1;
  row = writeBlock_(sheet, row, '③ 需要重溫的題目（錯誤次數由多至少；未作答不當錯）',
    ['題號 qid', '題目', '課題', '錯誤次數（累計）', '已作答', '錯誤率'], weakRows);
  sheet.setColumnWidth(2, 420);
  sheet.setFrozenRows(0);
}

/* ============================ 全班總覽 ============================ */

function updateOverview_(ss, allRows) {
  let sheet = ss.getSheetByName(CONFIG.OVERVIEW_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.OVERVIEW_SHEET);
  sheet.clear();

  // A. 學生對比
  const students = {};
  const topics = {};
  const qs = {};
  allRows.forEach(r => {
    const name = r[COL.nick];
    students[name] = students[name] || { subs: {}, a: 0, c: 0, s: 0, topics: {} };
    const s = students[name];
    s.subs[r[1] || (r[0] + r[5])] = 1;
    const ans = isAns_(r);
    if (ans) {
      s.a += 1;
      if (r[COL.ok] === 1) s.c += 1;
    } else {
      s.s += 1;
    }

    const t = r[COL.topic] || '未分類';
    s.topics[t] = s.topics[t] || { a: 0, c: 0, s: 0 };
    if (ans) {
      s.topics[t].a += 1;
      if (r[COL.ok] === 1) s.topics[t].c += 1;
    } else {
      s.topics[t].s += 1;
    }

    topics[t] = topics[t] || { a: 0, c: 0, s: 0 };
    if (ans) {
      topics[t].a += 1;
      if (r[COL.ok] === 1) topics[t].c += 1;
    } else {
      topics[t].s += 1;
    }

    const id = String(r[COL.qid]);
    qs[id] = qs[id] || { topic: r[COL.topic], a: 0, w: 0, s: 0 };
    if (ans) {
      qs[id].a += 1;
      if (r[COL.ok] !== 1) qs[id].w += 1;
    } else {
      qs[id].s += 1;
    }
  });

  const studentRows = Object.keys(students).map(name => {
    const s = students[name];
    let worst = '', worstRate = 101;
    Object.keys(s.topics).forEach(t => {
      const v = s.topics[t];
      const rate = v.a ? Math.round((v.c / v.a) * 100) : 100;
      if (v.a >= 3 && rate < worstRate) { worstRate = rate; worst = t; }
    });
    return [name, Object.keys(s.subs).length, s.a, s.s, s.c, s.a - s.c,
      s.a ? Math.round((s.c / s.a) * 100) + '%' : '0%',
      worst || '（資料不足）', worstRate === 101 ? '—' : worstRate + '%'];
  }).sort((a, b) => pctNum_(a[6]) - pctNum_(b[6])); // 最弱排最上

  // B. 全班課題（分母只計已作答）
  const topicRows = Object.keys(topics).map(t => {
    const v = topics[t];
    const rate = v.a ? Math.round((v.c / v.a) * 100) : 0;
    return [t, v.a, v.s, v.c, v.a - v.c, rate + '%', rate < 50 ? '⚠️ 全班弱項' : (rate < 70 ? '待加強' : 'OK')];
  }).sort((a, b) => pctNum_(a[5]) - pctNum_(b[5]));

  // C. 全班最差題目（題目文字 + 全班累計錯誤次數；未作答不當錯）
  const text = qtextMap_(allRows);
  const weakRows = Object.keys(qs)
    .filter(id => qs[id].w > 0)
    .map(id => [
      id,
      text[id] || '(題目未同步)',
      qs[id].topic,
      qs[id].w,
      qs[id].a,
      qs[id].a ? Math.round((qs[id].w / qs[id].a) * 100) + '%' : '0%'
    ])
    .sort((a, b) => (b[3] - a[3]) || (pctNum_(b[5]) - pctNum_(a[5])))
    .slice(0, 20);

  let row = 1;
  row = writeBlock_(sheet, row, '全班總覽（自動計算）', [], []);
  row += 1;
  row = writeBlock_(sheet, row, '① 學生成績（最弱排最上）',
    ['學生', '交卷次數', '已作答', '未作答', '答對', '答錯', '正確率', '最弱課題', '該課題正確率'], studentRows);
  row += 1;
  row = writeBlock_(sheet, row, '② 全班課題強弱（分母只計已作答）',
    ['課題', '已作答', '未作答', '答對', '答錯', '正確率', '評語'], topicRows);
  row += 1;
  row = writeBlock_(sheet, row, '③ 全班最差題目 Top 20（錯誤次數由多至少；未作答不當錯）',
    ['題號 qid', '題目', '課題', '錯誤次數（累計）', '已作答', '錯誤率'], weakRows);
  sheet.setColumnWidth(2, 420);
}

/* ============================ 每題 MC 答案分佈（全班） ============================ */

/** 由 qid 抽年份（例 202503 → 2025）；抽唔到就當最舊排最後 */
function yearOf_(qid) {
  const m = String(qid).match(/^(\d{4})\d{2}$/);
  return m ? m[1] : '';
}

function updateMCStats_(ss, allRows) {
  let sheet = ss.getSheetByName(CONFIG.MC_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.MC_SHEET);
  sheet.clear();

  const text = qtextMap_(allRows);
  const dist = {};
  allRows.forEach(r => {
    const id = String(r[COL.qid]);
    if (!id || id === 'undefined' || id === 'null') return;
    if (!dist[id]) {
      dist[id] = {
        topic: r[COL.topic], correct: r[COL.correctAns] || '',
        A: 0, B: 0, C: 0, D: 0, blank: 0
      };
    }
    const d = dist[id];
    const ua = r[COL.userAns];
    if (ua === 'A') d.A += 1;
    else if (ua === 'B') d.B += 1;
    else if (ua === 'C') d.C += 1;
    else if (ua === 'D') d.D += 1;
    else d.blank += 1;
    if (!d.correct) d.correct = r[COL.correctAns] || '';
  });

  const rows = Object.keys(dist).map(id => {
    const d = dist[id];
    const total = d.A + d.B + d.C + d.D;           // 只計已作答
    const order = [['A', d.A], ['B', d.B], ['C', d.C], ['D', d.D]].sort((x, y) => y[1] - x[1]);
    const topN = total ? order[0][1] : 0;
    const top = order.filter(x => x[1] === topN && x[1] > 0).map(x => x[0]).join('/');
    const topPct = total ? Math.round((topN / total) * 100) + '%' : '—';
    let verdict;
    if (!total) verdict = '（未有人作答）';
    else if (top.split('/').indexOf(d.correct) >= 0) verdict = '✓ 多數人答對';
    else verdict = '✗ 多數人答錯';
    return [
      yearOf_(id), id, text[id] || '(題目未同步)', d.topic || '', d.correct || '—',
      d.A, d.B, d.C, d.D, d.blank, total,
      total ? top : '—', topPct, verdict
    ];
  }).sort((x, y) => {
    // 年份由新到舊，同年份就按題號細到大
    if (x[0] !== y[0]) return x[0] > y[0] ? -1 : 1;
    return x[1] < y[1] ? -1 : (x[1] > y[1] ? 1 : 0);
  });

  let row = 1;
  row = writeBlock_(sheet, row, '每題 MC 答案分佈（全班累計，按年份由新到舊）',
    ['年份', '題號 qid', '題目', '課題', '正確答案',
      'A 人數', 'B 人數', 'C 人數', 'D 人數', '未作答', '作答人數',
      '多數人選', '最高票佔比', '多數答案'], rows);
  sheet.setColumnWidth(3, 420);
}

/* ============================ 小工具 ============================ */

function writeBlock_(sheet, startRow, title, headers, rows) {
  let r = startRow;
  if (title) {
    sheet.getRange(r, 1, 1, 1).setValue(title).setFontWeight('bold').setFontSize(12);
    r += 1;
  }
  if (headers && headers.length) {
    sheet.getRange(r, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#2D2D2D').setFontColor('#FFFFFF');
    r += 1;
  }
  if (rows && rows.length) {
    const cols = headers.length || rows[0].length;
    // 補齊欄數，避免 setValues 因參差長度報錯
    const fixed = rows.map(row => {
      const out = row.slice(0, cols);
      while (out.length < cols) out.push('');
      return out;
    });
    sheet.getRange(r, 1, fixed.length, cols).setValues(fixed);
    r += fixed.length;
  }
  return r;
}

/** 每個用家一個分頁：同名即係同一個人（相同暱稱會共用一個分頁） */
function ensureUserSheet_(ss, nickname) {
  const base = tabName_(nickname) || '訪客';
  const existing = ss.getSheetByName(base);
  if (existing) return existing;
  try {
    return ss.insertSheet(base);
  } catch (e) {
    // 撞名 race：再攞一次
    const again = ss.getSheetByName(base);
    if (again) return again;
    throw e;
  }
}

function tabName_(name) {
  const s = str_(name, CONFIG.TAB_NAME_LIMIT)
    .replace(/[[\]:*?/\\]/g, '')   // Sheet 分頁名唔可以有呢啲字元
    .replace(/^[=+\-@]+/, '')      // 開頭唔好係公式符號
    .trim();
  return s || '訪客';
}

/** 防公式注入：字串以 = + - @ 開頭會被 Sheets 當公式，加個 ' 前綴 */
function guard_(s, max) {
  const v = str_(s, max || CONFIG.MAX_STR);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function str_(v, max) {
  if (v === undefined || v === null) return '';
  let s = String(v).replace(/[\u0000-\u001F\u007F]/g, '');
  if (s.length > max) s = s.slice(0, max);
  return s;
}

function pctNum_(s) {
  const n = parseInt(String(s), 10);
  return isFinite(n) ? n : 0;
}

function openSheet_() {
  if (CONFIG.SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('未綁定試算表：請喺 Project Settings 綁定，或填 CONFIG.SPREADSHEET_ID');
  return ss;
}

function getRaw_(ss) {
  let sheet = ss.getSheetByName(CONFIG.RAW_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.RAW_SHEET);
    sheet.appendRow(RAW_HEADERS);
  }
  if (sheet.getLastRow() === 0) sheet.appendRow(RAW_HEADERS);

  const values = sheet.getDataRange().getValues();
  const hasHeader = values.length > 0 && String(values[0][0]) === 'ts';
  const rows = hasHeader ? values.slice(1) : values;
  // 課題統一：以題目庫分類為準（舊交卷紀錄都會跟新分類做統計）
  rows.forEach(r => {
    const t = TOPIC_BY_QID[String(r[COL.qid] === undefined ? '' : r[COL.qid] || '')];
    if (t) r[COL.topic] = t;
  });

  const seen = {};
  rows.forEach(r => { if (r[1]) seen[String(r[1])] = 1; });

  if (!hasHeader) {
    // 已有資料但冇 header → 插一行，唔好覆蓋第一行數據
    if (values.length === 0 || String(values[0] || '') !== '') sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, RAW_HEADERS.length).setValues([RAW_HEADERS]);
  } else if (values[0].length < RAW_HEADERS.length) {
    // 舊版 Raw 得 14 欄 → 補 header（唔會掂到舊數據）
    sheet.getRange(1, 1, 1, RAW_HEADERS.length).setValues([RAW_HEADERS]);
  }

  return { sheet: sheet, values: rows, seen: seen };
}

/** 由 Raw 建立 qid → 題目文字 對照表（取最新一筆非空白） */
function qtextMap_(rows) {
  const m = {};
  rows.forEach(r => {
    const id = r[COL.qid];
    const t = r[COL.qtext];
    if (id !== undefined && id !== null && String(id) !== '' && t) m[String(id)] = t;
  });
  return m;
}

function safeParse_(text) {
  if (!text) return null;
  try { return JSON.parse(text); } catch (e) { return null; }
}

function hash_(s) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s)).slice(0, 32);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================ 一次性工具（手動執行） ============================ */

/** 瀏覽器執行：一鍵重建所有分頁（改咗邏輯之後跑一次） */
function rebuildAll() {
  const ss = openSheet_();
  const raw = getRaw_(ss);
  const names = {};
  raw.values.forEach(r => { names[r[2]] = 1; });
  Object.keys(names).forEach(n => updateUserTab_(ss, n, raw.values));
  updateOverview_(ss, raw.values);
  updateMCStats_(ss, raw.values);
  Logger.log('已重建 %s 個分頁', Object.keys(names).length);
}
