/***********************************************************************
 * 가로세로 퀴즈 웹앱 - Google Apps Script 백엔드
 * ---------------------------------------------------------------------
 * 역할: teacher.html / student.html 과 Google Sheets 사이의 서버
 *  - Gemini API 키를 안전하게 보관 (브라우저에 노출 안 됨)
 *  - 퀴즈 저장 / 목록 / 불러오기
 *  - 학생 제출 채점 (정답은 서버에만 존재)
 *  - 제출 기록 저장 / 조회
 ***********************************************************************/

/* ===================== 설정 ===================== */
/* Gemini API 키
 * 권장: Apps Script 편집기 → 왼쪽 ⚙ 프로젝트 설정 → [스크립트 속성] 에
 *       속성 이름 GEMINI_API_KEY / 값 = 실제 키  로 저장해 두세요.
 *       그러면 나중에 이 코드를 통째로 갈아끼워도 키가 지워지지 않습니다.
 * 그게 번거로우면 아래 GEMINI_API_KEY_INLINE 줄에 직접 붙여넣어도 됩니다. */
const GEMINI_API_KEY_INLINE = '여기에_본인_제미나이_API_키_붙여넣기';
const GEMINI_MODEL          = 'gemini-3.1-flash-lite';

// 키는 AI 단어/힌트 생성 때만 필요하다. 학생 요청마다 읽지 않도록 그때만 꺼낸다.
function geminiKey_() {
  try {
    const p = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (p) return p;
  } catch (e) {}
  return GEMINI_API_KEY_INLINE;
}

// 시트 탭 이름 (자동 생성됨, 그대로 두면 됨)
const SHEET_QUIZ = '퀴즈목록';
const SHEET_SUB  = '제출기록';
/* ====================================================================== */


/* ---------- 메인 진입점 ---------- */
// student.html 공유링크 등이 GET으로 들어올 수도 있어 둘 다 처리
function doGet(e)  { return handle(e); }
function doPost(e) { return handle(e); }

/* 학생이 쓰는 동작만 열어둔다. 나머지(제출 기록 조회, 퀴즈 삭제 등)는
 * 선생님 비밀번호가 있어야 한다.
 * 이게 없으면 웹앱 주소만 아는 사람이 학생 이름과 점수를 전부 내려받거나
 * 퀴즈를 지울 수 있다 — student.html 소스에 주소가 그대로 들어 있으므로
 * 주소는 비밀이 아니다. index.html 의 잠금 화면은 눈에 보이는 1차 방어일 뿐이고,
 * 진짜 잠금은 여기다. (APPS_SCRIPT_보안패치.md 참고)
 * 비밀번호는 코드에 적지 않고 스크립트 속성 TEACHER_PW 에 둔다. */
const OPEN_ACTIONS = ['ping', 'getQuizForPlay', 'submitAnswer'];

function teacherPw_() {
  try { return PropertiesService.getScriptProperties().getProperty('TEACHER_PW') || ''; }
  catch (e) { return ''; }
}

function handle(e) {
  let res;
  try {
    const req = parseRequest_(e);
    const action = req.action;

    if (OPEN_ACTIONS.indexOf(action) === -1) {
      const pw = teacherPw_();
      // TEACHER_PW 를 아직 안 만들었으면 잠그지 않는다 (설정 전에 선생님이 갇히면 안 되니까).
      // 스크립트 속성에 TEACHER_PW 를 넣는 순간부터 잠금이 켜진다.
      if (pw && String(req.pw || '') !== pw) {
        return json_({ ok: false, error: '선생님 비밀번호가 필요합니다.' });
      }
    }

    switch (action) {
      case 'ping':            res = { ok: true, msg: 'pong' }; break;
      case 'generateWords':   res = generateWords_(req);       break;
      case 'generateClues':   res = generateClues_(req);       break;
      case 'saveQuiz':        res = saveQuiz_(req);            break;
      case 'getQuizList':     res = getQuizList_(req);         break;
      case 'getQuizForEdit':  res = getQuizForEdit_(req);      break;  // 선생님용(정답 포함)
      case 'getQuizForPlay':  res = getQuizForPlay_(req);      break;  // 학생용(정답 제거)
      case 'deleteQuiz':      res = deleteQuiz_(req);          break;
      case 'submitAnswer':    res = submitAnswer_(req);        break;
      case 'getSubmissions':  res = getSubmissions_(req);      break;
      case 'getWordStats':    res = getWordStats_(req);        break;
      default: res = { ok: false, error: '알 수 없는 action: ' + action };
    }
  } catch (err) {
    res = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return json_(res);
}


/* ---------- 요청 파싱 ---------- */
// HTML에서 fetch 호출 시 Content-Type을 지정하지 않아 text/plain 으로 전송됨
// → 브라우저 CORS preflight(OPTIONS) 가 발생하지 않아 GAS와 잘 통신됨
function parseRequest_(e) {
  if (e && e.postData && e.postData.contents) {
    return JSON.parse(e.postData.contents);
  }
  if (e && e.parameter && e.parameter.payload) {
    return JSON.parse(e.parameter.payload);
  }
  // GET 쿼리스트링으로 들어온 단순 요청 (예: ?action=getQuizForPlay&quizId=xxx)
  if (e && e.parameter && e.parameter.action) {
    return e.parameter;
  }
  return {};
}

// JSON 응답 (GAS는 별도 CORS 헤더 설정 불가하지만, 단순요청 응답은 브라우저가 읽을 수 있음)
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/* ===================== 시트 준비 ===================== */
function getSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  } else if (!cacheGet_('init_' + name)) {
    // 헤더 확인은 한 번만 하면 된다. 예전엔 요청마다 getLastRow() 를 호출했다.
    if (sh.getLastRow() === 0) {
      sh.appendRow(headers);
      sh.setFrozenRows(1);
    }
    cachePut_('init_' + name, 1);
  }
  return sh;
}

const QUIZ_HEADERS = ['quizId', 'quizName', 'topic', 'gridSize', 'difficulty', 'createdAt', 'quizData', 'folder'];
const SUB_HEADERS  = ['submissionId', 'quizId', 'quizName', 'studentName', 'studentId',
                      'submittedAt', 'score', 'totalCells', 'correctCells', 'wrongCells', 'timeSpent', 'wordResults'];

function quizSheet_() { return getSheet_(SHEET_QUIZ, QUIZ_HEADERS); }
function subSheet_()  { return getSheet_(SHEET_SUB,  SUB_HEADERS); }


/* ===================== 캐시 (동시접속 대응의 핵심) =====================
 * 한 반 30명이 같은 링크를 동시에 열면 예전에는 똑같은 퀴즈를 시트에서 30번 읽었다.
 * 내용이 같으므로 한 번만 읽어 캐시에 담고, 퀴즈를 저장/삭제할 때만 버린다.
 * 이 캐시로 학생 1명당 시트 접근이 (읽기 2회 + 쓰기 1회) → (쓰기 1회) 로 줄어든다.
 * ==================================================================== */
const CACHE_TTL = 21600;        // 6시간 (Apps Script 캐시 최대값)
const CACHE_MAX = 95 * 1024;    // 캐시 값 상한 100KB 보다 조금 작게

function cache_() { return CacheService.getScriptCache(); }

function cacheGet_(key) {
  try { const v = cache_().get(key); return v ? JSON.parse(v) : null; }
  catch (e) { return null; }
}

function cachePut_(key, value, ttl) {
  try {
    const v = JSON.stringify(value);
    if (v.length > CACHE_MAX) return;   // 너무 크면 캐시 없이 매번 읽는다 (동작은 동일)
    cache_().put(key, v, ttl || CACHE_TTL);
  } catch (e) {}
}

// 퀴즈 한 개와 관련된 캐시를 모두 버린다 (저장/삭제 직후 반드시 호출)
function bustQuiz_(quizId) {
  try { cache_().removeAll(['row_' + quizId, 'play_' + quizId, 'grade_' + quizId, 'subcount']); }
  catch (e) {}
}

/* 30명이 동시에 제출하면 시트 쓰기가 순간적으로 실패할 수 있다.
 * 실패를 학생 화면에 그대로 띄우지 말고 잠깐 쉬었다가 다시 시도한다.
 * (대기 시간을 랜덤하게 줘서 재시도가 또 한꺼번에 몰리지 않게 한다) */
function appendRowSafe_(sh, row) {
  var lastErr;
  for (var i = 0; i < 4; i++) {
    try { sh.appendRow(row); return; }
    catch (e) {
      lastErr = e;
      Utilities.sleep((i + 1) * (300 + Math.floor(Math.random() * 700)));
    }
  }
  throw lastErr;
}


/* ===================== 1) Gemini 단어/힌트 생성 ===================== */
function generateWords_(req) {
  const topic      = String(req.topic || '').trim();
  const difficulty = req.difficulty || '보통';     // 쉬움 / 보통 / 어려움
  const gridSize   = Number(req.gridSize) || 10;
  const target     = Number(req.count) || Math.max(8, Math.round(gridSize * 1.2));
  const poolCount  = target + 6;                 // 넉넉히 생성 → teacher.html이 퍼즐에 맞게 선별
  const maxLen     = Math.min(gridSize, 5);      // 짧을수록 글자 공유가 잘 되어 교차 ↑

  if (!topic) return { ok: false, error: '주제(topic)가 비어 있습니다.' };
  const apiKey = geminiKey_();
  if (!apiKey || apiKey.indexOf('여기에') === 0) {
    return { ok: false, error: 'Gemini API 키가 설정되지 않았습니다. (Apps Script 프로젝트 설정 → 스크립트 속성 → GEMINI_API_KEY)' };
  }

  const clueStyle = {
    '쉬움':   '아주 직접적이고 친절하게, 초등학생도 알 수 있게',
    '보통':   '적당한 수준으로',
    '어려움': '함축적이고 추론이 필요하게 (단, 정답이 한 개로 특정되도록)'
  }[difficulty] || '적당한 수준으로';

  const prompt =
`당신은 한국어 가로세로(크로스워드) 낱말 퀴즈 출제 전문가입니다.
아래 조건으로 퍼즐용 단어와 힌트를 만들어 주세요.

[주제] ${topic}
[난이도] ${difficulty}
[단어 수] ${poolCount}개 (퍼즐 구성을 위해 넉넉히)

[가장 중요한 규칙 — 글자 공유]
가로세로 퍼즐은 단어들이 같은 '글자(음절)'에서 교차합니다.
서로 전혀 겹치는 글자가 없는 단어들로만 구성하면 퍼즐을 만들 수 없습니다.
- 단어 집합 전체에서 같은 글자가 여러 단어에 반복 등장하도록 의도적으로 선정하세요.
  예) '이순신'·'신사임당'·'김유신' → '신' 글자 공유로 교차 가능
  예) '광합성'·'합창' → '합' 글자 공유
- 인물 이름처럼 글자 공유가 어려운 주제라면, 관련된 일반 용어(사건·장소·개념·작품·제도 등)도
  반드시 섞어서 공통 글자(예: 한, 국, 정, 사, 신, 대, 왕, 전)가 자주 나오게 만드세요.

[그 외 규칙]
- answer 는 띄어쓰기/숫자/영문/특수문자 없는 순수 한글 단어
- 글자 수는 2~${maxLen}글자 (2~3글자를 충분히 포함시켜 교차를 쉽게)
- 힌트(clue)는 ${clueStyle} 작성하되 정답 단어를 그대로 포함하지 말 것
- 같은 단어 중복 금지

반드시 아래 JSON 형식으로만, 다른 설명 없이 응답:
{"words":[{"answer":"단어","clue":"힌트 설명"}]}`;

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/'
            + GEMINI_MODEL + ':generateContent?key=' + apiKey;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: 'application/json',
      temperature: 0.9
    }
  };

  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code !== 200) {
    return { ok: false, error: 'Gemini 오류(' + code + '): ' + body.slice(0, 300) };
  }

  let words;
  try {
    const data = JSON.parse(body);
    const text = data.candidates[0].content.parts[0].text;
    const parsed = JSON.parse(text);
    words = parsed.words || [];
  } catch (err) {
    return { ok: false, error: 'Gemini 응답 파싱 실패: ' + String(err) };
  }

  // 정제: NFC 정규화(자모 분리형 방지) → 한글만 남기고, 글자수/중복 필터
  const seen = {};
  const cleaned = [];
  words.forEach(function (w) {
    const a = String(w.answer || '').normalize('NFC').replace(/[^가-힣]/g, '');
    const c = String(w.clue || '').normalize('NFC').trim();
    if (a.length >= 2 && a.length <= maxLen && c && !seen[a]) {
      seen[a] = true;
      cleaned.push({ answer: a, clue: c });
    }
  });

  return { ok: true, words: cleaned };
}


/* ===================== 1-2) Gemini 힌트만 생성 (수동 단어용) ===================== */
function generateClues_(req) {
  const topic = String(req.topic || '').trim();
  const difficulty = req.difficulty || '보통';
  const list = (req.words || []).map(function (s) { return String(s).normalize('NFC').replace(/[^가-힣]/g, ''); })
                                .filter(function (s) { return s.length >= 2; });
  if (!list.length) return { ok: false, error: '힌트를 만들 단어가 없습니다.' };
  const apiKey = geminiKey_();
  if (!apiKey || apiKey.indexOf('여기에') === 0) {
    return { ok: false, error: 'Gemini API 키가 설정되지 않았습니다. (Apps Script 프로젝트 설정 → 스크립트 속성 → GEMINI_API_KEY)' };
  }

  const clueStyle = {
    '쉬움': '아주 직접적이고 친절하게',
    '보통': '적당한 수준으로',
    '어려움': '함축적이고 추론이 필요하게'
  }[difficulty] || '적당한 수준으로';

  // 기존 힌트 예시(있으면) → 어투·길이·난이도를 맞추기 위한 참고
  const examples = (req.examples || [])
    .map(function (e) { return { answer: String(e.answer || '').normalize('NFC'), clue: String(e.clue || '').trim() }; })
    .filter(function (e) { return e.answer && e.clue; })
    .slice(0, 12);
  let exampleBlock = '';
  if (examples.length) {
    exampleBlock = '\n[기존 힌트 예시 — 아래 어투/문장 길이/난이도와 최대한 비슷하게 맞춰주세요]\n'
      + examples.map(function (e) { return '- ' + e.answer + ': ' + e.clue; }).join('\n') + '\n';
  }

  const prompt =
`아래 한국어 단어들에 대한 가로세로 퀴즈용 힌트(설명)를 만들어 주세요.
${topic ? '[퀴즈 주제] ' + topic + '\n' : ''}[힌트 스타일] ${clueStyle} 작성하되, 정답 단어 글자를 그대로 포함하지 말 것.
${exampleBlock}${examples.length ? '위 예시들과 어감·문장 길이·난이도가 자연스럽게 어울리도록 통일성 있게 작성하세요.\n' : ''}
[힌트를 만들 단어]
${list.join(', ')}

반드시 아래 JSON 형식으로만 응답 (단어를 key, 힌트를 value로):
{"clues":{"단어":"힌트"}}`;

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + apiKey;
  const payload = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { responseMimeType: 'application/json', temperature: 0.7 } };
  const resp = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });

  if (resp.getResponseCode() !== 200) return { ok: false, error: 'Gemini 오류: ' + resp.getContentText().slice(0, 200) };
  try {
    const data = JSON.parse(resp.getContentText());
    const parsed = JSON.parse(data.candidates[0].content.parts[0].text);
    return { ok: true, clues: parsed.clues || {} };
  } catch (err) {
    return { ok: false, error: '힌트 응답 파싱 실패: ' + String(err) };
  }
}


/* ===================== 2) 퀴즈 저장 ===================== */
function saveQuiz_(req) {
  const quizName   = String(req.quizName || '').trim();
  const quizData   = req.quizData;  // { gridSize, words:[{answer,clue,direction,row,col,number}] }
  const topic      = req.topic || '';
  const difficulty = req.difficulty || '';
  const folder     = String(req.folder || '').trim();

  if (!quizName) return { ok: false, error: '퀴즈 이름을 입력하세요.' };
  if (!quizData || !quizData.words) return { ok: false, error: '퀴즈 데이터가 없습니다.' };

  const sh = quizSheet_();
  const gridSize = quizData.gridSize
    ? (quizData.gridSize.rows + 'x' + quizData.gridSize.cols)
    : '';

  // quizId 가 있으면 덮어쓰기(수정), 없으면 새로 생성
  let quizId = req.quizId;
  const now = new Date();

  if (quizId) {
    const rowIdx = findRowById_(sh, quizId, 1);
    if (rowIdx > 0) {
      sh.getRange(rowIdx, 2).setValue(quizName);
      sh.getRange(rowIdx, 3).setValue(topic);
      sh.getRange(rowIdx, 4).setValue(gridSize);
      sh.getRange(rowIdx, 5).setValue(difficulty);
      sh.getRange(rowIdx, 7).setValue(JSON.stringify(quizData));
      sh.getRange(rowIdx, 8).setValue(folder);
      bustQuiz_(quizId);   // 수정했으니 이 퀴즈 캐시를 버린다 (학생이 옛 퀴즈를 받으면 안 됨)
      return { ok: true, quizId: quizId, updated: true };
    }
  }

  quizId = 'Q' + now.getTime();
  sh.appendRow([
    quizId, quizName, topic, gridSize, difficulty,
    Utilities.formatDate(now, 'GMT+9', 'yyyy-MM-dd HH:mm:ss'),
    JSON.stringify(quizData), folder
  ]);

  return { ok: true, quizId: quizId, updated: false };
}


/* ===================== 3) 퀴즈 목록 ===================== */
function getQuizList_(req) {
  const sh = quizSheet_();
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, quizzes: [] };

  // 1~6열과 8열만 읽는다. 예전에는 1~8열을 통째로 읽어서
  // 목록에 쓰지도 않는 7열(quizData, 퀴즈당 수 KB짜리 JSON)까지 매번 끌어왔다.
  const base    = sh.getRange(2, 1, last - 1, 6).getValues();
  const folders = sh.getRange(2, 8, last - 1, 1).getValues();
  const quizzes = base.map(function (r, i) {
    return {
      quizId: r[0], quizName: r[1], topic: r[2],
      gridSize: r[3], difficulty: r[4], createdAt: r[5], folder: folders[i][0] || ''
    };
  }).reverse(); // 최신순

  // 각 퀴즈 제출 수 합산
  const counts = submissionCounts_();
  quizzes.forEach(function (q) { q.submissionCount = counts[q.quizId] || 0; });

  return { ok: true, quizzes: quizzes };
}


/* ===================== 4) 퀴즈 불러오기 ===================== */
// 선생님 편집용: 정답 포함 전체 데이터
function getQuizForEdit_(req) {
  const q = loadQuizRow_(req.quizId);
  if (!q) return { ok: false, error: '퀴즈를 찾을 수 없습니다.' };
  return { ok: true, quiz: q };
}

// 학생 풀이용: 정답(answer) 제거, 글자 수만 전달
function getQuizForPlay_(req) {
  // 학생 전원이 똑같이 받아가는 응답이라 통째로 캐시한다.
  // 두 번째 학생부터는 시트를 아예 건드리지 않는다.
  const key = 'play_' + req.quizId;
  const cached = cacheGet_(key);
  if (cached) return { ok: true, quiz: cached };

  const q = loadQuizRow_(req.quizId);
  if (!q) return { ok: false, error: '퀴즈를 찾을 수 없습니다.' };

  const safeWords = (q.quizData.words || []).map(function (w) {
    return {
      number: w.number, clue: w.clue, direction: w.direction,
      row: w.row, col: w.col, length: w.answer.length
      // answer 는 의도적으로 제외 (정답 보호)
    };
  });

  const quiz = {
    quizId: q.quizId,
    quizName: q.quizName,
    gridSize: q.quizData.gridSize,
    words: safeWords
  };
  cachePut_(key, quiz);
  return { ok: true, quiz: quiz };
}

function loadQuizRow_(quizId) {
  if (!quizId) return null;
  const hit = cacheGet_('row_' + quizId);
  if (hit) return hit;
  const sh = quizSheet_();
  const rowIdx = findRowById_(sh, quizId, 1);
  if (rowIdx < 1) return null;
  const r = sh.getRange(rowIdx, 1, 1, 8).getValues()[0];
  let data = {};
  try { data = JSON.parse(r[6]); } catch (e) {}
  const q = {
    quizId: r[0], quizName: r[1], topic: r[2], gridSize: r[3],
    difficulty: r[4], createdAt: r[5], quizData: data, folder: r[7] || ''
  };
  cachePut_('row_' + quizId, q);
  return q;
}


/* ===================== 5) 퀴즈 삭제 ===================== */
function deleteQuiz_(req) {
  const sh = quizSheet_();
  const rowIdx = findRowById_(sh, req.quizId, 1);
  if (rowIdx < 1) return { ok: false, error: '퀴즈를 찾을 수 없습니다.' };
  sh.deleteRow(rowIdx);
  bustQuiz_(req.quizId);
  return { ok: true };
}


/* ===================== 6) 학생 제출 + 채점 ===================== */
/* 채점에 필요한 것(정답 셀맵 + 단어별 셀 목록)만 뽑아 캐시해 둔다.
 * 제출할 때마다 퀴즈 JSON을 시트에서 다시 읽고 파싱하던 것을 없앤다. */
function gradeSheet_(quizId) {
  const key = 'grade_' + quizId;
  const hit = cacheGet_(key);
  if (hit) return hit;

  const q = loadQuizRow_(quizId);
  if (!q) return null;

  const correctMap = {};   // "r-c" -> 정답 글자
  const words = [];        // [{answer, cells:["r-c", ...]}]
  (q.quizData.words || []).forEach(function (w) {
    const cells = [];
    for (var i = 0; i < w.answer.length; i++) {
      var rr = w.direction === 'down' ? w.row + i : w.row;
      var cc = w.direction === 'across' ? w.col + i : w.col;
      correctMap[rr + '-' + cc] = w.answer.charAt(i);
      cells.push(rr + '-' + cc);
    }
    words.push({ answer: w.answer, cells: cells });
  });

  const g = { quizName: q.quizName, correctMap: correctMap, words: words };
  cachePut_(key, g);
  return g;
}

function submitAnswer_(req) {
  const quizId      = req.quizId;
  const studentName = String(req.studentName || '').trim();
  const studentId   = String(req.studentId || '').trim();
  const answers     = req.answers || {};   // { "r-c": "글", ... }
  const timeSpent   = req.timeSpent || '';

  if (!studentName || !studentId) return { ok: false, error: '이름과 학번을 입력하세요.' };

  /* 학생 화면이 응답을 못 받아 같은 제출을 다시 보낸 경우(네트워크 지연·타임아웃),
   * 시트에 같은 기록이 두 줄 쌓이면 안 된다.
   * attemptId 는 '제출 버튼 한 번'에 하나씩 붙는 값이라, 이미 처리한 것이면
   * 새로 쓰지 않고 그때의 채점 결과를 그대로 돌려준다.
   * (학생이 '다시 제출'을 눌러 새로 채점받는 것은 attemptId 가 달라지므로 정상 기록된다) */
  const attemptId = String(req.attemptId || '');
  if (attemptId) {
    const done = cacheGet_('sub_' + attemptId);
    if (done) return done;
  }

  const g = gradeSheet_(quizId);
  if (!g) return { ok: false, error: '퀴즈를 찾을 수 없습니다.' };
  const correctMap = g.correctMap;

  const cellKeys = Object.keys(correctMap);
  const total = cellKeys.length;
  let correct = 0;
  const result = {};  // "r-c" -> true/false (학생에게 채점 표시용)

  cellKeys.forEach(function (key) {
    const ans = (answers[key] || '').trim();
    const ok = ans === correctMap[key];
    if (ok) correct++;
    result[key] = ok;
  });

  const wrong = total - correct;
  const score = total > 0 ? Math.round((correct / total) * 100) : 0;

  // 단어별 정답 여부 (정답률 통계용) — 모든 칸이 맞아야 그 단어를 맞은 것으로 본다
  const wordResults = {};
  g.words.forEach(function (w) {
    var okAll = true;
    for (var i = 0; i < w.cells.length; i++) {
      if ((answers[w.cells[i]] || '').trim() !== correctMap[w.cells[i]]) { okAll = false; break; }
    }
    wordResults[w.answer] = okAll ? 1 : 0;
  });

  // 기록 저장 (동시 제출로 한 번 실패해도 다시 시도한다)
  const sh = subSheet_();
  const now = new Date();
  // 같은 밀리초에 두 명이 제출해도 id가 겹치지 않게 뒤에 난수를 붙인다
  const submissionId = 'S' + now.getTime() + '-' + Math.floor(Math.random() * 1000);
  appendRowSafe_(sh, [
    submissionId, quizId, g.quizName, studentName, studentId,
    Utilities.formatDate(now, 'GMT+9', 'yyyy-MM-dd HH:mm:ss'),
    score, total, correct, wrong, timeSpent, JSON.stringify(wordResults)
  ]);
  try { cache_().remove('subcount'); } catch (e) {}   // 제출 수 캐시만 무효화

  // 채점 결과 + 정답맵 반환 (제출 후이므로 정답 공개 OK)
  const out = {
    ok: true,
    score: score,
    total: total,
    correct: correct,
    wrong: wrong,
    result: result,
    correctMap: correctMap
  };
  if (attemptId) cachePut_('sub_' + attemptId, out, 900);   // 15분간 같은 제출은 재사용
  return out;
}


/* ===================== 7) 제출 기록 조회 (선생님) ===================== */
function getSubmissions_(req) {
  const sh = subSheet_();
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, submissions: [] };

  // 12열(wordResults)은 제출 1건마다 붙는 JSON인데 이 화면에서는 쓰지 않는다.
  // 11열까지만 읽어 시트 읽기량과 응답 크기를 줄인다.
  const rows = sh.getRange(2, 1, last - 1, 11).getValues();
  let subs = rows.map(function (r) {
    return {
      submissionId: r[0], quizId: r[1], quizName: r[2],
      studentName: r[3], studentId: r[4], submittedAt: r[5],
      score: r[6], total: r[7], correct: r[8], wrong: r[9], timeSpent: r[10]
    };
  });

  // 특정 퀴즈만 필터 (옵션)
  if (req.quizId) {
    subs = subs.filter(function (s) { return s.quizId === req.quizId; });
  }
  subs.reverse(); // 최신순

  // 기록이 계속 쌓이면 응답이 무한정 커진다 → 최신 것부터 잘라서 보낸다
  const totalCount = subs.length;
  const limit = Math.max(1, Number(req.limit) || 500);
  if (subs.length > limit) subs = subs.slice(0, limit);

  return { ok: true, submissions: subs, totalCount: totalCount };
}


/* ===================== 7-2) 단어별 정답률 통계 ===================== */
function getWordStats_(req) {
  const quizId = req.quizId;
  const q = loadQuizRow_(quizId);
  if (!q) return { ok: false, error: '퀴즈를 찾을 수 없습니다.' };

  const sh = subSheet_();
  const last = sh.getLastRow();
  const agg = {};        // answer -> {correct, total}
  let subCount = 0;
  if (last >= 2) {
    // 필요한 두 열(quizId, wordResults)만 읽는다. 예전엔 12열 전체를 읽었다.
    const ids = sh.getRange(2, 2,  last - 1, 1).getValues();
    const wrs = sh.getRange(2, 12, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] !== quizId) continue;
      subCount++;
      var wr = {};
      try { wr = JSON.parse(wrs[i][0] || '{}'); } catch (e) { wr = {}; }
      Object.keys(wr).forEach(function (ans) {
        if (!agg[ans]) agg[ans] = { correct: 0, total: 0 };
        agg[ans].total++;
        if (wr[ans]) agg[ans].correct++;
      });
    }
  }

  const words = (q.quizData.words || []).map(function (w) {
    const a = agg[w.answer] || { correct: 0, total: 0 };
    return {
      answer: w.answer, clue: w.clue, direction: w.direction,
      correct: a.correct, total: a.total,
      rate: a.total ? Math.round((a.correct / a.total) * 100) : null
    };
  });

  return { ok: true, words: words, submissionCount: subCount };
}


/* ===================== 공통 유틸 ===================== */
function findRowById_(sh, id, colIdx) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, colIdx, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function submissionCounts_() {
  const hit = cacheGet_('subcount');
  if (hit) return hit;
  const sh = subSheet_();
  const last = sh.getLastRow();
  const counts = {};
  if (last < 2) return counts;
  const ids = sh.getRange(2, 2, last - 1, 1).getValues(); // quizId 열
  ids.forEach(function (r) {
    const id = r[0];
    counts[id] = (counts[id] || 0) + 1;
  });
  cachePut_('subcount', counts);   // 새 제출이 들어오면 submitAnswer_ 가 지운다
  return counts;
}
