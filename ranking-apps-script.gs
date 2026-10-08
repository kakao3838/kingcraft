// 제로 클락 공유 랭킹 서버 (Google Apps Script 웹 앱)
// 설정: 구글 시트 새로 만들기 → 확장 프로그램 → Apps Script → 이 코드 붙여넣기 →
//       배포 → 새 배포 → 유형 "웹 앱", 실행 계정 "나", 액세스 권한 "모든 사용자" → 나온 /exec 주소를
//       index.html 의 RANKING_URL 에 넣는다.
// 코드를 고친 뒤에는 "배포 관리"에서 새 버전으로 다시 배포해야 반영된다.

const SHEET_NAME = 'ranking';
const MAX_SCORE = 300;   // 이보다 큰 점수는 조작으로 보고 무시
const TOP_N = 10;

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(['name', 'score', 'time']);
  return sh;
}

function cleanName_(s) {
  return String(s || '')
    .replace(/[\u0000-\u001f<>]/g, '')
    .replace(/^[=+\-@\s]+/, '') // 시트 수식 주입 방지
    .trim()
    .slice(0, 10);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1)
    .filter(r => r[0] !== '' && Number.isFinite(Number(r[1])))
    .map(r => ({ name: String(r[0]), score: Number(r[1]), time: Number(r[2]) || 0 }))
    .sort((a, b) => b.score - a.score || a.time - b.time)
    .slice(0, TOP_N)
    .map(r => ({ name: r.name, score: r.score }));
  return json_(rows);
}

function doPost(e) {
  let data;
  try { data = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false }); }
  const name = cleanName_(data.name);
  const score = Math.floor(Number(data.score));
  if (!name || !Number.isFinite(score) || score < 1 || score > MAX_SCORE) return json_({ ok: false });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === name) {
        if (score > Number(values[i][1])) sh.getRange(i + 1, 2, 1, 2).setValues([[score, Date.now()]]);
        return json_({ ok: true });
      }
    }
    sh.appendRow([name, score, Date.now()]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
