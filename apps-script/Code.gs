/**
 * 호스피스간호실무 평가 결과 저장 서버 (Google Apps Script)
 * 1) 이 코드를 스프레드시트의 [확장 프로그램 > Apps Script]에 붙여 넣습니다.
 * 2) 아래 ADMIN_PASSWORD를 원하는 비밀번호로 바꾼 뒤 setup 함수를 한 번 실행합니다.
 * 3) [배포 > 새 배포 > 웹 앱] 실행 사용자: 나, 액세스 권한: 모든 사용자 로 배포합니다.
 */
const ADMIN_PASSWORD = "여기에-교수님-비밀번호";
const SHEET = "responses";
const HEAD = ["제출시각", "평가", "학번", "이름", "점수", "만점", "응답", "오답문항", "제출ID"];

function setup() {
  PropertiesService.getScriptProperties().setProperty("ADMIN_KEY", ADMIN_PASSWORD);
  sheet_();
  Logger.log("준비 완료. 이제 웹 앱으로 배포하세요.");
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight("bold");
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clean_(v, max) { return String(v == null ? "" : v).replace(/^[=+\-@]/, "'").slice(0, max); }

function doGet() { return out_({ ok: true, service: "hospice-eval" }); }

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: "bad_json" }); }

  if (body.action === "submit") {
    if (!body.sid || !body.name || !body.key) return out_({ ok: false, error: "missing" });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = sheet_();
      // 같은 제출ID가 이미 있으면 중복 저장하지 않음(재전송 대비)
      const last = sh.getLastRow();
      if (body.subId && last > 1) {
        const ids = sh.getRange(2, 9, last - 1, 1).getValues().flat();
        if (ids.indexOf(body.subId) !== -1) return out_({ ok: true, dup: true });
      }
      sh.appendRow([new Date(), clean_(body.key, 10), clean_(body.sid, 20), clean_(body.name, 30),
        Number(body.score) || 0, Number(body.total) || 0, clean_(body.answers, 60), clean_(body.wrong, 120), clean_(body.subId, 40)]);
    } finally { lock.releaseLock(); }
    return out_({ ok: true });
  }

  if (body.action === "results") {
    const key = PropertiesService.getScriptProperties().getProperty("ADMIN_KEY");
    if (!key || body.password !== key) return out_({ ok: false, error: "auth" });
    const sh = sheet_();
    const last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2, 1, last - 1, HEAD.length).getValues() : [];
    return out_({ ok: true, rows: rows.map(function (r) {
      return { at: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), key: r[1], sid: String(r[2]).replace(/^'/, ""), name: String(r[3]).replace(/^'/, ""),
               score: r[4], total: r[5], answers: String(r[6]), wrong: String(r[7]), subId: r[8] };
    }) });
  }
  return out_({ ok: false, error: "unknown_action" });
}
