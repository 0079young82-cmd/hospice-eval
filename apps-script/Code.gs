/**
 * 호스피스간호실무 진단·형성평가 서버 (Google Apps Script) — 2판
 * 저장하는 것: 로그인 기록, '강의 들었습니다' 확인 시각, 평가 시작·제출 시각, 소요 시간,
 *             문항별 풀이 시간, 화면 이탈 횟수, 응답과 점수, 응시 회차, 사용 기기
 *
 * 설치/수정 방법
 * 1) 스프레드시트 [확장 프로그램 > Apps Script]에 이 코드 전체를 붙여 넣고 저장
 * 2) ADMIN_PASSWORD를 바꾼 뒤 setup 함수 실행 (처음 한 번, 비밀번호를 바꿀 때도 다시 실행)
 * 3) [배포 > 배포 관리 > 연필 > 버전: 새 버전 > 배포]  (처음이면 [새 배포 > 웹 앱])
 *    실행 사용자: 나 / 액세스 권한: 모든 사용자  ← 'Google 계정이 있는 모든 사용자'가 아님!
 * 4) (선택) 'students' 탭에 학번·이름 명단을 넣으면 명단에 있는 학생만 로그인할 수 있습니다.
 */
const ADMIN_PASSWORD = "여기에-교수님-비밀번호";

const SUB = "submissions";
const SUB_HEAD = ["제출시각(서버)", "평가", "학번", "이름", "점수", "만점", "응답", "오답문항",
  "강의확인시각", "시작시각", "제출시각(학생기기)", "소요시간(초)", "문항별시간(초)", "화면이탈횟수",
  "응시회차", "기기", "제출ID"];
const LOG = "log";
const LOG_HEAD = ["시각(서버)", "학번", "이름", "이벤트", "평가", "상세", "기기"];
const ROSTER = "students";

function setup() {
  PropertiesService.getScriptProperties().setProperty("ADMIN_KEY", ADMIN_PASSWORD);
  sheet_(SUB, SUB_HEAD); sheet_(LOG, LOG_HEAD); sheet_(ROSTER, ["학번", "이름"]);
  Logger.log("준비 완료. 결과 시트 주소: " + ss_().getUrl());
  Logger.log("이제 [배포 > 새 배포 > 웹 앱]으로 배포하세요.");
}

// 스프레드시트에서 연 스크립트면 그 시트를, 따로 만든 스크립트면 결과용 시트를 새로 만들어 씁니다.
function ss_() {
  const a = SpreadsheetApp.getActiveSpreadsheet();
  if (a) return a;
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty("SHEET_ID");
  if (id) return SpreadsheetApp.openById(id);
  const s = SpreadsheetApp.create("호스피스간호실무_평가결과");
  p.setProperty("SHEET_ID", s.getId());
  return s;
}

function sheet_(name, head) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head); sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight("bold");
  }
  return sh;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function s_(v, max) { return String(v == null ? "" : v).trim().replace(/^[=+\-@]/, "'").slice(0, max || 200); }
function n_(v) { const x = Number(v); return isFinite(x) ? x : ""; }
function iso_(v) { return v instanceof Date ? v.toISOString() : String(v || ""); }

function roster_() {
  const sh = sheet_(ROSTER, ["학번", "이름"]); const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 2).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({ sid: String(r[0]).trim(), name: String(r[1]).trim() }));
}
function inRoster_(sid, name) {
  const r = roster_(); if (!r.length) return true;           // 명단이 비어 있으면 누구나
  return r.some(x => x.sid === String(sid).trim() && x.name === String(name).trim());
}
function log_(b, ev, detail) {
  sheet_(LOG, LOG_HEAD).appendRow([new Date(), s_(b.sid, 20), s_(b.name, 30), ev, s_(b.key, 10), s_(detail, 500), s_(b.ua, 150)]);
}

function doGet() { return out_({ ok: true, service: "hospice-eval", version: 2, time: new Date().toISOString() }); }

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: "bad_json" }); }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (b.action === "ping") return out_({ ok: true });

    if (b.action === "login") {
      if (!b.sid || !b.name) return out_({ ok: false, error: "missing" });
      if (!inRoster_(b.sid, b.name)) { log_(b, "login_denied", "명단에 없음"); return out_({ ok: false, error: "not_in_roster" }); }
      log_(b, "login", ""); return out_({ ok: true });
    }

    if (b.action === "event") {             // 강의 확인, 평가 시작 등
      if (!b.sid) return out_({ ok: false, error: "missing" });
      log_(b, s_(b.event, 30), b.detail || ""); return out_({ ok: true });
    }

    if (b.action === "submit") {
      if (!b.sid || !b.name || !b.key) return out_({ ok: false, error: "missing" });
      const sh = sheet_(SUB, SUB_HEAD); const last = sh.getLastRow();
      if (b.subId && last > 1) {             // 같은 제출 재전송은 한 번만 저장
        const ids = sh.getRange(2, 17, last - 1, 1).getValues().flat();
        if (ids.indexOf(b.subId) !== -1) return out_({ ok: true, dup: true });
      }
      sh.appendRow([new Date(), s_(b.key, 10), s_(b.sid, 20), s_(b.name, 30), n_(b.score), n_(b.total),
        s_(b.answers, 60), s_(b.wrong, 120), s_(b.checkedAt, 40), s_(b.startedAt, 40), s_(b.clientAt, 40),
        n_(b.durationSec), s_(b.perQ, 200), n_(b.leaveCount), n_(b.attempt), s_(b.ua, 150), s_(b.subId, 40)]);
      log_(b, "submit", "점수 " + b.score + "/" + b.total + ", " + b.durationSec + "초");
      return out_({ ok: true });
    }

    if (b.action === "results") {
      const key = PropertiesService.getScriptProperties().getProperty("ADMIN_KEY");
      if (!key || b.password !== key) return out_({ ok: false, error: "auth" });
      const read = (name, head) => { const sh = sheet_(name, head), l = sh.getLastRow();
        return l > 1 ? sh.getRange(2, 1, l - 1, head.length).getValues() : []; };
      const subs = read(SUB, SUB_HEAD).map(r => ({ at: iso_(r[0]), key: r[1], sid: String(r[2]).replace(/^'/, ""),
        name: String(r[3]).replace(/^'/, ""), score: r[4], total: r[5], answers: String(r[6]), wrong: String(r[7]),
        checkedAt: String(r[8]), startedAt: String(r[9]), clientAt: String(r[10]), dur: r[11], perQ: String(r[12]),
        leave: r[13], attempt: r[14], ua: String(r[15]), subId: r[16] }));
      const logs = read(LOG, LOG_HEAD).map(r => ({ at: iso_(r[0]), sid: String(r[1]).replace(/^'/, ""),
        name: String(r[2]).replace(/^'/, ""), ev: r[3], key: r[4], detail: String(r[5]), ua: String(r[6]) }));
      return out_({ ok: true, subs: subs, logs: logs, roster: roster_() });
    }
    return out_({ ok: false, error: "unknown_action" });
  } catch (err) {
    return out_({ ok: false, error: "server", message: String(err) });
  } finally { try { lock.releaseLock(); } catch (e2) {} }
}
