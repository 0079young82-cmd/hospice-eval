// ① 구글 Apps Script 웹 앱 주소 (…/exec 로 끝나야 합니다)
const API_URL = "https://script.google.com/macros/s/AKfycbwyi7kx2aBFlrJemmCvrDp1zczlZnkFD0z8fvyWDwNjzlssK6RXDi32hjnny4mqGQY1aA/exec";
// ② 화면에 보이는 과목 이름
const COURSE_LABEL = "2026-2 호스피스간호실무 · 간호학과 4학년";
// ③ (선택) 주차별 마감 시각. 비워 두면 언제 제출해도 '출석'으로 봅니다.
//    마감 뒤 제출은 교수용 출석부에 '지각'으로 표시됩니다. 형식: "2026-09-07T23:59"
const DEADLINES = {
  dx: "", w1: "", w2: "", w3: "", w4: "", w5: "", w6: "", w7: "",
  w8: "", w9: "", w10: "", w11: "", w12: "", w13: "", w14: ""
};
// ④ 너무 빨리 푼 제출을 표시하는 기준 (문항당 초). 5문항 × 8초 = 40초 미만이면 ⚠ 표시
const FAST_SEC_PER_ITEM = 8;
