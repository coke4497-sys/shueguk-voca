/* ===== 어휘 테스트 목록 (manifest) =====
   - weeks  : 학년당 주차 수 (1~25)
   - grades : 학년 코드 / 화면 표시 라벨
   - tests  : 제작이 끝나 "링크를 켤" 테스트만 등록.
              키는 "학년코드-주차" (예: m1-1), 값은 { title }.
              여기에 등록된 것만 링크 발송/학생 응시가 활성화됩니다.
              새 주차를 만들면 data/<코드>-<주차>.json 파일을 추가하고
              아래 tests 에 한 줄만 등록하면 끝.
*/
window.VOCA = {
  baseUrl: "https://coke4497-sys.github.io/shueguk-voca/test.html",
  weeks: 25,
  grades: [
    { code: "m1", label: "중1" },
    { code: "m2", label: "중2" },
    { code: "m3", label: "중3" },
    { code: "h1", label: "고1" },
    { code: "h2", label: "고2" },
    { code: "h3", label: "고3" }
  ],
  tests: {
    "m1-1": { title: "친구·우정 편" }
  }
};
