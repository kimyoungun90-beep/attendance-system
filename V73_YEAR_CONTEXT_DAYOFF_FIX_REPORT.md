# v73 year context and manager-confirmed dayoff fix

- v71 전체 파일 구조를 유지한 상태에서 `public/index.html`, `public/assets/app.js`, `public/assets/final-template.js`만 수정했습니다.
- 엑셀 생성 중 `buildContext()` 내부에서 `year/monthNo`가 전달되지 않아 발생할 수 있는 `year is not defined` 오류를 수정했습니다.
- Cloudflare/브라우저 캐시 회피를 위해 `app.js`와 `final-template.js` 캐시 키를 `v=73-year-context-dayoff-fix`로 변경했습니다.
- 공백·미입력 날짜를 앞에서부터 임의로 `휴무(공백)` 처리하지 않도록 자동 휴무 배정 호출을 제거했습니다.
- 매니저 월마감 수정본에서 명시된 휴무/휴가 값만 `상담사근태_관리자반영`에 확정 반영되도록 유지했습니다.
- D1 마이그레이션, API, 중간 저장, 인력 매칭, 연차 누적 관리 파일은 삭제하거나 변경하지 않았습니다.
