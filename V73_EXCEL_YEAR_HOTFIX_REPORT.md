# v73 Excel year hotfix

- v71 전체 파일 구조를 유지하고 엑셀 생성 오류만 수정했습니다.
- `buildFinalTemplateWorkbook()`에서 계산한 `year/monthNo`를 `buildContext()`에 전달하도록 수정했습니다.
- `year is not defined`가 엑셀 생성 중 발생할 수 있는 원인을 제거했습니다.
- 브라우저/Cloudflare 캐시 회피를 위해 `app.js`와 `final-template.js` 캐시 키를 `v=73-excel-year-hotfix`로 변경했습니다.
- D1, API, migrations, 인력 매칭, 연차 누적 관리, 자동 휴무 로직은 변경하지 않았습니다.
