# V61 중간 저장 이력 선택·삭제 기능

## 목적
- 7/10, 7/20처럼 월중 확인 완료본을 여러 번 저장하고, 다음 분석 때 원하는 저장본을 선택해서 이어 분석할 수 있게 함.
- 잘못 저장한 중간 확인본은 웹앱에서 삭제 가능하게 함.

## 변경 사항
1. 웹앱 왼쪽 파일 비교 설정 영역에 `중간 저장 관리` 추가
   - 최신 저장본 자동 사용
   - 특정 기준일 저장본 선택
   - 저장값 사용 안 함
   - 선택 저장본 삭제

2. `중간 확인 저장` 동작 변경
   - 기존처럼 `attendance_closures` 월 마감 테이블을 덮어쓰지 않음
   - 신규 `attendance_midmonth_snapshots`, `attendance_midmonth_snapshot_items`에 이력으로 저장
   - 같은 기준일을 다시 저장하면 해당 기준일 저장본만 교체

3. 분석 시 저장본 선택 적용
   - 저장본 선택: 해당 기준일까지 저장값 사용, 이후 날짜 새 분석
   - 최신 자동: 가장 최근 기준일 저장본 사용
   - 저장값 사용 안 함: 중간 저장값 무시
   - 전체 재분석 체크: 저장본 선택과 관계없이 모든 중간 저장값 무시

4. 삭제 기능
   - 선택한 중간 저장본과 날짜별 저장값을 D1에서 삭제
   - 다른 기준일 저장본은 유지

## D1 변경
- 신규 마이그레이션: `migrations/0012_midmonth_snapshots.sql`
- 런타임 `ensureSchema`에도 동일 테이블 자동 생성 로직 포함

## 신규 테이블
- `attendance_midmonth_snapshots`
- `attendance_midmonth_snapshot_items`

## 검증
- `public/assets/app.js` 문법 검사 통과
- `functions/api/midmonth-snapshots.js` 문법 검사 통과
- `functions/api/substitute-balances.js` 문법 검사 통과
- `functions/api/closures.js` 문법 검사 통과
- `functions/_lib/schema.js` 문법 검사 통과
