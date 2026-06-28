-- v22: 증빙 O 및 발생일 재계산용 월 마감 스냅샷 확장
--
-- 기존 운영 D1은 Pages Functions가 먼저 실행되었을 수도 있고,
-- wrangler migrations가 먼저 실행될 수도 있습니다. SQLite의 ADD COLUMN은
-- IF NOT EXISTS를 지원하지 않아 중복 ALTER가 기존 DB 배포를 막을 수 있으므로,
-- 실제 열 추가는 functions/_lib/schema.js의 ensureSchema()가 PRAGMA 확인 후
-- 필요한 열만 안전하게 추가합니다.
SELECT 1;
