-- 이전 최종본 상담사근태 시트의 수기 이월 잔여를 월 마감 원본에 저장합니다.
ALTER TABLE attendance_monthly_employee_facts ADD COLUMN imported_opening_balance_present INTEGER NOT NULL DEFAULT 0;
ALTER TABLE attendance_monthly_employee_facts ADD COLUMN imported_opening_substitute REAL NOT NULL DEFAULT 0;
ALTER TABLE attendance_monthly_employee_facts ADD COLUMN imported_opening_compensation REAL NOT NULL DEFAULT 0;
