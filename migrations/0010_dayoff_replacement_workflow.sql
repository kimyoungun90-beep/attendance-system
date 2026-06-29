-- v34: 기본휴무 초과분의 대체휴무·보상휴가 자동대체 결과 저장
ALTER TABLE attendance_monthly_summaries_v3 ADD COLUMN dayoff_replacement_used REAL NOT NULL DEFAULT 0;
ALTER TABLE attendance_monthly_summaries_v3 ADD COLUMN dayoff_replacement_shortage REAL NOT NULL DEFAULT 0;
ALTER TABLE attendance_monthly_summaries_v3 ADD COLUMN dayoff_replacement_substitute_used REAL NOT NULL DEFAULT 0;
ALTER TABLE attendance_monthly_summaries_v3 ADD COLUMN dayoff_replacement_compensation_used REAL NOT NULL DEFAULT 0;
