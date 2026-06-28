-- v19: 기존 부여 기록의 발생일이 비어 있는 경우 안전한 기본값을 채웁니다.
UPDATE attendance_leave_grants_v5
SET occurrence_date = COALESCE(NULLIF(criterion_date, ''), grant_month || '-01')
WHERE occurrence_date IS NULL OR TRIM(occurrence_date) = '';
