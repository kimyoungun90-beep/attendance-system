# v67 Auto Blank Dayoff Both Sheets

- Fixes 출/계 미입력 variants not being treated as blank dayoff candidates.
- Applies 휴무(공백) autofill to both 상담사근태 and 상담사근태_관리자반영 sheets.
- Reapplies autofill after manager finalization overrides so manager sheet does not leave under-allowance 출/계 미입력 rows.
- Excludes manager-aware auto 휴무(공백) rows from 출근증빙·휴무확인.
- No D1 schema change.
