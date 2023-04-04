SELECT
 timegenerated as Date, 
 extract_token(strings, 0, '|') as user,
 extract_token(strings, 2, '|') as sourceip
into term.csv
FROM term.evtx
WHERE EventID = 21 
and TO_DATE( timegenerated ) >= SUB(TO_LOCALTIME( SYSTEM_DATE() ), TIMESTAMP('2', 'd')) 
/*and sourceip not like '192.168.0.%'*/
ORDER BY Date DESC
