SELECT
eventid, 
case eventid
 WHEN 4624 THEN 'logon ok' 
 WHEN 4647 THEN 'logoff' 
 WHEN 4625 THEN 'logon failure' 
end as eventName,
TO_DATE(timegenerated) as Date, 
TO_TIME(timegenerated) as Time, 

case eventid
 WHEN 4647 THEN extract_token(Strings, 1, '|' )
 ELSE extract_token(Strings, 5, '|' )
end as LogonName, 

case eventid
 WHEN 4625 THEN extract_token(Strings, 19, '|' ) 
 ELSE extract_token(Strings, 18, '|' ) 
end as LogonIP, 

case extract_token(Strings, 8, '|' ) 
 WHEN '2' THEN 'interactive' 
 WHEN '3' THEN 'network' 
 WHEN '4' THEN 'batch' 
 WHEN '5' THEN 'service' 
 WHEN '7' THEN 'unlocked workstation' 
 WHEN '8' THEN 'network logon using a cleartext password' 
 WHEN '9' THEN 'impersonated logons' 
 WHEN '10' THEN 'remote access' 
 ELSE extract_token(Strings, 8, '|' ) 
end as LogonType, 

case eventid
 WHEN 4624 THEN 
  case extract_token(Strings, 1, '|' ) 
   WHEN 'SERVER$' THEN 'logon' 
   ELSE extract_token(Strings, 1, '|' ) 
  end 
 ELSE ''
end as Type /*, 
Strings */
INTO elog.csv
/*
FROM \\srv-ad\Security
*/
FROM \\srv-bc\Security, \\srv-ad\Security, \\srv-term\Security, \\srv-sql\Security
WHERE 
  EventID IN (4624; 4647; 4625)  AND LogonType not like 'service' and
  TO_DATE( TimeGenerated ) >= SUB(TO_LOCALTIME( SYSTEM_DATE() ), TIMESTAMP('2', 'd'))
 ORDER BY Date, Time DESC