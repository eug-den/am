SELECT
 [Event Time],[Event Type],[Local Address],[Remote Address],[Remote Host Name],
 [Local Port],[Remote Port],[Process ID],[Process Name],[Process Path],[Process User]
INTO tlv.csv
FROM TLV_srv-ad.csv, TLV_srv-term.csv
where [Remote Address] not like '192.168.0.%'
ORDER BY [Event Time] DESC
