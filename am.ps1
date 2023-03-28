#================================================================================
# зовусь € гордо: audit machine
# ver. 1.3.1  06/12/2021
#================================================================================
param(
[parameter(position=0)][string][alias("a")] $am_param = "*"
)

function find-option($to_check, $option)
{
  $keys = $to_check.ToCharArray()
  $options = $option.ToCharArray()

  foreach ($key in $keys) {if ($options -contains $key) {return $true}}
  return $false
}

if (!(find-option $am_param "*1udgs2lti0"))
{
  write-host 'usage: am.ps1 [[-a|-audit] <*|1udgs2lti>|0]'
  write-host '  -a, -audit выбор режимов работы, можно не указывать'
  write-host '    * все тесты (по умолчанию)'
  write-host '    1 тесты 1 типа'
  write-host '    2 тесты 2 типа'
  write-host '    u список всех пользователей'
  write-host '    d настройки DHCP сервера'
  write-host '    g члены групп типа "*admin*" "*јдмин*"'
  write-host '    s автозапуск на серверах по списку $server_list'
  write-host '    l логи с серверов'
  write-host '    t логи tcplogview'
  write-host '    i лог входов терминального сервера'
  write-host '    0 создание начальных файлов'
  return
}

# конфигурационные переменные START
#
$PSEmailServer = "srv-kerio.gr.local"
$mailto = "de@gr.local"
$mailfrom = "am@gr.local"
$path_audit = "\\srv-ad\d$\audit" # путь в сетевом формате!
$LogparserPatch = "C:\Program Files (x86)\Log Parser 2.2"
#$server_list = @("srv-bc") # укороченный список серверов дл€ ускорени€ отладки
$server_list = @("srv-bc", "srv-term") #, "srv-ad")
$DHCP_server_name  = "srv-bc"
$term_server_name  = "srv-term"
$AD_server_name    = "srv-dc"
$Start_server_name = $term_server_name
$local_network = '192.168.0.%'
$isExcel = $false      # установлен модуль import-excel
#$isExcel = $true      # установлен excel

#
# конфигурационные переменные END


<# дл€ удаленного запуска скрипта получим нужные модули
if ($psversionTable.psedition = "Desktop")
{
  $rs = New-PSSession -ComputerName srv-bc.gr.guord.local
#  import-module -PSSession $rs -Name ActiveDirectory
  import-module -PSSession $rs -Name dhcpServer
}
#>

$date = Get-Date -format yyyyMMdd
$summary_body = ""
$summary_error_count = 0

push-location $PSScriptRoot  # выполн€ть будем из папки, где располжен скрипт


#================================================================================
# вс€кие разные функции START
#
function Compare-Data($name_of_log, $ext_of_log, $subj = $name_of_log)
{
  $diff_file = $path_audit + "\" + $name_of_log + "_" + $date + "_diff.txt"
  $file1 = $path_audit + "\" + $name_of_log + $ext_of_log
  $file2 = $path_audit+"\"+$name_of_log+"_"+$date+$ext_of_log

  fc.exe $file1 $file2 > $diff_file

  
  if ( $LastExitCode -eq 0 )
  { 
    $global:summary_body += "+"+$name_of_log + "`n"
    remove-item $diff_file
    remove-item $file2
  }
  else
  {
    $global:summary_error_count++
    $global:summary_body += "-"+$name_of_log + "`n"
    Send-MailMessage 	 `
    -To $mailto 		 `
    -from $mailfrom		 `
    -subject "Audit machine: Warnings in $subj"  `
    -Body "See file in attachment." `
    -Encoding 'UTF8'	 `
    -Attachment $diff_file
  }
}

function Convert-csv2xlsx($name_with_path_csv, $name_with_path_xlsx, $TextFilePlatform = 2, $delimiter = ",", $Force = $True, $DeleteSource = $True)
{
  if ($isExcel)
  {
    Convert-csv2xlsx1 $name_with_path_csv $name_with_path_xlsx $TextFilePlatform  $delimiter $Force  $DeleteSource
  }
  else
  {
    Convert-csv2xlsx2 $name_with_path_csv $name_with_path_xlsx $TextFilePlatform  $delimiter $Force  $DeleteSource
  }
}

function Convert-csv2xlsx2($name_with_path_csv, $name_with_path_xlsx, $TextFilePlatform = 2, $delimiter = ",", $Force = $True, $DeleteSource = $True)
{
  If ($Force -AND (Test-Path -Path $name_with_path_xlsx)) { Remove-Item -Path $name_with_path_xlsx }

  $Excel = (gc $name_with_path_csv) | ConvertFrom-Csv -delimiter $delimiter |
          Export-Excel $name_with_path_xlsx -PassThru -AutoSize -AutoFilter -tablename am -tablestyle medium2 -AutoNameRange -freezetoprow

  Close-ExcelPackage $Excel
  If ($DeleteSource) { Remove-Item -Path $name_with_path_csv }
}

# use $TextFilePlatform = 65001 for unicode
function Convert-csv2xlsx1($name_with_path_csv, $name_with_path_xlsx, $TextFilePlatform = 2, $delimiter = ",", $Force = $True, $DeleteSource = $True)
{
  # Create a new Excel workbook with one empty sheet
  $excel = New-Object -ComObject excel.application 
  $excel.visible = $false
  $workbook = $excel.Workbooks.Add(1)
  $worksheet = $workbook.worksheets.Item(1)
  $worksheet.Name = $name_with_path_csv.Split("\")[-1].Split(".")[0] # часть имени до точки...

  #Build the QueryTables.Add command and reformat the data 
  $TxtConnector = ("TEXT;" + $name_with_path_csv)
  $Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
  $query = $worksheet.QueryTables.item($Connector.name)

  $query.TextFilePlatform = $TextFilePlatform
  $query.TextFileOtherDelimiter = $delimiter
  $query.TextFileParseType  = 1
  $query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
  $query.AdjustColumnWidth = 1

  # Execute & delete the import query
  [void] $query.Refresh()
  $query.Delete()
  [void] $worksheet.UsedRange.EntireColumn.AutoFit()
  $worksheet.Rows.Item(1).Font.Bold = $true
  [void] $worksheet.Rows.Item(1).AutoFilter()
 
  [void] $workSheet.Activate()
  $worksheet.Application.ActiveWindow.SplitRow = 1;
  $workSheet.Application.ActiveWindow.FreezePanes = $true;

  # Save & close the Workbook as XLSX.
  If ($Force -AND (Test-Path -Path $name_with_path_xlsx)) { Remove-Item -Path $name_with_path_xlsx }
  $workbook.SaveAs($name_with_path_xlsx, 51)
  $excel.quit()
  [void] [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
  If ($DeleteSource) { Remove-Item -Path $name_with_path_csv }
}

function Send-Logfile($name_of_Attachment, $subj)
{
  Send-MailMessage 	 `
  -To $mailto 		 `
  -from $mailfrom		 `
  -subject "Audit machine: Logging: $subj"  `
  -Body "See file in attachment." `
  -Encoding 'UTF8'	 `
  -Attachment $name_of_Attachment
}

function replace-servers($name, $replS, $replE)
{
  $replace_to = ""
  forEach ($server in $global:server_list) {$replace_to += $replS+$server+$replE}
  $a =  $replace_to.length-1 
  $replace_to = $replace_to.remove($a)

  (gc "$name.sql.template") -replace '@SERVERS@', $replace_to |sc "$name.sql"
}
#
# функции END

#================================================================================
# начало выполнени€ основной процедуры
#
# „ј—“№ 0: создание начальных файлов
#
if (find-option $am_param "0")
{
  invoke-command -ComputerName $AD_server_name -scriptblock {Get-AdUser -filter * -properties passwordLastset | select Name, PasswordLastSet | Export-csv -path $using:path_audit\user.csv -Encoding OEM}

  if (test-path "$path_audit\dhcp.xml") {Remove-Item -Path "$path_audit\dhcp.xml" -Force}
  invoke-command -ComputerName $DHCP_server_name -scriptblock {export-dhcpServer -ComputerName $Using:DHCP_server_name -file "$Using:path_audit\dhcp.xml"}

  invoke-command -ComputerName $AD_server_name -scriptblock {Get-AdGroup -Filter 'SamAccountName -Like "*admin*" -Or SamAccountName -like "*јдмин*"' | Get-AdGroupMember -Recursive | select-object -uniq |out-file $Using:path_audit\admin.txt -Encoding oem}

  forEach ($server in $server_list)
  {
    $outfile = $path_audit + "\ar-" + $server + ".txt"
    if ($server -eq $Start_server_name) 
    {
      autorunsc.exe -a * -h -s -t -nobanner -accepteula -o $outfile
    }
    else
    {
      invoke-command -ComputerName $server {autorunsc.exe -a * -h -s -t -nobanner -accepteula} | out-file $outfile -force -encoding ascii
    }
  }
}

# „ј—“№ I: проверка на изменение
#

#
# список всех пользователей
if (find-option $am_param "*u1")
{
  invoke-command -ComputerName $AD_server_name -scriptblock {Get-AdUser -filter * -properties passwordLastset | select Name, PasswordLastSet | Export-csv -path $using:path_audit\user_$using:date.csv -Encoding OEM}
  Compare-Data "user" ".csv"
}

# настройки DHCP сервера
if (find-option $am_param "*d1")
{
  if (test-path "$path_audit\dhcp_$date.xml") {Remove-Item -Path "$path_audit\dhcp_$date.xml" -Force}
  invoke-command -ComputerName $DHCP_server_name -scriptblock {export-dhcpServer -ComputerName $Using:DHCP_server_name -file "$Using:path_audit\dhcp_$Using:date.xml"}
  Compare-Data "dhcp" ".xml"
}

# члены групп типа "*admin*" "*јдмин*"
if (find-option $am_param "*g1")
{
  invoke-command -ComputerName $AD_server_name -scriptblock {Get-AdGroup -Filter 'SamAccountName -Like "*admin*" -Or SamAccountName -like "*јдмин*"' | Get-AdGroupMember -Recursive | select-object -uniq |out-file $Using:path_audit\admin_$Using:date.txt -Encoding oem}
  Compare-Data "admin" ".txt"
}

# автозапуск на серверах по списку $server_list
if (find-option $am_param "*s1")
{
  forEach ($server in $server_list)
  {
    $outfile = $path_audit + "\ar-"+$server+"_"+$date+".txt"
    if ($server -eq $Start_server_name) 
    {
      autorunsc.exe -a * -h -s -t -nobanner -accepteula -o $outfile
    }
    else
    {
      invoke-command -ComputerName $server {autorunsc.exe -a * -h -s -t -nobanner -accepteula} | out-file $outfile -force -encoding ascii
    }
    Compare-Data "ar-$server" ".txt" "autorun on \\$server"
  }
}

# суммарный отчет по первой части
if (find-option $am_param "*1udas")
{
  Send-MailMessage 	 `
  -To $mailto 		 `
  -from $mailfrom		 `
  -subject "Audit machine: Summary $summary_error_count" `
  -Body $summary_body `
  -Encoding 'UTF8'
}

#================================================================================
#
# „ј—“№ II: логи в EXCEL и на почту
#

#
# логи с серверов
if (find-option $am_param "*l2")
{
  replace-servers "elog" '\\' '\Security,'
  cmd /c $LogparserPatch\LogParser.exe -stats:OFF -i:EVT file:elog.sql -oCodepage:-1
  Remove-Item -Path elog.sql
  Convert-csv2xlsx $PSScriptRoot\elog.csv $path_audit\elog_$date.xlsx 65001
  Send-Logfile $path_audit\elog_$date.xlsx "All servers logon"
}

# логи tcplogview
if (find-option $am_param "*t2")
{
  forEach ($server in $server_list)
  {
    $in = '\\'+$server+'\admin$\TcpLogView.csv'
    move -path $in $PSScriptRoot\tlv_$server.csv -force
  }
  replace-servers "tlv" 'tlv_' '.csv,'
  (gc 'tlv.sql') -replace '@NETWORK@', $local_network |sc 'tlv.sql'
  cmd /c $LogparserPatch\LogParser.exe -stats:OFF -i:CSV file:tlv.sql
  Remove-Item -Path tlv.sql
  forEach ($server in $server_list) {Remove-Item -Path $PSScriptRoot\tlv_$server.csv}
  Convert-csv2xlsx $PSScriptRoot\tlv.csv $path_audit\tlv_$date.xlsx
  Send-Logfile $path_audit\tlv_$date.xlsx "tcplogview"
}

# лог входов терминального сервера
if (find-option $am_param "i")
{
  copy-item -path \\$term_server_name\c$\windows\System32\Winevt\Logs\Microsoft-Windows-TerminalServices-LocalSessionManager%4Operational.evtx term.evtx
  cmd /c $LogparserPatch\LogParser.exe -stats:OFF -i:EVT file:term.sql
  Remove-Item -Path term.evtx
  Convert-csv2xlsx $PSScriptRoot\term.csv $path_audit\term_$date.xlsx
  Send-Logfile $path_audit\term_$date.xlsx "Terminal server logon"
}

# основна€ процедура END

pop-location # возврат
