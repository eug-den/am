#================================================================================
# зовусь € гордо: audit machine
#
#================================================================================
param(
[parameter(position=0)][string][alias("a")] $audit = "*"
)

function find-option($to_check, $option)
{
  $keys = $to_check.ToCharArray()
  $options = $option.ToCharArray()

  foreach ($key in $keys) {if ($options -contains $key) {return $true}}
  return $false
}

if (!(find-option $audit "*udaslti0"))
{
  write-host 'usage: am.ps1 [[-a|-audit] <*|udaslti>|0]'
  write-host '  -a, -audit выбор режимов работы, можно не указывать'
  write-host '    * все тесты (по умолчанию)'
  write-host '    u список всех пользователей'
  write-host '    d настройки DHCP сервера'
  write-host '    a члены групп типа "*admin*" "*јдмин*"'
  write-host '    s автозапуск на серверах по списку $server_list'
  write-host '    l логи с серверов'
  write-host '    t логи tcplogview'
  write-host '    i лог входов терминального сервера'
  write-host '    0 создание начальных файлов'
  return
}

# конфигурационные переменные START
#
$PSEmailServer = "srv-ex.gr.guord.local"
$mailto = "de@gr.guord.local"
$mailfrom = "am@gr.guord.local"
# export-dhcpServer удаленно запустить быстро не удалось, поэтому путь локальный и запускать скрипт нужно с DHCP-сервера
$path_audit = "d:\audit" # папка, где будут складыватьс€ все выходные файлы
#$path_audit = "\\srv-ad\d$\audit"
$server_list = @("srv-bc", "srv-term", "srv-ad") # список серверов дл€ проверки автозапуска и сбора логов
$DHCP_server_name = "srv-ad"
$Start_server_name = $DHCP_server_name
$LogparserPatch = "C:\Program Files (x86)\Log Parser 2.2"
$term_server_name = "srv-term"
$local_network = '192.168.0.%'
##$psexec_path = d:\system\utl
#
# конфигурационные переменные END

$date = Get-Date -format yyyyMMdd
$summary_body = ""
$summary_error_count = 0

push-location $PSScriptRoot  # выполн€ть будем из папки, где располжен скрипт

<# дл€ удаленного запуска скрипта получим нужные модули
if ($psversionTable.psedition = "Desktop")
{
  $rs = New-PSSession -ComputerName srv-ad.gr.guord.local
  import-module -PSSession $rs -Name ActiveDirectory
  import-module -PSSession $rs -Name dhcpServer
}
#>

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

# use $TextFilePlatform = 65001 for unicode
function Convert-csv2xlsx($name_with_path_csv, $name_with_path_xlsx, $TextFilePlatform = 2, $delimiter = ",", $Force = $True, $DeleteSource = $True)
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
if (find-option $audit "0")
{
  Get-AdUser -filter * -properties passwordLastset | select Name, PasswordLastSet | Export-csv -path $path_audit\user.csv -Encoding OEM

  if (test-path "$path_audit\dhcp.xml") {Remove-Item -Path "$path_audit\dhcp.xml" -Force}
  export-dhcpServer -ComputerName $DHCP_server_name -file "$path_audit\dhcp.xml"

  Get-AdGroup -Filter 'SamAccountName -Like "*admin*" -Or SamAccountName -like "*јдмин*"' | Get-AdGroupMember -Recursive | select-object -uniq |out-file $path_audit\admin.txt -Encoding oem

  forEach ($server in $server_list)
  {
    $outfile = "ar-$server.txt"
    if ($server -eq $Start_server_name) 
    {
      autorunsc.exe -a * -h -s -t -nobanner -accepteula -o $outfile
      $in = "$PSScriptRoot\$outfile"
    }
    else
    {
      PsExec64.exe -accepteula -nobanner -c -e -f \\$server autorunsc.exe -a * -h -s -t -nobanner -accepteula -o $outfile
      $in = "\\$server\admin$\system32\$outfile"
    }
    move -path $in $path_audit -force
  }

}

# „ј—“№ I: проверка на изменение
#

#
# список всех пользователей
if (find-option $audit "*u")
{
  Get-AdUser -filter * -properties passwordLastset | select Name, PasswordLastSet | Export-csv -path $path_audit\user_$date.csv -Encoding OEM
  Compare-Data "user" ".csv"
}

# настройки DHCP сервера
if (find-option $audit "*d")
{
  if (test-path "$path_audit\dhcp_$date.xml") {Remove-Item -Path "$path_audit\dhcp_$date.xml" -Force}
  # export-dhcpServer удаленно не запустилс€
  export-dhcpServer -ComputerName $DHCP_server_name -file "$path_audit\dhcp_$date.xml"
  Compare-Data "dhcp" ".xml"
}

# члены групп типа "*admin*" "*јдмин*"
if (find-option $audit "*a")
{
  Get-AdGroup -Filter 'SamAccountName -Like "*admin*" -Or SamAccountName -like "*јдмин*"' | Get-AdGroupMember -Recursive | select-object -uniq |out-file $path_audit\admin_$date.txt -Encoding oem
  Compare-Data "admin" ".txt"
}

# автозапуск на серверах по списку $server_list
if (find-option $audit "*s")
{
  forEach ($server in $server_list)
  {
    $outfile = "ar-"+$server+"_"+$date+".txt"
    if ($server -eq $Start_server_name) 
    {
      autorunsc.exe -a * -h -s -t -nobanner -accepteula -o $outfile
      $in = "$PSScriptRoot\$outfile"
    }
    else
    {
      PsExec64.exe -accepteula -nobanner -c -e -f \\$server autorunsc.exe -a * -h -s -t -nobanner -accepteula -o $outfile
      $in = "\\$server\admin$\system32\$outfile"
    }
    move -path $in $path_audit -force
    Compare-Data "ar-$server" ".txt" "autorun on \\$server"
  }
}

# суммарный отчет по данному скрипту
if (find-option $audit "*udas")
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
if (find-option $audit "*l")
{
  replace-servers "elog" '\\' '\Security,'
  cmd /c $LogparserPatch\LogParser.exe -stats:OFF -i:EVT file:elog.sql -oCodepage:-1
  Remove-Item -Path elog.sql
  Convert-csv2xlsx $PSScriptRoot\elog.csv $path_audit\elog_$date.xlsx 65001
  Send-Logfile $path_audit\elog_$date.xlsx "All servers logon"
}

# логи tcplogview
if (find-option $audit "*t")
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
if (find-option $audit "i")
{
  copy-item -path \\$term_server_name\c$\windows\System32\Winevt\Logs\Microsoft-Windows-TerminalServices-LocalSessionManager%4Operational.evtx term.evtx
  cmd /c $LogparserPatch\LogParser.exe -stats:OFF -i:EVT file:term.sql
  Remove-Item -Path term.evtx
  Convert-csv2xlsx $PSScriptRoot\term.csv $path_audit\term_$date.xlsx
  Send-Logfile $path_audit\term_$date.xlsx "Terminal server logon"
}

# основна€ процедура END

pop-location # возврат
