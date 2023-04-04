# audit machine

# конфигурационные переменные START
#
$PSEmailServer = "srv-ex.gr.guord.local"
$mailto = "de@gr.guord.local"
$mailfrom = "am@gr.guord.local"
# export-dhcpServer удаленно запустить быстро не удалось, поэтому путь локальный и запускать скрипт нужно с DHCP-сервера
$path_audit = "d:\audit" 
#$path_audit = "\\srv-ad\d$\audit"
$server_list = @("srv-term")#, "srv-ad", "srv-bc", "srv-ex")
##$psexec_path = d:\system\utl
#
# конфигурационные переменные END


$date = Get-Date -format yyyyMMdd
$summary_body = ""
$summary_error_count = 0

<#
if ($psversionTable.psedition = "Desktop")
{
  $rs = New-PSSession -ComputerName srv-ad.gr.guord.local
  import-module -PSSession $rs -Name ActiveDirectory
  import-module -PSSession $rs -Name dhcpServer
}
#>

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

#<#
Get-AdUser -filter * -properties passwordLastset | select Name, PasswordLastSet | Export-csv -path $path_audit\user_$date.csv -Encoding OEM
Compare-Data "user" ".csv"

if (test-path "$path_audit\dhcp_$date.xml") {Remove-Item -Path "$path_audit\dhcp_$date.xml" -Force}
# export-dhcpServer удаленно не запустился
export-dhcpServer -ComputerName "srv-ad.gr.guord.local" -file "$path_audit\dhcp_$date.xml"
Compare-Data "dhcp" ".xml"

Get-AdGroup -Filter 'SamAccountName -Like "*admin*" -Or SamAccountName -like "*Админ*"' | Get-AdGroupMember -Recursive | select-object -uniq |out-file $path_audit\admin_$date.txt -Encoding oem
Compare-Data "admin" ".txt"
#>


forEach ($server in $server_list)
{
  $outfile = "ar-"+$server+"_"+$date+".txt"
#  PsExec64.exe -accepteula -nobanner -c -e -f \\$server autorunsc.exe -a l -h -s -t -nobanner -o $outfile
#  $in = '\\'+$server+'\admin$\system32\'+$outfile
#  move -path $in "$path_audit\$outfile" -force
  Compare-Data "ar-$server" ".txt" "autorun on \\$server"
}

Send-MailMessage 	 `
-To $mailto 		 `
-from $mailfrom		 `
-subject "Audit machine: Summary $summary_error_count" `
-Body $summary_body `
-Encoding 'UTF8'
