$PSEmailServer = "srv-ex.gr.guord.local"
$mailto = "de@gr.guord.local"
$mailfrom = "am@gr.guord.local"
$LogparserPatch = "C:\Program Files (x86)\Log Parser 2.2"
# export-dhcpServer удаленно запустить быстро не удалось, поэтому путь локальный и запускать скрипт нужно с DHCP-сервера
$path_audit = "d:\audit" 
$date = Get-Date -format yyyyMMdd

# use $TextFilePlatform = 65001 for unicode
function Convert-csv2xlsx($name_with_path_csv, $name_with_path_xlsx, $TextFilePlatform = 2, $delimiter = ",", $Force = $True, $DeleteSource = $True)
{
  # Create a new Excel workbook with one empty sheet
  $excel = New-Object -ComObject excel.application 
  $excel.visible = $false
  $workbook = $excel.Workbooks.Add(1)
  $worksheet = $workbook.worksheets.Item(1)
  $worksheet.Name = $name_with_path_csv.Split("\")[-1].Split(".")[0]

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

cmd /c $LogparserPatch\LogParser.exe -i:EVT file:terminal.sql -oCodepage:-1

Convert-csv2xlsx $PSScriptRoot\terminal.csv $PSScriptRoot\term.xlsx 65001 -DeleteSource $false

copy-item -path %SystemRoot%\System32\Winevt\Logs\Microsoft-Windows-TerminalServices-LocalSessionManager%4Operational.evtx