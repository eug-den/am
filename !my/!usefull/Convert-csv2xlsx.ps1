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

<#  тест функции  Convert-csv2xlsx

$path_csv = $PSScriptRoot 
$path_xlsx  = $path_csv
$name_with_path_csv = "$path_csv\test.csv"
#Convert-csv2xlsx $path_csv\test.csv $path_xlsx\test.xlsx 65001 -DeleteSource $false

#>
