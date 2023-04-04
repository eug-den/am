#requires -version 4.0

Function Convert-CSVtoXLS {
<#
.SYNOPSIS
This function converts a CSV file to an Excel workbook.
.DESCRIPTION
Convert-CSVtoXLS converts a csv file to a Excel workbook.
The first line of the CSV file is turned into a filtering header.
Excel must be installed on the computer.
.EXAMPLE
PS C:\> Convert-CSVtoXLS myfile.csv
myfile.csv will be converted to myfile.xslx
.EXAMPLE
PS C:\> foreach ($file in (ls *.csv)) { Convert-CSVtoXLS $file }
All csv files in the current folder will be converted.
.NOTES
NAME        :  Convert-CSVtoXLS
VERSION     :  0.8
LAST UPDATED:  20/02/2015
AUTHOR      :  Xavier Plantefève
.LINK
http://xavier.plantefeve.fr
.INPUTS
Either the file name of the string to be converted (System.String) or the file object (System.IO.FileInfo)
.OUTPUTS
No output.
#>

    [CmdletBinding(DefaultParameterSetName='fromstring')]
    Param(
        # Path of the CSV file to be converted.
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeline=$True,ParameterSetName='fromstring')]
        [string]$Path,
        # CSV file to be converted. Accepts pipeline.
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeline=$True,ParameterSetName='fromfile')]
        [System.IO.FileInfo]$File,
        # The source CSV file will be deleted.
        [switch]$DeleteSource,
        # If used, the Excel worksheet will be saved to the 97-2003 format.
        [switch]$LegacyFormat,
        # Delimiter used in the CSV. Defaults to 'SemiColon'
        [String][ValidateSet('Comma','Semicolon','Space','Tab')]$Delimiter = 'Semicolon',
        # Provides a way to use a non-standard delimiter char. Voids the -Delimiter parameter.
        [char]$DelimiterChar,
        # A name for the resulting excel worksheet. Defaults to the file base name.
        [string]$Name,
        # Full path (including filename) of the resulting Excel file.
        [string]$DestinationPath,
        # Allows overwriting of the Excel file.
        [switch]$Force
    )

    if ( $Path ) {
        $File = Get-Item -Path $Path
    }
    # We set $Path even if it exists, to translate it to a full path.
    $Path = $File.FullName

    # Format constants: https://msdn.microsoft.com/en-us/library/office/ff198017.aspx
    If ($LegacyFormat) {
        $XLfilext = '.xls'
        $FileFormat = 56
    } else {
        $XLfilext = '.xlsx'
        $FileFormat = 51
    }
    
    $excel = New-Object -ComObject excel.application
    $excel.Visible = $PSBoundParameters['Verbose']
    
    # Workbook creation
    $workbooks = $excel.Workbooks.Add()
    $worksheets = $workbooks.Worksheets
    $worksheets.Item(3).delete()
    $worksheets.Item(2).delete()
    $worksheet = $worksheets.Item(1)
    if ($Name) {
        $worksheet.Name = $Name
    } else {
        $worksheet.Name = $File.BaseName
    }

    # CSV Import.
    $TxtConnector = ("TEXT;${Path}")
    $CellRef = $worksheet.Range('A1')
    $Connector = $worksheet.QueryTables.add($TxtConnector,$CellRef)
    if ($DelimiterChar) {
        $worksheet.QueryTables.Item($Connector.Name).TextFileOtherDelimiter = $DelimiterChar
    } else {
        $worksheet.QueryTables.Item($Connector.Name)."TextFile${Delimiter}Delimiter" = $true
    }
    $worksheet.QueryTables.Item($Connector.Name).TextFileParseType = 1
    [void] $worksheet.QueryTables.Item($Connector.Name).Refresh()
    [void] $worksheet.QueryTables.Item($Connector.Name).delete()


    If ($worksheet.Cells.Item(1,1).Text -like '#TYPE*') {
        [void] $worksheet.Rows.Item(1).Delete()
    }

    # A bit of formatting, because we're shallow and like when things look nice.
    # (I'm joking, this is for the managers to be happy)
    [void] $worksheet.UsedRange.EntireColumn.AutoFit()
    $worksheet.Rows.Item(1).Font.Bold = $true
    [void] $worksheet.Rows.Item(1).AutoFilter()
    
    [void] $workSheet.Activate()
    $worksheet.Application.ActiveWindow.SplitRow = 1;
    $workSheet.Application.ActiveWindow.FreezePanes = $true;
    
    # We save the file and quit.
    if (!$DestinationPath) {
        $DestinationPath = "$($File.DirectoryName)\$($File.BaseName)${XLfilext}"
    } else {
        If ((Split-Path -Path $DestinationPath) -in '.','') {
            $DestinationPath = "$($pwd.ProviderPath)\$(Split-Path -Path $DestinationPath -Leaf)"
        }
    }
    
    If ($Force -AND (Test-Path -Path $DestinationPath)) { Remove-Item -Path $DestinationPath }
    $workbooks.SaveAs($DestinationPath,$FileFormat)
    $excel.quit()
    [void] [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    
    If ($DeleteSource) { Remove-Item -Path $Path }
} #function

# creates an alias for the function
Set-Alias -Name csv2xls -Value Convert-CSVtoXLS