#Requires -Version 7
if(-not (Import-Module 'ImportExcel' -PassThru -ea 'ignore')) {
    Install-Module 'ImportExcel' -Scope CurrentUser -Confirm
}
$Paths = [ordered]@{
    AppRoot = ($AppRoot = Get-Item $PSScriptRoot)
    ExportRoot = Join-Path $AppRoot '../export'
}
if( -not $function:Hr ) {    
    function Hr { "`n`n######`n`n" }
}
$Paths.ExportRoot_CurrentVersion = Join-Path $Paths.ExportRoot 'v1.0.6a'
$Paths.ExportRoot_CurrentVersion = Join-Path $Paths.ExportRoot 'v1.0.7a'

$Paths.Xlsx_Prefabs   = Join-Path $Paths.ExportRoot_CurrentVersion 'prefabs.xlsx'
$Paths.Xlsx_ChangeLog = Join-Path $Paths.ExportRoot_CurrentVersion 'changelog.xlsx'
$Paths.Md_ChangeLog   = Join-Path $Paths.ExportRoot_CurrentVersion 'changelog.md'
$Paths.Csv_ChangeLog  = Join-Path $Paths.ExportRoot_CurrentVersion 'csv/changelog.csv'
$Paths.json_ChangeLog = Join-Path $Paths.ExportRoot_CurrentVersion 'json/changelog.json'

$Paths.Game = [ordered]@{
    'ProgramData_Root' = Join-Path 'C:\Program Files (x86)\Steam\steamapps\common\Microtopia' 'Microtopia_Data'
}
<#
Main entry point.
Todo: refactor
#>

function md.Export.Changelog {
    <#
    .SYNOPSIS
        Exports changelog as '.csv', '.json', and '.md'
    #>
    param()
    $imXl = Import-Excel -path $Paths.Xlsx_ChangeLog -WorksheetName 'Changelog' -ImportColumns 1, 3 -HeaderName 'Version', 'English'
    $imXL
        # using BOM for best results when using Excel csv
        | ConvertTo-Csv
        | Set-Content -Path ($Paths.Csv_ChangeLog) -Encoding utf8BOM

    @( foreach($x in $imXl) {
        '| {0} | {1} |' -f @(
            $x.Version
            $x.English
        )
    }) | Set-Content -Path $Paths.Md_ChangeLog -Encoding utf8

    $imxl
        | Select-Object -Prop Code, English
        | ConvertTo-Json
        | Set-Content -Path $Paths.json_ChangeLog

    # $imXL | Join-String -p { $_.Version, $_.English } -sep "`n"


    # @( foreach($record in $imXl) { $record.'Code', $record.'English' -join ' ' } ) | Join-String -sep "`n"
}

. ( Get-Item -ea 'stop' (Join-Path ($PSScriptRoot) './MdUtils.ps1'))
hr

$pkg = Open-ExcelPackage -Path $Paths.xlsx_Prefabs
# $pkg.Workbook.Worksheets | %{ $_.Name }
$book = $pkg.Workbook
# $sheet = $pkg.workbook.Worksheets
md.Workbook.ListItems $Book
Close-ExcelPackage -ExcelPackage $pkg -NoSave

hr
$pkg = Open-ExcelPackage -Path $Paths.Xlsx_ChangeLog
$book = $pkg.Workbook
# $sheet = $pkg.workbook.Worksheets
# $pkg.Workbook.Worksheets | %{ $_.Name }
md.Workbook.ListItems $Book

$tableName = 'Table1'
$table = $book.Worksheets[1].Tables[ $tableName ]
Close-ExcelPackage -ExcelPackage $pkg -NoSave

# md.Export.Changelog # skip

$Paths | ft -AutoSize
hr;
$imxl | Select -Prop Code, English | ConvertTo-Json
