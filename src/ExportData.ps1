#Requires -Version 7
if(-not (Import-Module 'ImportExcel' -PassThru -ea 'silentlycontinue')) {
    Install-Module 'ImportExcel' -Scope CurrentUser -Confirm
}
$Paths = [ordered]@{
    AppRoot = ($AppRoot = Get-Item $PSScriptRoot)
    ExportRoot = Join-Path $AppRoot '../export'
}
if($true) { 
    # todo: refactor as module 
    $toImport = (Join-Path ($PSScriptRoot) './MdUtils.ps1')
    "DotSrc: `"$toImport`"" | Out-Host
    . ( Get-Item -ea 'stop' (Join-Path ($PSScriptRoot) './MdUtils.ps1'))
}

if( -not $function:Hr ) {    
    function Hr { "`n`n######`n`n" }
}
$Paths.ExportRoot_CurrentVersion = Join-Path $Paths.ExportRoot 'v1.0.6a'
$Paths.ExportRoot_CurrentVersion = Join-Path $Paths.ExportRoot 'v1.0.7a'

md.EnsureSubdirsExist -Path $paths.ExportRoot_CurrentVersion -verbose 

$Paths.Xlsx_Biome = Join-Path $Paths.ExportRoot_CurrentVersion 'biome.xlsx'
$Paths.Raw_Biome  = md.GetRawPath $Paths.Xlsx_Biome

$Paths.Xlsx_Prefabs   = Join-Path $Paths.ExportRoot_CurrentVersion 'prefabs.xlsx'
$Paths.Raw_Prefabs   = md.GetRawPath $Paths.Xlsx_Prefabs

$Paths.Xlsx_ChangeLog = Join-Path $Paths.ExportRoot_CurrentVersion 'changelog.xlsx'
$Paths.Md_ChangeLog   = Join-Path $Paths.ExportRoot_CurrentVersion 'changelog.md'
$Paths.Csv_ChangeLog  = Join-Path $Paths.ExportRoot_CurrentVersion 'csv/changelog.csv'
$Paths.json_ChangeLog = Join-Path $Paths.ExportRoot_CurrentVersion 'json/changelog.json'
$Paths.json_WorkbookSchema = Join-Path $Paths.ExportRoot_CurrentVersion 'json/workbook-schema.json'

# $Paths.Game = [ordered]@{
#     'ProgramData_Root' = Join-Path 'C:\Program Files (x86)\Steam\steamapps\common\Microtopia' 'Microtopia_Data'
# }




<#
    Main entry point.
    Todo: refactor
#>
# hr

'export excel XLSX' | out-host ; #  Write-Host -fg 'orange'
$pkg = Open-ExcelPackage -Path $Paths.Raw_Biome
$book = $pkg.Workbook
md.Workbook.ListItems $Book
$sheets = $pkg.workbook.Worksheets

# detect column counts
$curSheet = $pkg.Workbook.Worksheets['Biome Objects']

Close-ExcelPackage -ExcelPackage $pkg -NoSave


return

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
