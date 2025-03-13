#Requires -Version 7
if(-not (Import-Module 'ImportExcel' -PassThru -ea 'silentlycontinue')) {
    Install-Module 'ImportExcel' -Scope CurrentUser -Confirm
}
$Paths = [ordered]@{
    AppRoot = ($AppRoot = Get-Item $PSScriptRoot)
    ExportRoot = Join-Path $AppRoot '../export'
}
if($true) {
    # disable: Import-Module (Join-Path $PSScriptRoot 'Grouping.psm1') -ea 'stop'

    # todo: refactor as module
    $toImport = (Join-Path ($PSScriptRoot) './MdUtils.ps1')
    "DotSrc: `"$toImport`"" | Out-Host
    . ( Get-Item -ea 'stop' (Join-Path ($PSScriptRoot) './MdUtils.ps1'))
}

if( -not $function:Hr ) {
    function Hr { "`n`n######`n`n" }
}
$paths.Log = Join-Path $Paths.ExportRoot '../log.log'
$Paths.ExportRoot_CurrentVersion = Join-Path $Paths.ExportRoot 'v1.0.6a'
$Paths.ExportRoot_CurrentVersion = Join-Path $Paths.ExportRoot 'v1.0.7a'

md.EnsureSubdirsExist -Path $paths.ExportRoot_CurrentVersion -verbose

$Paths.Xlsx_Biome = Join-Path $Paths.ExportRoot_CurrentVersion 'biome.xlsx'
$Paths.Raw_Biome  = md.GetRawPath $Paths.Xlsx_Biome

$Paths.Xlsx_Prefabs = Join-Path $Paths.ExportRoot_CurrentVersion 'prefabs.xlsx'
$Paths.Raw_Prefabs  = md.GetRawPath $Paths.Xlsx_Prefabs

$Paths.Xlsx_Instinct = Join-Path $Paths.ExportRoot_CurrentVersion 'Instinct.xlsx'
$Paths.Raw_Instinct  = md.GetRawPath $Paths.Xlsx_Instinct

$Paths.Xlsx_ChangeLog               = Join-Path $Paths.ExportRoot_CurrentVersion 'changelog.xlsx'
$Paths.Md_ChangeLog                 = Join-Path $Paths.ExportRoot_CurrentVersion 'changelog.md'
$Paths.Csv_ChangeLog                = Join-Path $Paths.ExportRoot_CurrentVersion 'csv/changelog.csv'
$Paths.json_ChangeLog               = Join-Path $Paths.ExportRoot_CurrentVersion 'json/changelog.json'
$Paths.json_Biome_Objects           = Join-Path $Paths.ExportRoot_CurrentVersion 'json/biome-objects.json'
$Paths.json_Biome_Plants            = Join-Path $Paths.ExportRoot_CurrentVersion 'json/biome-plants.json'
$Paths.json_Biome_Plants_ColumnDesc = Join-Path $Paths.ExportRoot_CurrentVersion 'json/biome-plants-column-desc.json'
$Paths.json_Biome_Objects_Expanded  = Join-Path $Paths.ExportRoot_CurrentVersion 'json/biome-objects-expanded.json'
$Paths.json_WorkbookSchema          = Join-Path $Paths.ExportRoot_CurrentVersion 'json/workbook-schema.json'
$Paths.xlsx_WorkbookSchema          = Join-Path $Paths.ExportRoot_CurrentVersion 'workbook-schema.xlsx'

# $Paths.Game = [ordered]@{
#     'ProgramData_Root' = Join-Path 'C:\Program Files (x86)\Steam\steamapps\common\Microtopia' 'Microtopia_Data'
# }
<#
    Main entry point. refactor
#>
'export schemas for all *.xlsx' | Write-Host -fg 'gray60'


# never cache
Remove-Item $Paths.Xlsx_Biome -ea 'Ignore'
Clear-Content -path $Paths.Log -ea Ignore
md.Export.Biome.Biome_Objects -Paths $Paths -Verbose
md.Export.Biome.Plants -Paths $Paths -Verbose
# <nyi>: md.Export.Instinct -Paths $Paths -Verbose

$Paths.Log
    | Join-String -f 'See log for a list of changed files: "{0}"'
    | Write-Host -fg 'skyblue'

md.Export.Changelog -Verbose -Path $Paths

md.Export.WorkbookSchema
md.Export.WorkbookSchema.Xlsx -Paths $Paths -Verbose
md.Export.Readme.FileListing -Path $Paths.ExportRoot_CurrentVersion

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
