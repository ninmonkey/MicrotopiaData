#Requires -Version 7
if(-not (Import-Module 'ImportExcel' -PassThru -ea 'ignore')) {
    Install-Module 'ImportExcel' -Scope CurrentUser -Confirm
}
$Paths = [ordered]@{
    AppRoot = ($AppRoot = Get-Item $PSScriptRoot)
    ExportRoot = Join-Path $AppRoot '../export'
}

$Paths.ExportRoot_CurrentVersion = Join-Path $Paths.ExportRoot 'v1.0.6a'

$Paths.Xlsx_Prefabs   = Join-Path $Paths.ExportRoot_CurrentVersion 'prefabs.xlsx'
$Paths.Xlsx_ChangeLog = Join-Path $Paths.ExportRoot_CurrentVersion 'changelog.xlsx'

<#
Main entry point.
Todo: refactor
#>

. ( Get-Item -ea 'stop' (Join-Path ($PSScriptRoot) './MdUtils.ps1'))

$pkg = Open-ExcelPackage -Path $Paths.xlsx_Prefabs
# $pkg.Workbook.Worksheets | %{ $_.Name }
$book = $pkg.Workbook
# $sheet = $pkg.workbook.Worksheets
hr
md.Workbook.ListItems $Book
Close-ExcelPackage -ExcelPackage $pkg

$pkg = Open-ExcelPackage -Path $Paths.Xlsx_ChangeLog
$book = $pkg.Workbook
# $sheet = $pkg.workbook.Worksheets
# $pkg.Workbook.Worksheets | %{ $_.Name }
hr
md.Workbook.ListItems $Book
Close-ExcelPackage -ExcelPackage $pkg

hr;


$Paths | ft -AutoSize
