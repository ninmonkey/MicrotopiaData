#Requires -Version 7
if(-not (Import-Module 'ImportExcel' -PassThru -ea 'silentlycontinue')) {
    Install-Module 'ImportExcel' -Scope CurrentUser -Confirm
}
$Paths = [ordered]@{
    AppRoot = ($AppRoot = Get-Item $PSScriptRoot)
    ExportRoot = Join-Path $AppRoot '../export'
}
if($true) {
    Import-Module (Join-Path $PSScriptRoot 'Grouping.psm1') -ea 'stop'

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
$Paths.json_Biome_Objects = Join-Path $Paths.ExportRoot_CurrentVersion 'json/biome-objects.json'
$Paths.json_Biome_Objects_Expanded = Join-Path $Paths.ExportRoot_CurrentVersion 'json/biome-objects-expanded.json'
$Paths.json_WorkbookSchema = Join-Path $Paths.ExportRoot_CurrentVersion 'json/workbook-schema.json'

# $Paths.Game = [ordered]@{
#     'ProgramData_Root' = Join-Path 'C:\Program Files (x86)\Steam\steamapps\common\Microtopia' 'Microtopia_Data'
# }




<#
    Main entry point.
    Todo: refactor
#>
# hr

md.Export.WorkbookSchema -verbose

'export excel XLSX' | out-host ; #  Write-Host -fg 'orange'
$pkg = Open-ExcelPackage -Path $Paths.Raw_Biome
$book = $pkg.Workbook
md.Workbook.ListItems $Book
$sheets = $pkg.workbook.Worksheets

# detect column counts
$curSheet = $pkg.Workbook.Worksheets['Biome Objects']

remove-item $Paths.Xlsx_Biome -ea 'Ignore'

$importExcel_Splat = @{
    ExcelPackage  = $pkg
    WorksheetName = 'Biome Objects'
}
$rows =  Import-Excel @importExcel_Splat
# skip empty and non-data rows
$rows = @(
    $rows
        | ? { -not [string]::IsNullOrWhiteSpace( $_.CODE ) }
        | ? { $_.CODE -notmatch '^\s*//' }
        | ? { $_.CODE -notmatch '^\s*\?+\s*$' } # skip "???"
)

$exportExcel_Splat = @{
    InputObject   = @( $rows )
    Path          = $Paths.Xlsx_Biome
    Show          = $false
    WorksheetName = 'Biome_Objects'
    TableName     = 'Biome_Objects_Data'
    TableStyle    = 'Light5'
    AutoSize      = $True
}

Export-Excel @exportExcel_Splat

# ($src -split ',\s+').ForEach({
#   $segs = $_ -split '\s+', 2
#   [pscustomobject]@{ Name = $segs[0]; Quantity = $segs[1]  }
# })

# $record.PICKUPS     = @( $record.PICKUPS -split ',\s*' )



# json specific transforms
$sort_splat = @{
    Property = 'title', 'code', 'exchange_types'
}

$forJson = @(
    $Rows | %{
        $record = $_
        $record = md.Convert.BlankPropsToEmpty $Record
        $record = md.Convert.KeyNames $Record
        # $record = md.Convert.TruthyProps $Record

        # coerce blankables into empty strings for json
        $record.'pickups'             = md.Parse.IngredientsFromCsv $record.'pickups'
        $record.'exchange_types'      = md.Parse.ItemsFromList $record.'exchange_types'
        $record.'unclickable'         = md.Parse.Checkbox $record.'unclickable'
        $record.'trails_pass_through' = md.Parse.Checkbox $record.'trails_pass_through'
    }
) | Sort-Object @sortObjectSplat

$forJson
    | ConvertTo-Json -depth 9
    | Set-Content -path $Paths.Json_Biome_Objects # -Confirm

$Paths.Json_Biome_Objects | Join-String -f 'wrote: "{0}"' | write-host -fg 'gray50'


# also emit expanded records
$forJson = @(
    $Rows | %{
        $record = $_
        $expandProp = 'exchange_types'
        if( $record.$expandProp.count -gt 0) {
            $record.$expandProp | %{
                $curType            = $_
                $newObj             = $record | Select-Object -Prop *
                $newObj.$expandProp = $curType
                $newObj
            }
        } else {
            $record
        }
    }
)| Sort-Object @sortObjectSplat


$forJson
    | ConvertTo-Json -depth 9
    | Set-Content -path $Paths.json_Biome_Objects_Expanded # -Confirm

$Paths.json_Biome_Objects_Expanded | Join-String -f 'wrote: "{0}"' | write-host -fg 'gray50'

$exportExcel_Splat = @{
    InputObject   = @( $forJson )
    Path          = $Paths.Xlsx_Biome
    Show          = $false
    WorksheetName = 'Biome_Objects_Expanded'
    TableName     = 'Biome_Objects_Expanded_Data'
    TableStyle    = 'Light5'
    Title         = 'From Json'
    AutoSize      = $True
}

Export-Excel @exportExcel_Splat

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
