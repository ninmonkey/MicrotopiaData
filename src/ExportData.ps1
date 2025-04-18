﻿#Requires -Version 7
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
$Paths.ExportRoot_CurrentVersion = Join-Path $Paths.ExportRoot 'v1.0.8'

md.EnsureSubdirsExist -Path $paths.ExportRoot_CurrentVersion -verbose

$Paths.Xlsx_Biome = Join-Path $Paths.ExportRoot_CurrentVersion 'biome.xlsx'
$Paths.Raw_Biome  = md.GetRawPath $Paths.Xlsx_Biome

$Paths.Xlsx_Prefabs = Join-Path $Paths.ExportRoot_CurrentVersion 'prefabs.xlsx'
$Paths.Raw_Prefabs  = md.GetRawPath $Paths.Xlsx_Prefabs

$Paths.Xlsx_Instinct = Join-Path $Paths.ExportRoot_CurrentVersion 'Instinct.xlsx'
$Paths.Raw_Instinct  = md.GetRawPath $Paths.Xlsx_Instinct

$Paths.Xlsx_TechTree = Join-Path $Paths.ExportRoot_CurrentVersion 'techtree.xlsx'
$Paths.Raw_TechTree  = md.GetRawPath $Paths.Xlsx_TechTree

$Paths.Xlsx_Loc = Join-Path $Paths.ExportRoot_CurrentVersion 'loc.xlsx'
$Paths.Raw_Loc  = md.GetRawPath $Paths.Xlsx_Loc

$Paths.Xlsx_Sequences = Join-Path $Paths.ExportRoot_CurrentVersion 'sequences.xlsx'
$Paths.Raw_Sequences  = md.GetRawPath $Paths.Xlsx_Sequences

$Paths.Xlsx_ChangeLog                = Join-Path $Paths.ExportRoot_CurrentVersion 'changelog.xlsx'
$Paths.Md_ChangeLog                  = Join-Path $Paths.ExportRoot_CurrentVersion 'changelog.md'
$Paths.Csv_ChangeLog                 = Join-Path $Paths.ExportRoot_CurrentVersion 'csv/changelog.csv'
$Paths.json_ChangeLog                = Join-Path $Paths.ExportRoot_CurrentVersion 'json/changelog.json'

$Paths.Json_Crusher_Output = Join-Path $Paths.ExportRoot_CurrentVersion 'json/crusher-output.json'
$Paths.csv_Crusher_Output  = Join-Path $Paths.ExportRoot_CurrentVersion 'csv/crusher-output.csv'

$Paths.json_Prefabs_Buildings        = Join-Path $Paths.ExportRoot_CurrentVersion 'json/prefabs-buildings.json'
$Paths.json_Prefabs_FactoryRecipes   = Join-Path $Paths.ExportRoot_CurrentVersion 'json/prefabs-factoryrecipes.json'
$Paths.json_Prefabs_AntCastes        = Join-Path $Paths.ExportRoot_CurrentVersion 'json/prefabs-antcastes.json'
$Paths.json_Prefabs_Pickups          = Join-Path $Paths.ExportRoot_CurrentVersion 'json/prefabs-pickups.json'
$Paths.json_Prefabs_Trails           = Join-Path $Paths.ExportRoot_CurrentVersion 'json/prefabs-trails.json'
$Paths.json_Prefabs_PickupCategories = Join-Path $Paths.ExportRoot_CurrentVersion 'json/prefabs-pickupcategories.json'
$Paths.json_Prefabs_Hunger           = Join-Path $Paths.ExportRoot_CurrentVersion 'json/prefabs-hunger.json'
$Paths.json_Prefabs_StatusEffects    = Join-Path $Paths.ExportRoot_CurrentVersion 'json/prefabs-statuseffects.json'


$Paths.json_Sequences_Tutorial     = Join-Path $Paths.ExportRoot_CurrentVersion 'json/sequences-tutorial.json'
$Paths.json_Sequences_TutorialOld = Join-Path $Paths.ExportRoot_CurrentVersion 'json/sequences-tutorial-old.json'
$Paths.json_Sequences_Events       = Join-Path $Paths.ExportRoot_CurrentVersion 'json/sequences-events.json'


$Paths.json_Biome_Objects              = Join-Path $Paths.ExportRoot_CurrentVersion 'json/biome-objects.json'
$Paths.json_Biome_Objects_Expanded     = Join-Path $Paths.ExportRoot_CurrentVersion 'json/biome-objects-expanded.json'
$Paths.json_Biome_Plants               = Join-Path $Paths.ExportRoot_CurrentVersion 'json/biome-plants.json'
$Paths.json_Biome_Plants_ColumnDesc    = Join-Path $Paths.ExportRoot_CurrentVersion 'json/biome-plants-column-desc.json'

$Paths.json_Loc_legend                 = Join-Path $Paths.ExportRoot_CurrentVersion 'json/loc-legend.json'

$Paths.json_Loc_UI           = Join-Path $Paths.ExportRoot_CurrentVersion 'json/loc-ui.json'
$Paths.json_Loc_Objects      = Join-Path $Paths.ExportRoot_CurrentVersion 'json/loc-objects.json'
$Paths.json_Loc_tutorial     = Join-Path $Paths.ExportRoot_CurrentVersion 'json/loc-tutorial.json'
$Paths.json_Loc_instinct     = Join-Path $Paths.ExportRoot_CurrentVersion 'json/loc-instinct.json'
$Paths.json_Loc_techtree     = Join-Path $Paths.ExportRoot_CurrentVersion 'json/loc-techtree.json'
$Paths.json_Loc_credits      = Join-Path $Paths.ExportRoot_CurrentVersion 'json/loc-credits.json'
$Paths.json_Loc_Achievements = Join-Path $Paths.ExportRoot_CurrentVersion 'json/loc-achievements.json'

$Paths.json_TechTree_ResearchRecipes   = Join-Path $Paths.ExportRoot_CurrentVersion 'json/techtree-researchrecipes.json'
$Paths.json_TechTree_TechTree_Expanded = Join-Path $Paths.ExportRoot_CurrentVersion 'json/techtree-techtree-expanded.json'
$Paths.json_TechTree_TechTree          = Join-Path $Paths.ExportRoot_CurrentVersion 'json/techtree-techtree.json'

$Paths.json_WorkbookSchema           = Join-Path $Paths.ExportRoot_CurrentVersion 'json/workbook-schema.json'
$Paths.xlsx_WorkbookSchema           = Join-Path $Paths.ExportRoot_CurrentVersion 'workbook-schema.xlsx'

$Paths.Template_Readme = Join-Path $Paths.AppRoot './readme.template.md'
$Paths.Markdown_RootReadme = join-path $paths.AppRoot '../readme.md'

$build = $null
$Build ??= [ordered]@{ # auto 'show' certain files. nullish op lets you override defaults
    AutoOpen = [ordered]@{
        Biome_Objects            = $true
        Biome_Objects_Expanded   = $true # $true
        Biome_Plants             = $true # $true
        Loc                      = $true # $true
        Prefabs_Crusher          = $true
        TechTree_ResearchRecipes = $true # $true
        TechTree_TechTree        = $true # $true
        Prefabs                  = $true # $true
        Sequences                = $true
        WorkbookSchema           = $true
    }
    Export = [ordered]@{
        # main
        # Changelog                  = $false # $true # $false
        # Biome_Objects              = $false # $true # $false
        # Biome_Objects_Expanded     = $false # $false # $false1
        # Biome_Plants               = $false # $true # $false
        # Loc                        = $true # $true # $true
        # Prefabs_Crusher            = $false # $true # $false
        # Prefabs                    = $false # $true # $false
        # TechTree_ResearchRecipes   = $false # $true # $false
        # TechTree_TechTree          = $false # $true # $false
        # TechTree_TechTree_Expanded = $false # $false # $false
        # WorkbookSchema             = $false # $true # $false


        # all on
        Changelog                  = $true # true
        Biome_Objects              = $true # true
        Biome_Objects_Expanded     = $true # true
        Biome_Plants               = $true # true
        Loc                        = $true # true
        Prefabs_Crusher            = $true # true
        Prefabs                    = $true # true
        TechTree_ResearchRecipes   = $true # true
        TechTree_TechTree          = $true # true
        TechTree_TechTree_Expanded = $true
        WorkbookSchema             = $true
        Sequences                  = $True

        # only schema
        # Changelog                  = $false
        # Biome_Objects              = $false
        # Biome_Objects_Expanded     = $false
        # Biome_Plants               = $false
        # Loc                        = $false
        # Prefabs_Crusher            = $false
        # Prefabs                    = $false
        # TechTree_ResearchRecipes   = $false
        # TechTree_TechTree          = $false
        # TechTree_TechTree_Expanded = $false
        # WorkbookSchema             = $true
    }
}
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
# <nyi>: md.Export.Instinct -Paths $Paths -Verbose
$Paths.Log
    | Join-String -f 'See log for a list of changed files: "{0}"'
    | Write-Host -fg 'skyblue'

if( $Build.Export.Loc ) { # run loc first, others depend on it
    md.Export.Loc -Paths $Paths -Verbose
}
# main entry point for the script
if( $Build.Export.Biome_Objects ) {
    md.Export.Biome.Biome_Objects -Paths $Paths -Verbose
}
if( $Build.Export.Biome_Plants ) {
    md.Export.Biome.Plants -Paths $Paths -Verbose
}

if($Build.Export.TechTree_TechTree) {
    Remove-Item $Paths.Xlsx_TechTree -ea 'Ignore'
    md.Export.TechTree.TechTree -Paths $Paths -Verbose
}
if($Build.Export.Changelog) {
    md.Export.Changelog -Verbose -Path $Paths
}
if($Build.Export.Prefabs) {
    Remove-Item $Paths.Xlsx_Prefabs -ea 'Ignore'
    md.Export.Prefabs.Prefabs -Paths $Paths -Verbose
}
if($Build.Export.Sequences) {
    Remove-Item $Paths.Xlsx_Sequences -ea 'Ignore'
    md.Export.Sequences -Paths $Paths -Verbose
}

# final exports. Ran last to iterate all new exports
if($Build.Export.WorkbookSchema) {
    md.Export.WorkbookSchema -verbose # -Force #  -Paths $Paths -Verbose # -Force
    md.Export.WorkbookSchema.Xlsx -Paths $Paths -Verbose
}

md.Export.Readme.FileListing -Path $Paths

# log config at tail of log
$build
    | ConvertTo-Json -Depth 1
    | Join-string -op "`nBuildConfig: `n" -sep "`n"
    | Add-Content -path $paths.Log

$Paths
    | ConvertTo-Json -Depth 0
    | Join-string -op "`nPaths: `n" -sep "`n"
    | Add-Content -path $paths.Log

$Paths.Log
    | Join-String -f 'Done. See log for a list of changed files: "{0}"'
    | Write-Host -fg 'skyblue'
return

# $pkg = Open-ExcelPackage -Path $Paths.Xlsx_Prefabs
# $rows = ImportExcel -pkg $Pkg
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
