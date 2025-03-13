#Requires -Version 7

Import-Module Pansies -ea 'stop'
Import-Module ImportExcel -ea 'stop'

function md.Log.WroteFile {
    <#
    .SYNOPSIS
        Write host? Log to logfile? writes to "temp:\last.log" by default
    .EXAMPLE
        > $Paths.Json_Biome_Objects | md.Log.WroteFile
    #>
    param(
        [Parameter(ValueFromPipeline)]
        $InputObject
    )
    process {
        $msg = $InputObject | Join-String -f 'wrote: "{0}"'
        $msg | write-host -fg 'gray50'
        $msg | Add-Content -Path ($Paths.Log ?? 'temp:\last.log')
    }
}
function md.Workbook.ListItems {
    <#
    .SYNOPSIS
        Takes a workbook, echo the list of tables defined
    #>
    param(
        $InputObject
    )
    if( $book -isnot [OfficeOpenXml.ExcelWorkbook] ) {
        "Unexpected value type: $( $InputObject.GetType() )" | Write-Warning
    }
    [OfficeOpenXml.ExcelWorkbook] $book = $InputObject
    $joinStringSplat = @{
        Separator    = ', '
        Property     = 'Name'
        SingleQuote  = $True
        OutputPrefix = "WorkBook has Sheets: "
    }

    $book.Worksheets
        | Join-String @joinStringSplat

    $Book.Worksheets | % {
        [OfficeOpenXml.ExcelWorksheet] $curSheet = $_

        $joinStringSplat = @{
            Separator    = ', '
            Property     = 'Name'
            SingleQuote  = $True
            OutputPrefix = "ws $( $curSheet.Name ) has Tables: "
        }

        $curSheet.Tables
            | Join-String @joinStringSplat
    }

}

function md.GetRawPath {
    <#
    .Synopsis
        Path or object, modifies path: 'foo.xlsx' => 'foo-raw.xlsx'
    #>
    [CmdletBinding()]
    param(
        [object] $Path )

    $File = Get-Item -ea 'ignore' $Path
    if( -not $File ) {
        $File = [System.IO.FileInfo]::new( $Path )
    }
    if( -not $File) {
        throw "Unhandled path type: $( $Path )"
    }
    $rawPath = Join-Path $File.DirectoryName "$( $File.BaseName )-raw.xlsx"
    $rawPath
}

function md.Export.Changelog {
    <#
    .SYNOPSIS
        Exports changelog as '.csv', '.json', and '.md'
    #>
    [CmdletBinding()]
    param(
        # Paths hashtable
        [Parameter(Mandatory)] $Paths,

        # always write a fresh export
        [ValidateScript({throw 'nyi'})]
        [switch] $Force
    )
    $PSCmdlet.MyInvocation.MyCommand.Name
        | Join-String -f 'Enter: {0}' | Write-verbose

    # $rawPath = $Paths.xlsx_Changelog
    # $rawFullJoin-Path $rawPath.DirectoryName "$( $_.baseName )-raw.xlsx"
    $curOutput = $Paths.Xlsx_ChangeLog
    $rawSrc    = md.GetRawPath $curOutput

    "md.Export.Changelog => Parse: $( $rawSrc ), Output: $( $curOutput )" | Write-Host -fg 'gray60' -bg 'gray30'

    $importExcelSplat = @{
        Path          = $rawSrc
        WorksheetName = 'Changelog'
        ImportColumns = 1,         3
        HeaderName    = 'Code', 'English'
    }
    $regex = @{
        isVersion  = '\s*v\d+'
        dashPrefix = '^\s*?-\s*'
    }

    $imXl = Import-Excel @importExcelSplat

    $curVersionGroup =
        # $VersionGroup = $imxl[1].English
        $imXl | ? English -Match $regex.isVersion | Select -First 1 | % English

    $forJson = @(
        $imXL | %{
            $record = $_

            if( $record.English -eq 'English' ) { return }

            if( [string]::IsNullOrWhiteSpace( $record.English ) ) { return }

            if( $record.English -match $regex.isVersion ) {
                $curVersionGroup = $record.English
                return
            }

            [pscustomobject]@{
                Version = $curVersionGroup
                Code    = $record.Code
                English = $record.English -replace $regex.dashPrefix, ''
            }
        }
    )

    $importExcelSplat.Path | md.Log.WroteFile

    # tip: using BOM for best results when using Excel csv
    $forJson
        | ConvertTo-Csv
        | Set-Content -Path ($Paths.Csv_ChangeLog) -Encoding utf8BOM

    $Paths.Csv_ChangeLog | md.Log.WroteFile

    @(
        '| Version | Code | English | '
        '| - | - | - |'
        @(foreach($x in $forJson) {
            '| {0} | {1} | {2} |' -f @(
                $x.Version
                $x.Code
                $x.English
            )}
        )
    ) | Set-Content -Path $Paths.Md_ChangeLog -Encoding utf8

    $Paths.Md_ChangeLog | md.Log.WroteFile

    $forJson
        # | Select-Object -Prop Code, English
        | ConvertTo-Json
        | Set-Content -Path $Paths.json_ChangeLog

    $Paths.json_ChangeLog | md.Log.WroteFile

    # $imXL | Join-String -p { $_.Version, $_.English } -sep "`n"


    # @( foreach($record in $imXl) { $record.'Code', $record.'English' -join ' ' } ) | Join-String -sep "`n"
}

function md.Workbook.Schema {
    <#
    .SYNOPSIS
        filter files to *.xlsx, returns 'Get-ExcelFileSchema' as objects
    #>
    [CmdletBinding()]
    param(
        # Paths
        [object[]] $Path,
        [switch]$All
    )

    if($All) {
        $sources =
            $Paths.GetEnumerator()
            | ?{ $_.Value -match '.*xlsx$' }
    } else {
        $Sources = @( $Path )
    }

    # emit
    $Sources
        | ?{ Test-Path $_ }
        | %{
            Get-ExcelFileSchema -Path $_
                | ConvertFrom-Json
        }
}

function md.Export.WorkbookSchema {
    <#
    .synopsis
        a quick summary of all worksheets, in all files as json.
    #>
    [CmdletBinding()]
    param(
        # Paths, if not in $Paths.Values
        [object[]] $Path,
        $Destination,

        # always write a fresh export
        [switch] $Force,

        # also return the objects
        [switch] $PassThru
    )
    $Source = @(
        if( $Path ) { $Path }
        else { $Paths.Values }
    )
    if( -not $Destination ) {
        $Destination = $Paths.json_WorkbookSchema
    }
    if( -not $Force -and (Test-Path $Destination) ) {
        "Using cached schema: $( $Destination ) " | Write-Host -fg 'gray60'
    } else {
        $found = md.Workbook.Schema -Path $Source # $Paths.Values
        $found
            | ConvertTo-Json -Depth 9
            | Set-Content -Path $Destination

        $Destination | md.Log.WroteFile
    }

    if( -not $PassThru ) { return }
    Get-Content -path $Destination | ConvertFrom-Json -Depth 9
}

function md.Export.WorkbookSchema.Xlsx {
    <#
    .synopsis
        a quick summary of all worksheets, in all files as json.
    #>
    [CmdletBinding()]
    param(
        # Paths, if not in $Paths.Values
        [object[]] $Paths
        # $Destination,

        # # always write a fresh export
        # [switch] $Force,

        # # also return the objects
        # [switch] $PassThru
    )
    # also export schemas a an excel sheet
    $schema = md.Export.WorkbookSchema -PassThru
    remove-item $Paths.xlsx_WorkbookSchema -ea Ignore
    $exportExcelSplat = @{
        Path          = $Paths.xlsx_WorkbookSchema
        WorksheetName = 'Schema'
        AutoSize      = $true
        TableName     = 'Schema_data'
        TableStyle    = 'Light5'
        # Show          = $true
        Title         = 'Summary of xlsx schemas by file'
    }

    @(
        $schema
        | %{
            $record = $_
            $record.PropertyNames = $Record.PropertyNames | SOrt-Object -unique | Join-String -sep ', '
            $record
        }
        | Sort-Object ExcelFile, WorksheetName
    )
    | Export-Excel @exportExcelSplat

    md.Log.WroteFile $exportExcelSplat.Path
}

function md.EnsureSubdirsExist {
    <#
    .SYNOPSIS
        build any missing folders
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        # Root path of version directory
        $Path,

        [Parameter()]
        [string[]] $RequiredNames = @('json', 'csv')
    )
        # mkdir (join-path $Paths.ExportRoot_CurrentVersion -ChildPath 'json') -ea 'ignore'

    $versionRoot = Get-Item -ea 'stop' $Path
    foreach( $name in $RequiredNames ) {
        $newPath = Join-Path $versionRoot $name
        $newPath | Join-String -f 'create: "{0}"' -sep ', ' | Write-Verbose
        mkdir -ea 'ignore' $newPath
    }
}

function md.Parse.IngredientsFromCsv {
    <#
    .synopsis
        converts inputs like 'FIBER_SPIKETREE 100, CONCRETE_RAW 25' into ingredient lists
    #>
    param( [string]$Text )
    if( [string]::IsNullOrWhiteSpace( $Text ) ) { return ,@() }
    ,@(
        ($Text -split ',\s+').ForEach({
            $segs = $_ -split '\s+', 2
            [pscustomobject]@{
                Name     = $segs[0];
                Quantity = $segs[1]
            }
        })
    )
}
function md.Parse.ItemsFromList {
    <#
    .synopsis
        converts inputs like 'FIBER_SPIKETREE 100, CONCRETE_RAW 25' into ingredient lists
    #>
    param( [string]$Text )
    if( [string]::IsNullOrWhiteSpace( $Text ) ) { return ,@() }
    ,@(
        $Text -split ',\s+'
    )
}

function md.Parse.Checkbox {
     <#
    .synopsis
        converts boolean style inputs, like 'x' vs blank
    #>
    param( [string]$Text )
    if( $Text.Length -eq 0 ) { return $false }
    if( $Text -match '\s*x\s*') { return $true }
    return $false
}
function md.Format.NullAsString {
     <#
    .synopsis
        If null values, emit an empty string instead. For non-blanky, emit original value
    #>
    param( $Value )
    if($null -eq $Value){ return "" }
    if( [string]::IsNullOrWhiteSpace( $Value ) ) { return $false }

    return $value
}

function md.Convert.BlankPropsToEmpty {
     <#
    .synopsis
        coerce blankables into empty strings for json
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        [object] $InputObject
    )

    process {
        $InputObject.PSObject.Properties  | % {
            if( [string]::IsNullOrWhiteSpace( $_.Value ) ) {
                $InputObject.($_.Name) = ""
            }
        }
        $InputObject
    }
}
function md.Convert.KeyNames {
     <#
    .synopsis
        partially sanitize names, making it more json-ic
    .NOTES
        future: [1] coerce casing. Maybe [TextInfo.ToTitleCase] [2] tolower
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        [object] $InputObject
    )

    process {
        $newObj = [ordered]@{}
        $InputObject.PSObject.Properties  | % {
              $newName         = $_.Name -replace '[ ]+', '_'
              $newName         = $newName.toLower()
              $newObj.$newName = $_.Value
        }
        [pscustomobject]$newObj
    }
}

function _Convert.ExpandSingleProperty {
    param(
        $InputObject,
        [string] $Property
    )
    if( $InputObject.$expandProp.count -gt 0) {
        $InputObject.$expandProp | %{
            $curType            = $_
            $newObj             = $InputObject | Select-Object -Prop *
            $newObj.$expandProp = $curType
            $newObj
        }
    } else {
        $InputObject
    }
}

function md.Convert.ExpandProperty {
     <#
    .synopsis
        Expand nested lists to tables. Emits n-records for a list of n
    .NOTES
        future: [1] coerce casing. Maybe [TextInfo.ToTitleCase] [2] tolower
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        [object] $InputObject,

        # How many properties to expand ?
        [string] $PropertyName
    )

    process {
        _Convert.ExpandSingleProperty -InputObject $InputObject -Property @($PropertyName)[0]
        # $newObj = [ordered]@{}
        # $InputObject.PSObject.Properties  | % {
        #       $newName         = $_.Name -replace '[ ]+', '_'
        #       $newName         = $newName.toLower()
        #       $newObj.$newName = $_.Value
        # }
        # [pscustomobject]$newObj

    }
}

# function md.Convert.TruthyProps {
#      <#
#     .synopsis
#         converts boolean style inputs like 'x' or blank as true / false
#     #>
#     [CmdletBinding()]
#     param(
#         [Parameter(ValueFromPipeline)]
#         [object] $InputObject
#     )

#     process {
#         $InputObject.PSObject.Properties  | % {
#             if( $_.Value -match '^\s*x\s*$' ) {
#                 $InputObject.($_.Name) = $true
#             } elseif( $_.Value -is 'string' and $_.Value.Length -eq 0 )  {
#                 $InputObject.($_.Name) = $false
#             }
#             if( [string]::IsNullOrWhiteSpace( $_.Value ) ) {

#             }
#         }
#         $InputObject
#     }
# }

function md.Export.Biome.Biome_Objects {
    [CmdletBinding()]
    param(
        # Paths hashtable
        [Parameter(Mandatory)] $Paths,

        # always write a fresh export
        [ValidateScript({throw 'nyi'})]
        [switch] $Force
    )
    $PSCmdlet.MyInvocation.MyCommand.Name
        | Join-String -f 'Enter: {0}' | Write-verbose
    # Section: Export item: biome/Biome_Objects
    # todo: refactor like 'md.Export.Changelog'

    $pkg = Open-ExcelPackage -Path $Paths.Raw_Biome
    $book = $pkg.Workbook
    md.Workbook.ListItems $Book
    # $sheets = $pkg.workbook.Worksheets
    # # detect column counts
    # $curSheet = $pkg.Workbook.Worksheets['Biome Objects']

    remove-item $Paths.Xlsx_Biome -ea 'Ignore'

    # Section: Export item: biome/Biome_Objects
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

    # json specific transforms
    $sort_splat = @{
        Property = 'title', 'code', 'exchange_types'
    }

    $forJson = @(
        $Rows | %{
            $record = $_
            $record = md.Convert.BlankPropsToEmpty $Record
            $record = md.Convert.KeyNames $Record
            # coerce blankables into empty strings for json
            $record.'pickups'             = md.Parse.IngredientsFromCsv $record.'pickups'
            $record.'exchange_types'      = md.Parse.ItemsFromList $record.'exchange_types'
            $record.'unclickable'         = md.Parse.Checkbox $record.'unclickable'
            $record.'trails_pass_through' = md.Parse.Checkbox $record.'trails_pass_through'
            $record
        }
    ) | Sort-Object @sort_splat

    $forJson
        | ConvertTo-Json -depth 9
        | Set-Content -path $Paths.Json_Biome_Objects # -Confirm

    $Paths.Json_Biome_Objects | md.Log.WroteFile

    # also emit expanded records
    $forJson = @(
        $Rows | %{
            $record = $_
            $expandProp = 'exchange_types'
            md.Convert.ExpandProperty $Record -Prop $expandProp
        }
    )| Sort-Object @sort_splat
    write-warning 'todo: auto expand all properties dynamically: exchange_types, pickups, etc...'

    $forJson
        | ConvertTo-Json -depth 9
        | Set-Content -path $Paths.json_Biome_Objects_Expanded # -Confirm

    $Paths.json_Biome_Objects_Expanded | md.Log.WroteFile

    if( $false ) { <# test coercion from json to sheet #>
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
    }

    Close-ExcelPackage -ExcelPackage $pkg -NoSave
}

function md.Export.Biome.Plants {
    [CmdletBinding()]
    param(
        # Paths hashtable
        [Parameter(Mandatory)] $Paths,

        # always write a fresh export
        [ValidateScript({throw 'nyi'})]
        [switch] $Force
    )
    $PSCmdlet.MyInvocation.MyCommand.Name
        | Join-String -f 'Enter: {0}' | Write-verbose

    # Section: Export item: biome/Plants
    $pkg = Open-ExcelPackage -Path $Paths.Raw_Biome
    $book = $pkg.Workbook
    md.Workbook.ListItems $Book

    $importExcel_Splat = @{
        ExcelPackage  = $pkg
        WorksheetName = 'Plants'
    }
    $rows =  Import-Excel @importExcel_Splat

    # column descriptions are inline
    $description = $rows | ? Code -Match '^\s*//\s*$' | Select -First 1
    $description | ConvertTo-Json | Set-Content -path $Paths.json_Biome_Plants_ColumnDesc
    $paths.json_Biome_Plants_ColumnDesc | md.Log.WroteFile

    # skip empty and non-data rows
    $rows = @(
        $rows
            | ? { -not [string]::IsNullOrWhiteSpace( $_.CODE ) }
            | ? { $_.CODE -notmatch '^\s*//' }
            # | ? { $_.CODE -notmatch '^\s*\?+\s*$' } # skip "???"
    )

    $exportExcel_Splat = @{
        InputObject   = @( $rows )
        Path          = $Paths.Xlsx_Biome
        Show          = $false
        WorksheetName = 'Plants'
        TableName     = 'Plants_Data'
        TableStyle    = 'Light5'
        AutoSize      = $True
    }

    Export-Excel @exportExcel_Splat

    # json specific transforms
    $sort_splat = @{
        Property = 'code', 'mass'
    }

    $forJson = @(
        $Rows | %{
            $record = $_
            $record = md.Convert.BlankPropsToEmpty $Record
            $record = md.Convert.KeyNames $Record
            # coerce blankables into empty strings for json
            # $record.'pickups'             = md.Parse.IngredientsFromCsv $record.'pickups'
            # $record.'exchange_types'      = md.Parse.ItemsFromList $record.'exchange_types'
            $record.'ignore_grooves'         = md.Parse.Checkbox $record.'ignore_grooves'
            $record.'even_cluster'         = md.Parse.Checkbox $record.'even_cluster'
            # $record.'trails_pass_through' = md.Parse.Checkbox $record.'trails_pass_through'
            $record
        }
    ) | Sort-Object @sort_splat

    $forJson
        | ConvertTo-Json -depth 9
        | Set-Content -path $Paths.Json_Biome_Plants # -Confirm

    $Paths.json_Biome_Plants | md.Log.WroteFile




    # also emit expanded records
    $forJson = @(
        $Rows | %{
            $record = $_
            $record # md.Convert.ExpandProperty $Record -Prop $expandProp
        }
    )| Sort-Object @sort_splat

    Close-ExcelPackage -ExcelPackage $pkg -NoSave
}
function md.Export.TechTree.TechTree {
    <#
    .SYNOPSIS
        Parses and exports 'TechTree.xlsx/TechTree'
    #>
    [CmdletBinding()]
    param(
        # Paths hashtable
        [Parameter(Mandatory)] $Paths,

        # always write a fresh export
        [ValidateScript({throw 'nyi'})]
        [switch] $Force
    )
    $PSCmdlet.MyInvocation.MyCommand.Name
        | Join-String -f 'Enter: {0}' | Write-verbose

    $Regex = @{
        isTierNumber     = '^\s*//\s+tier\s+\d+'
        isGroupName      = '^\s*//' # '^\s*//\s*w\+' #  '^\s*//'
        toIgnoreHeader   = '//\s*unique\s*code'
        stripSlashPrefix = '\s*//\s+'
    }

    # Section: Export item: biome/Plants
    $pkg = Open-ExcelPackage -Path $Paths.Raw_TechTree
    $book = $pkg.Workbook
    md.Workbook.ListItems $Book

    remove-item $Paths.Xlsx_TechTree -ea 'Ignore'
    $importExcel_Splat = @{
        ExcelPackage  = $pkg
        WorksheetName = 'TechTree'

    }
    $rows =  Import-Excel @importExcel_Splat

    # column descriptions are inline
    # $description = $rows | ? Code -Match '^\s*//\s*$' | Select -First 1
    # $description | ConvertTo-Json | Set-Content -path $Paths.json_Biome_Plants_ColumnDesc

    # $paths.json_Biome_Plants_ColumnDesc | md.Log.WroteFile

    # skip empty and non-data rows
    $curGroupName = 'missing'
    $curTierNumber = 'missing'
    $curOrder = -1
    $rows = @(
        $rows
            | % {
                # capture grouping records, else add them to the data
                $record = $_
                $curOrder++
                if ($record.Code -match $Regex.isTierNumber) {
                    $curTierNumber = $record.Code  -replace $regex.StripSlashPrefix, ''
                    return
                } elseif ( $record.Code -match $Regex.isGroupName ) {
                    $curGroupName = $record.Code -replace $regex.stripSlashPrefix, ''
                    return
                } elseif ( $record.Code -match $Regex.toIgnoreHeader ) {
                    return
                }

                $record.PSObject.Properties.Add( [psnoteproperty]::new(
                    'Group', $curGroupName
                ), $true )
                $record.PSObject.Properties.Add( [psnoteproperty]::new(
                    'Tier', $curTierNumber
                ), $true )
                $record.PSObject.Properties.Add( [psnoteproperty]::new(
                    'Order', $curOrder
                ), $true )

                $record
            }
            | ? { $_.Code -notmatch $Regex.toIgnoreHeader }
            | ? { -not [string]::IsNullOrWhiteSpace( $_.CODE ) }
            # | ? { $_.CODE -notmatch  }
            # | ? { $_.CODE -notmatch '^\s*//' }
            # | ? { $_.CODE -notmatch '^\s*\?+\s*$' } # skip "???"
    )



    $exportExcel_Splat = @{
        InputObject   = @( $rows )
        Path          = $Paths.Xlsx_TechTree
        Show          = $true
        WorksheetName = 'TechTree'
        TableName     = 'TechTree_Data'
        TableStyle    = 'Light5'
        AutoSize      = $True
    }

    Export-Excel @exportExcel_Splat

    # json specific transforms
    $sort_splat = @{
        Property = 'code'
    }

    $forJson = @(
        $Rows | %{
            $record = $_
            $record = md.Convert.BlankPropsToEmpty $Record
            $record = md.Convert.KeyNames $Record
            # coerce blankables into empty strings for json
            # $record.'pickups'             = md.Parse.IngredientsFromCsv $record.'pickups'
            # $record.'exchange_types'      = md.Parse.ItemsFromList $record.'exchange_types'
            # $record.'ignore_grooves'         = md.Parse.Checkbox $record.'ignore_grooves'
            # $record.'even_cluster'         = md.Parse.Checkbox $record.'even_cluster'
            # $record.'trails_pass_through' = md.Parse.Checkbox $record.'trails_pass_through'
            $record
        }
    ) | Sort-Object @sort_splat

    $forJson
        | ConvertTo-Json -depth 9
        | Set-Content -path $Paths.json_TechTree_TechTree # -Confirm

    $Paths.json_TechTree_TechTree | md.Log.WroteFile


    # also emit expanded records
    # $forJson = @(
    #     $Rows | %{
    #         $record = $_
    #         $record # md.Convert.ExpandProperty $Record -Prop $expandProp
    #     }
    # )| Sort-Object @sort_splat

    Close-ExcelPackage -ExcelPackage $pkg -NoSave
}

function md.Invoke.FdFind {
    <#
    .SYNOPSIS
        call 'fd' find
    .LINK
        https://github.com/sharkdp/fd
    .NOTES
        # the usage to `fd [OPTIONS] --search-path <path> --search-path <path2> [<pattern>]`
        fd -e xlsx  --base-directory  '' --strip-cwd-prefix=never --search-path './src'
    #>
    [CmdletBinding()]
    param(
        # filter by File types
        [ArgumentCompletions('xlsx', 'csv', 'json', 'log')]
        [string[]] $Extension,

        # extra args
        [string[]] $ArgsList,

        # test cli generated arguments
        [switch] $WhatIf,

        # make paths relative, and linkable in markdown files
        [switch]$PathsAsMarkdown,

        # use '--no-ignore'
        [switch] $UsingNoIgnore
    )
    begin {
        $binFd = Get-Command 'fd' -CommandType Application -TotalCount 1 -ea 'stop'
    }
    end {
        $binArgs = @(
            if( $Extension ) {
                foreach ($ext in $Extension) {
                    "-e", $ext
                }
            }
            if($UsingNoIgnore) { '--no-ignore' }
            $ArgsList

            if( $PathsAsMarkdown ) {
                '--path-separator=/'
                '--strip-cwd-prefix=never'
            }
        )
        $binArgs | Join-String -op 'fd => ' -sep ' '  | Write-verbose
        if( $WhatIf ) {
            $binArgs | Join-String -op 'fd => ' -sep ' '  | Write-host -fg 'skyblue'
            return
        }
        # ...
        & $binFd @binArgs
    }
}

function Markdown.Write.Header {
    <#
    .SYNOPSIS
        Write a markdown H1-H6
    #>
    param(
        [int]$Depth = 2,
        [string] $Text = 'Default' )
    $Prefix = '#' * $Depth -join ''
    "`n${Prefix} ${Text}`n"
}
function Markdown.Write.Newline {
    <#
    .SYNOPSIS
        Write newlines/padding
    #>
    param( [int]$Count = 1 )
    "`n" * $Count -join ''
}
function Markdown.Write.Href {
    <#
    .SYNOPSIS
        Write href
    #>
    param(
        [string]$Text,
        # Allow non [System.Uri] types
        [Alias('Url', 'Href')]
        [string]$Link = '#'
    )
    $Link = $Link -replace '[ ]', '%20' # for github markdown relative links to work
    "[${Text}](${Link})"
}

function Markdown.Write.UL {
    <#
    .SYNOPSIS
        Writes a markdown unordered list
    #>
    param(
        [string[]] $Items
    )
    $Items | Join-String -f "- {0}" -sep "`n" -op "`n"
}

function Markdown.Format.LinksAsUL {
    <#
    .SYNOPSIS
        Converts a list of files to a list of UL links
    .EXAMPLE
        > Markdown.Format.LinksAsUL -Lines @(md.Invoke.FdFind -Extension json -PathsAsMarkdown)
    #>
    param(
        # list of relative paths, like "./json/biome-objects.json"
        [string[]] $Lines
    )
    if( $Lines.Count -eq 0 ) { return }
    Markdown.write.UL @(
        $Lines
        | %{
            $Text = $_ -split '/' | Select -last 1
            $Link = $_
            Markdown.Write.Href -Text $Text -Link $Link
        }
    )
}

function md.Export.Readme.FileListing {
    <#
    .SYNOPSIS
        Automatically build an index of all files generated as a markdown readme
    .LINK
        md.Invoke.FdFind
    #>
    [CmdletBinding()]
    param(
        # Root path to search
        [Parameter(Mandatory)]
        $Path
    )
    $PSCmdlet.MyInvocation.MyCommand.Name
        | Join-String -f 'Enter: {0}' | Write-verbose

    pushd -StackName 'export' $Paths.ExportRoot_CurrentVersion

    $Destination = Join-Path $Path 'readme.md'



    @(
        Markdown.Write.Header -Depth 2 -Text "About"
        "Files generated on: $( (get-date).tostring('yyyy-MM-dd') )"
        Markdown.Write.Header -Depth 2 -Text "Files by Type"

        Markdown.Write.Header -Depth 3 -Text "Json"
        Markdown.Format.LinksAsUL -Lines @(md.Invoke.FdFind -Extension json -PathsAsMarkdown -UsingNoIgnore)

        Markdown.Write.Header -Depth 3 -Text "Xlsx"
        Markdown.Format.LinksAsUL -Lines @(md.Invoke.FdFind -Extension xlsx -PathsAsMarkdown -UsingNoIgnore)

        Markdown.Write.Header -Depth 3 -Text "Csv"
        Markdown.Format.LinksAsUL -Lines @(md.Invoke.FdFind -Extension csv -PathsAsMarkdown -UsingNoIgnore)

        Markdown.Write.Header -Depth 3 -Text "Md"
        Markdown.Format.LinksAsUL -Lines @(md.Invoke.FdFind -Extension md -PathsAsMarkdown -UsingNoIgnore)

        # Markdown.Write.Header -Depth 3 -Text "Log"
        # Markdown.Format.LinksAsUL -Lines @(md.Invoke.FdFind -Extension log -PathsAsMarkdown -UsingNoIgnore)

    ) | Join-String -sep "`n" | Set-Content -Path $Destination
    $Destination | md.Log.WroteFile

    # md.Invoke.FdFind -WhatIf -Extension 'json' -ArgsList $Shared -PathsAsMarkdown -UsingNoIgnore

    popd -StackName 'export'
}
