#Requires -Version 7
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
    param()

    # $rawPath = $Paths.xlsx_Changelog
    # $rawFullJoin-Path $rawPath.DirectoryName "$( $_.baseName )-raw.xlsx"
    $curOutput = $Paths.Xlsx_ChangeLog
    $rawSrc = md.GetRawPath $curOutput

    "md.Export.Changelog => Parse: $( $rawSrc ), Output: $( $curOutput )" | Write-Host -fg 'gray60' -bg 'gray30'

    $imXl = Import-Excel -path $rawSrc -WorksheetName 'Changelog' -ImportColumns 1, 3 -HeaderName 'Version', 'English'

    # using BOM for best results when using Excel csv
    $imXL
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
        "Using cached schema: $( $Destination ) " | Write-Verbose
    } else {
        $found = md.Workbook.Schema -Path $Source # $Paths.Values
        $found
            | ConvertTo-Json -Depth 9
            | Set-Content -Path $Destination
    }

    if( -not $PassThru ) { return }
    Get-Content -path $Destination | ConvertFrom-Json -Depth 9
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
