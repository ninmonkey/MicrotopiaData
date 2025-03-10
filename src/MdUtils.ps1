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
