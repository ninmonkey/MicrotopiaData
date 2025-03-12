filter Expand-Property {
    <#
        .SYNOPSIS
            Expands an array property, creating a duplicate object for each value
        .LINK
            https://gist.github.com/Jaykul/f55f8418c6d03d426891dfaa5b424ac6
        .EXAMPLE
            [PSCustomObject]@{ Name = "A"; Value = @(1,2,3) } | Expand-Property Value
            Name Value
            ---- -----
            A        1
            A        2
            A        3
        .EXAMPLE
            [PSCustomObject]@{
                Name = "DevOps"
                Members = "Joe", "Phil", "Barb"
                MemberOf = "Ops", "Dev"
            } |
            Expand-Property MemberOf |
            Expand-Property Members
            Name   MemberOf Members
            ----   -------- -------
            DevOps Ops      Joe
            DevOps Ops      Phil
            DevOps Ops      Barb
            DevOps Dev      Joe
            DevOps Dev      Phil
            DevOps Dev      Barb
    #>
    param(
        # The name of a property on the input object, that has more than one value
        [Alias("Property")]
        [string]$Name,

        # The input object to duplicated
        [Parameter(ValueFromPipeline)]
        $InputObject
    )
    foreach ($Value in $InputObject.$Name) {
        $InputObject | Select-Object *, @{ Name = $Name; Expr = { $Value } } -Exclude $Name
    }
}

function Group-Property {
    <#
        .SYNOPSIS
            Given a collection of objects with where only one property differs,
            returns a single object with an array property containing the values
        .LINK
            https://gist.github.com/Jaykul/f55f8418c6d03d426891dfaa5b424ac6
        .EXAMPLE
            @(
                [PSCustomObject]@{ Name = "A"; Value = 1 }
                [PSCustomObject]@{ Name = "A"; Value = 2 }
                [PSCustomObject]@{ Name = "A"; Value = 3 }
            ) | Group-Property Value
            Name Value
            ---- -----
               A {1, 2, 3}
        .EXAMPLE
        @"
        Name,MemberOf,Members
        DevOps,Ops,Joe
        DevOps,Ops,Phil
        DevOps,Ops,Barb
        DevOps,Dev,Joe
        DevOps,Dev,Phil
        DevOps,Dev,Barb
        "@ | ConvertFrom-Csv | Group-Property Members | Group-Property MemberOf
        Name   MemberOf   Members
        ----   --------   -------
        DevOps {Ops, Dev} {Barb, Joe, Phil}
    #>
    [CmdletBinding()]
    param(
        # The input objects to deduplicated
        [Parameter(ValueFromPipeline)]
        [psobject]
        ${InputObject},

        # The name of the property that has unique values
        [Parameter(Position = 0)]
        [string]
        ${Name},

        # Specifies the culture to use when comparing strings.
        [string]
        ${Culture},

        # Indicates that the grouping is case-sensitive. Without this parameter, the property values of objects in a group might have different cases.
        [switch]
        ${CaseSensitive}
    )

    begin {
        try {
            $outBuffer = $null
            $null = $PSBoundParameters.Remove("Name")
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }

            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Microsoft.PowerShell.Utility\Group-Object', [System.Management.Automation.CommandTypes]::Cmdlet)
        } catch {
            throw
        }
    }

    process {
        try {
            if (!$steppablePipeline) {
                $PSBoundParameters["Property"] = $InputObject.PSObject.Properties.Name.Where{ $_ -ne $Name }
                $null = $PSBoundParameters.Remove("InputObject")
                Write-Verbose "Group by $($PSBoundParameters["Property"] -join ', ')"
                $scriptCmd = { & $wrappedCmd @PSBoundParameters | ForEach-Object { $_.Group[0].$Name = $_.Group.$Name; $_.Group[0] } }
                $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
                $steppablePipeline.Begin($PSCmdlet)
            }

            $steppablePipeline.Process($_)
        } catch {
            throw
        }
    }

    end {
        try {
            $steppablePipeline.End()
        } catch {
            throw
        }
    }

    clean {
        if ($null -ne $steppablePipeline) {
            $steppablePipeline.Clean()
        }
    }
}
