# ===== Lookup-Workstations.ps1 =====
[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$CsvPath = ".\workstations.csv",

    # Optional: paste a file of queries (one per line)
    [Parameter(Mandatory=$false)]
    [string]$InputFile,

    # Optional: export results to a CSV
    [Parameter(Mandatory=$false)]
    [string]$ExportPath
)

function Normalize([string]$s) {
    if ($null -eq $s) { return "" }
    return ($s.Trim())
}

function IsMachineLike([string]$q) {
    # Heuristic: contains digits and has a hyphen/underscore OR looks like sitecode-machine format
    # Adjust pattern to your naming convention if needed
    return ($q -match '^[A-Za-z0-9]+[-_][A-Za-z0-9-_.]+$' -and $q -match '\d')
}

function LoadInventory([string]$path) {
    if (!(Test-Path $path)) {
        throw "CSV not found: $path"
    }

    $rows = Import-Csv -Path $path

    # Ensure required columns exist
    $required = @("Name","Description")
    foreach ($c in $required) {
        if (-not ($rows | Get-Member -Name $c -MemberType NoteProperty)) {
            throw "CSV must contain columns: Name, Description. Missing: $c"
        }
    }

    # Normalize fields
    foreach ($r in $rows) {
        $r.Name = Normalize $r.Name
        $r.Description = $r.Description  # keep full text as-is (do not trim away details)
    }

    return $rows
}

function ReadQueriesInteractive {
    Write-Host ""
    Write-Host "Paste workstation numbers OR free-text (names, keywords)."
    Write-Host "One entry per line. Press ENTER on an empty line to run the search."
    Write-Host ""

    $list = New-Object System.Collections.Generic.List[string]
    while ($true) {
        $line = Read-Host
        if ([string]::IsNullOrWhiteSpace($line)) { break }
        $list.Add((Normalize $line))
    }
    return $list
}

function ReadQueriesFromFile([string]$path) {
    if (!(Test-Path $path)) { throw "Input file not found: $path" }
    return Get-Content -Path $path | ForEach-Object { Normalize $_ } | Where-Object { $_ }
}

function SearchInventory($inventory, [string]$query) {
    $q = Normalize $query
    if (-not $q) { return @() }

    $results = @()

    if (IsMachineLike $q) {
        # Machine lookup: exact match first
        $exact = $inventory | Where-Object { $_.Name -ieq $q }
        if ($exact) { return $exact }

        # Fallback: partial match on Name (contains)
        $partial = $inventory | Where-Object { $_.Name -ilike "*$q*" }
        if ($partial) { return $partial }

        # Also allow machine id to be found inside Description (sometimes embedded)
        $inDesc = $inventory | Where-Object { ($_.Description -as [string]) -imatch [regex]::Escape($q) }
        return $inDesc
    }
    else {
        # Free text search inside description (case-insensitive substring)
        # This supports "John Doe", "Ubuntu", "T14", etc.
        $pattern = [regex]::Escape($q)
        $textHits = $inventory | Where-Object { ($_.Description -as [string]) -imatch $pattern }

        # Extra: also allow searching in Name in case someone pastes partial machine codes without hyphen
        $nameHits = $inventory | Where-Object { $_.Name -ilike "*$q*" }

        # Merge + unique by Name+Description
        $combined = @($textHits + $nameHits) | Sort-Object Name, Description -Unique
        return $combined
    }
}

# ===== Main =====
try {
    $inventory = LoadInventory $CsvPath

    $queries =
        if ($InputFile) { ReadQueriesFromFile $InputFile }
        else { ReadQueriesInteractive }

    if (-not $queries -or $queries.Count -eq 0) {
        Write-Host "No queries provided. Exiting."
        exit 0
    }

    $allResults = New-Object System.Collections.Generic.List[object]

    foreach ($q in $queries) {
        $hits = SearchInventory -inventory $inventory -query $q

        if (-not $hits -or $hits.Count -eq 0) {
            $allResults.Add([pscustomobject]@{
                Query       = $q
                MatchType   = "NotFound"
                Name        = ""
                Description = ""
            })
            continue
        }

        foreach ($h in $hits) {
            $allResults.Add([pscustomobject]@{
                Query       = $q
                MatchType   = if (IsMachineLike $q) { "MachineLookup" } else { "TextSearch" }
                Name        = $h.Name
                Description = $h.Description
            })
        }
    }

    # Display nicely
    $allResults |
        Sort-Object Query, Name |
        Format-Table -AutoSize Query, MatchType, Name, Description

    if ($ExportPath) {
        $allResults | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $ExportPath
        Write-Host ""
        Write-Host "Exported results to: $ExportPath"
    }

} catch {
    Write-Error $_.Exception.Message
    exit 1
}
