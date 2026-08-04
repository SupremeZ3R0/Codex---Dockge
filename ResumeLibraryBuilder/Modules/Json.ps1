function Import-ApplicationsIndex {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        Write-Log "No previous JSON index found. This run will build all outputs." "INFO"
        return @()
    }

    try {
        $content = Get-Content -Path $Path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($content)) { return @() }
        $items = $content | ConvertFrom-Json
        Write-Log "Loaded previous JSON index: $Path" "INFO"
        return @($items)
    }
    catch {
        Write-Log "Could not read previous JSON index. This run will build all outputs. $($_.Exception.Message)" "WARNING"
        return @()
    }
}

function Get-DocumentSignature {
    param(
        [Parameter(Mandatory)]$Application,
        [ValidateSet("Resume", "CoverLetter")][string]$DocumentType
    )

    if ($DocumentType -eq "Resume") {
        if (-not $Application.ResumePath -or -not $Application.ResumeLastWriteTimeUtc) { return $null }
        return "$($Application.ResumePath)|$($Application.ResumeLastWriteTimeUtc)"
    }

    if (-not $Application.CoverLetterPath -or -not $Application.CoverLetterLastWriteTimeUtc) { return $null }
    return "$($Application.CoverLetterPath)|$($Application.CoverLetterLastWriteTimeUtc)"
}

function Get-NewDocumentApplications {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$CurrentApplications,
        [Parameter(Mandatory)]$PreviousApplications,
        [ValidateSet("Resume", "CoverLetter")][string]$DocumentType
    )

    if (-not $PreviousApplications -or @($PreviousApplications).Count -eq 0) {
        return @($CurrentApplications | Where-Object { Get-DocumentSignature -Application $_ -DocumentType $DocumentType })
    }

    $previousSignatures = New-Object 'System.Collections.Generic.HashSet[string]'
    $previousPaths = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($previous in @($PreviousApplications)) {
        $signature = Get-DocumentSignature -Application $previous -DocumentType $DocumentType
        if ($signature) { [void]$previousSignatures.Add($signature) }

        if ($DocumentType -eq "Resume" -and $previous.ResumePath) { [void]$previousPaths.Add($previous.ResumePath) }
        if ($DocumentType -eq "CoverLetter" -and $previous.CoverLetterPath) { [void]$previousPaths.Add($previous.CoverLetterPath) }
    }

    $newApplications = New-Object System.Collections.Generic.List[object]
    foreach ($current in @($CurrentApplications)) {
        $signature = Get-DocumentSignature -Application $current -DocumentType $DocumentType
        if (-not $signature) { continue }

        $path = if ($DocumentType -eq "Resume") { $current.ResumePath } else { $current.CoverLetterPath }

        # Older indexes did not store LastWriteTime values. If the old index has
        # only a matching path, treat it as already exported to avoid duplicates.
        if ($previousSignatures.Contains($signature)) { continue }
        if ($previousPaths.Contains($path) -and $previousSignatures.Count -eq 0) { continue }

        $newApplications.Add($current)
    }

    return $newApplications
}

function Export-ApplicationsIndex {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Applications,
        [Parameter(Mandatory)][string]$OutputPath
    )

    $Applications | ConvertTo-Json -Depth 5 | Set-Content -Path $OutputPath -Encoding UTF8
    Write-Log "Saved JSON index: $OutputPath" "SUCCESS"
}

function Export-ScanReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Applications,
        [Parameter(Mandatory)][string]$OutputPath
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("ResumeLibraryBuilder scan report")
    $lines.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $lines.Add("")

    foreach ($application in $Applications) {
        $lines.Add("Company: $($application.Company)")
        $lines.Add("Job: $($application.Job)")
        $lines.Add("Folder: $($application.RelativePath)")
        $lines.Add("Resume: $(if ($application.ResumeFile) { $application.ResumeFile } else { 'MISSING' })")
        $lines.Add("Cover Letter: $(if ($application.CoverLetterFile) { $application.CoverLetterFile } else { 'MISSING' })")
        $lines.Add("------------------------------------------------------------")
    }

    $lines | Set-Content -Path $OutputPath -Encoding UTF8
    Write-Log "Saved scan report: $OutputPath" "SUCCESS"
}
