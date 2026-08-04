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

function Import-ApplicationsIndex {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$InputPath
    )

    if (-not (Test-Path -LiteralPath $InputPath)) {
        Write-Log "No existing JSON index found: $InputPath" "INFO"
        return @()
    }

    try {
        $content = Get-Content -LiteralPath $InputPath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($content)) {
            Write-Log "Existing JSON index is empty: $InputPath" "WARN"
            return @()
        }

        $applications = @($content | ConvertFrom-Json)
        Write-Log "Loaded existing JSON index with $($applications.Count) application(s): $InputPath" "INFO"
        return $applications
    }
    catch {
        Write-Log "Could not read existing JSON index; this run will rebuild combined outputs. $($_.Exception.Message)" "WARN"
        return @()
    }
}
