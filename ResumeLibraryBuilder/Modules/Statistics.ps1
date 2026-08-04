function Write-BuildStatistics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Applications,
        [Parameter(Mandatory)][datetime]$StartedAt
    )

    $elapsed = (Get-Date) - $StartedAt
    $resumeCount = @($Applications | Where-Object { $_.ResumePath }).Count
    $coverCount = @($Applications | Where-Object { $_.CoverLetterPath }).Count
    $companyCount = @($Applications | Select-Object -ExpandProperty Company -Unique).Count

    Write-Log "Applications: $($Applications.Count)" "INFO"
    Write-Log "Companies: $companyCount" "INFO"
    Write-Log "Resumes found: $resumeCount" "INFO"
    Write-Log "Cover letters found: $coverCount" "INFO"
    Write-Log "Elapsed: $($elapsed.ToString())" "INFO"
}
