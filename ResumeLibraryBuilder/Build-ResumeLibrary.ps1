<#
.SYNOPSIS
Combines job-application resumes into one formatted Word document and one text file.

.DESCRIPTION
Scans a root folder structured as Root\CompanyName\JobName\resume.docx, identifies the
best resume for each job folder, preserves Word formatting in Combined_Resumes.docx, and
creates Combined_Resumes.txt for AI ingestion.
#>
[CmdletBinding()]
param(
    [string]$RootFolderOverride
)

$ErrorActionPreference = "Stop"

. "$PSScriptRoot\Config.ps1"
. "$PSScriptRoot\Modules\Logger.ps1"
. "$PSScriptRoot\Modules\Initialize.ps1"
. "$PSScriptRoot\Modules\ScanFolders.ps1"
. "$PSScriptRoot\Modules\Word.ps1"
. "$PSScriptRoot\Modules\Text.ps1"
. "$PSScriptRoot\Modules\Json.ps1"
. "$PSScriptRoot\Modules\Statistics.ps1"
. "$PSScriptRoot\Modules\Cleanup.ps1"

if ($RootFolderOverride) {
    $RootFolder = $RootFolderOverride
}

$State = Initialize-ResumeLibraryBuild -RootFolder $RootFolder -OutputFolder $OutputFolder

try {
    Write-Log "Starting ResumeLibraryBuilder" "INFO"
    Write-Log "Root folder: $RootFolder" "INFO"
    Write-Log "Output folder: $OutputFolder" "INFO"

    $ApplicationsIndexPath = Join-Path $OutputFolder $ApplicationsIndexName
    $CombinedResumeDocxPath = Join-Path $OutputFolder $CombinedResumeDocxName
    $CombinedResumeTxtPath = Join-Path $OutputFolder $CombinedResumeTxtName
    $CombinedCoverLetterDocxPath = Join-Path $OutputFolder $CombinedCoverLetterDocxName
    $CombinedCoverLetterTxtPath = Join-Path $OutputFolder $CombinedCoverLetterTxtName

    $ExistingApplications = @(Import-ApplicationsIndex -InputPath $ApplicationsIndexPath)

    $CanUseIncrementalScan = $ExistingApplications.Count -gt 0 -and `
        (Test-Path -LiteralPath $CombinedResumeDocxPath) -and `
        (Test-Path -LiteralPath $CombinedResumeTxtPath) -and `
        (Test-Path -LiteralPath $CombinedCoverLetterDocxPath) -and `
        (Test-Path -LiteralPath $CombinedCoverLetterTxtPath)

    $ExistingRelativePaths = @{}
    foreach ($application in $ExistingApplications) {
        if ($application.RelativePath) { $ExistingRelativePaths[$application.RelativePath] = $true }
    }

    $KnownRelativePaths = if ($CanUseIncrementalScan) { $ExistingRelativePaths } else { $null }
    $DiscoveredApplications = Find-JobApplications -RootFolder $RootFolder `
        -ResumeTitle $ResumeTitle `
        -CoverLetterKeywords $CoverLetterKeywords `
        -SupportedExtensions $SupportedExtensions `
        -ExistingRelativePaths $KnownRelativePaths `
        -PreferDocx:$PreferDocx

    if ($CanUseIncrementalScan) {
        $Applications = @($ExistingApplications + $DiscoveredApplications) | Sort-Object Company, Job, RelativePath
        $ResumeApplicationsToExport = @($DiscoveredApplications | Where-Object { $_.ResumePath })
        $CoverLetterApplicationsToExport = @($DiscoveredApplications | Where-Object { $_.CoverLetterPath })
    }
    else {
        $Applications = $DiscoveredApplications
        $ResumeApplicationsToExport = @($Applications | Where-Object { $_.ResumePath })
        $CoverLetterApplicationsToExport = @($Applications | Where-Object { $_.CoverLetterPath })
    }

    Write-Log "Found $($Applications.Count) total job application folder(s)." "INFO"
    Write-Log "New job application folder(s) to export: $($DiscoveredApplications.Count)" "INFO"

    $CanAppendResumes = $CanUseIncrementalScan
    $CanAppendCoverLetters = $CanUseIncrementalScan

    Write-Log "Resume export mode: $(if ($CanAppendResumes) { 'append new files only' } else { 'rebuild all resumes' })" "INFO"
    Write-Log "Resume file(s) to add: $($ResumeApplicationsToExport.Count)" "INFO"
    Write-Log "Cover-letter export mode: $(if ($CanAppendCoverLetters) { 'append new files only' } else { 'rebuild all cover letters' })" "INFO"
    Write-Log "Cover-letter file(s) to add: $($CoverLetterApplicationsToExport.Count)" "INFO"

    Export-ApplicationsIndex -Applications $Applications -OutputPath $ApplicationsIndexPath
    Export-ScanReport -Applications $Applications -OutputPath (Join-Path $OutputFolder $ScanReportName)

    $HasExports = $ResumeApplicationsToExport.Count -gt 0 -or $CoverLetterApplicationsToExport.Count -gt 0
    if (-not $HasExports) {
        Write-Log "No new resumes or cover letters to add; combined outputs were left unchanged." "INFO"
    }
    else {
        $WordContext = $null
        try {
            $WordContext = New-WordContext

            if ($ResumeApplicationsToExport.Count -gt 0) {
                Export-CombinedResumeDocx -Applications $ResumeApplicationsToExport `
                    -OutputPath $CombinedResumeDocxPath `
                    -WordContext $WordContext `
                    -MarginTopInches $ResumeDocMarginTopInches `
                    -MarginBottomInches $ResumeDocMarginBottomInches `
                    -MarginLeftInches $ResumeDocMarginLeftInches `
                    -MarginRightInches $ResumeDocMarginRightInches `
                    -Append:$CanAppendResumes

                Export-CombinedResumeText -Applications $ResumeApplicationsToExport `
                    -OutputPath $CombinedResumeTxtPath `
                    -WordContext $WordContext `
                    -Append:$CanAppendResumes
            }
            else {
                Write-Log "No new resumes to add." "INFO"
            }

            if ($CoverLetterApplicationsToExport.Count -gt 0) {
                Export-CombinedCoverLetterDocx -Applications $CoverLetterApplicationsToExport `
                    -OutputPath $CombinedCoverLetterDocxPath `
                    -WordContext $WordContext `
                    -MarginTopInches $CoverLetterDocMarginTopInches `
                    -MarginBottomInches $CoverLetterDocMarginBottomInches `
                    -MarginLeftInches $CoverLetterDocMarginLeftInches `
                    -MarginRightInches $CoverLetterDocMarginRightInches `
                    -Append:$CanAppendCoverLetters

                Export-CombinedCoverLetterText -Applications $CoverLetterApplicationsToExport `
                    -OutputPath $CombinedCoverLetterTxtPath `
                    -WordContext $WordContext `
                    -Append:$CanAppendCoverLetters
            }
            else {
                Write-Log "No new cover letters to add." "INFO"
            }
        }
        finally {
            if ($WordContext) {
                Close-WordContext -WordContext $WordContext
            }
        }
    }

    Write-BuildStatistics -Applications $Applications -StartedAt $State.StartedAt
    Write-Log "Done." "SUCCESS"
}
catch {
    Write-Log $_.Exception.Message "ERROR"
    throw
}
