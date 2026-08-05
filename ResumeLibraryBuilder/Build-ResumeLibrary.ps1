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

    $CanAppendResumeDocx = $ExistingApplications.Count -gt 0 -and (Test-Path -LiteralPath $CombinedResumeDocxPath)
    $CanAppendResumeText = $ExistingApplications.Count -gt 0 -and (Test-Path -LiteralPath $CombinedResumeTxtPath)
    $CanAppendCoverLetterDocx = $ExistingApplications.Count -gt 0 -and (Test-Path -LiteralPath $CombinedCoverLetterDocxPath)
    $CanAppendCoverLetterText = $ExistingApplications.Count -gt 0 -and (Test-Path -LiteralPath $CombinedCoverLetterTxtPath)

    $CanSkipKnownFolders = $CanAppendResumeDocx -and $CanAppendResumeText -and $CanAppendCoverLetterDocx -and $CanAppendCoverLetterText

    $ExistingRelativePaths = @{}
    $ExistingResumePaths = @{}
    $ExistingCoverLetterPaths = @{}
    foreach ($application in $ExistingApplications) {
        if ($application.RelativePath) { $ExistingRelativePaths[$application.RelativePath] = $true }
        if ($application.ResumePath) { $ExistingResumePaths[$application.ResumePath] = $true }
        if ($application.CoverLetterPath) { $ExistingCoverLetterPaths[$application.CoverLetterPath] = $true }
    }

    $KnownRelativePaths = if ($CanSkipKnownFolders) { $ExistingRelativePaths } else { $null }
    $DiscoveredApplications = Find-JobApplications -RootFolder $RootFolder `
        -ResumeTitle $ResumeTitle `
        -CoverLetterKeywords $CoverLetterKeywords `
        -SupportedExtensions $SupportedExtensions `
        -ExistingRelativePaths $KnownRelativePaths `
        -PreferDocx:$PreferDocx

    if ($CanSkipKnownFolders) {
        $Applications = @($ExistingApplications + $DiscoveredApplications) | Sort-Object Company, Job, RelativePath
    }
    else {
        $Applications = $DiscoveredApplications
    }

    $NewResumeApplications = @($Applications | Where-Object { $_.ResumePath -and -not $ExistingResumePaths.ContainsKey($_.ResumePath) })
    $NewCoverLetterApplications = @($Applications | Where-Object { $_.CoverLetterPath -and -not $ExistingCoverLetterPaths.ContainsKey($_.CoverLetterPath) })

    $ResumeDocxApplicationsToExport = if ($CanAppendResumeDocx) { $NewResumeApplications } else { @($Applications | Where-Object { $_.ResumePath }) }
    $ResumeTextApplicationsToExport = if ($CanAppendResumeText) { $NewResumeApplications } else { @($Applications | Where-Object { $_.ResumePath }) }
    $CoverLetterDocxApplicationsToExport = if ($CanAppendCoverLetterDocx) { $NewCoverLetterApplications } else { @($Applications | Where-Object { $_.CoverLetterPath }) }
    $CoverLetterTextApplicationsToExport = if ($CanAppendCoverLetterText) { $NewCoverLetterApplications } else { @($Applications | Where-Object { $_.CoverLetterPath }) }

    Write-Log "Found $($Applications.Count) total job application folder(s)." "INFO"
    Write-Log "Scanned $(if ($CanSkipKnownFolders) { 'only new folders' } else { 'all folders because at least one combined output must be rebuilt' })." "INFO"
    Write-Log "New job application folder(s) to export: $($DiscoveredApplications.Count)" "INFO"

    Write-Log "Resume DOCX export mode: $(if ($CanAppendResumeDocx) { 'append new files only' } else { 'rebuild all resumes' })" "INFO"
    Write-Log "Resume DOCX file(s) to add: $($ResumeDocxApplicationsToExport.Count)" "INFO"
    Write-Log "Resume TXT export mode: $(if ($CanAppendResumeText) { 'append new files only' } else { 'rebuild all resumes' })" "INFO"
    Write-Log "Resume TXT file(s) to add: $($ResumeTextApplicationsToExport.Count)" "INFO"
    Write-Log "Cover-letter DOCX export mode: $(if ($CanAppendCoverLetterDocx) { 'append new files only' } else { 'rebuild all cover letters' })" "INFO"
    Write-Log "Cover-letter DOCX file(s) to add: $($CoverLetterDocxApplicationsToExport.Count)" "INFO"
    Write-Log "Cover-letter TXT export mode: $(if ($CanAppendCoverLetterText) { 'append new files only' } else { 'rebuild all cover letters' })" "INFO"
    Write-Log "Cover-letter TXT file(s) to add: $($CoverLetterTextApplicationsToExport.Count)" "INFO"

    Export-ApplicationsIndex -Applications $Applications -OutputPath $ApplicationsIndexPath
    Export-ScanReport -Applications $Applications -OutputPath (Join-Path $OutputFolder $ScanReportName)

    $HasExports = $ResumeDocxApplicationsToExport.Count -gt 0 -or $ResumeTextApplicationsToExport.Count -gt 0 -or $CoverLetterDocxApplicationsToExport.Count -gt 0 -or $CoverLetterTextApplicationsToExport.Count -gt 0
    if (-not $HasExports) {
        Write-Log "No new resumes or cover letters to add; combined outputs were left unchanged." "INFO"
    }
    else {
        $WordContext = $null
        try {
            $WordContext = New-WordContext

            if ($ResumeDocxApplicationsToExport.Count -gt 0 -or $ResumeTextApplicationsToExport.Count -gt 0) {
                if ($ResumeDocxApplicationsToExport.Count -gt 0) {
                    Export-CombinedResumeDocx -Applications $ResumeDocxApplicationsToExport `
                    -OutputPath $CombinedResumeDocxPath `
                    -WordContext $WordContext `
                    -MarginTopInches $ResumeDocMarginTopInches `
                    -MarginBottomInches $ResumeDocMarginBottomInches `
                    -MarginLeftInches $ResumeDocMarginLeftInches `
                    -MarginRightInches $ResumeDocMarginRightInches `
                    -Append:$CanAppendResumeDocx
                }

                if ($ResumeTextApplicationsToExport.Count -gt 0) {
                    Export-CombinedResumeText -Applications $ResumeTextApplicationsToExport `
                    -OutputPath $CombinedResumeTxtPath `
                    -WordContext $WordContext `
                    -Append:$CanAppendResumeText
                }
            }
            else {
                Write-Log "No resume exports are needed." "INFO"
            }

            if ($CoverLetterDocxApplicationsToExport.Count -gt 0 -or $CoverLetterTextApplicationsToExport.Count -gt 0) {
                if ($CoverLetterDocxApplicationsToExport.Count -gt 0) {
                    Export-CombinedCoverLetterDocx -Applications $CoverLetterDocxApplicationsToExport `
                    -OutputPath $CombinedCoverLetterDocxPath `
                    -WordContext $WordContext `
                    -MarginTopInches $CoverLetterDocMarginTopInches `
                    -MarginBottomInches $CoverLetterDocMarginBottomInches `
                    -MarginLeftInches $CoverLetterDocMarginLeftInches `
                    -MarginRightInches $CoverLetterDocMarginRightInches `
                    -Append:$CanAppendCoverLetterDocx
                }

                if ($CoverLetterTextApplicationsToExport.Count -gt 0) {
                    Export-CombinedCoverLetterText -Applications $CoverLetterTextApplicationsToExport `
                    -OutputPath $CombinedCoverLetterTxtPath `
                    -WordContext $WordContext `
                    -Append:$CanAppendCoverLetterText
                }
            }
            else {
                Write-Log "No cover-letter exports are needed." "INFO"
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
