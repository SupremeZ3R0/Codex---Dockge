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

    $Applications = Find-JobApplications -RootFolder $RootFolder `
        -ResumeTitle $ResumeTitle `
        -CoverLetterKeywords $CoverLetterKeywords `
        -SupportedExtensions $SupportedExtensions `
        -PreferDocx:$PreferDocx

    Write-Log "Found $($Applications.Count) job application folder(s)." "INFO"

    Export-ApplicationsIndex -Applications $Applications -OutputPath (Join-Path $OutputFolder $ApplicationsIndexName)
    Export-ScanReport -Applications $Applications -OutputPath (Join-Path $OutputFolder $ScanReportName)

    $WordContext = New-WordContext
    try {
        Export-CombinedResumeDocx -Applications $Applications `
            -OutputPath (Join-Path $OutputFolder $CombinedResumeDocxName) `
            -WordContext $WordContext

        Export-CombinedResumeText -Applications $Applications `
            -OutputPath (Join-Path $OutputFolder $CombinedResumeTxtName) `
            -WordContext $WordContext

        Export-CombinedCoverLetterDocx -Applications $Applications `
            -OutputPath (Join-Path $OutputFolder $CombinedCoverLetterDocxName) `
            -WordContext $WordContext

        Export-CombinedCoverLetterText -Applications $Applications `
            -OutputPath (Join-Path $OutputFolder $CombinedCoverLetterTxtName) `
            -WordContext $WordContext
    }
    finally {
        Close-WordContext -WordContext $WordContext
    }

    Write-BuildStatistics -Applications $Applications -StartedAt $State.StartedAt
    Write-Log "Done." "SUCCESS"
}
catch {
    Write-Log $_.Exception.Message "ERROR"
    throw
}
