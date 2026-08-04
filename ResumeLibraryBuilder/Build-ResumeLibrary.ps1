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

    $IndexPath = Join-Path $OutputFolder $ApplicationsIndexName
    $ResumeDocxPath = Join-Path $OutputFolder $CombinedResumeDocxName
    $ResumeTxtPath = Join-Path $OutputFolder $CombinedResumeTxtName
    $CoverLetterDocxPath = Join-Path $OutputFolder $CombinedCoverLetterDocxName
    $CoverLetterTxtPath = Join-Path $OutputFolder $CombinedCoverLetterTxtName

    $HadPreviousIndex = Test-Path -Path $IndexPath -PathType Leaf
    $PreviousApplications = Import-ApplicationsIndex -Path $IndexPath

    $Applications = Find-JobApplications -RootFolder $RootFolder `
        -ResumeTitle $ResumeTitle `
        -CoverLetterKeywords $CoverLetterKeywords `
        -SupportedExtensions $SupportedExtensions `
        -PreferDocx:$PreferDocx

    Write-Log "Found $($Applications.Count) job application folder(s)." "INFO"

    $NewResumeApplications = @(Get-NewDocumentApplications -CurrentApplications $Applications -PreviousApplications $PreviousApplications -DocumentType Resume)
    $NewCoverLetterApplications = @(Get-NewDocumentApplications -CurrentApplications $Applications -PreviousApplications $PreviousApplications -DocumentType CoverLetter)

    if (-not $HadPreviousIndex) {
        Write-Log "No previous index is available. Rebuilding all combined outputs to avoid duplicates." "INFO"
        foreach ($outputPath in @($ResumeDocxPath, $ResumeTxtPath, $CoverLetterDocxPath, $CoverLetterTxtPath)) {
            if (Test-Path -Path $outputPath -PathType Leaf) { Remove-Item -Path $outputPath -Force }
        }
        $NewResumeApplications = @($Applications | Where-Object { $_.ResumePath })
        $NewCoverLetterApplications = @($Applications | Where-Object { $_.CoverLetterPath })
    }

    if (-not (Test-Path -Path $ResumeDocxPath -PathType Leaf) -or -not (Test-Path -Path $ResumeTxtPath -PathType Leaf)) {
        Write-Log "Resume outputs are missing or incomplete. Rebuilding resume outputs from all scanned resumes." "INFO"
        if (Test-Path -Path $ResumeDocxPath -PathType Leaf) { Remove-Item -Path $ResumeDocxPath -Force }
        if (Test-Path -Path $ResumeTxtPath -PathType Leaf) { Remove-Item -Path $ResumeTxtPath -Force }
        $NewResumeApplications = @($Applications | Where-Object { $_.ResumePath })
    }

    if (-not (Test-Path -Path $CoverLetterDocxPath -PathType Leaf) -or -not (Test-Path -Path $CoverLetterTxtPath -PathType Leaf)) {
        Write-Log "Cover-letter outputs are missing or incomplete. Rebuilding cover-letter outputs from all scanned cover letters." "INFO"
        if (Test-Path -Path $CoverLetterDocxPath -PathType Leaf) { Remove-Item -Path $CoverLetterDocxPath -Force }
        if (Test-Path -Path $CoverLetterTxtPath -PathType Leaf) { Remove-Item -Path $CoverLetterTxtPath -Force }
        $NewCoverLetterApplications = @($Applications | Where-Object { $_.CoverLetterPath })
    }

    Write-Log "New or changed resumes to add: $($NewResumeApplications.Count)" "INFO"
    Write-Log "New or changed cover letters to add: $($NewCoverLetterApplications.Count)" "INFO"

    Export-ScanReport -Applications $Applications -OutputPath (Join-Path $OutputFolder $ScanReportName)

    $WordContext = New-WordContext
    try {
        Export-CombinedResumeDocx -Applications $NewResumeApplications `
            -OutputPath $ResumeDocxPath `
            -WordContext $WordContext `
            -MarginTopInches $ResumeDocMarginTopInches `
            -MarginBottomInches $ResumeDocMarginBottomInches `
            -MarginLeftInches $ResumeDocMarginLeftInches `
            -MarginRightInches $ResumeDocMarginRightInches

        Export-CombinedResumeText -Applications $NewResumeApplications `
            -OutputPath $ResumeTxtPath `
            -WordContext $WordContext `
            -Append

        Export-CombinedCoverLetterDocx -Applications $NewCoverLetterApplications `
            -OutputPath $CoverLetterDocxPath `
            -WordContext $WordContext `
            -MarginTopInches $CoverLetterDocMarginTopInches `
            -MarginBottomInches $CoverLetterDocMarginBottomInches `
            -MarginLeftInches $CoverLetterDocMarginLeftInches `
            -MarginRightInches $CoverLetterDocMarginRightInches

        Export-CombinedCoverLetterText -Applications $NewCoverLetterApplications `
            -OutputPath $CoverLetterTxtPath `
            -WordContext $WordContext `
            -Append
    }
    finally {
        Close-WordContext -WordContext $WordContext
    }

    Export-ApplicationsIndex -Applications $Applications -OutputPath $IndexPath

    Write-BuildStatistics -Applications $Applications -StartedAt $State.StartedAt
    Write-Log "Done." "SUCCESS"
}
catch {
    Write-Log $_.Exception.Message "ERROR"
    throw
}
