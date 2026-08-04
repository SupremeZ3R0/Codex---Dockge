function Get-WordDocumentText {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$WordContext
    )

    $doc = $WordContext.Application.Documents.Open($Path, $false, $true)
    try {
        return $doc.Content.Text
    }
    finally {
        $doc.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    }
}

function Add-ApplicationTextBlock {
    param(
        [Parameter(Mandatory)][System.Text.StringBuilder]$Builder,
        [Parameter(Mandatory)]$Application,
        [Parameter(Mandatory)][string]$FileLabel,
        [Parameter(Mandatory)][string]$FileName,
        [Parameter(Mandatory)][string]$Text
    )

    $Builder.AppendLine("============================================================") | Out-Null
    $Builder.AppendLine("Company: $($Application.Company)") | Out-Null
    $Builder.AppendLine("Job: $($Application.Job)") | Out-Null
    $Builder.AppendLine("${FileLabel}: $FileName") | Out-Null
    $Builder.AppendLine("Folder: $($Application.RelativePath)") | Out-Null
    $Builder.AppendLine("============================================================") | Out-Null
    $Builder.AppendLine($Text) | Out-Null
    $Builder.AppendLine() | Out-Null
    $Builder.AppendLine() | Out-Null
}

function Export-CombinedResumeText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Applications,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)]$WordContext
    )

    $builder = New-Object System.Text.StringBuilder
    $withResumes = @($Applications | Where-Object { $_.ResumePath })
    $i = 0

    foreach ($application in $withResumes) {
        $i++
        Write-Progress -Activity "Building resume text file" -Status "$($application.Company) - $($application.Job)" -PercentComplete (($i / [Math]::Max($withResumes.Count, 1)) * 100)
        Write-Log "Extracting resume text: $($application.ResumePath)" "INFO"

        Add-ApplicationTextBlock -Builder $builder `
            -Application $application `
            -FileLabel "File" `
            -FileName $application.ResumeFile `
            -Text (Get-WordDocumentText -Path $application.ResumePath -WordContext $WordContext)
    }

    Write-Progress -Activity "Building resume text file" -Completed
    $builder.ToString() | Set-Content -Path $OutputPath -Encoding UTF8
    Write-Log "Saved resume TXT: $OutputPath" "SUCCESS"
}

function Export-CombinedCoverLetterText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Applications,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)]$WordContext
    )

    $builder = New-Object System.Text.StringBuilder
    $withCoverLetters = @($Applications | Where-Object { $_.CoverLetterPath })
    $i = 0

    foreach ($application in $withCoverLetters) {
        $i++
        Write-Progress -Activity "Building cover letter text file" -Status "$($application.Company) - $($application.Job)" -PercentComplete (($i / [Math]::Max($withCoverLetters.Count, 1)) * 100)
        Write-Log "Extracting cover letter text: $($application.CoverLetterPath)" "INFO"

        Add-ApplicationTextBlock -Builder $builder `
            -Application $application `
            -FileLabel "File" `
            -FileName $application.CoverLetterFile `
            -Text (Get-WordDocumentText -Path $application.CoverLetterPath -WordContext $WordContext)
    }

    Write-Progress -Activity "Building cover letter text file" -Completed
    $builder.ToString() | Set-Content -Path $OutputPath -Encoding UTF8
    Write-Log "Saved cover letter TXT: $OutputPath" "SUCCESS"
}
