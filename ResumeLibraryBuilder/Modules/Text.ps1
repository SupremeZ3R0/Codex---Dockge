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
        Write-Progress -Activity "Building text file" -Status "$($application.Company) - $($application.Job)" -PercentComplete (($i / [Math]::Max($withResumes.Count, 1)) * 100)
        Write-Log "Extracting text: $($application.ResumePath)" "INFO"

        $builder.AppendLine("============================================================") | Out-Null
        $builder.AppendLine("Company: $($application.Company)") | Out-Null
        $builder.AppendLine("Job: $($application.Job)") | Out-Null
        $builder.AppendLine("File: $($application.ResumeFile)") | Out-Null
        $builder.AppendLine("Folder: $($application.RelativePath)") | Out-Null
        $builder.AppendLine("============================================================") | Out-Null
        $builder.AppendLine((Get-WordDocumentText -Path $application.ResumePath -WordContext $WordContext)) | Out-Null
        $builder.AppendLine() | Out-Null
        $builder.AppendLine() | Out-Null
    }

    Write-Progress -Activity "Building text file" -Completed
    $builder.ToString() | Set-Content -Path $OutputPath -Encoding UTF8
    Write-Log "Saved TXT: $OutputPath" "SUCCESS"
}
