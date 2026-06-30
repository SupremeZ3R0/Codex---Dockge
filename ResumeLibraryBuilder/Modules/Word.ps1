function New-WordContext {
    [CmdletBinding()]
    param()

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    [pscustomobject]@{
        Application = $word
    }
}

function Add-ResumeSeparator {
    param(
        [Parameter(Mandatory)]$Document,
        [Parameter(Mandatory)]$Application
    )

    $range = $Document.Range()
    $range.Collapse(0)
    if ($Document.Content.End -gt 1) { $range.InsertBreak(7) }

    $range.InsertAfter("============================================================`r")
    $range.InsertAfter("Company: $($Application.Company)`r")
    $range.InsertAfter("Job: $($Application.Job)`r")
    $range.InsertAfter("File: $($Application.ResumeFile)`r")
    $range.InsertAfter("Folder: $($Application.RelativePath)`r")
    $range.InsertAfter("============================================================`r`r")
}

function Export-CombinedResumeDocx {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Applications,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)]$WordContext
    )

    $word = $WordContext.Application
    $document = $word.Documents.Add()

    try {
        $withResumes = @($Applications | Where-Object { $_.ResumePath })
        $i = 0
        foreach ($application in $withResumes) {
            $i++
            Write-Progress -Activity "Building Word document" -Status "$($application.Company) - $($application.Job)" -PercentComplete (($i / [Math]::Max($withResumes.Count, 1)) * 100)
            Write-Log "Adding resume to DOCX: $($application.ResumePath)" "INFO"

            Add-ResumeSeparator -Document $document -Application $application
            $range = $document.Range()
            $range.Collapse(0)
            $range.InsertFile($application.ResumePath)
        }

        Write-Progress -Activity "Building Word document" -Completed
        $document.SaveAs([ref]$OutputPath)
        Write-Log "Saved DOCX: $OutputPath" "SUCCESS"
    }
    finally {
        $document.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
    }
}
