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

function Convert-InchesToWordPoints {
    param([Parameter(Mandatory)][double]$Inches)
    return $Inches * 72
}

function Set-DocumentMargins {
    param(
        [Parameter(Mandatory)]$Document,
        [Parameter(Mandatory)][double]$TopInches,
        [Parameter(Mandatory)][double]$BottomInches,
        [Parameter(Mandatory)][double]$LeftInches,
        [Parameter(Mandatory)][double]$RightInches
    )

    foreach ($section in $Document.Sections) {
        $section.PageSetup.TopMargin = Convert-InchesToWordPoints -Inches $TopInches
        $section.PageSetup.BottomMargin = Convert-InchesToWordPoints -Inches $BottomInches
        $section.PageSetup.LeftMargin = Convert-InchesToWordPoints -Inches $LeftInches
        $section.PageSetup.RightMargin = Convert-InchesToWordPoints -Inches $RightInches
    }
}


function Remove-TrailingBlankContent {
    param([Parameter(Mandatory)]$Document)

    # Word documents always need a final paragraph mark. This removes only extra
    # trailing blank paragraphs and manual page/section breaks that can create
    # blank pages between inserted files.
    $maxDeletes = 20
    while ($maxDeletes -gt 0 -and $Document.Content.End -gt 3) {
        $end = $Document.Content.End
        $lastCharacter = $Document.Range($end - 2, $end - 1)
        $lastText = $lastCharacter.Text
        $lastTwoCharacters = $Document.Range($end - 3, $end - 1).Text

        if ($lastText -eq ([string][char]12)) {
            $lastCharacter.Delete() | Out-Null
        }
        elseif ($lastTwoCharacters -eq "`r`r") {
            $lastCharacter.Delete() | Out-Null
        }
        else {
            break
        }

        $maxDeletes--
    }
}

function Add-SimpleApplicationHeading {
    param(
        [Parameter(Mandatory)]$Document,
        [Parameter(Mandatory)]$Application
    )

    Remove-TrailingBlankContent -Document $Document

    $shouldStartNewPage = $Document.Content.End -gt 1
    $range = $Document.Range()
    $range.Collapse(0)

    $headingStart = $range.Start
    $range.InsertAfter("$($Application.Company) - $($Application.Job)`r`r")

    $headingRange = $Document.Range($headingStart, $range.End)
    $headingRange.Font.Bold = $true
    $headingRange.ParagraphFormat.Alignment = 0
    $headingRange.ParagraphFormat.LeftIndent = 0
    $headingRange.ParagraphFormat.FirstLineIndent = 0
    $headingRange.ParagraphFormat.SpaceBefore = 0
    $headingRange.ParagraphFormat.SpaceAfter = 0
    $headingRange.ParagraphFormat.LineSpacingRule = 0
    $headingRange.ParagraphFormat.TabStops.ClearAll()

    # Start the heading on a new page without inserting a standalone manual
    # page-break paragraph, which can show up as a blank page with one line.
    $headingRange.Paragraphs.Item(1).Format.PageBreakBefore = $shouldStartNewPage
}

function Export-CombinedResumeDocx {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Applications,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)]$WordContext,
        [Parameter(Mandatory)][double]$MarginTopInches,
        [Parameter(Mandatory)][double]$MarginBottomInches,
        [Parameter(Mandatory)][double]$MarginLeftInches,
        [Parameter(Mandatory)][double]$MarginRightInches,
        [switch]$Append
    )

    $word = $WordContext.Application
    $document = if ($Append -and (Test-Path -LiteralPath $OutputPath)) {
        Write-Log "Opening existing resume DOCX for append: $OutputPath" "INFO"
        $word.Documents.Open($OutputPath)
    }
    else {
        $word.Documents.Add()
    }

    try {
        Set-DocumentMargins -Document $document `
            -TopInches $MarginTopInches `
            -BottomInches $MarginBottomInches `
            -LeftInches $MarginLeftInches `
            -RightInches $MarginRightInches

        $withResumes = @($Applications | Where-Object { $_.ResumePath })
        $i = 0
        foreach ($application in $withResumes) {
            $i++
            Write-Progress -Activity "Building resume Word document" -Status "$($application.Company) - $($application.Job)" -PercentComplete (($i / [Math]::Max($withResumes.Count, 1)) * 100)
            Write-Log "Adding resume to DOCX: $($application.ResumePath)" "INFO"

            Add-SimpleApplicationHeading -Document $document -Application $application
            $range = $document.Range()
            $range.Collapse(0)
            $range.InsertFile($application.ResumePath)
            Remove-TrailingBlankContent -Document $document
        }

        Set-DocumentMargins -Document $document `
            -TopInches $MarginTopInches `
            -BottomInches $MarginBottomInches `
            -LeftInches $MarginLeftInches `
            -RightInches $MarginRightInches

        Write-Progress -Activity "Building resume Word document" -Completed
        $document.SaveAs([ref]$OutputPath)
        Write-Log "Saved resume DOCX: $OutputPath" "SUCCESS"
    }
    finally {
        $document.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
    }
}

function Export-CombinedCoverLetterDocx {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Applications,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)]$WordContext,
        [Parameter(Mandatory)][double]$MarginTopInches,
        [Parameter(Mandatory)][double]$MarginBottomInches,
        [Parameter(Mandatory)][double]$MarginLeftInches,
        [Parameter(Mandatory)][double]$MarginRightInches,
        [switch]$Append
    )

    $word = $WordContext.Application
    $document = if ($Append -and (Test-Path -LiteralPath $OutputPath)) {
        Write-Log "Opening existing cover letter DOCX for append: $OutputPath" "INFO"
        $word.Documents.Open($OutputPath)
    }
    else {
        $word.Documents.Add()
    }

    try {
        Set-DocumentMargins -Document $document `
            -TopInches $MarginTopInches `
            -BottomInches $MarginBottomInches `
            -LeftInches $MarginLeftInches `
            -RightInches $MarginRightInches

        $withCoverLetters = @($Applications | Where-Object { $_.CoverLetterPath })
        $i = 0
        foreach ($application in $withCoverLetters) {
            $i++
            Write-Progress -Activity "Building cover letter Word document" -Status "$($application.Company) - $($application.Job)" -PercentComplete (($i / [Math]::Max($withCoverLetters.Count, 1)) * 100)
            Write-Log "Adding cover letter to DOCX: $($application.CoverLetterPath)" "INFO"

            Add-SimpleApplicationHeading -Document $document -Application $application
            $range = $document.Range()
            $range.Collapse(0)
            $range.InsertFile($application.CoverLetterPath)
            Remove-TrailingBlankContent -Document $document
        }

        Set-DocumentMargins -Document $document `
            -TopInches $MarginTopInches `
            -BottomInches $MarginBottomInches `
            -LeftInches $MarginLeftInches `
            -RightInches $MarginRightInches

        Write-Progress -Activity "Building cover letter Word document" -Completed
        $document.SaveAs([ref]$OutputPath)
        Write-Log "Saved cover letter DOCX: $OutputPath" "SUCCESS"
    }
    finally {
        $document.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
    }
}
