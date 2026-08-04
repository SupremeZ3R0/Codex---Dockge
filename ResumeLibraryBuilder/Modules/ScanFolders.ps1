function Test-TemporaryWordFile {
    param([Parameter(Mandatory)][System.IO.FileInfo]$File)
    return $File.Name -like '~$*'
}

function ConvertTo-NormalizedFileTitle {
    param([Parameter(Mandatory)][string]$Value)
    return ($Value.ToLowerInvariant() -replace '[^a-z0-9]+', ' ').Trim()
}

function Get-KeywordScore {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [Parameter(Mandatory)][string[]]$Keywords
    )

    $name = $File.BaseName.ToLowerInvariant()
    $score = 0
    foreach ($keyword in $Keywords) {
        if ($name -like "*$($keyword.ToLowerInvariant())*") { $score += 10 }
    }
    if ($File.Extension -ieq ".docx") { $score += 2 }
    if ($File.Extension -ieq ".doc") { $score += 1 }
    $score += [Math]::Min([int]($File.Length / 1KB), 5)
    return $score
}

function Select-ResumeFile {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo[]]$Files,
        [Parameter(Mandatory)][string]$ResumeTitle,
        [switch]$PreferDocx
    )

    $expectedTitle = ConvertTo-NormalizedFileTitle -Value $ResumeTitle
    $matches = $Files | Where-Object {
        $fileTitle = ConvertTo-NormalizedFileTitle -Value $_.BaseName
        $fileTitle -eq $expectedTitle -or $fileTitle -like "$expectedTitle*"
    }

    if (-not $matches) { return $null }

    $matches |
        Sort-Object @{ Expression = { if ((ConvertTo-NormalizedFileTitle -Value $_.BaseName) -eq $expectedTitle) { 0 } else { 1 } } },
                    @{ Expression = { if ($PreferDocx -and $_.Extension -ieq ".docx") { 0 } else { 1 } } },
                    LastWriteTime -Descending |
        Select-Object -First 1
}

function Select-BestFile {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo[]]$Files,
        [Parameter(Mandatory)][string[]]$Keywords,
        [switch]$PreferDocx
    )

    $matches = $Files | Where-Object {
        $file = $_
        $Keywords | Where-Object { $file.BaseName -like "*$_*" }
    }

    if (-not $matches) { return $null }

    $matches |
        Sort-Object @{ Expression = { Get-KeywordScore -File $_ -Keywords $Keywords }; Descending = $true },
                    @{ Expression = { if ($PreferDocx -and $_.Extension -ieq ".docx") { 0 } else { 1 } } },
                    LastWriteTime -Descending |
        Select-Object -First 1
}

function Find-JobApplications {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RootFolder,
        [Parameter(Mandatory)][string]$ResumeTitle,
        [Parameter(Mandatory)][string[]]$CoverLetterKeywords,
        [Parameter(Mandatory)][string[]]$SupportedExtensions,
        [hashtable]$ExistingRelativePaths,
        [switch]$PreferDocx
    )

    $root = (Resolve-Path $RootFolder).Path.TrimEnd('\', '/')
    $jobFolders = Get-ChildItem -Path $root -Directory -Recurse | Where-Object {
        $_.FullName.Substring($root.Length).TrimStart('\', '/').Split([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar).Count -ge 2
    }

    $applications = New-Object System.Collections.Generic.List[object]
    $i = 0

    foreach ($folder in $jobFolders) {
        $i++
        Write-Progress -Activity "Scanning job folders" -Status $folder.FullName -PercentComplete (($i / [Math]::Max($jobFolders.Count, 1)) * 100)

        $relative = $folder.FullName.Substring($root.Length).TrimStart('\', '/')
        if ($ExistingRelativePaths -and $ExistingRelativePaths.ContainsKey($relative)) {
            continue
        }

        $parts = $relative -split '[\\/]'
        $company = $parts[0]
        $job = $parts[1]

        # Treat only folders that are at least Company\Job as applications; nested folders are included
        # only when they actually contain candidate files.
        $files = Get-ChildItem -Path $folder.FullName -File | Where-Object {
            $SupportedExtensions -contains $_.Extension.ToLowerInvariant() -and -not (Test-TemporaryWordFile -File $_)
        }

        if (-not $files) { continue }

        $resume = Select-ResumeFile -Files $files -ResumeTitle $ResumeTitle -PreferDocx:$PreferDocx
        $cover = Select-BestFile -Files $files -Keywords $CoverLetterKeywords -PreferDocx:$PreferDocx

        if ($resume -or $cover) {
            $applications.Add([pscustomobject]@{
                Company = $company
                Job = $job
                Folder = $folder.FullName
                RelativePath = $relative
                ResumePath = if ($resume) { $resume.FullName } else { $null }
                CoverLetterPath = if ($cover) { $cover.FullName } else { $null }
                ResumeFile = if ($resume) { $resume.Name } else { $null }
                CoverLetterFile = if ($cover) { $cover.Name } else { $null }
            })
        }
    }

    Write-Progress -Activity "Scanning job folders" -Completed
    return $applications | Sort-Object Company, Job, RelativePath
}
