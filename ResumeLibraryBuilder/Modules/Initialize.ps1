function Initialize-ResumeLibraryBuild {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RootFolder,
        [Parameter(Mandatory)][string]$OutputFolder
    )

    if (-not (Test-Path -Path $RootFolder -PathType Container)) {
        throw "Root folder does not exist: $RootFolder"
    }

    if (-not (Test-Path -Path $OutputFolder -PathType Container)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
    }

    $script:LogPath = Join-Path $OutputFolder "CombineLog.txt"
    Set-Content -Path $script:LogPath -Value "ResumeLibraryBuilder log" -Encoding UTF8

    [pscustomobject]@{
        RootFolder = (Resolve-Path $RootFolder).Path
        OutputFolder = (Resolve-Path $OutputFolder).Path
        StartedAt = Get-Date
    }
}
