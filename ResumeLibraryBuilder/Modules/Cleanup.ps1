function Close-WordContext {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$WordContext)

    if ($WordContext.Application) {
        $WordContext.Application.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WordContext.Application) | Out-Null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
