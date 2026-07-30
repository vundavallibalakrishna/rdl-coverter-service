param(
    [Parameter(Mandatory = $true)][string]$InputDocx,
    [Parameter(Mandatory = $true)][string]$OutputPdf,
    [Parameter(Mandatory = $true)][string]$ResultJson
)

$ErrorActionPreference = "Stop"
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repositoryRoot = [System.IO.Path]::GetFullPath((Join-Path $scriptRoot "..\..\.."))
$tmpRoot = [System.IO.Path]::GetFullPath((Join-Path $repositoryRoot "tmp"))
$outputDirectory = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($OutputPdf))
$resultDirectory = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($ResultJson))
if ($outputDirectory -ne $tmpRoot -or $resultDirectory -ne $tmpRoot) {
    throw "Word certification artifacts must be written directly under repository tmp/"
}
$word = $null
$document = $null
$started = Get-Date
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $word.Options.UpdateLinksAtOpen = $false
    $word.Options.SaveNormalPrompt = $false
    $document = $word.Documents.Open(
        [System.IO.Path]::GetFullPath($InputDocx),
        $false,
        $false,
        $false,
        "",
        "",
        $false,
        "",
        "",
        0,
        65001,
        $true,
        $false,
        0,
        $false,
        $false
    )
    $document.Repaginate()
    $pageCount = $document.ComputeStatistics(2)
    $document.ExportAsFixedFormat(
        [System.IO.Path]::GetFullPath($OutputPdf),
        17,
        $false,
        0,
        0,
        1,
        $pageCount,
        0,
        $true,
        $true,
        1,
        $true,
        $true,
        $false
    )
    @{
        passed = $true
        pageCount = $pageCount
        wordVersion = $word.Version
        compatibilityMode = $document.CompatibilityMode
        openedReadOnly = $document.ReadOnly
        openAndRepair = $false
        updateLinksAtOpen = $false
        displayAlerts = 0
        automationSecurity = 3
        elapsedMilliseconds = [int]((Get-Date) - $started).TotalMilliseconds
        input = [System.IO.Path]::GetFileName($InputDocx)
        output = [System.IO.Path]::GetFileName($OutputPdf)
    } | ConvertTo-Json | Set-Content -Encoding UTF8 $ResultJson
}
catch {
    @{
        passed = $false
        error = $_.Exception.Message
        elapsedMilliseconds = [int]((Get-Date) - $started).TotalMilliseconds
        input = [System.IO.Path]::GetFileName($InputDocx)
    } | ConvertTo-Json | Set-Content -Encoding UTF8 $ResultJson
    throw
}
finally {
    if ($document -ne $null) {
        $document.Close(0)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document)
    }
    if ($word -ne $null) {
        $word.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
