[CmdletBinding(DefaultParameterSetName = 'Convert')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Probe')]
    [switch]$Probe,

    [Parameter(Mandatory = $true, ParameterSetName = 'Convert')]
    [string]$Source,

    [Parameter(Mandatory = $true, ParameterSetName = 'Convert')]
    [string]$Destination
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

function Get-PowerPointExecutable {
    $candidates = @()
    foreach ($root in @(${env:ProgramFiles}, ${env:ProgramFiles(x86)})) {
        if ($root) {
            $candidates += Join-Path $root 'Microsoft Office\root\Office16\POWERPNT.EXE'
            $candidates += Join-Path $root 'Microsoft Office\Office16\POWERPNT.EXE'
            $candidates += Join-Path $root 'Microsoft Office\Office15\POWERPNT.EXE'
        }
    }
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }
    return $null
}

function Write-Result([hashtable]$Value) {
    [Console]::Out.WriteLine(($Value | ConvertTo-Json -Compress -Depth 4))
}

$executable = Get-PowerPointExecutable
$comType = [type]::GetTypeFromProgID('PowerPoint.Application', $false)
$available = ($null -ne $comType -and $null -ne $executable)
$fileVersion = $null
if ($executable) {
    $fileVersion = (Get-Item -LiteralPath $executable).VersionInfo.FileVersion
}

if ($Probe) {
    if (-not $available) {
        Write-Result @{
            available = $false
            backend = 'powerpoint'
            version = $fileVersion
            executable = $executable
            diagnostic = 'PowerPoint executable or COM registration is unavailable.'
        }
        exit 1
    }
    $probeBeforePids = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $probeApplication = $null
    $probeOwnsApplication = $false
    $probeResult = $null
    try {
        $probeApplication = New-Object -ComObject PowerPoint.Application
        Start-Sleep -Milliseconds 500
        $probeAfterPids = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $probeNewPids = @($probeAfterPids | Where-Object { $probeBeforePids -notcontains $_ })
        $probeOwnsApplication = ($probeNewPids.Count -eq 1)
        if (-not $probeOwnsApplication) {
            throw 'PowerPoint automation did not create an isolated process.'
        }
        $probeResult = @{
            available = $true
            backend = 'powerpoint'
            version = $fileVersion
            executable = $executable
            created_pids = $probeNewPids
        }
    }
    catch {
        $probeResult = @{
            available = $false
            backend = 'powerpoint'
            version = $fileVersion
            executable = $executable
            diagnostic = $_.Exception.Message
        }
    }
    finally {
        if ($null -ne $probeApplication) {
            if ($probeOwnsApplication) {
                try { $probeApplication.Quit() } catch { }
            }
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($probeApplication) } catch { }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    Write-Result $probeResult
    if ($probeResult.available) { exit 0 } else { exit 1 }
}

if (-not $available) {
    [Console]::Error.WriteLine('Microsoft PowerPoint desktop automation is unavailable.')
    exit 1
}

$sourcePath = [System.IO.Path]::GetFullPath($Source)
$destinationPath = [System.IO.Path]::GetFullPath($Destination)
if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
    [Console]::Error.WriteLine("Source PPTX does not exist: $sourcePath")
    exit 1
}
if ([System.IO.Path]::GetExtension($sourcePath) -ine '.pptx') {
    [Console]::Error.WriteLine('PowerPoint normalization accepts only .pptx input.')
    exit 1
}
$destinationDirectory = Split-Path -Parent $destinationPath
if (-not (Test-Path -LiteralPath $destinationDirectory -PathType Container)) {
    [Console]::Error.WriteLine("Destination directory does not exist: $destinationDirectory")
    exit 1
}
if ($sourcePath -eq $destinationPath) {
    [Console]::Error.WriteLine('Source and destination must be different paths.')
    exit 1
}

$beforePids = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$application = $null
$presentation = $null
$ownsApplication = $false
try {
    $application = New-Object -ComObject PowerPoint.Application
    Start-Sleep -Milliseconds 500
    $afterPids = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $newPids = @($afterPids | Where-Object { $beforePids -notcontains $_ })
    if ($newPids.Count -eq 1) {
        $ownsApplication = $true
    }
    elseif ($newPids.Count -gt 1) {
        $windowHandle = [int64]$application.HWND
        $matching = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Where-Object {
            [int64]$_.MainWindowHandle -eq $windowHandle -and $newPids -contains $_.Id
        })
        $ownsApplication = ($matching.Count -eq 1)
    }
    if (-not $ownsApplication) {
        throw 'PowerPoint did not create an isolated automation instance; export was aborted without touching an existing user instance.'
    }

    # Open(source, ReadOnly, Untitled, WithWindow). No save is ever issued to the PPTX.
    $presentation = $application.Presentations.Open($sourcePath, $true, $false, $false)
    $slideCount = [int]$presentation.Slides.Count
    # ppSaveAsPDF = 32. SaveCopyAs leaves the open source presentation bound to
    # its original path and never writes changes back to it.
    $presentation.SaveCopyAs($destinationPath, 32)
    $presentation.Close()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation)
    $presentation = $null

    Write-Result @{
        ok = $true
        backend = 'powerpoint'
        version = $fileVersion
        slide_count = $slideCount
        created_pids = $newPids
    }
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch { }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation) } catch { }
    }
    if ($null -ne $application) {
        if ($ownsApplication) {
            try { $application.Quit() } catch { }
        }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
