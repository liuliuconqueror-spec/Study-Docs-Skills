[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$SkillsRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($SkillsRoot)) {
    if (-not [string]::IsNullOrWhiteSpace($env:CODEX_HOME)) {
        $SkillsRoot = Join-Path $env:CODEX_HOME 'skills'
    }
    else {
        $SkillsRoot = Join-Path $HOME '.agents\skills'
    }
}

$SkillsRoot = [System.IO.Path]::GetFullPath($SkillsRoot)
$fileSystemRoot = [System.IO.Path]::GetPathRoot($SkillsRoot)
if ($SkillsRoot -ieq $fileSystemRoot) {
    throw "Refusing to use a filesystem root as SkillsRoot: $SkillsRoot"
}
$SkillsRoot = $SkillsRoot.TrimEnd(
    [System.IO.Path]::DirectorySeparatorChar,
    [System.IO.Path]::AltDirectorySeparatorChar
)
$skillNames = @('study-document-ingest', 'study-source-reader')

foreach ($skillName in $skillNames) {
    $destination = [System.IO.Path]::GetFullPath((Join-Path $SkillsRoot $skillName))
    if ([System.IO.Path]::GetDirectoryName($destination) -ne $SkillsRoot) {
        throw "Refusing unexpected uninstall target: $destination"
    }

    if (-not (Test-Path -LiteralPath $destination)) {
        Write-Host "Not installed: $destination"
        continue
    }

    if ($PSCmdlet.ShouldProcess($destination, 'Remove installed skill')) {
        Remove-Item -LiteralPath $destination -Recurse -Force
        Write-Host "Removed $destination"
    }
}

Write-Host 'Runtime state was preserved under ~/.codex/study-docs (or STUDY_DOCS_STATE_DIR).'
