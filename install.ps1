[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$SkillsRoot,
    [switch]$Force
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
$sourceRoot = [System.IO.Path]::GetFullPath($PSScriptRoot)

foreach ($skillName in $skillNames) {
    $source = Join-Path $sourceRoot $skillName
    if (-not (Test-Path -LiteralPath (Join-Path $source 'SKILL.md') -PathType Leaf)) {
        throw "Invalid package: $skillName/SKILL.md is missing."
    }

    $destination = Join-Path $SkillsRoot $skillName
    if ((Test-Path -LiteralPath $destination) -and -not $Force) {
        throw "Destination already exists: $destination. Re-run with -Force to replace it after creating a backup."
    }
}

if ($PSCmdlet.ShouldProcess($SkillsRoot, 'Install Study Docs Skills')) {
    New-Item -ItemType Directory -Path $SkillsRoot -Force | Out-Null
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'

    foreach ($skillName in $skillNames) {
        $source = Join-Path $sourceRoot $skillName
        $destination = Join-Path $SkillsRoot $skillName
        $backup = $null

        if (Test-Path -LiteralPath $destination) {
            $backup = "$destination.backup-$stamp"
            if (Test-Path -LiteralPath $backup) {
                throw "Backup destination already exists: $backup"
            }
            Move-Item -LiteralPath $destination -Destination $backup
        }

        try {
            Copy-Item -LiteralPath $source -Destination $destination -Recurse
            if (-not (Test-Path -LiteralPath (Join-Path $destination 'SKILL.md') -PathType Leaf)) {
                throw "Installed copy failed validation: $destination"
            }
        }
        catch {
            if (Test-Path -LiteralPath $destination) {
                Remove-Item -LiteralPath $destination -Recurse -Force
            }
            if ($backup -and (Test-Path -LiteralPath $backup)) {
                Move-Item -LiteralPath $backup -Destination $destination
            }
            throw
        }

        Write-Host "Installed $skillName -> $destination"
        if ($backup) {
            Write-Host "Previous copy preserved at $backup"
        }
    }

    $mineruScript = Join-Path $SkillsRoot 'mineru\scripts\mineru.py'
    $chunkerScript = Join-Path $SkillsRoot 'mineru\scripts\chunking.py'
    if (-not ((Test-Path -LiteralPath $mineruScript -PathType Leaf) -and (Test-Path -LiteralPath $chunkerScript -PathType Leaf))) {
        Write-Warning 'MinerU Skill is not installed beside these skills. Follow QUICKSTART.md before running ingestion.'
    }

    Write-Host 'Installation complete. Restart Codex if the skills are not detected automatically.'
}
