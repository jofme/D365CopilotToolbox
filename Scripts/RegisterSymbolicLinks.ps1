param(
	[string]$AOSMetadataPath = "C:\AOSService\PackagesLocalDirectory"
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $AOSMetadataPath)) {
	Write-Error "AOS metadata path not found: $AOSMetadataPath. Specify the correct path with -AOSMetadataPath."
	return
}

$RepoPath = Join-Path $PSScriptRoot ".."
$RepoMetadataPath = Join-Path $RepoPath "Metadata"
$RepoModelFolders = Get-ChildItem $RepoMetadataPath -Directory
foreach ($ModelFolder in $RepoModelFolders)
{
	$Target = Join-Path $RepoMetadataPath $ModelFolder.Name
	$LinkPath = Join-Path $AOSMetadataPath $ModelFolder.Name
	if (Test-Path $LinkPath) {
		Write-Warning "Symlink already exists: $LinkPath — skipping."
		continue
	}
	New-Item -ItemType SymbolicLink -Path $AOSMetadataPath -Name $ModelFolder.Name -Value $Target
	Write-Host "Created symlink: $LinkPath -> $Target"
}