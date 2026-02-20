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
	$LinkPath = Join-Path $AOSMetadataPath $ModelFolder.Name
	if (Test-Path $LinkPath) {
		$item = Get-Item $LinkPath -Force
		if ($item.LinkType -eq 'SymbolicLink') {
			cmd /c rmdir "$LinkPath"
			Write-Host "Removed symlink: $LinkPath"
		} else {
			Write-Warning "$LinkPath is not a symbolic link — skipping."
		}
	} else {
		Write-Warning "Path not found: $LinkPath — skipping."
	}
}