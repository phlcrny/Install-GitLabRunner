# Install-GitLabRunner

Automatically install or update gitlab-runner on Windows with hash-checking and optional backup of existing installs using Windows Powershell.

You might be able to get it to work on Powershell 6+ but I've not tried.

## Usage

```powershell
# Install or update to the latest version, no changes made if the latest version is already installed
.\Install-GitLabRunner.ps1

# Force install or update to the latest version, regardless of current version or hash mismatches, also with verbose messaging!
.\Install-GitLabRunner.ps1 -Force

# Install or update to the latest version, backs up the current version, no changes made if the latest version is already installed, and verbose messaging!
.\Install-GitLabRunner.ps1 -Backup -Verbose
```
