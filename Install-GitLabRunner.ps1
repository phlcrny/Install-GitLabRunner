<#
.SYNOPSIS
	Installs or updates gitlab-runner, after retrieving release details from the GitLab API, using basic Windows Powershell cmdlets.
.DESCRIPTION
	Using Invoke-RestMethod, Invoke-WebRequest, Get-FileHash
.EXAMPLE
	.\Install-GitLabRunner.ps1

	Installs with default values.
.EXAMPLE
	.\Install-GitLabRunner.ps1 -Force

	-Force is used to override any guardrail functions (hash mismatches, latest version already installed etc) and force the script to proceed.
.EXAMPLE
	.\Install-GitLabRunner.ps1 -Backup -Verbose

	Backs up any existing version of gitlab-runner and provides verbose messaging. One for the cautious.
.PARAMETER Path
	The directory in which gitlab-runner will be installed - the script assumes the executable is/will be named gitlab-runner.exe
.PARAMETER DownloadDirectory
	The location the executable will be downloaded to - defaults to a generated subdirectory of $ENV:Temp
.PARAMETER ApiEndpoint
	The GitLab project release API endpoint used to retrieve release details
.PARAMETER AllowPrerelease
	Indicates if prerelease versions should be considered for installation or not
.PARAMETER Backup
	Indicates if an existing version of the gitlab-runner executable in the install location should be backed up rather than overwritten
.PARAMETER Force
	Indicates that guardrail functions, such as hash or version checks, will be ignored
.INPUTS
	Strings, switches
.OUTPUTS
	N/A
.LINK
	https://gitlab.com/gitlab-org/gitlab-runner/-/releases
#>
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = 'Medium')]
param
(
	[Parameter(Mandatory = $false, HelpMessage = 'The directory in which gitlab-runner will be installed (assumes the executable is named gitlab-runner.exe)')]
	[alias('InstallLocation')]
	[string] $Path = "$ENV:ProgramFiles\gitlab-runner\",

	[Parameter(Mandatory = $false, HelpMessage = 'The location the executable will be downloaded to - defaults to a generated subdirectory of $ENV:Temp')]
	[string] $DownloadDirectory,

	[Parameter(Mandatory = $false, HelpMessage = 'The GitLab project release API endpoint')]
	[string] $ApiEndpoint = 'https://gitlab.com/api/v4/projects/250833/releases',

	[Parameter(Mandatory = $false, HelpMessage = 'Allow pre-release/release candidate versions')]
	[switch] $AllowPrerelease = $False,

	[Parameter(Mandatory = $false, HelpMessage = 'Backs up the existing gitlab-runner file')]
	[switch] $Backup,

	[Parameter(Mandatory = $false, HelpMessage = 'Force installation, ignores hash mismatches or standard break points')]
	[switch] $Force
)

#region Reusables
$IwrSplat = @{
	UseBasicParsing = $True
	ErrorAction 	= 'Stop'
	Verbose 		= $False
}
$ProcessSplat = @{
	NoNewWindow = $True
	Wait = $True
}
$InstallLocation = Join-Path -Path $Path -ChildPath 'gitlab-runner.exe'
#endregion

#region Check for admin
$CurrentUserID = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$WindowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($CurrentUserID)
$Admin = [System.Security.Principal.WindowsBuiltInRole]::Administrator
$IsAdmin = $WindowsPrincipal.IsInRole($Admin)
if (-not $IsAdmin)
{
	Write-Warning -Message 'You probably need to be admin to run this successfully, but crack on if you want'
}
#endregion

#region Retrieve release
try
{
	Write-Verbose 'Query GitLab API for releases'
	$ApiResponse = (Invoke-WebRequest -Uri $ApiEndpoint @IwrSplat).Content | ConvertFrom-Json
}
catch
{
	$PSCmdlet.ThrowTerminatingError($_)
}

Write-Verbose -Message 'Retrieving releases before sorting'
$Releases = $ApiResponse |
	Sort-Object 'created_at' -Descending |
	Select-Object 'name', 'tag_name', 'created_at', @{Label = 'version'; Expression = { [version] ($_.tag_name -replace '^v' -replace '\-rc(\d)+') } }

$LatestRelease = if ($AllowPrerelease)
{
	Write-Verbose -Message 'Selecting latest release by release creation date'
	$Releases | Sort-Object 'Version' -Descending | Select-Object -First 1
}
else
{
	Write-Verbose -Message 'Selecting latest non pre-release/rc version'
	$Releases | Where-Object 'name' -NotMatch 'rc' | Sort-Object 'Version' -Descending | Select-Object -First 1
}
Write-Verbose -Message "Latest release: $($LatestRelease.tag_name)"
#endregion

#region Retrieve installed version
Write-Verbose -Message "Checking for existing installation: '$InstallLocation'"
if (Test-Path -LiteralPath $InstallLocation)
{
	# Retrieve the version of the installed gitlab-runner.exe from $InstallLocation, regex it
	# convert it to a string then cast it to a version so we can compare it to the retrieved release version
	# (unless we've allowed pre-release tag then we're freewheeling a bit)
	$InstalledVersion = [version] ((. $InstallLocation -v |
		Select-String -Pattern '^Version\:.+$') -split ' +' |
		Select-Object -Last 1).ToString()
	Write-Verbose -Message "Existing installation: v$($InstalledVersion.ToString())"

	if (($LatestRelease.version -le $InstalledVersion) -and
		($LatestRelease.tag_name -notmatch '\-rc(\d)+'))
	{
		if ($Force)
		{
			Write-Verbose -Message 'Installed version already matches latest release. Installing anyway (-Force)'
		}
		else
		{
			Write-Verbose -Message 'Latest version is already installed!'
			Break
		}

	}
	elseif (($AllowPrerelease -eq $True) -and ($LatestRelease.tag_name -match '\-rc(\d)+'))
	{
		Write-Verbose -Message "Retrieved version is a prerelease and may supersede the installed version: $($LatestRelease.tag_name)"
	}
	else
	{
		Write-Warning -Message 'Unable to compare latest version and installed version.'
	}
}
#endregion

#region Download Prep
$DownloadFile = "gitlab-runner-windows-amd64-$($LatestRelease.tag_name).exe"
if (-not $DownloadDirectory)
{
	Write-Verbose -Message "Preparing download location "
	$DownloadDirectory = Join-Path -Path $ENV:Temp -ChildPath "gitlab-runner-download-$(Get-Date -Format 'yyyMMdd')"
}
if (-not (Test-Path -LiteralPath $DownloadDirectory))
{
	Write-Verbose -Message "Creating downloading directory '$DownloadDirectory'"
	[void] (New-Item -Path $DownloadDirectory -ItemType 'Directory')
}

try
{
	Write-Verbose -Message 'Setting download destination'
	$DownloadDestination = Join-Path -Path $DownloadDirectory -ChildPath $DownloadFile
}
catch
{
	$PSCmdlet.ThrowTerminatingError($_)
}

try
{
	Write-Verbose -Message 'Retrieving download details'
	$DownloadPageUrl = "https://gitlab-runner-downloads.s3.amazonaws.com/$($LatestRelease.tag_name)/index.html"
	$DownloadPage = Invoke-WebRequest -Uri $DownloadPageUrl @IwrSplat
	$DownloadBlock = ($DownloadPage.Content -split '<(\/)?li>' | Select-String -Pattern 'windows-amd64\.exe' -Context 1).Line
	# This is hacky but fits a fairly established pattern for relative links in these release pages.
	# I'm comfortable with it when the immediate alternative is using regex to parse tags.
	Write-Verbose -Message "Inferring download URL"
	$InferredDownloadLink = $DownloadPageUrl -replace 'index.html$', 'binaries/gitlab-runner-windows-amd64.exe'
	Write-Verbose -Message "Testing inferred download link: '$InferredDownloadLink'"
	$TestRequest = Invoke-WebRequest -Uri $InferredDownloadLink -DisableKeepAlive -Method 'HEAD' @IwrSplat
	if ($TestRequest.StatusCode -eq 200)
	{
		Write-Verbose -Message 'Inferred download link returns 200 OK, treating as valid download link'
		$DownloadUrl = $InferredDownloadLink
	}
	Write-Verbose -Message 'Retrieving expected file hash'
	$ExpectedFileHash = (Select-String -InputObject $DownloadBlock -Pattern '([a-z0-9]{64})').Matches.Value.ToLower()
}
catch
{
	$PSCmdlet.ThrowTerminatingError($_)
}
#endregion

#region Download
try
{
	Write-Verbose -Message "Downloading '$($LatestRelease.tag_name)' from '$DownloadUrl'"
	Invoke-WebRequest -Uri $DownloadUrl -OutFile $DownloadDestination @IwrSplat
	Write-Verbose -Message "Unblocking downloaded file"
	Unblock-File -LiteralPath $DownloadDestination
}
catch
{
	$PSCmdlet.ThrowTerminatingError($_)
}

try
{
	Write-Verbose -Message "Retrieving SHA256 hash of downloaded file: '$DownloadDestination'"
	$DownloadHash = (Get-FileHash -LiteralPath $DownloadDestination -Algorithm 'SHA256').Hash.ToLower()
	if (($Null -ne $ExpectedFileHash) -and ($DownloadHash -eq $ExpectedFileHash))
	{
		Write-Verbose -Message 'Downloaded and expected hashes match'
		Write-Verbose -Message "Expected: $ExpectedFileHash"
		Write-Verbose -Message "Download: $DownloadHash"
	}
	else
	{
		Write-Warning -Message "Hashes don't match!"
		Write-Warning -Message "Expected: $ExpectedFileHash"
		Write-Warning -Message "Download: $DownloadHash"
		if (-not $Force)
		{
			Write-Warning -Message 'Skipping install'
			Write-Warning -Message 'Removing bad downloaded file'
			Remove-Item -LiteralPath $DownloadDestination -Force
			Break
		}
		else
		{
			Write-Verbose -Message 'Continuing with install as Force is specified'
		}
	}
}
catch
{
	$PSCmdlet.ThrowTerminatingError($_)
}
#endregion

#region Install Prep
Write-Verbose -Message "Ensuring installation directory exists"
if (-not (Test-Path -LiteralPath $Path))
{
	Write-Verbose -Message "Creating installation directory: '$Path'"
	[void] (New-Item -LiteralPath $Path -ItemType 'Directory')
}

# Upgrade Prep
if (Test-Path -LiteralPath $InstallLocation)
{
	Write-Verbose -Message 'Retrieving file hash of existing installation'
	$InstalledVersionHash = (Get-FileHash -LiteralPath $InstallLocation -Algorithm 'SHA256').Hash.ToLower()
	if (($Null -ne $InstalledVersionHash) -and ($DownloadHash -eq $InstalledVersionHash))
	{
		Write-Warning -Message 'Current and downloaded hashes match'
		Write-Warning -Message "Current : $InstalledVersionHash"
		Write-Warning -Message "Download: $DownloadHash"
		if (-not $Force)
		{
			Write-Warning -Message 'Skipping install'
			Break
		}
		else
		{
			Write-Verbose -Message 'Continuing with install as Force is specified'
		}
	}
	else
	{
		Write-Verbose -Message 'Current and downloaded hashes match. Installing downloaded version'
		Write-Verbose -Message "Current: $InstalledVersionHash"
		Write-Verbose -Message "Download: $DownloadHash"
	}

	if (Get-Service -Name 'gitlab-runner')
	{
		try
		{
			$ServiceAlreadyInstalled = $True
			Write-Verbose -Message 'Retrieving installed service'
			$InstalledService = Get-CimInstance -Query "SELECT * FROM Win32_Service WHERE name='gitlab-runner'" | Where-Object 'PathName' -Match 'gitlab-runner'
			# Extract the service/installation path from the PathName property which also contains arguments etc.
			# We would just split on spaces and drop quotes then take the first element as the path, but 'Program Files' is likely to screw us, so we'll work around that.
			# This is still fragile to say the least but I'm comfortable making assumptions on paths and accepting edge cases for now.
			Write-Verbose -Message 'Retrieving installed service path'
			$InstalledServicePath = $InstalledService.PathName -replace 'Program Files', 'ProgramFiles' -split ' ' -replace '"' -replace 'ProgramFiles', 'Program Files' | Select-Object -First 1
			Write-Debug -Message 'Adding installed service path to $ProcessSplat'
			$ProcessSplat.Add('FilePath', $InstallLocation)
		}
		catch
		{
			$PSCmdlet.ThrowTerminatingError($_)
		}

		if ($InstalledServicePath -like $InstallLocation)
		{
			try
			{
				Write-Verbose -Message "Stopping gitlab-runner service in '$InstallLocation'"
				Start-Process -ArgumentList 'stop' @ProcessSplat
			}
			catch
			{
				$PSCmdlet.ThrowTerminatingError($_)
			}
		}
		else
		{
			Write-Warning -Message "Second gitlab-runner service found installed in '$InstalledServicePath'"
		}
	}

	if ($Backup)
	{
		try
		{
			try
			{
				$BackupLocation = $InstallLocation -replace '.exe$', "-$(Get-Date -Format 'yyyMMdd').exe.bak"
				Write-Verbose -Message "Backing up existing executable to '$BackupLocation'"
				Copy-Item -LiteralPath $InstallLocation -Destination $BackupLocation
			}
			catch
			{
				$PSCmdlet.ThrowTerminatingError($_)
			}
		}
		catch
		{
			$PSCmdlet.ThrowTerminatingError($_)
		}
	}
}
#endregion

#region Install
try
{
	Write-Verbose -Message "Moving downloaded file to '$InstallLocation'"
	Move-Item -LiteralPath $DownloadDestination -Destination $InstallLocation -Force
}
catch
{
	$PSCmdlet.ThrowTerminatingError($_)
}

try
{
	if ($ServiceAlreadyInstalled)
	{
		# gitlab-runner parameters are case sensitive for some reason.
		Start-Process -ArgumentList 'status' @ProcessSplat
		Write-Verbose -Message "Starting gitlab-runner: '$InstalledServicePath'"
		Start-Process -ArgumentList 'start' @ProcessSplat
		Start-Process -ArgumentList 'status' @ProcessSplat
	}
	else
	{
		Start-Process -ArgumentList 'status' @ProcessSplat
		Write-Verbose -Message "Installing gitlab-runner: '$InstalledServicePath'"
		Start-Process -ArgumentList 'install' @ProcessSplat
		Start-Process -ArgumentList 'status' @ProcessSplat
	}
}
catch
{
	$PSCmdlet.ThrowTerminatingError($_)
}
finally
{
	Write-Verbose -Message 'Outputting version details'
	Start-Process -ArgumentList '--version' @ProcessSplat
}
#endregion