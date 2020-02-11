<#
	This script is used to pull down the build artifacts produced by the 'Synapsys MSI' pipeline.
	
	The artifacts are zip files containing the msi's created by the pipeline, one per installer.
	After pulling down the zip files from VSTS, the zip files are extracted and copied to a network
	drive for safekeeping and easier access:  \\psmd0319cifs101\rnddata\rndprojects\InfoStratus\Releases
	
	The script handles multiple builds against different branches.  It always processes refs/heads/dev.
	Additionally, it proecsses all branches following the naming convention refs/heads/Releases/<major.minor.revision>
	For example, refs/heads/Releases/3.1, refs/heads/Releases/3.1.5, etc.

	Each branch has an associated folder in the Releases share:
		* refs/heads/dev => 3.x Dev
		* refs/heads/Releases/3.1 => 3.1
		* refs/heads/Releases/3.1.5 => 3.1.5
		* etc.
#>


function Decrypt
{
    param ($encryptedString)

    # convert encrypted text back to a secure string object
    $secureString = ConvertTo-SecureString $encryptedString

    # decrypt secure string
    $marshal = [System.Runtime.InteropServices.Marshal]
    $ptr = $marshal::SecureStringToBSTR($secureString)
    $plainText = $marshal::PtrToStringBSTR($ptr)
    $marshal::ZeroFreeBSTR($ptr)

    return $plainText
}

function Create-FolderShortcut {
    param (
        [string] $folderName,
        [string] $shortcutFolder,
        [string] $targetParentFolder
    )

    $wScriptShell = New-Object -ComObject WScript.Shell
    $shortcut = $wScriptShell.CreateShortcut("$shortcutFolder\$folderName.lnk" )
    $shortcut.TargetPath = "$targetParentFolder\$folderName"
    $shortcut.IconLocation = "%SystemRoot%\system32\imageres.dll,137"
    $shortcut.Save()
}

function Write-Log {
    param( $message )

    $timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    Add-Content $logfile -Value "$timestamp $message"
    Write-Host "$timestamp $message"
}


# start 

cls
$error.Clear()

if (-Not (Test-Path Variable:PSise)) { # only run this in the console and not in the ISE
    $progressPreference = "SilentlyContinue"  # suppress progress bar for Invoke-Request downloads
}


$installProperties = Get-ItemProperty "HKLM:\SOFTWARE\Becton Dickinson\Synapsys\Install"
$releaseFolder = $installProperties.'Source Folder'
if (-not $releaseFolder.EndsWith('\')) { $releaseFolder += '\' }
$devBranchFolderName = $installProperties.'Dev Branch Folder'
$autoInstallBranch = $installProperties.'Auto-Install Branch'

$devOpsProperties = Get-ItemProperty "HKLM:\SOFTWARE\Becton Dickinson\Synapsys\Install\DevOps"
$synapsysMsiBuildDefinitionName = $devOpsProperties.'Build Definition'
$synapsysMsiBuildDefinitionNum = $devOpsProperties.'Build Definition Number'
$vsts_username = $devOpsProperties.'API User'
$encrypted_vsts_pat = $devOpsProperties.'API Password'
$apiVersion =$devOpsProperties.'API Version'
$organization = $devOpsProperties.Organization
$projectName = $devOpsProperties.ProjectName
$repositoryName = $devOpsProperties.RepositoryName
$baseUrl = "$($devOpsProperties.'API Base URL')/$organization"

$vsts_pat = Decrypt $encrypted_vsts_pat
$encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$($vsts_username):$vsts_pat"))
$authorizationHeader = "Basic $encodedCredentials"

# API version can be specified either in the header of the HTTP request or as a URL query parameter:
# https://docs.microsoft.com/en-us/azure/devops/integrate/how-to/call-rest-api?view=vsts
# e.g., $url = "$baseUrl/_apis/projects?api-version=$apiVersion
$headers = @{ Authorization = $authorizationHeader; Accept = $apiVersion }


$logfile = "$releaseFolder\InstallersDownload.log"
if (Test-Path $logfile) { Remove-Item $logfile -Force }

Write-Log "Get list of release branches for '$projectName' project"
$url = "$baseUrl/$projectName/_apis/git/repositories/$repositoryName/refs?filter=heads/releases"
$response = $null
$response = Invoke-RestMethod -Uri "$url" -Headers $headers

if ($response -ne $null) {
    $branches = @("refs/heads/dev")
    foreach ($branch in $response.value) {
        if (-not $branch.isLocked) { $branches += $branch.name }
    }
}

foreach ($branch in $branches) {
    Write-Log "Get last successful '$synapsysMsiBuildDefinitionName' build of the $branch' branch..."
    $url = "$baseUrl/$projectName/_apis/build/builds?definitions=$synapsysMsiBuildDefinitionNum&branchName=$branch&statusFilter=completed&resultFilter=succeeded&`$top=1"
    # Write-Log "`t$url"
    $response = $null
    $response = Invoke-RestMethod -Uri "$url" -Headers $headers

    if ($response.count -eq 0) { continue }  # no builds for this branch (yet)

    $lastBuildId = $response.value.id
    $lastBuildNumber = $response.value.buildNumber

    # clear out download folder
    $dirInfo = New-Item -Path $env:TEMP -Name $lastBuildNumber -ItemType Directory -Force
    $downloadFolder = $dirInfo.FullName
    Remove-Item "$downloadFolder\*" -Force -Recurse

    # create release folder, as necessary
    $destFolder = $releaseFolder
    if ($branch -eq "refs/heads/dev") {
        $destFolder += $devBranchFolderName
    } else {
        $destFolder += $branch.substring(20)  # omit "refs/heads/releases/"; by convention, the remainder should be the release #
    }

    if (!(Test-Path $destFolder)) {
        Write-Log "Creating directory '$destFolder'..."
        New-Item -ItemType directory -Path $destFolder | Out-Null

        Create-FolderShortcut -folderName "DBDefence Redistributable" -shortcutFolder $destFolder -targetParentFolder $documentationFolder
        Create-FolderShortcut -folderName "Installation Documentation" -shortcutFolder $destFolder -targetParentFolder $documentationFolder
    }

    # create/clear out Latest Build folder
    $destFolder += "\Latest Build"
    if (Test-Path $destFolder) {
        Write-Log "Deleting existing files in '$destFolder'..."
        Remove-Item -Path "$destFolder\*" -Recurse
    } else {
        Write-Log "Creating directory '$destFolder'..."
        New-Item -ItemType directory -Path $destFolder | Out-Null
    }

    Write-Log "Downloading installer artifacts for buildId $lastBuildId to '$downloadFolder'"
    $url = "$baseUrl/$projectName/_apis/build/builds/$lastBuildId/artifacts"
    # Write-Log "`t$url"
    $response = $null
    $response = Invoke-RestMethod -Uri "$url" -Headers $headers

    if ($response.count -gt 0) {
        $artifacts = $response.value

        # Expand-Archive cmdlet is only available in PowerShell 5.0, and above
        # so use .Net Framework to unzip to be safe
        Add-Type -AssemblyName System.IO.Compression.FileSystem

        $artifacts | foreach-object {
            $artifactId = $_.id
            $artifactName = $_.name
            $artifactUrl = $_.resource.downloadUrl

            $url = "$baseUrl/$projectName/_apis/build/builds/$lastBuildId/artifacts"
            $downloadPath = [System.IO.Path]::Combine($downloadFolder, "$artifactName.zip")

            Write-Log "`tDownloading $artifactName..."
            Invoke-WebRequest -Uri $artifactUrl -Headers $headers -OutFile $downloadPath 
            [System.IO.Compression.ZipFile]::ExtractToDirectory($downloadPath, $downloadFolder)
        }
    } else {
        Write-Warning "'$synapsysMsiBuildDefinitionName' build $lastBuildId has no artifacts"
        exit 1
    }

    # structure files to our liking/convention:
    # zip files have folders that contain the files.
    # We don't want these, so move the files and whack the folders

    Write-Log "Moving extracted files to their final resting place '$destFolder'..."
    $directories = Get-ChildItem -Path $downloadFolder -Directory
    $directories | foreach-object { 
        $files = Get-ChildItem -Path $_.FullName -File

        $files | foreach-object {
            if ($_.Name -eq 'readme.txt') {

                # generate readme file name
                $synapsysInstaller = $($files | where-object -Property Name -like "Synapsys Installer*")[0].Name  # don't care if it's the exe or msi
                $readmeName = [System.IO.Path]::ChangeExtension($synapsysInstaller, 'readme')

                Copy-Item -Path $_.FullName -Destination "$destFolder\$readmeName"

            } else {

                Copy-Item -Path $_.FullName -Destination $destFolder
            }
        }
    }

    Write-Log "Removing temporary download folder $downloadFolder"
    Remove-Item $downloadFolder -Force -Recurse  # remove download folder and its content

    # Write-Log "Extracting msi for FOD submission"
    # $msi = (Get-ChildItem -Path $destFolder -Filter "Synapsys Installer*.msi").Name
    # msiexec /qb /a `"$destFolder\$msi`" TARGETDIR=`"$releaseFolder\FoD`"
}


Write-Log "Installing latest version of Synapsys $autoInstallBranch..."
$scriptPath = Split-Path $MyInvocation.MyCommand.Path -Parent
. "$scriptPath\InstallSynapsys.ps1" -branch $autoInstallBranch

Write-Log "Done"