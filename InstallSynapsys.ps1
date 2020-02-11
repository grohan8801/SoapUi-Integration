param ( $branch = "3.6" )  # ToDo:  this will be passed in

cls
$error.Clear()
$lastExitCode = 0

# get config from registry
$configProperties = Get-ItemProperty "HKLM:\SOFTWARE\Becton Dickinson\Synapsys\Config"
$fqdn = $configProperties.FQDN

$dbProperties = Get-ItemProperty "HKLM:\SOFTWARE\Becton Dickinson\Synapsys\Database"
$dataSource = $dbProperties.DataSource
$dbDefencePassword = $dbProperties.DbDefencePassword
$initialCatalog = $dbProperties.InitialCatalog
$dbUserID = $dbProperties.UserID
$encryptedDbPassword = $dbProperties.Password

$scriptPath = Split-Path $MyInvocation.MyCommand.Path -Parent
Add-Type -path "$scriptPath\Common.Core.dll"
$dbPassword = [BD.InfoStratus.Common.Utilities.EncryptionHelper]::Decrypt($encryptedDbPassword)

$installProperties = Get-ItemProperty "HKLM:\SOFTWARE\Becton Dickinson\Synapsys\Install"
$baseInstallerPath = $installProperties.'Source Folder'
if (-not $baseInstallerPath.EndsWith('\')) { $baseInstallerPath += '\' }

# locate installer
$fileObject = Get-ChildItem -Path "$baseInstallerPath\$branch\Latest Build" -Filter "Synapsys Installer $branch*.msi"
if ( $fileObject -ne $null ) { 
    $msiPath = $fileObject.FullName
} else {
    Write-Host "Couldn't locate a $branch Synapsys installer in `"$baseInstallerPath\$branch\Latest Build`"" 
    exit 1
}


# locate most recent lab file
$fileObject = Get-ChildItem -Path "$baseInstallerPath" -Filter "*.lab" | sort LastWriteTime | select -last 1
if ( $fileObject -ne $null ) { 
    $labFilepath = $fileObject.FullName
} else {
    Write-Host "Couldn't locate a lab file in `"$baseInstallerPath`"" 
    exit 2
}

# install
$msiArgs = "/qb /i `"$msiPath`" FQDN=`"$fqdn`" DB_APP_USER_PWD_ENC=`"$dbPassword`" IDM_DB_PWD_ENC=`"$dbPassword`" DB_DEFENCE_PWD_ENC=`"$dbDefencePassword`" LAB_SETTINGS_PATH=`"$labFilepath`" SKIP_VALIDATION=1"
Write-Host "msiexec.exe $msiArgs"
$exitCode = (Start-Process -FilePath "msiexec.exe" -ArgumentList $msiArgs -Wait -Passthru).ExitCode

if ($exitCode -ne 0)
{
    Write-Host "Error $exitCode installing '$msiPath' with lab file '$labFilepath'."
	exit $exitCode
}


# change SynapsysAdmin password to "default"
try {
    $sql = "UPDATE BD_IDM3.mr.UserAccounts SET IsLoginAllowed = 1, IsLoginAllowedChanged = NULL, LastFailedLogin = NULL, FailedLoginCount = 0, HashedPassword = 'C350.ALGwyGUAV+psEQtmJlpBl5Aa8mhtmJin50Uu88sQDeHIosbnCExqCAzGk65idIxEig==' WHERE Username = 'SynapsysAdmin'"

 	if ($dbUserId -ne "") {
		$output = Invoke-Sqlcmd -ServerInstance "$dataSource" -Database "$initialCatalog" -Username "$dbUserId" -Password "$dbPassword" -Query "$sql"
	} else {
		$output = Invoke-Sqlcmd -ServerInstance "$dataSource" -Database "$initialCatalog" -Query "$sql"
	}
} catch {
	Write-Host $error
    exit 4
}