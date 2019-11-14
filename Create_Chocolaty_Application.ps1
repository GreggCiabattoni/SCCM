#Gregg Ciabattoni
#9/17/2019
#Create Chocklaty Application 
#Example uses a install.ps1
#InstallationCommand    = "powershell -executionpolicy bypass -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command 'iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))' && SET 'PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"



[BOOL]$DEBUG=$false
$SiteCode = "TEST" # Site code 
$ProviderMachineName = "sccmserver.example.com" # SMS Provider machine name
$initParams = @{}

if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams #-ErrorAction SilentlyContinue
}

# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams
############################################################
Clear

$AppName 	= "Chocolatey"

$AppDesc 	=  "Chocolatey Software is focused on helping our community, customers, and partners with solutions that help fill the gaps that are often ignored. We offer a simple, pragmatic, and open approach to software managemen"
$AppPath 	= "TEST:\Application\Chocolatey"
$AppDTName  = "Install $AppName"



$NewApplicationParams = @{
	Name 					= $AppName
	Description 			= $AppDesc 
	Publisher 				= "Chocolatey Softeare" 
	ReleaseDate 			= (Get-Date) 
	AutoInstall 			= $true 
	SoftwareVersion 		= "0.10.15"
	LocalizedName 			= "Chocolatey"
	LocalizedDescription 	= $AppDesc 	
	IsFeatured  			= $true
}

#New-CMApplication @NewApplicationParams 


$NewDeploymentTypeParams = @{
	ApplicationName 		= $AppName 
	ContentLocation 		= "\\sccmserver\Packages\Chocolatey"
	#ContentLocation 		= ""
	DeploymentTypeName 		= $AppDTName
	AdministratorComment 	= $AppDesc 
	InstallCommand 			= "InstallChocolatey.psq"
	scriptfile 				= "\\sccmserver\Packages\Chocolatey\InstallChocolatey.ps1"
	ScriptLanguage 			= "Powershell"
	
}
	
New-CMApplication @NewApplicationParams 
Add-CMScriptDeploymentType  @NewDeploymentTypeParams 
Set-CMScriptDeploymentType -ApplicationName $AppName -DeploymentTypeName $AppDTName -Comment "UPDATED "
$App = Get-CMApplication $AppName
Move-CMObject -FolderPath $AppPath -InputObject  $App



