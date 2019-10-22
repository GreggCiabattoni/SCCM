#Gregg Ciabattoni
#2/22/2018
#Download new msi's to stageing location
#Get MSI Product Code
#Create Folder with version name on Distribution Point
#Update Deployment Type Content Location
#Update Deployment Type Detection Method's MSI PRoduct Code
#Update Version In Application
#Update Date Published In Application
#Redistibute Content
#Update distibution point
Clear

#Load SCCM Module
$SiteCode = "TES" # Site code 
$ProviderMachineName = "myserver.example.com" # SMS Provider machine name
$initParams = @{}

if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}

# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams

###########################################################################################
#Get-CMSite
$ApplicationName= "Google Chrome Update TEST" 
##$DTNAME="TEST Install Google Chrome"
$DP = "\\dp\share\Google_Chrome\"
$TempDirectory = "C:\Temp"
[BOOL]$DEBUG=$false
$START=Get-Date
$ScriptName="Update_Google_Chrome_Deployments.ps1"
Write-Host -BackgroundColor Black -ForegroundColor White "Script Started" (Get-Date) "DEBUG = $DEBUG"
###########################################################################################
#Functions	
Function GetLatestGoogleVersion(){
$WebResponse = Invoke-WebRequest "https://omahaproxy.appspot.com/all" 
$csv = $WebResponse.Content | convertfrom-csv
return ($csv | ?{($_.os -eq 'win64') -and ($_.channel -eq 'stable')}).current_version
}

function Get-FileMetaData 
{ 
    param([Parameter(Mandatory=$True)][string]$File = $(throw "Parameter -File is required.")) 
 
    if(!(Test-Path -Path $File)) 
    { 
        throw "File does not exist: $File" 
        Exit 1 
    } 
 
    $tmp = Get-ChildItem $File 
    $pathname = $tmp.DirectoryName 
    $filename = $tmp.Name 
 
    $shellobj = New-Object -ComObject Shell.Application 
    $folderobj = $shellobj.namespace($pathname) 
    $fileobj = $folderobj.parsename($filename) 
    $results = New-Object PSOBJECT 
    for($a=0; $a -le 294; $a++) 
    { 
        if($folderobj.getDetailsOf($folderobj, $a) -and $folderobj.getDetailsOf($fileobj, $a))  
        { 
            $hash += @{$($folderobj.getDetailsOf($folderobj, $a)) = $($folderobj.getDetailsOf($fileobj, $a))} 
            $results | Add-Member $hash -Force 
        } 
    } 
    return $results 
}

Function Download-Chrome {    
    # Test internet connection
    if (Test-Connection google.com -Count 3 -Quiet) {
		 foreach($file in ('GoogleChromeStandaloneEnterprise64.msi', 'GoogleChromeStandaloneEnterprise.msi')){
				$lnk = "http://dl.google.com/edgedl/chrome/install/" + $file
				Write-Host -ForegroundColor Gray "Downloading Google Chrome $file ... $lnk " -NoNewLine
				if(Test-Path $TempDirectory\$file){
					Write-Host -ForegroundColor DarkGreen "Removing $TempDirectory\$file"
					Remove-Item -Path $TempDirectory\$file -Force 
				}else{
					Write-Host -ForegroundColor Green "$TempDirectory\$file does not exist"
				}

				# Download the installer from Google
		        try {
			        New-Item -ItemType Directory "$TempDirectory" -Force | Out-Null
			        (New-Object System.Net.WebClient).DownloadFile($lnk, "$TempDirectory\$file")
		            Write-Host 'success!' -ForegroundColor Green
		        } catch {
			        Write-Host 'failed. There was a problem with the download.' -ForegroundColor Red
		            if ($RunScriptSilent -NE $True){
		                Read-Host 'Press [Enter] to exit'
		            }
			        exit
		        }
		}
	} else {
		        Write-Host "failed. Unable to connect to Google's servers." -ForegroundColor Red
		        if ($RunScriptSilent -NE $True){
		            Read-Host 'Press [Enter] to exit'
		        }
			    exit
		    }
		
	Write-Host -ForegroundColor Gray "Download Complete"
}

Function GetMSICode([string] $path){
	$comObjWI = New-Object -ComObject WindowsInstaller.Installer
	$MSIDatabase = $comObjWI.GetType().InvokeMember("OpenDatabase","InvokeMethod",$Null,$comObjWI,@($Path,0))
	$Query = "SELECT Value FROM Property WHERE Property = 'ProductCode'"
	$View = $MSIDatabase.GetType().InvokeMember("OpenView","InvokeMethod",$null,$MSIDatabase,($Query))
	$View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
	$Record = $View.GetType().InvokeMember("Fetch","InvokeMethod",$null,$View,$null)
	$Value = $Record.GetType().InvokeMember("StringData","GetProperty",$null,$Record,1)
	#Write-Host -ForegroundColor red $Value
	return $Value
}


function GetCurrentLocation([string]$Application){
$DeploymentType = Get-CMDeploymentType -ApplicationName $Application
$Count =  $DeploymentType.Count
#If ($Count -lt 1){
    $xml = [xml]$DeploymentType.SDMpackageXML
#}Else{
 #   $xml = [xml]$DeploymentType[0].SDMpackageXML
#}
 #$xml.AppMgmtDigest.DeploymentType.Installed.Contents. |ft
$ContentLocation = $xml.AppMgmtDigest.DeploymentType.Installer.Contents.Content.Location
return $ContentLocation
}

###########################################################################################
#Get Latest Version of Chrome from 
$LATEST = GetLatestGoogleVersion
#Get Currently Deployed Version from SCCM
$ApplicationName = "Google Chrome Update Test"
$APP = Get-CMApplication -Name $ApplicationName 
$Deployed = $APP.SoftwareVersion

if(($DEBUG)){
		$SCCMCOLOR = "Blue"
		#Write-Host -BackgroundColor White -ForegroundColor Black "Script started $Start DEBUG enabled"
		#Ssleep 1
	}else{
		$SCCMCOLOR = "Yellow"
}
		
Write-Host -ForegroundColor $SCCMCOLOR "Currently deployed version of $ApplicationName"$APP.SoftwareVersion ($APP.DateLastModified |Get-Date -Format 'dd-mm-yyyy')

Write-Host -ForegroundColor DarkGreen "The latest version of Google Chrome is $LATEST"
if ($LATEST -gt $Deployed){
	Write-Host -ForegroundColor DarkGreen "A newer version of Chrome ($LATEST) is available downloading..."
	Download-Chrome
	}else {
	Write-Host -BackgroundColor Black -ForegroundColor White "$Deployed is the latest version, exiting..."
	exit
	}

#Download and Stage Chrome MSI's
$ContentRoot = $DP + $LATEST
Set-Location -Path $TempDirectory
if ((Test-Path -Path $ContentRoot)){
		Write-Host -BackgroundColor Black -ForegroundColor White "$ContentRoot exiting..."
		exit 
	}else{
		Write-Host -ForegroundColor DarkGreen "A $LATEST folder on $DP does not exist`nCreating $ContentRoot"
		New-Item -ItemType dir -Path $DP -Name $LATEST |Out-Null
		Write-Host -ForegroundColor DarkGreen "Creating $ContentRoot\x32"
		New-Item -ItemType dir -Path $ContentRoot -Name 'x32' |Out-Null
		Write-Host -ForegroundColor DarkGreen "Copying $TempDirectory\GoogleChromeStandaloneEnterprise.msi to $ContentRoot\x32\"
		Copy-Item -Path $TempDirectory\'GoogleChromeStandaloneEnterprise.msi' $ContentRoot\x32\
		Write-Host -ForegroundColor DarkGreen "Creating $ContentRoot\x64"
		New-Item -ItemType dir -Path $ContentRoot -Name 'x64' |Out-Null
		Write-Host -ForegroundColor DarkGreen "Copying $TempDirectory\GoogleChromeStandaloneEnterprise64.msi to $ContentRoot\x64\"
		Copy-Item -Path $TempDirectory\'GoogleChromeStandaloneEnterprise64.msi' $ContentRoot\x64\
	}
###########################################################################################
#break
#Make Changes in SCCM for 3 versions of chrome

#$Apps = @("Google Chrome Update Test")
$Apps = @("Google Chrome Update Test", "Google Chrome (x86)", "Google Chrome")
foreach ($ApplicationName in $Apps){
	Write-Host -ForegroundColor $SCCMCOLOR "Updateing $ApplicationName SCCM propertiesG"
	
	if($ApplicationName.EndsWith("(x86)")){
		Write-Host -ForegroundColor DarkGreen "$ApplicationName is 32 bit"
		$type = '\x32'
		$Installer= "\GoogleChromeStandaloneEnterprise.msi"
	}else{
		Write-Host -ForegroundColor DarkGreen "$ApplicationName is 64 bit"
		$type = '\x64'
		$Installer= "\GoogleChromeStandaloneEnterprise64.msi"
	}

	$ContentLocation = $ContentRoot + $type + $Installer
	
	
	$NOW = Get-Date
	#Google Chrome Update TEST
	Set-Location -Path $TempDirectory
	$MSI = GetMSICode $ContentLocation
	#$MSI |fl
	#rite-Host -ForegroundColor DarkGreen "MSI Code is $MSI"
	$MSI = $MSI |select # .ToString().TrimStart()
	Write-Host -ForegroundColor DarkGreen "MSI Code is $MSI"
	Set-Location -Path pur:
	#$MSI =( $MSI.ToString().trim()# |select 
	$DTYPE = Get-CMDeploymentType -ApplicationName $ApplicationName 
	$DTNAME = (Get-CMDeploymentType -ApplicationName $ApplicationName).LocalizedDisplayName
	#Write-Host -ForegroundColor $SCCMCOLOR "Updating $ApplicationName in SCCM"
	#Set New Content Location in Deployment Type, and update MSI Code Detection Method
	Write-Host -ForegroundColor $SCCMCOLOR "Setting content location to $ContentLocation`nUpdateing MSI detection code to $MSI"
	#Write-Host -ForegroundColor $SCCMCOLOR "Updateing MSI detection code to $MSI"
	#break
	$UPDATED = "Last updated $NOW  by $ScriptName"
	if(!($DEBUG)){Set-CMMSIDeploymentType  -ApplicationName $ApplicationName -DeploymentTypeName $DTNAME -ProductCode $MSI -ContentLocation $ContentLocation -Comment $UPDATED }
	#Set Version and update release date in Application properties
	Write-Host -ForegroundColor $SCCMCOLOR "Setting version to $LATEST`nUpdateing last updated date to" (Get-Date -Format 'MM:dd:yyyy::hh:mm:ss')
	if(!($DEBUG)){Set-CMApplication  -Name $ApplicationName -SoftwareVersion $LATEST -ReleaseDate (Get-Date) -LocalizedDescription $UPDATED  }
	#Update Distribution Points
	Write-Host -ForegroundColor $SCCMCOLOR "Updating Distribution Points"
	if(!($DEBUG)){Update-CMDistributionPoint -ApplicationName $ApplicationName -DeploymentTypeName $DTYPE.LocalizedDisplayName }
	Write-Host -ForegroundColor Cyan "Completed $ApplicationName $LATEST Update"
} 
Write-Host -BackgroundColor Black -ForegroundColor White "Script Ended" (Get-Date) "runtime::"(New-TimeSpan -Start $START -End (Get-Date)) 
