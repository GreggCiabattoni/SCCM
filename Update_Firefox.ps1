#Gregg Ciabattoni
#3/15/2018
#Update Mozila Firefox ESR to latest version
#Download new msi's to stageing location
#Get File Version from file metadata 
#Create Folder with version name on Distribution Point
#Update Deployment Type Content Location
#Update Deployment Type Detection Method's MSI PRoduct Code
#Update Version In Application
#Update Date Published In Application
#Redistibute Content
#Update distibution point
Clear
$SiteCode = "TES" # Site code 
$ProviderMachineName = "sccmserver.example.com" # SMS Provider machine name
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
###
$EMAILBODY 	= ""
$START		= Get-Date
$ScriptName = "UpdateFireFox_ESR.ps1"
$ApplicationName= "Mozilla Firefox Update Test" 
$DP = "\\dp\share\Mozilla_Firefox"
$TempDirectory = "C:\Temp"
###
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

Function Download-FireFoxESR([string]$version) {
#https://download.mozilla.org/?product=firefox-esr-latest-ssl&os=win64&lang=en-US
#https://download.mozilla.org/?product=firefox-esr-latest-ssl&os=win&lang=en-US
    # Test internet connection
	if(Test-Path $TempDirectory\$version){
				Write-Host -ForegroundColor DarkGreen "Removing $TempDirectory\$version"
				Remove-Item -Path $TempDirectory\$version -Recurse -Force 
	}else{
				Write-Host -ForegroundColor Blue "$TempDirectory\$file does not exist"
	}
		
    if (Test-Connection mozilla.org -Count 3 -Quiet) {
		 foreach($type in ('x32', 'x64')){
			switch($type){
				'x64'{
				$lnk ="https://download.mozilla.org/?product=firefox-esr-latest-ssl&os=win64&lang=en-US"
				$file ="Firefox_Setup_"+ $version + ".exe"
				}
				'x32'{
				$lnk ="https://download.mozilla.org/?product=firefox-esr-latest-ssl&os=win&lang=en-US"
				$file ="Firefox_Setup_" + $version + "_x32.exe"
				}
			}
			
				if(Test-Path $TempDirectory\$file){
					Write-Host -ForegroundColor DarkGreen "Removing $TempDirectory\$file"
					Remove-Item -Path $TempDirectory\$file -Force 
				}else{
					Write-Host -ForegroundColor Green "$TempDirectory\$file does not exist"
				}
				
				Write-Host -ForegroundColor Magenta "Downloading $type $file ... $lnk " -NoNewLine
		       
				# Download the installer from Mozilla
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
				
				
				
				Write-Host -ForegroundColor DarkGreen "Extracting $file to $TempDirectory\$version\$type"
				Invoke-Expression -Command "C:\'Program Files'\7-Zip\7z.exe X $TempDirectory\$file -o$TempDirectory\$version\$type -y" |Out-Null
		}
	} else {
		        Write-Host "failed. Unable to connect to Mozilla servers." -ForegroundColor Red
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
	return $Value
}


Function GetDetectionClauses([string]$AppName){
$DeploymentType = (Get-CMDeploymentType -ApplicationName $AppName).LocalizedDisplayName
$SDMPackageXML = (Get-CMDeploymentType -ApplicationName "$($ApplicationName)" -DeploymentTypeName "$($DTNAME)").SDMPackageXML
[string[]]$OldDetections = (([regex]'(?<=SettingLogicalName=.)([^"]|\\")*').Matches($SDMPackageXML)).Value
return $OldDetections
}
function SendEmail([string]$TO, [string]$SUBJECT, [string]$MSGTXT){
	Set-Location C:\Scripts\Servers
	$HTML 			= ""
	$FROM 			= "$ScriptName@example.com"
	#$TO				= "gregg.ciabattoni@gmail.com"
	$HTMLHEADER		= Get-Content Header.html
	$HTMLFOOTER		= Get-Content Footer.html
	$SMTPSERVER		= "mail.example.com"
	$MAIL			= new-object Net.Mail.SmtpClient($SMTPSERVER)
	$HTML 			= $HTMLHEADER + $MSGTXT + $HTMLFOOTER
	#$HTML 			= $MSGTXT
	#$HTML 			=  $MSGTXT + $HTMLFOOTER
	$MSG 			= new-object Net.Mail.MailMessage($FROM,$TO,$SUBJECT,$HTML)
	$MSG.IsBodyHTML = $true
	$MAIL.Send($MSG)
	#Write-Host -ForegroundColor Cyan "HTML HEADER is `n$HTMLHEADER`nHTML Footer  is `n$HTMLFOOTER`nHTML is `n$HTML"
}

#End Functions
##########################################################################
#Get Latest version from mozilla
#$LATESTVERSION = (Invoke-WebRequest  "https://product-details.mozilla.org/1.0/firefox_versions.json" | ConvertFrom-Json).psobject.properties.value[-1]
$LATESTVERSION = (Invoke-WebRequest  "https://product-details.mozilla.org/1.0/firefox_versions.json" | ConvertFrom-Json).FIREFOX_ESR 
Write-Host -ForegroundColor DarkGreen "The latest version of Firefox ESR is $LATESTVERSION"

#Get Currently Deployed Version from application
$APP = Get-CMApplication -Name $ApplicationName 
$Deployed = $APP.SoftwareVersion
$DTNAME = (Get-CMDeploymentType -ApplicationName $ApplicationName).LocalizedDisplayName
Write-Host -ForegroundColor Yellow "The currently deployed version of Firefox ESR is $Deployed"
#Write-Host -ForegroundColor Yellow "$ApplicationName "$APP.SoftwareVersion $APP.DateLastModified "`nDeployment ID Name is $DTNAME"


#Download Firefox 

if(($LATESTVERSION -replace ".{3}$") -gt ($Deployed -replace ".{3}$")){
	Write-Host -ForegroundColor DarkGreen "A newer version of $ApplicationName is available Downloading $$LATESTVERSION ..."
	Download-FireFoxESR $LATESTVERSION
}else{ 
	Write-Host -ForegroundColor Red "$Deployed is the latest version exising..."
	exit
}

#break

#Stage Firefox
Set-Location $TempDirectory
$PUR_PREF = $DP + "\Settings\MySettings-settings.js"
$PUR_CFG = $DP + "\Settings\MySettings.cfg"
$UNINSTALL = $DP + "\Settings\UninstallOldFirefoxVersions.bat"
if(!(Test-Path $DP\$LATESTVERSION)){
	Write-Host -ForegroundColor DarkGreen "Copying $TempDirectory\$LATESTVERSION to $DP\$LATESTVERSION"
	Copy-Item  -Path  $TempDirectory\$LATESTVERSION $DP -Recurse -Force
	Write-Host -ForegroundColor DarkGreen "Copying config and prefrences files to distribution point..."
	Copy-Item -Path $PUR_PREF $DP\$LATESTVERSION\x64\core\defaults\pref -Force
	Copy-Item -Path $PUR_PREF $DP\$LATESTVERSION\x32\core\defaults\pref -Force
	Copy-Item -Path $PUR_CFG $DP\$LATESTVERSION\x64\core -Force
	Copy-Item -Path $PUR_CFG $DP\$LATESTVERSION\x32\core -Force
	Copy-Item -Path $UNINSTALL $DP\$LATESTVERSION\x64 -Force
	Copy-Item -Path $UNINSTALL $DP\$LATESTVERSION\x32 -Force
}else{
	Write-Host -ForegroundColor DarkRed "$DP\$LATESTVERSION already exists"
	#exit
}
#break
$Apps = @("Mozilla Firefox Update Test","Mozilla Firefox 60 x86" ,"Mozilla Firefox 60")
#$Apps = @("Mozilla Firefox Update Test" , "Mozilla Firefox 60 x86")
#clear
#$Apps = @("Mozilla Firefox Update Test")
foreach ($ApplicationName in $Apps){
	Write-Host -ForegroundColor yellow "Updateing $ApplicationName SCCM Properties"
	if($ApplicationName.EndsWith("x86")){
		Write-Host -ForegroundColor DarkGreen "$ApplicationName is 32bit"
		$type = '\x32'
	}else{
		Write-Host -ForegroundColor DarkGreen "$ApplicationName is 64bit"
		$type = '\x64'
	}
	
	Set-Location -Path pur:
	$NOW = Get-Date
	$DTYPE= Get-CMDeploymentType -ApplicationName $ApplicationName 
	$DTNAME = (Get-CMDeploymentType -ApplicationName $ApplicationName).LocalizedDisplayName
	$ContentLocation = $DP + "\" + $LATESTVERSION + $type
	$UPDATED = "Last updated $NOW  by $ScriptName"
	
	#Get Metadata needed for file ve rsion Check
	$EXE = $TempDirectory + "\" + $LATESTVERSION + $type + "\core\firefox.exe"
	$MD = Get-FileMetaData $EXE
	$FileVersion = $MD.'File version'
	Write-Host -ForegroundColor DarkGreen "$EXE File version is $FileVersion"
	
	Write-Host -ForegroundColor Yellow "Updateing Content location $ApplicationName`n$ContentLocation"
	Set-CMscriptDeploymentType  -ApplicationName $ApplicationName -DeploymentTypeName $DTNAME -ContentLocation $ContentLocation 
	##Get Existing Detection Rules from Application Deployment Type
	#clear
	$DC = GetDetectionClauses $ApplicationName
	##Create new Detection Rule for new file version
	$DM = New-CMDetectionClauseFile   -FileName 'firefox.exe' -Path 'C:\Program Files\Mozilla Firefox' -PropertyType Version -Value -ExpressionOperator IsEquals -ExpectedValue $FileVersion 
	Write-Host -ForegroundColor Yellow "Updating Detection rules for $ApplicationName"
	##Remove Old Detection Rule(s) and new rule
	Set-CMscriptDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName $DTNAME -RemoveDetectionClause ($DC) -AddDetectionClause $DM -Comment $UPDATED 
	##Set Version and update release date in Application properties
	Write-Host -ForegroundColor Yellow "Set Version to $LATESTVERSION and updateing release date to $NOW in $ApplicationName properties"
	Set-CMApplication  -Name $ApplicationName -SoftwareVersion $LATESTVERSION -ReleaseDate $NOW -LocalizedDescription $UPDATED 
	
	##Update Distribution Points
	Write-Host -ForegroundColor Yellow "Updating Distribution Points"
	Update-CMDistributionPoint -ApplicationName $ApplicationName -DeploymentTypeName $DTYPE.LocalizedDisplayName 
	Write-Host -ForegroundColor Magenta "Done updating $ApplicationName"



}

Write-Host -ForegroundColor DarkGray "Sending Email....."
$SBJ = "$ApplicationName was updated to $FileVersion"
$EMAILBODY 	+= "<FONT SIZE=`"3`" COLOR=`"blue`">$ApplicationName was updated to $FileVersion<BR>Updated by $ScriptName<BR></FONT>"
SendEmail 'gregg.ciabattoni@gmail.com'  $SBJ $EMAILBODY	
Write-Host -ForegroundColor DarkGray "Script Ended" (Get-Date) "runtime::"(New-TimeSpan -Start $START -End (Get-Date)) 




