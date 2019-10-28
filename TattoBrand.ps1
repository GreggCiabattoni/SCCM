#############################################################################
# Gregg Ciabattoni
# 
# Augast 30th 2017
# TattoBrand.ps1
# This Script recieves vaules from a task sequence
# and Tattoos them to the Deployment and OEM Registry Keys
#
# It Then Copies Logos to the local computer and hack the 
# Defaul img0.jpg. This should be done later with unattended.xml
#
# The script then sets the defualt Lockscreen backroud
# Logos are stored in package. 
# TS Step passes values in this order 
# %Make %Model% %SerialNumber% %AssetTag% %UUID% %_SMSTSPackageName%
#############################################################################

param([string]$MAKE,[string]$MODEL,[string]$SERIAL,[string]$ASSETTAG,[string]$UUID,[string]$TSID)
#DEBUG echo "$TSID ran $NOW `r`n`tSERIAL NUMBER $SERIAL`r`n`tService Tag`t$ASSETTAG`r`n`tMake`t$MAKE`r`n`tModel`t$MODEL`r`n`tUUID`t$UUID" > C:\Logs\BrandingSrart.log
clear
$NOW = date 

function AddUpdateReg([string] $RPATH, [string] $NAME, [string] $VALUE){
	$KEY = $RPATH + "\" + $NAME
	Write-Host -ForegroundColor Gray "$KEY"
	if((Get-ItemProperty -Path  $RPATH -Name $NAME -ErrorAction SilentlyContinue )){
		Write-Host -ForegroundColor Green "$KEY Existis updating $NAME with $VALUE"
		Set-ItemProperty -Path $RPATH -Name $NAME -Value $VALUE #  `-PropertyType DWORD -Force | Out-Null
	}else{
		Write-Host -ForegroundColor Blue "$KEY Does not exist adding $KEY with $VALUE"
		New-ItemProperty -Path $RPATH -Name $NAME -Value $VALUE | Out-Null
	}


}

#Tattoo
$OEMREGPATH	="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation"

$REGVAL 	="C:\Windows\Web\4K\Wallpaper\Windows\pclogo.jp"
$REGNAME 	="Logo"
AddUpdateReg $OEMREGPATH $REGNAME $REGVALL

$REGVAL 	="YOUR ORGANIZATION NAME"
$REGNAME 	="Manufacturer"
AddUpdateReg $OEMREGPATH $REGNAME $REGVALL

$REGVAL 	="Mon/Fri 8AM to 8PM"
$REGNAME	="SupportHours"
AddUpdateReg $OEMREGPATH $REGNAME $REGVAL

$REGVAL		="www.du/wot"
$REGNAME	="SupportURL"
AddUpdateReg $OEMREGPATH $REGNAME $REGVAL

$REGVAL		="(914)251-6465"
$REGNAME	="SupportPhone"
AddUpdateReg $OEMREGPATH $REGNAME $REGVAL

$REGVAL		=$MODEL
$REGNAME 	="Model"
AddUpdateReg $OEMREGPATH $REGNAME $REGVAL

$REGVAL		=$NOW
$REGNAME	="Build Date"
AddUpdateReg $OEMREGPATH $REGNAME $REGVAL

$REGVAL		=$SERIAL
$REGNAME 	="SerialNumber"
AddUpdateReg $OEMREGPATH $REGNAME $REGVAL

$REGVAL		=$ASSETTAG
$REGNAME 	="Service Tag #"
AddUpdateReg $OEMREGPATH $REGNAME $REGVAL

$REGVAL		=$UUID
$REGNAME 	="UUID"
AddUpdateReg $OEMREGPATH $REGNAME $REGVAL

#Default MDTP REGISTRY PATH also done with zittatto, but doesn't add these vaules
#GNC This will replace ZTI tatto eventually
$OEMREGPATH	="HKLM:\SOFTWARE\Microsoft\Deployment 4"

$REGVAL		=$SERIAL
$REGNAME 	="SerialNumber"
AddUpdateReg $OEMREGPATH $REGNAME $REGVAL

$REGVAL		=$ASSETTAG
$REGNAME 	="Service Tag #"
AddUpdateReg $OEMREGPATH $REGNAME $REGVAL

$REGVAL		=$UUID
$REGNAME 	="UUID"
AddUpdateReg $OEMREGPATH $REGNAME $REGVAL

$REGVAL		=$TSID
$REGNAME 	="Task Sequence Name"
AddUpdateReg $OEMREGPATH $REGNAME $REGVAL

$REGVAL		=$NOW
$REGNAME	="Build Date"
AddUpdateReg $OEMREGPATH $REGNAME $REGVAL


#Branding
#Set Default Wallpapper HACK
#Copy Logs to Local Hard Drive
#clear
takeown /f C:\Windows\Web\Wallpaper\Windows\img0.jpg 
takeown /f C:\Windows\Web\4K\Wallpaper\Windows\*.*
icacls C:\Windows\Web\Wallpaper\Windows\img0.jpg  /Grant 'System:(F)'
icacls C:\Windows\Web\4K\Wallpaper\Windows\*.* /Grant 'System:(F)'
Remove-Item C:\Windows\Web\Wallpaper\Windows\img0.jpg -ErrorAction SilentlyContinue
Remove-Item C:\Windows\Web\4K\Wallpaper\Windows\*.* -ErrorAction SilentlyContinue
#Needed For testing Set-Location C:\ScriptRoot
Copy-Item $PSScriptRoot\img0.jpg C:\Windows\Web\Wallpaper\Windows\img0.jpg -ErrorAction SilentlyContinue
Copy-Item $PSScriptRoot\images\*.* C:\Windows\Web\4K\Wallpaper\Windows\ -ErrorAction SilentlyContinue

#Set LockScreen Backround to  Logo
$REGPATH 	="HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization"
$REGROOT 	="HKLM:\SOFTWARE\Policies\Microsoft\Windows"

#Check that personlization key exists 
if(!(Get-ItemProperty -Path  $REGPATH  -ErrorAction SilentlyContinue )){
		Write-Host -ForegroundColor Red "$REGPATH Does Not exist creating...."
		New-Item -Path $REGROOT -Name "Personalization" -Force #| Out-Null
	}else{
		Write-Host -ForegroundColor Blue "$REGPATH exits"
	}

$REGNAME	="LockScreenImage"
$REGVAL		="C:\Windows\Web\4K\Wallpaper\Windows\LockScreen.jpg"
AddUpdateReg $REGPATH $REGNAME $REGVAL

#Leave Log On Local Machine
echo "$TSID ran $NOW `r`n`tSERIAL NUMBER $SERIAL`r`n`tService Tag`t$ASSETTAG`r`n`tMake`t$MAKE`r`n`tModel`t$MODEL`r`n`tUUID`t$UUID" > C:\Logs\Branding.log
