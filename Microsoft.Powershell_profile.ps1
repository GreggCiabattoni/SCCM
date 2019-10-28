#Powershell Profile Example
Clear

if (Test-Path "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"){
	Write-Host -ForegroundColor Green "SCCM Console installed loading module"
	$SiteCode = "TES" # Site code 
	$ProviderMachineName = "sccmserver.example.com" # SMS Provider machine name
	$initParams = @{}
	
	if((Get-Module ConfigurationManager) -eq $null) {
	    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams  -ErrorAction SilentlyContinue
	}

	# Connect to the site's drive if it is not already present
	if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
	    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
	}
	# Set the current location to be the site code.
	Set-Location "$($SiteCode):\" @initParams
}



$Shell = $Host.UI.RawUI
$size = $Shell.WindowSize
$size.width=80
$size.height=25
#$Shell.WindowSize = $size
$size = $Shell.BufferSize
$size.width=70
$size.height=500
#$Shell.BufferSize = $size
$shell.BackgroundColor = “Black”
$shell.ForegroundColor = “Cyan”
#$psISE.Options.FontSize = 16
function LocalS { Set-Location "C:\Scripts" }
function HowTo {Set-Location "\\SERVER\SHARE\NAME\My Documents\PowerShell" }
Set-Alias ssh "C:\Program Files (x86)\PuTTY\putty.exe"
function wget ($url) {(new-object Net.WebClient).DownloadString("$url")}
#Set-Alias wget "Invoke-WebRequest"
function which ($cmd) { get-command $cmd | select path }
#function prompt { $Admin + "PS $(get-location)> " }
function global:prompt {"PS $(Get-Location) [$env:COMPUTERNAME]>"}
function global:getwinrm {Get-Service -Name winrm | format-list }
function global:gets{param($status) Get-Service | where {$_.status -eq $status} } 
function global:inverntory {param($name = ".") Get-WmiObject -Class Win32_OperatingSystem -Namespace root\cimv2 -ComputerName $name |Format-List * }
New-Alias gr getwinrm -Scope global -ErrorAction SilentlyContinue
New-Alias gs gets -Scope global -ErrorAction SilentlyContinue
New-Alias inv inverntory -Scope global -ErrorAction SilentlyContinue
# Get Group Members
function global:GetPURGroupMembers([string] $GroupName){
	if((Get-ADGroupMember  $GroupName).count -gt 0){
		#Write-Host (Get-ADGroupMember  $GroupName).count
		$GroupMembers=(Get-ADGroupMember   $GroupName -ErrorAction SilentlyContinue)
		foreach($GroupMember in $GroupMembers){	
			$GroupMemberName=$GroupMemberName+(Get-ADUser -Identity $GroupMember).Name+";"#"`n"
		}
		$CNT = (Get-ADGroupMember  $GroupName).count
		return "$GroupName contains $CNT Members`n$GroupMemberName"
		}
	else {
		Write-Host "$GroupName has no members"
		}
}



#$Global:Admin=''
#$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
#$principal = new-object System.Security.principal.windowsprincipal($CurrentUser)
#if ($principal.IsInRole("Administrators")) { $Admin='ADMIN ' }
#$host.ui.rawui.WindowTitle = "$TITLE PS $pwd"
#$ (Get-Host).UI.RawUI
#net use B: /d
#net use B: "\\SERVER\SHARE\USERNAME\My Documents\How-To\PowerShell\"
if(!(Test-Path "B:\")){
	New-PSDrive -PSProvider FileSystem -Root "\\SERVER\SHARE\USERNAME\My Documents\How-To\PowerShell\" -Name B
}
#clear
#Set-Location C:
#Remove-PSDrive -Name B 
Set-Location B:

Write-Host -ForegroundColor Blue -BackgroundColor Yellow "Loaded GC functions: LocalS, HowTo, Wget, Which, GetWin, Gets, inverntory, GetPURGroup"

# Chocolatey profile
$ChocolateyProfile = "$env:ChocolateyInstall\helpers\chocolateyProfile.psm1"
if (Test-Path($ChocolateyProfile)) {
  Import-Module "$ChocolateyProfile"
}
