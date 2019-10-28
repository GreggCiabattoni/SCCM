#Gregg Ciabattoni
#Downloads MS Patches that have came out in the last n days
#Adds updates to existing Software Update Groups (Deployments made twice a year)
#Removes expired updates from all Software Update Groups
#Moves expired updates to expired folders
#Moves deployed updates to the current years folder
Clear
[BOOL]$DEBUG=$false
$SiteCode = "TES" # Site code 
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
#Clear

function DownloadUpdates(){	
	# Get-CMSoftwareUpdateSyncStatus
	$YEAR = Get-Date -Format yyyy
	$YEAR_FOLDER = "PUR:\SoftwareUpdate\$YEAR"
	if (($LASTSYNC = Get-CMSoftwareUpdateSyncStatus).LastSyncErrorCode -eq 0){
		Write-Host -ForegroundColor Green "Last Sync completed successfully on" $LASTSYNC.LastSuccessfulSyncTime
	}else{
		Write-Host -ForegroundColor Red "Last Sync failed last successfully on" $LASTSYNC.LastSuccessfulSyncTime
		#break
	}

	#$UPDATES = Get-CMSoftwareUpdate -fast
	#$UPDATES.Count
	$NEWUPDATES = Get-CMSoftwareUpdate -fast | Where{ (($_.DateLastModified -gt ((Get-Date).AddDays(-$DAYS))) -and ($_.IsExpired -eq $false) -and ($_.IsSuperseded  -eq $false) -and ($_.IsDeployed -eq $false) ) } | Sort-Object -Property DatePosted  
	foreach ($NEWUPDATE in $NEWUPDATES){
		$CNT ++
		if(!($NEWUPDATE.IsDeployable)){
			$DOWNLOADED ++
			Write-Host -ForegroundColor Magenta "$DOWNLOADED Downloading " $NEWUPDATE.LocalizedDisplayName "to $DeploymentPackageName ........."
			Save-CMSoftwareUpdate -DeploymentPackageName $DeploymentPackageName -SoftwareUpdateId $NEWUPDATE.CI_ID
		}
		if(!($NEWUPDATE.IsDeployed)){
			Write-Host -ForegroundColor Green "$CNT Adding" $NEWUPDATE.LocalizedDisplayName "to $SUGROUP"
			if(!($DEBUG)){$NEWUPDATE | Add-CMSoftwareUpdateToGroup -Confirm:$false  -SoftwareUpdateGroupName $SUGROUP }
			Write-Host -ForegroundColor Green "Moving" $NEWUPDATE.LocalizedDisplayName "to $YEAR_FOLDER"
			if(!($DEBUG)){$NEWUPDATE | Move-CMObject -Confirm:$false -FolderPath $YEAR_FOLDER }
		}
		#break
	}
	
	if($CNT -gt 0){
		$UPDATED = "Last Updated " + (Get-Date) + " via Munthly_Patch_Deployment.ps1" 
		Set-CMSoftwareUpdateGroup -Name $SUGROUP -Description $UPDATED
		Set-CMSoftwareUpdateDeploymentPackage -Name $DeploymentPackageName -Description $UPDATED
	}else{
		Write-Host -ForegroundColor Green "No new updates avialable"
	}
}

function RemoveExpiredUpdates(){
	$EXPTOTAL=0; $GRPCNT=0; $YEAR = Get-Date -Format yyyy
	$EXPIRED_FOLDER = "PUR:\SoftwareUpdate\$YEAR\Expired"
	Get-CMSoftwareUpdateGroup | foreach {
		$GRPCNT++;$LEXPCNT=0
		Write-Host -ForegroundColor DarkGreen "$GRPCNT Checking" $_.LocalizedDisplayName "for expired updates:" -NoNewline
		$GROUPNAME = $_.LocalizedDisplayName
		Get-CMSoftwareUpdate -fast  -UpdateGroupName $GROUPNAME | Where { (($_.IsExpired -eq 'True') -or ($_.IsSuperseded -eq 'True') ) }| foreach{
			$EXPTOTAL ++
			$LEXPCNT ++
			Write-Host -ForegroundColor Yellow "`n$EXPTOTAL Removing expired item" $_.LocalizedDisplayName $_.ObjectPath $_.IsExpired "FROM $GROUPNAME"
			$_ | Remove-CMSoftwareUpdateFromGroup -Confirm:$false -SoftwareUpdateGroupName $GROUPNAME -Force 
			
			Write-Host -ForegroundColor Yellow "Moving" $_.LocalizedDisplayName "to $EXPIRED_FOLDER"
			$_ | Move-CMObject -Confirm:$false -FolderPath $EXPIRED_FOLDER 
		}
		
		if($LEXPCNT -eq 0){
			Write-Host -ForegroundColor Green "{OK}"
		}else{
			Write-Host -ForegroundColor DarkRed  $_.LocalizedDisplayName "had $LEXPCNT expired updates"
		}
	}

	$LEXPCNT=0
	Write-Host -ForegroundColor DarkGreen "Checking for non deployed expired updates in (root) / folder:" -NoNewline
	Get-CMSoftwareUpdate -fast | Where { ((($_.ObjectPath -eq "/") -and ($_.IsExpired -eq 'True')) -or (($_.ObjectPath -eq "/") -and ($_.IsSuperseded -eq 'True')) )} | Sort-Object -Property DatePosted |foreach{
		$LEXPCNT++
		$EXPTOTAL++
		#Write-Host -ForegroundColor Magenta $_.LocalizedDisplayName $_.ObjectPath
		Write-Host -ForegroundColor Yellow "Moving" $_.LocalizedDisplayName "to $EXPIRED_FOLDER"
		$_ | Move-CMObject  -FolderPath  $EXPIRED_FOLDER -Confirm:$false
	}
	
	if($LEXPCNT -eq 0){
			Write-Host -ForegroundColor Green "{OK}"
		}else{
			Write-Host -ForegroundColor DarkRed $_.LocalizedDisplayName "$LEXPCNT expired updates"
		}
	
	if($EXPTOTAL -eq 0){
			Write-Host -ForegroundColor Green "All $GRPCNT Software Update Groups already up to date"
		}else{
			Write-Host -ForegroundColor DarkGreen "$EXPTOTAL expired updates from $GRPCNT groups were undeployed and moved to expired folder"
		}
	
}

#End Functions 
############################################################
$X=0;$CNT=0;$DOWNLOADED=0;$DAYS=28
$SUGROUP = "Microsoft Software Updates 2019 2/2  (Jul-Dec)"
#$SUG= Get-CMSoftwareUpdateGroup -Name $SUGROUP
#$SUG.Updates
$DeploymentPackageName = "ALL_Microsoft_Updates_2019-SS27"
#Get-CMUpdateGroupDeployment -Name $SUGROUP


DownloadUpdates
RemoveExpiredUpdates



Write-Host -ForegroundColor Blue "Script Ended "#$CNT new updates exist. $DOWNLOADED were downloaded"
#sleep 10
#$NEWUPDATES | Select-Object -Property @{N='Date Posted';E={$_.DatePosted}},@{N='Name';E={$_.LocalizedDisplayName}} 
