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
############################################################
$ColPath = "TES:\DeviceCollection\Staff"
$Collection = "My Collection"
#Devices can come from AD OU query or any database table
$DEVICES = ("WORKSTATION1","WORKSTATION2")
#Create Collection
$LimitingCollection = "ALL_Staff_Workstations"
$LimitingCollectionId = (Get-CMDeviceCollection -Name $LimitingCollection).CollectionID
#######New-CMDeviceCollection -Name $Collection -LimitingCollectionId $LimitingCollectionId 

$CollectionId = (Get-CMDeviceCollection -Name $Collection).CollectionID

#######Move-CMObject  -ObjectId $CollectionId -FolderPath  $ColPath

foreach ($DEVICE in $DEVICES){
	Write-Host -ForegroundColor Blue "Adding $DEVICE to $Collection"
	Add-CMDeviceCollectionDirectMembershipRule -CollectionID "$CollectionID" -ResourceID (Get-CMDevice -Name $DEVICE).ResourceID 
}

#Get-Content "C:\TEMP\CollectionMembers.txt" | foreach { Add-CMDeviceCollectionDirectMembershipRule -CollectionID "CollectionID" -ResourceID (Get-CMDevice -Name $_).ResourceID }
