#region Connect to CM Console

# Site configuration
$SiteCode = "GRE" # Site code 
$ProviderMachineName = "MGSCCM01.intra.grenergy.com" # SMS Provider machine name

# Customizations
$initParams = @{}
#$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
#$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

# Do not change anything below this line

# Import the ConfigurationManager.psd1 module 
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}

# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams


#endregion

#region Global Variables

CLS
$SUG = 'Server Quarterly Patches 2021-07-09 14:35:33'
#$SUG = Read-Host -Prompt 'What is the Software Update Group Name?'
$ActivityID = Read-Host -Prompt 'What is the Activity ID?'
$Mitigated = Read-Host -Prompt 'Are these patches Mitigated?'
$WhoAmI = $env:UserName
$OUTFile = 'C:\Users\' + $WhoAmI + '\desktop\ListPatches.csv'

#endregion

#region Get Software Updates

$ListSoftwareUpdates = (Get-CMSoftwareUpdate -UpdateGroupName $SUG -Fast) | Sort LocalizedDisplayName
ForEach ($i in $ListSoftwareUpdates)
{

$GREPatchDescription = $i.LocalizedDisplayName -replace ",", ' - '
$GRERevisedDate = $i.DateRevised
$GRERevisedDate1 = $GRERevisedDate.GetDateTimeFormats()[86]
$GREPatchID = $i.ArticleID

$msg = $ActivityID + "," + $GREPatchID + "," + $GREPatchDescription + "," + $Mitigated + "," + "," + "," + "," + $GRERevisedDate1
$msg | Out-File $OUTFile -Append
}

#endregion

#region Create Header for CSV File

$NewFile = 'C:\Users\' + $WhoAmI + '\Desktop\Patches.csv'
Add-Content $NewFile “Activity_ID,Patch_Number,Patch_Description,Mitigated,Does Not Apply,Comp_Measures,Notes,Release_Date”
Get-Content $OutFile | Add-Content $NewFile
Remove-Item $OutFile -Force:$true -Confirm:$false

#endregion
