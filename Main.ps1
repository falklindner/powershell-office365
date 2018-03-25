. $PSScriptRoot/variables.ps1
. $PSScriptRoot/remotefunctions.ps1
. $PSScriptRoot/localfunctions.ps1
. $PSScriptRoot/exchangeinterface.ps1

Class Person
{
 [String]$Vorname
 [String]$Nachname
 [String]$Mail
 [String]$AG
 [String]$IndividualDG
}


$ExcelPath = "$(Get-Item $PSScriptRoot)\work"
$LOG = "$ExcelPath\Log\Log.txt"
$TestCMDs = "$ExcelPath\Log\CMDs.txt"

Function WriteToLog ($Text) {
    Out-File -FilePath $LOG -Append -Encoding utf8 -InputObject "$Text"
}
Function WriteToCMD ($Text) {
    Out-File -FilePath $TestCMDs -Append -Encoding utf8 -InputObject "$Text"
}

# Main routine
$startDTM = (Get-Date)

LoginFHH_Exchange($UserCredential)
Connect-PnPOnline -Url $SPOnline -Credentials ($UserCredential)

If (Test-Path $ExcelPath) {Remove-Item $ExcelPath -Force -Recurse}
New-Item -ItemType Directory -Path $ExcelPath -Force

New-Item -ItemType File -Path $LOG -Force
WriteToLog -Text "###########################################################################################"
WriteToLog -Text "###########################################################  Generating Local Address List "
WriteToLog -Text "###########################################################################################"


ForEach ($file in  $(Get-PnPFolderItem -FolderSiteRelativeUrl "Dokumentbibliothek" -ItemType File)) {
    Get-PnPFile -Url /Dokumentbibliothek/$($file.Name) -AsFile -Force -Path $ExcelPath -Filename $file.Name
}

$ExcelList = Get-ChildItem -Path $ExcelPath -File  | Where-Object {($_.Extension -eq ".xlsx") -and ($_.Name -notlike "*~*")}
WriteToLog -Text "Found $($ExcelList.Count) Excel Files"
$LocalAddressList = BuildLAL -ExcelList $ExcelList 
$LocalAddressList | Out-File -FilePath $ExcelPath\log\lal.txt 
WriteToLog -Text "###########################################################################################"
WriteToLog -Text "###########################################################  Obtaining Global Address List "
WriteToLog -Text "###########################################################################################"


$GlobalAddressList = BuildGAL

$GlobalAddressList | Out-File -FilePath $ExcelPath\log\gal.txt 

$Comparison = Compare-Object -ReferenceObject $GlobalAddressList -DifferenceObject $LocalAddressList -IncludeEqual -Property Mail -PassThru
$Comparison | Select-Object Mail,SideIndicator | Where-Object { $_.SideIndicator -ne "==" } |  Out-File -FilePath $ExcelPath\log\comp.txt

WriteToLog -Text "###########################################################################################"
WriteToLog -Text "############################################################  Mangaging Global Address List"
WriteToLog -Text "###########################################################################################"

RemoveFromGAL ($Comparison)
AddToGAL ($Comparison)


WriteToLog -Text "###########################################################################################"
WriteToLog -Text "############################################################  Mangaging Distribution Groups"
WriteToLog -Text "###########################################################################################"

ManageDistribtionGroups -LAL $LocalAddressList -ExcelList $ExcelList

CloseFHH
Disconnect-PnPOnline
$endDTM = (Get-Date)

WriteToLog -Text "Elapsed Time: $(($endDTM-$startDTM).totalseconds) seconds"