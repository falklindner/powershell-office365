. $PSScriptRoot/variables.ps1
. $PSScriptRoot/localfunctions.ps1

Class Person
{
 [String]$Vorname
 [String]$Nachname
 [String]$Mail
 [String]$AG
 [String]$IndividualDG
}

Function WriteToLog ($Text) {
    Out-File -FilePath $LOG -Append -Encoding utf8 -InputObject "$Text"
}
Function WriteToCMD ($Text) {
    Out-File -FilePath $TestCMDs -Append -Encoding utf8 -InputObject "$Text"
}

################################################ Main routine

##

#### Main loop


### External XLSX loading
$answer = Read-Host -Prompt "Load Excel files from Server? (y/n)" 
If ($answer -contains "y"){
    If (Test-Path $ExcelPath) {Remove-Item $ExcelPath -Force -Recurse}
    New-Item -ItemType Directory -Path $ExcelPath -Force

    New-Item -ItemType File -Path $LOG -Force
    Write-Host "###########################################################################################"
    Write-Host "###########################################################  Generating Local Address List "
    Write-Host "###########################################################################################"  

    Connect-PnPOnline -Url $SPOnline -Credentials ($UserCredential)

    ForEach ($file in  $(Get-PnPFolderItem -FolderSiteRelativeUrl "Dokumentbibliothek" -ItemType File)) {
        Get-PnPFile -Url /Dokumentbibliothek/$($file.Name) -AsFile -Force -Path $ExcelPath -Filename $file.Name
    }
}



### LAL Generating
If (Test-Path $ExcelPath\lal.csv) {
    $redo = Read-Host -Prompt "LAL CSV found, redo LAL processing? (y/n)" 
    
    If ($redo -contains "y") {        

        $ExcelList = Get-ChildItem -Path $ExcelPath -File  | Where-Object {($_.Extension -eq ".xlsx") -and ($_.Name -notlike "*~*")}
        Write-Host "Found $($ExcelList.Count) Excel Files, importing them takes ca. 80s"
        $GlobalIDGList = @()
        $GlobalIDGList = Get-GlobalIDGs -ExcelList $ExcelList 
        Write-Host "Global IDGs done"
        $LocalAddressList = Get-LAL -ExcelList $ExcelList -GlobalIDGList $GlobalIDGList
        $LocalAddressList | Select-Object -Property Vorname,Nachname,Mail,AG | Export-Csv -Path $ExcelPath\lal.csv -Encoding Unicode
        Write-Host "Export to CSV is done. Exported lines:" $LocalAddressList.Count        
        }

}
Else {
        $ExcelList = Get-ChildItem -Path $ExcelPath -File  | Where-Object {($_.Extension -eq ".xlsx") -and ($_.Name -notlike "*~*")}
        Write-Host "Found $($ExcelList.Count) Excel Files, importing them takes ca. 80s"
        $GlobalIDGList = @()
        $GlobalIDGList = Get-GlobalIDGs -ExcelList $ExcelList 
        Write-Host "Global IDGs done"
        $LocalAddressList = Get-LAL -ExcelList $ExcelList -GlobalIDGList $GlobalIDGList
        $LocalAddressList | Select-Object -Property Vorname,Nachname,Mail,AG | Export-Csv -Path $ExcelPath\lal.csv -Encoding Unicode
        Write-Host "Export to CSV is done. Exported lines:" $LocalAddressList.Count 

}




$LocalAddressList = Import-Csv -Path $ExcelPath\lal.csv -Encoding Unicode

$string = Read-Host -Prompt "String to look for" 

$list = $LocalAddressList | Where-Object { $_.Nachname -match $string -or $_.Vorname -match $string -or $_.Mail -match $string} 

ForEach ($item in $list)
{
    Write-Host "############################################################"
    Write-Host "Name: "$item.Nachname","$item.Vorname
    Write-Host "Mail: "$item.Mail
    foreach ($ag in $item.AG.Split(",")) {Write-Host "AG  : "$ag}
    Write-Host "############################################################"
}