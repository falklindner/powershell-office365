Function BuildLAL ($ExcelList) {
    $LocalAddressList = New-Object System.Collections.Generic.List[Person] 
    # Importing the Contacts from each Excel File in the Base path into a list of instances of type person
    $i=0
    foreach ($ExcelFile in $ExcelList) {
      ExcelToLAL -ExcelFile $ExcelFile -AddressList $LocalAddressList
      $InternalContacts = $LocalAddressList | Where-Object { $_.Mail -like "*fhh-portal*" }
      Write-Progress -activity "Importing Local Address List" -status "Imported $i of $($Excellist.Count)" -percentComplete (($i / $Excellist.Count)  * 100)
      WriteToLog -Text "###########  $($ExcelFile.Name) done. $($LocalAddressList.Count) Persons identified ($($InternalContacts.Count) FHH Members)."
      }     
    Return $LocalAddressList
}
Function ExcelToLAL ($ExcelFile, $AddressList){
    
    $Header = @{Label="Nachname"; Expression = {$_.Nachname}},`
              @{Label="Vorname" ; Expression = {$_.Vorname}},`
              @{Label="Mail"    ; Expression = {$_."E-Mail-Adresse"}},`
              @{Label="AGName"  ; Expression = {$ExcelFile.BaseName}}

    $Header = AddIDGToHeader -Header $Header -ExcelFile $ExcelFile


    #Importing contact details from Excel File and writing them into global list $AddressList
    $ExcelImport = Import-Excel -Path $ExcelFile.FullName  | `
    Select-Object -Property $Header | `
    Where-Object {($_.Nachname -or $_.Vorname) -and $_.Mail}
    
    
    [Int]$skipped = 0
    WriteToLog -Text "###########################################################################################"
    WriteToLog -Text "############################################################  $($ExcelFile.Name)"
    WriteToLog -Text "###########################################################################################"
    
    foreach ($ExcelLine in $ExcelImport) {
        $Person = LineToPerson($ExcelLine)
        $PersonExists = CheckExist -Email $Person.Mail -AddressList $AddressList
        

        If( $PersonExists -eq $true ) {
            WriteToLog -Text "Skipping $($Person.Mail), adding AG $($Person.AG)"
            $entry = $AddressList | Where-Object { $_.Mail -eq "$($Person.Mail)"}
            $entry.AG += ",$($Person.AG)"
            $skipped = $skipped+1
        }
        If ( $PersonExists -eq $false ) {
            WriteToLog -Text "Adding $($Person.Mail)"
            $AddressList.Add($Person)
        }
    }
    

    # Statistics for the Excel Import
    $ExcelCount = $ExcelImport.Count
    WriteToLog -Text "###########################################################################################"
    WriteToLog -Text "Excel lines: $ExcelCount`t Skipped $skipped`t$($ExcelFile.Name)"
    WriteToLog -Text "###########################################################################################"

}
Function AddIDGToHeader ($Header,$ExcelFile) {
        
    #Importing the individual distribution group tags from Excel Files. 
    #Column Header is asssumed to Start with V:, Tag is a "ü" (chechmark in Windings)

    #Typical Output of Import-Excel -Path $ExcelFile.FullName -NoHeader | Get-Member | Select-Object Definition
    #     
    # ----------
    # bool Equals(System.Object obj)
    # int GetHashCode()
    # type GetType()
    # string ToString()
    # string P1=Nachname
    # string P2=Vorname
    # string P3=E-Mail-Adresse
    # string P4=V:Portal
    # string P5=V:Hardware
    # string P6=V:Jimdo
    $pattern = '^string P\d=V:'
    $IDGList = $(Import-Excel -Path $ExcelFile.FullName -NoHeader | Get-Member | Select-Object Definition |  Where-Object {$_ -like "*V:*"}).Definition -replace $pattern,''
    If ($IDGList -eq "") { $IDGList = $null }
    $Header_add = @{}
    ForEach ($idgtag in $IDGList) {
        $expr = 
        $Header_add.Add( @{Label="$($ExcelFile.BaseName)-$idgtag"; Expression = {$_."$columnname"}})
    }
    $Header += $Header_add
    return $Header
}
Function LineToPerson ($ExcelLine) {
    # Converts a line of the Import-Excel script to an instance of person
 $retPerson = New-Object -TypeName Person 
 $retPerson.Vorname = $ExcelLine.Vorname
 $retPerson.Nachname = $ExcelLine.Nachname
 $retPerson.Mail = $($ExcelLine.Mail).tolower()
 $retPerson.AG = $ExcelLine.AGName
 $retPerson.IndividualDG = ""
 return $retPerson
}
Function CheckExist ($Email,$AddressList) {
        return $($AddressList.Mail -contains "$Email")
}
Function CheckAndClean ($LAL) {
    
}
