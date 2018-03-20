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
Function LineToPerson ($ExcelLine) {
    # Converts a line of the Import-Excel script to an instance of person
 $retPerson = New-Object -TypeName Person 
 $retPerson.Vorname = $ExcelLine.FirstName
 $retPerson.Nachname = $ExcelLine.LastName
 $retPerson.Mail = $($ExcelLine.WindowsEmailAddress).tolower()
 $retPerson.AG = $ExcelLine.AGName
 return $retPerson
}
Function CheckExist ($Email,$AddressList) {
        return $($AddressList.Mail -contains "$Email")
}
Function ExcelToLAL ($ExcelFile, $AddressList){
    
    $ExcelImport = Import-Excel -Path $ExcelFile.FullName -HeaderName LastName,FirstName,WindowsEmailAddress -StartRow 2 | `
    Where-Object {($_.LastName -or $_.FirstName) -and $_.WindowsEmailAddress} | `
    Select-Object -Property LastName,FirstName,WindowsEmailAddress,@{Label="AGName";Expression={$ExcelFile.BaseName}}
    
    # A separate list for statistics purposes. The return list is $AddressList
    $DummyList = New-Object System.Collections.Generic.List[Person] 
    [Int]$skipped = 0
    WriteToLog -Text "###########################################################################################"
    WriteToLog -Text "############################################################  $($ExcelFile.Name)"
    WriteToLog -Text "###########################################################################################"
    
    foreach ($ExcelLine in $ExcelImport) {
        $Person = LineToPerson($ExcelLine)
        $DummyList.Add($Person)
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
    $ImportCount = $DummyList.Count
    WriteToLog -Text "###########################################################################################"
    WriteToLog -Text "Excel lines: $ExcelCount`tImported Contacts: $ImportCount`t Skipped $skipped`t$($ExcelFile.Name)"
    WriteToLog -Text "###########################################################################################"

}

Function CheckAndClean ($LAL) {
    
}