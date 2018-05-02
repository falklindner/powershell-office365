Function Get-LAL ($ExcelList, $GlobalIDGList) {
    $LocalAddressList = New-Object System.Collections.Generic.List[Person] 
    # Importing the Contacts from each Excel File in the Base path into a list of instances of type person

    foreach ($ExcelFile in $ExcelList) {
      Convert-ExcelToLAL -ExcelFile $ExcelFile -AddressList $LocalAddressList -GIDGL $GlobalIDGList
      $InternalContacts = $LocalAddressList | Where-Object { $_.Mail -like "*fhh-portal*" }
      WriteToLog -Text "###########  $($ExcelFile.Name) done. $($LocalAddressList.Count) Persons identified ($($InternalContacts.Count) FHH Members)."
      }     
    Return $LocalAddressList
}

Function Get-GlobalIDGs  ($ExcelList) {
    $GIDGL = @()
 ForEach ($ExcelFile in $ExcelList) {
        #Prepares an Expression List for the Select Statement of Import-Excel (defines header of columes)
        #Obtaining raw headers from xlsx
        $raw = Import-Excel -Path $ExcelFile.FullName -NoHeader -StartRow 1 | Select-Object -First 1

        #Filtering for custom members, i.e. the columns     
        #
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
        $pattern = 'string P\d='
        $formatted_header = $($raw | Get-Member | Where-Object { $_.Definition -match $pattern}) | `
                                                Select-Object @{Label = "AG"; Expression= { $ExcelFile.BaseName }},`
                                                                @{Label = "ID"; Expression = {$_.Name}},`
                                                                @{Label = "XLSXHead"; Expression ={$_.Definition -replace $pattern,''}}
            
        
        #Adding ind. Distribution Group columns, which start with "V:". 
        $IDGList = $formatted_header | Where-Object { $_.XLSXHead -like "V:*"}

        

        $GIDGL += $IDGList
        # IDGList then looks like 
        # ID XLSXHead
        # -- --------
        # P4 V:Portal
        # P5 V:Hardware
        # P6 V:Jimdo
    }    
    return $GIDGL
}

Function Convert-ExcelToLAL ($ExcelFile, $AddressList, $GIDGL){

    $IDGList = $GIDGL | Where-Object { $_.AG -eq "$($ExcelFile.BaseName)" }
    $Header = PrepareExcelHeader -ExcelFile $ExcelFile -IDGList $IDGList
   

    #Importing contact details from Excel File and writing them into global list $AddressList
    $ExcelImport = Import-Excel -Path $ExcelFile.FullName -NoHeader -StartRow 2 | `
    Select-Object -Property $Header | `
    Where-Object {($_.Nachname -or $_.Vorname) -and $_.Mail}
    
    
    [Int]$skipped = 0
    WriteToLog -Text "###########################################################################################"
    WriteToLog -Text "############################################################  $($ExcelFile.Name)"
    WriteToLog -Text "###########################################################################################"
    
    foreach ($ExcelLine in $ExcelImport) {
        $IDGString = Get-Idgstring -ExcelLine $ExcelLine -IDGList $IDGList
        $Person = LineToPerson -ExcelLine $ExcelLine -IDGString $IDGString
        $PersonExists = CheckExist -Email $Person.Mail -AddressList $AddressList
        

        If( $PersonExists -eq $true ) {
            WriteToLog -Text "Skipping $($Person.Mail), adding AG $($Person.AG) and IDGs"
            $entry = $AddressList | Where-Object { $_.Mail -eq "$($Person.Mail)"}
            $entry.AG += ",$($Person.AG)"
            $entry.IndividualDG += $IDGString
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
Function PrepareExcelHeader  ($ExcelFile, $IDGList) {
    
    #Typical header looks like
    #
        # ID XLSXHead
        # -- --------
        # P1 Nachname
        # P2 Vorname
        # P3 E-Mail-Adresse
        # P4 V:Portal
        # P5 V:Hardware
        # P6 V:Jimdo
    #Assumption is, that P1 is always Nachname, P2 Vorname and P3 Mail address

    $Expr_Head =  @{Label="Nachname"; Expression = {$_.P1}},`
                  @{Label="Vorname" ; Expression = {$_.P2}},`
                  @{Label="Mail"    ; Expression = {$_.P3}},`
                  @{Label="AGName"  ; Expression = {$ExcelFile.BaseName}}
    


    #Generating additional Expression statements, containing the IDGs from $IDGList
    #Typical Output of 
    Foreach ($head in $IDGList) {
        $raw_expr = "`$_.$($head.ID)"
        $expr = [scriptblock]::Create($raw_expr)
        $idgname = $head.XLSXHead.Trim(" ") -replace 'V:',''
        $idgexpr = @{Label="$($ExcelFile.BaseName)-$idgname"; Expression = $expr}
        $Expr_Head += $idgexpr
    }
    return $Expr_Head
}

Function Get-Idgstring ($ExcelLine,$IDGList) {
    $string = ""
    ForEach ($idg in $IDGList) {
        $idgtag = $idg.XLSXHead -replace "V:", "$($ExcelLine.AGName)-"
        If ($ExcelLine.$idgtag -eq "ü") {
            $string += "$idgtag,"
        }
    }
    return $string
}


Function LineToPerson ($ExcelLine,$IDGString) {
    # Converts a line of the Import-Excel script to an instance of person
 $retPerson = New-Object -TypeName Person 
 $retPerson.Vorname = $ExcelLine.Vorname
 $retPerson.Nachname = $ExcelLine.Nachname
 $retPerson.Mail = $($ExcelLine.Mail).tolower()
 $retPerson.AG = $ExcelLine.AGName
 $retPerson.IndividualDG = $IDGString
 return $retPerson
}
Function CheckExist ($Email,$AddressList) {
        return $($AddressList.Mail -contains "$Email")
}