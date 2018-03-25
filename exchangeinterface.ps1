Function RemoveFromGAL ($Comparison) {
    $ToRemove = @($Comparison | Where-Object  {$_.SideIndicator -eq "<="})
    ForEach ($Contact in $ToRemove) {
        #Remove-MailContact -Confirm:$false -Identity $c.WindowsEmailAddress
        WriteToCMD -Text "Remove-MailContact -Confirm:$false -Identity $($Contact.Mail)"
        WriteToLog -Text "$($Contact.Vorname) $($Contact.NachName) ($($Contact.Mail)) was removed from  Global Address List. "
    }
    WriteToLog -Text "$($ToRemove.Count) Contacts have been deleted from the Global Address List."
}

Function AddToGAL ($Comparison) {
    $ToAdd = ($Comparison | Where-Object  {$_.SideIndicator -eq "=>"} | Where-Object {($_.Mail -notlike "*fhh-portal*")})
    Foreach ($Contact in $ToAdd) {
        WriteToCMD -Text "New-MailContact -Name $($Contact.VorName) $($Contact.NachName) -ExternalEmailAddress $($Contact.Mail) -LastName $($Contact.NachName) -FirstName $($Contact.VorName)"
        #New-MailContact -Name "$($Contact.VorName) $($Contact.NachName)" -ExternalEmailAddress $Contact.Mail -LastName $Contact.NachName -FirstName $Contact.VorName
        WriteToLog -Text "$($Contact.Vorname) $($Contact.NachName) ($($Contact.Mail)) was added to Global Address List."
}
    WriteToLog -Text "$($ToAdd.Count) Contacts have been added to the Global Address List."
}

Function CheckDG ($ag) {
    if (!(Get-DistributionGroup | Where-Object {$_.Name -eq "Mitglieder AG $ag"}))
    {
        WriteToCMD -Text "New-DistributionGroup -Name Mitglieder AG $ag -Type Security -DisplayName Mitglieder AG $ag -ManagedBy $User -PrimarySmtpAddress mitglieder_$ag@fhh-portal.de" 
       # New-DistributionGroup -Name "Mitglieder AG $ag" -Type Security -DisplayName "Mitglieder AG $ag" -ManagedBy $User -PrimarySmtpAddress "mitglieder_$ag@fhh-portal.de"
    }
}

Function DeleteDGMembers ($Comparison,$DGName) {
    $ToRemove = @($Comparison | Where-Object  {$_.SideIndicator -eq "<="})
    ForEach ($Contact in  $ToRemove) {        
       # Remove-DistributionGroupMember -Identity "Mitglieder AG $DGName" -Member $Contact.PrimarySMTPAddress -Confirm:$false
        WriteToCMD -Text "Remove-DistributionGroupMember -Identity Mitglieder AG $DGName -Member $($Contact.PrimarySMTPAddress) -Confirm:$false"
        WriteToLog -Text "Removed from generic DG: $($Contact.PrimarySMTPAddress)"
    }
}

Function AddDGMemebers ($Comparison,$DGName)  {
    $ToAdd = @($Comparison| Where-Object  {$_.SideIndicator -eq "=>"})
    ForEach ($Contact in  $ToAdd) {
       # Add-DistributionGroupMember -Identity "Mitglieder AG $DGName" -Member $Contact.PrimarySMTPAddress
       WriteToCMD -Text "Add-DistributionGroupMember -Identity Mitglieder AG $DGName -Member $($Contact.PrimarySMTPAddress)"
       WriteToLog -Text "Added to generic DG: $($Contact.PrimarySMTPAddress)"
    }
}



Function CheckIDG ($IDG,$AG) {
    if (!(Get-DistributionGroup | Where-Object {$_.Name -eq "Verteiler $IDG für AG $AG"}))
    {
        WriteToCMD -Text "New-DistributionGroup -Name Mitglieder AG $AG -Type Security -DisplayName Mitglieder AG $AG -ManagedBy $User -PrimarySmtpAddress mitglieder_$ag@fhh-portal.de" 
       # New-DistributionGroup -Name "Mitglieder AG $ag" -Type Security -DisplayName "Mitglieder AG $ag" -ManagedBy $User -PrimarySmtpAddress "mitglieder_$ag@fhh-portal.de"
    }
}

Function DeleteIDGMembers ($Comparison,$DGName) {
    $ToRemove = @($Comparison | Where-Object  {$_.SideIndicator -eq "<="})
    ForEach ($Contact in  $ToRemove) {        
       # Remove-DistributionGroupMember -Identity "Mitglieder AG $DGName" -Member $Contact.PrimarySMTPAddress -Confirm:$false
        WriteToCMD -Text "Remove DistributionGroupMember -Identity Mitglieder AG $DGName -Member $($Contact.PrimarySMTPAddress) -Confirm:$false"
        WriteToLog -Text "Removed from individual DG: $($Contact.PrimarySMTPAddress)"
    }
}

Function AddIDGMemebers ($Comparison,$DGName)  {
    $ToAdd = @($Comparison| Where-Object  {$_.SideIndicator -eq "=>"})
    ForEach ($Contact in  $ToAdd) {
       # Add-DistributionGroupMember -Identity "Mitglieder AG $DGName" -Member $Contact.PrimarySMTPAddress
       WriteToCMD -Text "Add-DistributionGroupMember -Identity Mitglieder AG $DGName -Member $($Contact.PrimarySMTPAddress)"
       WriteToLog -Text "Added to generic DG: $($Contact.PrimarySMTPAddress)"
    }
}


Function ManageDistribtionGroups ($LAL,$ExcelList) {
    ForEach ($file in $ExcelList){
        $FilePath = $file.FullName
        $AGName = $file.BaseName
        WriteToLog -Text "###########################################################################################"
        WriteToLog -Text "########################################################  $AGName ($($IDGList.Count) IDGs)"
        WriteToLog -Text "###########################################################################################"
        
        CheckDG -ag $AGName
        
        $LocalList = $LAL | Where-Object { $_.AG -cmatch "$AGName" } | Select-Object @{Name="PrimarySMTPAddress";Expression={$_.Mail}}
        WriteToLog -Text "Found $($LocalList.Count) Members in XLSX"
        $RemoteList = @(Get-DistributionGroupMember -ResultSize Unlimited -Identity "Mitglieder AG $AGName" | Select-Object PrimarySMTPAddress)
        WriteToLog -Text "Found $($RemoteList.Count) Members on Generic DG"
        $Comp = @(Compare-Object -ReferenceObject $RemoteList -DifferenceObject $LocalList -Property PrimarySMTPAddress -PassThru)
        WriteToLog -Text "Therefore we have $($Comp.Count) Actions!"

        DeleteDGMembers -Comparison $Comp -DGName $AGName
        AddDGMemebers -Comparison $Comp -DGName $AGName
        
        ForEach ($element in $IDGList) {
            If( $element -eq $null ) {Break}
            $idg = $element.Split(":")[1]
            $columnname = "V:$($idg)"
            WriteToLog -Text "###########################################################################################"
            WriteToLog -Text "########################################################  IDG $idg of $AGName"
            WriteToLog -Text "###########################################################################################"
            CheckIDG -IDG $idg -AG $AGName

            $LocalIDGList = $LAL | Where-Object { ($_.AG -cmatch "$AGName") -and ($_.($columnname) -eq "ü") } | Select-Object @{Name="PrimarySMTPAddress";Expression={$_.Mail}} 
            WriteToLog -Text "Found $($LocalIDGList.Count) IDG Members in XLSX"
            $RemoteIDGList = @(Get-DistributionGroupMember -ResultSize Unlimited -Identity "Verteiler $idg für AG $AGName" | Select-Object PrimarySMTPAddress)
            WriteToLog -Text "Found $($RemoteIDGList.Count) IDG Members in XLSX"
            $Comp = @(Compare-Object -ReferenceObject $RemoteIDGList -DifferenceObject $LocalIDGList -Property PrimarySMTPAddress -PassThru)
        }
    }
}


