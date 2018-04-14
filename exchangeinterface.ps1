Function Remove-FromGAL ($Comparison) {
    $ToRemove = @($Comparison | Where-Object  {$_.SideIndicator -eq "<="})
    ForEach ($Contact in $ToRemove) {
        Remove-MailContact -Confirm:$false -Identity $c.WindowsEmailAddress
        #WriteToCMD -Text "Remove-MailContact -Confirm:$false -Identity $($Contact.Mail)"
        WriteToLog -Text "$($Contact.Vorname) $($Contact.NachName) ($($Contact.Mail)) was removed from  Global Address List. "
    }
    WriteToLog -Text "$($ToRemove.Count) Contacts have been deleted from the Global Address List."
}

Function Add-ToGAL ($Comparison) {
    $ToAdd = ($Comparison | Where-Object  {$_.SideIndicator -eq "=>"} | Where-Object {($_.Mail -notlike "*fhh-portal*")})
    Foreach ($Contact in $ToAdd) {
        #WriteToCMD -Text "New-MailContact -Name $($Contact.VorName) $($Contact.NachName) -ExternalEmailAddress $($Contact.Mail) -LastName $($Contact.NachName) -FirstName $($Contact.VorName)"
        New-MailContact -Name "$($Contact.VorName) $($Contact.NachName)" -ExternalEmailAddress $Contact.Mail -LastName $Contact.NachName -FirstName $Contact.VorName
        WriteToLog -Text "$($Contact.Vorname) $($Contact.NachName) ($($Contact.Mail)) was added to Global Address List."
}
    WriteToLog -Text "$($ToAdd.Count) Contacts have been added to the Global Address List."
}

Function Add-DG ($ag) {
    if (!(Get-DistributionGroup | Where-Object {$_.Name -eq "Mitglieder AG $ag"}))
    {
       # WriteToCMD -Text "New-DistributionGroup -Name Mitglieder AG $ag -Type Security -DisplayName Mitglieder AG $ag -ManagedBy $User -PrimarySmtpAddress mitglieder_$ag@fhh-portal.de" 
       New-DistributionGroup -Name "Mitglieder AG $ag" -Type Security -DisplayName "Mitglieder AG $ag" -ManagedBy $User -PrimarySmtpAddress "mitglieder_$ag@fhh-portal.de"
    }
}

Function Remove-DGMembers ($Comparison,$DGName) {
    $ToRemove = @($Comparison | Where-Object  {$_.SideIndicator -eq "<="})
    ForEach ($Contact in  $ToRemove) {        
        Remove-DistributionGroupMember -Identity "Mitglieder AG $DGName" -Member $Contact.PrimarySMTPAddress -Confirm:$false
        #WriteToCMD -Text "Remove-DistributionGroupMember -Identity Mitglieder AG $DGName -Member $($Contact.PrimarySMTPAddress) -Confirm:$false"
        WriteToLog -Text "Removed from generic DG: $($Contact.PrimarySMTPAddress)"
    }
}

Function Add-DGMemebers ($Comparison,$DGName)  {
    $ToAdd = @($Comparison| Where-Object  {$_.SideIndicator -eq "=>"})
    ForEach ($Contact in  $ToAdd) {
       Add-DistributionGroupMember -Identity "Mitglieder AG $DGName" -Member $Contact.PrimarySMTPAddress
       #WriteToCMD -Text "Add-DistributionGroupMember -Identity Mitglieder AG $DGName -Member $($Contact.PrimarySMTPAddress)"
       WriteToLog -Text "Added to generic DG: $($Contact.PrimarySMTPAddress)"
    }
}



Function Add-IDG ($IDG,$AG) {
    if (!(Get-DistributionGroup | Where-Object {$_.Name -eq "Verteiler $IDG für AG $AG"}))
    {
       New-DistributionGroup -Name "Verteiler $IDG für AG $AG" -Type Security -DisplayName "Verteiler $IDG für AG $AG" -ManagedBy $User -PrimarySmtpAddress "verteiler_$($IDG)_$($AG)@fhh-portal.de"
       WriteToLog -Text "Added NEW individual DG Verteiler $IDG für AG $AG"
    }
}

Function Remove-IDGMembers ($Comparison,$AGName,$IDGName) {
    $ToRemove = @($Comparison | Where-Object  {$_.SideIndicator -eq "<="})
    ForEach ($Contact in  $ToRemove) {        
        Remove-DistributionGroupMember -Identity "Verteiler $IDGName für AG $AGName" -Member $Contact.PrimarySMTPAddress -Confirm:$false
       # WriteToCMD -Text "Remove DistributionGroupMember -Identity Verteiler $IDGName für AG $AGName -Member $($Contact.PrimarySMTPAddress) -Confirm:$false"
        WriteToLog -Text "Removed from individual DG: $($Contact.PrimarySMTPAddress)"
    }
}

Function Add-IDGMemebers ($Comparison,$AGName,$IDGName)  {
    $ToAdd = @($Comparison| Where-Object  {$_.SideIndicator -eq "=>"})
    ForEach ($Contact in  $ToAdd) {
       Add-DistributionGroupMember -Identity "Verteiler $IDGName für AG $AGName" -Member $Contact.PrimarySMTPAddress
       #WriteToCMD -Text "Add-DistributionGroupMember -Identity Verteiler $IDGName für AG $AGName -Member $($Contact.PrimarySMTPAddress)"
       WriteToLog -Text "Added to individual DG: $($Contact.PrimarySMTPAddress)"
    }
}


Function Set-DistribtionGroups ($LAL,$ExcelList,$GIDGL) {
    # Adding and Removing contacts from remote DG according to the xlsx (AG).

    ForEach ($file in $ExcelList){
        $AGName = $file.BaseName
        $IDGList = $GIDGL | Where-Object { $_.AG -eq "$AGName"}
        WriteToLog -Text "###########################################################################################"
        WriteToLog -Text "########################################################  $AGName ($($IDGList.Count) IDGs)"
        WriteToLog -Text "###########################################################################################"
        
        Add-DG -ag $AGName
        
        $LocalList = $LAL | Where-Object { $_.AG -cmatch "$AGName" } | Select-Object @{Name="PrimarySMTPAddress";Expression={$_.Mail}}
        WriteToLog -Text "Found $($LocalList.Count) Members in XLSX"
        $RemoteList = @(Get-DistributionGroupMember -ResultSize Unlimited -Identity "Mitglieder AG $AGName" | Select-Object PrimarySMTPAddress)
        WriteToLog -Text "Found $($RemoteList.Count) Members on Generic DG"
        $Comp = @(Compare-Object -ReferenceObject $RemoteList -DifferenceObject $LocalList -Property PrimarySMTPAddress -PassThru)
        WriteToLog -Text "Therefore we have $($Comp.Count) Actions!"

        Remove-DGMembers -Comparison $Comp -DGName $AGName
        Add-DGMemebers -Comparison $Comp -DGName $AGName
    }
    ForEach ($element in $GIDGL) {
        $IDGName = $($element.XLSXHead).Split(":")[1]
        $AGName = $element.AG
        WriteToLog -Text "###########################################################################################"
        WriteToLog -Text "########################################################  IDG $IDGName of $AGName"
        WriteToLog -Text "###########################################################################################"
        Add-IDG -IDG $IDGName -AG $AGName

        $LocalIDGList = @($LAL | Where-Object { ($_.AG -cmatch "$AGName") -and ($_.IndividualDG -cmatch "$AGName-$IDGName")  } | Select-Object @{Name="PrimarySMTPAddress";Expression={$_.Mail}})
        WriteToLog -Text "Found $($LocalIDGList.Count) IDG Members in XLSX"
        $RemoteIDGList = @(Get-DistributionGroupMember -ResultSize Unlimited -Identity "Verteiler $IDGName für AG $AGName" | Select-Object PrimarySMTPAddress)
        WriteToLog -Text "Found $($RemoteIDGList.Count) IDG Members on Exchange"
        $Comp = @(Compare-Object -ReferenceObject $RemoteIDGList -DifferenceObject $LocalIDGList -Property PrimarySMTPAddress -PassThru)
        Remove-IDGMembers -Comparison $Comp -AGName $AGName -IDGName $IDGName
        Add-IDGMemebers -Comparison $Comp -AGName $AGName -IDGName $IDGName
        
    }
}


