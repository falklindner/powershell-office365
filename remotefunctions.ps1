Function LoginFHH_Exchange ($cred) 
{
    if (!(Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" })) {
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication "Basic" -AllowRedirection
    Import-PSSession $exchangeSession -DisableNameChecking
    if (!(Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" })) {
        Throw "Connection unsuccessful"
        }   
    }
}

Function CloseFHH 
{
    foreach ($exchangesession in $(Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange" }))
    {   
        Remove-PSSession -Id $exchangeSession.Id
    }
  #  Disconnect-MsolService
}

Function ContactToPerson ($Contact) {
    # Converts a line of the Import-Excel script to an instance of person
 $retPerson = New-Object -TypeName Person 
 $retPerson.Vorname = $Contact.FirstName
 $retPerson.Nachname = $Contact.LastName
 $retPerson.Mail = $($Contact.WindowsEmailAddress).tolower()
 return $retPerson
}

Function BuildGAL {
    $ContactList = Get-Contact -ResultSize Unlimited| Select-Object LastName,FirstName,WindowsEmailAddress
    $GlobalAddressList = New-Object System.Collections.Generic.List[Person] 
    # Importing the Contacts from Office 365 Server into a list of instances of type person
    foreach ($Contact in $ContactList) {
     $Person = ContactToPerson -Contact $Contact 
     $GlobalAddressList.Add($Person)
    }
    WriteToLog -Text "###########################################################################################"
    WriteToLog -Text "################## Exchange GAL imported. $($GlobalAddressList.Count) Persons identified."
    WriteToLog -Text "###########################################################################################"

    Return $GlobalAddressList
}

Function IsInternal ([Person] $Person) {
    [bool] $IsInternal = $false
    If ($Person.Mail -like "*fhh-portal*") {
        $IsInternal = $true
    }
    return $IsInternal
}