# office365-powershell
Powershell scripts to manage Office 365 Exchange Server
by Falk Lindner

Workflow is as follows:

- Connects to a SharePoint Folder and Downloads all xlsx files
- Imports all Contacts from xlsx Files (Expected Colums are Nachname, Vorname, E-Mail Adrese)
- Connects to Exchange Server on Office 365
- Compares Global Address List with List of all Contacts in all the xlsx's
- Adds / Removes Contacts accordingly
- Creates Distribution groups according to xlsx Files (Each file is one DG)
- Adds / Removes Contacts from DGs accordingly
- Takes into account sub-distribution groups via special colums (Starting with V:) 
- Adds/ Removes Contacts from sub-distribtution groups


Depends on 
- Import-Excel https://github.com/dfinke/ImportExcel
- PnP-PowerShell https://github.com/SharePoint/PnP-PowerShell


