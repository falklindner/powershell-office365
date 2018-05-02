<# 
 .Synopsis
  Displays a visual representation of a calendar.

 .Description
  Displays a visual representation of a calendar. This function supports multiple months
  and lets you highlight specific date ranges or days.

 .Parameter Start
  The first month to display.

 .Parameter End
  The last month to display.

 .Parameter FirstDayOfWeek
  The day of the month on which the week begins.

 .Parameter HighlightDay
  Specific days (numbered) to highlight. Used for date ranges like (25..31).
  Date ranges are specified by the Windows PowerShell range syntax. These dates are
  enclosed in square brackets.

 .Parameter HighlightDate
  Specific days (named) to highlight. These dates are surrounded by asterisks.


 .Example
   # Show a default display of this month.
   Show-Calendar

 .Example
   # Display a date range.
   Show-Calendar -Start "March, 2010" -End "May, 2010"

 .Example
   # Highlight a range of days.
   Show-Calendar -HighlightDay (1..10 + 22) -HighlightDate "December 25, 2008"
#>

function Convert-AddressLists (
    [System.IO.FileInfo]  $ExcelList )
{
    $LocalAddressList = New-Object System.Collections.Generic.List[Person]
    $IDGList = @()

    ForEach ($ExcelFile in $ExcelList) {
    $Raw_Excel =  Import-Excel -Path $ExcelFile.FullName -NoHeader -StartRow 1
    $IDGList += Get-IDGList -Raw $Raw_Excel -ExcelFile $ExcelFile
    $Header = Get-ExcelHeader -Raw $Raw_Excel -IDGList $IDGList
    $LocalAddressList += Write-PersonList -Raw $Raw_Excel -Header $Header  -IDGList $IDGList -ExcelFile $ExcelFile -AddressList $LocalAddressList
    return $LocalAddressList
    }
}

function Get-IDGList ($Raw, [System.IO.FileInfo] $ExcelFile) 
{
    $raw_header = $Raw | Select-Object -First 1
    $pattern = 'string P\d='
    $formatted_header = $($raw_header  | Get-Member | Where-Object { $_.Definition -match $pattern}) | `
                                      Select-Object @{Label = "AG"; Expression= { $ExcelFile.BaseName }},`
                                                    @{Label = "ID"; Expression = {$_.Name}},`
                                                    @{Label = "XLSXHead"; Expression ={$_.Definition -replace $pattern,''}},`
                                                    @{Label = "IDGName"; Expression ={$_.Definition -replace $pattern,'' -replace "V:",''}}
    $IDGList = $formatted_header | Where-Object { $_.XLSXHead -like "V:*"}
    return  $IDGList
}

function Get-ExcelHeader ($Raw, $IDGList) 
{
    $Expr_Head =    @{Label="Nachname"; Expression = {$_.P1}},`
                    @{Label="Vorname" ; Expression = {$_.P2}},`
                    @{Label="Mail"    ; Expression = {$_.P3}},`
                    @{Label="AGName"  ; Expression = {$ExcelFile.BaseName}}

    Foreach ($IDG in $IDGList) {
        $raw_expr = "`$_.$($IDG.ID)"
        $expr = [scriptblock]::Create($raw_expr)
        $idgexpr = @{Label="$($IDG.AG)-$($IDG.IDGNAME)"; Expression = $expr}
        $Expr_Head += $idgexpr
    }
    return $Expr_Head
}

function Write-PersonList  (
    $Raw,
    $Header,
    $IDGList,
    [System.IO.FileInfo] $ExcelFile,
    [System.Collections.Generic.List[Person]] $AddresList )
    
{
    $ExcelImport = $Raw | Select-Object -Property $Header -Skip 1 | Where-Object {($_.Nachname -or $_.Vorname) -and $_.Mail}
      
    [Int]$skipped = 0
    WriteToLog -Text "###########################################################################################"
    WriteToLog -Text "############################################################  $($ExcelFile.Name)"
    WriteToLog -Text " "
    
    foreach ($ExcelLine in $ExcelImport) {
        $IDGString = Get-Idgstring -ExcelLine $ExcelLine -IDGList $IDGList
        $Person = Convert-LineToPerson -ExcelLine $ExcelLine -IDGString $IDGString
        $PersonExists = Test-Existing -Email $Person.Mail -AddressList $AddressList
        

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
    WriteToLog -Text " "

}

Function Get-Idgstring ($ExcelLine,$IDGList) {
    $string = ""
    ForEach ($IDG in $IDGList) {
        $idgtag = "$($IDG.AG)-$($IDG.IDGName)"
        If ($ExcelLine.$idgtag -eq "ü") {
            $string += "$idgtag,"
        }
    }
    return $string
}


Function Convert-LineToPerson ($ExcelLine,$IDGString) {
    # Converts a line of the Import-Excel script to an instance of person
 $retPerson = New-Object -TypeName Person 
 $retPerson.Vorname = $ExcelLine.Vorname
 $retPerson.Nachname = $ExcelLine.Nachname
 $retPerson.Mail = $($ExcelLine.Mail).tolower()
 $retPerson.AG = $ExcelLine.AGName
 $retPerson.IndividualDG = $IDGString
 return $retPerson
}

Function Test-Existing ($Email,$AddressList) {
    return $($AddressList.Mail -contains "$Email")
}