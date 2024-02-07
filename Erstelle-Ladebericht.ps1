#Requires -Modules ImportExcel

#Parameters
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$CSVPath
)

Import-Module ImportExcel

#Parameter einlesen
$parameter = Get-Content -Path '.\param.json' | ConvertFrom-Json

#Parameter setzten
$Kennzeichen = $parameter.Kennzeichen
$BetragkWh = $parameter.BetragkWh
$Mitarbeiter = $parameter.Mitarbeiter
$Culture = [cultureinfo]::GetCultureInfo('de-AT')

# CSV Importieren
$Ladungen = Import-Csv -Path $CSVPath

# Prüfung ob es leere Elemente bei den geladenen kWh gibt
$Ladungen | ForEach-Object {
    if ($_.used -eq "") {
        Write-Error "CSV enthält leere Werte bei Ladungen, bitte das CSV kontrollieren. Aufgabe wird abgebrochen"
        Start-Sleep -Seconds 60
        Exit
    }
}

# Spalte Used auf Dezimalzahl ändern
$Ladungen | ForEach-Object {

    $_.used = $_.used.substring(0, $_.used.length -4).replace(".",",")
}

#Monat finden
$datum = $Ladungen.date[0].Substring(0,  $ladungen.date[0].length - 10)
$datum = $datum.ToDateTime($Culture)
$berichtsmonat = (Get-Culture).DateTimeFormat.GetMonthName(($datum.Month))
$Jahr = Get-Date -Format yyyy
# Notendinge Spalten aus dem CSV übernehmen
$stub = $Ladungen | Select-Object Date, End_date, Location, Used

#CSV Manipulieren
$stub | ForEach-Object {
    $_ | Add-Member -Type NoteProperty -Name 'Kennzeichen' -Value $Kennzeichen
    $Kosten = [int]$_.used.replace(",",".") * $BetragkWh
    $gesamtkosten = $gesamtkosten + $Kosten
    $_ | Add-Member -Type NoteProperty -Name 'Kosten' -Value $Kosten
}
$newHeader = "Startdatum", "Enddatum", "Ort", "geladene kWh", "Kennzeichen", "Kosten"
$stub = $stub | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | ConvertFrom-CSV -Header $newHeader

# Excel Export
$gesamtkosten = [math]::Round($gesamtkosten,2)
$rows = $stub.rows.count
$summe = $rows + 2
$summe = "F$($summe)"
$Filename = "Ladebericht-$($berichtsmonat)-$($Jahr)-$($Mitarbeiter).xlsx"
$stub | Export-Excel -Path .\$Filename -AutoSize -WorksheetName "$($berichtsmonat)-$($Mitarbeiter)" -TableStyle Light10 -BoldTopRow -KillExcel
$excel = Open-ExcelPackage .\$Filename
$excel."$berichtsmonat-$Mitarbeiter".Cells["$($summe)"].Value = $gesamtkosten
Close-ExcelPackage -ExcelPackage $excel
. .\$Filename