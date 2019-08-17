<#
http://www.powershell-monster.com/scripts/misc/011-create-ics-calendar-file/
http://icalendar.org
RFC5545
#>

$destination = "$env:Temp\TTCSPCalendar.ics"

$header = @"
BEGIN:VCALENDAR
VERSION:2.0
METHOD:PUBLISH
X-MS-OLK-FORCEINSPECTOROPEN:TRUE
PRODID:PowerDbaKlaas
BEGIN:VTIMEZONE
TZID:Europe/Brussels
BEGIN:STANDARD
DTSTART:16011028T030000
RRULE:FREQ=YEARLY;BYDAY=-1SU;BYMONTH=10
TZOFFSETFROM:+0200
TZOFFSETTO:+0100
END:STANDARD
BEGIN:DAYLIGHT
DTSTART:16010325T020000
RRULE:FREQ=YEARLY;BYDAY=-1SU;BYMONTH=3
TZOFFSETFROM:+0100
TZOFFSETTO:+0200
END:DAYLIGHT
END:VTIMEZONE
"@

Set-Content -Value $header -Path $destination -Encoding UTF8

<#
UID moet globally uniek zijn! => datum+tijd+ploeg+@ttcsp.be??
deze UID moet ook gebruikt worden bij method update of cancel
DTStamp: tijdstip waarop de ics file gemaakt werd, wijzigt dus bij elke update
CREATED: tijdstip waarop de informatie van het event gecreëerd werd in een calendar application,
wijzigt niet bij updates
LAST-MODIFIED: tijdstip laatste wijziging van een event
CREATED & LAST-MODIFIED worden door de applicatie ingevuld
#>

$games = Import-Csv $env:Temp\Wedstrijden.csv -Delimiter ';'
foreach ( $game in $games )
{
    $startdate = [datetime]::parseexact($game.datum, 'd/MM/yyyy H:mm:ss', $null)
    $starttime = [datetime]::parseexact($game.beginuur, 'H:mm', $null)
    $start = $startdate.tostring("yyyyMMdd") + 'T' + $starttime.tostring("HHmmss")
    $endtime = $starttime.AddMinutes(209)
    $end = $startdate.tostring("yyyyMMdd") + 'T' + $endtime.tostring("HHmmss")
    $now = $(Get-Date).tostring("yyyyMMdd") + 'T' + $(Get-Date).tostring("HHmmss")
    $location = "$($game.locatie ),`n $($game.adres), $($game.gemeente)"
    $summary = "$($game.thuisploeg) - $($game.bezoekers)"
    $UID = "$($game.thuisploeg.Replace(' ',''))$start@TTCSP.be"
    $description = ""

    $event = @"
BEGIN:VEVENT
DTSTART;TZID=Europe/Brussels:$start
DTEND;TZID=Europe/Brussels:$end
UID:$UID
DTSTAMP;TZID=Europe/Brussels:$now
LOCATION:$location
SEQUENCE:0
BEGIN:VALARM
ACTION:DISPLAY
DESCRIPTION:REMINDER
TRIGGER;RELATED=START:-PT02H00M00S
END:VALARM
DESCRIPTION:$description
SUMMARY:$summary
END:VEVENT
"@

    Add-Content -Value $event -Path $destination -Encoding UTF8
}

$closing = @"
END:VCALENDAR
"@

Add-Content -Value $closing -Path $destination -Encoding UTF8