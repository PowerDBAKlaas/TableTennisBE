<#
.SYNOPSIS
    Gets the members and matches for a Belgian tabletennis club. 

.DESCRIPTION
    Queries the matches and members of a Belgian tabletennis club, and the addresses of the venues of all clubs, both on the VTTL and Sporta web sites.
    The results are flattened into a csv and saved in the users Temp directory.
    From there they can be imported by Power Query or any other application of choice.

.NOTES
    Author: Klaas Vandenberghe ( @PowerDBAKlaas )
    Date:   2019-07-28
    Website: https://www.powerdba.eu

.LINK
    https://ttonline.sporta.be/api/?wsdl
    https://api.vttl.be/0.7/?wsdl
    http://tabt.frenoy.net/index.php?l=NL&display=TabTAPI_NL
    http://api.frenoy.net/group__TabTAPIfunctions.html

.EXAMPLE
    PS> .\Get-TTFrenoy.ps1 -VTTLID OVL125 -SportaID 4046

    Gets the members and matches for TTC Sint-Pauwels.
#>
[cmdletbinding()]
Param (
    [Parameter(Mandatory = $false, position = 0)]
    [String]$VTTLID,

    [Parameter(Mandatory = $false, position = 1)]
    [String]$SPORTAID
)

$TargetDir = New-Item -ItemType Directory -Force -Path "$env:TEMP\TTFrenoy"
get-childitem -Path $TargetDir -include *.csv -recurse | remove-item
$matchescsv = "$TargetDir\Wedstrijden.csv"
$memberscsv = "$TargetDir\leden.csv"
$addressescsv = "$TargetDir\venues.csv"

$leagues = @(
    @{
        uri      = "https://ttonline.sporta.be/api/?wsdl"
        name     = "SPORTA"
        ClubId   = $SPORTAID
        matchURI = "https://ttonline.sporta.be/wedstrijd/"
    }
    ,
    @{
        uri      = "http://api.vttl.be/0.7/?wsdl"
        name     = "VTTL"
        ClubId   = $VTTLID
        matchURI = "http://competitie.vttl.be/match/"
    }
)

foreach ( $league in $leagues ) {
    if ($league.clubID) {
        $prox = New-WebServiceProxy -Uri $($league.uri)
        $NS = $prox.GetType().namespace

        $Cl = New-Object ($NS + '.GetClubs')
        $Clubs = $prox.GetClubs($Cl)
        $SelectedClubs = $clubs.ClubEntries | Where-Object { $_.name -notin ('Buitenland', 'KBTTB/FRBTT') }
        foreach ($club in $SelectedClubs) {
            $null = $adres
            $ClubProp = @{
                League       = $league.name;
                UniqueIndex  = $club.UniqueIndex; 
                Name         = $club.name;
                LongName     = $club.LongName;
                Category     = $club.Category;
                CategoryName = $club.CategoryName;
            }
            foreach ($venue in $club.venueentries) {
                $VenueProp = $ClubProp.Clone()
                $VenueProp.add('VenueID', $venue.id)
                $VenueProp.add('VenueVolgNr', $venue.clubvenue)
                $VenueProp.add('VenueNaam', $VenueProp.name.Replace("`"", ""))
                $VenueProp.add('Adres', $venue.Street.Replace(",", ""))
                $VenueProp.add('Gemeente', $venue.town)
                $VenueProp.add('Telefoon', $venue.phone)
                $VenueProp.add('Commentaar', $($venue.Comment.Replace("`"", "") -replace '\r\n', ' / '))
            
                $Adres = New-Object -TypeName PSObject -Property $VenueProp
                $adres | Export-Csv -Path $addressescsv -NoClobber -NoTypeInformation -Delimiter ';' -Force -Append -Encoding UTF8
            }
        }
    
        $Mb = New-Object ($NS + '.GetMembersRequest')
        $Mb.Club = $($league.ClubId)
        $Mb.WithResults = $true
        $leden = $prox.GetMembers($Mb)
        $leden.MemberEntries | Select-Object @{l = 'League'; e = { $($league.name) } }, @{l = 'Club'; e = { $Mb.Club } }, UniqueIndex, Position, LastName, firstname, Ranking, RankingIndex |
        Export-Csv -Path $memberscsv -NoClobber -NoTypeInformation -Delimiter ';' -Force -Append -Encoding UTF8
        if ($league.name -eq 'Sporta') {
            # Add doubles player to join on Wedstrijden
            "`"SPORTA`";`"$SportaID`";`"222222`";`"999`";`"Dubbel`";`"Dubbel`";`"Q8`";`"999`"" |
            Add-Content -Path $memberscsv
        }
            
        $Ma = New-Object ($NS + '.GetMatchesRequest')
        $Ma.Club = $($league.ClubId)
        $Ma.ShowDivisionNameSpecified = $true
        $Ma.ShowDivisionName = "yes"
        $Ma.WithDetails = "yes"
        $Ma.WithDetailsSpecified = "yes"
        foreach ($seizoen in 20, 19, 18, 17, 16, 15) {  
            $Ma.Season = $seizoen
            $wedstr = $null
            $wedstr = $prox.GetMatches($Ma)

            foreach ($match in $wedstr.TeamMatchesEntries) {
                $null = $detail
                #$adres = $addresses | Where-Object { $_.clubID -match $match.Homeclub -and $_.clubvenue -eq $match.Venue }
                $prop = @{
                    Datum         = "{0:yyyy'-'MM'-'dd}" -f $match.date;
                    Beginuur      = $match.time.toshorttimestring();
                    # Bond          = $($league.name);
                    Afdeling      = $match.DivisionName;
                    ThuisClub     = $match.HomeClub;
                    Thuisploeg    = $match.hometeam;
                    BezoekersClub = $match.AwayClub;
                    Bezoekers     = $match.AwayTeam;
                    OnzeClub      = $league.ClubId;
                    Uitslag       = $match.Score;
                    MatchID       = $match.MatchUniqueID;
                    ClubVenue     = $match.VenueClub;
                    Venue         = $match.Venue;
                    #Locatie       = if ($adres) { $adres.name.Replace("`"", "") } else { "" };
                    #Adres         = if ($adres) { $adres.Street.Replace(",", "") } else { "" };
                    #Gemeente      = $adres.Town;
                    #Info          = if ($adres) { $adres.Comment.Replace("`"", "") -replace '\r\n', ' / ' } else { "" };
                }
                foreach ( $detail in $match.MatchDetails.individualMatchResults) {
                    $Det = $prop.clone();
                    if ($detail) {
                        $Det.add('ThuisSets', $detail.Homesetcount);
                        $Det.add('UitSets', $detail.Awaysetcount);
                        $Det.add('WedstrijdVolgNr', $detail.Position);
                        $Det.add('Scores', $detail.Scores);
                        $Det.add('ThuisForfait', $detail.IsHomeForfeited);
                        $Det.add('UitForfait', $detail.IsAwayForfeited);
                        $Det.add('ThuisVolgNr', $($detail.HomePlayerMatchIndex[0]));
                        $Det.add('UitVolgNr', $($detail.AwayPlayerMatchIndex[0]));
                        if ($detail.Position -eq 7 -and $league.name -eq "Sporta") {
                            # Add dummy doubles player
                            $Det.add('ThuisSpeler', 222222);
                            $Det.add('UitSpeler', 222222); 
                        }
                        else {
                            $Det.add('ThuisSpeler', $($detail.HomePlayerUniqueIndex[0]));
                            $Det.add('UitSpeler', $($detail.AwayPlayerUniqueIndex[0]));
                        }
                    }
                    else {
                        $Det.add('ThuisSets', 0);
                        $Det.add('UitSets', 0);
                        $Det.add('WedstrijdVolgNr', 0);
                        $Det.add('Scores', '0-0');
                        $Det.add('ThuisForfait', $false);
                        $Det.add('UitForfait', $false);
                        $Det.add('ThuisVolgNr', 0);
                        $Det.add('UitVolgNr', 0);
                        $Det.add('ThuisSpeler', $null);
                        $Det.add('UitSpeler', $null);
                    }
                    $Wedstrijd = New-Object -TypeName PSObject -Property $Det
                    if ($wedstrijd.Datum -notlike '*01-01-01*') { 
                        $wedstrijd | Export-Csv -Path $matchescsv -NoClobber -NoTypeInformation -Delimiter ';' -Force -Append -Encoding UTF8
                    }
                }   # end foreach detail
            }   # end foreach match
        }   # end foreach season
        $prox.Dispose()
    }   # end IF ClubID
}   # end foreach league