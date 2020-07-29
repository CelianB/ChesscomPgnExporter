$user = Read-Host "Type your chess.com username"
## Download folder
$outFolder = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$outFile = "chess_com_games_$user.pgn"
$outPath = Join-Path -Path $outFolder -ChildPath $outFile
New-Item -Path $outPath -ItemType "file" -Force

$chessComArchivesUrl = "https://api.chess.com/pub/player/$user/games/archives"

$archivesReq = (Invoke-WebRequest -URI $chessComArchivesUrl | ConvertFrom-Json).archives 
$i =0
foreach ($archive_url in $archivesReq) {
    ## archive_url : format : https://api.chess.com/pub/player/user/games/yyyy/mm
    $pourcent = ($i*100)/$archivesReq.length
    $splitter = $archive_url -split '/'
    $m = $splitter[-2]
    $y = $splitter[-1]
    Write-Progress -Activity Updating -Status "Downloading $m/$y" -PercentComplete $pourcent
    $i+=1
    $games = (Invoke-WebRequest -URI $archive_url | ConvertFrom-Json).games

    $u = 0
    foreach ($game in $games) {
        $subPourcent = ($u*100)/$games.length
        Write-Progress -Activity Updating -Status "Loading..." -PercentComplete $subPourcent
        $pgn = "$($game.pgn)`r`n"
        Add-Content $outPath $pgn
        $u+=1
    }
}