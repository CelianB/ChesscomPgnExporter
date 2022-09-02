$Host.UI.RawUI.WindowTitle = "ChesscomPgnExporter | $([char]0x00A9) C$([char]0x00E9)lian BASTIEN & MauLotu"

$user = Read-Host "Username"
try {                                             ## Send the request to Chesscom api endpoint
	$chessComArchivesUrl = "https://api.chess.com/pub/player/$user/games/archives"
	$archivesReq = (Invoke-WebRequest -URI $chessComArchivesUrl | ConvertFrom-Json).archives 
}
catch { Write-Error("Chesscom username is unknown") }

$outFolder = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$year      = Read-Host "Year"
$outPath   = Join-Path -Path $outFolder -ChildPath "$user $year.pgn"
New-Item -Path $outPath -ItemType "file" -Force   ## Create file to Downloads folder
write-host "Saving games to $outPath"
$i = 0
foreach ($archive_url in $archivesReq) {          ## archive_url : https://...$user/games/yyyy/mm
    $months_percent = ($i++ * 100)/$archivesReq.length
    $splitter = $archive_url -split '/'
    $m = $splitter[-1]
    $y = $splitter[-2]
    Write-Progress -ID 2 -Activity Updating -Status "$m/$y" -PercentComplete $months_percent
	if ($y -eq $year) {
		$games = (Invoke-WebRequest -URI $archive_url | ConvertFrom-Json).games
		write-host "Downloaded $($games.length) games from $archive_url"
		$u = 0
		foreach ($game in $games) {
			$gamesPercent = ($u++ * 100) / $games.length
			Write-Progress -ID 1 -Activity Updating -Status "Storing games..." -PercentComplete $gamesPercent
			$pgn = "$($game.pgn)`r`n"
			Add-Content $outPath $pgn
		}
	} else { write-host "Skip $y/$m" }
}
pause