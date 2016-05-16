Push-Location $PSScriptRoot

$latestDataUri = "http://himawari8-dl.nict.go.jp/himawari8/img/D531106/latest.json"
$imageBaseUri = "http://himawari8-dl.nict.go.jp/himawari8/img/D531106/2d/550/"

$wallpaperFile = Join-Path $env:USERPROFILE "h8-wallpaper.png"

$imageMagickPath = "C:\Program Files\ImageMagick-7.0.1-Q16\magick.exe"

# Grab the metadata file telling us when the latest imagery was uploaded
$latestData = (Invoke-WebRequest $latestDataUri -ErrorAction Stop).Content | ConvertFrom-Json -ErrorAction Stop

# Build up the file system path for the images
$latest = get-date $latestData[0].date -f "yyyy/MM/dd/HHmmss"

"Grabbing image dated $latest"

# Image paths
$images = @{
    "00" = $imageBaseUri + $latest + "_0_0.png"
    "01" = $imageBaseUri + $latest + "_1_0.png"
    "02" = $imageBaseUri + $latest + "_0_1.png"
    "03" = $imageBaseUri + $latest + "_1_1.png"
}

$cmd = ""
# Get the images
foreach ($image in ($images.GetEnumerator() | Sort -Property name )) {
    $imgOutPath = "$env:userprofile\h8-$($image.name).png"
    Invoke-WebRequest -Uri $image.Value -OutFile $imgOutPath
    $cmd+= "$imgOutPath "
}

# ImageMagick doesn't launch properly from PowerShell for some reason, so spit out a batch file and use that
$batchFile = @"
"$imageMagickPath" montage $cmd -mode concatenate -tile 2x $wallpaperFile
"@

$batchFile | Set-content -Encoding String $env:Temp\h8magick.cmd 

&$env:Temp\h8magick.cmd 

# If the wallpaper exists ...
if (Test-Path $wallpaperFile) {
    # Set it to fit and not tile
    Set-ItemProperty 'HKCU:\Control Panel\Desktop' -Name "TileWallpaper" -Value 0
    Set-ItemProperty 'HKCU:\Control Panel\Desktop' -Name "WallPaperStyle" -Value 6

    # Assign the wallpaper
    $dllBits = @'
		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
'@

	$SetWallpaper = Add-Type -MemberDefinition $dllbits -Name "SysParamsInfo" -Namespace "User32Functions" -PassThru

    $SetWallpaper::SystemParametersInfo(20, 0, $wallpaperFile, 3)
}

Pop-Location