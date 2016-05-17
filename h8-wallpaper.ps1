[void][reflection.assembly]::loadwithpartialname("system.drawing")

$latestDataUri = "http://himawari8-dl.nict.go.jp/himawari8/img/D531106/latest.json"
$imageBaseUri = "http://himawari8-dl.nict.go.jp/himawari8/img/D531106/2d/550/"

$wallpaperFile = Join-Path $env:USERPROFILE "h8-wallpaper.png"

$h8info = @{
    "XCount" = 2
    "YCount" = 2
    }
    
# Grab the metadata file telling us when the latest imagery was uploaded
$latestData = (Invoke-WebRequest $latestDataUri -ErrorAction Stop).Content | ConvertFrom-Json -ErrorAction Stop

# Build up the file system path for the images
$latest = get-date $latestData[0].date -f "yyyy/MM/dd/HHmmss"

"Grabbing image dated $latest"

# Image paths
$images = @{
    "0_0" = $imageBaseUri + $latest + "_0_0.png"
    "1_0" = $imageBaseUri + $latest + "_1_0.png"
    "0_1" = $imageBaseUri + $latest + "_0_1.png"
    "1_1" = $imageBaseUri + $latest + "_1_1.png"
}

# Get the images
$imagesToProcess = @()
foreach ($image in $images.GetEnumerator()) {
    $imgOutPath = "$env:userprofile\h8_$($image.name).png"
    Invoke-WebRequest -Uri $image.Value -OutFile $imgOutPath
    $imagesToProcess += $imgOutPath
}

$wallpaperSize = ([System.Drawing.Bitmap]::FromFile($imagesToProcess[0])).PhysicalDimension

$imgResult = [System.Drawing.Bitmap]::new(
                [int]($h8info.XCount * $wallpaperSize.Width),
                [int]($h8info.YCount * $wallpaperSize.Height)
                )

$g = [System.Drawing.Graphics]::FromImage($imgResult)
$g.Clear([System.Drawing.Color]::Black)


foreach ($image in $imagesToProcess) {
    $imgbmp = [System.Drawing.Bitmap]::FromFile($image)
    
    $xCoord = [int]($image.Split("_")[-2]) * [int]$wallpaperSize.Width
    $yCoord = [int]($image.Split("_")[-1]).Split(".")[0] * [int]$wallpaperSize.Height
    "$image`: $xCoord $yCoord"
    $g.DrawImage($imgbmp, [System.Drawing.Rectangle]::new($xCoord,$yCoord,$wallpaperSize.Width, $wallpaperSize.Height))
}

$imgResult.Save($wallpaperFile)


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

# Clean up $WallpaperSize as we don't need it any more
$wallpaperSize = $null
