
######
# Settings
######
$workingDirectory = Join-Path $env:USERPROFILE ".h8-wallpaper"


$latestDataUri = "http://himawari8-dl.nict.go.jp/himawari8/img/D531106/latest.json"
$imageBaseUri = "http://himawari8-dl.nict.go.jp/himawari8/img/D531106/2d/550/"

######
# Script Logic
######

$h8info = @{
    "XCount" = 2
    "YCount" = 2
    "LatestStampFile" = Join-Path $workingDirectory "latest.txt"
    }

$AlreadyUpToDate = $false

# Create the working directory if it does not exist
if (-not (Test-Path $workingDirectory)) {
    $wd = New-Item $workingDirectory -ItemType Directory
    $wd.Attributes = "hidden"
    $wd = $null
    "Created working directory: '$workingDirectory"
}

$wallpaperFile = Join-Path $workingDirectory "h8-wallpaper.png"

# Load the System.Drawing namespace so we can do cool image stuff
[void][reflection.assembly]::loadwithpartialname("system.drawing")

# Grab the metadata file telling us when the latest imagery was uploaded
$latestData = (Invoke-WebRequest $latestDataUri -ErrorAction Stop).Content | 
                ConvertFrom-Json -ErrorAction Stop

# Build up the file system path for the images
$latest = get-date $latestData[0].date -f "yyyy/MM/dd/HHmmss"

# Check to see if we already have the latest imagery 
# (based on timestamp from retrieved latest.json)
if (Test-Path $h8info.LatestStampFile) {
    $StampContent = Get-Content $h8info.LatestStampFile
    if ((Get-Date $latestData[0].date) -eq (Get-Date $stampContent)) {
        # If the timestamp in the file on disk matches the latest.json,
        # It means we're already up to date, no changes
        $AlreadyUpToDate = $true
    }
}

if ($AlreadyUpToDate) {
    "No new imagery available. Exiting"
    break;
}

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
    "Grabbing: $($image.value)"
    $imgOutPath = "$workingDirectory\h8_$($image.name).png"
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
    "Placing $image at $xCoord,$yCoord"
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

# Write out the timestamp file so we don't update unnecessarily
(Get-Date $latestData[0].date).ToString() | Set-Content $h8info.LatestStampFile