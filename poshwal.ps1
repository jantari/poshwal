<#
    .SYNOPSIS
    Generates and optionally applies a terminal theme from the colors found in an image.

    .LINK
    https://www.github.com/jantari

    .PARAMETER ImageFile
    The full path to the image you want to feed as input.
    You can easily get the path of a file from Explorer by holding Shift while right-clicking it and selecting "Copy as Path".
    If no file is provided the script will try to use your current wallpaper as its input.

    .PARAMETER SimilarColorThreshold
    Specify a number between 0 and 1.
    This number will be multiplied with 255 to get the deviation range (plus and minus for the R, G and B values)
    in which a color is considered too similar to another and will therefore be discarded. Setting this Parameter to 0
    means if the colors RGB(100,100,100) and RGB(100,100,101) are extracted from the image, they will both be kept
    and may end up in the color scheme despite being indistinguishable.
    Setting this parameter to 1 means even RGB(0,0,0) and RGB(255,255,255) as well as every color in between will
    be considered too similar to the first one checked and be discarded. The script would only extract one color from
    the image. The default value of this parameter is 0.08 (meaning above 8% deviation is considered disimilar enough).
    The deviation factor is applied to all R, G and B values. A color is only discarded if all of its R, G and B values
    are too similar to a preexisting one.

    .PARAMETER RequiredBrightnessDifference
    Specify a number between 0 and 1.
    The minimum required Brightness difference of all other colors to the background color.
    For dark themes, this means all other colors will have to be X brighter than the background
    and for light themes this means all colors that aren't darker than the background by X will not be used.
    The default value of this parameter is 0.14

    .PARAMETER SimilarColorAlgorithm
    Specify 'RGBTotalAverage' or 'RGBIndividual'.
    With RGBTotalAverage, a color in the image is discarded as being too similar to an already extracted one if the average
    deviation of its Red, Green and Blue values from the already extracted one is within SimilarColorThreshold percent.
    This means the individual value for say Green is allowed to exceed a derivation of 12% even if SimilarColorThreshold is set to 0.12
    if Red and Blue are similar enough to average it out. This is the default because from my testing it seems to produce better results,
    however it is slightly slower.

    With RGBIndividual, no outliers are allowed and each Red, Green and Blue values are checked to be within the derivation percent
    set by SimilarColorThreshold. If only one of the values is over the threshold the color is discarded from further evaluation for
    being too similar to an already known one.

    .PARAMETER NotEnoughColors
    Specify the behavior of the script if not enough (less than 16) colors were extracted from the image.
    Options are "Reuse" to use the same colors multiple times or "Extrapolate" to try and generate fitting colors
    not in the image by variating dominant colors that are in the image. If you are not happy with the results of
    either option, lower the value of "SimilarColorThreshold" to allow the algorithm to extract more colors from the image.

    .PARAMETER LightTheme
    Create a light theme (light background and possibly also pick other colors differently for contrast).
    By default the script will always produce a dark theme.

    .PARAMETER MapColorsByHue
    Enabled by default. Specify -MapColorsByHue:$false to disable this.
    This will "try" to assign the colors in the palette that aren't gray or black to
    the best fitting color from the image so that errors stay red and colored ASCII
    art such as from neofetch somewhat retains its intended appearance.
    If this is disabled, the 16 colors are assigned in order of their prevalence
    (Color 1 - most prevalent) to least prevalent (Color 15) with no regards to
    their appearance/hue. The background and primary text color are
    always calculated seperately and not affected by this setting.

    .PARAMETER Apply
    Apply the generated theme (current user only).

    .PARAMETER SetWallpaper
    Set the specified ImageFile as the desktop wallpaper.

    .PARAMETER dbgOldBrightnessCurve
    This parameter switches to an older brightness curve for deciding the background and primary text colors that isn't as strict on
    backgrounds that are too bright for a dark theme or too dark for a light theme. Will probably remove, only kept for my testing.
#>

Param (
    [CmdletBinding()]
    [Parameter(Position = 0, ValueFromPipelineByPropertyName)]
	[Alias('FullName')]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$ImageFile,
    [ValidateRange(0,1)]
    [float]$SimilarColorThreshold = 0.08,
    [ValidateRange(0,1)]
    [float]$RequiredBrightnessDifference = 0.14,
    [ValidateSet('RGBTotalAverage', 'RGBIndividual')]
    [string]$SimilarColorAlgorithm = 'RGBIndividual',
    [ValidateSet('Reuse', 'Extrapolate')]
    [string]$NotEnoughColors = 'Extrapolate',
    [switch]$LightTheme,
    [switch]$MapColorsByHue = $true,
    [switch]$Apply,
    [ValidateSet('Fill', 'Fit', 'Stretch', 'Center', 'Tile', 'Span')]
    [string]$SetWallpaper,
    [switch]$dbgOldBrightnessCurve
)

# Needed to print the Information-Stream in the console instead of ignoring it
$InformationPreference = "Continue"

# Function ported from C# code of: http://www.blackwasp.co.uk/RGBHSL_3.aspx
function ComponentFromHue {
	Param (
		[decimal]$m1,
		[decimal]$m2,
		[decimal]$h
	)

    $h = ($h + 1) % 1

	if (($h * 6) -lt 1) {
		return $m1 + ($m2 - $m1) * 6 * $h
	} elseif (($h * 2) -lt 1) {
		return $m2
	} elseif (($h * 3) -lt 2) {
		return $m1 + ($m2 - $m1) * ((2 / 3) - $h) * 6
	} else {
		return $m1
	}
}

# Function ported from C# code of: http://www.blackwasp.co.uk/RGBHSL_3.aspx
function Convert-HSLtoColor {
    Param (
        [Parameter(Mandatory = $true)]
        [float]$Hue,
        [Parameter(Mandatory = $true)]
        [float]$Brightness,
        [Parameter(Mandatory = $true)]
        [float]$Saturation
    )

    $Hue = $Hue / 360

    $max = if ($brightness -lt 0.5) {
        $brightness * (1 + $saturation)
    } else {
        $brightness + $saturation - $brightness * $saturation
    }
    $min = $brightness * 2 - $max

    $colorTable = @(
        [Math]::Round(255 * (ComponentFromHue -m1 $min -m2 $max -h ($Hue + (1/3)))) # RED
        [Math]::Round(255 * (ComponentFromHue -m1 $min -m2 $max -h $Hue))           # GREEN
        [Math]::Round(255 * (ComponentFromHue -m1 $min -m2 $max -h ($Hue - (1/3)))) # BLUE
    )

    $colorObject = [System.Drawing.Color]::FromArgb($colorTable[0], $colorTable[1], $colorTable[2])
    return $colorObject
}

# This function compares an array of colors to 6 basic ones and finds
# the most ideal hue match among the input to each of the 6 basic colors
function Get-BestColorMatch {
    Param (
        $Colors,
        [switch]$NoDuplicates,
        [switch]$NoRejects
    )

    $bestFittingColors = @(
        $null, # BLUE
        $null, # GREEN
        $null, # CYAN
        $null, # RED
        $null, # MAGENTA
        $null  # YELLOW
    )

    # We calculate the average saturation of the colors passed in to later maybe not pick an outlier but one that's somewhat close to the avg in the image
    [float]$sat = 0
    foreach ($color in $Colors) {
        $sat += $Color.GetSaturation()
    }
    $avgSaturation = $sat / $Colors.Count
    Write-Verbose "Average saturation of $($Colors.Count) input Colors: $avgSaturation"

    $diffMatrix = New-Object float[][] -ArgumentList 6, $Colors.Count

    # Calculate the diffs from each color to each reference color and store them in a 2D array
    # The information in the array will be laid out like this:
    # InputColor1 InputColor2 InputColor3 ...
    # DIFFVALUE   DIFFVALUE   DIFFVALUE   # DIFFS TO PERFECT BLUE
    # DIFFVALUE   DIFFVALUE   DIFFVALUE   # DIFFS TO PERFECT GREEN
    # ...
    for ($i = 0; $i -lt 6; $i++) {
        for ($j = 0; $j -lt $Colors.Count; $j++) {
            $HueDiff = [Math]::Abs($perfectColorHues[$i] - $Colors[$j].GetHue())
            if ($HueDiff -gt 180) {
                $HueDiff = 360 - $HueDiff
            }
            $SatDiff = [Math]::Abs($avgSaturation - $Colors[$j].GetSaturation())

            # Put very gray colors at a disadvantage because their hue hardly matters
            if ($Colors[$j].GetSaturation() -lt 0.08) {
                $SatDiff += 10
            }

            # Put very dark colors at a disadvantage because their hue hardly matters
            if ($Colors[$j].GetBrightness() -lt 0.10) {
                $SatDiff += 10
            }

            $diff = $HueDiff + $SatDiff * 45

            $diffMatrix[$i][$j] = $diff
        }
    }

    $alreadyUsedCOORD = New-Object int[][] -ArgumentList 2, 0

    # Decide on colors by traversing the 2D array for the lowest diffs, remembering which columns and rows we've already used
    do {
        $lowestDiffCOORD = @(-1, -1)
        $lowestDiff = 361

        for ($i = 0; $i -lt 6; $i++) {
            if ($i -notin $alreadyUsedCOORD[0] -and $NoDuplicates) {
                for ($j = 0; $j -lt $Colors.Count; $j++) {
                    if ($j -notin $alreadyUsedCOORD[1] -and $NoDuplicates) {
                        if ($diffMatrix[$i][$j] -lt $lowestDiff -or $lowestDiff -eq -1) {
                            $lowestDiff      = $diffMatrix[$i][$j]
                            $lowestDiffCOORD = @($i, $j)
                        }
                    }
                }
            }
        }

        if ($lowestDiff -lt 80) {
            # LowestDiffCOORD[0] is the index of one of the 6 colors we want to find ($i)
            # and LowestDiffCOORD[1] is the index of the Color we fill it with from $Colors input ($j)
            Write-Verbose "Decided on $($Colors[$lowestDiffCOORD[1]].Name) for color #$($lowestDiffCOORD[0]) with diff $lowestDiff"
            $bestFittingColors[$lowestDiffCOORD[0]] = $Colors[$lowestDiffCOORD[1]]
            $alreadyUsedCOORD[0] += $lowestDiffCOORD[0]
            $alreadyUsedCOORD[1] += $lowestDiffCOORD[1]
        }

    } while ($bestFittingColors.Contains($null) -and ( $alreadyUsedCOORD[1].Length -ne $Colors.Length ) -and $lowestDiff -lt 80)

    return $bestFittingColors
}

# This function is just for me to be able to define constants in a more C-like syntax because the PoSh syntax is weird
function CONST ($type, $varName, $value) {
    New-Variable -Name $varName -Value ($value -as "$type") -Option ReadOnly -Scope Script
}

# The entire code and logic to set the Windows Dekstop wallpaper is taken from the script:
# Set-RedditWallpapersAsDesktop.ps1 - JorgTheElder@Outlook.com, who himself says he built it with:
# great help from: https://www.kittell.net/code/powershell-remove-windows-wallpaper-per-user/
if ($SetWallpaper) {
    Add-Type -TypeDefinition '
        using System;
        using System.Runtime.InteropServices;
        using Microsoft.Win32;
        namespace Wallpaper {
            public class Setter {
                public const int SetDesktopWallpaper = 20;
                public const int UpdateIniFile       = 0x01;
                public const int SendWinIniChange    = 0x02;

                [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
                private static extern int SystemParametersInfo (int uAction, int uParam, string lpvParam, int fuWinIni);

                public static void SetWallpaper ( string path ) {
                    SystemParametersInfo( SetDesktopWallpaper, 0, path, UpdateIniFile | SendWinIniChange );
                }
            }
        }
    '

    function Set-WallPaper {
        Param (
            [Parameter(Mandatory)]
            [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
            [string]$WallPaper,
            [ValidateSet('Fill', 'Fit', 'Stretch', 'Center', 'Tile', 'Span')]
            [string]$WallpaperStyle = 'Fill',
            [switch]$AutoColorization
        )

        #Affected Reg Keys
        #AutoColorization REG_DWORD 1 or 0
        #WallpaperStyle REG_SZ 10, 6, 2, 0, 22
        #Wallpaper REG_SZ path-to-image

        #remove cached files to help change happen
        #Remove-Item -Path "$($env:APPDATA)\Microsoft\Windows\Themes\CachedFiles" -Recurse -Force -ErrorAction SilentlyContinue    

        $fit = @{ 'Fill' = 10; 'Fit' = 6; 'Stretch' = 2; 'Center' = 0; 'Tile' = '99'; 'Span' = '22' }

        if ($AutoColorization) {
            Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name AutoColorization -value 1;
        } else {
            Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name AutoColorization -value 0;
        }

        if ($WallpaperStyle -eq 'Tile') {
            Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name WallpaperStyle -value 0;
            Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name TileWallpaper -value 1;
        } else {
            Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name WallpaperStyle -value $fit[$WallpaperStyle];
            Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name TileWallpaper -value 0;
        }

        [Wallpaper.Setter]::SetWallpaper($WallPaper);
    }
}

function Test-SimilarColorInArray {
    Param (
        [Parameter(Mandatory = $true)]
        $Array,
        [Parameter(Mandatory = $true)]
        $Color,
        [hashtable]$CountPrevalence,
        [int[]]$CountHuePrevalence
    )

    $maxRGBdeviation = $SimilarColorThreshold * 255

    if ($similarColorAlgorithm -eq 'RGBTotalAverage') {
        foreach ($existingColor in $Array) {
            $TotalRedDeviation   = [System.Math]::Abs($Color.R - $existingColor.R)
            $TotalGreenDeviation = [System.Math]::Abs($Color.G - $existingColor.G)
            $TotalBlueDeviation  = [System.Math]::Abs($Color.B - $existingColor.B)
            $TotalRGBDeviation   = $TotalRedDeviation + $TotalGreenDeviation + $TotalBlueDeviation
            if (($TotalRGBDeviation / 3) -lt $maxRGBdeviation) {
                # A Color like this is already in hashtable
                if ($CountPrevalence) {
                    if ($CountPrevalence.Contains($existingColor.Name)) {
                        $CountPrevalence.Set_Item($existingColor.Name, $CountPrevalence.Get_Item($existingColor.Name) + 1)
                    } else {
                        $CountPrevalence.Add($existingColor.Name, 1)
                    }
                }
                return $true
            }
        }
    } elseif ($similarColorAlgorithm -eq 'RGBIndividual') {
        $similarColorsRange = @{
            'Rmin' = $Color.R - $maxRGBdeviation
            'Rmax' = $Color.R + $maxRGBdeviation
            'Gmin' = $Color.G - $maxRGBdeviation
            'Gmax' = $Color.G + $maxRGBdeviation
            'Bmin' = $Color.B - $maxRGBdeviation
            'Bmax' = $Color.B + $maxRGBdeviation
        }
    
        foreach ($existingColor in $Array) {
            if ($existingColor.R -ge $similarColorsRange.Get_Item('Rmin') -and $existingColor.R -le $similarColorsRange.Get_Item('Rmax') -and
                $existingColor.G -ge $similarColorsRange.Get_Item('Gmin') -and $existingColor.G -le $similarColorsRange.Get_Item('Gmax') -and
                $existingColor.B -ge $similarColorsRange.Get_Item('Bmin') -and $existingColor.B -le $similarColorsRange.Get_Item('Bmax')
            ) {
                # A Color like this is already in hashtable
                if ($CountPrevalence) {
                    if ($CountPrevalence.Contains($existingColor.Name)) {
                        $CountPrevalence.Set_Item($existingColor.Name, $CountPrevalence.Get_Item($existingColor.Name) + $colorCounts.Get_Item($Color.Name))
                    } else {
                        $CountPrevalence.Add($existingColor.Name, $colorCounts.Get_Item($Color.Name))
                    }
                }
                
                if ($CountHuePrevalence) {
                    $hueOfColor = [int]$existingColor.GetHue()
                    $hueCounts[$hueOfColor] += $colorCounts.Get_Item($color.Name)
                }
                return $true
            }
        }
    }

    # Color like this is not in hashtable yet
    return $false
}

function Write-ColorSample {
    Param (
        [Parameter( Mandatory = $true )]
        $Color,
        [ValidateSet('Stdout', 'Verbose')]
        [string]$Stream = 'Stdout'
    )
    
    switch ($Stream) {
        'Stdout' {
            Write-Output "$ESC[48;2;$($Color.R);$($Color.G);$($Color.B)m        $ESC[0m $($Color.Name)"
        }
        'Verbose' {
            Write-Verbose "$ESC[48;2;$($Color.R);$($Color.G);$($Color.B)m        $ESC[0m $($Color.Name)"
        }
    }
}

Write-Verbose "Minimum required RGB value deviation: $($SimilarColorThreshold * 255) with method $similarColorAlgorithm"
$null = [Reflection.Assembly]::LoadWithPartialName('System.Drawing')

if (-not $ImageFile) {
    # If not input image file was specified try to use the curent desktop wallpaper
    $unicodeEnc  = [system.Text.Encoding]::Unicode
    $encodedPath = [Microsoft.Win32.Registry]::GetValue('HKEY_CURRENT_USER\Control Panel\Desktop', 'TranscodedImageCache', -1) | Select-Object -Skip 24
    # The regex is needed because the registry value contains lots of trailing \0 terminators
    $ImageFile   = $unicodeEnc.GetString($encodedPath) -replace '\0*$'
    if (-not (Test-Path -LiteralPath $ImageFile -PathType Leaf)) {
        $ImageFile = (Get-ItemProperty -Path 'HKCU:\Control Panel\Desktop').Wallpaper
        if (-not (Test-Path -LiteralPath $ImageFile -PathType Leaf)) {
            throw "No image file was provided as input and could not get the current wallpaper, aborting."
        }
    }
}

$SrcImg = [System.Drawing.Image]::FromFile($ImageFile)

# resize image for faster processing
# target size is 400 pixels wide
if ($SrcImg.Width -gt 400) {
    [double]$ratio  = $SrcImg.Height / $SrcImg.Width
    [int]$newWidth  = 400
    [int]$newHeight = 400 * $ratio
    $SrcImg         = $SrcImg.GetThumbnailImage($newWidth, $newHeight, $null, 2)
}

[int]$img_totalPixels = $SrcImg.Width * $SrcImg.Height

# allColors will contain every unique color pixel we find in the image (deduplicated)

# Interestingly enough, declaring it in one of the following 3 ways will change the entire algorithm???
#$allColors = [System.Collections.Hashtable]::new($img_totalPixels)
#$allColors = [System.Collections.Hashtable]::new()
$allColors = @{}

# colorCounts will preserve information about how often a duplicate pixel is encountered so we can
# combine this information with $allColors later to work out heuristics on the image and its colors
#$colorCounts = [System.Collections.Generic.Dictionary[string, int]]::new($img_totalPixels)
$colorCounts = [System.Collections.Hashtable]::new()

# Define CONSTANT "ESC" with literal escape character
CONST char ESC 0x1b

# Store the chosen colors from the image after removing too similar ones and sort them by prevalence
$eligibleColors = [System.Collections.Generic.List[object]]::new()
$eligibleColors_Prevalences = @{}

# Initialize empty int-array with a size of 361 so we have all indexes from 0-360
$hueCounts = New-Object -TypeName int[] -ArgumentList 361

$perfectColorHues = @(
    240, # BLUE    0000FF
    120, # GREEN   00FF00
    180, # CYAN    00FFFF
    0,   # RED     FF0000
    300, # MAGENTA FF00FF
    60   # YELLOW  FFFF00
)

for ($horizontalPixel = 0; $horizontalPixel -lt $SrcImg.Width; $horizontalPixel++) {
    for ($verticalPixel = 0; $verticalPixel -lt $SrcImg.Height; $verticalPixel++) {
        $currentPixel = $SrcImg.GetPixel($horizontalPixel, $verticalPixel)

        if ($allColors.Contains($currentPixel.Name)) {
            $colorCounts.Set_Item($currentPixel.Name, $colorCounts.Get_Item($currentPixel.Name) + 1)
        } else {
            $allColors.Add($currentPixel.Name, $currentPixel)
            $colorCounts.Add($currentPixel.Name, 1)
        }
    }
}

Write-Verbose "Total pixels: $(($colorCounts.Values | Measure-Object -Sum).Sum)"

foreach ($color in ($allColors.GetEnumerator() | Sort-Object { $colorCounts.Get_Item($_.Key) } -Descending ) ) {
    if ($colorCounts.Get_Item($color.Name) -gt 1) {
        if (-not (Test-SimilarColorInArray -Array $eligibleColors -Color $color.Value -CountPrevalence $eligibleColors_Prevalences -CountHuePrevalence $hueCounts) ) {
            $eligibleColors.Add($color.Value)
        }
    }
}

Write-Verbose "$($eligibleColors.Count) unique enough colors found in the image."

# Weeding out otherwise unfit colors (too low contrast etc)
# Kick low saturation colors (all gray tones)
# Remove nearly pitch black colors
$eligibleColors.Where{ $_.GetBrightness() -lt 0.05 } |
    ForEach-Object { $null = $eligibleColors.Remove($_) }

Write-Verbose "$($eligibleColors.Count) colors left after removing too dark (pitch black) ones."

Write-Verbose 'Colors sorted by prevalence:'
foreach ($color in $eligibleColors) {
    $formattedPercent = "{0:00.00}" -f (($colorCounts.Get_Item($color.Name) + $eligibleColors_Prevalences.Get_Item($color.Name)) / $img_totalPixels * 100)
    #Write-Host "prevalence score for this color: $( $eligibleColors_Prevalences.Get_Item($color.Name) ) - total score: $( $colorCounts.Get_Item($color.Name) + $eligibleColors_Prevalences.Get_Item($color.Name) )"
    Write-Verbose "$ESC[48;2;$($Color.R);$($Color.G);$($Color.B)m        $ESC[0m $($Color.Name) - Hue makes up $formattedPercent% of total image"
}

# Choosing a background color
$backgroundColor = $eligibleColors | Where-Object {
    (($colorCounts.Get_Item($_.Name) + $eligibleColors_Prevalences.Get_Item($_.Name)) / $img_totalPixels * 100) -gt 0.02
} | Sort-Object {
    $colorHue            = [int]$_.GetHue()
    $rawColorPrevalence  = ($colorCounts.Get_Item($_.Name) + $eligibleColors_Prevalences.Get_Item($_.Name)) /  $img_totalPixels * 100
    $approxHuePrevalence = ($hueCounts[($colorHue - 4)..($colorHue + 4)] | Measure-Object -Sum).Sum
    [float]$HuePrevalenceInImage = $approxHuePrevalence / $img_totalPixels * 100
    if ($dbgOldBrightnessCurve) {
        [float]$BrightnessFactor = $_.GetBrightness() * 100
        if (-not $LightTheme) {
            [float]$BrightnessFactor = 100 - $BrightnessFactor
        }
    } else {
        if (-not $LightTheme) {
            $BrightnessFactor = [Math]::Pow(1 - $_.GetBrightness(), 2) * 100
        } else {
            $BrightnessFactor = [Math]::Pow($_.GetBrightness(), 2) * 100
        }
    }
    Write-Verbose ("Color $($_.Name) Background Score: {0} (Hue: {1}, Bri: {2})" -f ($HuePrevalenceInImage + $BrightnessFactor), $HuePrevalenceInImage, $BrightnessFactor)
    $HuePrevalenceInImage + $BrightnessFactor + $rawColorPrevalence
} | Select-Object -Last 1

# Removing the chosen background color from the mix
$null = $eligibleColors.Remove($backgroundColor)

# More detail about the chosen background color
$bgColorRoughHue = [int]$backgroundColor.GetHue()
Write-Verbose "Hue of background color makes up:     $("{0:N2}" -f ($hueCounts[$bgColorRoughHue] / $img_totalPixels * 100))% of total image."
Write-Verbose "Total pixels of this color in image:  $($colorCounts.Get_Item($backgroundColor.Name)) ($($colorCounts.Get_Item($backgroundColor.Name) / $img_totalPixels * 100 )%)"
Write-Verbose "Background colors brightness is:      $($backgroundColor.GetBrightness())"

# After having selected a background color, we remove all colors that are too close
# to it in brightness (required difference configurable) to ensure some minimum contrast
$eligibleColors.Where{
    $_.GetBrightness() -gt ($backgroundColor.GetBrightness() - $RequiredBrightnessDifference) -and $_.GetBrightness() -lt ($backgroundColor.GetBrightness() + $RequiredBrightnessDifference)
} | Foreach-Object { $null = $eligibleColors.Remove($_) }

# Choosing a primary text color
$primaryTextColor = $eligibleColors | Sort-Object {
    $colorHue = [int]$_.GetHue()
    $surroundingHuesPrevalence = ($hueCounts[($colorHue - 5)..($colorHue + 5)] | Measure-Object -Sum).Sum
    [float]$HuePrevalenceInImage = ( $surroundingHuesPrevalence / $img_totalPixels * 100)
    if ($dbgOldBrightnessCurve) {
        [float]$BrightnessFactor = $_.GetBrightness() * 100
        if ($LightTheme) {
            [float]$BrightnessFactor = 100 - $BrightnessFactor
        }
    } else {
        if ($LightTheme) {
            $BrightnessFactor = [Math]::Pow(1 - $_.GetBrightness(), 2) * 100
        } else {
            $BrightnessFactor = [Math]::Pow($_.GetBrightness(), 2) * 100
        }
    }
    Write-Verbose ("Color $($_.Name) text color Score: {0} (Hue: {1}, Bri: {2})" -f ($HuePrevalenceInImage + $BrightnessFactor), $HuePrevalenceInImage, $BrightnessFactor)
    #$HuePrevalenceInImage + $BrightnessFactor
    if ($HuePrevalenceInImage -lt 10) {
        10 + $BrightnessFactor
    } else {
        30 + $BrightnessFactor
    }
} | Select-Object -Last 1

# Removing the chosen foreground color from the mix
$null = $eligibleColors.Remove($primaryTextColor)

$color15white = $eligibleColors | Sort-Object {
    $colorHue = [int]$_.GetHue()
    $surroundingHuesPrevalence = ($hueCounts[($colorHue - 5)..($colorHue + 5)] | Measure-Object -Sum).Sum
    [float]$HuePrevalenceInImage = ( $surroundingHuesPrevalence / $img_totalPixels * 100)
    if ($dbgOldBrightnessCurve) {
        [float]$BrightnessFactor = $_.GetBrightness() * 100
        if ($LightTheme) {
            [float]$BrightnessFactor = 100 - $BrightnessFactor
        }
    } else {
        if ($LightTheme) {
            $BrightnessFactor = [Math]::Pow(1 - $_.GetBrightness(), 2) * 100
        } else {
            $BrightnessFactor = [Math]::Pow($_.GetBrightness(), 2) * 100
        }
    }
    Write-Verbose ("Color $($_.Name) text color Score: {0} (Hue: {1}, Bri: {2})" -f ($HuePrevalenceInImage + $BrightnessFactor), $HuePrevalenceInImage, $BrightnessFactor)
    #$HuePrevalenceInImage * 2 + $BrightnessFactor
    if ($HuePrevalenceInImage -lt 10) {
        10 + $BrightnessFactor
    } else {
        30 + $BrightnessFactor
    }
} | Select-Object -Last 1

# Removing the chosen foreground color from the mix
$null = $eligibleColors.Remove($color15white)

Write-Verbose "$($eligibleColors.Count) colors left after removing 1 background, 1 text color and colors with not enough brightness difference to the background."
Write-Verbose 'Remaining colors sorted by prevalence:'
$eligibleColors | ForEach-Object {
    Write-ColorSample -Color $_ -Stream Verbose
}

# This array will hold our 16 colors (16 objects of type System.Drawing.Color)
$themeColorPalette = @(
    $backgroundColor,  # Black
    $null,             # DarkBlue
    $null,             # DarkGreen
    $null,             # DarkCyan
    $null,             # DarkRed
    $null,             # DarkMagenta
    $null,             # DarkYellow
    $primaryTextColor, # Gray
    $null,             # DarkGray
    $null,             # Blue
    $null,             # Green
    $null,             # Cyan
    $null,             # Red
    $null,             # Magenta
    $null,             # Yellow
    $color15white      # White
)

if ($MapColorsByHue) {
    $colorGroupsBySatandHue = $eligibleColors | Group-Object -Property { $_.GetBrightness() -gt 0.4 } | Sort-Object Count -Descending
    $colorGroupsBySatandHue | Select-Object Count, Name | Out-String | Write-Verbose
    $colorGroupsBySatandHue | ForEach-Object {
        $colorSamples = foreach ($color in $_.group) {
            "$ESC[48;2;$($Color.R);$($Color.G);$($Color.B)m        $ESC[0m"
        }
        Write-Verbose ($colorSamples -join " ")
    }
    # The function returns the colors it matched best
    # from $eligibleColors in the order that they must go
    # into the color palette
    $brightColorsPalette = Get-BestColorMatch -Colors $colorGroupsBySatandHue.Where{ $_.Name -eq 'true' }.Group -NoDuplicates
    $darkColorsPalette   = Get-BestColorMatch -Colors $colorGroupsBySatandHue.Where{ $_.Name -eq 'false' }.Group -NoDuplicates

    # If no bright color at all was returned, take the dark colors and brighten them to create lighter variants
    if (-not $brightColorsPalette -ne $nulll -and $darkColorsPalette -eq $null) {
        for ($i = 0; $i -lt 6; $i++) {
            if ($null -ne $darkColorsPalette[$i]) {
                $brightColorsPalette[$i] = Convert-HSLtoColor -Hue $darkColorsPalette[$i].GetHue() -Brightness ($darkColorsPalette[$i].GetBrightness() + 0.24) -Saturation $darkColorsPalette[$i].GetSaturation()
            }
        }
    }

    # If no dark colors at all were returned, take the light colors and darken them
    if (-not $darkColorsPalette -eq $null -and $brightColorsPalette -eq $null) {
        for ($i = 0; $i -lt 6; $i++) {
            if ($null -ne $brightColorsPalette[$i]) {
                $darkColorsPalette[$i] = Convert-HSLtoColor -Hue $brightColorsPalette[$i].GetHue() -Brightness ($brightColorsPalette[$i].GetBrightness() + 0.24) -Saturation $brightColorsPalette[$i].GetSaturation()
            }
        }
    }

    for ($i = 0; $i -lt 6; $i++) {
        $themeColorPalette[$i + 9] = $brightColorsPalette[$i]
    }

    $themeColorPalette[1] = $darkColorsPalette[0]
    $themeColorPalette[2] = $darkColorsPalette[1]
    $themeColorPalette[3] = $darkColorsPalette[2]
    $themeColorPalette[4] = $darkColorsPalette[3]
    $themeColorPalette[5] = $darkColorsPalette[4]
    $themeColorPalette[6] = $darkColorsPalette[5]

    $brightColorsPalette | Sort-Object -Descending | ForEach-Object {
        $null = $eligibleColors.Remove($_)
    }
    $darkColorsPalette | Sort-Object -Descending | ForEach-Object {
        $null = $eligibleColors.Remove($_)
    }
} else {
    $themeColorPalette[1] = $eligibleColors[0]
    $themeColorPalette[2] = $eligibleColors[1]
    $themeColorPalette[3] = $eligibleColors[2]
    $themeColorPalette[4] = $eligibleColors[3]
    $themeColorPalette[5] = $eligibleColors[4]
    $themeColorPalette[6] = $eligibleColors[5]

    $themeColorPalette[9]  = $eligibleColors[7]
    $themeColorPalette[10] = $eligibleColors[8]
    $themeColorPalette[11] = $eligibleColors[9]
    $themeColorPalette[12] = $eligibleColors[10]
    $themeColorPalette[13] = $eligibleColors[11]
    $themeColorPalette[14] = $eligibleColors[12]
}
$themeColorPalette[8] = $eligibleColors[6]

# Handling too few colors found:
# if not enough colors are left to fill in all 16 required for a console theme:
if ($themeColorPalette.Contains($null)) {
    switch ($NotEnoughColors) {
        'Reuse' {
            $i = 0
            while ($themeColorPalette.Contains($null)) {
                # Append as many of the colors already in the list to the list again as needed to get to 14
                $index = $themeColorPalette.IndexOf($null)

                Write-Verbose "Warning: Had to reuse a color at position $index -> $($eligibleColors[$i].Name)"
                $themeColorPalette[$index] = $eligibleColors[$i]
                if (++$i -eq $eligibleColors.Count) {
                    $i = 0
                }
            }
        }
        'Extrapolate' {
            while ($themeColorPalette.Contains($null)) {
                $index = $themeColorPalette.IndexOf($null)
                if ($index -in 1, 2, 3, 4, 5, 6, 9, 10, 11, 12, 13, 14) {
                    $highestPrevalence  = 0
                    $mostPrevalentIndex = 0
                    if ($index -gt 8) {
                        # We have to fill in a bright and "colorful" color
                        $starting = 9
                        $until    = 15
                    } else {
                        # We have to fill in a bright and "colorful" color
                        $starting = 1
                        $until    = 7
                    }
                    for ($i = $starting; $i -lt $until; $i++) {
                        # Find the most prevalent dark or light color to sample our new one off of
                        if ($null -ne $themeColorPalette[$i]) {
                            $colorHue = [int]$themeColorPalette[$i].GetHue()
                            if (($hueCounts[($colorHue - 5)..($colorHue + 5)] | Measure-Object -Sum).Sum -gt $highestPrevalence) {
                                $mostPrevalentIndex = $i
                            }
                        }
                    }

                    # Change hue of most prevalent color to that of the color we are looking for then fill it in!
                    if ($index -gt 8) {
                        $themeColorPalette[$index] = Convert-HSLtoColor -Hue $perfectColorHues[$index - 9] -Brightness $themeColorPalette[$mostPrevalentIndex].GetBrightness() -Saturation $themeColorPalette[$mostPrevalentIndex].GetSaturation()
                    } else {
                        $themeColorPalette[$index] = Convert-HSLtoColor -Hue $perfectColorHues[$index - 1] -Brightness $themeColorPalette[$mostPrevalentIndex].GetBrightness() -Saturation $themeColorPalette[$mostPrevalentIndex].GetSaturation()
                    }
                } elseif ($index -eq 8) {
                    # We have to fill in dark gray
                    if ($LightTheme) {
                        $themeColorPalette[$index] = Convert-HSLtoColor -Hue $primaryTextColor.GetHue() -Brightness ($primaryTextColor.GetBrightness() + 0.24) -Saturation $primaryTextColor.GetSaturation()
                    } else {
                        $themeColorPalette[$index] = Convert-HSLtoColor -Hue $backgroundColor.GetHue() -Brightness ($backgroundColor.GetBrightness() + 0.24) -Saturation $backgroundColor.GetSaturation()
                    }
                } else {
                    # Index 15 missing -> white
                    $themeColorPalette[$index] = $primaryTextColor
                }
                Write-Verbose "Warning: Had to extrapolate a color at position $index -> $($themeColorPalette[$mostPrevalentIndex].Name) to $($themeColorPalette[$index].Name)"
            }
        }
        'Auto' {
            # The idea behind auto is that it extrapolates if few colots are found (to not have a theme of only 2-3 colors),
            # and reuses if "enough" colors are found to create an interesting theme, like e.g. about 6+
            # With L O T S of other improvements done in the meantime though, this shouldn't be necessary anymore
        }
    }
}

if ($Apply) {
    Add-Type -TypeDefinition '
        using System;
        using System.Runtime.InteropServices;

        public class ConsoleAPI
        {
            [StructLayout(LayoutKind.Sequential)]
            public struct COORD
            {
                public short X;
                public short Y;
            }

            public struct SMALL_RECT
            {
                public short Left;
                public short Top;
                public short Right;
                public short Bottom;
            }

            [StructLayout(LayoutKind.Sequential)]
            public struct CONSOLE_SCREEN_BUFFER_INFO_EX
            {
                public uint cbSize;
                public COORD dwSize;
                public COORD dwCursorPosition;
                public ushort wAttributes;
                public SMALL_RECT srWindow;
                public COORD dwMaximumWindowSize;

                public ushort wPopupAttributes;
                public bool bFullscreenSupported;

                [MarshalAs(UnmanagedType.ByValArray, SizeConst = 16)]
                public uint[] ColorTable;
            }

            [DllImport("kernel32.dll", SetLastError = true)]
            public static extern IntPtr GetStdHandle(int nStdHandle);

            [DllImport("kernel32.dll", SetLastError = true)]
            public static extern bool GetConsoleScreenBufferInfoEx(IntPtr hConsoleOutput, ref CONSOLE_SCREEN_BUFFER_INFO_EX ConsoleScreenBufferInfo);

            [DllImport("kernel32.dll", SetLastError = true)]
            public static extern bool SetConsoleScreenBufferInfoEx(IntPtr ConsoleOutput, ref CONSOLE_SCREEN_BUFFER_INFO_EX ConsoleScreenBufferInfoEx);
        }
    '

    function RGB ($r, $g, $b) {
		return [uint32]$r + ([uint32]$g -shl 8) + ([uint32]$b -shl 16);
    }

    # https://docs.microsoft.com/en-us/windows/console/getstdhandle
    $conoutHandle = [ConsoleAPI]::GetStdHandle( -11 )
    $csbiex = New-Object ConsoleAPI+CONSOLE_SCREEN_BUFFER_INFO_EX -Property @{'cbSize' = 96}

    # Delete registry subkeys which could contain custom themes
    Remove-Item -Path 'HKCU:\Console\*'
    # Make sure ColorTable07 is selected as the primary text color
    [Microsoft.Win32.Registry]::SetValue("HKEY_CURRENT_USER\Console", "ScreenColors", 7, "DWORD")
    # Set the CursorColor
    [Microsoft.Win32.Registry]::SetValue("HKEY_CURRENT_USER\Console", "CursorColor", (RGB $primaryTextColor.R $primaryTextColor.G $primaryTextColor.B), "DWORD")

    $getCSBIEXresult = [ConsoleAPI]::GetConsoleScreenBufferInfoEx($conoutHandle, [ref]$csbiex)

    # Apply theme to new console windows
    for ($i = 0; $i -lt 16; $i++) {
        $COLORREFval = RGB $themeColorPalette[$i].R $themeColorPalette[$i].G $themeColorPalette[$i].B
        [Microsoft.Win32.Registry]::SetValue("HKEY_CURRENT_USER\Console", ("ColorTable{0:00}" -f $i), $COLORREFval, "DWORD")
        $csbiex.ColorTable[$i] = $COLORREFval
        Write-ColorSample -Color $themeColorPalette[$i] -Stream Stdout
    }

    if ($getCSBIEXresult) {
        # Apply theme to current console window and restore original size
        # because SetConsoleScreenBufferInfoEx() inexplicably shrinks the window
        $windowSize = $host.UI.Rawui.WindowSize
        $null = [ConsoleAPI]::SetConsoleScreenBufferInfoEx($conoutHandle, [ref]$csbiex)
        [console]::SetWindowSize($windowSize.Width, $windowSize.Height)
    }
    
    # Get visible console windows besides the current one (where this is running)
    [array]$openConsoleWindows = (Get-CimInstance -Class Win32_Process -Filter "Name = 'conhost.exe'").ParentProcessID | Where-Object { $_ -ne $PID -and (Get-Process -Id $_).MainWindowHandle -ne 0 }

    # If there's open console windows besides the current one
    if ($openConsoleWindows) {
        Write-Verbose "$($openConsoleWindows.Count) open console windows besides this one were found"
        # Apply theme to all other open console windows
        if (Test-Path -Path "$PSScriptRoot\cpp_consoleattacher.exe") {
            $cppHelper = (Get-Item -LiteralPath cpp_consoleattacher.exe).FullName
            $startup   = [wmiclass]"Win32_ProcessStartup"
            $startup.Properties['CreateFlags'].value = 0x8

            foreach ($CONPID in $openConsoleWindows) {
                $null = ([wmiclass]"Win32_Process").Create("$cppHelper $CONPID $($csbiex.ColorTable)", 'C:\', $startup)
            }
        } else {
            Write-Warning -Message "The complementary application 'cpp_consoleattacher.exe' could not be found in the directory of this script.`r`nWithout it, currently open console windows other than this one cannot be re-colored on the fly."
        }
    }

    # Experimental support for re-coloring of the new Windows Terminal app (wt.exe)
    if ($wt = Get-AppxPackage -Name 'Microsoft.WindowsTerminal') {
        $newThemeforWT = [PSCustomObject]@{
            background   = "#{0}" -f $themeColorPalette[0].Name.Substring(2)
            black        = "#{0}" -f $themeColorPalette[0].Name.Substring(2)
            blue         = "#{0}" -f $themeColorPalette[1].Name.Substring(2)
            brightBlack  = "#{0}" -f $themeColorPalette[7].Name.Substring(2)
            brightBlue   = "#{0}" -f $themeColorPalette[9].Name.Substring(2)
            brightCyan   = "#{0}" -f $themeColorPalette[11].Name.Substring(2)
            brightGreen  = "#{0}" -f $themeColorPalette[10].Name.Substring(2)
            brightPurple = "#{0}" -f $themeColorPalette[13].Name.Substring(2)
            brightRed    = "#{0}" -f $themeColorPalette[12].Name.Substring(2)
            brightWhite  = "#{0}" -f $themeColorPalette[15].Name.Substring(2)
            brightYellow = "#{0}" -f $themeColorPalette[14].Name.Substring(2)
            cyan         = "#{0}" -f $themeColorPalette[3].Name.Substring(2)
            foreground   = "#{0}" -f $themeColorPalette[7].Name.Substring(2)
            green        = "#{0}" -f $themeColorPalette[2].Name.Substring(2)
            name         = 'Generated by poshwal'
            purple       = "#{0}" -f $themeColorPalette[5].Name.Substring(2)
            red          = "#{0}" -f $themeColorPalette[4].Name.Substring(2)
            white        = "#{0}" -f $themeColorPalette[15].Name.Substring(2)
            yellow       = "#{0}" -f $themeColorPalette[6].Name.Substring(2)
        }
        
        [string]$wtCfgFile = "$env:LOCALAPPDATA\Packages\$($wt.PackageFamilyName)\RoamingState\profiles.json"
        if (Test-Path -LiteralPath $wtCfgFile -PathType Leaf) {
            $wtSettings = Get-Content -LiteralPath $wtCfgFile | ConvertFrom-Json
            if ($wtSettings.schemes.name -contains 'Generated by poshwal') {
                $wtSettings.schemes = $wtSettings.schemes.Where{ $_.name -ne 'Generated by poshwal' }
            }
        
            $wtSettings.schemes += $newThemeforWT
        
            foreach ($a in $wtSettings.profiles) {
                $a.colorScheme = 'Generated by poshwal'
            }
        
            $wtSettings | ConvertTo-Json -Depth 100 | Set-Content -Path $wtCfgFile
        } else {
            Write-Warning "Configuration file for the new Windows Terminal not found, skipping. You are probably running a version of wt that is not supported yet."
        }
    }

    # Remove theme settings from existing Shortcuts to PowerShell so the new theme applies everywhere including Win + X menu and the file explorer shortcuts
    $wscriptCom = New-Object -ComObject WScript.Shell
    $shortcuts  = Get-ChildItem -Path ([Environment]::GetFolderPath('Programs') + '\Windows PowerShell\*.lnk') -Exclude "*ISE*"
    foreach ($shortcut in $shortcuts) {
        $OldShortcut = $wscriptCom.CreateShortcut($shortcut.FullName)
        if (Test-Path -Path ($shortcut.FullName + '.bak')) {
            Remove-Item -Path $shortcut.FullName
        } else {
            Rename-Item -Path $shortcut.FullName -NewName "$shortcut.bak"
        }
        $NewShortcut = $wscriptCom.CreateShortcut($shortcut.FullName)
        $NewShortcut.Description      = $OldShortcut.Description
        $NewShortcut.TargetPath       = $OldShortcut.TargetPath
        $NewShortcut.WorkingDirectory = $OldShortcut.WorkingDirectory
        $NewShortcut.Save()
    }
} else {
    foreach ($currentColor in $themeColorPalette) {
        Write-ColorSample -Color $currentColor -Stream Stdout
    }
    Write-Information -Message "INFORMATION: To apply the generated theme specify the '-Apply' parameter.`r`nLearn more about available parameters by utilizing tab-completion or by running 'Get-Help' on this script."
}

if ($SetWallpaper) {
    Set-Wallpaper -Wallpaper $ImageFile -WallpaperStyle $SetWallpaper -AutoColorization
}
