# poshwal


Automatically generates and applies themes for the built-in Windows terminal (`conhost.exe`) based on the colors found in an image.

Although the name is a reference to [pywal](https://github.com/dylanaraps/pywal) and this script is supposed to provide a similar feature for the Windows platform it borrows no code or algorithms from pywal - I wanted to learn and create something of my own (this is also a good excuse for why it's not perfect I guess).

## Syntax

```powershell
./poshwal.ps1
    [[-ImageFile] <String>]
    [-Apply]
    [-LightTheme]
    [-MapColorsByHue]
    [-SetWallpaper {Fill | Fit | Stretch | Center | Tile | Span}]
    [-SimilarColorThreshold <float>]
    [-SimilarColorAlgorithm { RGBTotalAverage | RGBIndividual }]
    [-RequiredBrightnessDifference <float>]
    [-NotEnoughColors { Extrapolate | Reuse }]
    [<CommonParameters>]
```

## Usage

Download the latest release from [here](https://github.com/jantari/poshwal/releases/latest) and run:

Simplest example:
```powershell
./poshwal.ps1 -Apply
```

The script also accepts an image file over the pipeline:
```powershell
ls "$env:USERPROFILE\Pictures" -Filter "*.jpg" | Get-Random | ./poshwal.ps1 -Apply -SetWallpaper 'Fill'
```

If you wish to see more details on how a theme was decided on to tweak parameters for future runs, looking at the `-Verbose` output will be necessary:
```powershell
./poshwal.ps1 -ImageFile "\\nas\Share\example.png" -Verbose
```

By default the script will try to match colors from the image to their most appropriate spot in the theme
so that output from commands like `Write-Host "Success" -ForegroundColor Green` is still somewhat green
and even generate colors not actually in the image if nothing close at all was found. This is a safe default
because it ensures any console apps you have that rely on colored output to convey information will work.

However, if you do not want colors generated but strictly use what's in the image you can run:
```powershell
./poshwal.ps1 -ImageFile "\\nas\Share\example.png" -NotEnoughColors 'Reuse'
```
which will reuse colors instead where no good enough fit was found, or:
```powershell
./poshwal.ps1 -ImageFile "\\nas\Share\example.png" -MapColorsByHue:$false
```
which disables any kind of "smart" assignment of colors and instead assigns them in descending order of
their prevalence in the image. Since no colors are being discarded here for being a "bad fit" it maximizes
your chances of enough colors having been extracted from the image to fill in all 16 slots. However if
there weren't the behavior specified by `-NotEnoughColors` will still apply which is "Extrapolate" by default.
To combat this without reusing colors, you can set the `-SimilarColorThreshold` to a lower value to allow
more colors to be extracted from the image despite being increasingly similar and hard to differentiate.

## Sample screenshots

I have uploaded some samples here: https://imgur.com/a/emb95A3

## Building

The `.ps1` script can be run autonomously, is easily editable and does not require compilation.
The helper binary 'cpp_consoleattacher.exe' is only needed to re-color already opened terminal windows other than the one you're running `poshwal` from on the fly.

To compile the C++ binary yourself, download and install the [Microsoft Visual C++ Build Tools](https://aka.ms/buildtools) or the full Visual Studio 2017.
Open a "x64 Native Tools Command prompt for VS 2017" from the Start menu, `cd` to the directory with the source file and run `cl cpp_consoleattacher.cpp /EHsc` to compile it.
For more details please refer to the [Microsoft documentation](https://docs.microsoft.com/en-us/cpp/build/walkthrough-compiling-a-native-cpp-program-on-the-command-line?view=vs-2017) on the VC++ Build tools cli.

## Miscellaneous

1. If you do not specify an input image, the script will try to use your current wallpaper.
2. If you are not happy with a generated theme from a particular image try different values for the `SimilarColorThreshold`, `RequiredBrightnessDifference` and `SimilarColorAlgorithm` parameters as well as possibly `-MapColorsByHue:$false` (consult `Get-Help`) - that said the algorithm is not magic and won't be able to pull a theme you love out of every image ☹ please try a few!
3. The 'cpp_consoleattacher.exe' file is not neccessary for the script to function, it is only needed to re-color other already opened terminals on the fly.