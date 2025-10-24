
param(
    [Parameter(Mandatory=$true, Position=0)][string]$Input1,
    [Parameter(Mandatory=$true, Position=1)][string]$Input2,
    [Parameter(Mandatory=$true, Position=2)][string]$Input3,
    [Parameter(Mandatory=$true, Position=3)][string]$Input4,
    [Parameter(Mandatory=$true, Position=4)][string]$Output
)

$capSide = 16 * [math]::Floor([math]::Sqrt(139264))
$tile    = [math]::Floor($capSide / 2)
$tilePad = "scale=w=${tile}:h=${tile}:force_original_aspect_ratio=decrease,scale=trunc(iw/2)*2:trunc(ih/2)*2,pad=${tile}:${tile}:(ow-iw)/2:(oh-ih)/2"

$fc = @"
[0:v]$tilePad[v0];
[1:v]$tilePad[v1];
[2:v]$tilePad[v2];
[3:v]$tilePad[v3];
[v0][v1][v2][v3]xstack=inputs=4:layout=0_0|${tile}_0|0_${tile}|${tile}_${tile}[out]
"@

ffmpeg -hide_banner -loglevel error -stats -y -i "$Input1" -i "$Input2" -i "$Input3" -i "$Input4" -filter_complex $fc -map "[out]" -c:v libx264 -x264-params "level=6.2" -pix_fmt yuv420p "$Output"