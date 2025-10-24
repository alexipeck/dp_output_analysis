$framerate = 10
$vfMaxH264 = 'scale=w=min(iw\,16*floor(sqrt(139264*iw/ih))):h=min(ih\,16*floor(sqrt(139264*ih/iw))):force_original_aspect_ratio=decrease,scale=trunc(iw/2)*2:trunc(ih/2)*2,setsar=1'
if (-not (Test-Path "./videos")) {
    New-Item -ItemType Directory -Path "./videos" -Force | Out-Null
}

ffmpeg -hide_banner -loglevel error -stats -y -framerate $framerate -i "./infection/%d.png" -vf $vfMaxH264 -c:v libx264 -x264-params "level=6.2" -pix_fmt yuv420p "./videos/infection.mp4"
ffmpeg -hide_banner -loglevel error -stats -y -framerate $framerate -i "./foi/%d.png" -vf $vfMaxH264 -c:v libx264 -x264-params "level=6.2" -pix_fmt yuv420p "./videos/foi.mp4"
