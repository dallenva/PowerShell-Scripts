# Combine multiple media files into one file

#scoop install ffmpeg
$path = "C:\Users\dalle\Documents\Audiobooks\Star Wars - The Radio Drama"
#$path = "C:\Users\dalle\Documents\Audiobooks\The Empire Strikes Back - The Radio Drama"
#$path = "C:\Users\dalle\Documents\Audiobooks\Return Of The Jedi - The Radio Drama"

$filename = $path.split("\")[-1] + ".mp4"

Set-Location $path

$rawlist = get-childitem 

remove-item filelist.txt
remove-item out.mp4
$rawlist | foreach-object {
            add-content -path $path/filelist.txt -value("file '"  + $_.name + "'")
            }
get-content -path .\filelist.txt

ffmpeg -f concat -safe 0 -i filelist.txt -c copy out.mp4
rename-item out.mp4 $filename