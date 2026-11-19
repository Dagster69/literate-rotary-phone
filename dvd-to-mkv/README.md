# DVD to MKV

Converts DVD folders (`VIDEO_TS`) into single MKV files for Jellyfin, with a window to pick movies and watch progress. It uses FFmpeg to copy the video, audio, subtitles and chapters without re-encoding, so there is no quality loss. Original DVD folders are never touched, moved or deleted.

## One-time setup (the PC that will do the converting)

1. Install FFmpeg: open Command Prompt and run `winget install Gyan.FFmpeg`. It needs version 7.1 or newer.
2. Copy `DvdToMkv.ps1` and `Run-DvdToMkv.bat` into the same folder.

## Using it

1. Double-click **Run-DvdToMkv.bat**. If Windows warns you, click **More info > Run anyway**.
2. Check the two folders at the top:
   - **Movies folder**: where the DVD folders are, e.g. `\\NASDE4EB7\Multimedia\Danes vids`
   - **Save MKV files to**: where the new files go. The genre folders are copied, e.g. `...\Action\BLOODSHOT\BLOODSHOT.mkv`
3. Press **Scan**. It lists every DVD it finds (a few minutes over the network).
4. **Do a test first.** Press **Untick all**, then tick 3 or 4 movies and press **Start**. Play the results in VLC and check the length, picture, sound and subtitles.
5. When happy, press **Tick all** and **Start**. It runs one movie at a time, shows progress and time left, and can be stopped at any time with **Stop**.

## Working through a big library in rounds

If the output drive can't hold everything at once, convert a few genres at a time:

1. Pick a genre in the drop-down (it shows how many movies and how many GB) and press **Tick genre**. Repeat for more genres. The line under the list shows the total ticked and the free space, and turns red if it won't fit.
2. Press **Start**.
3. When it finishes, review the results (below) and watch a sample of files.
4. Only then delete that round's original DVD folders in Explorer, **and empty the NAS recycle bin** (`@Recycle`), otherwise the space isn't freed.
5. Press **Scan** again and do the next round.

The converter never deletes anything itself.

## Checking the results

- Each finished movie is compared with the length the DVD itself reports. If they differ by more than 90 seconds or 3%, the status shows an orange **Done - CHECK** and the file is kept, so you can look at it.
- `DvdToMkv_results.csv` in Documents has one row per movie: result, DVD and MKV sizes and lengths, number of audio and subtitle tracks, and the paths. Open it in Excel and filter the Result column for CHECK or FAILED, and sort by SubtitleTracks to spot any that came out with none.

## Deleting the originals (only after you have checked the results)

`Run-DeleteConvertedOriginals.bat` reads `DvdToMkv_results.csv` and lists the original DVD folders whose movie converted OK and whose MKV exists and is a sensible size. Nothing is deleted until you type `DELETE`. CHECK and FAILED movies are never touched. Only the `VIDEO_TS` folder is removed; the movie folder is removed only if it is then empty, so moviesheets and posters stay. The plan is saved as `DvdToMkv_delete_plan.csv` so you can look at it in Excel first. Afterwards, empty the NAS recycle bin or the space is not freed.

## Retrying failures

Press **Scan**, then **Tick retries**. It ticks every movie whose latest result in `DvdToMkv_results.csv` was FAILED or CHECK and that has no MKV now. To redo a movie that finished as CHECK, delete its MKV and folder first.

## Two ways of reading a DVD

1. **Normal method** (fast, about 1 to 5 minutes): FFmpeg reads the DVD through its own navigation. This keeps chapters and the language names of the audio and subtitle tracks.
2. **Direct VOB method** (slower, about 15 to 20 minutes): used automatically when the normal method fails, gives an empty file, or comes out shorter than the DVD says. Some DVD Cloner copies have broken index files that stop the normal method. This method joins the `VTS_xx_n.VOB` files of the main title and ignores the index files. The movie, audio and subtitles are kept, but chapters and language names are lost. Audio tracks are put in DVD order, with the first as the default, and uncompressed (LPCM) audio is converted to lossless PCM so it fits in MKV.

The results spreadsheet has a **Method** column showing which one each movie used.

## How it behaves

- **Safe to stop and restart.** Each file is written as `.partial` and only renamed once it passes checks (sensible length, has audio, size close to the DVD). Finished movies show as Done and are skipped next time.
- **Failed movies** show in red with the reason. Tick them and press Start to retry.
- **Skipped automatically:** folders with no `VIDEO_TS.IFO` or under 500 MB (empty or broken copies).
- **Free space check** before each movie. If the output drive is nearly full it stops cleanly.
- **Subtitle timing fix:** subtitle packets stamped later than 6 hours are dropped. This fixes DVD Cloner copies that showed a 13-hour length.
- **Log** is saved in your Documents folder as `DvdToMkv_<date>.log`.

## Without a window

```powershell
.\DvdToMkv.ps1 -NoGui -Source "\\NASDE4EB7\Multimedia\Danes vids" -Output "F:\MKV-test" -Limit 3
```

Converts the first 3 movies still waiting, printing progress to the console.
