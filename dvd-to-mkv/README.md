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
