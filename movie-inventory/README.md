# Movie inventory

Makes a spreadsheet of every movie in the NAS movie share, so you can see what needs sorting before setting up Jellyfin. It only reads; nothing on the NAS is changed, moved or deleted.

## Running it (Windows PC)

1. Copy `MovieInventory.ps1` and `Run-MovieInventory.bat` into the same folder on the PC, e.g. `Documents\MovieInventory`.
2. Double-click `Run-MovieInventory.bat`.
3. When asked, type the movies folder, e.g. `\\NAS\Movies`, or a mapped drive like `M:\`. You can also drag the folder from File Explorer onto the window. Use the same path you'd type into File Explorer's address bar to open the movies.
4. Wait. Over the network a big library can take a few minutes.
5. The spreadsheet opens in Excel when it's done. Both files are saved to your Documents folder:
   - `MovieInventory_<date>.csv` has one row per movie.
   - `MovieInventory_<date>_summary.txt` has totals by format and genre, plus any folders it couldn't read.

If Windows shows "Windows protected your PC", click **More info → Run anyway**. This is normal for scripts downloaded from the internet.

## What's in the spreadsheet

| Column | Meaning |
|---|---|
| Genre | The first folder under the movies share, e.g. `Action` |
| Title, Year | Worked out from the folder or file name. Check any that look wrong. |
| Format | `DVD (VIDEO_TS)`, `MKV`, `MP4`, `AVI`, `ISO image`, `Blu-ray (BDMV)`, `VOB (loose)`, ... |
| Files, SizeGB | Number of files and total size. A DVD counts all its VOB/IFO/BUP files. |
| HasMoviesheet | A `*sheet*.jpg` image is next to the movie |
| HasPoster | Another image (e.g. `folder.jpg`, `Movie.jpg`) is next to the movie |
| HasNfo | A `.nfo` info file is next to the movie |
| Notes | Things to check: possible duplicates, multi-disc sets, multi-part files, small files that may be extras or samples, incomplete DVDs, older formats |
| Path | Where it is on the NAS |

Use Excel's **Data → Filter** to filter by Format or Notes.

## What it skips

QNAP system folders (`@Recycle`, `@Recently-Snapshot`, `.@__thumb`, `.streams`, anything starting with `@`) and the Windows recycle bin, so deleted files and snapshots aren't counted twice.

## Options

Run from PowerShell for more control:

```powershell
.\MovieInventory.ps1 -Root "\\NAS\Movies" -OutDir "D:\Reports" -NoOpen
```
