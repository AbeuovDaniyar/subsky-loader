# Sub Sky Loader

Turns the daily **GT Trips** export (.xlsx or .csv) into the Sub Sky loader workbook — the same 26-column table the old "NH VA Loader" Excel sheet produced, without Power Query or changing file paths.

## Download

Get **SubSkyLoader.exe** from the [latest release](../../releases/latest). No install needed.

The first time, Windows may show "Windows protected your PC" because the app isn't signed. Click **More info → Run anyway**.

## Use

1. Double-click `SubSkyLoader.exe`.
2. Choose today's GT Trips file (.xlsx or .csv).
3. A **Save As** window opens with a suggested name (e.g. `Sub Sky 4th.xlsx`) next to your trips file. Change the folder or the file name if you want, then click **Save**.
   - The sheet inside is named after the file you save.
   - When you picked an .xlsx, a `.csv` copy is saved in the same folder.
   - The app remembers the last folder you saved to.
4. Open the loader, or convert another file.

If the file you're saving over is open in Excel, it tells you to close it or pick another name.

## What it does

- Reads either the leg export (`Leg ID/Leg #`, `Passenger Name`, `Requested PickUp Time`, …) or the VA manifest (`Passenger`, `Estimated Pickup Time`, …) format. Header matching ignores case, spaces and punctuation.
- Fills **Booking Id** from `Leg ID/Leg #` (or a `Trip ID` column) as `1-<trip #>-<leg>`, e.g. `896699 / Leg-A` → `1-896699-A`. If there's no ID column it numbers the trips 1, 2, 3… and says so.
- Removes commas from passenger names, turns times like `12:00 EST` or `2:30 PM` into `12:00` / `14:30`.
- Fills the fixed fields (Hartford County polygon, NH, SUB, CLI1, TL1, status 10, zero costs).
- Writes a formatted Excel table (TableStyleMedium7) on a sheet named after the file, e.g. `Sub Sky 4th`.

No data leaves the computer.

## Build from source

```
python -m pip install -r requirements.txt
pyinstaller --onefile --windowed --name SubSkyLoader subsky_loader.py
```

The .exe is written to `dist\`.
