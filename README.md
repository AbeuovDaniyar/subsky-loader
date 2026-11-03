# Sub Sky Loader

Turns the daily **GT Trips** export (.xlsx or .csv) into the Sub Sky loader workbook — the same 26-column table the old "NH VA Loader" Excel sheet produced, without Power Query or changing file paths.

## Download

Get **SubSkyLoader.exe** from the [latest release](../../releases/latest). No install needed.

The first time, Windows may show "Windows protected your PC" because the app isn't signed. Click **More info → Run anyway**.

## Use

1. Double-click `SubSkyLoader.exe`.
2. Choose today's GT Trips file (.xlsx or .csv).
3. It saves, next to that file:
   - a `.csv` copy (when you picked an .xlsx)
   - the loader, e.g. `Sub Sky 4th.xlsx` (named from the file name)
4. Open the loader, or convert another file.

If a loader with that name is already open in Excel, it saves as `Sub Sky 4th (2).xlsx` instead.

## What it does

- Reads either the leg export (`Leg ID/Leg #`, `Passenger Name`, `Requested PickUp Time`, …) or the VA manifest (`Passenger`, `Estimated Pickup Time`, …) format. Header matching ignores case, spaces and punctuation.
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
