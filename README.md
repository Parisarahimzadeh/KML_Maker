

# Excel → KML (short)

* Converts each Excel sheet to a KML file.
* Preserves cell background colors (RGB, Indexed, Theme+Tint).
* Placemark description shows a vertical table of all columns.
* `DotColor` sets marker color (`#RRGGBB`, `R,G,B`, CSS names).
* Only headers starting with `BAPath` or `BHPath` are forced to black when empty/white.
* Simple circle icon for points.

## Requirements

```bash
python 3.9+
pip install simplekml openpyxl
# Linux: install python3-tk if Tkinter is missing
```

## Run

```bash
python kml_from_excel.py
```

* Pick an Excel file in the dialog.
* Output: `KML_Output/<SheetName>.kml` next to the Excel.

## Columns

* Required: headers containing `lat` and `lon` (case-insensitive).
* Optional: `DotColor` for marker color.

## Customize

* Edit the “Creator” row link/text in the code.
* Change table width via the `<table style="width:...">`.

## License

MIT
# KML_Maker

