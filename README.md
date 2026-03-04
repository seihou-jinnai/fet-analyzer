# fet-analyzer

Desktop application for extracting mobility (μ) and threshold voltage (Vth) 
from FET transfer characteristics using the saturation-regime model.

## Download

Download the latest version from:
https://github.com/seihou-jinnai/fet-analyzer/releases/latest

## Features

- Import Keithley Excel files
- P-type / N-type support
- Mobility and Vth extraction
- Batch processing
- Graph export

## How to Use

1. Drag and drop Keithley output xls file(s).
2. Select a file from “Select file.”
3. Select a sheet from “Select sheet.”
4. Input W, L, C, fit window, and Type.
5. From the sheet preview, select the columns for Isd and Vg, then click “Set as I-SD” and “Set as V-G,” respectively.
6. Click “add” to include in the analysis list.
7. Select a folder for exporting, then click “execute.”

## Output

- Isd–Vg plot
- √Isd–Vg plot with fitting
- Extracted parameters (μ, Vth)

## Author

Seihou Jinnai

## License

MIT License
