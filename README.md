# Experimental Plan Builder

A standalone Tkinter application for creating experimental plans and exporting them to formatted `.xlsx` workbooks.

## Run in VS Code on Windows

Install the **Python** extension for VS Code, open this folder, then press **F5**
or choose **Terminal > Run Task > Run Experimental Plan Builder**. On the first
run, the project downloads Python 3.13.14 from python.org into the local
`.python` folder and installs the packages from `requirements.txt`. Later runs
reuse that runtime. No global Python installation, PATH changes, or manual
interpreter selection are needed.

The first run requires an internet connection and permission to run the setup
script. The project-local runtime takes up disk space and is not included when
you copy or share the source folder; each PC bootstraps its own runtime.

To prepare the runtime without launching the app, run `setup.ps1` in PowerShell.
The project-local interpreter is `.python\python.exe`.

The app starts with `Copy of MA_PPDT2FBT_PCBM.xlsx` as a five-sheet template. Each sheet is available in its own tab and can be edited directly in the grid:

- `Ingredients`: project title, active materials, solvents, and substrates
- `Sample list`: constant and varying preparation parameters
- `Characterization`: film and solar-cell characterization methods
- `Evaluation and Analysis`: evaluation outputs and comments
- `Tasks & People involved`: responsibilities and working plan

Use **Open workbook** to edit another workbook. Use **Export XLSX** to save all sheets while preserving the workbook's existing formatting and merged cells.

To merge cells in the editor, click the first cell, then **Shift-click** the last cell in the rectangular range and choose **Merge cells**. If other selected cells contain values, the editor asks before merging; the top-left value is kept. A range that overlaps an existing merged area cannot be merged again.

## Test the export without opening the GUI

```powershell
& .\.python\python.exe -m unittest discover -s tests -v
```
