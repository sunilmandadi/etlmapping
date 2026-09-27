ETL Mapping Copilot Prompt Builder (WinPython Portable, no Streamlit)

What it does
- Opens an Excel .xlsx or .xlsm mapping workbook.
- Lets you select a worksheet and extracts recognized source-to-target fields.
- Preserves other populated mapping columns under additional_fields.
- Reads mapping rows even if Excel omitted worksheet dimension metadata.
- Displays JSON and creates EDA, EDW, Datamart, and ETL script prompts.
- Copies prompts to the clipboard or saves prompts/JSON to files.
- Saves all four prompts and the JSON into one ZIP.
- Processes the workbook locally; it does not send data to Copilot.

Prerequisites
- Windows 10/11 and WinPython Portable with Tkinter included.
- Python 3.9 or later recommended.
- openpyxl installed in the WinPython environment.
- An .xlsx or .xlsm mapping workbook.

Run with WinPython
1. Extract this ZIP, for example to D:\etl\ETL_Mapping_Copilot_Tkinter_WinPython.
2. Open Command Prompt and confirm that WinPython's python\python.exe exists
   and reports the expected version. Running python without its full path may
   select a different Python installation.
3. Install the Excel reader once using that same Python executable:
   "D:\WinPython\WPy64-31470\python\python.exe" -m pip install -r "D:\etl\ETL_Mapping_Copilot_Tkinter_WinPython\requirements.txt"
4. Start the app using that executable (adjust the WinPython path if needed):
   "D:\WinPython\WPy64-31470\python\python.exe" "D:\etl\ETL_Mapping_Copilot_Tkinter_WinPython\app.py"

Do not type the app.py path alone: Windows may open it in an editor instead.
If python\python.exe does not exist, extract a complete WinPython distribution
before running these commands. If you use an already installed Python instead,
use the same python command for both pip installation and running app.py.

If your bank uses an internal package repository, install openpyxl from that
repository instead of the public internet. Tkinter is part of Python and is not
installed with pip. If the GUI fails with "No module named tkinter", use a
WinPython build that includes Tkinter or ask IT to provide it.

In the app
1. Browse to the workbook and choose its mapping worksheet.
2. Click Extract JSON.
3. Choose EDA, EDW, Datamart, or ETL Scripts.
4. Optionally enter the platform and layer notes.
5. Click Generate / Preview, then Copy prompt or save the prompt files.

Header recognition includes common names such as Src_Table / Source Table,
Src_Column / Source Column, Mapping / Transform_Rule, Target Column / Tgt_Column,
Datatype / Data_Type, Null Rule, and Key Flag. Unrecognized populated columns
are retained in additional_fields. Review the extracted JSON before using it.
