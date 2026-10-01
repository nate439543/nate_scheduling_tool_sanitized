# Nate's Scheduling Tool

A Windows desktop app that automates a weekly Primavera P6 schedule update
workflow: building update-input Excel workbooks from P6 data, collecting
completed updates back into an import file, analyzing update progress, and
publishing/renaming finished schedule files out to OneDrive/SharePoint and
Procore.

The app is a [NiceGUI](https://nicegui.io/) desktop window (backed by
`pywebview`) that drives Excel over COM automation (`pywin32`) and
`openpyxl`, so it requires Windows with Excel installed.

## Quick start

1. Install [Python 3.13](https://www.python.org/downloads/) (check "Add
   Python to PATH" during install).
2. Run `setup.bat` — creates a `.venv` virtual environment and installs
   dependencies from `requirements.txt`.
3. Edit `config.ini` with your paths, project list, and file naming
   conventions (see [Configuration](#configuration)).
4. Launch the app with `Nate's Scheduling Tool.bat` (or the `.lnk` shortcut).

## What it does

The GUI is organized into three tabs, each wrapping one or more scripts:

### Build
![alt text]([https://github.com/nate439543/nate_scheduling_tool_sanitized/tree/main/Extra/README_Images/analyze_tab.png] "Build Tab")
- **Build P6 Update Sheets**  — pulls P6 export data and generates the per-project
  weekly update Excel workbooks for the given data date range.
- Shortcuts to open the fragnet template, the P6 data source file, the
  output folder, and the update template.

### Analyze
![alt text](https://github.com/nate439543/nate_scheduling_tool_sanitized/tree/main/Extra/README_Images/analyze_tab.png "Analyze Tab")
- **Grab Update Data For Analysis**  — Pulls comments/user input data from each project's
  published update sheet into a source file which is connected to any reports/analysis files via
  Power Query.
- Shortcuts to open the activities-riding-date, update analysis, and comment
  analysis workbooks.

### Process
![alt text](https://github.com/nate439543/nate_scheduling_tool_sanitized/tree/main/Extra/README_Images/process_tab.png "Process Tab")
- **Generate Import Files for Current Updates** — Collects weekly updates from OFCI,
  Weekly Update, and Auto Updater excel files. Converts this data into excel files for each
  project which can be directly imported into P6.
- **Generate Publish Nomenclature**  — builds the baseline/schedule ID/schedule name lookup
  spreadsheet used when publishing schedules for a given data date.
- **Move and Rename Publish Files**  — matches files dropped in the publish input folder
  against `config.ini`'s schedule name map, renames/copies them, and pushes them out
  to the SharePoint and Procore OneDrive folders (with project-specific routing rules).
- Shortcuts to the OFCI Updater, Auto Updater, and import/publish-input
  folders.

## Project layout

```
app/
  nate_scheduling_tool_gui.py      # NiceGUI desktop app / entry point
  build_p6_update_sheets/          # Build weekly update workbooks from P6 data
  collect_updates_to_import_file/  # Collect OFCI/weekly/auto-updater rows into an import file
  get_update_data/                 # Pull comments/progress into the analysis workbook
  reorganize_to_onedrive/          # Rename + publish schedule files to OneDrive/Procore
  sched_update_file_org/           # Generate schedule publish nomenclature
  frag_logger/                     # Log fragnet entries to the fragnet log
  excel_logic/                     # Excel read/write helpers (openpyxl + COM)
  onedrive_logic/                  # File copy helpers for OneDrive-synced folders
  file_manipulation_logic/         # Generic file copy/rename helpers
  utilities/                       # Config loading, date math, timestamped printing
Excel Files/                       # Templates, nomenclature, and import/publish staging folders
config.ini                         # All paths, file-naming templates, and per-project settings
setup.bat                          # Creates .venv and installs requirements.txt
Nate's Scheduling Tool.bat          # Launches the GUI
```

## Configuration

Everything project- and path-specific lives in `config.ini`, including:

- `[Update Sheet Names]` — project code → update workbook file name template.
- `[Excel File Paths]` / `[Publish Output Paths]` / `[details]` — local and
  OneDrive paths for templates, data sources, analysis workbooks, and
  publish destinations. `{PROJECT_ROOT}` resolves to the repo root, `~`
  resolves to the user's home directory, and `{ONEDRIVE_PATH_SEGMENT}` /
  `{SCHEDULE_TOOL_FILES_PATH_SEGMENT}` resolve based on the
  `use_user_onedrive_path_segment` / `use_schedule_tool_files_path_segment_host`
  toggles (so the same config can point at a personal or shared OneDrive
  layout).
- `[Update List]` — which update sources (`OFCI`, `Weekly_Updates`,
  `Auto_Updater`) are collected.
- `[schedule names]` — short file name → full published file name template,
  used when renaming/routing files to OneDrive and Procore.
- `[Fragnets Details]` / `[Frag Sheet Ranges]` / `[Frag Log Ranges]` — paths
  and cell ranges for fragnet logging.

Filename templates support placeholders like `{date_slashes}`, `{date_dots}`,
`{date_dashes}`, `{date_swaped}`, `{date_underscores_yearfirst}`, and
`{data_date}`, which are filled in from the date(s) entered in the GUI.

Config is loaded via `app/utilities/get_config.py`, which parses `config.ini`
into a typed `ConfigOutput` dataclass and resolves all path placeholders.

## Requirements

- Windows with Microsoft Excel installed (COM automation via `pywin32`).
- Python 3.13+.
- See `requirements.txt` for Python dependencies (`nicegui`, `pywebview`,
  `openpyxl`, `pywin32`, `pandas`, `msal`, `requests`).

## Notes

This is a personal automation tool tailored to a specific project's file
naming conventions and OneDrive/SharePoint folder layout (see `config.ini`),
not a general-purpose scheduling library.
