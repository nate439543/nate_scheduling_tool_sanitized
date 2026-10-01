import os
import sys
from openpyxl.utils import get_column_letter
from excel_logic.modify_excel_logic import collect_weekly_update_input_by_project, set_excel_cell, refresh_data_in_excel_file
from excel_logic.modify_excel_logic import add_row_of_data_to_sheet, unlock_excel_sheet, lock_excel_sheet
from excel_logic.modify_excel_logic import expand_table_in_sheet_using_com, add_range_of_data_to_sheet_using_com


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_ROOT)

import utilities.ts_print

def build_update_excel_file(
	project: str,
  p6_data: list[list[str]],
  update_sheet_path: str,
  last_data_date_slashes: str,
  next_data_date_slashes: str,
  update_sheet_password: str,
  first_data_row: int = 4,
):
  "Main Modify Excel Function"

  
  try:
    print(f"Modifying Excel file for project: {project}")

    


    

    unlock_excel_sheet(update_sheet_path, "Updates", update_sheet_password)
    add_range_of_data_to_sheet_using_com(
      update_sheet_path,
      "Updates",
      starting_col=1,
      data=p6_data,
      starting_row=first_data_row,
    )
    print(f"Added {len(p6_data)} rows of P6 data to UPDATE INPUT sheet starting at row {first_data_row}.")
    if p6_data and p6_data[0]:
      header_row = max(1, first_data_row - 1)
      end_col = get_column_letter(len(p6_data[0]))
      end_row = first_data_row + len(p6_data) - 1
      new_table_range = f"A{header_row}:{end_col}{end_row}"
      # print(f"Expanding table in UPDATE INPUT sheet to range: {new_table_range}")
      # expand_table_in_sheet_using_com(update_sheet_path, "Updates", new_table_range)
    else:
      print("No P6 data rows were provided. Skipping table expansion.")
      
    print(f"Locking UPDATE INPUT sheet with password: {update_sheet_password}")
    lock_excel_sheet(update_sheet_path, "Updates", update_sheet_password)
  except Exception as e:
    print(
            f"\n<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<\n"
            f"\n\n\n\n Error in build_update_excel_file: {str(e)}"
            f"\n\n\n\n>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>\n"
          )