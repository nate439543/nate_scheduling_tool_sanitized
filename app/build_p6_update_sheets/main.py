"""Abandon all hope, ye who enter here."""
# Created by Nate Clark, June 2026

import datetime
from pathlib import Path
import sys
import os
import argparse

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_ROOT)

import utilities.ts_print
from onedrive_logic.copy_files_to_onedrive_logic import copy_file_to_new_path, copy_fragnet_template
from excel_logic.build_update_excel_file import build_update_excel_file
from excel_logic.modify_excel_logic import collect_weekly_update_input_by_project, unlock_excel_sheet, lock_excel_sheet
from utilities.get_config import get_config
from excel_logic.modify_excel_logic import set_excel_cell

def build_p6_update_sheet(old_data_date, new_data_date):


    try:
        config_output = get_config()

        print(f"Template path: {config_output.template_path}")
        projects_to_process = [project for project in config_output.excel_files_object.keys()]
      
        
        grouped_p6_data = collect_weekly_update_input_by_project(
            projects=projects_to_process,
            project_designator=config_output.weekly_update_project_designator,
            input_sheet_path=config_output.p6_data_path,
            input_sheet_name=config_output.p6_data_sheet_name,
            search_column_int=1,
            start_row=2,
            end_col=config_output.weekly_update_number_of_columns,
            start_col=1
        )
        if grouped_p6_data[projects_to_process[0]] == []:
            print(f"No data found for project {projects_to_process[0]}")
            return False

        old_data_date_obj = datetime.datetime.strptime(old_data_date, "%m%d%y")
        old_date_swaped = old_data_date_obj.strftime("%y%m%d")
        old_date_slashes = old_data_date_obj.strftime("%m/%d/%Y")
        old_date_dots = old_data_date_obj.strftime("%m.%d.%y")
        old_date_dashes = old_data_date_obj.strftime("%m-%d-%y")
        old_date_underscores_yearfirst = old_data_date_obj.strftime("%Y_%m_%d")
        old_date_numbers = old_data_date_obj.strftime("%y%m%d")
        old_date_long_dots = old_data_date_obj.strftime("%Y.%m.%d")

        next_data_date_obj = datetime.datetime.strptime(new_data_date, "%m%d%y")

        next_data_date_slashes = next_data_date_obj.strftime("%m/%d/%Y")
        set_excel_cell(
            config_output.template_path,
            "Config",
            "A2",
            old_date_slashes,
            recalculate_after_write=True,
            recalc_visible=True,
            focus_window_for_recalc=False,
            home_sheet="Updates",
            home_cell="A1",
        )
        print(f"Setting next data date in UPDATE INPUT sheet to {next_data_date_slashes}")
        set_excel_cell(
            config_output.template_path,
            "Config",
            "B2",
            next_data_date_slashes,
            recalculate_after_write=True,
            recalc_visible=True,
            focus_window_for_recalc=False,
            home_sheet="Updates",
            home_cell="A1",
        )

        for project, update_sheet in config_output.excel_files_object.items():
            
            
            new_file_name = (update_sheet
                .replace('{date_slashes}', old_date_slashes)
                .replace('{date_dots}',   old_date_dots)
                .replace('{date_dashes}', old_date_dashes)
                .replace('{date_numbers}', old_date_numbers)
                .replace('{data_date}',   old_data_date)
                .replace('{date_swaped}', old_date_swaped)
                .replace('{date_underscores_yearfirst}', old_date_underscores_yearfirst))
            
            
            archive_folder_base_path_dated = os.path.join(config_output.archive_folder_base_path, old_date_long_dots)
            archive_folder_path_dated = os.path.join(config_output.archive_folder_base_path, old_date_long_dots, new_file_name)
            
            
            print(f"Copying template for {project} to archive folder with new name: {new_file_name}")
            copy_file_to_new_path(config_output.template_path, archive_folder_path_dated)
            
            p6_data_file_path = os.path.join(config_output.template_folder_path, "Current_p6_Data.xlsm")


            next_data_date_slashes = next_data_date_obj.strftime("%m/%d/%Y")

            build_update_excel_file(project, grouped_p6_data[project], archive_folder_path_dated, old_date_slashes, next_data_date_slashes, config_output.weekly_update_password)

            archive_folder_path_dated = Path(archive_folder_path_dated)
            
            # print(f"Recalculating workbook for cached values to ensure formulas are updated before upload to OneDrive\nPath: {archive_folder_path_dated}")
            # _recalculate_workbook_for_cached_values(archive_folder_path_dated)
            
            print(f"Updated Excel file for {project} with new data date {next_data_date_slashes} and saved to archive folder")
            output_folder_path = os.path.join(config_output.output_folder_base_path, old_date_long_dots)
            output_file_path = os.path.join(output_folder_path, new_file_name)
            print(f"Copying updated file for {project} from archive to output folder: {output_file_path}")
            copy_file_to_new_path(archive_folder_path_dated, output_file_path)

        copy_fragnet_template(config_output.fragnet_template_path, output_folder_path, old_date_long_dots)
        return True


    except Exception as e:
        print(f"Error in build_p6_update_sheet: {str(e)}")
        return False
def main():
    # Set up command line argument parsing
    print(f"---Setup Excel Update Files---")
    parser = argparse.ArgumentParser(description='Rename files to OneDrive')
    parser.add_argument('--old_data_date', type=str, help='Last Data Date in MMDDYY format (e.g., 060526)')
    parser.add_argument('--new_data_date', type=str, help='Next Data Date in MMDDYY format (e.g., 060626)')

    args = parser.parse_args()
    
    # Use provided date or default to today's date
    if args.old_data_date:
        data_date = args.old_data_date
    else:
        # Default to today's date
        print("No date provided... \n\n\nThat's, like, an important part.")
        return
    
    if args.new_data_date: 
        data_date = args.new_data_date
    else :
        print("No new data date provided... \n\n\nThat's, like, an important part.")
        return

        
    
    print(f"Renaming files with a previous data date: {data_date}  and a next data date: {args.new_data_date}")
    



    result = build_p6_update_sheet(
        old_data_date = args.old_data_date,
        new_data_date = args.new_data_date

    )
    
    
    if result:
        
        print(f"Files renamed")
    else:
        print("Renaming process failed.")

if __name__ == "__main__":
    main()