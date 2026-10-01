import sys
import os
import argparse
from timeit import main
from datetime import datetime
from datetime import timedelta






SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_ROOT)
from excel_logic.modify_excel_logic import collect_comments_and_user_input
from utilities.get_config import get_config
from utilities.date_calculations import determine_next_data_date
from file_manipulation_logic.file_manipulation_logic import copy_file_to_new_path
from excel_logic.modify_excel_logic import collect_comments_and_user_input, remove_old_sheets, remove_rows_containing_value


    

def get_update_data(old_data_date):
  
  print("---Get Update Data---")

  config_output = get_config()
  try:
    

    old_data_date_obj = datetime.strptime(old_data_date, "%m%d%y")
    old_date_swaped = old_data_date_obj.strftime("%y%m%d")
    old_date_slashes = old_data_date_obj.strftime("%m/%d/%Y")
    old_date_dots = old_data_date_obj.strftime("%m.%d.%y")
    old_date_dashes = old_data_date_obj.strftime("%m-%d-%y")
    old_date_underscores_yearfirst = old_data_date_obj.strftime("%Y_%m_%d")
    old_date_numbers = old_data_date_obj.strftime("%y%m%d")
    old_date_long_dots = old_data_date_obj.strftime("%Y.%m.%d")
    
   
    config_output = get_config()

    print(f"Analysis Source Path: {config_output.update_analysis_source_path}")
    
    staging_path = config_output.archive_folder_base_path
    copy_file_to_new_path(config_output.update_analysis_source_path, staging_path)
    staging_path = os.path.join(staging_path, os.path.basename(config_output.update_analysis_source_path))
    print(f"Copied update analysis source file to staging path: {staging_path}")
    
      

    if not config_output.update_analysis_source_path:
      raise ValueError("update_analysis_source_path is missing in config.ini [details]")
    

    print("Removing Rows Containing Old Data for the Current Data Date")
    for sheet_name, sheet_range in config_output.update_analysis_ranges.items():
      print(f"Removing rows containing '{old_date_dots}' from sheet '{sheet_name}'")
      remove_rows_containing_value(staging_path, sheet_name, "A", old_date_dots)
    
    for project, update_sheet in config_output.excel_files_object.items():
      
      update_sheet_name = (update_sheet
        .replace('{date_slashes}', old_date_slashes)
        .replace('{date_dots}',   old_date_dots)
        .replace('{date_dashes}', old_date_dashes)
        .replace('{date_numbers}', old_date_numbers)
        .replace('{data_date}',   old_data_date)
        .replace('{date_swaped}', old_date_swaped)
        .replace('{date_underscores_yearfirst}', old_date_underscores_yearfirst))
      
      sp_file_path = os.path.join(config_output.sp_folder_path, old_date_long_dots, update_sheet_name)
      
      print(f"Collecting comments and user input for {project} from {sp_file_path} to {config_output.update_analysis_path}")
      collect_comments_and_user_input(sp_file_path, staging_path, old_date_dots, config_output.update_analysis_ranges, old_date_dots,True)
      copy_file_to_new_path(staging_path, config_output.update_analysis_source_path)
      print(f"Copied updated analysis file to: {config_output.update_analysis_source_path}")
  except Exception as e:
    print(f"Error in get_update_data: {str(e)}")
    

    
    return
    
    
  
  
  
  
  
def main():
  # Set up command line argument parsing
  print(f"---Get Update Data---")
  parser = argparse.ArgumentParser(description='Get Update Data')
  parser.add_argument('--old_data_date', type=str, help='Last Data Date in MMDDYY format (e.g., 060526)')
  args = parser.parse_args()


  get_update_data(args.old_data_date)

if __name__ == "__main__":
  main()