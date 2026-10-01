import sys
import os
import argparse
from timeit import main
from datetime import datetime
from datetime import timedelta






SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_ROOT)
from onedrive_logic.copy_files_to_onedrive_logic import read_files_in_directory, copy_file_to_new_path, create_folder_in_directory
from excel_logic.modify_excel_logic import _find_first_empty_row, _get_cell_value, collect_frag_data, get_sheet_by_path_and_name, add_data_to_frag_log
from utilities.get_config import get_config
from utilities.date_calculations import determine_next_data_date
from file_manipulation_logic.file_manipulation_logic import copy_file_to_new_path, rename_file_in_directory




def frag_logger(old_data_date):
  


  config_output = get_config()
  try:
    
    # Initialize -------------------------------------------
    old_data_date_obj = datetime.strptime(old_data_date, "%m%d%y")

    old_date_dots = old_data_date_obj.strftime("%m.%d.%y")

    old_date_long_dots = old_data_date_obj.strftime("%Y.%m.%d")
    
   
    config_output = get_config()

    
      
    # Read Frag Files ----------------------------------------------
    if not config_output.sp_folder_path:
      raise ValueError("sp_folder_path is missing in config.ini [details]")
    

    sp_frag_folder_path = os.path.join(config_output.sp_folder_path, old_date_long_dots, "Fragnets")
      
    print(f"Reading all files in {sp_frag_folder_path}")
    
    frag_path_strings = read_files_in_directory(sp_frag_folder_path)
    
    frag_template_file_name = config_output.fragnet_template_path.split("/")[-1]
    print(f"Frag Template File Name: {frag_template_file_name}")
    frag_template_path = os.path.join(sp_frag_folder_path, frag_template_file_name)
    print(f"Frag template path: \n{frag_template_path}")
    
    print(f"Frag Path Strings: \n{frag_path_strings}")
    print(f"Removing fragnet template path from frag_path_strings")
    
    frag_path_strings.remove(frag_template_path)
    print(f"Total Frags Detected: {len(frag_path_strings)}")
    
    
    # Get Frag Log Data ----------------------------------------------
    frag_log_path = config_output.fragnet_log_path
    print(f"Getting Frag Sheet from: \n{frag_log_path}")
    frag_log_sheet = get_sheet_by_path_and_name(frag_log_path, config_output.fragnet_log_sheet_name)
    frag_log_first_empty_row = _find_first_empty_row(frag_log_sheet, 6, 1)
    print(f"First empty row in frag log: {frag_log_first_empty_row}")
    log_number = int(_get_cell_value(frag_log_sheet, frag_log_first_empty_row-1, 2))
    print(f"Last log number in frag log: {log_number}")
    
    
    frag_data = []
    
    for frag_path_string in frag_path_strings:
      #Collect frag data
      frag_sheet = get_sheet_by_path_and_name(frag_path_string, config_output.fragnet_template_sheet_name, data_only=True)
      frag_data.append(collect_frag_data(frag_sheet, config_output.frag_sheet_ranges))
      print(f"Collected frag data from {frag_path_string}:\n {frag_data[-1]}")
      #Copy to OD
      copy_file_to_new_path(frag_path_string, os.path.join(config_output.fragnet_bank_path, f"{log_number+1}_{os.path.basename(frag_path_string)}"))
      
    
    print(f"Adding {len(frag_data)} frag data entries to frag log at {frag_log_path}")
    
    for frag_entry in frag_data:
      add_data_to_frag_log(frag_log_path, config_output.fragnet_log_sheet_name, frag_entry, log_number+1, config_output.frag_log_ranges)
      log_number += 1

    
    

  except Exception as e:
    print(f"Error in frag_logger: {str(e)}")
    

    
    return
    
    
  
  
  
  
  
def main():
  # Set up command line argument parsing
  print(f"---Get Update Data---")
  parser = argparse.ArgumentParser(description='Get Update Data')
  parser.add_argument('--old_data_date', type=str, help='Last Data Date in MMDDYY format (e.g., 060526)')
  args = parser.parse_args()


  frag_logger(args.old_data_date)

if __name__ == "__main__":
  main()