
# Created by Nate Clark, September 2025

import sys
import os
import argparse
from datetime import datetime

import shutil
from onedrive_logic.copy_files_to_onedrive_logic import copy_file_to_new_path
from utilities.get_config import get_config


sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def reorganize_files_to_onedrive(data_date, copy=True):

  try:
      

    date_obj = datetime.strptime(data_date, "%m%d%y")
    date_swaped = date_obj.strftime("%y%m%d")
    date_slashes = date_obj.strftime("%m/%d/%y")
    date_dots = date_obj.strftime("%m.%d.%y")
    date_dashes = date_obj.strftime("%m-%d-%y")
    date_numbers = date_obj.strftime("%y%m%d")
    date_long_dots = date_obj.strftime("%Y.%m.%d")

    config_output = get_config()

    if config_output.publish_files_input_path is None:
      print("Error: 'publish_files_input_path' not found in config.ini")
      return False
    
        
        
    # create_onedrive_folder(os.path.join(sharepoint_base_path), date_long_dots, headers, requests)
    
    
    if data_date[2:4] == '01':
        suffix = " Monthly"
    else:
        suffix = ""

    for pre_short_name, long_name in config_output.schedule_names.items():

      new_file_name = (long_name
                        .replace('{date_slashes}', date_slashes)
                        .replace('{date_dots}',   date_dots)
                        .replace('{date_dashes}', date_dashes)
                        .replace('{date_numbers}', date_numbers)
                        .replace('{data_date}',   data_date)
                        .replace('{date_swaped}', date_swaped))

      short_name = pre_short_name.replace('{date_swaped}', date_swaped)
      folder_directory = os.path.join(config_output.publish_files_input_path)
      print(f"Searching for {short_name} in {folder_directory}")
      
      
      


      # Only act on an exact file match; avoid accidental matches like PROJ2ms.pdf
      old_path = os.path.join(folder_directory, short_name)
      print(f"Checking existence of: {old_path}")
      if not os.path.isfile(old_path):
          # Nothing to do for this mapping if the exact short name isn't present
          print(f"File not found: {old_path}")
          continue

      print(f"Matched exact file -> renaming/copying: {short_name} -> {new_file_name}")
      new_path = os.path.join(folder_directory, new_file_name)

      if copy:
        print(f"Copying {short_name} -> {new_file_name}")
        try:
          shutil.copy2(old_path, new_path)
        except Exception as e:
          if short_name == new_file_name:
            print(f"Source and destination are the same file: {short_name}")
            # Continue with upload logic below
          else:
            print(f"Error copying {short_name} to {new_file_name}: {str(e)}")
            continue
      else:
        print(f"--------------------------------------\nRenaming \n {old_path} to \n{new_path}\n--------------------------------------")
        os.rename(old_path, new_path)

      #push to onedrive
      sharepoint_base_path = os.path.join(config_output.sharepoint_base_path)
      sharepoint_path = os.path.join(sharepoint_base_path, date_long_dots, new_file_name)
      procore_base_path = os.path.join(config_output.procore_base_path)

      procore_subfolder = config_output.procore_routing.get(pre_short_name, "").strip()
      if procore_subfolder:
        alt_onedrive_path = os.path.join(procore_base_path, procore_subfolder, f"{date_long_dots}{suffix}", new_file_name)
        print(f"Routing {short_name} to Procore subfolder '{procore_subfolder}': {alt_onedrive_path}")
      else:
        alt_onedrive_path = False
        print(f"No Procore routing configured for {short_name}, skipping alternate OneDrive copy")

      if data_date[2:4] == '01':
        sharepoint_path = sharepoint_path.replace("WEEKLY","MONTHLY")
        alt_onedrive_path = alt_onedrive_path.replace("WEEKLY","MONTHLY") if alt_onedrive_path else False
        print(f"Adjusted paths for monthly suffix where applicable.\n sharepoint_path: {sharepoint_path}\n alt_onedrive_path: {alt_onedrive_path}")

      # Copy the file to its OneDrive-synced destinations
      if not ("(alt ms)" in new_file_name.lower()):
        print(f"(alt ms) not in {new_file_name}, copying to sharepoint_path: {sharepoint_path}")
        copy_file_to_new_path(new_path, sharepoint_path)

      if alt_onedrive_path:
        copy_file_to_new_path(new_path, alt_onedrive_path)
    return True
  except Exception as e:
    print(f"Error: {str(e)}")
    return False
def main():
    # Set up command line argument parsing
    print(f"rename_to_onedrive_main called")
    parser = argparse.ArgumentParser(description='Rename files to OneDrive')
    parser.add_argument('--date', type=str, help='Date in MMDDYY format (e.g., 091525)')
    args = parser.parse_args()
    
    # Use provided date or default to today's date
    if args.date:
        data_date = args.date
    else:
        # Default to today's date
        print("No date provided, ask Nate what to do.")
        return
    
    print(f"Renaming files with data date: {data_date}")
    



    result = reorganize_files_to_onedrive(
        data_date=data_date,

    )
    
    if result:
        
        print(f"Files renamed")
    else:
        print("Renaming process failed.")

if __name__ == "__main__":
    main()