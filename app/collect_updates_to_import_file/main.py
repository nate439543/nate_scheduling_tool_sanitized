import sys
import os
import argparse
from timeit import main
from datetime import datetime
from datetime import timedelta






SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_ROOT)
from excel_logic.modify_excel_logic import collect_comments_and_user_input, create_import_file
from utilities.get_config import get_config
from utilities.date_calculations import determine_next_data_date
from excel_logic.modify_excel_logic import (
  collect_updates,
  create_import_file, 
  add_row_of_data_to_sheet,
)


    

def collect_updates_to_import_file(old_data_date, current_time):
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
    
    update_list = config_output.update_list


    for project, update_sheet in config_output.excel_files_object.items():
      import_folder_path = config_output.import_folder_path
      import_file_path = create_import_file(
        save_path=import_folder_path,
        file_name=f"{project}_Import_{current_time.strftime('%Y%m%d_%H%M%S')}.xlsx",
        headers=config_output.update_sheet_columns,
        userdata=config_output.userdata_sheet_data
      )
      update_data = []
      for update in update_list:
        print(f"Collecting update data for: {update}")

          
        if update == "OFCI":
          for ofci_sheet_type, ofci_sheet_name in config_output.ofci_file_sheet_names.items():
            project_number = ''.join(filter(str.isdigit, project))
            sheet_name = ofci_sheet_name.replace('{project}', project_number)
            print(f"Collecting update data for: {sheet_name}")
            new_data = collect_updates(
              workbook_path= config_output.ofci_updater_path,
              sheet_name= sheet_name,
              number_of_columns= config_output.import_number_of_columns,
            )
            print(f"Collected {len(new_data)} lines of data for {sheet_name}")
            print(f"Adding {len(new_data)} lines to update data, current length: {len(update_data)}")
            update_data.extend(new_data)

            
        elif update == "Weekly_Updates":
          update_sheet_name = (update_sheet
          .replace('{date_slashes}', old_date_slashes)
          .replace('{date_dots}',   old_date_dots)
          .replace('{date_dashes}', old_date_dashes)
          .replace('{date_numbers}', old_date_numbers)
          .replace('{data_date}',   old_data_date)
          .replace('{date_swaped}', old_date_swaped)
          .replace('{date_underscores_yearfirst}', old_date_underscores_yearfirst))

          sp_file_path = os.path.join(config_output.sp_folder_path, old_date_long_dots, update_sheet_name)
          new_data = collect_updates(
            workbook_path= sp_file_path,
            sheet_name= "Import",
            number_of_columns= config_output.import_number_of_columns,
          )
          print(f"Collected {len(new_data)} lines of data for {update_sheet_name}")
          print(f"Adding {len(new_data)} lines to update data, current length: {len(update_data)}")
          update_data.extend(new_data)

        elif update == "Auto_Updater":
          project_number = ''.join(filter(str.isdigit, project))
          sheet_name = config_output.auto_updater_sheet_name.replace('{project}', project_number)
          new_data = collect_updates(
            workbook_path= config_output.auto_updater_path,
            sheet_name= sheet_name,
            number_of_columns= config_output.import_number_of_columns,
          )
          print(f"Collected {len(new_data)} lines of data for {sheet_name}")
          print(f"Adding {len(new_data)} lines to update data, current length: {len(update_data)}")
          update_data.extend(new_data)
          
        
      for line in update_data:
        
        add_row_of_data_to_sheet(
          workbook_path=import_file_path,
          sheet_name="TASK",
          starting_col=1,
          data=line
        )
      print(f"{len(update_data)} Lines of Update data for {project} collected and saved to {import_file_path}")
  except Exception as e:
    print(f"Error collecting update data: {str(e)}")


def main():
  # Set up command line argument parsing
  print(f"---Get Update Data---")
  parser = argparse.ArgumentParser(description='Get Update Data')
  parser.add_argument('--old_data_date', type=str, help='Last Data Date in MMDDYY format (e.g., 060526)')
  # parser.add_argument('--update_list', type=str, nargs='+', help='List of updates to collect (e.g., OFCI Weekly_Updates)')
  # parser.add_argument('--project_blacklist', type=str, nargs='+', help='List of projects to exclude from processing (e.g., Project1 Project2)')
  args = parser.parse_args()


  collect_updates_to_import_file(args.old_data_date, current_time=datetime.now())

if __name__ == "__main__":
  main()