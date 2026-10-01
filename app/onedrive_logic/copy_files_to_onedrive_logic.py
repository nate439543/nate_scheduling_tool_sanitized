import os
import sys
import shutil

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_ROOT)

import utilities.ts_print


def create_folder_in_directory(directory_path, folder_name):
  """Create folder_name inside directory_path and return the created path."""
  if not directory_path:
    raise ValueError("directory_path cannot be empty")
  if not folder_name:
    raise ValueError("folder_name cannot be empty")

  folder_path = os.path.join(directory_path, folder_name)
  os.makedirs(folder_path, exist_ok=True)
  return folder_path

def read_files_in_directory(directory_path) -> list[str]:
  """Return a list of file paths in the given directory."""
  if not os.path.isdir(directory_path):
    raise ValueError(f"{directory_path} is not a valid directory")
  
  return [os.path.join(directory_path, f) for f in os.listdir(directory_path) if os.path.isfile(os.path.join(directory_path, f))]
 
def copy_file_to_new_path (old_path, new_path):
      print(f"Checking existence of: {old_path}")
      if not os.path.isfile(old_path):
          # Nothing to do for this mapping if the exact short name isn't present
          print(f"File not found: {old_path}")
          return 

      # print(f"Matched exact file -> renaming/copying: {short_name} -> {new_file_name}")
      # new_path = os.path.join(folder_directory, new_file_name)

      print(f"Copying {old_path} -> {new_path}")
      try:
        destination_dir = os.path.dirname(new_path)
        if destination_dir:
          parent_dir, leaf_dir = os.path.split(destination_dir)
          if parent_dir and leaf_dir:
            create_folder_in_directory(parent_dir, leaf_dir)
          else:
            os.makedirs(destination_dir, exist_ok=True)
        shutil.copy2(old_path, new_path)
      except Exception as e:
        print(f"Error copying {old_path} to {new_path}: {str(e)}")
        return
def copy_fragnet_template(fragnet_template_path, dest_path, old_date_long_dots):
  try:
    if not os.path.isfile(fragnet_template_path):
        print(f"Fragnet template file not found: {fragnet_template_path}")
        return

    destination_dir = os.path.join(dest_path,"Fragnets")
    os.makedirs(destination_dir, exist_ok=True)
    print(f"Ensured destination directory exists: {destination_dir}")

    destination_path = os.path.join(destination_dir, os.path.basename(fragnet_template_path))
    shutil.copy2(fragnet_template_path, destination_path)
    print(f"Template file found: {fragnet_template_path}")
    print(f"Copied fragnet template to: {destination_path}")
  except Exception as e:
    print(f"Error copying fragnet template: {str(e)}")

