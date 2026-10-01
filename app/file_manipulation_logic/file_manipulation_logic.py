import os
import sys
import shutil

import configparser
import json
from datetime import datetime
from datetime import timedelta
import time
from pathlib import Path
from dataclasses import dataclass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_ROOT)


def create_folder_in_directory(directory_path, folder_name):
  """Create folder_name inside directory_path and return the created path."""
  if not directory_path:
    raise ValueError("directory_path cannot be empty")
  if not folder_name:
    raise ValueError("folder_name cannot be empty")

  folder_path = os.path.join(directory_path, folder_name)
  os.makedirs(folder_path, exist_ok=True)
  return folder_path

def rename_file_in_directory(directory_path, old_file_name, new_file_name):
  """Rename a file in the given directory."""
  if not directory_path:
    raise ValueError("directory_path cannot be empty")
  if not old_file_name:
    raise ValueError("old_file_name cannot be empty")
  if not new_file_name:
    raise ValueError("new_file_name cannot be empty")

  old_file_path = os.path.join(directory_path, old_file_name)
  new_file_path = os.path.join(directory_path, new_file_name)

  if not os.path.isfile(old_file_path):
    raise FileNotFoundError(f"{old_file_path} does not exist")

  os.rename(old_file_path, new_file_path)
  return new_file_path

def move_file_to_new_path(old_path, new_path):
    """Move a file from old_path to new_path."""
    if not os.path.isfile(old_path):
        raise FileNotFoundError(f"{old_path} does not exist")

    destination_dir = os.path.dirname(new_path)
    if destination_dir:
        os.makedirs(destination_dir, exist_ok=True)

    shutil.move(old_path, new_path)
    return new_path
  
def copy_file_to_new_path(old_path, new_path):
    """Copy a file from old_path to new_path."""
    if not os.path.isfile(old_path):
        raise FileNotFoundError(f"{old_path} does not exist")

    destination_dir = os.path.dirname(new_path)
    if destination_dir:
        os.makedirs(destination_dir, exist_ok=True)

    shutil.copy2(old_path, new_path)
    return new_path