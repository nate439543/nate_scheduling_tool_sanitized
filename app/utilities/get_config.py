
import os
import sys
import shutil

import configparser

from dataclasses import dataclass, field


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(os.path.dirname(SCRIPT_DIR))
sys.path.insert(0, PROJECT_ROOT)

class CaseSensitiveConfigParser(configparser.ConfigParser):
  def optionxform(self, optionstr):
    return optionstr
@dataclass
class ConfigOutput:
  excel_files_object: dict[str, str] = field(default_factory=dict)
  update_analysis_ranges: dict[str, str] = field(default_factory=dict)
  update_sheet_columns: dict[str, str] = field(default_factory=dict)
  userdata_sheet_data: dict[str, str] = field(default_factory=dict)
  frag_sheet_ranges: dict[str, str] = field(default_factory=dict)
  frag_log_ranges: dict[str, str] = field(default_factory=dict)
  ofci_file_sheet_names: dict[str, str] = field(default_factory=dict)
  update_list: list[str] = field(default_factory=list)
  template_path: str = ""
  output_folder_base_path: str = ""
  archive_folder_base_path: str = ""
  template_folder_path: str = ""
  p6_data_path: str = ""
  p6_data_sheet_name: str = ""
  sp_folder_path: str = ""
  update_analysis_path: str = ""
  fragnet_template_path: str = ""
  fragnet_template_sheet_name: str = ""
  fragnet_log_path: str = ""
  fragnet_log_sheet_name: str = ""
  fragnet_bank_path: str = ""
  ofci_updater_path: str = ""
  auto_updater_path: str = ""
  auto_updater_sheet_name: str = ""
  import_folder_path: str = ""
  import_number_of_columns: int = 8
  weekly_update_project_designator: str = ""
  weekly_update_password: str = ""
  weekly_update_number_of_columns: int = 17
  schedule_names: dict[str, str] = field(default_factory=dict)
  procore_routing: dict[str, str] = field(default_factory=dict)
  schedule_publish_nomenclature: dict[str, str] = field(default_factory=dict)
  publish_files_output_path: str = ""
  publish_files_input_path: str = ""
  comment_analysis_path: str = ""
  ofci: bool = False
  weekly_updates: bool = False
  auto_updater: bool = False
  sharepoint_base_path: str = ""
  procore_base_path: str = ""
  activities_riding_dates_path: str = ""
  schedule_publish_nomenclature_path: str = ""
  update_analysis_source_path: str = ""
  onedrive_path_segment: str = ""
  user_onedrive_path_segment: str = ""
  use_user_onedrive_path_segment: str = "false"
  publish_parent_folder_path: str = ""
  schedule_tool_files_path_segment_external_user: str = ""
  schedule_tool_files_path_segment_host: str = ""
  use_schedule_tool_files_path_segment_host: str = "false"
  schedule_tool_files_path: str = ""
  local_tool_path: str = ""



def get_config():
    """
    Load configuration from config.ini 
    """
    config_output = ConfigOutput()
    config_file = os.path.join(PROJECT_ROOT, 'config.ini')
    print(f"Looking for config file at: {config_file}")

    if os.path.exists(config_file):
      try:
        print(f"Reading config file: {config_file}")
        
        
        
        parser = CaseSensitiveConfigParser()
        parser.read(config_file)
        
        # Update Sheet Names (project code -> update sheet file name)
        config_output.excel_files_object = dict(parser.items('Update Sheet Names'))

        # Schedule Names (short name -> published file name)
        config_output.schedule_names = dict(parser.items('schedule names'))

        # Procore Routing (short name -> Procore subfolder under procore_base_path, blank = none)
        config_output.procore_routing = dict(parser.items('Procore Routing'))

        # Schedule Publish Nomenclature (ID prefix -> schedule name template)
        config_output.schedule_publish_nomenclature = dict(parser.items('Schedule Publish Nomenclature'))

        # Excel File Paths
        config_output.template_path = parser["Excel File Paths"].get('template_path', "")
        config_output.archive_folder_base_path = parser["Excel File Paths"].get('archive_folder_base_path', "")
        config_output.template_folder_path = parser["Excel File Paths"].get('template_folder_path', "")
        config_output.p6_data_path = parser["Excel File Paths"].get('p6_data_path', "")
        config_output.update_analysis_path = parser["Excel File Paths"].get('update_analysis_path', "")
        config_output.comment_analysis_path = parser["Excel File Paths"].get('comment_analysis_path', "")
        config_output.activities_riding_dates_path = parser["Excel File Paths"].get('activities_riding_dates_path', "")
        config_output.schedule_publish_nomenclature_path = parser["Excel File Paths"].get('schedule_publish_nomenclature_path', "")
        config_output.update_analysis_source_path = parser["Excel File Paths"].get('update_analysis_source_path', "")
        config_output.local_tool_path = parser["Excel File Paths"].get('local_tool_path', "")
        # Details
        config_output.p6_data_sheet_name = parser["details"].get('p6_data_sheet_name', "")
        config_output.sp_folder_path = parser["details"].get('sp_folder_path', "")
        config_output.publish_files_output_path = parser["details"].get('publish_files_output_path', "")
        config_output.publish_files_input_path = parser["details"].get('publish_files_input_path', "")
        config_output.publish_files_input_path = parser["details"].get('publish_files_input_path', "")
        config_output.weekly_update_number_of_columns = parser["details"].getint('weekly_update_number_of_columns', 17)
        config_output.weekly_update_project_designator = parser["details"].get('weekly_update_project_designator', "")
        config_output.weekly_update_password = parser["details"].get('weekly_update_password', "")
        config_output.output_folder_base_path = parser["details"].get('output_folder_base_path', "")
        config_output.onedrive_path_segment = parser["details"].get('onedrive_path_segment', "")
        config_output.user_onedrive_path_segment = parser["details"].get('user_onedrive_path_segment', "")
        config_output.use_user_onedrive_path_segment = parser["details"].get('use_user_onedrive_path_segment', "false")
        config_output.schedule_tool_files_path_segment_external_user = parser["details"].get('schedule_tool_files_path_segment_external_user', "")
        config_output.schedule_tool_files_path_segment_host = parser["details"].get('schedule_tool_files_path_segment_host', "")
        config_output.use_schedule_tool_files_path_segment_host = parser["details"].get('use_schedule_tool_files_path_segment_host', "false")
        
        # Publish Output Paths
        config_output.sharepoint_base_path = parser["Publish Output Paths"].get('sharepoint_base_path', "")
        config_output.procore_base_path = parser["Publish Output Paths"].get('procore_base_path', "")
        config_output.publish_parent_folder_path = parser["Publish Output Paths"].get('publish_parent_folder_path', "")

        # Collect Updates Paths
        config_output.import_folder_path = parser["Collect Updates Paths"].get('import_folder_path', "")
        config_output.import_number_of_columns = parser["Collect Updates Paths"].getint('import_number_of_columns', 8)
        config_output.ofci_updater_path = parser["Collect Updates Paths"].get('ofci_updater_path', "")
        config_output.auto_updater_path = parser["Collect Updates Paths"].get('auto_updater_path', "")
        config_output.auto_updater_sheet_name = parser["Collect Updates Paths"].get('auto_updater_sheet_name', "")

        # Update List (only the updates flagged true, in file order)
        config_output.update_list = [
            key for key in parser['Update List']
            if parser['Update List'].getboolean(key, fallback=False)
        ]

        # Fragnets Details
        config_output.fragnet_template_path = parser["Fragnets Details"].get('fragnet_template_path', "")
        config_output.fragnet_template_sheet_name = parser["Fragnets Details"].get('fragnet_template_sheet_name', "")
        config_output.fragnet_log_path = parser["Fragnets Details"].get('fragnet_log_path', "")
        config_output.fragnet_log_sheet_name = parser["Fragnets Details"].get('fragnet_log_sheet_name', "")
        config_output.fragnet_bank_path = parser["Fragnets Details"].get('fragnet_bank_path', "")
        
        # Frag Sheet Ranges
        config_output.frag_sheet_ranges = dict(parser.items('Frag Sheet Ranges'))

        # Frag Log Ranges
        config_output.frag_log_ranges = dict(parser.items('Frag Log Ranges'))

        # Update Analysis Ranges
        config_output.update_analysis_ranges = dict(parser.items('Update Analysis Ranges'))

        # Update Sheet Columns
        config_output.update_sheet_columns = dict(parser.items('Update Sheet Columns'))

        # Userdata Sheet Data
        config_output.userdata_sheet_data = dict(parser.items('Userdata Sheet Data'))

        # OFCI File Sheet Names
        config_output.ofci_file_sheet_names = dict(parser.items('OFCI File Sheet Names'))

        onedrive_path_segment = config_output.user_onedrive_path_segment if config_output.use_user_onedrive_path_segment.lower() == "true" else config_output.onedrive_path_segment
        config_output.schedule_tool_files_path = config_output.schedule_tool_files_path_segment_host if config_output.use_schedule_tool_files_path_segment_host.lower() == "true" else config_output.schedule_tool_files_path_segment_external_user
        schedule_tool_files_path_segment = config_output.schedule_tool_files_path
        path_fields = [
            'template_path', 'archive_folder_base_path', 'template_folder_path',
            'sp_folder_path', 'update_analysis_path', 'p6_data_path',
            'fragnet_template_path', 'fragnet_log_path', 'fragnet_bank_path',
            'ofci_updater_path', 'output_folder_base_path', 'import_folder_path',
            'auto_updater_path',
            'sharepoint_base_path', 'procore_base_path',
            'activities_riding_dates_path', 'comment_analysis_path', 'publish_files_input_path',
            'schedule_publish_nomenclature_path','update_analysis_source_path', 'publish_parent_folder_path','schedule_tool_files_path','local_tool_path'
        ]
        
        for field_name in path_fields:
            setattr(config_output, field_name, getattr(config_output, field_name).replace("{ONEDRIVE_PATH_SEGMENT}", onedrive_path_segment))
            setattr(config_output, field_name, getattr(config_output, field_name).replace("{SCHEDULE_TOOL_FILES_PATH_SEGMENT}", schedule_tool_files_path_segment))
            setattr(config_output, field_name, os.path.expanduser(getattr(config_output, field_name)))
            setattr(config_output, field_name, getattr(config_output, field_name).replace("{PROJECT_ROOT}", PROJECT_ROOT))

      except Exception as e:
        print(f"Warning: Error reading config file: {str(e)}")
    else:
      print(f"Error: Config file not found: {config_file}.")
    return config_output