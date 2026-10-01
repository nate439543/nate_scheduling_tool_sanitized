from datetime import datetime
import pandas as pd
import sys
import os
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_ROOT)
from utilities.get_config import get_config
config = get_config()

if len(sys.argv) < 2:
    print("Usage: python generate_schedules.py MMDDYY")
    sys.exit(1)

date_input = sys.argv[1]

try:
    date_obj = datetime.strptime(date_input, "%m%d%y")
except ValueError:
    print("Invalid date format. Please use MMDDYY format (e.g., 090525)")
    sys.exit(1)

filename_date = date_obj.strftime("%y%m%d")
schedule_date = date_obj.strftime("%m.%d.%y")
# LA date in ISO format (YYYY-MM-DD)
la_date = date_obj.strftime("%Y-%m-%d")

# Check if the third and fourth characters are '01' (first day of month)
frequency = "MONTHLY" if date_input[2:4] == "01" else "WEEKLY"
prefix = f"{filename_date} " if date_input[2:4] == "01" else ""

schedule_names = []
schedule_ids = []
baseline_ids = []

for id_prefix, name_template in config.schedule_publish_nomenclature.items():
    name = (name_template
            .replace('{prefix}', prefix)
            .replace('{frequency}', frequency)
            .replace('{schedule_date}', schedule_date)
            .replace('{la_date}', la_date))

    schedule_names.append(name)
    schedule_ids.append(f"{id_prefix}.{filename_date}")
    baseline_ids.append(f"{id_prefix}-{filename_date}")

data = {
    'Baseline IDs': baseline_ids,
    'Schedule IDs': schedule_ids,
    'Schedule Names': schedule_names
}
df = pd.DataFrame(data)
os.makedirs(os.path.join(config.schedule_publish_nomenclature_path, f"{filename_date}_schedules.xlsx"), exist_ok=True)
output_file = os.path.join(config.schedule_publish_nomenclature_path, f"{filename_date}_schedules.xlsx")
df.to_excel(output_file, index=False)
print(f"Schedule information created successfully and saved to {output_file}")