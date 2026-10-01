from datetime import datetime
from datetime import timedelta
import argparse



def determine_next_data_date(current_date_str, monthly_bool):
  if isinstance(monthly_bool, str):
    is_monthly = monthly_bool.strip().lower() in ("true", "t", "1", "y", "yes")
  else:
    is_monthly = bool(monthly_bool)

  if is_monthly:
    return get_monthly_next_data_date(current_date_str)

  return get_default_next_data_date(current_date_str)


def get_monthly_next_data_date(current_date_str):
  date_obj = datetime.strptime(current_date_str, "%m%d%y")

  if date_obj.month == 12:
    next_date_obj = date_obj.replace(year=date_obj.year + 1, month=1, day=1)
  else:
    next_date_obj = date_obj.replace(month=date_obj.month + 1, day=1)

  return next_date_obj.strftime("%m%d%y")

def get_default_next_data_date(current_date_str):
  date_obj = datetime.strptime(current_date_str, "%m%d%y")
  next_date_obj = date_obj + timedelta(days=7)
  return next_date_obj.strftime("%m%d%y")


def main():
  parser = argparse.ArgumentParser(description="Date helper utilities")
  parser.add_argument("--old_data_date", type=str, required=True, help="Previous data date in MMDDYY format")
  parser.add_argument("--monthly_bool", type=str, default="false", help="Set true for monthly next date")
  args = parser.parse_args()

  # Print only the computed date so batch callers can capture it reliably.
  print(determine_next_data_date(args.old_data_date, args.monthly_bool))


if __name__ == "__main__":
  main()