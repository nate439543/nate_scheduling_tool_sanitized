import os
import datetime as dt
from pathlib import Path
import subprocess
import gc
import time
import win32com
import win32com.client as win32
import sys
from typing import Any


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_ROOT)
import utilities.ts_print

from openpyxl import load_workbook, Workbook, workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.cell import column_index_from_string, range_boundaries
from openpyxl.utils.datetime import to_excel
from openpyxl.utils import get_column_letter
from win32com.client import Dispatch, DispatchEx



def _build_atomic_temp_path(target_path: Path) -> Path:
	return target_path.with_name(
		f".{target_path.name}.tmp-{os.getpid()}-{int(time.time() * 1000)}"
	)


def _replace_file_atomically(temp_path: Path, target_path: Path) -> None:
	os.replace(temp_path, target_path)

def refresh_data_in_excel_file(file_path: str) -> None:
	workbook_path = Path(file_path)
	if not workbook_path.exists():
		raise FileNotFoundError(f"Excel file not found: {workbook_path}")

	excel = win32.DispatchEx("Excel.Application")
	excel.Visible = False
	excel.DisplayAlerts = False
	excel.EnableEvents = False
	excel.AskToUpdateLinks = False

	try:
		wb = None
		last_error: Exception | None = None
		for attempt in range(1, 4):
			try:
				wb = excel.Workbooks.Open(
					str(workbook_path),
					UpdateLinks=0,
					ReadOnly=False,
					IgnoreReadOnlyRecommended=True,
					AddToMru=False,
				)
				break
			except Exception as exc:
				last_error = exc
				print(
					f"Failed to open Excel workbook '{workbook_path}' "
					f"(attempt {attempt}/3): {exc}"
				)
				gc.collect()
				if attempt < 3:
					time.sleep(1)

		if wb is None:
			raise RuntimeError(f"Unable to open Excel workbook: {workbook_path}") from last_error

		wb.RefreshAll()
		excel.CalculateUntilAsyncQueriesDone()
		wb.Save()
		wb.Close(SaveChanges=False)
		print(f"Data refreshed and saved for Excel file: {workbook_path}")
	finally:
		excel.Quit()

def unlock_excel_sheet(file_path: str, sheet_name: str, password: str) -> None:
	workbook_path = Path(file_path)
	if not workbook_path.exists():
		raise FileNotFoundError(f"Excel file not found: {workbook_path}")

	excel = win32.DispatchEx("Excel.Application")
	excel.Visible = False
	excel.DisplayAlerts = False
	excel.EnableEvents = False

	try:
		wb = excel.Workbooks.Open(
			str(workbook_path),
			UpdateLinks=0,
			ReadOnly=False,
			IgnoreReadOnlyRecommended=True,
			AddToMru=False,
		)

		if sheet_name not in [sheet.Name for sheet in wb.Worksheets]:
			raise ValueError(f"Sheet '{sheet_name}' not found in workbook '{workbook_path}'.")

		ws = wb.Worksheets(sheet_name)
		ws.Unprotect(Password=password)
		wb.Save()
		print(f"Sheet '{sheet_name}' unlocked and saved in Excel file: {workbook_path}")
	except Exception as e:
		print(f"Error unlocking sheet '{sheet_name}' in Excel file '{workbook_path}': {e}")
	finally:
		wb.Close(SaveChanges=False)
		excel.Quit()
  
def lock_excel_sheet(file_path: str, sheet_name: str, password: str) -> None:
	workbook_path = Path(file_path)
	if not workbook_path.exists():
		raise FileNotFoundError(f"Excel file not found: {workbook_path}")

	excel = win32.DispatchEx("Excel.Application")
	excel.Visible = False
	excel.DisplayAlerts = False
	excel.EnableEvents = False

	try:
		wb = excel.Workbooks.Open(
			str(workbook_path),
			UpdateLinks=0,
			ReadOnly=False,
			IgnoreReadOnlyRecommended=True,
			AddToMru=False,
		)

		if sheet_name not in [sheet.Name for sheet in wb.Worksheets]:
			raise ValueError(f"Sheet '{sheet_name}' not found in workbook '{workbook_path}'.")

		ws = wb.Worksheets(sheet_name)
		ws.Protect(Password=password)
		wb.Save()
		print(f"Sheet '{sheet_name}' locked and saved in Excel file: {workbook_path}")
	except Exception as e:
		print(f"Error locking sheet '{sheet_name}' in Excel file '{workbook_path}': {e}")
	finally:
		wb.Close(SaveChanges=False)
		excel.Quit()
  
def set_excel_cell(
	file_path: str,
	sheet_name: str,
	cell_address: str,
	value: str,
	output_path: str | None = None,
	recalculate_after_write: bool = True,
	recalc_visible: bool = False,
	focus_window_for_recalc: bool = False,
	home_sheet: str | None = None,
	home_cell: str | None = None,
) -> Path:
	workbook_path = Path(file_path)
	if not workbook_path.exists():
		raise FileNotFoundError(f"Excel file not found: {workbook_path}")

	if recalculate_after_write:
		return _set_excel_cell_via_excel(
			workbook_path,
			sheet_name,
			cell_address,
			value,
			output_path=output_path,
			visible=recalc_visible,
			focus_window=focus_window_for_recalc,
			home_sheet=home_sheet,
			home_cell=home_cell,
		)

	# Keep VBA project data when editing macro-enabled workbooks.
	keep_vba = workbook_path.suffix.lower() == ".xlsm"
	workbook = load_workbook(workbook_path, keep_vba=keep_vba)

	if sheet_name not in workbook.sheetnames:
		raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {workbook.sheetnames}")

	sheet = workbook[sheet_name]
	sheet[cell_address] = value
	_set_home_view_openpyxl(workbook, home_sheet, home_cell)

	save_path = Path(output_path) if output_path else workbook_path
	_save_workbook_with_retries(workbook, save_path)
	workbook.close()

	return save_path


def _set_home_view_openpyxl(workbook: Any, home_sheet: str | None, home_cell: str | None) -> None:
	if not home_sheet and not home_cell:
		return

	target_sheet = home_sheet or workbook.sheetnames[0]
	if target_sheet not in workbook.sheetnames:
		raise ValueError(f"Sheet '{target_sheet}' not found. Available sheets: {workbook.sheetnames}")

	target_cell = home_cell or "A1"
	home_worksheet = workbook[target_sheet]
	workbook.active = workbook.sheetnames.index(target_sheet)
	if home_worksheet.sheet_view.selection:
		home_worksheet.sheet_view.selection[0].activeCell = target_cell
		home_worksheet.sheet_view.selection[0].sqref = target_cell
	home_worksheet.sheet_view.topLeftCell = target_cell


def _set_excel_cell_via_excel(
	workbook_path: Path,
	sheet_name: str,
	cell_address: str,
	value: str,
	output_path: str | None = None,
	visible: bool = True,
	focus_window: bool = False,
	home_sheet: str | None = None,
	home_cell: str | None = None,
) -> Path:
	excel = None
	workbook = None
	worksheet = None
	pending_replace: tuple[Path, Path] | None = None
	try:
		excel = DispatchEx("Excel.Application")
		excel.Visible = visible
		excel.DisplayAlerts = False
		excel.ScreenUpdating = True
		excel.EnableEvents = True

		workbook = excel.Workbooks.Open(str(workbook_path), ReadOnly=False)
		try:
			workbook.ForceFullCalculation = True
			workbook.FullCalculationOnLoad = True
		except Exception:
			pass
		worksheet_names = [sheet.Name for sheet in workbook.Worksheets]
		if sheet_name not in worksheet_names:
			raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {worksheet_names}")
		if home_sheet and home_sheet not in worksheet_names:
			raise ValueError(f"Sheet '{home_sheet}' not found. Available sheets: {worksheet_names}")

		worksheet = workbook.Worksheets(sheet_name)
		worksheet.Activate()
		if focus_window and visible:
			workbook.Windows(1).Activate()
			excel.WindowState = -4143  # xlNormal
			try:
				excel.Goto(worksheet.Range(cell_address), True)
			except Exception:
				pass

		worksheet.Range(cell_address).Value = value
		save_path = Path(output_path) if output_path else workbook_path

		# Force Excel to register dependent changes the way it would after an interactive edit.
		try:
			excel.Calculate()
		except Exception:
			pass
		excel.CalculateFullRebuild()
		excel.CalculateUntilAsyncQueriesDone()

		target_sheet = home_sheet or sheet_name
		target_cell = home_cell or "A1"
		home_worksheet = workbook.Worksheets(target_sheet)
		home_worksheet.Activate()
		try:
			home_worksheet.Range(target_cell).Select()
		except Exception:
			pass

		temp_save_path = _build_atomic_temp_path(save_path)
		os.makedirs(save_path.parent, exist_ok=True)
		workbook.SaveAs(str(temp_save_path))
		pending_replace = (temp_save_path, save_path)
		return save_path
	finally:
		if workbook is not None:
			workbook.Close(SaveChanges=False)
		if pending_replace is not None:
			temp_path, target_path = pending_replace
			try:
				_replace_file_atomically(temp_path, target_path)
			finally:
				if temp_path.exists():
					try:
						temp_path.unlink()
					except OSError:
						pass
		if excel is not None:
			excel.Quit()


def _value_matches(cell_value: Any, filter_value: Any, case_sensitive: bool) -> bool:
	if isinstance(cell_value, str) and isinstance(filter_value, str):
		if case_sensitive:
			return cell_value == filter_value
		return cell_value.lower() == filter_value.lower()
	return cell_value == filter_value

def add_row_of_data_to_sheet(workbook_path, sheet_name, starting_col: int, data: list[Any], starting_row: int = 1) -> None:
	workbook_path = Path(workbook_path)	
	if not workbook_path.exists():
		raise FileNotFoundError(f"Excel file not found: {workbook_path}")

	keep_vba = workbook_path.suffix.lower() == ".xlsm"
	workbook = load_workbook(workbook_path, keep_vba=keep_vba)

	try:
		if sheet_name not in workbook.sheetnames:
			raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {workbook.sheetnames}")


		sheet = workbook[sheet_name]
		row_index = _find_first_empty_row(sheet, starting_row, starting_col)
		# print(f"Adding data to sheet '{sheet_name}'  at row {row_index}, column {starting_col}: {data}")
		for col_index, value in enumerate(data):
			sheet.cell(row=row_index, column=starting_col + col_index, value=value)
			#print(f"Set cell at row {row_index}, column {starting_col + col_index} to value: {value}")
		_save_workbook_with_retries(workbook, workbook_path)
	finally:
		# print(f"Finished writing data to sheet '{sheet_name}'")
		workbook.close()

def add_range_of_data_to_sheet(workbook_path_str:str, sheet_name:str, starting_col: int, data: list[list[Any]], starting_row: int = 1, create_table: bool = False, table_name: str = "UpdateTable") -> None:
	workbook_path = Path(workbook_path_str)	
	if not workbook_path.exists():
		raise FileNotFoundError(f"Excel file not found: {workbook_path}")

	keep_vba = workbook_path.suffix.lower() == ".xlsm"
	workbook = load_workbook(workbook_path, keep_vba=keep_vba)

	try:
		if sheet_name not in workbook.sheetnames:
			raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {workbook.sheetnames}")

		sheet = workbook[sheet_name]
		row_index = _find_first_empty_row(sheet, starting_row, starting_col)
		print(f"Adding range of data to sheet '{sheet_name}' starting at row {row_index}, column {starting_col}: {data}")
		for row_offset, row_data in enumerate(data):
			for col_offset, value in enumerate(row_data):
				sheet.cell(row=row_index + row_offset, column=starting_col + col_offset, value=value)
				#print(f"Set cell at row {row_index + row_offset}, column {starting_col + col_offset} to value: {value}")
		if create_table:
			end_col_index = starting_col + len(data[0]) - 1
			end_row_index = row_index + len(data) - 1
			table_range = f"{get_column_letter(starting_col)}{row_index}:{get_column_letter(end_col_index)}{end_row_index}"
			print(f"Creating table '{table_name}' in sheet '{sheet_name}' with range: {table_range}")
			if table_name in sheet.tables:
				print(f"Table '{table_name}' already exists. Removing existing table before creating a new one.")
				del sheet.tables[table_name]
			table = Table(displayName="updates", ref=table_range)

			sheet.add_table(table)
		_save_workbook_with_retries(workbook, workbook_path)
	finally:
		print(f"Finished writing range of data to sheet '{sheet_name}'")
		workbook.close()

def add_range_of_data_to_sheet_using_com(workbook_path_str:str, sheet_name:str, starting_col: int, data: list[list[Any]], starting_row: int = 1) -> None:
	
	workbook_path = Path(workbook_path_str)
	if not workbook_path.exists():
		raise FileNotFoundError(f"Excel file not found: {workbook_path}")
	if not data:
		print(f"No data provided for sheet '{sheet_name}'. Skipping COM write.")
		return

	workbook = None
	excel = DispatchEx("Excel.Application")
	excel.Visible = False
	excel.DisplayAlerts = False
	excel.EnableEvents = False

	try:
		workbook = excel.Workbooks.Open(workbook_path, ReadOnly=False)
		worksheet_names = [sheet.Name for sheet in workbook.Worksheets]
		if sheet_name not in worksheet_names:
			raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {worksheet_names}")

		worksheet = workbook.Worksheets(sheet_name)
		row_index = starting_row
		lengths = [len(row) for row in data]
		if len(set(lengths)) != 1:
			raise ValueError(f"All rows must have the same length for COM range writes. Got lengths: {sorted(set(lengths))}")

		column_count = lengths[0]
		start_address = worksheet.Cells(row_index, starting_col).Address
		end_address = worksheet.Cells(row_index + len(data) - 1, starting_col + column_count - 1).Address
		print(f"Writing {len(data)}x{column_count} to {start_address} -> {end_address}")

		# for row_offset, row_data in enumerate(data):
		# 	row_target = worksheet.Range(
		# 		worksheet.Cells(row_index + row_offset, starting_col),
		# 		worksheet.Cells(row_index + row_offset, starting_col + column_count - 1),
		# 	)
		# 	if (row_index + row_offset) % 10 == 0:
		# 		print(f"Writing row {row_index + row_offset} to {row_target.Address}")
		# 	row_target.Value = (tuple(row_data),)
  
		for col_offset in range(column_count):
			col_target = worksheet.Range(
				worksheet.Cells(row_index, starting_col + col_offset),
				worksheet.Cells(row_index + len(data) - 1, starting_col + col_offset),
			)
			

			col_values = tuple((row[col_offset],) for row in data)
			col_target.Value = col_values
		print(f"Finished writing range of data to sheet '{sheet_name}' starting at row {row_index}, column {starting_col}.")
		workbook.Save()
		print(f"Saved workbook after writing range of data to sheet '{sheet_name}'.")

	except Exception as e:
		print(f"Error adding range of data to sheet '{sheet_name}': {e}")
		raise
	finally:
		if workbook is not None:
			workbook.Close(SaveChanges=False)
		if excel is not None:
			excel.Quit()
   

def _parse_column_reference(column_value: str) -> int:
	"""Accept a column letter (e.g. 'C') or 1-based index string (e.g. '3')."""
	raw = str(column_value).strip()
	if raw.isdigit():
		index = int(raw)
		if index < 1:
			raise ValueError(f"Column index must be >= 1, got: {raw}")
		return index

	return column_index_from_string(raw.upper())

def _convert_cell_address_to_indices(cell_address: str) -> tuple[int, int]:
	"""Convert a cell address (e.g., 'B3') to 1-based row and column indices."""
	if not isinstance(cell_address, str) or not cell_address:
		raise ValueError(f"Invalid cell address: {cell_address}")

	column_part = ''.join(filter(str.isalpha, cell_address)).upper()
	row_part = ''.join(filter(str.isdigit, cell_address))

	if not column_part or not row_part:
		raise ValueError(f"Invalid cell address: {cell_address}")

	column_index = column_index_from_string(column_part)
	row_index = int(row_part)

	return row_index, column_index


def _save_workbook_with_retries(workbook: Any, save_path: Path, retries: int = 3, delay_seconds: float = 0.6) -> None:
	"""Save workbook atomically with retries to handle transient Windows file locks."""
	last_error: PermissionError | None = None
	for attempt in range(1, retries + 1):
		temp_path = _build_atomic_temp_path(save_path)
		try:
			workbook.save(temp_path)
			os.makedirs(save_path.parent, exist_ok=True)
			os.replace(temp_path, save_path)
			return
		except PermissionError as e:
			last_error = e
			print(
				f"Permission denied while saving '{save_path}' "
				f"(attempt {attempt}/{retries}). Ensure the file is not open in Excel."
			)
			gc.collect()
			if attempt < retries:
				time.sleep(delay_seconds)
		finally:
			if temp_path.exists():
				try:
					temp_path.unlink()
				except OSError:
					pass

	if last_error is not None:
		raise last_error


def _find_first_empty_row(sheet, start_row: int, column_index: int) -> int:
	row = start_row
	while sheet.cell(row=row, column=column_index).value not in (None, ""):
		row += 1
	return row

def get_sheet_by_path_and_name(workbook_path_str: str, sheet_name: str, data_only: bool = False) -> Any:
	workbook_path = Path(workbook_path_str)	
	if not workbook_path.exists():
		raise FileNotFoundError(f"Excel file not found: {workbook_path}")

	keep_vba = workbook_path.suffix.lower() == ".xlsm"
	workbook = load_workbook(workbook_path, keep_vba=keep_vba, data_only=data_only)

	try:
		if sheet_name not in workbook.sheetnames:
			raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {workbook.sheetnames}")
		print(f"Sheet Found. Returning {workbook[sheet_name]}")
		return workbook[sheet_name]
	finally:
		workbook.close()

def _get_cell_value(sheet, row: int, column_index: int) -> Any:
  try:
    return sheet.cell(row=row, column=column_index).value
  except Exception as e:
    print(f"Error getting cell value at row {row}, column {column_index}: {e}")
    return None


def _convert_datetime_cell_to_excel_number(cell_value: Any) -> Any:
	"""Convert Excel date/time values loaded as Python datetime objects to Excel serial numbers."""
	if isinstance(cell_value, (dt.datetime, dt.date, dt.time, dt.timedelta)):
		return to_excel(cell_value)
	return cell_value

def create_import_file(save_path: str, file_name: str, headers: dict[str, str], userdata: dict[str, str]) -> str:
	try:
		os.makedirs(save_path, exist_ok=True)
		output_path = os.path.join(save_path, file_name)
		print(f"Creating import file at {output_path} with headers: {headers}")
		if os.path.exists(output_path):
			print(f"Import file already exists at {output_path}. Throwing error")
			raise FileExistsError(f"Import file already exists at {output_path}.")

		workbook = Workbook()
		sheet = workbook.active
		if sheet is None:
			sheet = workbook.create_sheet(title="TASK")
		sheet.title = "TASK"
		for cell_ref, header in headers.items():
			row_number, column_number = _convert_cell_address_to_indices(cell_ref)
			sheet.cell(row=row_number, column=column_number, value=header)
		userdata_sheet = workbook.create_sheet(title="USERDATA")
		for cell_ref, value in userdata.items():
			row_number, column_number = _convert_cell_address_to_indices(cell_ref)
			userdata_sheet.cell(row=row_number, column=column_number, value=value)
		workbook.save(output_path)
		workbook.close()
		print(f"Import file created at {output_path}")
		return output_path
	except Exception as e:
		print(f"Error creating import file: {e}")
		raise

def expand_table_in_sheet(file_path: str, sheet_name: str, new_range: str) -> None:
	workbook_path = Path(file_path)
	if not workbook_path.exists():
		raise FileNotFoundError(f"Excel file not found: {workbook_path}")

	keep_vba = workbook_path.suffix.lower() == ".xlsm"
	workbook = load_workbook(workbook_path, keep_vba=keep_vba)

	try:
		if sheet_name not in workbook.sheetnames:
			raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {workbook.sheetnames}")

		sheet = workbook[sheet_name]
		sheet_table_count = len(sheet.tables)
		if sheet_table_count > 1:
			print(f"!\n!\n!\nWarning: Sheet '{sheet_name}' has {sheet_table_count} tables. Expanding all tables to new range '{new_range}'.!\n!\n!\n")
		elif sheet_table_count == 0:
			print(f"!\n!\n!\nWarning: Sheet '{sheet_name}' has no tables. No tables to expand.!\n!\n!\n")
			return
 
		for table in sheet.tables.values():
			print(f"Expanding table '{table.name}' to new range '{new_range}'")
			table.ref = new_range

		_save_workbook_with_retries(workbook, workbook_path)
	finally:
		workbook.close()
def expand_table_in_sheet_using_com(file_path: str, sheet_name: str, new_range: str) -> None:
	workbook_path = Path(file_path)
	if not workbook_path.exists():
		raise FileNotFoundError(f"Excel file not found: {workbook_path}")

	workbook = None
	excel = DispatchEx("Excel.Application")
	excel.Visible = False
	excel.DisplayAlerts = False
	excel.EnableEvents = False

	try:
		workbook = excel.Workbooks.Open(str(workbook_path), ReadOnly=False)
		worksheet_names = [sheet.Name for sheet in workbook.Worksheets]
		if sheet_name not in worksheet_names:
			raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {worksheet_names}")

		worksheet = workbook.Worksheets(sheet_name)
		tables = worksheet.ListObjects
		table_count = tables.Count
		if table_count > 1:
			print(f"!\n!\n!\nWarning: Sheet '{sheet_name}' has {table_count} tables. Expanding all tables to new range '{new_range}'.!\n!\n!\n")
		elif table_count == 0:
			print(f"!\n!\n!\nWarning: Sheet '{sheet_name}' has no tables. No tables to expand.!\n!\n!\n")
			return

		# Use the provided range only for start/end rows.
		# Preserve each table's existing columns to avoid removing formula columns
		# and breaking structured references.
		new_rows_range = worksheet.Range(new_range)
		start_row = new_rows_range.Row
		end_row = start_row + new_rows_range.Rows.Count - 1

		for i in range(1, table_count + 1):
			table = tables.Item(i)
			table_start_col = table.Range.Column
			table_end_col = table_start_col + table.Range.Columns.Count - 1
			resized_range = worksheet.Range(
				worksheet.Cells(start_row, table_start_col),
				worksheet.Cells(end_row, table_end_col),
			)
			print(
				f"Expanding table '{table.Name}' rows to {start_row}:{end_row} "
				f"while preserving columns {table_start_col}:{table_end_col}"
			)
			table.Resize(resized_range)

		workbook.Save()
	finally:
		if workbook is not None:
			workbook.Close(SaveChanges=False)
		if excel is not None:
			excel.Quit()
def copy_base_range_until_special_blank(
	file_path: str,
	base_worksheet: str,
	base_range: str,
 	base_special_column: str,
	new_worksheet: str,
	new_file_path: str,
 	new_column: str = "A",
  front_label_column: str = "",
  back_label_column: str = "",
  
) -> str:
	"""
	Copy values from the first column of base_range downwards while base_special_column has data,
	then append them into new_column on new_worksheet starting at the first empty row.
	"""
	print(f"Copying values from:\n'{base_worksheet}'!{base_range}\n to \n'{new_worksheet}' column {new_column} until special column {base_special_column} is blank.")
	workbook_path = Path(file_path)
	if not workbook_path.exists():
		raise FileNotFoundError(f"Excel file not found: {workbook_path}")
	output_path = Path(new_file_path)
	if not output_path.exists():
		raise FileNotFoundError(f"Output path does not exist: {output_path}")

	keep_vba = workbook_path.suffix.lower() == ".xlsm"
	keep_output_vba = output_path.suffix.lower() == ".xlsm"
	workbook = load_workbook(workbook_path, keep_vba=keep_vba)
	source_values_workbook = load_workbook(workbook_path, keep_vba=keep_vba, data_only=True)
	output_workbook = load_workbook(output_path, keep_vba=keep_output_vba)

	try:
		if base_worksheet not in workbook.sheetnames:
			raise ValueError(f"Base worksheet '{base_worksheet}' not found. Available: {workbook.sheetnames}")
		else:
			print(f"{base_worksheet} found in {workbook_path.name}. Proceeding with copy operation.")
		if new_worksheet not in output_workbook.sheetnames:
			print(f"Creating new worksheet '{new_worksheet}' in output workbook.")
			output_workbook.create_sheet(new_worksheet)
			if new_worksheet not in output_workbook.sheetnames:
				raise ValueError(f"Failed to create new worksheet '{new_worksheet}' in output workbook.")
		else:
			print(f"New worksheet '{new_worksheet}' already exists in output workbook. Appending data to it.")

		source_sheet = workbook[base_worksheet]
		source_values_sheet = source_values_workbook[base_worksheet]
		destination_sheet = output_workbook[new_worksheet]

		min_col, min_row, _max_col, _max_row = range_boundaries(base_range)

		source_column_index = min_col
		base_special_column_index = _parse_column_reference(base_special_column)
		destination_column_index = _parse_column_reference(new_column)

		if min_row is None or min_col is None or _max_row is None or _max_col is None:
			raise ValueError(f"Invalid base range '{base_range}' for worksheet '{base_worksheet}'.")
		if min_row is None or min_col is None or _max_row is None or _max_col is None:
			raise ValueError(f"Invalid base range '{base_range}' for worksheet '{base_worksheet}'.")

		start_row = int(min_row)
		source_column_index = int(min_col)
		base_special_column_index = _parse_column_reference(base_special_column)
		destination_column_index = _parse_column_reference(new_column)
		row = start_row
	
		print(f"Start Row: {start_row}\nSource Column Index: {source_column_index}\nSpecial Column Index: {base_special_column_index}\nDestination Column Index: {destination_column_index}\nSource Sheet Max Row: {source_sheet.max_row}\nDestination Sheet Max Row: {destination_sheet.max_row} \nMax Column: {_max_col}\nMin Column: {min_col}\nMax Row: {_max_row}\nMin Row: {min_row}")
		values_to_copy = []
		while row <= _max_row:
			special_value = source_values_sheet.cell(row=row, column=base_special_column_index).value
			if special_value is None or (isinstance(special_value, str) and special_value.strip() == ""):
				print(f"Value not detected at row {row}, column {base_special_column_index} in '{base_worksheet}'. Stopping copy operation.")
				break
			else:
				if row % 10 == 0:
					print(f"Value detected at row {row}, column {base_special_column_index} in '{base_worksheet}': '{special_value}'. Copying value from column {source_column_index}.")
				
			column = min_col
			this_row_values = []
			if front_label_column != "":
				this_row_values.append(front_label_column)
			while column <= _max_col:
				cell_value = source_values_sheet.cell(row=row, column=column).value
				if row % 10 == 0:
					print(f"Row {row}, Column {column} value: '{cell_value}'")
				this_row_values.append(cell_value)	
				column += 1
			if back_label_column != "":
				this_row_values.append(back_label_column)
			values_to_copy.append(this_row_values)
			row += 1
		if len(values_to_copy) > 0:
			print(f"Copying {len(values_to_copy)} values from '{base_worksheet}'!{base_range} to '{new_worksheet}' column {new_column}.")
		else:
			print(f"No values found to copy from '{base_worksheet}'!{base_range} to '{new_worksheet}' column {new_column}.")

		if not values_to_copy:

			return destination_sheet.calculate_dimension()

		destination_row = _find_first_empty_row(destination_sheet, 1, destination_column_index)
		while destination_sheet.cell(row=destination_row, column=destination_column_index).value not in (None, ""):
			destination_row += 1

		for row_values in values_to_copy:
			for col_index, value in enumerate(row_values, start=destination_column_index):
				destination_sheet.cell(row=destination_row, column=col_index).value = value
			destination_row += 1

		save_path = Path(output_path) 
		_save_workbook_with_retries(output_workbook, save_path)
	
		return destination_sheet.calculate_dimension()
	except Exception as e:
		print(f"Error in copy_base_range_until_special_blank: {str(e)}")
		raise
	finally:
		workbook.close()
		source_values_workbook.close()
		output_workbook.close()


def collect_comments_and_user_input(
  input_sheet_path: str, 
  output_sheet_path: str, 
  data_date: str, 
  update_analysis_ranges: dict[str, str], 
  front_label: str = "", 
  sheet_name_as_back_label: bool = False
) -> None:
	""""""
	for sheet_name, sheet_range in update_analysis_ranges.items():
		print(f"Copying Range {sheet_range} from sheet '{sheet_name}'")
		try:
			if front_label != "":
				front_label_column = front_label
			else:
				front_label_column = ""
			if sheet_name_as_back_label:
				back_label_column = sheet_name
			else:
				back_label_column = ""
			new_sheet_name = sheet_name #+ "_" + data_date
			destination_range = copy_base_range_until_special_blank(input_sheet_path, 
        sheet_name, 
        sheet_range,
        "A", 
        new_sheet_name, 
        output_sheet_path, 
        "A", 
        front_label_column, 
        back_label_column
      )
			print(f"Copied data to '{new_sheet_name}'!{destination_range} in \n{output_sheet_path}")
			#expand_table_in_sheet(output_sheet_path, new_sheet_name, destination_range)
		except Exception as e:
			print(f"Error in collect_comments_and_user_input for sheet '{sheet_name}': {str(e)}")
			print("Continuing execution after logging this error.")

	# Excel open/save/close pass to normalize workbook internals for external data connections.
	_recalculate_workbook_for_cached_values(Path(output_sheet_path))
 
def collect_weekly_update_input_by_project(
	projects: list[str],
	project_designator: str,
	input_sheet_path: str,
	input_sheet_name: str,
	search_column_int: int,
	start_row: int,
	end_col: int = 20,
	start_col: int = 1,
) -> dict[str, list[list[str]]]:
   
	weekly_updates = {}
	try:
		for project in projects:
			other_projects = [project_designator.replace("{project}", p).strip().lower() for p in projects if p != project]
			print(f"Collecting weekly update input for project: {project}")
			weekly_updates[project] = []
			update_sheet = get_sheet_by_path_and_name(input_sheet_path, input_sheet_name, data_only=True)
			project_start_row = start_row
			for row in range(start_row, update_sheet.max_row + 1):
				cell_value = _get_cell_value(update_sheet, row, search_column_int)
				if cell_value is None or str(cell_value).strip() == "":
					print(f"Empty cell detected at row {row}, column {search_column_int}. Stopping data collection for project '{project}'.")
					break
				elif str(cell_value).strip().lower() == project_designator.replace("{project}", project).strip().lower():
					project_start_row = row
					# print(f"Found project '{project}' at row {row}. Starting data collection from this row.")
					for project_row in range(project_start_row, update_sheet.max_row + 1):
						cell_value = _get_cell_value(update_sheet, project_row, search_column_int)
						normalized_value = str(cell_value).strip().lower()
						if normalized_value in other_projects:
							print(f"Encountered next project '{normalized_value}' at row {project_row}. Stopping data collection for project '{project}'.")
							row = project_row - 1  # Adjust the outer loop to continue from the last processed row
							break
						else:
							row_values = []
							for col in range(start_col, end_col + 1):
								cell_value = _get_cell_value(update_sheet, project_row, col)
								row_values.append(_convert_datetime_cell_to_excel_number(cell_value))
							weekly_updates[project].append(row_values)
       
		return weekly_updates
  
	except Exception as e:
    
		print(f"Error in collect_weekly_update_input_by_project: {str(e)}")
		return {}

def collect_updates(
	workbook_path: str,
	sheet_name: str,
	number_of_columns: int
) -> list[list[Any]]:
	"""
	Collect updates from a specified sheet in an Excel workbook.

	- workbook_path: path to the Excel workbook
	- sheet_name: name of the sheet to collect updates from
	- number_of_columns: number of columns to collect from each row

	Returns a list of lists containing the collected data.
	"""
	try:
    
		update_sheet = get_sheet_by_path_and_name(workbook_path, sheet_name, data_only=True)
		max_row = _find_first_empty_row(update_sheet, 1, 1) -1
		print(f"Collecting updates from sheet '{sheet_name}' in workbook '{workbook_path}'. Max row with data: {max_row}. (In Column 1)")
		row = 3
		column = 1
	
		update_data = []
		for row_offset in range(max_row):
			row_values = []
   
			if _get_cell_value(update_sheet, row, column) in (None, ""):
				print(f"Empty cell detected at row {row}, column {column}. Stopping data collection.")
				break
			# print(f"Collecting data from row {row} in sheet '{sheet_name}'")

			for col_offset in range(number_of_columns):
				cell_value = _get_cell_value(update_sheet, row, column + col_offset)
				row_values.append(_convert_datetime_cell_to_excel_number(cell_value))
			
			update_data.append(row_values)
			row += 1
		return update_data
	except Exception as e:
		print(f"Error in collect_updates for sheet '{sheet_name}': {str(e)}")
		return []



def collect_frag_data(
	frag_sheet,
	frag_sheet_ranges: dict[str, str]
	) -> dict[str, str]:
	""""""
	try:
		frag_data_list: dict[str, str] = {}
		for data_type, data_location in frag_sheet_ranges.items():
			col_value = _parse_column_reference(data_location[0])
			frag_data_list[data_type] = _get_cell_value(frag_sheet, int(data_location[1:2].strip()), col_value)
	except Exception as e:
		print(f"Error in collect_frag_data for frag sheet '{frag_sheet.title}': {str(e)}")		
	return frag_data_list

def add_data_to_frag_log(
  frag_log_path: str, 
  frag_log_sheet_name: str,
  frag_data_list: dict[str, str],
  log_number: int,
  frag_log_ranges: dict[str, str]
)-> None:
	""""""
	workbook_path = Path(frag_log_path)
	if not workbook_path.exists():
		raise FileNotFoundError(f"Excel file not found: {workbook_path}")

	keep_vba = workbook_path.suffix.lower() == ".xlsm"
	workbook = load_workbook(workbook_path, keep_vba=keep_vba)

	try:
		if frag_log_sheet_name not in workbook.sheetnames:
			raise ValueError(f"Sheet '{frag_log_sheet_name}' not found. Available sheets: {workbook.sheetnames}")

		sheet = workbook[frag_log_sheet_name]
		first_empty_row = _find_first_empty_row(sheet, 6, 1)
  
		print(f"Adding frag data to '{frag_log_sheet_name}' at row {first_empty_row} with log number {log_number}")

		sheet.cell(row=first_empty_row, column=int(frag_log_ranges["frag_log_number_cell"]), value= str(log_number))
		sheet.cell(row=first_empty_row, column=int(frag_log_ranges["frag_entry_confirmation_cell"]), value= "Y")
		for type, data in frag_data_list.items():
			if type in frag_log_ranges:
				col_index = int(frag_log_ranges[type])

				sheet.cell(row=first_empty_row, column=col_index, value=data)
			else:
				print(f"Warning: Data type '{type}' not found in frag_log_ranges. Skipping this data entry.")

		_save_workbook_with_retries(workbook, workbook_path)
	finally:
		workbook.close()

def remove_rows_containing_value(file_path: str, sheet_name: str, column_letter: str, value_to_remove: str) -> None:
		"""
		Remove rows from the specified sheet in the Excel file where the specified column contains the given value.
		"""
		workbook_path = Path(file_path)
		if not workbook_path.exists():
				raise FileNotFoundError(f"Excel file not found: {workbook_path}")

		keep_vba = workbook_path.suffix.lower() == ".xlsm"
		workbook = load_workbook(workbook_path, keep_vba=keep_vba)

		try:
				if sheet_name not in workbook.sheetnames:
						print(f"Sheet '{sheet_name}' not found. Available sheets: {workbook.sheetnames}. No rows removed.")
						return

				sheet = workbook[sheet_name]
				column_index = column_index_from_string(column_letter)

				rows_to_delete = []
				for row in range(1, sheet.max_row + 1):
						cell_value = sheet.cell(row=row, column=column_index).value
						if cell_value == value_to_remove:
								rows_to_delete.append(row)

				for row in reversed(rows_to_delete):
						sheet.delete_rows(row)

				_save_workbook_with_retries(workbook, workbook_path)
		finally:
				workbook.close()
    
    
def remove_old_sheets(path: str, data_date: str, ranges: dict[str, str]) -> None:
	for sheet_name, sheet_range in ranges.items():
		old_sheet_name = sheet_name + "_" + data_date
		try:
			workbook_path = Path(path)
			if not workbook_path.exists():
				raise FileNotFoundError(f"Excel file not found: {workbook_path}")
			keep_vba = workbook_path.suffix.lower() == ".xlsm"
			workbook = load_workbook(workbook_path, keep_vba=keep_vba)
			if old_sheet_name in workbook.sheetnames:
				print(f"Removing old sheet '{old_sheet_name}' from workbook.")
				del workbook[old_sheet_name]
				_save_workbook_with_retries(workbook, workbook_path)
			else:
				print(f"Sheet '{old_sheet_name}' not found in workbook. No action taken.")
		except Exception as e:
				print(f"Error removing old sheet '{old_sheet_name}': {str(e)}")
		finally:
			if 'workbook' in locals():
				workbook.close()
    

def _rows_to_tsv(headers: list[Any], rows: list[list[Any]], include_headers: bool) -> str:
	def to_text(value: Any) -> str:
		if value is None:
			return ""
		return str(value)

	lines: list[str] = []
	if include_headers:
		lines.append("\t".join(to_text(v) for v in headers))

	for row in rows:
		lines.append("\t".join(to_text(v) for v in row))

	return "\n".join(lines)


def _copy_text_to_clipboard(text: str) -> None:
	if not text:
		return

	if os.name != "nt":
		raise RuntimeError("Clipboard copy is only implemented for Windows in this project.")

	subprocess.run(["clip"], input=text, text=True, check=True)


def _recalculate_workbook_for_cached_values(file_path: Path, visible: bool = False, activate_window: bool = False) -> None:
	excel = None
	workbook = None
	shell = None
	pending_replace: tuple[Path, Path] | None = None
	try:
		excel = DispatchEx("Excel.Application")
		excel.Visible = visible
		excel.DisplayAlerts = False
		excel.ScreenUpdating = True
		workbook = excel.Workbooks.Open(str(file_path), ReadOnly=False)

		if activate_window and visible:
			# Bringing Excel to the foreground can improve calc speed on some systems.
			workbook.Activate()
			excel.WindowState = -4143  # xlNormal
			shell = Dispatch("WScript.Shell")
			shell.AppActivate(excel.Caption)

		excel.CalculateFullRebuild()
		excel.CalculateUntilAsyncQueriesDone()

		temp_save_path = _build_atomic_temp_path(file_path)
		workbook.SaveAs(str(temp_save_path))
		pending_replace = (temp_save_path, file_path)
	finally:
		shell = None
		if workbook is not None:
			workbook.Close(SaveChanges=False)
		if pending_replace is not None:
			temp_path, target_path = pending_replace
			try:
				_replace_file_atomically(temp_path, target_path)
			finally:
				if temp_path.exists():
					try:
						temp_path.unlink()
					except OSError:
						pass
		if excel is not None:
			excel.Quit()


def filter_and_copy_table_rows(
	file_path: str,
	sheet_name: str,
	table_name: str,
	filter_column: str,
	filter_value: Any,
 	columns_copied: int = 16,
	output_sheet_name: str = "FilteredTableData",
	output_path: str | None = None,
	case_sensitive: bool = False,
	clear_output_sheet: bool = True,
	copy_to_clipboard: bool = False,
	include_headers_in_clipboard: bool = True,

) -> Path:
	"""
	Filter rows from a named Excel table and copy matches to an output sheet.

	- file_path: source workbook path (.xlsx or .xlsm)
	- sheet_name: sheet containing the table
	- table_name: Excel table name (not display text)
	- filter_column: header name inside the table used for filtering
	- filter_value: value to match in filter_column
	- output_sheet_name: destination sheet to write filtered data
	- output_path: optional save-as path; defaults to in-place save
	- columns_copied: number of leading table columns to copy to the output
	"""
	workbook_path = Path(file_path)
	if not workbook_path.exists():
		raise FileNotFoundError(f"Excel file not found: {workbook_path}")

	keep_vba = workbook_path.suffix.lower() == ".xlsm"
	workbook = load_workbook(workbook_path, keep_vba=keep_vba)
	values_workbook = load_workbook(workbook_path, keep_vba=keep_vba, data_only=True)

	if sheet_name not in workbook.sheetnames:
		raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {workbook.sheetnames}")
	if sheet_name not in values_workbook.sheetnames:
		raise ValueError(f"Sheet '{sheet_name}' not found in data-only workbook.")

	source_sheet = workbook[sheet_name]
	values_sheet = values_workbook[sheet_name]
	if table_name not in source_sheet.tables:
		raise ValueError(f"Table '{table_name}' not found in sheet '{sheet_name}'.")

	table = source_sheet.tables[table_name]
	min_col, min_row, max_col, max_row = range_boundaries(table.ref) 
	if min_col is None or min_row is None or max_col is None or max_row is None:
		raise ValueError(f"Invalid table range for '{table_name}': {table.ref}")

	if columns_copied < 1:
		raise ValueError("columns_copied must be at least 1.")

	copy_max_col = min(min_col + columns_copied - 1, max_col)


	all_headers = [source_sheet.cell(row=min_row, column=col).value for col in range(min_col, max_col + 1)]
	if filter_column not in all_headers:
		raise ValueError(f"Column '{filter_column}' not found in table '{table_name}'. Headers: {all_headers}")

	filter_col_offset = all_headers.index(filter_column)
	filter_col = min_col + filter_col_offset
	output_headers = [source_sheet.cell(row=min_row, column=col).value for col in range(min_col, copy_max_col + 1)]

	# If the filter column is formula-based and cached results are missing, force an
	# Excel recalc/save so openpyxl data_only values are populated.
	has_formula_in_filter_col = False
	has_cached_filter_value = False
	for row in range(min_row + 1, max_row + 1):
		source_value = source_sheet.cell(row=row, column=filter_col).value
		if isinstance(source_value, str) and source_value.startswith("="):
			has_formula_in_filter_col = True
		cached_value = values_sheet.cell(row=row, column=filter_col).value
		if cached_value is not None:
			has_cached_filter_value = True
		if has_formula_in_filter_col and has_cached_filter_value:
			break

	if has_formula_in_filter_col and not has_cached_filter_value:
		values_workbook.close()
		workbook.close()
		_recalculate_workbook_for_cached_values(workbook_path)
		workbook = load_workbook(workbook_path, keep_vba=keep_vba)
		values_workbook = load_workbook(workbook_path, keep_vba=keep_vba, data_only=True)
		source_sheet = workbook[sheet_name]
		values_sheet = values_workbook[sheet_name]

	filtered_rows: list[list[Any]] = []
	for row in range(min_row + 1, max_row + 1):
		cell_value = values_sheet.cell(row=row, column=filter_col).value
		#print(f"Checking row {row}: filter column value='{cell_value}' against filter_value='{filter_value}'")
		if _value_matches(cell_value, filter_value, case_sensitive):
			if row % 50 == 0:
				print(f"Row {row} matches filter: {cell_value}")
			row_values = [source_sheet.cell(row=row, column=col).value for col in range(min_col, copy_max_col + 1)]
			filtered_rows.append(row_values)

	if output_sheet_name in workbook.sheetnames:
		output_sheet = workbook[output_sheet_name]
		if clear_output_sheet and output_sheet.max_row > 0:
			output_sheet.delete_rows(1, output_sheet.max_row)
	else:
		output_sheet = workbook.create_sheet(output_sheet_name)
		print(f"Created new sheet: {output_sheet_name}")
	output_sheet.append(output_headers)
	for row_values in filtered_rows:
		output_sheet.append(row_values)

	if copy_to_clipboard:
		print(f"Copying {len(filtered_rows)} filtered rows to clipboard with headers included: {include_headers_in_clipboard}")
		clipboard_text = _rows_to_tsv(output_headers, filtered_rows, include_headers_in_clipboard)
		_copy_text_to_clipboard(clipboard_text)
	os.makedirs(workbook_path.parent, exist_ok=True)
	save_path = Path(output_path) if output_path else workbook_path
	if save_path.parent and not save_path.parent.exists():
		print(f"Creating directories for save path: {save_path.parent}")
		save_path.parent.mkdir(parents=True, exist_ok=True)

	_save_workbook_with_retries(workbook, save_path)
	print(f"Saved workbook to: {save_path}")
	return save_path


def run_excel_macro(
	file_path: str,
	macro_name: str,
	save_after_run: bool = True,
	visible: bool = True,
	read_only: bool = False,
) -> Path:
	"""
	Open an Excel workbook and run a VBA macro.

	- file_path: workbook path (.xlsm, .xlsb, .xlsx with macros in add-in, etc.)
	- macro_name: VBA macro name (example: Module1.MyMacro)
	- save_after_run: save workbook after macro completes
	- visible: show Excel window while running
	- read_only: open workbook as read-only
	"""
	workbook_path = Path(file_path)
	if not workbook_path.exists():
		raise FileNotFoundError(f"Excel file not found: {workbook_path}")

	excel = None
	workbook = None
	pending_replace: tuple[Path, Path] | None = None
	try:
		excel = DispatchEx("Excel.Application")
		excel.Visible = visible
		excel.DisplayAlerts = False

		workbook = excel.Workbooks.Open(str(workbook_path), ReadOnly=read_only)

		qualified_macro = f"'{workbook.Name}'!{macro_name}"
		try:
			excel.Application.Run(qualified_macro)
		except Exception:
			# Fallback for callers that already provide a fully qualified macro path.
			excel.Application.Run(macro_name)

		if save_after_run and not read_only:
			temp_save_path = _build_atomic_temp_path(workbook_path)
			workbook.SaveAs(str(temp_save_path))
			pending_replace = (temp_save_path, workbook_path)

		return workbook_path
	finally:
		if workbook is not None:
			workbook.Close(SaveChanges=False)
		if pending_replace is not None:
			temp_path, target_path = pending_replace
			try:
				_replace_file_atomically(temp_path, target_path)
			finally:
				if temp_path.exists():
					try:
						temp_path.unlink()
					except OSError:
						pass
		if excel is not None:
			excel.Quit()


