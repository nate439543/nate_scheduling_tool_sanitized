from nicegui import ui
import os
import subprocess
import sys
from datetime import datetime, date, timedelta
from utilities.get_config import get_config

config = get_config()


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_ROOT)
sys.path.insert(0, SCRIPT_DIR)

from build_p6_update_sheets.main import build_p6_update_sheet
from reorganize_to_onedrive.main import main as reorganize_to_onedrive_main
from get_update_data.main import get_update_data
from collect_updates_to_import_file.main import collect_updates_to_import_file
from frag_logger.main import frag_logger
from reorganize_to_onedrive.main import reorganize_files_to_onedrive


def run_step(step_fn, step_name: str, args=None) -> None:
    """Run a step and surface status in the GUI."""

    print(f'Running step: {step_name}')
    try:
        if args is None:
            step_fn()
        else:
            step_fn(*args)
        ui.notify(f'{step_name} finished.', type='positive')
        print(f'{step_name} finished.')
    except SystemExit as exc:
        if exc.code in (0, None):
            ui.notify(f'{step_name} finished.', type='positive')
            print(f'{step_name} finished successfully.')
        else:
            ui.notify(f'{step_name} exited with code {exc.code}.', type='warning')
            print(f'{step_name} exited with code {exc.code}.')
    except Exception as exc:
        ui.notify(f'{step_name} failed: {exc}', type='negative')
        print(f'{step_name} failed: {exc}')


def run_generate_schedules(date_text: str) -> None:
    script_path = os.path.join(SCRIPT_DIR, 'sched_update_file_org', 'generate_schedules.py')
    subprocess.run([sys.executable, script_path, date_text], check=True)


def open_path(path_to_open) -> None:
    if not os.path.exists(path_to_open):
        ui.notify(f'Folder/File not found: {path_to_open}', type='negative')
        return
    os.startfile(path_to_open)


def get_last_wednesday() -> str:
    today = date.today()
    offset = (today.weekday() - 2) % 7  # Wednesday is 2
    last_wednesday = today - timedelta(days=offset)
    return last_wednesday.strftime('%m/%d/%y')


def get_analysis_wednesday() -> str:
    today = date.today()
    calc_date = today + timedelta(days=-5)
    offset = (calc_date.weekday() - 2) % 7  # Wednesday is 2
    last_wednesday = calc_date - timedelta(days=offset)
    return last_wednesday.strftime('%m/%d/%y')


def _open_ofci_updater() -> None:
    path = config.ofci_updater_path
    if os.path.exists(path):
        open_path(path)
    else:
        ui.notify(f'OFCI Updater not found: {path}', type='negative')

def _open_auto_updater() -> None:
    path = config.auto_updater_path
    if os.path.exists(path):
        open_path(path)
    else:
        ui.notify(f'Auto Updater not found: {path}', type='negative')

def _on_generate_import_files_click() -> None:
    process_start_dd_value = process_start_dd.value if process_start_dd.value else ""
    if process_start_dd_value:
        run_step(
            lambda: collect_updates_to_import_file(process_start_dd_value.replace('/', ''), current_time=datetime.now()),
            'Generate Import Files for This Week\'s Updates',
        )
    else:
        ui.notify('Enter a date in MMDDYY format first.', type='warning')

def _on_generate_publish_nomenclature_click() -> None:
    if process_end_dd.value:
        run_generate_schedules(process_end_dd.value.replace('/', ''))
    else:
        ui.notify('Enter a date in MM/DD/YY format first.', type='warning')

def _open_publish_nomenclature_files() -> None:
    path = config.schedule_publish_nomenclature_path
    if os.path.exists(path):
        open_path(path)
    else:
        ui.notify(f'Publish Nomenclature folder not found: {path}', type='negative')

def _on_move_and_rename_click() -> None:
    if process_end_dd.value:
        reorganize_files_to_onedrive(
            process_end_dd.value.replace('/', '')
        )
    else:
        ui.notify('Enter a date in MM/DD/YY format first.', type='warning')

def _open_import_files() -> None:
    path = config.import_folder_path
    if os.path.exists(path):
        open_path(path)
    else:
        ui.notify(f'Import folder not found: {path}', type='negative')

ui.colors(primary="#368181")

ui.label("Nate's Scheduling Tool").classes('text-h5 q-mb-md w-full text-center')


with ui.tabs().classes('w-full').props('align="justify"') as tabs:
    build_tab = ui.tab('Build')
    analyze_tab = ui.tab('Analyze')
    process_tab = ui.tab('Process')
    


with ui.tab_panels(tabs, value=build_tab).classes('w-full'):

    with ui.tab_panel(build_tab):
        with ui.row().classes('w-full no-wrap gap-4 justify-center'):
            with ui.column().classes('w-1/2 items-center'):
                build_start_dd = (
                    ui.date_input('Starting Data Date', value=get_last_wednesday())
                    .classes('w-88 q-mb-md')
                    .style('background-color: #002C4100; color: white; ')
                )
                build_start_dd.picker.props('mask="MM/DD/YY"')



                def _on_build_p6_click() -> None:
                    start, end = build_start_dd.value, build_end_dd.value
                    if start and end:
                        run_step(
                            lambda: build_p6_update_sheet(start.replace('/', ''), end.replace('/', '')),
                            'Build P6 Update Sheets',
                        )
                    else:
                        ui.notify('Enter both start and end dates in MMDDYY format first.', type='warning')

                ui.button(
                    'Build P6 Update Sheets',
                    on_click=_on_build_p6_click,
                ).style('color: white !important; background-color:#770031  !important;')

                ui.button(
                    'Open P6 Data Source File',
                    on_click=lambda: open_path(config.p6_data_path),
                ).style('color: white !important; background-color: #6PROJ55f !important;')
                
                ui.button(
                    'Open Build Update Sheets Output Folder',
                    on_click=lambda: open_path(config.output_folder_base_path),
                ).style('color: white !important; background-color: #6PROJ55f !important;')

            with ui.column().classes('w-1/2 items-center'):
                
                build_end_default = (
                    datetime.strptime(get_last_wednesday(), '%m/%d/%y') + timedelta(days=7)
                ).strftime('%m/%d/%y')
                build_end_dd = (
                    ui.date_input('Ending Data Date', value=build_end_default)
                    .classes('w-88 q-mb-md')
                    .style('background-color: #88888900; color: white; ')
                )
                build_end_dd.picker.props('mask="MM/DD/YY"')
                ui.button(
                    'Open Frag Template',
                    on_click=lambda: open_path(config.fragnet_template_path),
                ).style('color: white !important; background-color: #0d474a !important;')
                

                
                ui.button(
                    'Open Excel Update Template',
                    on_click=lambda: open_path(config.template_path),
                ).style('color: white !important; background-color: #0d474a !important;')

    with ui.tab_panel(analyze_tab):
        with ui.row().classes('w-full no-wrap gap-4 justify-center'):
            with ui.column().classes('w-1/2 items-center'):
                analyze_start_dd = (
                    ui.date_input('Starting Data Date', value=get_analysis_wednesday())
                    .classes('w-88 q-mb-md')
                    .style('background-color: #002C4100; color: white; ')
                )
                analyze_start_dd.picker.props('mask="MM/DD/YY"')

                def _on_grab_update_data_click() -> None:

                    analyze_start_dd_value = analyze_start_dd.value if analyze_start_dd.value else ""
                    if analyze_start_dd.value:
                        print(f"Grabbing update data for analysis starting from {analyze_start_dd_value}")
                        run_step(
                            lambda: get_update_data(analyze_start_dd_value.replace('/', '')),
                            'Grab Update Data For Analysis',
                        )
                    else:
                        ui.notify('Enter a date in MMDDYY format first.', type='warning')

                ui.button(
                    'Grab Update Data For Analysis',
                    on_click=_on_grab_update_data_click,
                ).style('color: white !important; background-color: #770031 !important;')

            with ui.column().classes('w-1/2 items-center'):
                def _open_update_analysis_source() -> None:
                    path = config.update_analysis_source_path
                    if os.path.exists(path):
                        open_path(path)
                    else:
                        ui.notify(f'Update Analysis Source file not found: {path}', type='negative')

                ui.button(
                    'Open Activities Riding Date',
                    on_click=lambda: open_path(config.activities_riding_dates_path),
                ).style('color: white !important; background-color: #6PROJ55f !important;')
                ui.button(
                    'Open Update Analysis File',
                    on_click=lambda: open_path(config.update_analysis_path),
                ).style('color: white !important; background-color: #6PROJ55f !important;')
                ui.button(
                    'Open Comment Analysis File',
                    on_click=lambda: open_path(config.comment_analysis_path),
                ).style('color: white !important; background-color: #6PROJ55f !important;')
                ui.button(
                    'Open Update Analysis Source File',
                    on_click=_open_update_analysis_source,
                ).style('color: white !important; background-color: #193153 !important;')
    with ui.tab_panel(process_tab):
        with ui.row().classes('w-full no-wrap gap-4 justify-center'):
            with ui.column().classes('w-1/2 items-center'):
                process_start_default = (
                    datetime.strptime(get_last_wednesday(), '%m/%d/%y') - timedelta(days=7)
                ).strftime('%m/%d/%y')
                process_start_dd = (
                    ui.date_input('Starting Data Date (Generate Import)', value=process_start_default)
                    .classes('w-88 q-mb-md')
                    .style('background-color: #002C4100; color: white; ')
                )
                process_start_dd.picker.props('mask="MM/DD/YY"')



                ui.button(
                    'Generate Import Files for Current Updates',
                    on_click=lambda: run_step(
                        _on_generate_import_files_click,
                        'Generate Import Files for Current Updates',
                    ),
                ).style('color: white !important; background-color: #770031 !important;')

                ui.button(
                    'Open OFCI Updater',
                    on_click=lambda: run_step(_open_ofci_updater, 'Open OFCI Updater'),
                ).style('color: white !important; background-color: #6PROJ55f !important;')
                ui.button(
                    'Open Auto Updater',
                    on_click=lambda: run_step(_open_auto_updater, 'Open Auto Updater'),
                ).style('color: white !important; background-color: #6PROJ55f !important;')
                ui.button(
                    'Open Import Files',
                    on_click=lambda: run_step(_open_import_files, 'Open Import Files'),
                ).style('color: white !important; background-color: #0d474a !important;')
                # ui.button(
                #     'Log Frags',
                #     on_click=lambda: run_step(
                #         lambda: frag_logger(process_start_dd.value.replace('/', ''))
                #         if process_start_dd.value
                #         else ui.notify('Enter a date in MM/DD/YY format first.', type='warning'),
                #         'Log Frags',
                #     ),
                # ).style('color: white !important; background-color: #002C41 !important;')

            with ui.column().classes('w-1/2 items-center'):
                process_end_default = datetime.strptime(get_last_wednesday(), '%m/%d/%y').strftime('%m/%d/%y')
                process_end_dd = (
                    ui.date_input(
                        'Ending Data Date (Move and Rename / Nomenclature)',
                        value=process_end_default,
                    )
                    .classes('w-88 q-mb-md')
                    .style('background-color: #002C4100; color: white; ')
                )
                process_end_dd.picker.props('mask="MM/DD/YY"')
                
                ui.button(
                    'Generate Publish Nomenclature',
                    on_click=lambda: run_step(
                        _on_generate_publish_nomenclature_click,
                        'Generate Publish Nomenclature',
                    ),
                ).style('color: white !important; background-color: #770031 !important;')
                
                ui.button(
                    'Open Publish Nomenclature Files',
                    on_click=lambda: run_step(
                        _open_publish_nomenclature_files,
                        'Open Publish Nomenclature Files',
                    ),
                ).style('color: white !important; background-color: #0d474a !important;')
                                
                ui.button(
                    'Move and Rename Publish Files',
                    on_click=lambda: run_step(_on_move_and_rename_click, 'Move and Rename Publish Files'),
                ).style('color: white !important; background-color: #770031 !important;')

                ui.button(
                    'Open Publish File Input Folder',
                    on_click=lambda: open_path(config.publish_files_input_path),
                ).style('color: white !important; background-color: #0d474a !important;')

with ui.row().classes('w-full gap-4 justify-center q-mt-lg'):
    ui.button(
        'Open Local Tool Files',
        on_click=lambda: open_path(config.local_tool_path),
    ).style('color: white !important; background-color: #0d474a !important;')
    
    ui.button(
        'Open OneDrive Tool Files',
        on_click=lambda: open_path(config.schedule_tool_files_path),
    ).style('color: white !important; background-color: #6PROJ55f !important;')
    
    ui.button(
        'Open OneDrive Publish Folder',
        on_click=lambda: open_path(config.publish_parent_folder_path),
    ).style('color: white !important; background-color: #193153 !important;')




ui.run(native=True, title='Nate\'s Scheduling Tool', dark=True, reload=False, favicon= "app/favicon.ico")