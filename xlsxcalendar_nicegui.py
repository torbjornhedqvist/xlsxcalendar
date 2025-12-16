#!/usr/bin/env python3
"""
NiceGUI frontend for xlsxcalendar configuration and execution
"""
import argparse
import os
import subprocess
import yaml
from nicegui import ui

parser = argparse.ArgumentParser(description='NiceGUI frontend for xlsxcalendar')
parser.add_argument('--root', default='.',
                    help='Path to xlsxcalendar root directory (default: .)')
args = parser.parse_args()

REPO_PATH = args.root
CONFIG_PATH = os.path.join(REPO_PATH, 'xlsxcalendar.yaml')
PROGRAM_PATH = os.path.join(REPO_PATH, 'xlsxcalendar.py')
IMPORTS_PATH = os.path.join(REPO_PATH, 'imports')

TEXTBOX_COLOR = '#FFFAFA'


def load_config():
    """Load current configuration from YAML file or return defaults"""
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as file:
            return yaml.safe_load(file) or {}
    except FileNotFoundError:
        return {}
    except yaml.YAMLError:
        return {}


def save_config(config_data):
    """Save configuration to YAML file"""
    with open(CONFIG_PATH, 'w', encoding='utf-8') as file:
        yaml.dump(config_data, file, default_flow_style=False, sort_keys=False)


def get_available_imports():
    """Scan imports directory for holiday and theme files"""
    holiday_files = []
    theme_files = []

    if os.path.exists(IMPORTS_PATH):
        for file in os.listdir(IMPORTS_PATH):
            if file.endswith('.yaml'):
                full_path = f'./imports/{file}'
                if file.startswith('holidays_'):
                    holiday_files.append(full_path)
                elif file.startswith('theme_'):
                    theme_files.append(full_path)

    return sorted(holiday_files), sorted(theme_files)


@ui.page('/')
def main_page():
    """Main page for xlsxcalendar configuration interface"""
    config = load_config()
    holiday_files, theme_files = get_available_imports()

    ui.label('XlsxCalendar Configuration').classes('text-h4 q-mb-md')

    def close_all_sections():
        """Close all expanded sections"""
        dates_expansion.value = False
        worksheet_expansion.value = False
        content_expansion.value = False
        holidays_expansion.value = False
        themes_expansion.value = False
        import_expansion.value = False
        output_expansion.value = False

    ui.button('Close All Sections', on_click=close_all_sections).classes('q-mb-md')

    with ui.card().classes('w-full max-w-4xl').style('background-color: #F0F8FF'):
        dates_expansion = ui.expansion('Dates', icon='calendar_today').classes('w-full')
        with dates_expansion:
            ui.label('Provide the start and end date of the complete calendar to be '
                     'generated. Note! If you are using an importer module ensure the '
                     'dates are matching.').classes('text-caption text-blue-8 q-mb-sm')
            start_date = ui.input('Start Date',
                                  value=config.get('start_date', '')).props('type=date')
            end_date = ui.input('End Date',
                                value=config.get('end_date', '')).props('type=date')
            ui.separator().style('background-color: #003366; height: 2px;')

        worksheet_expansion = ui.expansion('Worksheet Internal Settings',
                          icon='description').classes('w-full')
        with worksheet_expansion:
            ui.label('Inside the Excel, just the name of the worksheet, defaults to '
                     '- Calendar -').classes('text-caption text-blue-8 q-mb-sm')
            worksheet_name = ui.input('Worksheet Name',
                                      value=config.get('worksheet_name', ''),
                                      placeholder='- Calendar -').classes('w-full').style(
                                          f'font-size: 100%; background-color: {TEXTBOX_COLOR}')
            
            # ui.separator().style('background-color: #A0A0A0; height: 1px;')

            ui.label('And the color of the Worksheet tab. Click on the line to the left of '
                     'the currently configured color to change.').classes(
                     'text-caption text-blue-8 q-mb-sm')
            with ui.row().classes('w-full items-center'):
                ui.label('Worksheet Tab Color:').classes('w-32')
                worksheet_tab_color = ui.input('',
                                               value=config.get('worksheet_tab_color',
                                                                '#ff9966')).props(
                                                   'type=color').classes('w-20')
                color_preview = ui.label('').classes('w-8 h-8 rounded border').style(
                    f'background-color: {config.get("worksheet_tab_color", "#ff9966")}')
                
                worksheet_tab_color.on_value_change(lambda e: color_preview.style(
                    f'background-color: {e.value}'))
                
            ui.separator().style('background-color: #A0A0A0; height: 1px;')
            ui.label('The language of the day of the week.').classes(
                'text-caption text-blue-8 q-mb-sm')
            with ui.row().classes('w-full items-center'):
                ui.label('Day Language:').classes('w-32')
                language = ui.radio(['en', 'sv', 'fi', 'es'],
                                    value=config.get('worksheet_day_of_week_language', 'en'))
            ui.separator().style('background-color: #003366; height: 2px;')

        content_expansion = ui.expansion('Content Settings', icon='edit').classes('w-full')
        with content_expansion:
            ui.label('The heading in the left-most row of the calendar where you could '
                     'put your team name, activity or any other smart heading. '
                     'In Content Entries you can put a list of names or activities to '
                     'be put under the heading. Note! this can be overridden by an '
                     'importer plugin.').classes('text-caption text-blue-8 q-mb-sm')
            content_heading = ui.input('Content Heading',
                                       value=config.get('content_heading', ''),
                                       placeholder='Team/Activity').classes('w-full').style(
                                           f'font-size: 100%; background-color: {TEXTBOX_COLOR}')
            content_entries = ui.textarea('Content Entries (one per line)',
                                          value='\n'.join(config.get('content_entries', []))
                                          if config.get('content_entries') else '').classes(
                                              'w-full').style(
                                              f'background-color: {TEXTBOX_COLOR}')
            ui.separator().style('background-color: #003366; height: 2px;')

        holidays_expansion = ui.expansion('Holidays', icon='event').classes('w-full')
        with holidays_expansion:
            ui.label('Predefined holiday schedules. The dates specified in these files will be '
                     'marked with the weekend color and get a special note below calendar in '
                     'the column for the specific day, indicated by an exclamation character '
                     '[!] with the associated text. Multiple selections possible. '
                     'If there is a conflict in dates with Custom Holidays below, the '
                     'Custom Holidays take precedence.').classes(
                         'text-caption text-blue-8 q-mb-sm')
            ui.label('Holiday Files:')
            holiday_checkboxes = []
            for holiday_file in holiday_files:
                checked = (config.get('holiday_imports') and
                           holiday_file in config.get('holiday_imports', []))
                checkbox = ui.checkbox(holiday_file, value=checked)
                holiday_checkboxes.append((checkbox, holiday_file))

            ui.separator().style('background-color: #A0A0A0; height: 1px;')
            ui.label('Customized holiday days which is not part of any standard holiday template '
                     'from the selection above. It can be any special days or non-business days '
                     'which is happening any day in the week.').classes(
                         'text-caption text-blue-8 q-mb-sm')
            ui.label('Example: 2025-12-11: \'Bob\\\'s birthday\'').classes(
                     'text-caption text-blue-8 q-mb-sm')
            custom_holidays = ui.textarea('Custom Holidays, one per line',
                    value='\n'.join([f'{k}: {v}' for k, v in config.get('holidays',
                        {}).items()])).classes('w-full').style(
                            f'background-color: {TEXTBOX_COLOR}')
            ui.separator().style('background-color: #003366; height: 2px;')

        themes_expansion = ui.expansion('Themes', icon='palette').classes('w-full')
        with themes_expansion:
            ui.label('Various pre-defined color themes for the calendars').classes(
                'text-caption text-blue-8 q-mb-sm')
            theme_radio = ui.radio(['None (default settings)'] + theme_files,
                                   value='None' if not config.get('theme_imports')
                                   else config.get('theme_imports'))
            ui.separator().style('background-color: #003366; height: 2px;')

        import_expansion = ui.expansion('Data Import', icon='upload').classes('w-full')
        with import_expansion:
            ui.label('Plugins which takes an importer file as input and populates the calendar '
                     'with data. Can be highly specialized and you need to read the README '
                     'for the plugin to understand how it is supposed to be used.').classes(
                'text-caption text-blue-8 q-mb-sm')
            ui.label('Plugins')
            importer_module = ui.radio(['None', 'plugins.ess_importer'],
                                       value='None' if not config.get('importer_module')
                                       else config.get('importer_module'))
            with ui.row().classes('w-full items-center gap-2'):
                importer_file = ui.input('Importer File',
                                         value=config.get('importer_file', ''),
                                         placeholder='./tests/ess_test_within_year.csv'
                                         ).classes('flex-grow')

                async def browse_file():
                    # pylint: disable=import-outside-toplevel
                    import tkinter as tk
                    from tkinter import filedialog
                    root = tk.Tk()
                    root.withdraw()
                    file_path = filedialog.askopenfilename(
                        filetypes=[("CSV files", "*.csv"),
                                   ("Excel files", "*.xlsx;*.xls"),
                                   ("All files", "*.*")]
                    )
                    root.destroy()
                    if file_path:
                        importer_file.value = file_path

                ui.button('Browse', on_click=browse_file).classes('w-20')
            ui.separator().style('background-color: #003366; height: 2px;')

        output_expansion = ui.expansion('Output', icon='save').classes('w-full')
        with output_expansion:
            ui.label('Provide the file name of the Excel file which will contain the '
                     'results in Output File below. The file will be placed in your '
                     'Downloads folder.').classes('text-caption text-blue-8 q-mb-sm')
            output_file = ui.input('Output File',
                                   value=config.get('output_file', ''),
                                   placeholder='./output.xlsx').classes('w-full').style(
                                       f'font-size: 100%; background-color: {TEXTBOX_COLOR}')
            ui.separator().style('background-color: #003366; height: 2px;')

    result_label = ui.label('')

    async def generate_calendar():
        """Generate calendar based on configuration"""
        try:
            # Build configuration
            config_data = {
                'start_date': start_date.value,
                'end_date': end_date.value
            }

            if output_file.value:
                config_data['output_file'] = output_file.value
            if worksheet_name.value:
                config_data['worksheet_name'] = worksheet_name.value
            if worksheet_tab_color.value:
                config_data['worksheet_tab_color'] = worksheet_tab_color.value
            if language.value != 'en':
                config_data['worksheet_day_of_week_language'] = language.value
            if content_heading.value:
                config_data['content_heading'] = content_heading.value

            if content_entries.value:
                entries = [e.strip() for e in content_entries.value.split('\n') if e.strip()]
                if entries:
                    config_data['content_entries'] = entries

            # Holiday imports
            selected_holidays = [hf for cb, hf in holiday_checkboxes if cb.value]
            if selected_holidays:
                config_data['holiday_imports'] = selected_holidays

            # Theme
            if theme_radio.value and theme_radio.value != 'None':
                config_data['theme_imports'] = theme_radio.value

            # Custom holidays
            if custom_holidays.value:
                holidays = {}
                for line in custom_holidays.value.split('\n'):
                    if ':' in line:
                        date_key, desc = line.split(':', 1)
                        holidays[date_key.strip()] = desc.strip()
                if holidays:
                    config_data['holidays'] = holidays

            # Importer
            if importer_module.value and importer_module.value != 'None':
                config_data['importer_module'] = importer_module.value
            if importer_file.value:
                config_data['importer_file'] = importer_file.value

            # Save and execute
            save_config(config_data)

            result = subprocess.run([
                'python3', PROGRAM_PATH, '-c', CONFIG_PATH
            ], cwd=REPO_PATH, capture_output=True, text=True, check=False)

            if result.returncode == 0:
                result_label.text = '✅ Calendar generated successfully!'
                result_label.classes('text-green')
                output_path = config_data.get('output_file', './output.xlsx')
                full_path = os.path.join(REPO_PATH, output_path.lstrip('./'))
                if os.path.exists(full_path):
                    ui.download(full_path)
            else:
                result_label.text = f'❌ Error: {result.stderr}'
                result_label.classes('text-red')

        except FileNotFoundError as e:
            result_label.text = f'❌ File Error: {str(e)}'
            result_label.classes('text-red')
        except subprocess.SubprocessError as e:
            result_label.text = f'❌ Subprocess Error: {str(e)}'
            result_label.classes('text-red')

    ui.button('Generate Calendar', on_click=generate_calendar).classes('q-mt-md')


if __name__ in {"__main__", "__mp_main__"}:
    # Remove show=False if you want the local web browser to automatically launch.
    ui.run(host='0.0.0.0', port=8080, title='XlsxCalendar Configuration', show=False)
