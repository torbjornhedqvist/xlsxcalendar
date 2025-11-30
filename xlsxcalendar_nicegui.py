#!/usr/bin/env python3
"""
NiceGUI frontend for xlsxcalendar configuration and execution
"""
from nicegui import ui, app
import yaml
import subprocess
import os
from datetime import datetime

CONFIG_PATH = '/home/etorhed/repos/xlsxcalendar/xlsxcalendar.yaml'
PROGRAM_PATH = '/home/etorhed/repos/xlsxcalendar/xlsxcalendar.py'
REPO_PATH = '/home/etorhed/repos/xlsxcalendar'
IMPORTS_PATH = '/home/etorhed/repos/xlsxcalendar/imports'

def load_config():
    """Load current configuration from YAML file or return defaults"""
    try:
        with open(CONFIG_PATH, 'r') as f:
            return yaml.safe_load(f) or {}
    except FileNotFoundError:
        return {}
    except Exception:
        return {}

def save_config(config_data):
    """Save configuration to YAML file"""
    with open(CONFIG_PATH, 'w') as f:
        yaml.dump(config_data, f, default_flow_style=False, sort_keys=False)

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
    config = load_config()
    holiday_files, theme_files = get_available_imports()
    
    ui.label('XlsxCalendar Configuration').classes('text-h4 q-mb-md')
    
    with ui.card().classes('w-full max-w-4xl'):
        with ui.expansion('Basic Settings', icon='settings').classes('w-full'):
            ui.label('Provide the start and end date of the complete calendar to be generated. ' \
            'Note! If you are using an importer module ensure the dates are matching.').\
            classes('text-caption text-blue-8 q-mb-sm')
            start_date = ui.input('Start Date', value=config.get('start_date', '')).props('type=date')
            end_date = ui.input('End Date', value=config.get('end_date', '')).props('type=date')
            ui.separator()
            ui.label('Provide the path and file name of the Excel file which will contain the ' \
            'results in Output File below.').classes('text-caption text-blue-8 q-mb-sm')
            output_file = ui.input('Output File', value=config.get('output_file', ''), placeholder='./output.xlsx')
            ui.separator()
        
        with ui.expansion('Worksheet Settings', icon='description').classes('w-full'):
            ui.label('Information section').classes('text-caption text-grey-6 q-mb-sm')
            worksheet_name = ui.input('Worksheet Name', value=config.get('worksheet_name', ''), placeholder='- Calendar -').classes('w-full')
            with ui.row().classes('w-full items-center'):
                ui.label('Tab Color:').classes('w-32')
                worksheet_tab_color = ui.input('', value=config.get('worksheet_tab_color', '#ff9966')).props('type=color').classes('w-20')
            with ui.row().classes('w-full items-center'):
                ui.label('Day Language:').classes('w-32')
                language = ui.select(['en', 'sv', 'fi', 'es'], 
                                    value=config.get('worksheet_day_of_week_language', 'en')).classes('w-32')
        
        with ui.expansion('Content Settings', icon='edit').classes('w-full'):
            ui.label('Information section').classes('text-caption text-grey-6 q-mb-sm')
            content_heading = ui.input('Content Heading', value=config.get('content_heading', ''), placeholder='Team/Activity')
            content_entries = ui.textarea('Content Entries (one per line)', 
                                        value='\n'.join(config.get('content_entries', [])) if config.get('content_entries') else '')
        
        with ui.expansion('Holidays & Themes', icon='event').classes('w-full'):
            ui.label('Information section').classes('text-caption text-grey-6 q-mb-sm')
            ui.label('Holiday Files:')
            holiday_checkboxes = []
            for holiday_file in holiday_files:
                checked = config.get('holiday_imports') and holiday_file in config.get('holiday_imports', [])
                cb = ui.checkbox(holiday_file, value=checked)
                holiday_checkboxes.append((cb, holiday_file))
            
            theme_options = [''] + theme_files
            theme_select = ui.select(theme_options, 
                                   value=config.get('theme_imports', ''),
                                   label='Theme')
            
            custom_holidays = ui.textarea('Custom Holidays (YYYY-MM-DD: Description)', 
                                        value='\n'.join([f'{k}: {v}' for k, v in config.get('holidays', {}).items()]))
        
        with ui.expansion('Data Import', icon='upload').classes('w-full'):
            ui.label('Information section').classes('text-caption text-grey-6 q-mb-sm')
            importer_module = ui.select(['', 'plugins.template_importer', 'plugins.ess_importer'],
                                      value=config.get('importer_module', ''),
                                      label='Importer Module')
            with ui.row().classes('w-full items-center gap-2'):
                importer_file = ui.input('Importer File', value=config.get('importer_file', ''),
                                       placeholder='./tests/ess_test_within_year.csv').classes('flex-grow')
                
                async def browse_file():
                    import tkinter as tk
                    from tkinter import filedialog
                    root = tk.Tk()
                    root.withdraw()
                    file_path = filedialog.askopenfilename(
                        filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
                    )
                    root.destroy()
                    if file_path:
                        importer_file.value = file_path
                
                ui.button('Browse', on_click=browse_file).classes('w-20')
    
    result_label = ui.label('')
    
    async def generate_calendar():
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
            if theme_select.value:
                config_data['theme_imports'] = theme_select.value
            
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
            if importer_module.value:
                config_data['importer_module'] = importer_module.value
            if importer_file.value:
                config_data['importer_file'] = importer_file.value
            
            # Save and execute
            save_config(config_data)
            
            result = subprocess.run([
                'python3', PROGRAM_PATH, '-c', CONFIG_PATH
            ], cwd=REPO_PATH, capture_output=True, text=True)
            
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
                
        except Exception as e:
            result_label.text = f'❌ Error: {str(e)}'
            result_label.classes('text-red')
    
    ui.button('Generate Calendar', on_click=generate_calendar).classes('q-mt-md')

if __name__ in {"__main__", "__mp_main__"}:
    # Remove show=False if you want the local web browser to automatically launch. 
    ui.run(host='0.0.0.0', port=8080, title='XlsxCalendar Configuration', show=False)
