#!/usr/bin/env python3
"""
Multi-user web service for xlsxcalendar
Provides session isolation, authentication, and secure file handling
"""
import hashlib
import logging
import os
import secrets
import shutil
import subprocess
import sys
import tempfile
from contextlib import contextmanager, suppress
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

import yaml
from nicegui import ui, app  # pylint: disable=import-error

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Configuration
UPLOAD_MAX_SIZE = 10 * 1024 * 1024  # 10MB
SESSION_TIMEOUT = 3600  # 1 hour
MAX_CONCURRENT_USERS = 50
TEXTBOX_COLOR = '#FFFAFA'

class Language(Enum):
    """Supported languages for calendar."""
    ENGLISH = "en"
    SWEDISH = "sv"
    FINNISH = "fi"
    SPANISH = "es"

class FileExtension(Enum):
    """Allowed file extensions."""
    YAML = ".yaml"
    YML = ".yml"

@dataclass
class CalendarConfig:  # pylint: disable=too-many-instance-attributes
    """Configuration for calendar generation."""
    start_date: str
    end_date: str
    output_file: str
    worksheet_name: str = "- Calendar -"
    worksheet_tab_color: str = "#ff9966"
    worksheet_day_of_week_language: str = "en"
    content_heading: str = ""
    content_entries: List[str] = field(default_factory=list)
    holiday_imports: List[str] = field(default_factory=list)
    holidays: Dict[str, str] = field(default_factory=dict)
    theme_imports: Optional[str] = None
    importer_module: Optional[str] = None
    importer_file: str = ""

class XlsxCalendarError(Exception):
    """Base exception for xlsxcalendar errors."""

class SessionLimitExceededError(XlsxCalendarError):
    """Raised when maximum concurrent users is reached."""

class AuthenticationError(XlsxCalendarError):
    """Raised when authentication fails."""

class UserSession:
    """Manages individual user session with isolated temporary directory."""
    __slots__ = ('session_id', 'created_at', 'last_access', 'temp_dir',
                 'config_path', 'output_path')
    def __init__(self, session_id: str):
        """Initialize user session with unique temporary directory."""
        self.session_id = session_id
        self.created_at = datetime.now()
        self.last_access = datetime.now()
        self.temp_dir = Path(tempfile.mkdtemp(prefix=f'xlsxcal_{session_id}_'))
        self.config_path = self.temp_dir / 'config.yaml'
        self.output_path = self.temp_dir / 'output.xlsx'

    @property
    def is_expired(self) -> bool:
        """Check if session has expired (older than 1 hour)."""
        return (datetime.now() - self.last_access).total_seconds() > SESSION_TIMEOUT

    def cleanup(self) -> None:
        """Clean up temporary directory and files."""
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir, ignore_errors=True)

class SessionManager:
    """Manages user sessions, API keys, and authentication."""
    def __init__(self):
        self.sessions: Dict[str, UserSession] = {}
        self.api_keys: Set[str] = set()
        self.api_keys_file = Path('/app/keys/xlsxcalendar_api_keys')
        self.api_keys_plaintext_file = Path('/app/keys/xlsxcalendar_api_keys.txt')
        self._load_api_keys()

    def _load_api_keys(self) -> None:
        """Load API keys from environment or persistent file"""
        # Load from environment
        key = os.getenv('XLSXCAL_API_KEY')
        if key:
            self.api_keys.add(hashlib.sha256(key.encode()).hexdigest())

        # Load from persistent file
        if self.api_keys_file.exists():
            with suppress(FileNotFoundError, IOError):
                content = self.api_keys_file.read_text(encoding='utf-8')
                self.api_keys.update(line.strip() for line in content.splitlines() if line.strip())

    def _save_api_keys(self) -> None:
        """Save API keys to persistent file"""
        with suppress(OSError, IOError):
            self.api_keys_file.parent.mkdir(parents=True, exist_ok=True)
            self.api_keys_file.write_text('\n'.join(self.api_keys) + '\n', encoding='utf-8')

    def _save_plaintext_key(self, api_key: str) -> None:
        """Save plaintext API key for user reference"""
        with suppress(OSError, IOError):
            self.api_keys_plaintext_file.parent.mkdir(parents=True, exist_ok=True)
            with self.api_keys_plaintext_file.open('a', encoding='utf-8') as txt_file:
                txt_file.write(f"{api_key}\n")

    def add_api_key(self, api_key: str) -> str:
        """Add new API key and persist it"""
        key_hash = hashlib.sha256(api_key.encode()).hexdigest()
        self.api_keys.add(key_hash)
        self._save_api_keys()
        self._save_plaintext_key(api_key)
        return api_key

    def authenticate(self, api_key: str) -> bool:
        """Authenticate user with API key."""
        if not api_key:
            return False
        key_hash = hashlib.sha256(api_key.encode()).hexdigest()
        return key_hash in self.api_keys

    def create_session(self) -> str:
        """Create new user session with resource limits."""
        if len(self.sessions) >= MAX_CONCURRENT_USERS:
            self._cleanup_expired()
            if len(self.sessions) >= MAX_CONCURRENT_USERS:
                raise SessionLimitExceededError("Maximum concurrent users reached")

        session_id = secrets.token_urlsafe(32)
        self.sessions[session_id] = UserSession(session_id)
        return session_id

    def get_session(self, session_id: str) -> Optional[UserSession]:
        """Get or create user session."""
        session = self.sessions.get(session_id)
        if session and not session.is_expired:
            session.last_access = datetime.now()
            return session
        if session:
            session.cleanup()
            del self.sessions[session_id]
        return None

    def _cleanup_expired(self) -> None:
        """Clean up expired sessions."""
        expired_sessions = {
            sid: session for sid, session in self.sessions.items()
            if session.is_expired
        }
        for sid, session in expired_sessions.items():
            session.cleanup()
            del self.sessions[sid]

session_manager = SessionManager()

@contextmanager
def secure_session():
    """Context manager for secure session handling"""
    session_id = app.storage.user.get('session_id')
    if not session_id:
        session_id = session_manager.create_session()
        app.storage.user['session_id'] = session_id

    session = session_manager.get_session(session_id)
    if not session:
        session_id = session_manager.create_session()
        app.storage.user['session_id'] = session_id
        session = session_manager.get_session(session_id)

    try:
        yield session
    except Exception as e:
        ui.notify(f'Error: {str(e)}', type='negative')
        raise

def validate_config(config_data: dict) -> bool:
    """Validate configuration data"""
    required_fields = ['start_date', 'end_date']
    for required_field in required_fields:
        if required_field not in config_data:
            return False

    # Validate date format
    try:
        datetime.strptime(config_data['start_date'], '%Y-%m-%d')
        datetime.strptime(config_data['end_date'], '%Y-%m-%d')
    except ValueError:
        return False

    return True

def sanitize_filename(filename: str) -> str:
    """Sanitize filename to prevent path traversal"""
    return "".join(c for c in filename if c.isalnum() or c in (' ', '.', '_', '-')).strip()

def get_available_imports(_session: UserSession) -> Tuple[List[str], List[str]]:
    """Scan imports directory for holiday and theme files"""
    holiday_files = []
    theme_files = []

    imports_path = Path('/xlsxcalendar/imports')
    if imports_path.exists():
        for file_path in imports_path.iterdir():
            if file_path.suffix in {FileExtension.YAML.value, FileExtension.YML.value}:
                relative_path = f"./imports/{file_path.name}"
                if file_path.name.startswith('holidays_'):
                    holiday_files.append(relative_path)
                elif file_path.name.startswith('theme_'):
                    theme_files.append(relative_path)

    return sorted(holiday_files), sorted(theme_files)

def load_session_config(session: UserSession) -> Dict:
    """Load current configuration from session file or return defaults"""
    if session.config_path.exists():
        with suppress(FileNotFoundError, IOError, yaml.YAMLError):
            return yaml.safe_load(session.config_path.read_text(encoding='utf-8')) or {}
    return {}

def save_session_config(session: UserSession, config_data: Dict) -> None:
    """Save configuration to session file"""
    with suppress(FileNotFoundError, IOError, yaml.YAMLError):
        session.config_path.write_text(
            yaml.dump(config_data, default_flow_style=False, sort_keys=False),
            encoding='utf-8'
        )

@ui.page('/')
def main_page():  # pylint: disable=too-many-locals,too-many-statements
    """Main application page with authentication"""

    # Simple API key authentication
    if not app.storage.user.get('authenticated'):
        with ui.card().classes('w-96 mx-auto mt-20'):
            ui.label('XlsxCalendar Multi-User Service').classes('text-2xl mb-4')
            api_key_input = ui.input('API Key', password=True).classes('w-full')

            def authenticate():
                if session_manager.authenticate(api_key_input.value):
                    app.storage.user['authenticated'] = True
                    ui.navigate.reload()
                else:
                    ui.notify('Invalid API key', type='negative')

            ui.button('Login', on_click=authenticate).classes('w-full')
        return

    # Main application UI - exact copy from xlsxcalendar_nicegui.py
    with secure_session() as session:
        config = load_session_config(session)
        holiday_files, theme_files = get_available_imports(session)

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
                                              f'font-size: 100%; '
                                              f'background-color: {TEXTBOX_COLOR}')

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
                                               f'font-size: 100%; '
                                               f'background-color: {TEXTBOX_COLOR}')
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
                ui.label('Customized holiday days which is not part of any standard '
                         'holiday template from the selection above. It can be any special '
                         'days or non-business days which is happening any day in the week.'
                         ).classes('text-caption text-blue-8 q-mb-sm')
                ui.label('Example: 2025-12-11: \'Bob\\\'s birthday\''
                         ).classes('text-caption text-blue-8 q-mb-sm')
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
                ui.separator().style('background-color: #A0A0A0; height: 1px;')
                ui.label('If you run the web gui in a docker container make sure you map the path '
                         'for the importer file to match the configured mount point for the '
                         'container.').classes('text-caption text-blue-8 q-mb-sm')
                importer_file = ui.input('Importer File',
                                         value=config.get('importer_file', ''),
                                         placeholder='/data/input.xlsx .csv or appropriate '
                                         'type for your plugin').classes('w-full').style(
                                             f'font-size: 100%; background-color: {TEXTBOX_COLOR}')
                ui.separator().style('background-color: #003366; height: 2px;')

            output_expansion = ui.expansion('Output', icon='save').classes('w-full')
            with output_expansion:
                ui.label('Provide the file name of the Excel file which will contain the '
                         'results in Output File below. The file will be placed in your '
                         'Downloads folder.').classes('text-caption text-blue-8 q-mb-sm')
                output_file = ui.input('Output File',
                                       value=config.get('output_file', ''),
                                       placeholder='output.xlsx').classes('w-full').style(
                                           f'font-size: 100%; background-color: {TEXTBOX_COLOR}')
                ui.separator().style('background-color: #003366; height: 2px;')

        result_label = ui.label('')

        async def generate_calendar():  # pylint: disable=too-many-branches
            """Generate calendar based on configuration"""
            try:
                # Build configuration - exact copy from original
                config_data = {
                    'start_date': start_date.value,
                    'end_date': end_date.value,
                    'output_file': str(session.output_path)  # Use session output path
                }

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
                if theme_radio.value and theme_radio.value != 'None (default settings)':
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
                save_session_config(session, config_data)

                result = subprocess.run([
                    'python3', '/xlsxcalendar/xlsxcalendar.py', '-c', str(session.config_path)
                ], capture_output=True, text=True, timeout=30, check=False)

                if result.returncode == 0 and session.output_path.exists():
                    result_label.text = '✅ Calendar generated successfully!'
                    result_label.classes('text-green')

                    # Provide download with original filename or default
                    download_filename = output_file.value or 'calendar.xlsx'
                    safe_filename = sanitize_filename(download_filename)
                    ui.download(str(session.output_path), filename=safe_filename)
                else:
                    error_msg = result.stderr or 'Unknown error occurred'
                    result_label.text = f'❌ Error: {error_msg[:200]}'
                    result_label.classes('text-red')

            except subprocess.TimeoutExpired:
                result_label.text = '❌ Error: Process timeout'
                result_label.classes('text-red')
            except (subprocess.SubprocessError, OSError) as e:
                result_label.text = f'❌ Error: {str(e)[:200]}'
                result_label.classes('text-red')

        ui.button('Generate Calendar', on_click=generate_calendar).classes('q-mt-md')

if __name__ == '__main__':
    # Check for --generate-key flag
    if '--generate-key' in sys.argv:
        # Create a separate session manager instance for key generation
        key_manager = SessionManager()
        new_key = secrets.token_urlsafe(32)
        key_manager.add_api_key(new_key)
        logger.info("Generated new API key: %s", new_key)
        logger.info("Key has been saved and will persist across restarts")
        sys.exit(0)

    # Load existing keys or generate first one
    if not session_manager.api_keys:
        default_key = secrets.token_urlsafe(32)
        session_manager.add_api_key(default_key)
        logger.info("Generated initial API key: %s", default_key)
        logger.info("This key will persist across restarts")
    else:
        logger.info("Using existing API keys from previous sessions")
        # Show existing keys from plaintext file
        plaintext_file = Path('/app/keys/xlsxcalendar_api_keys.txt')
        if plaintext_file.exists():
            with suppress(FileNotFoundError, IOError):
                keys = [line.strip() for line in
                       plaintext_file.read_text(encoding='utf-8').splitlines()
                       if line.strip()]
                if keys:
                    logger.info("Available API keys: %s", ', '.join(keys))
        logger.info("Use --generate-key flag to create additional keys")

    # Generate storage secret for sessions
    storage_secret = os.getenv('STORAGE_SECRET', secrets.token_urlsafe(32))

    ui.run(
        host='0.0.0.0',
        port=8080,
        title='XlsxCalendar Multi-User Service',
        show=False,
        reload=False,
        storage_secret=storage_secret
    )
