#!/usr/bin/env python3
"""
Copyright (C) 2022 Torbjorn Hedqvist
All Rights Reserved You may use, distribute and modify this code under the
terms of the MIT license. See LICENSE file in the project root for full
license information.
"""
import os
import logging
import importlib
from datetime import date
import yaml

log = logging.getLogger(__name__)
class Config:
    """A class that holds all configuration loaded from the YAML configuration file
    or set default values for static or optional configuration parameters.
    """
    # pylint: disable=too-many-instance-attributes
    # pylint: disable=too-many-branches
    # Reasonable amount in this kind of utility class.


    def __init__(self):
        """Constructor.
        """

        # Defaults, might be overridden later by configuration file.
        self._filename = 'dummy_x4zt-Bl' # Dummy filename to avoid loading a config file
        self._log_level = 'INFO'
        self._start_date = None
        self._end_date = None
        self._output_file = './output.xlsx'
        self._worksheet_name = '- Calendar -'
        self._worksheet_tab_color = '#ff9966'
        self._week_days = ["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]
        self._content_heading = 'Title/Heading'
        self._content_col_size = len(self._content_heading)
        self._content_entries = []
        self._content_num_rows = 10 # Empty content rows if nothing else is provided
        self._theme_imports = None
        self._holidays_imports = {}
        self._holidays = {}
        self._importer_module = None
        self._importer_file = None
        self._importer = None

        # cell_format default is set to None as it has it's defaults in cell_format.py,
        # and we only change that if a configuration file overrides it later.
        self._cell_formats = None
        self._cell_formats_day = None
        self._cell_formats_weekend = None
        self._cell_formats_week_odd = None
        self._cell_formats_week_even = None
        self._cell_formats_month_odd = None
        self._cell_formats_month_even = None
        self._cell_formats_year_odd = None
        self._cell_formats_year_even = None
        self._cell_formats_content_heading = None

        # Static config which should not be tampered with.
        self._start_col = 1 # "Sheet Column B"

        # These are also static config but more safe to tamper with.
        self._year_row = 2  # "Sheet Row 3"
        self._month_row = self._year_row + 1
        self._week_row = self._year_row + 2
        self._day_of_week_row = self._year_row + 3
        self._day_row = self._year_row + 4

    def load_config(self, args: dict) -> bool:
        """Try to load a config file if it exists.
        If there is no config file, first check if the default output filename and path is
        overridden from command line and set it accordingly and then test if the minimal 
        requirement of start_date and end_date is provided, it this is not met return False.
        """

        if args.get('config_file') is not None:
            self._filename = args.get('config_file')
        try:
            with open(self._filename, 'r', encoding='utf-8') as file:
                loaded_config = yaml.load(file, Loader=yaml.FullLoader)

            if args.get('start_date') is not None and args.get('end_date') is not None:
                log.debug('Got dates from command line args')
                self._start_date = self.__str_to_date(args.get('start_date'))
                self._end_date = self.__str_to_date(args.get('end_date'))
            elif 'start_date' in loaded_config and 'end_date' in loaded_config:
                self._start_date = self.__str_to_date(loaded_config.get('start_date'))
                self._end_date = self.__str_to_date(loaded_config.get('end_date'))
            else:
                log.error('Start and/or end date missing, must be provided either in config,')
                log.error("or as command line arguments.")
                return False

            if loaded_config.get('output_file') is not None:
                self._output_file = loaded_config.get('output_file')

            # Override above if an output file has been provided as a command line argument
            if args.get('output_file') is not None:
                log.debug('output_file "%s" provided from command line args, override '
                          'configuration file settings.', args.get('output_file'))
                self._output_file = args.get('output_file')

            if loaded_config.get('worksheet_name') is not None:
                self._worksheet_name = loaded_config.get('worksheet_name')

            if loaded_config.get('worksheet_tab_color') is not None:
                self._worksheet_tab_color = loaded_config.get('worksheet_tab_color')

            # Check for supported week day language in ISO 639 two letter format
            if loaded_config.get('worksheet_day_of_week_language') is not None:
                supported_languages = ['en', 'sv', 'es', 'fi']
                language = loaded_config.get('worksheet_day_of_week_language')
                if language in supported_languages:
                    if language == 'en':
                        pass # Default, do nothing
                    elif language == 'sv':
                        self._week_days = ["Må", "Ti", "On", "To", "Fr", "Lö", "Sö"]
                    elif language == 'es':
                        self._week_days = ["Lu", "Ma", "Mi", "Ju", "Vi", "Sá", "Do"]
                    else: # Only finnish remains
                        self._week_days = ["Ma", "Ti", "Ke", "To", "Pe", "La", "Su"]
                else:
                    log.error('Not a supported language: %s', language)
                    return False

            if loaded_config.get('content_heading') is not None:
                self._content_heading = loaded_config.get('content_heading')
                self._content_col_size = len(self._content_heading)

            if loaded_config.get('content_entries') is not None:
                self._content_entries = loaded_config.get('content_entries')
                self._content_num_rows = len(self._content_entries)
                for entry in self._content_entries:
                    if len(entry) > self._content_col_size:
                        # This entry have more characters than the previous, update.
                        self._content_col_size = len(entry)

            # Handling of external themes imports if they are configured.
            # These will be overridden if additional cell_formats configuration is done
            # in the central configuration.
            self._theme_imports = loaded_config.get('theme_imports')
            if self._theme_imports is not None:
                try:
                    with open(self._theme_imports, 'r', encoding='utf-8') as file:
                        config = yaml.load(file, Loader=yaml.FullLoader)
                        self.update_cell_formats(config.get('cell_formats'))
                except IOError as error:
                    print(error, "Abandon theme imports, please fix the error.")
                    self._cell_formats = None

            # Now check the central configuration and override if needed.
            self.update_cell_formats(loaded_config.get('cell_formats'))

            # Handling of external holiday import files if they are configured.
            # The content of all imported files will be merged into the
            # self._holidays dict attribute
            self._holidays_imports = loaded_config.get('holiday_imports')
            if self._holidays_imports is not None:
                try:
                    for filename in self._holidays_imports:
                        with open(filename, 'r', encoding='utf-8') as file:
                            config = yaml.load(file, Loader=yaml.FullLoader)
                            holidays = config.get('holidays')
                            self._holidays.update(holidays)
                except IOError as error:
                    print(error, "Abandon all holiday imports, please fix the error.")
                    self._holidays = {}

            # And now add (merge) the configuration file's local holidays.
            if loaded_config.get('holidays') is not None:
                self._holidays.update(loaded_config.get('holidays'))

            # Check for importer plugins and files
            self._importer_module = loaded_config.get('importer_module')
            self._importer_file = loaded_config.get('importer_file')
            if self._importer_module is not None and self._importer_file is not None:
                # Import the plugin and load the file, update the content column related
                # attributes on success.
                plugin_module = importlib.import_module(self._importer_module)
                self._importer = plugin_module.Importer.get_instance()
                tmp_content_entries = self._importer.load(self._importer_file)
                if tmp_content_entries is not None: # Success
                    self._content_col_size = len(self._content_heading)
                    self._content_entries = tmp_content_entries
                    self._content_num_rows = len(self._content_entries)
                    for entry in self._content_entries:
                        if len(entry) > self._content_col_size:
                            # This entry have more characters than the previous, update.
                            self._content_col_size = len(entry)
                else: # Loading failed for some reason, resetting module attribute
                    self._importer_module = None

            return True

        except IOError as error:
            if args.get('output_file') is not None:
                log.debug('output_file "%s" provided from command line args', 
                          args.get('output_file'))
                self._output_file = args.get('output_file')
            if args.get('start_date') is not None and args.get('end_date') is not None:
                # This is the last resort to produce a calendar with minimal input provided
                log.debug('start_date=%s, end_date=%s provided from command line args',
                          args.get('start_date'), args.get('end_date'))
                self._start_date = self.__str_to_date(args.get('start_date'))
                self._end_date = self.__str_to_date(args.get('end_date'))
                return True
            log.error('%s', error)
            print(error)
            print("Add configuration file or provide required args. Use option -h for help")
            return False

    def __str_to_date(self, date_str: str) -> date:
        """Create a date object from a string formatted as YYYY-MM-DD.
        """
        year, month, day = date_str.split('-')
        return date(int(year), int(month), int(day))

    def update_cell_formats(self, config_formats: dict):
        """Update the cell_formats values if input is provided for each key.
        """
        if config_formats is not None:
            if config_formats.get('day') is not None:
                self._cell_formats_day = config_formats.get('day')
            if config_formats.get('weekend') is not None:
                self._cell_formats_weekend = config_formats.get('weekend')
            if config_formats.get('week_odd') is not None:
                self._cell_formats_week_odd = config_formats.get('week_odd')
            if config_formats.get('week_even') is not None:
                self._cell_formats_week_even = config_formats.get('week_even')
            if config_formats.get('month_odd') is not None:
                self._cell_formats_month_odd = config_formats.get('month_odd')
            if config_formats.get('month_even') is not None:
                self._cell_formats_month_even = config_formats.get('month_even')
            if config_formats.get('year_odd') is not None:
                self._cell_formats_year_odd = config_formats.get('year_odd')
            if config_formats.get('year_even') is not None:
                self._cell_formats_year_even = config_formats.get('year_even')
            if config_formats.get('content_heading') is not None:
                self._cell_formats_content_heading = config_formats.get('content_heading')


    @property
    def start_date(self) -> date:
        """Start date of calendar.
        """
        return self._start_date

    @property
    def end_date(self) -> date:
        """End date of calendar.
        """
        return self._end_date

    @property
    def output_file(self) -> str:
        """Name of the excel file to store the calendar in.
        """
        return self._output_file

    @property
    def worksheet_name(self) -> str:
        """Name of worksheet in the workbook.
        """
        return self._worksheet_name

    @property
    def worksheet_tab_color(self) -> str:
        """Color of worksheet tab in the workbook.
        """
        return self._worksheet_tab_color

    @property
    def content_heading(self) -> str:
        """Title or heading of the first column of the calendar.
        """
        return self._content_heading

    @property
    def content_entries(self) -> list:
        """A list of entries to populate the content column with, can be empty.
        """
        return self._content_entries

    @property
    def content_num_rows(self) -> int:
        """The number of rows the content column should be. As the entries list can be empty
        and still content rows should be printed, this attribute will be used for this
        purpose.
        """
        return self._content_num_rows

    @property
    def content_col_size(self) -> float:
        """The size of the content column which is calculated based on the longest string,
        (number of characters), in either the content_heading or any of the provided entries in
        content_entries. This value is multiplied with the average character size for
        Calibri 11p Font Size which is ~1.1p.
        """
        return self._content_col_size * 1.1

    @property
    def cell_formats_day(self) -> dict:
        """Overridden value of the day format from config.
        """
        return self._cell_formats_day

    @property
    def cell_formats_weekend(self) -> dict:
        """Overridden value of the weekend format from config.
        """
        return self._cell_formats_weekend

    @property
    def cell_formats_week_odd(self) -> dict:
        """Overridden value of the odd weeks format from config.
        """
        return self._cell_formats_week_odd

    @property
    def cell_formats_week_even(self) -> dict:
        """Overridden value of the even weeks format from config.
        """
        return self._cell_formats_week_even

    @property
    def cell_formats_month_odd(self) -> dict:
        """Overridden value of the odd month format from config.
        """
        return self._cell_formats_month_odd

    @property
    def cell_formats_month_even(self) -> dict:
        """Overridden value of the even weeks format from config.
        """
        return self._cell_formats_month_even

    @property
    def cell_formats_year_odd(self) -> dict:
        """Overridden value of the odd years format from config.
        """
        return self._cell_formats_year_odd

    @property
    def cell_formats_year_even(self) -> dict:
        """Overridden value of the even years format from config.
        """
        return self._cell_formats_year_even

    @property
    def cell_formats_content_heading(self) -> dict:
        """Overridden value of the content heading format from config.
        """
        return self._cell_formats_content_heading

    @property
    def holidays(self) -> dict:
        """Dictionary of all holidays dates keyed as <date>: 'String'
        where <date> can be either a full date in the format YYYY-MM-DD resulting in a returned
        key in datetime.date format, or in the format MM-DD resulting in a returned key as a string.
        """
        return self._holidays

    @property
    def importer_module(self) -> str:
        """
        :return: An importer module name or None
        """
        return self._importer_module

    @property
    def importer_file(self) -> str:
        """
        :return: An importer module name or None
        """
        return self._importer_file

    @property
    def importer(self) -> None:
        """
        :return: An importer module
        """
        return self._importer

    # The static ones
    @property
    def start_col(self) -> int:
        """First column of the calendar, static value and should not be changed.
        """
        return self._start_col

    @property
    def week_days(self) -> list:
        """A list with shortcut strings for the week days, 
        static value and should not be changed.
        """
        return self._week_days


    @property
    def year_row(self) -> int:
        """Year heading row.
        """
        return self._year_row

    @property
    def month_row(self) -> int:
        """Month heading row.
        """
        return self._month_row

    @property
    def week_row(self) -> int:
        """Week heading row.
        """
        return self._week_row

    @property
    def day_of_week_row(self) -> int:
        """Week heading row.
        """
        return self._day_of_week_row

    @property
    def day_row(self) -> int:
        """Day heading row.
        """
        return self._day_row

    # All setter methods
    @start_date.setter
    def start_date(self, value: str):
        """Set the start date of the calendar.
        """
        self._start_date = value

    @end_date.setter
    def end_date(self, value: str):
        """Set the end date of the calendar.
        """
        self._end_date = value
