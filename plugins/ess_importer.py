#!/usr/bin/env python
"""
Copyright (C) 2022 Torbjorn Hedqvist
All Rights Reserved You may use, distribute and modify this code under the
terms of the MIT license. See LICENSE file in the project root for full
license information.

An importer implementation
"""
from datetime import date, timedelta
import logging
import csv
import xlsxwriter
import pandas
from plugins.abstract_importer import AbstractImporter
from util.config import Config
from util.cell_format import CellFormat

log = logging.getLogger(__name__)
class Importer(AbstractImporter):
    """
    Realization class to to handle ESS comma separated value (CSV) files and
    write them into the currently created calendar.
    """
    __instance = None

    def __init__(self):
        """Constructor.
        """
        # Instance attributes
        self._records = {}
        self._content_num_rows = None
        self._date_row = None
        self._start_date = None

    @classmethod
    def get_instance(cls):
        """
        If instance is None create an instance of this class
        and return it, else return the existing instance.

        :return: An Importer instance.
        """
        if cls.__instance is None:
            log.debug("Create singleton instance")
            cls.__instance = cls()
        return cls.__instance


    def __get_months_and_days(self) -> list:
        """
        Pull out start day and month and end day and month from the stored list of
        provided dates in the heading in the imported file. They have the following
        format: "DD.MM"
        """
        s_day, s_month = self._date_row[0].split('.') # first element (start date)
        e_day, e_month = self._date_row[-1].split('.') # last element (end date)
        log.debug('s%s.%s, e%s.%s', s_day, s_month, e_day, e_month)
        return int(s_day), int(s_month), int(e_day), int(e_month)


    def __is_dates_in_range(self, conf: Config) -> bool:
        """
        Verify that the imported file dates are within the range of the calendar
        start_date and end_date.

        :return: True if within range, else False
        """
        s_day, s_month, e_day, e_month = self.__get_months_and_days()
        log.debug('s%s, e%s', conf.start_date, conf.end_date)

        if conf.start_date.year == conf.end_date.year:
            if (e_month - s_month) >= 0: # Also within one year range
                if s_month >= conf.start_date.month and e_month <= conf.end_date.month:
                    # OK unless completely out of range, add extra check to see if
                    # weekends are matching in plot()
                    if s_month == conf.start_date.month and s_day < conf.start_date.day:
                        log.debug('import start_day < calendar')
                        return False
                    if e_month == conf.end_date.month and e_day > conf.end_date.day:
                        log.debug('import end_day > calendar')
                        return False
                    log.debug('Both calendar and imports in same year and within range.')
                    return True
            else:
                # Has to be out of range if calendar is within a year and the import
                # is crossing year boundary
                log.debug('Calendar start & end year is same but not imports.')
                return False
        else:
            if (e_month - s_month) >= 0:
                log.debug('Calendar cross year boundary but not imports.')
                if s_month >= conf.start_date.month:
                    if s_month == conf.start_date.month and s_day < conf.start_date.day:
                        log.debug('import start_day < calendar')
                        return False
                    log.debug('import start within calendar range.')
                    return True
                log.debug('import start month less than calendar start month')
                return False

            log.debug('Calendar and imports cross year boundary.')
            if s_month >= conf.start_date.month and e_month <= conf.end_date.month:
                if s_month == conf.start_date.month and s_day < conf.start_date.day:
                    log.debug('import start_day < calendar')
                    return False
                if e_month == conf.end_date.month and e_day > conf.end_date.day:
                    log.debug('import end_day > calendar')
                    return False
                log.debug('imports within range.')
                return True
            return False
        return True

    def __set_start_date(self, conf: Config):
        """
        Set the start date of the imported data based on current best knowledge
        from the __is_dates_in_range(...). Guessing the year for the start date.
        Additional checks will be made in plot to see if weekends are in sync.
        """
        # pylint: disable=W0612 # e_day not used
        s_day, s_month, e_day, e_month = self.__get_months_and_days()
        # pylint: enable=W0612
        if conf.start_date.year == conf.end_date.year:
            if (e_month - s_month) >= 0:
                # Within one year range for both calendar and imports.
                self._start_date = date(conf.start_date.year, s_month, s_day)
        else:
            # Calendar crossing year boundaries
            if s_month <= 12:
                self._start_date = date(conf.start_date.year, s_month, s_day)
            else:
                self._start_date = date(conf.end_date.year, s_month, s_day)
        log.debug('%s', self._start_date)


    def __is_weekend(self, a_date: date) -> bool:
        """
        :return: True if the provided date is a Saturday or Sunday, else False.
        """
        if int(a_date.strftime('%u')) > 5:
            return True
        return False


    def load(self, filename: str) -> list:
        """
        Load a semicolon separated table with the following layout:
        Personnel No;Name;Org Unit;Company;Country;19.08;20.08;21.08;22.08;23.08...
        12345678;Kalle Karlsson;The unit;XYZ;SE;;O;O;A;;...
        ...
        where 3 dots (...) represents arbitrary extension in that direction.

        As there are no year in the heading it will require some extra checks to
        verify that the imported file is within the range of the calendar.

        :return: A list of strings which will be used to populate the content_entries
        in the calendar, or None if failing to read file.
        """
        log.debug('Enter')
        if '.xlsx' in filename:
            log.debug('%s seems to be in xlsx format, needs to be converted to a cvs file',
                      filename)
            read_xlsx_file = pandas.read_excel(filename)
            filename = filename.replace('.xlsx', '.csv')
            log.debug('New filename is %s', filename)
            read_xlsx_file.to_csv(filename, sep = ';',
                                  index = None,
                                  header = True,
                                  encoding='iso-8859-1')
            log.info('xlsx formatted file converted to semicolon separated csv file')

        try:
            with open(filename, encoding='ISO-8859-1', newline='') as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=';', quotechar='|')
                for row in csv_reader:
                    # pop will left shift between operations
                    row.pop(0) # Remove Personnel No
                    row.pop(1) # Remove Org unit
                    row.pop(1) # Remove Company
                    row.pop(1) # Remove Country
                    self._records.update({row[0]: row[1:]})

                if 'Name' in self._records:
                    self._date_row = self._records.get('Name')
                    log.debug('self._date_row = %s', self._date_row)
                    self._records.pop('Name')
                log.debug('self._records = %s', self._records)
                content_entries = list(self._records.keys())
                self._content_num_rows = (len(content_entries))
                log.debug('Exit: %s', content_entries)
                return content_entries
        except IOError as error:
            log.error(error)
        log.debug('Exit: None')
        return None


    def plot(self, conf: Config, workbook: xlsxwriter.Workbook, cform: CellFormat) -> bool:
        """
        Plot the imported ESS csv formatted input in the currently created calendar.

        :return: True if it succeeds with all operations, else False
        """
        log.debug('Enter')
        if self.__is_dates_in_range(conf) is True:
            self.__set_start_date(conf)
            worksheet = workbook.get_worksheet_by_name(conf.worksheet_name)
            approved_absence_format = workbook.\
                add_format({'border': 1, 'align': 'center', 'fg_color': '#00FF00'})
            planned_absence_format = workbook.\
                add_format({'border': 1, 'align': 'center', 'fg_color': '#00B0F0'})
            legend_format = workbook.\
                add_format({'bold': True, 'border': 2,'align': 'center', 'fg_color': '#D9E1F2'})
            legend_approved_absence_format = workbook.\
                add_format({'border': 1, 'fg_color': '#00FF00'})
            legend_planned_absence_format = workbook.\
                add_format({'border': 1, 'fg_color': '#00B0F0'})
            delta = (self._start_date - conf.start_date).days
            log.debug('delta=%s', delta)
            row = 0
            for record in self._records.items():
                row +=1
                for day in range(len(record[1])):
                    entry = record[1][day]
                    if entry == '':
                        # TODO: Add check here if this is different in generated calendar.
                        # Can be different weekend rules for different countries.
                        continue # empty, skip

                    log.debug('entry=%s', entry)

                    # This is an extra check to see that the weekends are aligned as the
                    # __is_dates_in_range() only catches if the imports and the calendar is
                    # within reasonable diff ~6 month, larger diff can slip through.
                    current_date = (conf.start_date + timedelta(days = day + delta))
                    if entry == 'O' and not self.__is_weekend(current_date):
                        log.debug('current_date=%s', current_date)
                        log.error('The weekends in calendar and imports are not in sync.')
                        log.error('Cannot continue the plotting with unreliable data, skip plot.')
                        return False

                    if entry == 'O':
                        log.debug('O, weekend and skip')
                        continue # weekend, skip
                    if entry == 'H':
                        # This is a holiday and it might be an imported local holiday not covered by
                        # configured holidays, make sure it will be marked in the calendar as well.
                        log.debug("entry row=%s", conf.day_row + row)
                        log.debug("entry column=%s", conf.start_col + delta + day)
                        worksheet.write(conf.day_row + row, conf.start_col + delta + day,
                                        '', cform.weekend)
                    elif entry == 'P':
                        worksheet.write(conf.day_row + row, conf.start_col + delta + day,
                                        entry, planned_absence_format)
                    elif entry == 'A':
                        worksheet.write(conf.day_row + row, conf.start_col + delta + day,
                                        entry, approved_absence_format)
                    else:
                        log.warning('Cannot recognize this input=%s, skip plot', entry)
                        return False
            # Finally write a Legend explaining the the output
            worksheet.write(conf.day_row + self._content_num_rows + 2,
                            conf.start_col - 1, 'Legend', legend_format)
            worksheet.write(conf.day_row + self._content_num_rows + 3,
                            conf.start_col - 1, 'Approved absence="A"',
                            legend_approved_absence_format)
            worksheet.write(conf.day_row + self._content_num_rows + 4,
                            conf.start_col - 1, 'Planned absence="P"',
                            legend_planned_absence_format)
            log.debug('Exit: True')
            return True
        log.error('Imported file date range outside calendar range, skip plot.')
        return False
