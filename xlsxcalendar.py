#!/usr/bin/env python3
"""
Copyright (C) 2022 Torbjorn Hedqvist
All Rights Reserved You may use, distribute and modify this code under the
terms of the MIT license. See LICENSE file in the project root for full
license information.

TODO: Add a better description

"""
import logging
import logging.config
import argparse
from datetime import timedelta
import sys
import xlsxwriter
from util.config import Config
from util.cell_format import CellFormat
from util.dateinfo import Dateinfo
from util.static_layout import set_static_layout
from util.merge_trim import *

def parse_args() -> dict:
    """Parse the command line arguments. If any of the rules defined by this function
    is broken, the program will abort with a clear error message given by argparse.
    Returns: vars as a dict with all available arguments as keys.
    """
    parser = argparse.ArgumentParser(
        description="Creates a calendar in Excel",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("-s", "--start-date", type=str,
                        help="Start date of calendar, using format of YYYY-MM-DD")
    parser.add_argument("-e", "--end-date", type=str,
                        help="End date of calendar, using format YYYY-MM-DD")
    parser.add_argument("-c", "--config-file", type=str,
                        help="Specify alternate configuration input file in yaml format")
    parser.add_argument("-i", "--import-file", type=str,
                        help="data file to be imported into the calendar.")
    parser.add_argument("-l", "--log-level", type=str,
                        help="Set the log level, default is INFO")
    return vars(parser.parse_args())


def update_forms(cform: CellFormat, conf: Config):
    """Since we can't create the workbook until we have read filename information stored
    in configuration we have to wait and update the cell forms from default values after
    the configuration is loaded
    """
    if conf.cell_formats_day:
        cform.day = conf.cell_formats_day
    if conf.cell_formats_weekend:
        cform.weekend = conf.cell_formats_weekend
    if conf.cell_formats_week_odd:
        cform.week_odd = conf.cell_formats_week_odd
    if conf.cell_formats_week_even:
        cform.week_even = conf.cell_formats_week_even
    if conf.cell_formats_month_odd:
        cform.month_odd = conf.cell_formats_month_odd
    if conf.cell_formats_month_even:
        cform.month_even = conf.cell_formats_month_even
    if conf.cell_formats_year_odd:
        cform.year_odd = conf.cell_formats_year_odd
    if conf.cell_formats_year_even:
        cform.year_even = conf.cell_formats_year_even
    if conf.cell_formats_content_heading:
        cform.content_heading = conf.cell_formats_content_heading


def main():
    """Main program"""
    args = parse_args()
    logging.config.fileConfig('./util/logging.conf')
    if args.get('log_level') is not None:
        numeric_level = getattr(logging, args.get('log_level').upper())
        logging.getLogger().setLevel(numeric_level)
        # pylint: disable=E1101
        for logger in logging.root.manager.loggerDict:
            logging.getLogger(logger).setLevel(numeric_level)
        # pylint: enable=E1101

    # Logging in main, example below
    # logging.info('Just a dummy INFO test')
    # logging.debug('Just a dummy DEBUG test')
    conf = Config()
    if conf.load_config(args) is False:
        logging.error('Failed to')
        sys.exit(1)

    total_days = (conf.end_date - conf.start_date).days + 1 # +1 to include the last day
    workbook = xlsxwriter.Workbook(conf.output_file)
    cform = CellFormat(workbook) # with default settings
    update_forms(cform, conf) # and update from conf if needed
    worksheet = workbook.add_worksheet(conf.worksheet_name)
    set_static_layout(workbook, conf, cform)

    dateinfo = Dateinfo(conf)
    # Loop over all dates in the range from start date to end date day by day and
    # populate the calendar heading, (year/month/week/day) and all the rows below.
    for day in range(total_days):
        current_date = conf.start_date + timedelta(days = day)
        dateinfo.next.week = int(current_date.strftime('%V'))
        if is_week_merged(worksheet, conf, cform, day, dateinfo) is True:
            dateinfo.offset.week = day + 1
            dateinfo.current.week = dateinfo.next.week

        dateinfo.next.month = current_date.month
        if is_month_merged(worksheet, conf, cform, day, current_date, dateinfo) is True:
            dateinfo.offset.month = day + 1
            dateinfo.current.month = dateinfo.next.month

        dateinfo.next.year = current_date.year
        if is_year_merged(worksheet, conf, cform, day, dateinfo) is True:
            dateinfo.offset.year = day + 1
            dateinfo.current.year = dateinfo.next.year

        # Finally write the next date and if it's a weekend update all cells for all defined
        # rows with a "weekend" formatting.
        if int(current_date.strftime('%u')) < 6: # Weekday
            worksheet.write(conf.day_of_week_row, conf.start_col + day,
                            conf.week_days[current_date.weekday()], cform.day)
            worksheet.write(conf.day_row, conf.start_col + day, current_date.day, cform.day)
        else: # Weekend
            worksheet.write(conf.day_of_week_row, conf.start_col + day,
                            conf.week_days[current_date.weekday()], cform.weekend)
            worksheet.write(conf.day_row, conf.start_col + day, current_date.day, cform.weekend)
            for rows in range(conf.content_num_rows):
                worksheet.write(conf.day_row + rows + 1, conf.start_col + day, '', cform.weekend)

        # See if this day is found in the list of holidays, if yes mark it as a weekend and
        # add a note with the reason text.
        if conf.holidays.get(current_date) is not None:
            worksheet.write(conf.day_of_week_row, conf.start_col + day,
                            conf.week_days[current_date.weekday()], cform.weekend)
            worksheet.write(conf.day_row, conf.start_col + day, current_date.day, cform.weekend)
            for rows in range(conf.content_num_rows):
                worksheet.write(conf.day_row + rows + 1, conf.start_col + day, '', cform.weekend)
            worksheet.write((conf.day_row + conf.content_num_rows + 1), conf.start_col + day,
                            '!', cform.bold_border_center)
            worksheet.write_comment((conf.day_row + conf.content_num_rows + 1),
                                     conf.start_col + day, conf.holidays.get(current_date))

    # Wrap up the end of the calendar
    trim_end_of_week(worksheet, conf, cform, day, dateinfo)
    trim_end_of_month(worksheet, conf, cform, day, current_date, dateinfo)
    trim_end_of_year(worksheet, conf, cform, day, dateinfo)

    # If we have an importer loaded it means we can plot the content from the imported file.
    if conf.importer:
        if conf.importer.plot(conf, workbook, cform) is True:
            print(f'Calendar successfully created. To be stored in {conf.output_file}')
        else:
            print('Parts of the imports seems to have failed, see the log for more info.')

    while True:
        try:
            workbook.close()
        except xlsxwriter.exceptions.FileCreateError as error:
            decision = input(f"Exception caught in workbook.close(): {error}\n"
                             "Please close the file if it is open in Excel.\n"
                             "Try to write file again? [Y/n]: ")
            if decision != 'n':
                continue
            sys.exit(0)
        print('File stored successfully.')
        break


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt as err:
        print('Interrupted')
        sys.exit(err)
