#!/usr/bin/env python3
"""
Copyright (C) 2022 Torbjorn Hedqvist
All Rights Reserved You may use, distribute and modify this code under the
terms of the MIT license. See LICENSE file in the project root for full
license information.

Supporting functions to merge and trim the heading rows for weeks, months and years.

"""
from datetime import date, timedelta
import xlsxwriter
from util.config import Config
from util.cell_format import CellFormat
from util.dateinfo import Dateinfo

def is_week_merged(worksheet: xlsxwriter.Workbook.worksheet_class, conf: Config, cform: CellFormat,
                      day: int, dateinfo: Dateinfo) -> bool:
    """Check if we have a new week, if yes merge all the cells for the current week and print
    the week number. Toggle between two background colors for even and odd weeks
    """

    if dateinfo.next.week != dateinfo.current.week:
        if dateinfo.next.week % 2: # even
            if dateinfo.offset.week != day:
                worksheet.merge_range(conf.week_row, dateinfo.offset.week, conf.week_row, day,
                                    f"W{dateinfo.current.week}", cform.week_even)
            else:
                # Just one cell, can't merge and no place for text
                worksheet.write_string(conf.week_row, day, '', cform.week_even)
        else: # odd
            if dateinfo.offset.week != day:
                worksheet.merge_range(conf.week_row, dateinfo.offset.week, conf.week_row, day,
                                    f"W{dateinfo.current.week}", cform.week_odd)
            else:
                # Just one cell, can't merge and no place for text
                worksheet.write_string(conf.week_row, day, '', cform.week_odd)
        return True
    return False


def is_month_merged(worksheet: xlsxwriter.Workbook.worksheet_class, conf: Config, cform: CellFormat,
                    day: int, current_date: date, dateinfo: Dateinfo) -> bool:
    #pylint: disable-msg=too-many-arguments
    """Check if we have a new month, if yes merge all the cells for the current month and print
    the month name. Toggle between two background colors for even and odd months
    """

    if dateinfo.next.month != dateinfo.current.month:
        # New month, time to merge the cells for previous month and print month name.
        # As current_date is already set to next month we adjust using timedelta(days = -1).
        if dateinfo.next.month % 2:
            if dateinfo.offset.month != day:
                worksheet.merge_range(conf.month_row, dateinfo.offset.month, conf.month_row, day,
                                    (current_date + timedelta(days = -1)).strftime('%b'),
                                    cform.month_even)
            else:
                # Just one cell, can't merge and no place for text
                worksheet.write_string(conf.month_row, day, '', cform.month_even)
        else:
            if dateinfo.offset.month != day:
                worksheet.merge_range(conf.month_row, dateinfo.offset.month, conf.month_row, day,
                                    (current_date + timedelta(days = -1)).strftime('%b'),
                                    cform.month_odd)
            else:
                # Just one cell, can't merge and no place for text
                worksheet.write_string(conf.month_row, day, '', cform.month_odd)
        return True
    return False


def is_year_merged(worksheet: xlsxwriter.Workbook.worksheet_class, conf: Config, cform: CellFormat,
                   day: int, dateinfo: Dateinfo) -> bool:
    """Check if we have a new month, if yes merge all the cells for the current month and print
    the month name. Toggle between two background colors for even and odd months
    """

    if dateinfo.next.year != dateinfo.current.year:
        # New year, time to merge the cells for previous year and print the year.
        if dateinfo.next.year % 2:
            if dateinfo.offset.year != day:
                worksheet.merge_range(conf.year_row, dateinfo.offset.year, conf.year_row, day,
                                      dateinfo.current.year, cform.year_even)
            else:
                # Just one cell, can't merge and no place for text
                worksheet.write_string(conf.year_row, day, '', cform.year_even)
        else:
            if dateinfo.offset.year != day:
                worksheet.merge_range(conf.year_row, dateinfo.offset.year, conf.year_row, day,
                                    dateinfo.current.year, cform.year_odd)
            else:
                # Just one cell, can't merge and no place for text
                worksheet.write_string(conf.year_row, day, '', cform.year_odd)
        return True
    return False


def trim_end_of_week(worksheet: xlsxwriter.Workbook.worksheet_class, conf: Config,
                     cform: CellFormat, day: int, dateinfo: Dateinfo):
    """Wrap up the end of the week row merging and print the last week or just add
    a single cell with week formatting if that remains
    """
    if dateinfo.offset.week <= day:
        if dateinfo.next.week % 2:
            worksheet.merge_range(conf.week_row, dateinfo.offset.week, conf.week_row, day + 1,
                                  f"W{dateinfo.current.week}", cform.week_odd)
        else:
            worksheet.merge_range(conf.week_row, dateinfo.offset.week, conf.week_row, day + 1,
                                  f"W{dateinfo.current.week}", cform.week_even)
    elif dateinfo.offset.week == day + 1:
        # Just one cell, can't merge and no place for text
        if dateinfo.next.week % 2:
            worksheet.write_string(conf.week_row, day + 1, '', cform.week_odd)
        else:
            worksheet.write_string(conf.week_row, day + 1, '', cform.week_even)


def trim_end_of_month(worksheet: xlsxwriter.Workbook.worksheet_class, conf: Config,
                      cform: CellFormat, day: int, current_date: date, dateinfo: Dateinfo):
    #pylint: disable-msg=too-many-arguments
    """Wrap up the end of the month row merging and print the last month or just add
    a single cell with month formatting if that remains
    """
    if dateinfo.offset.month <= day:
        if dateinfo.next.month % 2:
            worksheet.merge_range(conf.month_row, dateinfo.offset.month, conf.month_row, day + 1,
                                (current_date + timedelta(days = -1)).strftime('%b'),
                                cform.month_odd)
        else:
            worksheet.merge_range(conf.month_row, dateinfo.offset.month, conf.month_row, day + 1,
                                (current_date + timedelta(days = -1)).strftime('%b'),
                                cform.month_even)
    elif dateinfo.offset.month > day:
        # Just one cell, can't merge and no place for text
        if dateinfo.next.month % 2:
            worksheet.write_string(conf.month_row, day + 1, '', cform.month_odd)
        else:
            worksheet.write_string(conf.month_row, day + 1, '', cform.month_even)


def trim_end_of_year(worksheet: xlsxwriter.Workbook.worksheet_class, conf: Config,
                     cform: CellFormat, day: int, dateinfo: Dateinfo):
    """Wrap up the end of the year row merging and print the last year or just add
    a single cell with year formatting if that remains
    """
    if dateinfo.offset.year <= day:
        if dateinfo.next.year % 2:
            worksheet.merge_range(conf.year_row, dateinfo.offset.year, conf.year_row, day + 1,
                                  dateinfo.current.year, cform.year_odd)
        else:
            worksheet.merge_range(conf.year_row, dateinfo.offset.year, conf.year_row, day + 1,
                                  dateinfo.current.year, cform.year_even)
    elif dateinfo.offset.year > day:
        # Just one cell, can't merge and no place for text
        if dateinfo.next.year % 2:
            worksheet.write_string(conf.year_row, day + 1, '', cform.year_odd)
        else:
            worksheet.write_string(conf.year_row, day + 1, '', cform.year_even)
