#!/usr/bin/env python3
"""
Copyright (C) 2022 Torbjorn Hedqvist
All Rights Reserved You may use, distribute and modify this code under the
terms of the MIT license. See LICENSE file in the project root for full
license information.
"""
import logging
import xlsxwriter

log = logging.getLogger(__name__)
class CellFormat:
    """
    A collection of formats to be used in Xlsxwriter to populate cells in
    an Excel workbook creating a calendar.
    """
    # pylint: disable=too-many-instance-attributes
    # Reasonable amount in this kind of utility class.

    def __init__(self, workbook: xlsxwriter.Workbook):
        """
        Constructor.
        """
        self._workbook = workbook
        # Static attributes, have no setter methods
        self._bold = self._workbook.add_format({'bold': True})
        self._bold_border = self._workbook.add_format({'bold': True, 'border': 1})
        self._bold_border_center = self._workbook.\
            add_format({'bold': True, 'border': 1, 'align': 'center'})
        self._border = self._workbook.add_format({'border': 1})
        self._border_center = self._workbook.add_format({'border': 1, 'align': 'center'})
        self._bold_italic = self._workbook.add_format({'bold': True, 'italic': True})

        # Default attributes, can be changed via setter methods
        self._day = self._workbook.\
            add_format({'border': 1, 'align': 'center', 'fg_color': '#D9D9D9'})
        self._weekend = self._workbook.\
            add_format({'border': 1, 'align': 'center', 'fg_color': '#cf1020'})
        self._week_odd = self._workbook.\
            add_format({'bold': True, 'border': 2, 'align': 'center', 'fg_color': '#fae7b5'})
        self._week_even = self._workbook.\
            add_format({'bold': True, 'border': 2, 'align': 'center', 'fg_color': '#C5D9F1'})
        self._month_odd = self._workbook.\
            add_format({'bold': True, 'border': 2, 'align': 'center', 'fg_color': '#B8CCE4'})
        self._month_even = self._workbook.\
            add_format({'bold': True, 'border': 2, 'align': 'center', 'fg_color': '#95B3D7'})
        self._year_odd = self._workbook.\
            add_format({'bold': True, 'border': 2, 'align': 'center', 'fg_color': '#B7DEE8'})
        self._year_even = self._workbook.\
            add_format({'bold': True, 'border': 2, 'align': 'center', 'fg_color': '#DAEEF3'})
        self._content_heading = self._workbook.\
            add_format({'bold': True, 'border': 1, 'fg_color': '#ffa700'})


    @property
    def bold(self) -> xlsxwriter.format.Format:
        """Format cell with bold"""
        return self._bold

    @property
    def bold_border(self) -> xlsxwriter.format.Format:
        """Format cell with bold and single line border"""
        return self._bold_border

    @property
    def bold_border_center(self) -> xlsxwriter.format.Format:
        """Format cell with bold and single line border and center alignment"""
        return self._bold_border_center

    @property
    def border(self) -> xlsxwriter.format.Format:
        """Format cell with a single line border"""
        return self._border

    @property
    def border_center(self) -> xlsxwriter.format.Format:
        """Format cell with a single line border and center alignment"""
        return self._border_center

    @property
    def bold_italic(self) -> xlsxwriter.format.Format:
        """Format cell with bold, italic and right alignment"""
        return self._bold_italic

    @property
    def day(self) -> xlsxwriter.format.Format:
        """Used to format the top level row of dates"""
        return self._day

    @property
    def weekend(self) -> xlsxwriter.format.Format:
        """Used to format all column cells, (including top day row), for the days representing
        weekend (Saturday and Sunday) days"""
        return self._weekend

    @property
    def week_odd(self) -> xlsxwriter.format.Format:
        """Used to format the merged collection of an odd weeks cells in the heading"""
        return self._week_odd

    @property
    def week_even(self) -> xlsxwriter.format.Format:
        """Used to format the merged collection of an even weeks cells in the heading"""
        return self._week_even

    @property
    def month_odd(self) -> xlsxwriter.format.Format:
        """Used to format the merged collection of an odd months cells in the heading"""
        return self._month_odd

    @property
    def month_even(self) -> xlsxwriter.format.Format:
        """Used to format the merged collection of an even months cells in the heading"""
        return self._month_even

    @property
    def year_odd(self) -> xlsxwriter.format.Format:
        """Used to format the merged collection of an odd years cells in the heading"""
        return self._year_odd

    @property
    def year_even(self) -> xlsxwriter.format.Format:
        """Used to format the merged collection of an even years cells in the heading"""
        return self._year_even

    @property
    def content_heading(self) -> xlsxwriter.format.Format:
        """Format cell for the content heading"""
        return self._content_heading

    # pylint: disable=missing-function-docstring
    # All setter methods
    @day.setter
    def day(self, value: dict):
        log.debug(value)
        self._day = self._workbook.add_format(value)

    @weekend.setter
    def weekend(self, value: dict):
        log.debug(value)
        self._weekend = self._workbook.add_format(value)

    @week_odd.setter
    def week_odd(self, value: dict):
        log.debug(value)
        self._week_odd = self._workbook.add_format(value)

    @week_even.setter
    def week_even(self, value: dict):
        log.debug(value)
        self._week_even = self._workbook.add_format(value)

    @month_odd.setter
    def month_odd(self, value: dict):
        log.debug(value)
        self._month_odd = self._workbook.add_format(value)

    @month_even.setter
    def month_even(self, value: dict):
        log.debug(value)
        self._month_even = self._workbook.add_format(value)

    @year_odd.setter
    def year_odd(self, value: dict):
        log.debug(value)
        self._year_odd = self._workbook.add_format(value)

    @year_even.setter
    def year_even(self, value: dict):
        log.debug(value)
        self._year_even = self._workbook.add_format(value)

    @content_heading.setter
    def content_heading(self, value: dict):
        log.debug(value)
        self._content_heading = self._workbook.add_format(value)
