#!/usr/bin/env python
"""
Copyright (C) 2022 Torbjorn Hedqvist
All Rights Reserved You may use, distribute and modify this code under the
terms of the MIT license. See LICENSE file in the project root for full
license information.

A template importer implementation with bare minimum mandatory API's
"""
import logging
import xlsxwriter
from plugins.abstract_importer import AbstractImporter
from util.config import Config

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
        # To be added

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

    def load(self, filename: str) -> list:
        """
        Load the intended import file.

        :return: MUST return a list of strings which will be used to populate the content_entries
        in the calendar, or None if failing to read file.
        """
        log.debug('Enter')
        # Add your load implementation here and return the list of strings.
        print('template_importer.load()')
        log.debug('Exit: None')
        # return a_list_of_content_entries where each entry is a string.


    def plot(self, conf: Config, workbook: xlsxwriter.Workbook) -> bool:
        """
        Plot the content of the imported file into the currently created calendar.

        :return: True if it succeeds with all operations, else False
        """
        log.debug('Enter')
        # worksheet = workbook.get_worksheet_by_name(conf.worksheet_name)
        print('template_importer.plot()')
        log.debug('Exit: True')
        return True
