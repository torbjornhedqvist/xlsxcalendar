#!/usr/bin/env python
"""
Copyright (C) 2022 Torbjorn Hedqvist
All Rights Reserved You may use, distribute and modify this code under the
terms of the MIT license. See LICENSE file in the project root for full
license information.

Abstract class defining the mandated method(s) and associated signatures to import a
custom defined calendar file and write the results in to the currently created calendar.
"""
from abc import ABC, abstractmethod
import xlsxwriter
from util.config import Config

class AbstractImporter(ABC):
    """
    Abstract Importer Interface
    """

    @abstractmethod
    def load(self, filename: str) -> list:
        """
        A load method to read the file content provided in filename. This method should
        be called before the main program starts to create the calendar as it will provide
        input on how many rows (content_entries) the calendar should have.

        :return: A list of strings which will be used to populate the content_entries
        in the calendar.
        """


    @abstractmethod
    def plot(self, conf: Config, workbook: xlsxwriter.Workbook) -> bool:
        """
        A plot method which will plot out the content in the newly created calender.
        This method should be called after the main calendar is created.
        All needed input is provided in the Config and access to the workbook to plot.
        """
