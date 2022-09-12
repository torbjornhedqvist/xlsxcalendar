#!/usr/bin/env python3
"""
Copyright (C) 2022 Torbjorn Hedqvist
All Rights Reserved You may use, distribute and modify this code under the
terms of the MIT license. See LICENSE file in the project root for full
license information.
"""
import logging
from util.config import Config

# pylint: disable=missing-function-docstring
log = logging.getLogger(__name__)
class Offset:
    """Wrapper class for all offset values.
    """

    def __init__(self, conf: Config):
        """Constructor.
        """
        self._week = conf.start_col
        self._month = conf.start_col
        self._year = conf.start_col

    @property
    def week(self) -> int:
        return self._week

    @property
    def month(self) -> int:
        return self._month

    @property
    def year(self) -> int:
        return self._year

    # Setter methods
    @week.setter
    def week(self, value: int):
        log.debug('Offset=%s', value)
        self._week = value

    @month.setter
    def month(self, value: int):
        log.debug('Offset=%s', value)
        self._month = value

    @year.setter
    def year(self, value: int):
        log.debug('Offset=%s', value)
        self._year = value


class Current:
    """Wrapper class for all current date values.
    """

    def __init__(self, conf: Config):
        """Constructor.
        """
        self._week = int(conf.start_date.strftime('%V'))
        self._month = conf.start_date.month
        self._year = conf.start_date.year

    @property
    def week(self) -> int:
        return self._week

    @property
    def month(self) -> int:
        return self._month

    @property
    def year(self) -> int:
        return self._year

    # Setter methods
    @week.setter
    def week(self, value: int):
        log.debug('Current=%s', value)
        self._week = value

    @month.setter
    def month(self, value: int):
        log.debug('Current=%s',value)
        self._month = value

    @year.setter
    def year(self, value: int):
        log.debug('Current=%s', value)
        self._year = value


class Next:
    """Wrapper class for all next date values.
    """

    def __init__(self):
        """Constructor.
        """
        self._week = None
        self._month = None
        self._year = None

    @property
    def week(self) -> int:
        return self._week

    @property
    def month(self) -> int:
        return self._month

    @property
    def year(self) -> int:
        return self._year

    # Setter methods
    @week.setter
    def week(self, value: int):
        self._week = value

    @month.setter
    def month(self, value: int):
        self._month = value

    @year.setter
    def year(self, value: int):
        self._year = value

# pylint: enable=missing-function-docstring
class Dateinfo:
    """Wrapper class for all date info counters used in the main loop of calendar creation.
    """

    def __init__(self, conf: Config):
        """Constructor.
        """
        self._offset = Offset(conf)
        self._current = Current(conf)
        self._next = Next()

    @property
    def offset(self) -> Offset:
        """Returns an Offset instance.
        """
        return self._offset

    @property
    def current(self) -> Current:
        """Returns a Current instance.
        """
        return self._current

    @property
    def next(self) -> Current:
        """Returns a Next instance.
        """
        return self._next
