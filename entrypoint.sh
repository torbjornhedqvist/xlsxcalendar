#!/bin/bash

# Check if first argument is --web or --gui
if [ "$1" = "--web" ] || [ "$1" = "--gui" ]; then
    echo "Starting XlsxCalendar Web Interface..."
    shift  # Remove --web/--gui from arguments
    exec python xlsxcalendar_nicegui.py --root /xlsxcalendar "$@"
else
    echo "Starting XlsxCalendar CLI..."
    exec python xlsxcalendar.py "$@"
fi
