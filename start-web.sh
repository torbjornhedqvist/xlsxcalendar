#!/bin/bash
# Start the web interface in Docker container

echo "Starting XlsxCalendar Web Interface..."
echo "Access the interface at: http://localhost:8080"
echo "Press Ctrl+C to stop"

docker-compose up xlsxcalendar-web
