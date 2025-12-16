
FROM python:3.12-slim

WORKDIR /xlsxcalendar

COPY . /xlsxcalendar

RUN pip install --no-cache-dir -r requirements.txt

# Make entrypoint script executable
RUN chmod +x entrypoint.sh

# Expose port for web interface
EXPOSE 8080

# Use the flexible entrypoint script
ENTRYPOINT ["./entrypoint.sh"]