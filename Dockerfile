
FROM python:3.12-slim

ARG USER_ID=1000
ARG GROUP_ID=1000
RUN groupadd -g $GROUP_ID appgroup && useradd -u $USER_ID -g appgroup appuser

WORKDIR /xlsxcalendar

COPY . /xlsxcalendar
RUN chown -R appuser:appgroup /xlsxcalendar

RUN pip install --no-cache-dir -r requirements.txt

# Make entrypoint script executable
RUN chmod +x entrypoint.sh

# Expose port for web interface
EXPOSE 8080

USER appuser

# Use the flexible entrypoint script
ENTRYPOINT ["./entrypoint.sh"]
