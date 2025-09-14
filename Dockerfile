
FROM python:3.12-slim

WORKDIR /xlsxcalendar

COPY . /xlsxcalendar

RUN pip install --no-cache-dir -r requirements.txt

# Command to run your application
ENTRYPOINT ["python", "xlsxcalendar.py"]