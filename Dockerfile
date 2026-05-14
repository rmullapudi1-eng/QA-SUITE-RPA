FROM mcr.microsoft.com/playwright/python:v1.44.0-jammy

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Install Playwright Chromium (binary already present in base image)
RUN playwright install chromium

COPY . .

EXPOSE 8080
CMD ["python", "app_ui.py"]
