FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy your app code
COPY . .

EXPOSE 5000

# Run Gunicorn with a factory function
# This uses the syntax: "module:callable()" for create_app()
CMD ["gunicorn", "--workers=2", "--bind=0.0.0.0:5050", "run:app"]
