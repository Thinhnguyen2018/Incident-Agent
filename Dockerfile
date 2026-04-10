FROM python:3.13-slim

WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy source
COPY . .

# Create directories for uploads and history
RUN mkdir -p uploads history

EXPOSE 5000

CMD ["python", "app.py"]
