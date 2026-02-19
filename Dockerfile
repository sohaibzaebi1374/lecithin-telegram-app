FROM python:3.11-slim

# Prevent Python from buffering stdout/stderr (useful for logs on Fly)
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# System deps (optional but helpful)
RUN apt-get update && apt-get install -y --no-install-recommends \
    ca-certificates \
 && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copy app
COPY . .

# Run the bot (polling)
CMD ["python", "bot.py"]
