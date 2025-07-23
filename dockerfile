FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    cron \
    wget \
    unzip \
    curl \
    gnupg \
    fonts-liberation \
    libglib2.0-0 \
    libnss3 \
    libgconf-2-4 \
    libfontconfig1 \
    libxss1 \
    libasound2 \
    libxtst6 \
    libxrandr2 \
    libu2f-udev \
    libatk-bridge2.0-0 \
    libgtk-3-0 \
    && rm -rf /var/lib/apt/lists/*

RUN wget -q -O - https://dl.google.com/linux/linux_signing_key.pub | gpg --dearmor -o /etc/apt/trusted.gpg.d/google.gpg && \
    echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" > /etc/apt/sources.list.d/google-chrome.list && \
    apt-get update && \
    apt-get install -y google-chrome-stable chromium-driver && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

RUN mkdir -p /app/logs /app/relatorios

COPY requirements.txt .
RUN pip install --upgrade pip && \
    pip install -r requirements.txt

COPY . .

COPY crontab /etc/cron.d/relatorio-cron
RUN chmod 0644 /etc/cron.d/relatorio-cron && \
    crontab /etc/cron.d/relatorio-cron

RUN touch /var/log/cron.log

CMD ["cron", "-f"]
