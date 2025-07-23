FROM python:3.11-slim

# Instala dependências do sistema
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

# Instala Chrome e ChromeDriver
RUN apt-get update && apt-get install -y wget gnupg2 ca-certificates curl unzip && \
    wget -q -O - https://dl.google.com/linux/linux_signing_key.pub | gpg --dearmor -o /etc/apt/trusted.gpg.d/google.gpg && \
    echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" > /etc/apt/sources.list.d/google-chrome.list && \
    apt-get update && \
    apt-get install -y google-chrome-stable chromium-driver

# Cria diretório de trabalho
WORKDIR /app

COPY requirements.txt requirements.txt
RUN pip install --upgrade pip
RUN pip install -r requirements.txt

COPY . .
COPY entrypoint.sh /app/entrypoint.sh
RUN chmod +x /app/entrypoint.sh



# Torna entrypoint executável
RUN chmod +x /app/entrypoint.sh



# Cria cronjob: roda a cada minuto (ajuste depois para "0 4 * * *" se quiser só às 04h)
RUN echo "*/1 * * * * /app/entrypoint.sh >> /var/log/cron.log 2>&1" > /etc/cron.d/relatorio-cron

RUN chmod 0644 /etc/cron.d/relatorio-cron
RUN crontab /etc/cron.d/relatorio-cron

# Inicia cron no modo foreground
CMD ["cron", "-f"]
