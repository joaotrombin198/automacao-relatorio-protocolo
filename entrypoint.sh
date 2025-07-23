#!/bin/bash
export PATH=/usr/local/bin:$PATH

echo "[INFO] Rodando script em $(date)" >> /app/log.txt
echo "PATH=$PATH" >> /app/log.txt
echo "Which python3: $(which python3)" >> /app/log.txt
python3 --version >> /app/log.txt
python3 -m pip show selenium >> /app/log.txt 2>&1
python3 -m pip list >> /app/log.txt 2>&1

cd /app
python3 gerar-relatorio.py >> /app/log.txt 2>&1
