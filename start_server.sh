#!/bin/bash
set -e
cd "$(dirname "$0")"
pip3 install -r requirements.txt
if [ ! -f server.env ]; then
  echo "Создайте server.env из server.env.example"
  cp server.env.example server.env
  echo "Заполните BOT_TOKEN и ADMIN_IDS в server.env"
  exit 1
fi
python3 run_server.py
