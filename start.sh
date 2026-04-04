#!/bin/bash
cd "$(dirname "$0")"
# 先殺掉佔用 5678 port 的舊進程
lsof -i :5678 -t 2>/dev/null | xargs kill -9 2>/dev/null
sleep 1
python3 reconcile_app.py
