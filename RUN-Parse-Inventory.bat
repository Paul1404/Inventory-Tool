@echo off
set PS_SCRIPT_PATH=.\Parse-Inventory.ps1
powershell -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -Command "& { . '%PS_SCRIPT_PATH%'; exit $LASTEXITCODE }"
