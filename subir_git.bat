@echo off
cd /d %~dp0
git add .
git commit -m "Actualización automática %date% %time%"
git push origin main
