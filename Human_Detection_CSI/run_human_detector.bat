@echo off
title Wi-Fi CSI Human Detection & Vitals System
echo =====================================================================
echo  Camera-Free Wi-Fi CSI Human Presence & Vital Signs Sensing
echo =====================================================================
echo.
echo Starting sensing server on port 8888...
start http://localhost:8888
python -u csi_detector_server.py --http-port 8888
pause
