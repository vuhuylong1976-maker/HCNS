@echo off
rem TAPACO Nhan su - gan URL may chu tu webapp.txt vao index.html
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0GanURL_MayChu.ps1" %*
pause
