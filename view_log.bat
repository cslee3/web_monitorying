@echo off
powershell -Command "$OutputEncoding=[Console]::OutputEncoding=[Text.Encoding]::UTF8; Get-Content -Path '%~dp0log\web_monitor.log' -Wait -Tail 100 -Encoding UTF8"
