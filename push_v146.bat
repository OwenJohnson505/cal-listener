@echo off
cd /d C:\Users\jowen\cal-listener
powershell -Command "(Get-Content cal_listener\__main__.py) -replace '1\.4\.5', '1.4.6' | Set-Content cal_listener\__main__.py"
powershell -Command "(Get-Content pyproject.toml) -replace 'version = \"1\.4\.5\"', 'version = \"1.4.6\"' | Set-Content pyproject.toml"
git add -A
git commit -m "v1.4.6: port bookings_report + maersk_report file-processor engines"
git tag v1.4.6
git push origin main
git push origin v1.4.6
