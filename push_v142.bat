@echo off
cd /d C:\Users\jowen\cal-listener
echo === Bumping version to 1.4.2 ===
powershell -Command "(Get-Content cal_listener\__main__.py) -replace '1\.4\.1', '1.4.2' | Set-Content cal_listener\__main__.py"
powershell -Command "(Get-Content pyproject.toml) -replace 'version = \"1\.4\.1\"', 'version = \"1.4.2\"' | Set-Content pyproject.toml"
echo === Compile check ===
py -m py_compile cal_listener\handlers\dm_probe_nav.py || goto :fail
py -m py_compile cal_listener\handlers\revenue_breakdown_scraper.py || goto :fail
py -m py_compile cal_listener\handlers\__init__.py || goto :fail
echo === Git commit + tag + push ===
git add -A
git commit -m "v1.4.2: dm_probe_nav handler + 2-step Booking->Customer Invoice nav fix"
git tag v1.4.2
git push origin main
git push origin v1.4.2
echo === Done ===
goto :eof
:fail
echo COMPILE FAILED
exit /b 1
