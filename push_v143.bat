@echo off
cd /d C:\Users\jowen\cal-listener
echo === Bumping version to 1.4.3 ===
powershell -Command "(Get-Content cal_listener\__main__.py) -replace '1\.4\.2', '1.4.3' | Set-Content cal_listener\__main__.py"
powershell -Command "(Get-Content pyproject.toml) -replace 'version = \"1\.4\.2\"', 'version = \"1.4.3\"' | Set-Content pyproject.toml"
echo === Compile check ===
py -m py_compile cal_listener\handlers\dm_probe_all_screens.py || goto :fail
py -m py_compile cal_listener\handlers\__init__.py || goto :fail
echo === Git commit + tag + push ===
git add -A
git commit -m "v1.4.3: unified dm_probe_all_screens handler — walks every screen, captures titles + control surface"
git tag v1.4.3
git push origin main
git push origin v1.4.3
echo === Done ===
goto :eof
:fail
echo COMPILE FAILED
exit /b 1
