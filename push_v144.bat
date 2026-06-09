@echo off
cd /d C:\Users\jowen\cal-listener
echo === Bumping version to 1.4.4 ===
powershell -Command "(Get-Content cal_listener\__main__.py) -replace '1\.4\.3', '1.4.4' | Set-Content cal_listener\__main__.py"
powershell -Command "(Get-Content pyproject.toml) -replace 'version = \"1\.4\.3\"', 'version = \"1.4.4\"' | Set-Content pyproject.toml"
echo === Compile check ===
py -m py_compile cal_listener\handlers\dm_probe_all_screens.py || goto :fail
echo === Git commit + tag + push ===
git add -A
git commit -m "v1.4.4: probe captures empty-text left-strip icons with tooltip/legacy-name (Customers nav)"
git tag v1.4.4
git push origin main
git push origin v1.4.4
echo === Done ===
goto :eof
:fail
echo COMPILE FAILED
exit /b 1
