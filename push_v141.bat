@echo off
cd /d C:\Users\jowen\cal-listener
echo === Bumping version to 1.4.1 ===
powershell -Command "(Get-Content cal_listener\__main__.py) -replace '1\.4\.0', '1.4.1' | Set-Content cal_listener\__main__.py"
powershell -Command "(Get-Content pyproject.toml) -replace 'version = \"1\.4\.0\"', 'version = \"1.4.1\"' | Set-Content pyproject.toml"
echo === Compile check ===
py -m py_compile cal_listener\handlers\customer_email_audit.py || goto :fail
echo === Git commit + tag + push ===
git add -A
git commit -m "v1.4.1: fix customer_email_audit run_audit signature + 3-sheet xlsx"
git tag v1.4.1
git push origin main
git push origin v1.4.1
echo === Done ===
goto :eof
:fail
echo COMPILE FAILED
exit /b 1
