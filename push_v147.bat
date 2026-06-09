@echo off
cd /d C:\Users\jowen\cal-listener
powershell -Command "(Get-Content cal_listener\__main__.py) -replace '1\.4\.6', '1.4.7' | Set-Content cal_listener\__main__.py"
powershell -Command "(Get-Content pyproject.toml) -replace 'version = \"1\.4\.6\"', 'version = \"1.4.7\"' | Set-Content pyproject.toml"
git add -A
git commit -m "v1.4.7: port ClearBooks Playwright driver + 3 cb_* handlers wired (create_bill, edit_bill, credit_note) + cb_login"
git tag v1.4.7
git push origin main
git push origin v1.4.7
