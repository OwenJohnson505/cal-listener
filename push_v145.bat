@echo off
cd /d C:\Users\jowen\cal-listener
echo === Bumping version to 1.4.5 ===
powershell -Command "(Get-Content cal_listener\__main__.py) -replace '1\.4\.4', '1.4.5' | Set-Content cal_listener\__main__.py"
powershell -Command "(Get-Content pyproject.toml) -replace 'version = \"1\.4\.4\"', 'version = \"1.4.5\"' | Set-Content pyproject.toml"
echo === Git commit + tag + push ===
git add -A
git commit -m "v1.4.5: port tariff_assigner engine + handler"
git tag v1.4.5
git push origin main
git push origin v1.4.5
echo === Done ===
