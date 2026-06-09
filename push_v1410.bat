@echo off
cd /d C:\Users\jowen\cal-listener
powershell -Command "(Get-Content cal_listener\__main__.py) -replace '1\.4\.9', '1.4.10' | Set-Content cal_listener\__main__.py"
git add -A
git commit -m "v1.4.10: fix truncated pyproject.toml that broke v1.4.9 build"
git tag v1.4.10
git push origin main
git push origin v1.4.10
