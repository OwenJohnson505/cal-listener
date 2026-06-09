@echo off
cd /d C:\Users\jowen\cal-listener
powershell -Command "(Get-Content cal_listener\__main__.py) -replace '1\.4\.7', '1.4.8' | Set-Content cal_listener\__main__.py"
powershell -Command "(Get-Content pyproject.toml) -replace 'version = \"1\.4\.7\"', 'version = \"1.4.8\"' | Set-Content pyproject.toml"
git add -A
git commit -m "v1.4.8: BIG BANG — wire every remaining cb_*, consignment_cross_ref, file-report handlers (no more stubs)"
git tag v1.4.8
git push origin main
git push origin v1.4.8
