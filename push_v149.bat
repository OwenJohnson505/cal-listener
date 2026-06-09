@echo off
cd /d C:\Users\jowen\cal-listener
git add -A
git commit -m "v1.4.9: point bundled Playwright at user-profile ms-playwright cache"
git tag v1.4.9
git push origin main
git push origin v1.4.9
