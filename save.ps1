# Morning Arena — שמור ודחוף לגיטהאב
# הרצה: .\save.ps1 "תיאור השינוי"

param([string]$msg = "update")

Set-Location $PSScriptRoot
git add Code.gs Index.html
git commit -m $msg
git push
Write-Host "`n✅ נשמר וגובה ב-GitHub!" -ForegroundColor Green
