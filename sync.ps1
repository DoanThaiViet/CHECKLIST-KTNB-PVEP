# Tu dong: sinh lai checklists.json tu file Excel M365 roi day len GitHub.
# Chay dinh ky boi Task Scheduler.
$ErrorActionPreference = 'Stop'
$repo = 'C:\Users\vietdt\checklist-sync'
$node = 'C:\Program Files\nodejs\node.exe'
$git  = 'C:\Users\vietdt\AppData\Local\Programs\Git\cmd\git.exe'
Set-Location $repo

# 1. Sinh lai checklists.json tu Excel moi nhat (OneDrive da dong bo ve may)
& $node "$repo\auto-update.js"
if ($LASTEXITCODE -ne 0) { Write-Output "Convert failed"; exit 1 }

# 2. Neu khong co thay doi thi dung
& $git add checklists.json
& $git diff --staged --quiet
if ($LASTEXITCODE -eq 0) { Write-Output "No change"; exit 0 }

# 3. Commit + push
$stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
& $git commit -m "Auto-update checklist $stamp" | Out-Null
& $git push
Write-Output "Pushed at $stamp"
