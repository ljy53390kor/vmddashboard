# vmddata deploy script
$ErrorActionPreference = "Stop"

Write-Host "vmddata deploy start..." -ForegroundColor Cyan

$vmddataUrl = "https://gitlab.tde.sktelecom.com/PGPRIVATE/playground-vmddata-app.git"
$remotes = git remote
if ($remotes -notcontains "vmddata") {
    git remote add vmddata $vmddataUrl
    Write-Host "vmddata remote added" -ForegroundColor Green
} else {
    git remote set-url vmddata $vmddataUrl
    Write-Host "vmddata remote updated" -ForegroundColor Green
}

Copy-Item Diyfile.yaml Diyfile.vmddashboard.yaml -Force
Copy-Item Diyfile.vmddata.yaml Diyfile.yaml -Force
Write-Host "Diyfile.yaml replaced" -ForegroundColor Green

git add Diyfile.yaml
git commit -m "deploy: vmddata playground config" --allow-empty

git push vmddata main:master --force
Write-Host "vmddata push done" -ForegroundColor Green

Copy-Item Diyfile.vmddashboard.yaml Diyfile.yaml -Force
Remove-Item Diyfile.vmddashboard.yaml
git add Diyfile.yaml
git commit -m "restore: vmddashboard diyfile" --allow-empty
git push origin

Write-Host ""
Write-Host "Done! After build, access Studio at:" -ForegroundColor Green
Write-Host "https://playground.idcube.sktelecom.com/vmddata-app/studio/" -ForegroundColor Cyan
Write-Host "Settings > API > Project API keys > anon key" -ForegroundColor Yellow
