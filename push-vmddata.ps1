# vmddata Playground 앱 배포 스크립트
$ErrorActionPreference = "Stop"

Write-Host "vmddata 배포 시작..." -ForegroundColor Cyan

# 1. vmddata remote 없으면 추가
$remotes = git remote
if ($remotes -notcontains "vmddata") {
    git remote add vmddata https://gitlab.tde.sktelecom.com/PGPRIVATE/playground-vmddata.git
    Write-Host "vmddata remote 추가됨" -ForegroundColor Green
}

# 2. Diyfile.yaml 백업 후 vmddata용으로 교체
Copy-Item Diyfile.yaml Diyfile.vmddashboard.yaml -Force
Copy-Item Diyfile.vmddata.yaml Diyfile.yaml -Force
Write-Host "Diyfile.yaml 교체됨 (vmddata용)" -ForegroundColor Green

# 3. 커밋
git add Diyfile.yaml
git commit -m "deploy: vmddata playground config" --allow-empty

# 4. vmddata repo에 push
git push vmddata main:master --force
Write-Host "vmddata push 완료" -ForegroundColor Green

# 5. Diyfile.yaml 원복
Copy-Item Diyfile.vmddashboard.yaml Diyfile.yaml -Force
Remove-Item Diyfile.vmddashboard.yaml
git add Diyfile.yaml
git commit -m "restore: vmddashboard diyfile" --allow-empty
git push origin

Write-Host ""
Write-Host "✅ 배포 완료!" -ForegroundColor Green
Write-Host "빌드 완료 후 아래 URL에서 Studio 접속해서 anon key 확인:" -ForegroundColor Yellow
Write-Host "https://playground.idcube.sktelecom.com/vmddata/studio/" -ForegroundColor Cyan
Write-Host "Settings > API > Project API keys > anon key 복사" -ForegroundColor Yellow
