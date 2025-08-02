
foreach ($domain in @(
"api.automox.com"
"console.automox.com"
"rc.automox.com"
"ct.automox.com"
"storage-cdn.prod.automox.com"
"worklet-signing.prod.automox.com"
"installation-reporting-service.prod.automox.com"
"llm.automox.com"
"downloadexport.automox.com"
"policyreport.automox.com"
"download-export-cdn.prod.automox.com"
"rtt.automox.com"
"agent-content-cdn.automox.com"
)) {
    Write-Host "Testing $domain..." -ForegroundColor Cyan
    Test-NetConnection -ComputerName $domain -Port 443
    Write-Host "`n"
}
