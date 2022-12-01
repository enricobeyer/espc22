# Authenticate
Connect-MgGraph -ClientId "75a9bbec-4b9c-40db-8135-b88813cdd631" -TenantId "dcd0637d-e799-4e34-9ce1-179f1f036bc3" -CertificateName "CN=PowerShellScriptCert"

Write-Host "USERS:"
Write-Host "======================================================"
# List first 50 users
Get-MgUser -Property "id,displayName" -PageSize 50 | Format-Table DisplayName, Id

Write-Host "GROUPS:"
Write-Host "======================================================"
# List first 50 groups
Get-MgGroup -Property "id,displayName" -PageSize 50 | Format-Table DisplayName, Id

# Disconnect
Disconnect-MgGraph