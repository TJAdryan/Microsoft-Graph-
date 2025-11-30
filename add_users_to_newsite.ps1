#adds a group 
# Add_Matrix_Users_To_newsite.ps1
# 1. Reads Site_Access_Matrix_REVIEW.xlsx
# 2. Adds every user found in the list to "sitename" as a MEMBER
# 3. FIX: Corrected Connection URL (Removed double .com)

# --- CONFIGURATION ---
$envPath = "C:\Python_Scripts\N-Able\.env"
$MatrixPath = "C:\Python_Scripts\api_calls\reports\Site_Access_Matrix_REVIEW.xlsx"
$TargetTeamName = ""
$TenantName = ""
# FIXED URL: Hardcoded to prevent variable expansion errors
$RootSiteUrl = "https://yoursite.sharepoint.com"
# ---------------------

# 1. Load Environment
if (Test-Path $envPath) {
    Write-Host "Loading .env..." -ForegroundColor Gray
    Get-Content $envPath | Where-Object { $_ -match "=" -and $_ -notmatch "^#" } | ForEach-Object {
        $parts = $_ -split "=", 2
        if ($parts.Count -eq 2) { Set-Item -Path "env:$($parts[0].Trim())" -Value $parts[1].Trim().Trim('"').Trim("'") }
    }
}
$clientId = $env:clientId_jrco_cert
$thumbprint = $env:thumbprint_jrco

# 2. Import Data
if (-not (Test-Path $MatrixPath)) { 
    Write-Error " File not found: $MatrixPath"
    exit 
}

$users = Import-Excel $MatrixPath
Write-Host " Reading Matrix: $(Split-Path $MatrixPath -Leaf)" -ForegroundColor Cyan
Write-Host "   Found $(@($users).Count) users." -ForegroundColor Cyan

# 3. Connect & Get Target Team
try {
    Write-Host "Connecting to Tenant ($RootSiteUrl)..." -ForegroundColor Gray
    # Fixed Connection String
    Connect-PnPOnline -Url $RootSiteUrl -ClientId $clientId -Thumbprint $thumbprint -Tenant $TenantName -ErrorAction Stop
    
    Write-Host "Looking for Team: '$TargetTeamName'..." -NoNewline
    $targetGroup = Get-PnPMicrosoft365Group | Where-Object { $_.DisplayName -eq $TargetTeamName } | Select-Object -First 1
    
    if (-not $targetGroup) { 
        Write-Host " Not Found!" -ForegroundColor Red
        Write-Error "Could not find a team named '$TargetTeamName'"
        exit 
    }
    Write-Host " Found (ID: $($targetGroup.Id))" -ForegroundColor Green
    
} catch {
    Write-Error "Connection Failed: $($_.Exception.Message)"; exit
}

# 4. Add Users Loop
foreach ($row in $users) {
    $email = $row.user_email
    
    if ([string]::IsNullOrWhiteSpace($email)) { continue }

    Write-Host "------------------------------------------------"
    Write-Host "User: $($row.user_displayName) ($email)" -ForegroundColor Yellow
    
    try {
        Write-Host "  Adding to '$TargetTeamName'..." -NoNewline
        
        # Add as Member
        Add-PnPMicrosoft365GroupMember -Identity $targetGroup.Id -Users $email -ErrorAction Stop
        
        Write-Host " Success." -ForegroundColor Green
    }
    catch {
        if ($_.Exception.Message -like "*already*") {
            Write-Host " Already a member." -ForegroundColor Gray
        } else {
            Write-Host " Failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

Write-Host "`n✅ Complete." -ForegroundColor Green
