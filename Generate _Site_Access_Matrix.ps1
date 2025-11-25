# ==================================================================================
# SCRIPT: Generate Site Access Matrix (Fixed for Null URLs)
# ==================================================================================

# --- 1. CONFIGURATION ---
$envPath = "C:\Python_Scripts\N-Able\.env"
$InputExcel = "C:\Python_Scripts\api_calls\JRCo_data\reports\Sites_For_Permissioning_REVIEW.xlsx"
$OutputExcel = "C:\Python_Scripts\api_calls\JRCo_data\reports\PowerShell_Permission_Matrix_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx"

# Load Environment Variables
if (Test-Path $envPath) {
    Get-Content $envPath | ForEach-Object {
        if ($_ -match "^(.*?)=(.*)$") {
            [System.Environment]::SetEnvironmentVariable($matches[1].Trim(), $matches[2].Trim())
        }
    }
}

$clientId = $env:clientId_jrco_cert
$thumbprint = $env:thumbprint_jrco
$tenantName = "rosecommunity.com"

# Check modules
if (-not (Get-Module -ListAvailable -Name ImportExcel)) { Write-Error "Install-Module ImportExcel required"; exit }
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) { Write-Error "Install-Module PnP.PowerShell required"; exit }

# --- 2. LOAD SITES ---
Write-Host "Reading Site List: $InputExcel" -ForegroundColor Cyan
try {
    $sites = Import-Excel $InputExcel
    if (-not $sites) { Write-Error "Excel file is empty."; exit }
    
    # DEBUG: Print the column names found to help troubleshoot
    $columns = $sites[0].PSObject.Properties.Name
    Write-Host "ℹ️  Found Columns in Excel: $($columns -join ', ')" -ForegroundColor Gray
}
catch {
    Write-Error "Could not read Excel file: $($_.Exception.Message)"
    exit
}

# --- 3. DATA COLLECTION ---
$MasterUserList = @{} 
$ProcessedSites = @() 

foreach ($row in $sites) {
    # 1. Try to find the URL property (Case-insensitive check)
    $siteUrl = $null
    if ($row.PSObject.Properties['webUrl']) { $siteUrl = $row.webUrl }
    elseif ($row.PSObject.Properties['Url']) { $siteUrl = $row.Url }
    elseif ($row.PSObject.Properties['SiteUrl']) { $siteUrl = $row.SiteUrl }
    
    # 2. CRITICAL FIX: Skip if URL is empty (handles blank rows)
    if ([string]::IsNullOrWhiteSpace($siteUrl)) { 
        # Write-Warning "Skipping a row with no URL..."
        continue 
    }

    # 3. Determine Site Name safely
    $siteName = $null
    if ($row.PSObject.Properties['name'] -and -not [string]::IsNullOrWhiteSpace($row.name)) { 
        $siteName = $row.name 
    } 
    elseif ($row.PSObject.Properties['SiteName'] -and -not [string]::IsNullOrWhiteSpace($row.SiteName)) { 
        $siteName = $row.SiteName 
    } 
    else { 
        # Only split if we actually have a URL
        $siteName = $siteUrl.Split('/')[-1] 
    }
    
    Write-Host "Processing: $siteName" -NoNewline
    $ProcessedSites += $siteName
    
    try {
        # Connect
        Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantName -ErrorAction Stop
        
        $siteUsers = @()
        
        # Get Owners
        $ownerGroup = Get-PnPGroup -AssociatedOwnerGroup -ErrorAction SilentlyContinue
        if ($ownerGroup) {
            $owners = Get-PnPGroupMember -Identity $ownerGroup.Id -ErrorAction SilentlyContinue
            foreach ($u in $owners) { $siteUsers += [PSCustomObject]@{ Email = $u.Email; Name = $u.Title; Role = "Owner" } }
        }
        
        # Get Members
        $memberGroup = Get-PnPGroup -AssociatedMemberGroup -ErrorAction SilentlyContinue
        if ($memberGroup) {
            $members = Get-PnPGroupMember -Identity $memberGroup.Id -ErrorAction SilentlyContinue
            foreach ($u in $members) { $siteUsers += [PSCustomObject]@{ Email = $u.Email; Name = $u.Title; Role = "Member" } }
        }
        
        Write-Host " -> Found $(($siteUsers | Select-Object -Unique Email).Count) users" -ForegroundColor Green
        
        # Process Users
        foreach ($user in $siteUsers) {
            if ([string]::IsNullOrWhiteSpace($user.Email) -or $user.Email -like "*@$tenantName") { continue }
            
            $email = $user.Email
            
            if (-not $MasterUserList.ContainsKey($email)) {
                $MasterUserList[$email] = @{ Name = $user.Name; Email = $email; Permissions = @{} }
            }
            
            # Prioritize Owner over Member if user is in both
            $current = $MasterUserList[$email].Permissions[$siteName]
            if ($user.Role -eq "Owner") {
                $MasterUserList[$email].Permissions[$siteName] = "Owner"
            }
            elseif (-not $current) {
                $MasterUserList[$email].Permissions[$siteName] = "Member"
            }
        }
    }
    catch {
        Write-Host " -> ❌ Failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# --- 4. EXPORT ---
Write-Host "`nGenerating Matrix..." -ForegroundColor Cyan

$MatrixData = @()
foreach ($email in $MasterUserList.Keys) {
    $userData = $MasterUserList[$email]
    $rowObj = [ordered]@{ "User Name" = $userData.Name; "User Email" = $userData.Email }
    
    foreach ($site in $ProcessedSites) {
        $rowObj[$site] = if ($userData.Permissions.ContainsKey($site)) { $userData.Permissions[$site] } else { "" }
    }
    $MatrixData += [PSCustomObject]$rowObj
}

if ($MatrixData.Count -gt 0) {
    $MatrixData | Export-Excel -Path $OutputExcel -AutoSize -FreezeTopRow
    Write-Host "`n✅ Success! Matrix saved to:" -ForegroundColor Green
    Write-Host "   $OutputExcel" -ForegroundColor Gray
}
else {
    Write-Host "❌ No data collected. Check column headers in the 'Found Columns' message above." -ForegroundColor Red
}
