# Script to retrieve all SharePoint sites with names and URLs - Improved version for handling permission issues
# Look for 'fix' to search for variables that need to be changed
# Load environment variables from .env file
$envPath = "C:\yourenvlocation\.env"

if (Test-Path $envPath) {
    Write-Host "Loading environment variables from $envPath"
    Get-Content $envPath | ForEach-Object {
        if ($_ -match "^(.*?)=(.*)$") {
            $name = $matches[1].Trim()
            $value = $matches[2].Trim() -replace '^"|"$','' -replace "^'|'$",''
            [System.Environment]::SetEnvironmentVariable($name, $value)
        }
    }
}

# Get authentication parameters from environment
$clientId = $env:clientId_cert#fix
$thumbprint = $env:thumbprint #fix 
$tenantName = "yoursite.com" #fix
$adminSiteUrl = "https://yoursite-admin.sharepoint.com"

# Initialize arrays to store site information
$siteInfo = @()
$errors = @()
$progressCounter = 0

# Function to test if we have access to a site
function Test-SiteAccess {
    param(
        [string]$SiteUrl
    )
    
    try {
        Connect-PnPOnline -Url $SiteUrl -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantName -WarningAction SilentlyContinue -ErrorAction Stop
        $web = Get-PnPWeb -Includes Id, ServerRelativeUrl -ErrorAction Stop
        $siteId = $web.Id.ToString()
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        return @{
            HasAccess = $true
            IsArchived = $false
            Message = "Access OK"
            SiteId = $siteId
            ServerRelativeUrl = $web.ServerRelativeUrl
        }
    }
    catch {
        # Try to determine why access failed
        $errorMessage = $_.Exception.Message
        $isArchived = $errorMessage -match "archived" -or $errorMessage -match "read-only"
        
        return @{
            HasAccess = $false
            IsArchived = $isArchived
            Message = $errorMessage
        }
    }
}

# Connect to SharePoint Admin site
try {
    Write-Host "Connecting to SharePoint Admin site..." -ForegroundColor Yellow
    Connect-PnPOnline -Url $adminSiteUrl -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantName -WarningAction SilentlyContinue
    
    # Test if we have admin access
    try {
        # Get all site collections (we might not have access to all of them)
        Write-Host "Retrieving site collections..." -ForegroundColor Cyan
        $siteCollections = Get-PnPTenantSite -ErrorAction Stop
        $hasAdminAccess = $true
        $totalSites = $siteCollections.Count
        Write-Host "Found $totalSites site collections at admin level" -ForegroundColor Green
    }
    catch {
        Write-Host "Warning: Limited admin access. Will retrieve accessible sites only." -ForegroundColor Yellow
        $hasAdminAccess = $false
        # If we can't get tenant sites, we'll use a simpler approach
        Write-Host "Using alternative method to find sites..." -ForegroundColor Yellow
    }
    
    # If we have admin access, proceed with the full site collection retrieval
    if ($hasAdminAccess) {
        # Basic site info retrieval without accessing site content (avoids most permission issues)
        foreach ($site in $siteCollections) {
            $progressCounter++
            $percentComplete = [math]::Round(($progressCounter / $totalSites) * 100, 2)
            Write-Progress -Activity "Retrieving Site Information" -Status "$progressCounter of $totalSites ($percentComplete%)" -PercentComplete $percentComplete
            
            # Check if site is locked or archived
            $isLocked = $site.LockState -ne "Unlock"
            $isReadOnly = $site.ReadOnly
            
            # Test if we can access the site
            $accessResult = Test-SiteAccess -SiteUrl $site.Url
            
            # Add the site to our list with basic info we already have
            $siteObject = [PSCustomObject]@{
                Title = $site.Title
                URL = $site.Url
                Type = $site.Template
                SiteId = $accessResult.SiteId
                ServerRelativeUrl = $accessResult.ServerRelativeUrl
                Owner = $site.Owner
                Created = $site.Created
                LastModified = $null
                Description = $null
                HasAccess = $accessResult.HasAccess
                IsLocked = $isLocked
                IsReadOnly = $isReadOnly
                IsArchived = $accessResult.IsArchived
                Status = if ($isLocked) { "Locked" } elseif ($isReadOnly) { "Read-Only" } elseif ($accessResult.IsArchived) { "Archived" } elseif (!$accessResult.HasAccess) { "No Access" } else { "Active" }
                AccessMessage = $accessResult.Message
                ListCount = $null
                HasSubsites = $null
            }
            
            # Record any access errors
            if (!$accessResult.HasAccess) {
                $errors += "Access denied to $($site.Url): $($accessResult.Message)"
            }
            
            $siteInfo += $siteObject
        }
    }
    else {
        # Alternative approach - search for sites we have access to
        Write-Host "Attempting to find sites through search..." -ForegroundColor Yellow
        
        # Try to find sites through search
        try {
            # Since we're connected to admin, disconnect first
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            
            # Connect to the root site, which most users have access to
            $rootSiteUrl = "https://yoursite.sharepoint.com"
            Connect-PnPOnline -Url $rootSiteUrl -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantName -WarningAction SilentlyContinue
            
            # Get hub sites (less permissions required)
            $hubSites = Get-PnPHubSite -ErrorAction SilentlyContinue
            if ($hubSites -and $hubSites.Count -gt 0) {
                Write-Host "Found $($hubSites.Count) hub sites" -ForegroundColor Green
                foreach ($hub in $hubSites) {
                    $siteObject = [PSCustomObject]@{
                        Title = $hub.Title
                        URL = $hub.SiteUrl
                        Type = "Hub Site"
                        Created = $null
                        LastModified = $null
                        Description = $hub.Description
                        HasAccess = (Test-SiteAccess -SiteUrl $hub.SiteUrl)
                        ListCount = $null
                        HasSubsites = $null
                    }
                    $siteInfo += $siteObject
                }
            }
            
            # Try to get site information through Search (works with less permissions)
            try {
                # Use search to find sites (this works with regular permissions)
                $searchQuery = "*"
                $searchResults = Submit-PnPSearchQuery -Query $searchQuery -SourceId "yoursourceid" -ErrorAction Stop #fix id

                
                if ($searchResults -and $searchResults.ResultRows) {
                    Write-Host "Found sites through search: $($searchResults.ResultRows.Count)" -ForegroundColor Green
                    
                    foreach ($result in $searchResults.ResultRows) {
                        $siteObject = [PSCustomObject]@{
                            Title = $result.Title
                            URL = $result.SPSiteUrl
                            Type = "Found via Search"
                            Created = $null
                            LastModified = $null
                            Description = $result.Description
                            HasAccess = $true # If we found it via search, we likely have some level of access
                            ListCount = $null
                            HasSubsites = $null
                        }
                        $siteInfo += $siteObject
                    }
                }
            }
            catch {
                Write-Host "Could not retrieve sites through search: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "Alternative site discovery failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    # Export results to CSV with additional sorting options
    if ($siteInfo.Count -gt 0) {
        # Remove duplicate URLs
        $siteInfo = $siteInfo | Sort-Object URL -Unique
        
        # Add additional stats
        $totalSites = $siteInfo.Count
        $accessibleSites = ($siteInfo | Where-Object { $_.HasAccess -eq $true }).Count
        $deniedSites = ($siteInfo | Where-Object { $_.HasAccess -eq $false }).Count
        $archivedSites = ($siteInfo | Where-Object { $_.IsArchived -eq $true }).Count
        $lockedSites = ($siteInfo | Where-Object { $_.IsLocked -eq $true }).Count
        $readOnlySites = ($siteInfo | Where-Object { $_.IsReadOnly -eq $true }).Count
        
        Write-Host "`nSite Statistics:" -ForegroundColor Cyan
        Write-Host "  Total Sites: $totalSites" -ForegroundColor White
        Write-Host "  Accessible Sites: $accessibleSites" -ForegroundColor Green
        Write-Host "  Denied Sites: $deniedSites" -ForegroundColor Yellow
        Write-Host "  Archived Sites: $archivedSites" -ForegroundColor Yellow
        Write-Host "  Locked Sites: $lockedSites" -ForegroundColor Yellow
        Write-Host "  Read-Only Sites: $readOnlySites" -ForegroundColor Yellow
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        
        # Create a more detailed CSV with additional columns useful for site copying
        $csvPath = "C:\Powershell_scripts\SharePoint_Sites_$timestamp.csv"#fixpath
        $siteInfo | 
            Select-Object Title, URL, SiteId, ServerRelativeUrl, Type, Status, Owner, Created, 
                        IsLocked, IsReadOnly, IsArchived, HasAccess, AccessMessage, ListCount, HasSubsites |
            Export-Csv -Path $csvPath -NoTypeInformation
        
        # Create a simple list for copying purposes (accessible sites only)
        $copyCsvPath = "C:\Python_Scripts\api_calls\JRCo_data\Powershell_scripts\SharePoint_Sites_ForCopy_$timestamp.csv"
        $siteInfo | 
            Where-Object { $_.HasAccess -eq $true } |
            Sort-Object Title |
            Select-Object Title, URL, SiteId, @{Name="SourceUrl";Expression={$_.URL}}, @{Name="TargetUrl";Expression={""}} |
            Export-Csv -Path $copyCsvPath -NoTypeInformation
        
        # Create a list of denied sites
        $deniedCsvPath = "C:\Powershell_scripts\SharePoint_Sites_Denied_$timestamp.csv"#fixpath
        $siteInfo | 
            Where-Object { $_.HasAccess -eq $false } |
            Sort-Object Title |
            Select-Object Title, URL, Type, Status, Owner, Created, IsLocked, IsReadOnly, IsArchived, AccessMessage |
            Export-Csv -Path $deniedCsvPath -NoTypeInformation
        
        Write-Host "`nFound $($siteInfo.Count) unique sites" -ForegroundColor Green
        Write-Host "Full site information exported to: $csvPath" -ForegroundColor Green
        Write-Host "Simplified copy template exported to: $copyCsvPath" -ForegroundColor Green
        if ($deniedSites -gt 0) {
            Write-Host "Denied sites ($deniedSites) exported to: $deniedCsvPath" -ForegroundColor Yellow
        }
        
        # Display results in grid view
        $siteInfo | Out-GridView -Title "SharePoint Sites"
    }
    else {
        Write-Host "No site information found. Check permissions or connection." -ForegroundColor Red
    }
    
    # Display error summary if any
    if ($errors.Count -gt 0) {
        $groupedErrors = $errors | Group-Object | Select-Object Name, Count
        Write-Host "`nEncountered $($errors.Count) errors during retrieval" -ForegroundColor Yellow
        
        # Export full error list to CSV
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $errorsCsvPath = ".\Powershell_scripts\SharePoint_Errors_$timestamp.csv"#fixpath
        $errors | ForEach-Object {
            [PSCustomObject]@{
                Error = $_
            }
        } | Export-Csv -Path $errorsCsvPath -NoTypeInformation
        
        Write-Host "Detailed error list exported to: $errorsCsvPath" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "Failed to connect to SharePoint Admin site: $($_.Exception.Message)" -ForegroundColor Red
    
    # Attempt alternative approach by connecting directly to root site
    try {
        Write-Host "Attempting to connect to root site instead..." -ForegroundColor Yellow
        Connect-PnPOnline -Url "https://yoursite.sharepoint.com" -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantName -WarningAction SilentlyContinue
        
        # Try to get web properties
        $web = Get-PnPWeb -Includes Title, Created, LastItemModifiedDate
        
        Write-Host "Connected to root site: $($web.Title)" -ForegroundColor Green
        Write-Host "Try running the script with user credentials that have higher permissions." -ForegroundColor Yellow
    }
    catch {
        Write-Host "Could not connect to any SharePoint site. Please check credentials and permissions." -ForegroundColor Red
    }
}
finally {
    # Disconnect from any open connections
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}

Write-Host "`nScript execution completed." -ForegroundColor Green
