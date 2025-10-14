# This uses certificate auth but you can probably do this with a more conventional log in too.


# Load your environment variables and auth parameters
Import-Module PnP.PowerShell

# Load environment variables from .env file
$envPath = "C:\Python_Scripts\N-Able\.env"
if (Test-Path $envPath) {
    Write-Host "Loading environment variables..." -ForegroundColor Gray
    Get-Content $envPath | ForEach-Object {
        if ($_ -match "^(.*?)=(.*)$") {
            $name = $matches[1].Trim()
            $value = $matches[2].Trim() -replace '^"|"$|^''|''$', ''
            # Use the Process scope explicitly to avoid errors
            [System.Environment]::SetEnvironmentVariable($name, $value, [System.EnvironmentVariableTarget]::Process)
        }
    }
}

# Auth parameters
$clientId = $env:clientId_cert
$thumbprint = $env:thumbprint
$tenantName = "yourdomain.com"

# Function to verify and create folders if needed
function Verify-SPOFolder {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl,
        [Parameter(Mandatory=$true)]
        [string]$FolderPath
    )
    
    Write-Host "Verifying folder path: $FolderPath" -ForegroundColor Yellow
    
    try {
        # Connect to site if not already connected
        $connection = Get-PnPConnection -ErrorAction SilentlyContinue
        if (-not $connection -or $connection.Url -ne $SiteUrl) {
            Connect-PnPOnline -Url $SiteUrl -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantName
        }
        
        # Split the path into segments
        $segments = $FolderPath.Trim('/').Split('/')
        $currentPath = $segments[0]  # Start with root (usually the document library)
        
        # First check if the root folder/library exists
        $rootFolder = Get-PnPFolder -Url $currentPath -ErrorAction SilentlyContinue
        
        if (-not $rootFolder) {
            Write-Host "❌ Root folder/library '$currentPath' does not exist!" -ForegroundColor Red
            return $false
        }
        
        # Now traverse and create each subfolder as needed
        for ($i = 1; $i -lt $segments.Count; $i++) {
            $folderName = $segments[$i]
            $parentPath = $currentPath
            $currentPath = "$currentPath/$folderName"
            
            $folder = Get-PnPFolder -Url $currentPath -ErrorAction SilentlyContinue
            
            if (-not $folder) {
                Write-Host "  Creating missing folder: $folderName" -ForegroundColor Yellow
                Add-PnPFolder -Name $folderName -Folder $parentPath
                
                # Verify the folder was created
                $folder = Get-PnPFolder -Url $currentPath -ErrorAction SilentlyContinue
                if (-not $folder) {
                    Write-Host "❌ Failed to create folder: $currentPath" -ForegroundColor Red
                    return $false
                }
            }
            else {
                Write-Host "  Folder exists: $currentPath" -ForegroundColor Green
            }
        }
        
        Write-Host "✓ Folder path verified successfully: $FolderPath" -ForegroundColor Green
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Host "❌ Error verifying folder path: $errorMsg" -ForegroundColor Red
        return $false
    }
}

# Proper implementation of Copy-SPOSingleFile
function Copy-SPOSingleFile {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SourceSite,
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        [Parameter(Mandatory=$true)]
        [string]$DestinationSite,
        [Parameter(Mandatory=$true)]
        [string]$DestinationFolder
    )
    
    Write-Host "`n==== FILE COPY OPERATION ====" -ForegroundColor Cyan
    Write-Host "Source site: $SourceSite" -ForegroundColor White
    Write-Host "Source path: $FilePath" -ForegroundColor White
    Write-Host "Destination: $DestinationSite : $DestinationFolder" -ForegroundColor White
    
    # Get just the filename
    $fileName = Split-Path -Path $FilePath -Leaf
    Write-Host "Target file: $fileName" -ForegroundColor White
    
    # Create temp folder for download
    $tempFolder = Join-Path $env:TEMP "SPSingleFileCopy_$(Get-Date -Format 'MMddHHmmss')"
    New-Item -Path $tempFolder -ItemType Directory -Force | Out-Null
    $tempPath = Join-Path $tempFolder $fileName
    Write-Host "Created temp folder: $tempFolder" -ForegroundColor Gray
    
    try {
        # 1. Connect to source site
        Write-Host "`n[Step 1/4] Connecting to source site..." -ForegroundColor Yellow
        Connect-PnPOnline -Url $SourceSite -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantName
        Write-Host "✓ Connected to source site" -ForegroundColor Green
        
        # 2. Download the file
        Write-Host "`n[Step 2/4] Downloading file..." -ForegroundColor Yellow
        Get-PnPFile -Url $FilePath -Path $tempFolder -Filename $fileName -AsFile -Force
        
        if (Test-Path $tempPath) {
            $fileInfo = Get-Item $tempPath
            Write-Host "✓ File downloaded successfully ($([math]::Round($fileInfo.Length/1KB, 2)) KB)" -ForegroundColor Green
            
            # 3. Connect to destination site
            Write-Host "`n[Step 3/4] Connecting to destination site..." -ForegroundColor Yellow
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            Connect-PnPOnline -Url $DestinationSite -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantName
            Write-Host "✓ Connected to destination site" -ForegroundColor Green
            
            # 4. Make sure the destination folder exists and upload
            Write-Host "`n[Step 4/4] Creating folders and uploading file..." -ForegroundColor Yellow
            if (Verify-SPOFolder -SiteUrl $DestinationSite -FolderPath $DestinationFolder) {
                # Upload file
                $file = Add-PnPFile -Path $tempPath -Folder $DestinationFolder
                Write-Host "✓ File copied successfully to $DestinationFolder" -ForegroundColor Green
                Write-Host "  File URL: $($file.ServerRelativeUrl)" -ForegroundColor Green
                return $true
            }
            else {
                Write-Host "❌ Destination folder verification failed" -ForegroundColor Red
                return $false
            }
        }
        else {
            Write-Host "❌ File download failed" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
        # Show detailed error info for troubleshooting
        Write-Host "`nDetailed error information:" -ForegroundColor Yellow
        Write-Host $_.Exception -ForegroundColor Red
        Write-Host "`nStack trace:" -ForegroundColor Yellow
        Write-Host $_.ScriptStackTrace -ForegroundColor Gray
        return $false
    }
    finally {
        # Cleanup
        if (Test-Path $tempFolder) {
            Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "`nTemporary files cleaned up" -ForegroundColor Gray
        }
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-Host "`n==== OPERATION COMPLETED ====`n" -ForegroundColor Cyan
    }
}

# Add a new function to verify the destination after copy
function Verify-FileInDestination {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl,
        [Parameter(Mandatory=$true)]
        [string]$FolderPath,
        [Parameter(Mandatory=$true)]
        [string]$FileName
    )
    
    Write-Host "`n==== VERIFYING FILE IN DESTINATION ====" -ForegroundColor Cyan
    
    try {
        # Connect to site
        Connect-PnPOnline -Url $SiteUrl -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantName
        
        # First try direct path to check if file exists
        $filePath = "$FolderPath/$FileName"
        Write-Host "Checking for file at: $filePath" -ForegroundColor Yellow
        
        try {
            $file = Get-PnPFile -Url $filePath -AsListItem -ErrorAction Stop
            Write-Host "✓ File verified! Found at: $filePath" -ForegroundColor Green
            Write-Host "  Last modified: $($file['Modified'])" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Host "❌ File not found at expected path" -ForegroundColor Red
            
            # Try searching for the file by name in the destination site
            Write-Host "`nSearching for file by name across site..." -ForegroundColor Yellow
            
            # Get all document libraries
            $docLibs = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 }
            foreach ($lib in $docLibs) {
                Write-Host "  Checking library: $($lib.Title)" -ForegroundColor Gray
                
                # Try CAML query to find the file
                $query = "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>$FileName</Value></Eq></Where></Query></View>"
                $items = Get-PnPListItem -List $lib -Query $query
                
                if ($items -and $items.Count -gt 0) {
                    foreach ($item in $items) {
                        $foundPath = $item["FileRef"]
                        Write-Host "✓ Found file in different location: $foundPath" -ForegroundColor Green
                        return $true
                    }
                }
            }
            
            Write-Host "❌ File not found anywhere in the site" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "Error during verification: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
    finally {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
}

# Simple parameters - adjust destination folder
$sourceSite = "https://rosecommunity.sharepoint.com/teams/RDS-TestSiteCleanupA"
$destSite = "https://rosecommunity.sharepoint.com/teams/RDS-TestSiteABC"
$destLib = "Shared Documents"

# IMPORTANT: Adjusted the destination folder structure to match what SharePoint expects
# The folder path needs to exactly match SharePoint's structure
$destFolder = "$destLib"  # Try putting it directly in the root of Shared Documents

$specificFilePath = "/Shared Documents/000 Final Documents/Annual Financials/Audits/SharePoint_Sites_20251012_135403.xlsx"
$fileName = Split-Path -Path $specificFilePath -Leaf  # Extract just the filename

# Optional: Add more detailed logging about the destination
Write-Host "`n==== OPERATION DETAILS =====" -ForegroundColor Cyan
Write-Host "Copying file from: $sourceSite" -ForegroundColor White
Write-Host "              to: $destSite" -ForegroundColor White
Write-Host "     File path: $specificFilePath" -ForegroundColor White
Write-Host "     Destination folder: $destFolder" -ForegroundColor White
Write-Host "==========================" -ForegroundColor Cyan

# Call the function directly with clear result indication
$result = Copy-SPOSingleFile -SourceSite $sourceSite -FilePath $specificFilePath -DestinationSite $destSite -DestinationFolder $destFolder

if ($result) {
    Write-Host "▶▶▶ SUCCESS: File was copied successfully! ◀◀◀" -ForegroundColor Green -BackgroundColor Black
    
    # Verify the file is actually there
    Write-Host "`nVerifying file is in destination site..." -ForegroundColor Yellow
    $verified = Verify-FileInDestination -SiteUrl $destSite -FolderPath $destFolder -FileName $fileName
    
    if ($verified) {
        Write-Host "`n✓ CONFIRMATION: File is confirmed to be in the destination site." -ForegroundColor Green -BackgroundColor Black
        Write-Host "📂 Please check: $destSite" -ForegroundColor Cyan
    }
    else {
        Write-Host "`n⚠️ WARNING: File copy reported success but verification failed." -ForegroundColor Yellow -BackgroundColor Black
        Write-Host "This could be due to:" -ForegroundColor Yellow
        Write-Host "1. Permissions issues - you might not have access to see the file" -ForegroundColor Yellow
        Write-Host "2. Delay in SharePoint's system - the file might appear after a few minutes" -ForegroundColor Yellow
        Write-Host "3. The file was saved to a different location than expected" -ForegroundColor Yellow
    }
} else {
    Write-Host "▶▶▶ FAILED: File copy operation did not complete successfully ◀◀◀" -ForegroundColor Red -BackgroundColor Black
}

# Additional code to list all files in the destination site's document libraries
Write-Host "`n==== LISTING ALL FILES IN DESTINATION SITE ====" -ForegroundColor Cyan
try {
    Connect-PnPOnline -Url $destSite -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantName
    $allLists = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false }
    
    Write-Host "Found $($allLists.Count) document libraries in destination site" -ForegroundColor White
    $totalFiles = 0
    
    foreach ($list in $allLists) {
        Write-Host "`nChecking library: $($list.Title)" -ForegroundColor Cyan
        $query = "<View Scope='RecursiveAll'><RowLimit>5000</RowLimit></View>"
        $items = Get-PnPListItem -List $list -Query $query -PageSize 100
        $files = $items | Where-Object { $_["FileSystemObjectType"] -eq "File" }
        
        if ($files -and $files.Count -gt 0) {
            Write-Host "  Found $($files.Count) files in $($list.Title)" -ForegroundColor Green
            $totalFiles += $files.Count
            
            foreach ($file in $files | Sort-Object { $_["FileRef"] }) {
                $filePath = $file["FileRef"]
                $fileModified = $file["Modified"]
                Write-Host "    • $filePath (Modified: $fileModified)" -ForegroundColor Gray
                
                # Highlight our target file if found
                if ($filePath -match $fileName) {
                    Write-Host "    ✅ THIS IS OUR TARGET FILE ✅" -ForegroundColor Green -BackgroundColor Black
                }
            }
        } else {
            Write-Host "  No files found in $($list.Title)" -ForegroundColor Yellow
        }
    }
    
    if ($totalFiles -eq 0) {
        Write-Host "`n❌ NO FILES FOUND in any document library in the destination site!" -ForegroundColor Red -BackgroundColor Black
    } else {
        Write-Host "`nTotal files in destination site: $totalFiles" -ForegroundColor Cyan
    }
}
catch {
    Write-Host "Error listing files: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
