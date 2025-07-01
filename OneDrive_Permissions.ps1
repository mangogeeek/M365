<#
.SYNOPSIS
Assigns OneDrive-related permissions to an Azure AD application (without admin consent)
#>

# Function to validate GUID format
function Test-IsGuid {
    param([string]$GuidString)
    $guid = [System.Guid]::empty
    return [System.Guid]::TryParse($GuidString, [ref]$guid)
}

# Connect to Microsoft Graph
try {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes "Application.ReadWrite.All" -NoWelcome -ErrorAction Stop
    Write-Host "Connected successfully." -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
    exit
}

# Prompt for Client ID with validation
do {
    $clientId = Read-Host -Prompt "`nEnter your Application (Client) ID"
    if (-not (Test-IsGuid $clientId)) {
        Write-Host "Invalid Client ID format. Please enter a valid GUID." -ForegroundColor Red
    }
} until (Test-IsGuid $clientId)

# Get service principals
try {
    Write-Host "`nFetching application details..." -ForegroundColor Cyan
    $app = Get-MgServicePrincipal -Filter "appId eq '$clientId'" -ErrorAction Stop
    if (-not $app) {
        Write-Host "Application with Client ID $clientId not found." -ForegroundColor Red
        exit
    }
    
    $graphSpn = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -ErrorAction Stop
    Write-Host "Found application: $($app.DisplayName)" -ForegroundColor Green
}
catch {
    Write-Host "Error fetching application details: $_" -ForegroundColor Red
    exit
}

# Define OneDrive-related permissions
$permissions = @(
    # Application permissions (Role)
    @{ Name = "Files.ReadWrite.All"; Type = "Role"; Id = "75359482-378d-4052-8f01-80520e7db3cd" },
    @{ Name = "User.Read.All"; Type = "Role"; Id = "df021288-bdef-4463-88db-98f22de89214" },
    @{ Name = "Application.Read.All"; Type = "Role"; Id = "9a5d68dd-52b0-4cc2-bd40-abcf44ac3a30" }
)

# Display permissions summary
Write-Host "`nThe following OneDrive permissions will be assigned:" -ForegroundColor Cyan
$permissions | Format-Table Name, Type -AutoSize
Write-Host "`nNote: These are application permissions that require admin consent." -ForegroundColor Yellow

# Confirm before proceeding
$confirmation = Read-Host "`nDo you want to proceed with these permissions? (Y/N)"
if ($confirmation -ne 'Y') {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    exit
}

# Prepare the resource access object
$resourceAccess = $permissions | ForEach-Object {
    @{
        id   = $_.Id
        type = $_.Type
    }
}

# Update the application with permissions
try {
    Write-Host "`nAssigning permissions to application..." -ForegroundColor Cyan
    
    # Get current permissions to merge (not overwrite)
    $currentAccess = (Get-MgApplication -ApplicationId $app.AppId).RequiredResourceAccess | 
        Where-Object { $_.ResourceAppId -eq $graphSpn.AppId }
    
    if ($currentAccess) {
        # Merge existing permissions with new ones
        $combinedAccess = $currentAccess.ResourceAccess + $resourceAccess
        $uniqueAccess = $combinedAccess | Group-Object id | ForEach-Object { $_.Group[0] }
        
        Update-MgApplication -ApplicationId $app.AppId `
            -RequiredResourceAccess @(
                @{
                    resourceAppId  = $graphSpn.AppId
                    resourceAccess = $uniqueAccess
                }
            ) -ErrorAction Stop
    }
    else {
        Update-MgApplication -ApplicationId $app.AppId `
            -RequiredResourceAccess @(
                @{
                    resourceAppId  = $graphSpn.AppId
                    resourceAccess = $resourceAccess
                }
            ) -ErrorAction Stop
    }
    
    Write-Host "Successfully assigned permissions." -ForegroundColor Green
}
catch {
    Write-Host "Error assigning permissions: $_" -ForegroundColor Red
    exit
}

# Completion message
Write-Host "`nProcess completed. Summary:" -ForegroundColor Cyan
Write-Host "- Assigned $($permissions.Count) application permissions" -ForegroundColor White

Write-Host "`nNext steps:" -ForegroundColor Green
Write-Host "1. An Azure AD administrator must grant admin consent for these permissions:" -ForegroundColor Yellow
Write-Host "   - Files.ReadWrite.All (Access all files user can access)" -ForegroundColor Yellow
Write-Host "   - User.Read.All (Read all users' full profiles)" -ForegroundColor Yellow
Write-Host "   - Application.Read.All (Read applications)" -ForegroundColor Yellow
Write-Host "2. Verify in Azure Portal under:" -ForegroundColor Green
Write-Host "   Azure AD > App Registrations > $($app.DisplayName) > API permissions" -ForegroundColor Green

# Disconnect (optional)
# Disconnect-MgGraph
