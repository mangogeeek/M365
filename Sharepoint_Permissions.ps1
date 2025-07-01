<#
.SYNOPSIS
Assigns SharePoint and related permissions to an Azure AD application (without admin consent)
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
    $sharepointSpn = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0ff1-ce00-000000000000'" -ErrorAction Stop
    $mipSpn = Get-MgServicePrincipal -Filter "appId eq '870c4f2e-85b6-4d43-bdda-6ed9a579b725'" -ErrorAction Stop
    $officeMgmtSpn = Get-MgServicePrincipal -Filter "appId eq 'c5393580-f805-4401-95e8-94b7a6ef2fc2'" -ErrorAction Stop
    
    Write-Host "Found application: $($app.DisplayName)" -ForegroundColor Green
}
catch {
    Write-Host "Error fetching service principals: $_" -ForegroundColor Red
    exit
}

# Define all permissions for each service
$permissions = @(
    # Microsoft Graph Permissions
    @{ Service = "Microsoft Graph"; Name = "Directory.Read.All"; Type = "Role"; Id = "7ab1d382-f21e-4acd-a863-ba3e13f7da61" },
    @{ Service = "Microsoft Graph"; Name = "Files.ReadWrite.All"; Type = "Role"; Id = "75359482-378d-4052-8f01-80520e7db3cd" },
    @{ Service = "Microsoft Graph"; Name = "Group.Read.All"; Type = "Role"; Id = "5b567255-7703-4780-807c-7be8301ae99b" },
    @{ Service = "Microsoft Graph"; Name = "Reports.Read.All"; Type = "Role"; Id = "230c1aed-a721-4c5d-9cb4-a90514e508ef" },
    @{ Service = "Microsoft Graph"; Name = "Sites.FullControl.All"; Type = "Role"; Id = "a82116e5-55eb-4c41-a434-62fe8a61c773" },
    @{ Service = "Microsoft Graph"; Name = "Sites.Manage.All"; Type = "Role"; Id = "65e50fdc-43b7-4915-933e-e8138f11f40a" },
    @{ Service = "Microsoft Graph"; Name = "Sites.ReadWrite.All"; Type = "Role"; Id = "9492366f-7969-46a4-8d15-ed1a20078fff" },
    @{ Service = "Microsoft Graph"; Name = "User.Read.All"; Type = "Role"; Id = "df021288-bdef-4463-88db-98f22de89214" },
    
    # SharePoint Online Permissions
    @{ Service = "SharePoint Online"; Name = "Sites.FullControl.All"; Type = "Role"; Id = "678536fe-1083-478a-9c59-b99265e6b0d3" },
    @{ Service = "SharePoint Online"; Name = "TermStore.ReadWrite.All"; Type = "Role"; Id = "6c37c71d-f50f-4bff-8fd3-8a41da390140" },
    @{ Service = "SharePoint Online"; Name = "User.ReadWrite.All"; Type = "Role"; Id = "2d4d3d8e-2be3-4bef-9f87-7875a61c29de" },
    
    # Microsoft Information Protection Sync Service
    @{ Service = "MIP Sync Service"; Name = "UnifiedPolicy.Tenant.Read"; Type = "Role"; Id = "ed9faba0-d5a9-474d-b526-6ac2b0e5c690" },
    
    # Office 365 Management APIs
    @{ Service = "Office 365 Mgmt APIs"; Name = "ActivityFeed.Read"; Type = "Role"; Id = "5943d56c-6821-4672-b2fc-6e0db6d499a2" },
    @{ Service = "Office 365 Mgmt APIs"; Name = "ActivityReports.Read"; Type = "Role"; Id = "9e6b5982-b3e4-4daf-b5e2-6a4f6b42f828" },
    @{ Service = "Office 365 Mgmt APIs"; Name = "ServiceHealth.Read"; Type = "Role"; Id = "55896846-df78-47a7-aa94-8d3d4442ca7f" }
)

# Display permissions summary
Write-Host "`nThe following permissions will be assigned (admin consent will be required later):" -ForegroundColor Cyan
$permissions | Format-Table Service, Name, Type -AutoSize

# Confirm before proceeding
$confirmation = Read-Host "`nDo you want to proceed with these permissions? (Y/N)"
if ($confirmation -ne 'Y') {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    exit
}

# Group permissions by service
$services = $permissions | Group-Object Service

# Process each service
foreach ($service in $services) {
    Write-Host "`nProcessing $($service.Name) permissions..." -ForegroundColor Cyan
    
    $spn = switch ($service.Name) {
        "Microsoft Graph" { $graphSpn }
        "SharePoint Online" { $sharepointSpn }
        "MIP Sync Service" { $mipSpn }
        "Office 365 Mgmt APIs" { $officeMgmtSpn }
    }
    
    # Prepare the resource access object
    $resourceAccess = $service.Group | ForEach-Object {
        @{
            id   = $_.Id
            type = $_.Type
        }
    }
    
    # Update the application with permissions for this service
    try {
        $currentAccess = (Get-MgApplication -ApplicationId $app.AppId).RequiredResourceAccess | 
            Where-Object { $_.ResourceAppId -eq $spn.AppId }
        
        if ($currentAccess) {
            # Merge existing permissions with new ones
            $combinedAccess = $currentAccess.ResourceAccess + $resourceAccess
            $uniqueAccess = $combinedAccess | Group-Object id | ForEach-Object { $_.Group[0] }
            
            Update-MgApplication -ApplicationId $app.AppId `
                -RequiredResourceAccess @(
                    @{
                        resourceAppId  = $spn.AppId
                        resourceAccess = $uniqueAccess
                    }
                ) -ErrorAction Stop
        }
        else {
            Update-MgApplication -ApplicationId $app.AppId `
                -RequiredResourceAccess @(
                    @{
                        resourceAppId  = $spn.AppId
                        resourceAccess = $resourceAccess
                    }
                ) -ErrorAction Stop
        }
        
        Write-Host "  Successfully assigned $($service.Group.Count) permissions" -ForegroundColor Green
    }
    catch {
        Write-Host "  Error assigning $($service.Name) permissions: $_" -ForegroundColor Red
        continue
    }
}

# Completion message
Write-Host "`nProcess completed. Summary:" -ForegroundColor Cyan
Write-Host "- Assigned permissions across $($services.Count) services" -ForegroundColor White
Write-Host "- Total permissions assigned: $($permissions.Count)" -ForegroundColor White

Write-Host "`nNext steps:" -ForegroundColor Green
Write-Host "1. An Azure AD administrator must grant admin consent for these permissions" -ForegroundColor Yellow
Write-Host "2. Verify in Azure Portal under:" -ForegroundColor Green
Write-Host "   Azure AD > App Registrations > $($app.DisplayName) > API permissions" -ForegroundColor Green

# Disconnect (optional)
# Disconnect-MgGraph
