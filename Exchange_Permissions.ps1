<#
.SYNOPSIS
Assigns Microsoft Graph permissions to an Azure AD application with interactive Client ID input
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
    Connect-MgGraph -Scopes "Application.ReadWrite.All AppRoleAssignment.ReadWrite.All" -NoWelcome -ErrorAction Stop
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

# Define all permissions
$permissions = @(
    # Application permissions (Role)
    @{ Name = "APIConnectors.ReadWrite.All"; Type = "Role"; Id = "f431331c-49a6-499f-be1c-62af19c34a9f" },
    @{ Name = "Directory.ReadWrite.All"; Type = "Role"; Id = "78c8a3c8-a07e-4b9e-af1b-b5ccab50a175" },
    @{ Name = "Files.ReadWrite.All"; Type = "Role"; Id = "75359482-378d-4052-8f01-80520e7db3cd" },
    @{ Name = "MailboxFolder.ReadWrite.All"; Type = "Role"; Id = "3db474e9-6946-4e69-a614-738289b6a695" },
    @{ Name = "MailboxItem.ImportExport.All"; Type = "Role"; Id = "7aa9d49a-6b49-4dad-a3f9-3e4a2685049e" },
    @{ Name = "MailboxItem.Read.All"; Type = "Role"; Id = "435644c6-a5b1-40bf-8de8-4a2af15b5570" },
    @{ Name = "People.Read.All"; Type = "Role"; Id = "d04bb851-cb7c-4146-97c7-ca3e71baf56c" },
    @{ Name = "RecordsManagement.ReadWrite.All"; Type = "Role"; Id = "44914903-9b0d-40f1-a6b9-a56a41a1e6a5" },
    @{ Name = "Tasks.ReadWrite.All"; Type = "Role"; Id = "44e666d1-d276-445b-a5fc-8815eeb81d55" },
    @{ Name = "Mail.ReadWrite"; Type = "Role"; Id = "024d486e-b451-40bb-833d-3e66d98c5c73" },
    @{ Name = "User.Read.All"; Type = "Role"; Id = "df021288-bdef-4463-88db-98f22de89214" },
    
    # Delegated permissions (Scope)
    @{ Name = "Calendars.ReadWrite"; Type = "Scope"; Id = "1d5bb343-7c8b-4b35-9598-73aaa7208215" },
    @{ Name = "Contacts.ReadWrite"; Type = "Scope"; Id = "afb6c84b-06be-49af-80bb-8f3f77004eab" },
    @{ Name = "Directory.ReadWrite.All"; Type = "Scope"; Id = "863451e7-0667-486c-a5d6-d135439485f0" },
    @{ Name = "Files.ReadWrite.All"; Type = "Scope"; Id = "863451e7-0667-486c-a5d6-d135439485f0" },
    @{ Name = "Files.ReadWrite.AppFolder"; Type = "Scope"; Id = "8019c312-3263-48e6-825e-2b833497195b" },
    @{ Name = "Files.SelectedOperations.Selected"; Type = "Scope"; Id = "3b5f3d61-589b-4a3c-a359-5dd4b5ee5bd5" },
    @{ Name = "Mail.ReadWrite"; Type = "Scope"; Id = "e2a3a72e-5f79-4c64-b1b1-878b674786c9" },
    @{ Name = "Mail.Send"; Type = "Scope"; Id = "e383f46e-2787-4529-855e-0e479a3ffac0" },
    @{ Name = "MailboxSettings.ReadWrite"; Type = "Scope"; Id = "818c620a-27a9-40bd-a6a5-d96f7d610b4b" },
    @{ Name = "People.Read.All"; Type = "Scope"; Id = "8f6a01e7-0391-4ee5-aa22-a3af122cef27" },
    @{ Name = "ShortNotes.ReadWrite"; Type = "Scope"; Id = "328438b7-4c01-4c07-a840-e7a6d1d2d1e7" },
    @{ Name = "Tasks.ReadWrite.All"; Type = "Scope"; Id = "f45671fb-e0fe-4b4b-be20-3d3ce43f1bcb" },
    @{ Name = "User.Read.All"; Type = "Scope"; Id = "a154be20-db9c-4678-8ab7-66f6cc099a59" }
)

# Display permissions summary
Write-Host "`nThe following permissions will be assigned:" -ForegroundColor Cyan
$permissions | Format-Table Name, Type -AutoSize

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

# Update the application with all permissions
try {
    Write-Host "`nAssigning permissions to application..." -ForegroundColor Cyan
    Update-MgApplication -ApplicationId $app.AppId `
        -RequiredResourceAccess @{
            resourceAppId  = "00000003-0000-0000-c000-000000000000"
            resourceAccess = $resourceAccess
        } -ErrorAction Stop
    
    Write-Host "Successfully assigned all permissions." -ForegroundColor Green
}
catch {
    Write-Host "Error assigning permissions: $_" -ForegroundColor Red
    exit
}

# Grant admin consent for application permissions
Write-Host "`nGranting admin consent for application permissions..." -ForegroundColor Cyan
foreach ($perm in $permissions | Where-Object { $_.Type -eq "Role" }) {
    try {
        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $app.Id `
            -PrincipalId $app.Id `
            -ResourceId $graphSpn.Id `
            -AppRoleId $perm.Id -ErrorAction Stop
        
        Write-Host "  Granted: $($perm.Name)" -ForegroundColor Green
    }
    catch {
        Write-Host "  Failed to grant $($perm.Name): $_" -ForegroundColor Yellow
    }
}

# Completion message
Write-Host "`nProcess completed. Summary:" -ForegroundColor Cyan
Write-Host "- Assigned $($permissions.Count) permissions total" -ForegroundColor White
Write-Host "- Granted admin consent for $($permissions.Where{$_.Type -eq 'Role'}.Count) application permissions" -ForegroundColor White

Write-Host "`nVerify in Azure Portal under:" -ForegroundColor Green
Write-Host "Azure AD > App Registrations > $($app.DisplayName) > API permissions" -ForegroundColor Green
Write-Host "`nNote: Delegated permissions will require user consent during login." -ForegroundColor Yellow

# Disconnect (optional)
# Disconnect-MgGraph
