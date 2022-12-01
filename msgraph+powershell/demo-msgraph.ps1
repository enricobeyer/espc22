#region Install the Microsoft Graph PowerShell SDK

# PowerShell SDK for Microsoft Graph GitHub repository
start 'https://github.com/microsoftgraph/msgraph-sdk-powershell'

# Microsoft Graph PowerShell module
start 'https://www.powershellgallery.com/packages/Microsoft.Graph'

# Graph Explorer
start 'https://aka.ms/ge'

Find-Module Microsoft.Graph

# Installing the main module of the SDK will install all 40 sub-modules. (~850MB)
# Consider only installing the necessary modules, including Microsoft.Graph.Authentication.

# You can install the SDK in PowerShell 7 or Windows PowerShell
Install-Module Microsoft.Graph -Scope CurrentUser -Verbose 

# Verify the installation
Get-Module Microsoft.Graph -ListAvailable
Get-InstalledModule Microsoft.Graph

# Updating the SDK
Update-Module Microsoft.Graph

# Uninstalling the SDK

# Uninstall the main module
Uninstall-Module Microsoft.Graph
# Remove all of the dependency modules
Get-InstalledModule Microsoft.Graph.* | ForEach-Object { if ($_.Name -ne "Microsoft.Graph.Authentication") { Uninstall-Module $_.Name } }
Uninstall-Module Microsoft.Graph.Authentication

#endregion

#region API version

# By default, the SDK uses the Microsoft Graph REST API v1.0.
# You can change this by using the Select-MgProfile command.
Get-MgProfile
# Select-MgProfile -Name beta 

#endregion

#region Finding commands

Get-Module microsoft.graph.* -ListAvailable
Get-Module microsoft.graph.* -ListAvailable | Measure-Object

Get-Command -Module Microsoft.Graph.Authentication |
Sort-Object noun |
Format-Table -GroupBy noun

Get-Command -Module Microsoft.Graph.Users |
Sort-Object noun |
Format-Table -GroupBy noun

#endregion

#region Authentication

# The PowerShell SDK supports two types of authentication: delegated access, and app-only access.
# First, we will use delegated access to login as a user,
# grant consent to the SDK to act on our behalf, and call the Microsoft Graph.

# Determine required permission scopes

# Each API in the Microsoft Graph is protected by one or more permission scopes.
# The user logging in must consent to one of the required scopes for the APIs you plan to use.

Find-MgGraphCommand -Command Get-MgUser
Find-MgGraphCommand -Command Get-MgUser -ApiVersion v1.0 | Select-Object -ExpandProperty permissions -First 1
Find-MgGraphCommand -Command Get-MgUserJoinedTeam -ApiVersion v1.0 | Select-Object -ExpandProperty permissions -First 1
# Use a URI to get all related cmdlets
Find-MgGraphCommand -Uri '/users/{user-id}'

# What permissions are applicable to a certain domain, for example, user, application, directory? 
Find-MgGraphPermission user
Find-MgGraphPermission User.Read
Find-MgGraphPermission User.Read -ExactMatch -PermissionType Delegated
Find-MgGraphPermission User.Read -ExactMatch -PermissionType Delegated | Format-List 
Find-MgGraphPermission User.Read.All -ExactMatch -PermissionType Delegated | Format-List

# Launch detailed permissions documentation
Get-Help Find-MgGraphPermission -Online
# Launch Microsoft Graph Permission Explorer
start https://graphpermissions.merill.net

# Sign in 

# Use the Connect-MgGraph command to sign in with the required scopes.
# You'll need to sign in with an admin account to consent to the required scopes.

Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All"

# The command prompts you to go to a web page to sign in using a device code.
# Once you've done that, the command indicates success with a Welcome To Microsoft Graph! message.
# You only need to do this once per session.

# You can add additional permissions by repeating the Connect-MgGraph command with the new permission scopes.

Get-MgContext
(Get-MgContext).Scopes

# List users in your Microsoft 365 organization
Get-MgUser

# Get the specific user
Get-MgUser -Filter "displayName eq 'Megan Bowen'" -OutVariable user

Get-MgUserJoinedTeam -UserId $user.Id

# Select one of the user's joined Teams and use its DisplayName to filter the list
$team = Get-MgUserJoinedTeam -UserId $user.Id -Filter "displayName eq 'mo3ak'"

# List all channels, then filter the list to get the specific channel you want
Get-MgTeamChannel -TeamId $team.Id
$channel = Get-MgTeamChannel -TeamId $team.Id -Filter "displayName eq 'General'"

# post a message to the channel
New-MgTeamChannelMessage -TeamId $team.Id -ChannelId $channel.Id -Body @{ Content="Hello World" }
New-MgTeamChannelMessage -TeamId $team.Id -ChannelId $channel.Id -Body @{ Content="Hello, ESPC22!" } -Importance "urgent"
start https://teams.microsoft.com
#endregion

#region Configuring app-only access for a simple script to list users and groups in your Microsoft 365 tenant

# Certificate

# You will need an X.509 certificate installed in your user's trusted store on the machine where you will run the script.
# You'll also need the certificate's public key exported in .cer, .pem, or .crt format.
# You'll need the value of the certificate subject or its thumbprint.

# Create and export your public certificate without a private key

# Use this method to authenticate from an application running from your machine
# For example, authenticate from PowerShell 7 or Windows PowerShell

$cert = New-SelfSignedCertificate -Subject "CN=PowerShellScriptCert" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256

$certPath = "C:\demo-msgraph\PowerShellScriptCert.cer"

Export-Certificate -Cert $cert -FilePath $certPath

# Register the application

# First, you're using the PowerShell SDK with delegated access, logging in as an administrator, and creating the app registration.
# Then, using that app registration, you're able to use the PowerShell SDK with app-only access, allowing for unattended scripts.

code ./RegisterAppOnly.ps1

.\RegisterAppOnly.ps1 -AppName "Graph PowerShell Script" -CertPath $certPath

# Authenticate
Connect-MgGraph -ClientId "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXX" -TenantId "XXXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXX" -CertificateName "CN=PowerShellScriptCert"

Get-MgContext

# Let's test the script

code .\GraphAppOnly.ps1

.\GraphAppOnly.ps1

#endregion


# Graph PowerShell Samples Community
start https://aka.ms/graphsamples

# Microsoft Entra admin center
start https://entra.microsoft.com/#home
# Graph X-Ray
start https://graphxray.merill.net/

# Graph PowerShell Conversion Analyzer
start https://graphpowershell.merill.net/

#region Various examples

Invoke-MgGraphRequest -Method GET https://graph.microsoft.com/v1.0/me
Invoke-MgGraphRequest -Method GET https://graph.microsoft.com/v1.0/users
Invoke-MgGraphRequest -Method GET https://graph.microsoft.com/v1.0/users -OutVariable result
$result.value | gm
Invoke-MgGraphRequest -Method GET https://graph.microsoft.com/v1.0/users -OutputType PSObject -outvariable result2 
$result2.value | gm
$result2.value | ft *name
$result.value | select -first 1
$result2.value | select -first 1

# Tenant Information

## Organization Contact Details
Get-MgOrganization | Select-Object DisplayName, City, State, Country, PostalCode, BusinessPhones

## Organization Assigned Plans
Get-MgOrganization | Select-Object -expand AssignedPlans

## List application registrations in the tenant
Get-MgApplication | Select-Object DisplayName, Appid, SignInAudience

## List service principals in the tenant
Get-MgServicePrincipal | Select-Object id, AppDisplayName | Where-Object { $_.AppDisplayName -like "*graph*" }

# Microsoft Graph Users and Groups Snippets

# List of Users
Get-MgUser -top 999 | Select-Object id, displayName, OfficeLocation, BusinessPhones

# List of users with no Office Location
Get-MgUser | Select-Object id, displayName, OfficeLocation, BusinessPhones | Where-Object {!$_.OfficeLocation }

# Update the location of the user
Update-MgUser -UserId $UserId -OfficeLocation $NewLocation

# Get all Groups
Get-MgGroup -top 999 | Select-Object id, DisplayName, GroupTypes

# Get all unified (Microsoft 365 Groups) Groups
Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')"

<# Search-MgGroup by Justin Grote
#requires -Module Microsoft.Graph.Groups
function Search-MGGroup {
  param(
    $Name,
    #Only returns results that are Microsoft Teams
    [Switch]$Team
  )
  if ($Team) {
    $filter = "resourceProvisioningOptions/Any(x:x eq 'Team')"
  }
  if ($Name) {
    $search = "displayname:$Name"
  }
  Get-MgGroup -Filter $filter -Search $search -ConsistencyLevel eventual
}
#>

. .\Search-MgGroup.ps1
Search-MGGroup -Name "mo3ak"

# Get-Details of a single Group
Get-MgGroup -GroupId $groupId | Format-List | more

# Get Owners of a Group
Get-MgGroupOwner -GroupId $GroupId 

# Translate Directory Objects to Users 
Get-MgGroupOwner -GroupId $GroupId | ForEach-Object { @{ UserId=$_.Id}} | get-MgUser | Select-Object id, DisplayName, Mail

# Could do the same for Group Members
Get-MgGroupMember -GroupId $GroupId 

# Get your mail
Get-MgUserMessage -UserId $UserId -Filter "contains(subject,'Marketing')" | Select-Object sentDateTime, subject

# New Group
$group = New-MgGroup -DisplayName "PowerFam" -MailEnabled:$false -mailNickName "powerfam" -SecurityEnabled

# Add member to Group  
New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $UserId

# View new member to Group
Get-MgGroupMember -GroupId $group.Id  | ForEach-Object { @{ UserId=$_.Id}} | Get-Mguser | Select-Object id, DisplayName, Mail

#Remove Group
Remove-MgGroup -GroupId $Group.Id

# Create a new User
New-MgUser -displayName "Bob Brown" -AccountEnabled -PasswordProfile @{"Password"="{password}"} `
         -MailNickname "Bob.Brown" -UserPrincipalName "bob.brown@{tenantdomain}"
#endregion