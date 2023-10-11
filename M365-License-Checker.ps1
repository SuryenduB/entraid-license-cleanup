
param(
    [Parameter(Mandatory=$true)]
    [string]$licenseId="c2273bd0-dff7-4215-9ef5-2c7bcfb06425"
  
)
# Note: Details user and application sign-in activity for a tenant (directory). You must have an Azure AD Premium P1 or P2 license to download sign-in logs using the Microsoft Graph API.
# Reference: https://learn.microsoft.com/en-us/graph/api/resources/signin?view=graph-rest-1.0

# Connect (you will need admin permissions for this)
#CSS codes
$header = @"
<style>

    h1 {

        font-family: Arial, Helvetica, sans-serif;
        color: #e68a00;
        font-size: 28px;

    }

    
    h2 {

        font-family: Arial, Helvetica, sans-serif;
        color: #000099;
        font-size: 16px;

    }

    
    
   table {
		font-size: 12px;
		border: 0px; 
		font-family: Arial, Helvetica, sans-serif;
	} 
	
    td {
		padding: 4px;
		margin: 0px;
		border: 0;
	}
	
    th {
        background: #395870;
        background: linear-gradient(#49708f, #293f50);
        color: #fff;
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
	}

    tbody tr:nth-child(even) {
        background: #f0f0f2;
    }
    


    #CreationDate {

        font-family: Arial, Helvetica, sans-serif;
        color: #ff3300;
        font-size: 12px;

    }



    .StopStatus {

        color: #ff0000;
    }
    
  
    .RunningStatus {

        color: #008000;
    }




</style>
"@
Connect-MgGraph -Scopes User.Read.All,Directory.Read.All,AuditLog.Read.All,Reports.Read.All

$header = "<h1>M365 Enterprise Apps License Report: $(Get-Date) </h1>"

# Enterprise - c2273bd0-dff7-4215-9ef5-2c7bcfb06425

$license = $LicenseID # Enterprise License 

# For more information about the filter query check out the following resource:
#   https://learn.microsoft.com/en-us/graph/filter-query-parameter?context=graph%2Fapi%2F1.0&view=graph-rest-1.0
$filter = "assignedLicenses/any(s:s/skuId eq " + $license + ")"

# Licenses


# The -All Flag returns all, so avoid whilst testing out your query
#$users = Get-MgUser -Property "displayName","signInActivity","assignedLicenses","accountEnabled"
#$users

# Unfiltered List
#$users = Get-MgUser -Property "displayName","signInActivity","assignedLicenses","accountEnabled" | Select DisplayName,accountEnabled,signInActivity,assignedLicenses
#$users

# High recommend looking at the graph explorer to help you understand the properties you can query
#   https://developer.microsoft.com/en-us/graph/graph-explorer
#   e.g. https://graph.microsoft.com/beta/users?$select=displayName,signInActivity&$filter=assignedLicenses/any(s:s/skuId eq  '6634e0ce-1a9f-428c-a498-f84ec7b8aa2e')
Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/users?`$filter=`"assignedLicenses/any(s:s/skuId eq  $license  )&`$expand=`"memberOf(`$select=id,displayName)" -Headers @{"ConsistencyLevel"="eventual"}
# Filtered List
$licensedusers = Get-MgBetaUser -Property "displayName","signInActivity","assignedLicenses","accountEnabled","userPrincipalName","officeLocation","memberOf" -ExpandProperty "memberOf(`$select=id,displayName)"   -Filter $filter -Debug
$Period = 'D7'
$OfficeGroup = Get-MgBetaGroup -ConsistencyLevel eventual -Count groupCount -Search '"DisplayName:lic_office"'
# Export the list to CSV File

Get-MgReportM365AppUserDetail -Period $Period -OutFile Office365UsersReport.csv
$Office365UsersReport = Import-Csv Office365UsersReport.csv
$result = New-Object System.Collections.ArrayList

foreach($user in $licensedusers)
{
    
    
    #Lookup values in $Office365UsersReport and get the Last Activity Date
    $property = (New-Object PSObject -Property @{
        DisplayName = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        AccountEnabled = $user.AccountEnabled
        OfficeLocation = $user.OfficeLocation
        # Look if the user is a member of the Office Group by comparing Id Values
        
        
        LastActivityDate = $Office365UsersReport | Where-Object {$_.'User Principal Name' -eq $user.UserPrincipalName} | Select-Object -ExpandProperty 'Last Activity Date'
    } 
    )
    $result.Add($property) | Out-Null

    
}
$result | ConvertTo-Html  -Property DisplayName,UserPrincipalName,AccountEnabled,OfficeLocation,OfficeGroup,LastActivityDate -Fragment -PreContent $header  | Out-File .\Office365UsersReport.html
.\Office365UsersReport.html
