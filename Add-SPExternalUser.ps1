#Script#Add-SPExternalUser#
Function Add-SPExternalUser {

<#
.SYNOPSIS
   Uses CSOM to send an invite to the named user
.DESCRIPTION
   This function will check that thes SharePoint Client Side Object Model dll's are installed and if so will load them into memory
   It will then attempt to send a Sharing invite to the specified user
.EXAMPLE
   Add-SPExternalUser -UserEmail kiasuit@hotmail.co.uk -Role Edit -SiteUrl https://kilasuit.sharepoint.com/sites/csom -Message 'Project Access' -Credential (Get-Credential ryan@kilasuit.onmicrosoft.com)
.EXAMPLE
   Import-CSV C:\temp\externaluserrequests.csv | Foreach-Object { Add-SPExternalUser -UserEmail $_.Email -Role $_.Role -SiteUrl $_.SiteUrl -Message $_.Message -Credential $MySharePointCredential }
#>


[cmdletbinding()]
param(
    
    [Parameter(Mandatory = $true)]
    [String]$UserEmail,

    [Parameter(Mandatory = $true)]
    [Microsoft.SharePoint.Client.Sharing.Role]$Role,

    [Parameter(Mandatory = $true)]
    [String]$SiteURL,

    [Parameter(Mandatory = $true)]
    [String]$Message,

    [Parameter(Mandatory = $true)]
    [PSCredential]$Credential

     )

    $Clientdll = “C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll”

    $ClientRuntimeDll = “C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll”

    If ( (Test-Path $Clientdll) -and (Test-Path $ClientRuntimeDll) ) 
         {  Add-Type –Path $Clientdll
            Add-Type –Path $ClientRuntimeDll
        
            Write-Verbose -Message "Attempting to give $UserEmail the following $Role rights on $SiteURL"
        
            If ($ClientContext.site.url -ne $SiteURL) 
                {   $Global:ClientContext = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
                    $ClientContext.Credentials  = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.UserName,$Credential.Password)
                }

            $userList = New-Object "System.Collections.Generic.List``1[Microsoft.SharePoint.Client.Sharing.UserRoleAssignment]"

            $userRoleAssignment = New-Object Microsoft.SharePoint.Client.Sharing.UserRoleAssignment
            $userRoleAssignment.UserId = $UserEmail
            $userRoleAssignment.Role = $role
            $userList.Add($userRoleAssignment)

            [Microsoft.SharePoint.Client.Sharing.WebSharingManager]::UpdateWebSharingInformation($ClientContext, $ClientContext.Web, $userList, $true, $message, $true, $true)
            $ClientContext.ExecuteQuery()
        }
        else {
                Write-Warning -Message "SharePoint Client Dll's not found please install them from http://www.microsoft.com/en-us/download/details.aspx?id=42038"
             }
}

