[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)]
	[string]
	$TenantName = "sierracedarinc",

	[Parameter(Mandatory = $false)]
	[string]
	$CredentialId = "Tenant Automation"
)

# Static Attributes
$POLICY_TYPE_ACTION = "GuestUserOrphan"
$SB_URI = Get-AutomationVariable -Name "ServiceBusUri"
$SB_ACCESS_POLICY_NAME = Get-AutomationVariable -Name "ServiceBusAccessPolicyName"
$SB_ACCESS_POLICY_KEY = Get-AutomationVariable -Name "ServiceBusAccessPolicyKey"
$SB_TOKEN_DURATION_SECONDS = Get-AutomationVariable -Name "ServiceBusTokenDuration"
$RB_CONNECTION_DELAY_SECONDS = Get-AutomationVariable -Name "RunbookConnectionDelay"
$RB_CONNECTION_MAX_RETRY = Get-AutomationVariable -Name "RunbookConnectionMaxRetry"

# Template Properties (Used by all runbooks)
$TenantFqdn = "$($TenantName).onmicrosoft.com"
$SbHeader = $null
$Credential = Get-AutomationPSCredential -Name $CredentialId

###############################
# TEMPLATE FUNCTIONS

function Send-ServiceBusQueue {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[System.Collections.IDictionary]
		$Message
	)

	process {

		function Init-ServiceBusQueue {

			# Calculate token expiry Now + 5 mins
			$expires = [Int64](([DateTime]::UtcNow)-(Get-Date "1/1/1970")).TotalSeconds + $SB_TOKEN_DURATION_SECONDS

			# Create token
			$sigStr = [System.Web.HttpUtility]::UrlEncode($SB_URI) + "`n" + [String]$expires
			$key = [Text.Encoding]::ASCII.GetBytes($SB_ACCESS_POLICY_KEY)
			$hmac = New-Object System.Security.Cryptography.HMACSHA256(,$key)

			$signature = $hmac.ComputeHash([Text.Encoding]::ASCII.GetBytes($sigStr))
			$signature = [Convert]::ToBase64String($signature)
			$sasToken = [String]::Format("SharedAccessSignature sig={0}&se={1}&skn={2}&sr={3}", [System.Web.HttpUtility]::UrlEncode($signature), $expires, $SB_ACCESS_POLICY_NAME, [System.Web.HttpUtility]::UrlEncode($SB_URI))

			# Construct HTTP request
			Set-Variable -Scope 1 -Name "SbHeader" -Value @{
				"Authorization" = $sasToken;
				"BrokerProperties" = "{`"Label`": `"$($POLICY_TYPE_ACTION)`"}"
			}
		}

		function Send-ServiceBusQueuePayload {

			#Convert message to JSON payload
			$payload = $Message | ConvertTo-Json

			# Send the message to service bus
			$response = Invoke-RestMethod "$($SB_URI)/messages" -Method Post -Headers $SbHeader -Body $payload -ContentType "application/json" -ErrorVariable $thisError -ErrorAction SilentlyContinue

			# token may have expired
			if($thisError) {
				Init-ServiceBusQueue
				$response = Invoke-RestMethod "$($SB_URI)/messages" -Method Post -Headers $SbHeader -Body $payload -ContentType "application/json" -ErrorVariable $thisError -ErrorAction SilentlyContinue

				# if still an error
				if($thisError) {
					Write-Error "There was an error sending the payload to the service bus queue."
					exit
				}
			}

		}

		if($SbHeader -eq $null) {
			Init-ServiceBusQueue
		}

		Send-ServiceBusQueuePayload

	}
}

function Disconnect-Office365 {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $false)]
		[bool]
		$SPOService = $false,

		[Parameter(Mandatory = $false)]
		[bool]
		$MsolService = $false,

		[Parameter(Mandatory = $false)]
		[bool]
		$PnPOnline = $false,

		[Parameter(Mandatory = $false)]
		[bool]
		$ExOnline = $false,

		[Parameter(Mandatory = $false)]
		[bool]
		$SecurityComplianceCenter = $false
	)

	Disconnect-AzureAD

	if($SPOService) {
		Disconnect-AzureAD
	}

	if($SPOService) {
		Disconnect-SPOService
	}

	if($MsolService) {
		Disconnect-MsolService
	}

	if($PnpOnline) {
		Disconnect-PnPOnline
	}

	if($ExOnline -or $SecurityComplianceCenter) {
		Get-PSSession | Remove-PSSession
	}

}

function Connect-Office365 {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $false)]
		[bool]
		$SPOService = $false,

		[Parameter(Mandatory = $false)]
		[bool]
		$MsolService = $false,

		[Parameter(Mandatory = $false)]
		[bool]
		$PnPOnline = $false,

		[Parameter(Mandatory = $false)]
		[String]
		$PnPOnlineSiteUri = [String]::Empty,

		[Parameter(Mandatory = $false)]
		[bool]
		$ExOnline = $false,

		[Parameter(Mandatory = $false)]
		[bool]
		$SecurityComplianceCenter = $false,

		[Parameter(Mandatory = $false)]
		[bool]
		$ResetAllSessions = $false
	)

	process {
		function Connect-EXOPSSession {
			$sessionName = [String]::Format("$(Get-AutomationVariable -Name "PSSessionExchangeOnlineName")", $TenantName)
			$sessionUri = "https://outlook.office365.com/powershell-liveid/"

			# Remove all previous sessions
			if($ResetAllSessions) {
				Get-PSSession -Name $sessionName | Remove-PSSession
			}

			# Locate an existing session
			$session = Get-PSSession | Where-Object {$_.Name -eq $sessionName}

			# Create the session if it does not exist
			if($session -eq $null) {

				"Session does not exist. Creating a new session ..."
				$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $sessionUri -Credential $Credential -Authentication Basic -AllowRedirection -Name $sessionName

				"Importing session to console ..."
				Import-Module (Import-PSSession -Session $session -AllowClobber -DisableNameChecking)

			} else {
				"Session exists. Re-using the '$($session.Name)' session ..."
			}
		}

		function Connect-IPPSSession {
			$sessionName = [String]::Format("$(Get-AutomationVariable -Name "PSSessionSecurityComplianceCenterName")", $TenantName)
			$sessionUri = "https://ps.compliance.protection.outlook.com/powershell-liveid"

			# Remove all previous sessions
			if($ResetAllSessions) {
				Get-PSSession -Name $sessionName | Remove-PSSession
			}

			# Locate an existing session
			$session = Get-PSSession | Where-Object {$_.Name -eq $sessionName}

			# Create the session if it does not exist
			if($session -eq $null) {

				"Session does not exist. Creating a new session ..."
				$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $sessionUri -Credential $Credential -Authentication Basic -AllowRedirection -Name $sessionName -SessionOption (New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck)

				"Importing session to console ..."
				Import-Module (Import-PSSession -Session $session -AllowClobber -DisableNameChecking)

				if($session -eq $null) {
					Write-Error "Could not create PowerShell session."
					exit;
				}

			} else {
				"Session exists. Re-using the '$($session.Name)' session ..."
			}
		}

		# Connect to Powershell Modules
		if($Credential -eq $null) {
			Write-Error "Credential object was not loaded properly and is null"
			exit
		}

		"Connecting to Azure AD ..."
		Connect-AzureAD -Credential $Credential

		if($SPOService) {
			$retry = 0
			
			do {
				try {
					Write-Output "Connecting to SPO Service ..."
					Connect-SPOService -Url https://$TenantName-admin.sharepoint.com -Credential $Credential

					# if successful, then this will be reached.  Set to max
					# retries to exit loop.
					$retry = $RB_CONNECTION_MAX_RETRY

				} catch {

					#if($_ -match "The remote server returned an error: (403) Forbidden") {
					if($_ -match  "A command that prompts the user failed because the host program or the command type does not") {
						$retry++
						$Credential = Get-AutomationPSCredential -Name $CredentialId

						Start-Sleep -Seconds $RB_CONNECTION_DELAY_SECONDS

					} else {
						Write-Error $_
						exit
					}

				}
			} while($retry -lt $RB_CONNECTION_MAX_RETRY)
		}

		if($PnPOnline) {
			$retry = 0

			$site = [String]::Empty

			# if no string provided, use central admin site
			if($PnPOnlineSiteUri -eq [String]::Empty) {
				$site = "https://$($TenantName)-admin.sharepoint.com"
			} else {
				$site = $PnPOnlineSiteUri
			}

			do {
				try {
					Write-Output "Connecting to PnP Online ..."
					Connect-PnPOnline -Url $site -Credential $Credential

					# if successful, then this will be reached.  Set to max
					# retries to exit loop.
					$retry = $RB_CONNECTION_MAX_RETRY

				} catch {
					if($_ -match "The remote server returned an error: (403) Forbidden") {
						$retry++
						$Credential = Get-AutomationPSCredential -Name $CredentialId

						Start-Sleep -Seconds $RB_CONNECTION_DELAY_SECONDS
					} else {
						Write-Error $_
						exit
					}

				}
			} while($retry -lt $RB_CONNECTION_MAX_RETRY)

		}


		if($MsolService) {
			"Connecting to Msol Service ..."
			Connect-MsolService -Credential $Credential
		}

		if($ExOnline) {
			"Connecting to Exchange Online ..."
			Connect-EXOPSSession
		}

		if($SecurityComplianceCenter) {
			"Connecting to Office 365 Security and Compliance Center ..."
			Connect-IPPSSession
		}


		# Verify tenant
		$domains = Get-AzureADDomain | Where-Object { $_.Name -match $TenantFqdn}

		if($domains.Count -eq 0) {
			Write-Error "The service account is not associated with the $($TenantFqdn) tenant."
			exit
		}
	}
}


###############################
# START SCAN

# Pre-create messages in a Hashtable with a referenceable key
# Any guest user that is not part of an existing ACL will be removed

Connect-Office365 -SPOService $true -PnPOnline $true

Write-Output "Getting guest users from Azure AD ..."

$guestUsers = @{}
Get-AzureADUser -All $true -Filter "userType eq 'Guest'" | ForEach-Object {

    # Messaging object
    $message = @{
        objectId = $_.UserPrincipalName;
        objectType = "GuestUser";
        displayName = $_.DisplayName;
		tenant = $tenantFqdn;
        policyActionType = $POLICY_TYPE_ACTION;
    }
    $guestUsers.Add($_.UserPrincipalName, $message)
}

# -- GET ALL UNIFIED GROUPS IN TENANT
## Placeholder.  Currently no guest users can be added and granted functional
## access to a Unified Group.  To be added later.

#Get-AzureADMSGroup -All $true | Where-Object {$_.GroupTypes -eq "Unified"} | ForEach-Object {
#	Get-AzureADGroupMember -ObjectId $_.Id | Where-Object {$_.UserType -eq "Guest"} | ForEach-Object {
#		$_
#	}
#}

# -- GET ALL SHAREPOINT SITES IN TENANT

Write-Output "Scanning classic SharePoint site collections ..."

#Get all templates
Get-SPOWebTemplate | Foreach-Object {
    "Scanning for site collections with the $($_.Name) template ..."
    Get-SPOSite -Template $_.Name | Where-Object { $_.SharingCapability -ne "Disabled" } | ForEach-Object {

        Write-Output "Scanning site collection $($_.Url) ..."

		# Any user that has visited a SharePoint site will be recorded in the user information list (retrieved by Get-SPOUser)
		# To filter out guest users that still have access, you need to:
		# 1: Query for all guest users (LoginName match #ext#)
		# 2: SharePoint creates a sharing group to grant an end user acces.  Check to see if the user is part of a group. (GroupCount > 0)

        Get-SPOUser -Site $_.Url -Limit All | Where-Object { $_.LoginName -match "#ext#" -and $_.Groups.Count -gt 0 } | ForEach-Object {

            Write-Output "Guest user $($_.LoginName) is active"
            $guestUsers.Remove($_.LoginName)
        }
    }
}

# -- GET ALL ONEDRIVE FOR BUSINESS SITES IN TENANT

Write-Output "Scanning OneDrive for Business site collections ..."

# Get all OneDrive URLs
Get-SPOSite -Filter { "Url -like '*-my.sharepoint.com*'"} -IncludePersonalSite $true -Limit All | ForEach-Object {

	Write-Output "Scanning site collection $($_.Url) ..."

    Get-SPOUser -Site $_.Url -Limit All | Where-Object { $_.LoginName -match "#ext#" -and $_.Groups.Count -gt 0 } | ForEach-Object {
        Write-Progress "Guest user $($_.LoginName) is active"
        $guestUsers.Remove($_.LoginName)
    }
}


# -- SEND TO SERVICE BUS QUEUE

# Iterate and send the message
foreach($message in $guestUsers.Values) {

	Send-ServiceBusQueue -Message $message

}

###############################