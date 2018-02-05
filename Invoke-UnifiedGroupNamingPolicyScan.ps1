[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)]
	[String]
	$TenantName = "devflin",

	[Parameter(Mandatory = $false)]
	[String]
	$CredentialId = "Tenant Automation",

	[Parameter(Mandatory = $false)]
	[Object]
	$DomainCompanyJson = "{`"itserveco.com`": `"ITS`", 
		`"sierra-cedar.com`": `"SCI`", 
		`"sierrasystems.com`": `"SSG`", 
		`"sierracedarinc.onmicrosoft.com`": `"ITS`",
		`"devflin.onmicrosoft.com`": `"DEV`", 
		`"itscorpsysdev01.onmicrosoft.com`": `"DEV`" }",

	[Parameter(Mandatory = $false)]
	[String]
	$GroupDisplayNameTemplate = "{{CompanyId}} {{DisplayName}}",

	[Parameter(Mandatory = $false)]
	[String]
	$GroupEmailTemplate = "ogrp-{{CompanyId}}-{{MailNickname}}@{{CompanyMailDomain}}",

	[Parameter(Mandatory = $false)]
	[int]
	$MessagingLimit = 10,

	[Parameter(Mandatory = $false)]
	[bool]
	$ReadOnly = $true

)

# Static Properties
$POLICY_TYPE_ACTION = "UnifiedGroupNamingPolicy"
$SB_URI = Get-AutomationVariable -Name "ServiceBusUri"
$SB_ACCESS_POLICY_NAME = Get-AutomationVariable -Name "ServiceBusAccessPolicyName"
$SB_ACCESS_POLICY_KEY = Get-AutomationVariable -Name "ServiceBusAccessPolicyKey"
$SB_TOKEN_DURATION_SECONDS = Get-AutomationVariable -Name "ServiceBusTokenDuration"
$RB_CONNECTION_DELAY_SECONDS = Get-AutomationVariable -Name "RunbookConnectionDelay"
$RB_CONNECTION_MAX_RETRY = Get-AutomationVariable -Name "RunbookConnectionMaxRetry"
$DEBUG_MODE = $false

# Template Properties (Used by all runbooks)
$TenantFqdn = "$($TenantName).onmicrosoft.com"
$Credential = Get-AutomationPSCredential -Name $CredentialId
$CorrelationId = [Guid]::NewGuid()
$SessionId = [Guid]::NewGuid()
$MessageQueue = New-Object System.Collections.Queue
$MessageKeys = {@()}.Invoke()
$Statistics = @{}

# Function Properties
$MailDomainRegEx = "@(.+)$"


###############################
# TEMPLATE FUNCTIONS

function Test-TenantDomainVanityCheck {
	[CmdletBinding()]
	[OutputType([Boolean])]
	param()
	# Verify tenant
	$domains = Get-AzureADDomain | Where-Object { $_.Name -match $TenantFqdn}

	if($domains.Count -eq 0) {
		return $false
	}

	return $true
}

function Send-ServiceBusQueue {
	[CmdletBinding()]
	param(
	)

	process {
		$sbHeader = $null

		function Initialize-ServiceBusQueue {

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

			Set-Variable -Scope 1 -Name "sbHeader" -Value @{
				"Authorization" = $sasToken;
			}


		}

		function Send-ServiceBusQueuePayload {
			$payload = [String]::Empty

			#Convert message to JSON payload
			if($MessageQueue.Count -gt 1) {

				$payload = $MessageQueue | ConvertTo-Json
			} elseif($MessageQueue.Count -eq 1) {

				$payload = $MessageQueue.Dequeue() | ConvertTo-Json
				$payload = [String]::Format("[{0}]", $payload)
			} else {

				Write-Output "No messages have been queued."
				return
			}

			try {
				# Send the message to service bus
				$response = Invoke-RestMethod "$($SB_URI)/messages" `
				-Method Post `
				-Headers $sbHeader `
				-Body $payload `
				-ContentType "application/vnd.microsoft.servicebus.json" `
				-ErrorVariable $thisError `
				-ErrorAction SilentlyContinue

			} catch {
				Write-Error "There was an error sending the payload to the service bus queue."
				Write-Error $_
				$streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
				$errResp = $streamReader.ReadToEnd()
				$streamReader.Close()
				Write-Error $errResp
				exit
			}

		}

		if($sbHeader -eq $null) {
			Initialize-ServiceBusQueue
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
			$retry = 0

			# Remove all previous sessions
			if($ResetAllSessions) {
				Get-PSSession -Name $sessionName | Remove-PSSession
			}

			# Locate an existing session
			$session = Get-PSSession | Where-Object {$_.Name -eq $sessionName}

			# Create the session if it does not exist
			if($session -eq $null) {

				do {
					try {
						"Session does not exist. Creating a new session ..."
						$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $sessionUri -Credential $Credential -Authentication Basic -AllowRedirection -Name $sessionName -ErrorAction Stop

						"Importing session to console ..."
						Import-Module (Import-PSSession -Session $session -AllowClobber -DisableNameChecking)

						# if successful, then this will be reached.  Set to max
						# retries to exit loop.
						$retry = $RB_CONNECTION_MAX_RETRY

					} catch {
						# Wasn't able to get the credential object. Need to try again.
						Write-Output "Error connecting.  Retrying ..."
						$retry++
						$Credential = Get-AutomationPSCredential -Name $CredentialId

						if($retry -ge $RB_CONNECTION_MAX_RETRY) {
							Write-Error "Max retries exceeded.  Terminating script."
							exit
						}

						Start-Sleep -Seconds $RB_CONNECTION_DELAY_SECONDS			
					}

				} while($retry -lt $RB_CONNECTION_MAX_RETRY)

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
		Write-Output "Connecting to Azure AD ..."
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

					Write-Output "Error connecting.  Retrying ..."

					$retry++
					$Credential = Get-AutomationPSCredential -Name $CredentialId

					if($retry -ge $RB_CONNECTION_MAX_RETRY) {
						Write-Error "Max retries exceeded.  Terminating script."
						exit
					}

					Start-Sleep -Seconds $RB_CONNECTION_DELAY_SECONDS

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

					Write-Output "Error connecting.  Retrying ..."

					$retry++
					$Credential = Get-AutomationPSCredential -Name $CredentialId

					if($retry -ge $RB_CONNECTION_MAX_RETRY) {
						Write-Error "Max retries exceeded.  Terminating script."
						exit
					}

					Start-Sleep -Seconds $RB_CONNECTION_DELAY_SECONDS


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
	}
}

function Add-Statistics {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[String]
		$Key
	)

	process {
		# Add to Statistics
		if($Statistics.ContainsKey($Key)) {

			# Increment counter
			$count = $Statistics[$Key]
			$count++

			# Save to variable
			$Statistics[$Key] = $count

		} else {

			# add one occurance of domain
			$Statistics.Add($Key, 1)
		}
	}
}

###############################
# LOCAL FUNCTIONS


function Set-TemplateValues {
	[CmdletBinding()]
	[OutputType([String])]
	param(
		[Parameter(Mandatory = $true)]
		[String]
		$Template,

		[Parameter(Mandatory = $false)]
		[String]
		$CompanyId,

		[Parameter(Mandatory = $false)]
		[String]
		$DisplayName,

		[Parameter(Mandatory = $false)]
		[String]
		$MailNickname,

		[Parameter(Mandatory = $false)]
		[String]
		$CompanyMailDomain
	)

	process {
		$str = $Template
		$str = $str -replace "{{CompanyId}}", $CompanyId
		$str = $str -replace "{{DisplayName}}", $DisplayName
		$str = $str -replace "{{MailNickname}}", $MailNickname
		$str = $str -replace "{{CompanyMailDomain}}", $CompanyMailDomain

		$str
		
	}
}

function Get-UnifiedGroupOwnerDomainCount {
	[CmdletBinding()]
	[OutputType([Hashtable])]
	param(
		[Parameter(Mandatory = $true)]
		[String]
		$ObjectId
	)

	process {
		$owners = @{}

		# Count domain usage for the owner
		Get-AzureADGroupOwner -ObjectId $ObjectId | ForEach-Object {

			# Get the mail domain of the owner
			$ownerMailDomain = [Regex]::Match($_.Mail, $MailDomainRegEx).Groups[1].Value
			$ownerMailDomain = $ownerMailDomain.Replace("@", [String]::Empty)

			# Add or update the domain count
			if($owners.ContainsKey($ownerMailDomain)) {

				# Increment counter
				$count = $owners[$ownerMailDomain]
				$count++

				# Save to variable
				$owners[$ownerMailDomain] = $count

			} else {

				# add one occurance of domain
				$owners.Add($ownerMailDomain, 1)
			}
		}

		return $owners

	}
}

function Get-UnifiedGroupMajorityOrg {
	[CmdletBinding()]
	[OutputType([Hashtable])]
	param(
		[Parameter(Mandatory = $true)]
		[String]
		$ObjectId
	)

	process {
		$ht = @{}

		$owners = Get-UnifiedGroupOwnerDomainCount -ObjectId $ObjectId

		$companyId = [String]::Empty
		$companyMailDomain = [String]::Empty


		# Get the most recurring domain by count
		$companyDomain = $owners.GetEnumerator() | Sort-Object -Descending Value | Select-Object -First 1

		# Get the domain references in an object
		$domainRef = $DomainCompanyJson | ConvertFrom-Json

		# Process only if the company domain has been
		# determined.  Otherwise, there may not  be any
		# owners or errored logic.
		if($companyDomain -ne $null) {

			# Set template variables for this iteration
			# of the unified group.  The value will be different
			# for each loop.

			$ht.Add("CompanyMailDomain", $companyDomain.Name)
			$ht.Add("CompanyId", $domainRef.PSObject.Properties[$companyDomain.Name].Value)

			return $ht
		}

		return $null
	}
}


function Invoke-UnifiedGroupNamingPolicyAction {
	[CmdletBinding()]
	param()

	foreach($msg in $MessageQueue) {

		try {
			$body = $msg.Body | ConvertFrom-Json
			$setDisplayName = $false
			$setEmail = $false

			# Set display name
			if($body.DisplayName -ne $body.NewDisplayName) {

				Write-Output "Set-AzureADMSGroup -Id $($body.ObjectId) -DisplayName ""$($body.NewDisplayName)"""

				if($ReadOnly -eq $false) {
					Set-AzureADMSGroup -Id $($body.ObjectId) -DisplayName $body.NewDisplayName
					$setDisplayName = $true
				}
			}

			# Set the email
			if($body.Email -ne $body.NewEmail) {

				Write-Output "Set-UnifiedGroup -Identity $($body.ObjectId) -PrimarySmtpAddress ""$($body.NewEmail)"""

				if($ReadOnly -eq $false) {
					Set-UnifiedGroup -Identity $($body.ObjectId) -PrimarySmtpAddress $body.NewEmail
					$setEmail = $true
				}
			
			}
		} catch {
				Write-Error [String]::Format("There was an error writing the change to the tenant. DisplayName: {0}, Email: {1}", $setDisplayName, $setEmail)
				Write-Error $_
		}
		
		# Write-Output "Set-UnifiedGroup -Identity $($_.Id) -EmailAddresses @{remove=""$($_.Mail)""}"
		# $msg.Body | ConvertFrom-Json | Select-Object @{Name="Id";Expression={$_.objectId}}, @{Name="DisplayName";Expression={$_.newDisplayName}}, @{Name="Mail";Expression={$_.newEmail}}, @{Name="MailNickname";Expression={$_.mailNickname}},  @{Name="GroupTypes";Expression={"Unified"}} | Export-Csv C:\Projects\Common\Office365\ITServeco.O365.Governance\Runbooks\unit-tests\datasets\azureadmsgroup-11152017-modified.csv -Append
	}

	#foreach($msg in $MessageQueue) {
	#	$msg.Body | ConvertFrom-Json | Select-Object @{Name="Id";Expression={$_.objectId}}, @{Name="DisplayName";Expression={$_.newDisplayName}}, @{Name="Mail";Expression={$_.newEmail}}, @{Name="MailNickname";Expression={$_.mailNickname}},  @{Name="GroupTypes";Expression={"Unified"}} | Export-Csv C:\Projects\Common\Office365\ITServeco.O365.Governance\Runbooks\unit-tests\datasets\azureadmsgroup-11152017-modified.csv -Append
	#}


}

function Add-ServiceBusMessageQueue {
	[CmdletBinding()]	
	param(
		[Parameter(Mandatory = $true)]
		[String]
		$Id,

		[Parameter(Mandatory = $true)]
		[String]
		$DisplayName,

		[Parameter(Mandatory = $true)]
		[String]
		$NewDisplayName,

		[Parameter(Mandatory = $true)]
		[String]
		$Email,

		[Parameter(Mandatory = $true)]
		[String]
		$NewEmail,

		[Parameter(Mandatory = $true)]
		[String]
		$MailNickname,

		[Parameter(Mandatory = $false)]
		[Boolean]
		$NewEmailExists = $false,

		[Parameter(Mandatory = $true)]
		[Object]
		$Owners
	)

	# for duplicate detection. if service bus queue finds an identical
	# message ID, overwrite it with this one
	$messageId = [String]::Format("{0}::{1}", $POLICY_TYPE_ACTION, $Id)

	if(!$MessageKeys.Contains($messageId)) {

		# the JSON body of the message
		$body = @{
			objectId = $Id;
			objectType = "UnifiedGroup";
			displayName = $DisplayName;
			tenant = $TenantFqdn;
			policyActionType = $POLICY_TYPE_ACTION;
			owners = {@()}.Invoke()

			newDisplayName = $NewDisplayName;
			email = $Email;
			newEmail = $NewEmail;
			mailNickname = $MailNickname;
			newEmailExists = $NewEmailExists
		}

		$body.owners = $Owners

		$bodySerialized = $body | ConvertTo-Json

		# Messaging object
		$message = @{
			BrokerProperties = @{
				Label = $POLICY_TYPE_ACTION;
				CorrelationId = $CorrelationId;
				MessageId = $messageId;
				SessionId = $SessionId;
				PartitionKey = $SessionId;
				ContentType = "application/json"
			}

			UserProperties = @{}

			Body = "$bodySerialized"

		}

		$MessageKeys.Add($messageId)
		$MessageQueue.Enqueue($message)

	} else {
		Write-Warning "Duplicate message in queue detected for $messageId.  Skipping."
	}
}

function Get-BaseDisplayName {
	[CmdletBinding()]
	[OutputType([Hashtable])]
	param(
		[Parameter(Mandatory = $true)]
		[String]
		$DisplayName
	)

	process {
		# serialize the domain references
		# for regular expressions
		$strValue = [String]::Empty

		$domainRef = $DomainCompanyJson | ConvertFrom-Json

		foreach($key in $domainRef.PSObject.Properties.Name) {
			$strValue += $domainRef.PSObject.Properties[$key].Value
			$strValue += " |"
		}

		# Assumption made here.  The company id is
		# at the start of the pattern "^".  
		# If there are multiple prefixes (ie. use case when the script
		# has been misconfigured and has run multiple times in error), this
		# strips them all out as well as long as the orgid is specified
		# in the DomainCompanyJson
		#
		# MUST CHANGE if pattern can be anywhere
		#
		$strValue = "^(" + $strValue.Substring(0,$strValue.Length-1) + ")*"

		# Process Display Name
		$myValue = $DisplayName -replace $strValue, [String]::Empty
		$myValue = $myValue.Trim()

		return $myValue
	}
}


function Invoke-UnifiedGroupNamingPolicyViolation {
	[CmdletBinding()]
	param()

	process {
		# Iterate through each Unified Group
		$groups = Get-AzureADMSGroup -All $true | Where-Object { $_.GroupTypes -contains 'Unified' }

		$groups | ForEach-Object {

			Write-Output "Analyzing Office 365 Group: ""$($_.DisplayName)"" ..."

			# Determine the majority organization
			$majorityOrg = Get-UnifiedGroupMajorityOrg -ObjectId $_.Id

			# Process only if the company domain has been
			# determined.  If null, there is no owner
			if($majorityOrg -ne $null) {

				$params = @{
					Id = $_.Id
					DisplayName = $_.DisplayName
					NewDisplayName = [String]::Empty
					Email = $_.Mail
					NewEmail = [String]::Empty
					MailNickname = $_.Mailnickname
					NewEmailExists = $false
				}

				$owners = Get-AzureADGroupOwner -ObjectId $_.Id | Select-Object DisplayName, UserPrincipalName
				$ownerCol = {@()}.Invoke()

				$owners | ForEach-Object {
					$ownerCol.Add(@{
						userPrincipalName = $_.UserPrincipalName
						displayName = $_.DisplayName
					})
				}


				# Display name was in compliant but is changing organization ID?
				# Example: DisplayName - DEV Project ABCD
				# This prevents it from applying DEV DEV Project ABCD
				$baseDisplayName = Get-BaseDisplayName -DisplayName $_.DisplayName

				# Get Group Display Name
				$params["NewDisplayName"] = Set-TemplateValues -Template $GroupDisplayNameTemplate `
					-CompanyId $majorityOrg["CompanyId"] `
					-DisplayName $baseDisplayName `
					-MailNickname $_.MailNickname `
					-CompanyMailDomain $majorityOrg["CompanyMailDomain"]

				# Get Group Email
				$params["NewEmail"] = $(Set-TemplateValues -Template $GroupEmailTemplate `
					-CompanyId $majorityOrg["CompanyId"] `
					-DisplayName $baseDisplayName `
					-MailNickname $_.MailNickname `
					-CompanyMailDomain $majorityOrg["CompanyMailDomain"]).ToLower()


				# Does the new email exist as a proxy address?
				if($_.ProxyAddresses -match [String]::Format("smtp:{0}", $params["NewEmail"])) {
					$params["NewEmailExists"] = $true
				}

				# Is there a change in the name?
				if($params["DisplayName"] -ne $params["NewDisplayName"] -or
					$params["Email"] -ne $params["NewEmail"]) {

					Add-Statistics -Key "Change"
					Add-Statistics -Key $([String]::Format("ORGID:{0}", $majorityOrg["CompanyId"]))

					Add-ServiceBusMessageQueue -Id $params["Id"] `
						-DisplayName $params["DisplayName"] `
						-NewDisplayName $params["NewDisplayName"] `
						-Email $params["Email"] `
						-NewEmail $params["NewEmail"] `
						-MailNickname $params["MailNickname"] `
						-NewEmailExists $params["NewEmailExists"] `
						-Owners $ownerCol

				} else {
					Add-Statistics -Key "Pass"
					Write-Output "Group name is OK!"
				}

			} else {
				Add-Statistics -Key "Inconclusive"
				Write-Output "An organization cannot be identified.  The group may not have any owners.  Skipping."
			}
		}

	}
}


###############################
# START SCAN

if(!$DEBUG_MODE) {
	Connect-Office365 -ExOnline $true
}

if(!(Test-TenantDomainVanityCheck)) {
	Write-Error "The service account is not associated with the $($TenantFqdn) tenant."
	exit
}

if($ReadOnly) {
	Write-Output "Running in READ ONLY mode.  No changes will be made."
} else {
	Write-Output "Running in WRITE mode.  Changes will be made."
}

# Find all violations to the naming policy
Invoke-UnifiedGroupNamingPolicyViolation

if($Statistics["Change"] -le $MessagingLimit) {
	Send-ServiceBusQueue
	Invoke-UnifiedGroupNamingPolicyAction
} else {
	Write-Warning "Messaging threshold reached at $($Statistics["Change"]) changes.  Name changes aborted.  The threshold is $MessagingLimit.  Change the MessagingLimit parameter to a higher value to allow processing."
}

if(!$DEBUG_MODE) { 
	Disconnect-Office365 -ExOnline $true 
}

Write-Output $Statistics