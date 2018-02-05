[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)]
	[string]
	$TenantName = "devflin",

	[Parameter(Mandatory = $false)]
	[string]
	$CredentialId = "Tenant Automation",

	[Parameter(Mandatory = $false)]
	[int]
	$MinOwnerCount = 2


)

# 0: INITIALIZE WORKFLOW

# Static Attributes
$POLICY_TYPE_ACTION = "UnifiedGroupMinOwner"
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
			owners = {@()}.Invoke();
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


###############################
# START SCAN

if(!$DEBUG_MODE) {
	Connect-Office365
}

if(!(Test-TenantDomainVanityCheck)) {
	Write-Error "The service account is not associated with the $($TenantFqdn) tenant."
	exit
}

$groups = Get-AzureADMSGroup -All $true

$groups | Where-Object { $_.GroupTypes -contains "Unified" } | ForEach-Object {

	Write-Output "Scanning Unified Group ""$($_.DisplayName)"" ... "

	$owners = Get-AzureADGroupOwner -ObjectId $_.Id
	$ownerCol = {@()}.Invoke()

	if($owners.Count -lt $MinOwnerCount) {

		# Add all owners to messaging object
		foreach($owner in $owners) {
			$ownerCol.Add(@{
				userPrincipalName = $owner.UserPrincipalName
				displayName = $owner.DisplayName
			})
		}

		Write-Output "Unified Group has $($owners.Count) owner.  Policy violation."
		Add-Statistics -Key "Change"
		Add-ServiceBusMessageQueue -Id $_.Id -DisplayName $_.DisplayName -Owners $ownerCol
	} else {
		Write-Output "Unified Group has $($owners.Count) owners.  OK!"
		Add-Statistics -Key "Pass"
	}
}

Send-ServiceBusQueue

if(!$DEBUG_MODE) { 
	Disconnect-Office365
}

Write-Output $Statistics

###############################
