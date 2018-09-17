###################
# PRIVATE FUNCTIONS
###################
function Test-PrivateIPAddress ([System.Net.IPAddress]$IPAddress) {
	$PrivateA = "^(10\.){1}(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){2}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
	$PrivateB = "^(172\.){1}(?:(?:1[6-9]|2[0-9]|31?)\.){1}(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.)(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
	$PrivateC = "^(192\.168\.){1}(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){1}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"

	$PrivateIPAddress = $false

	if (($IPAddress -imatch $PrivateA) -or ($IPAddress -imatch $PrivateB) -or ($IPAddress -imatch $PrivateC)) {
		$PrivateIPAddress = $true
	} else {
		$PrivateIPAddress = $true
	}

	Write-Output $PrivateIPAddress 
}

function Get-ComputerDomain ([string]$ComputerName) {

	$Win32_ComputerSystem = Get-WmiObject -Namespace root\CIMV2 -Class Win32_ComputerSystem -Property Domain, DomainRole -ComputerName $ComputerName
	$ComputerDomain = $null

	if ($Win32_ComputerSystem.DomainRole) {
		$ComputerDomain = $Win32_ComputerSystem.Domain
	} else {
		$ComputerDomain = $null
	}

	Write-Output $ComputerDomain

	# The .NET way
	# http://msdn.microsoft.com/en-us/library/system.directoryservices.activedirectory.domain.aspx
	# This doesn't work if the host computer is not joined to a domain or unable to contact the AD controller
	#Write-Output ([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain())
}

function Get-ActiveDirectoryDnsConfiguration {
	[CmdletBinding()]
	param
	(
		[Parameter(
			ValueFromPipeline=$True,
			ValueFromPipelineByPropertyName=$True,
			HelpMessage='What computer name would you like to target?')]
		[Alias('host')]
		[ValidateLength(3,30)]
		[string[]]$ComputerName = "localhost"
	)
	process {

		$ActiveDirectoryDnsConfiguration = @()
		$ComputerDomain = $null
		$Domain = $null
		$Computer = $null

		foreach ($Computer in $ComputerName) {

			$ComputerDomain = Get-ComputerDomain -ComputerName $Computer

			Get-WmiObject -Namespace root\CIMV2 -Class Win32_NetworkAdapterConfiguration `
			-Property DnsServerSearchOrder, DNSDomain `
			-Filter "(IPEnabled = True) and (DNSDomain <> 'domain_not_set.invalid') and (DNSDomain <> '')" `
			-ComputerName $Computer | 
			ForEach-Object {

				# Use the Computer's Domain unless there's a specific domain for this adapter (e.g. VPN connections)
				if ($_.DNSDomain) {
					$Domain = $_.DNSDomain
				} else {
					$Domain = $ComputerDomain
				}

				if (($Domain) -and ($_.DnsServerSearchOrder)) {
					$ActiveDirectoryDnsConfiguration += (
						New-Object -TypeName psobject -Property @{ 
							ComputerName = $Computer
							Domain = $Domain
							DnsServer = $_.DnsServerSearchOrder
						}
					)
				}
			}
		}
		Write-Output $ActiveDirectoryDnsConfiguration
	}
}

function Get-IPv4SubnetConfiguration {
	[CmdletBinding()]
	param
	(
		[Parameter(
			ValueFromPipeline=$True,
			ValueFromPipelineByPropertyName=$True,
			HelpMessage='What computer name would you like to target?')]
		[Alias('host')]
		[ValidateLength(3,30)]
		[string[]]$ComputerName = $env:COMPUTERNAME
	)
	process {
		$IPv4SubnetConfiguration = @()
		$pos = $null

		Get-WmiObject -Namespace root\CIMV2 `
		-Class Win32_NetworkAdapterConfiguration `
		-Property IPAddress, IPSubnet `
		-Filter '(IPEnabled = True)' `
		-ComputerName $ComputerName | 
		ForEach-Object {
			for ($pos = 0; $pos -lt @($_.IPAddress).Count; $pos++) { 

				# Only work w\ IPv4 addresses
				if (($_.IPAddress[$pos] -as [System.Net.IPAddress]).AddressFamily -ieq 'InterNetwork') {
					$IPv4SubnetConfiguration += (
						New-Object -TypeName psobject -Property @{ 
							IPAddress = $_.IPAddress[$pos]
							SubnetMask = $_.IPSubnet[$pos]
						}
					) 
				}
			}
		}

		Write-Output $IPv4SubnetConfiguration
	}
}

function Get-DnsARecord {
	[CmdletBinding()]
	param
	(
		[Parameter(
			Mandatory=$true,
			HelpMessage='What is the IP Address of the DNS Server you would like to target?')]
		[Alias('ip')]
		[System.Net.IPAddress[]]$DnsServerIPAddress
		,
		[Parameter(
			Mandatory=$true,
			HelpMessage='What is the Domain you would like to target?')]
		[string]$Domain
	)
	begin {
		$DnsServer = $null
	}
	process {
		Write-Output (
			$DnsServerIPAddress | ForEach-Object {
				try {
					$DnsServer = $_
					Get-WMIObject -Namespace root\MicrosoftDNS -Class MicrosoftDNS_AType -Computer $_ -Filter "ContainerName= '$Domain'" -ErrorAction Stop | 
					Where-Object {$_.DomainName -ine $_.OwnerName } | 
					Select-Object OwnerName, IPAddress | 
					ForEach-Object {
						New-Object -TypeName psobject -Property @{ 
							OwnerName = $_.OwnerName
							IPAddress = $_.IPAddress
						}
					}
				}
				catch {
					$ThisException = $_.Exception
					while ($ThisException.InnerException) {
						$ThisException = $ThisException.InnerException
					}
					Write-NetworkScanLog -Message "Error querying for A records from DNS Server $DnsServer for domain '$Domain': $($ThisException.Message)" -MessageLevel Warning
				}
			} | Select-Object -Property OwnerName, IPAddress -Unique
		)
	}
	end {
		Remove-Variable -Name DnsServer
	}
}

function Get-DeviceScanObject {
	[CmdletBinding()]
	[OutputType([System.Int32])]
	param(
		[Parameter(Mandatory=$false)]
		[System.String]
		$DnsRecordName = $null
		,
		[Parameter(Mandatory=$false)]
		[System.String]
		$WmiMachineName = $null
		,
		[Parameter(Mandatory=$true)]
		[System.Net.IPAddress]
		$IPAddress
		,
		[Parameter(Mandatory=$false)]
		[Boolean]
		$IsPingAlive = $false
		,
		[Parameter(Mandatory=$false)]
		[Boolean]
		$IsWmiAlive = $false
	)
	process {
		Write-Output (
			New-Object -TypeName PSObject -Property @{
				DnsRecordName = $DnsRecordName
				WmiMachineName = $WmiMachineName
				IPAddress = $IPAddress
				IsPingAlive = $IsPingAlive
				IsWmiAlive = $IsWmiAlive 
			}
		)
	}
}

function Write-NetworkScanLog {
	[CmdletBinding()]
	param(
		[Parameter(Position=0, Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$Message
		,
		[Parameter(Position=1, Mandatory=$true)] 
		[alias('level')]
		[ValidateSet('information','verbose','debug','error','warning')]
		[System.String]
		$MessageLevel
	)
	try {
		if ((Test-Path -Path 'function:Write-Log') -eq $true) {
			Write-Log -Message $Message -MessageLevel $MessageLevel
		}
	}
	catch {
		Write-Host $Message
		#Throw
	}
}



###################
# PUBLIC FUNCTIONS
###################

function Find-IPv4Device {
	<#
	.SYNOPSIS
		Synchronously look for IPv4 devices on a network.

	.DESCRIPTION
		This function executes a synchronus scan for IPv4 devices on a network.
		
		When an IPv4 device is found connectivity via Windows Management Interface (WMI) is also verified.

	.PARAMETER  DnsServer
		'Automatic', or the Name or IP address of an Active Directory DNS server to query for a list of hosts to test for connectivity

		When 'Automatic' is specified the function will use WMI queries to discover the current computer's DNS server(s) to query.
		
	.PARAMETER  DnsDomain
		'Automatic' or the Active Directory domain name to use when querying DNS for a list of hosts.
		
		When 'Automatic' is specified the function will use the current computer's AD domain.
		
		'Automatic' will be used by default if DnsServer is specified but DnsDomain is not provided.

	.PARAMETER  Subnet
		'Automatic' or a comma delimited list of subnets (in CIDR notation) to scan for connectivity.
		
		When 'Automatic' is specified the function will use the current computer's IP configuration to determine subnets to scan. 
		
		A quick refresher on CIDR notation:

			BITS	SUBNET MASK			USABLE HOSTS PER SUBNET
			----	---------------		-----------------------
			/20		255.255.240.0		4094
			/21		255.255.248.0		2046 
			/22		255.255.252.0		1022
			/23		255.255.254.0		510 
			/24		255.255.255.0		254 
			/25		255.255.255.128		126 
			/26		255.255.255.192		62
			/27		255.255.255.224		30
			/28		255.255.255.240		14
			/29		255.255.255.248		6
			/30		255.255.255.252		2
			/32		255.255.255.255		1		

	.PARAMETER  ComputerName
		A comma delimited list of computer names to test for connectivity.
		
	.PARAMETER  ExcludeSubnet
		A comma delimited list of subnets (in CIDR notation) to exclude when testing for connectivity.
		
	.PARAMETER  LimitSubnet
		A comma delimited list of subnets (in CIDR notation) to limit the scope of connectivity tests. Only hosts with IP Addresses that fall within the specified subnet(s) will be included in the results.

	.PARAMETER  ExcludeComputerName
		A comma delimited list of computer names to exclude when testing for connectivity. Wildcards are accepted.
		
		An attempt will be made to resolve the IP Address(es) for each computer in this list and those addresses will also be used when determining if a host should be included or excluded when testing for connectivity.		
		
	.PARAMETER  MaxConcurrencyThrottle
		Number between 1-100 to indicate how many instances to collect information from concurrently.

		If not provided then the number of logical CPUs present to your session will be used.

	.PARAMETER  PrivateOnly
		Only include hosts with private class A, B, or C IP addresses

	.PARAMETER  ResolveAliases
		When a mismatch between host name and WMI machine name occurs query DNS for the machine name
		
	.PARAMETER  ParentProgressId
		If the caller is using Write-Progress then all progress information will be written using ParentProgressId as the ParentID		

	.EXAMPLE
		Find-IPv4Device -DNSServer automatic -DNSDomain automatic -PrivateOnly
		
		Description
		-----------
		Queries Active Directory for a list of hosts to scan for IPv4 connectivity. The list of hosts will be restricted to private IP addresses only.

	.EXAMPLE
		Find-IPv4Device -Subnet 172.20.40.0/28
		
		Description
		-----------
		Scans all hosts in the subnet 172.20.40.0/28 for IPv4 connectivity.
		
	.EXAMPLE
		Find-IPv4Device -Computername Server1,Server2,Server3
		
		Description
		-----------
		Scanning Server1, Server2, and Server3 for IPv4 connectivity.

	.OUTPUTS
		System.Management.Automation.PSObject

	.NOTES

#>
	[CmdletBinding(DefaultParametersetName='dns')]
	param(
		[Parameter(
			Mandatory=$true,
			ParameterSetName='dns',
			HelpMessage='DNS Server(s)'
		)] 
		[alias('dns')]
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$|^auto$|^automatic$')]
		[string[]]
		$DnsServer = 'automatic'
		,
		[Parameter(
			Mandatory=$false,
			ParameterSetName='dns',
			HelpMessage='DNS Domain Name'
		)] 
		[alias('domain')]
		[string]
		$DnsDomain = 'automatic'
		,
		[Parameter(
			Mandatory=$true,
			ParameterSetName='subnet',
			HelpMessage='Subnet (in CIDR notation)'
		)] 
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)[\\/]\d{1,2}$|^auto$|^automatic$')]
		[string[]]
		$Subnet = 'Automatic'
		,
		[Parameter(
			Mandatory=$true,
			ParameterSetName='computername',
			HelpMessage='Computer Name(s)'
		)] 
		[alias('Computer')]
		[string[]]
		$ComputerName
		,
		[Parameter(Mandatory=$false, ParameterSetName='dns')]
		[Parameter(Mandatory=$false, ParameterSetName='subnet')]
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)[\\/]\d{1,2}$')]
		[string[]]
		$ExcludeSubnet
		,
		[Parameter(Mandatory=$false, ParameterSetName='dns')]
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)[\\/]\d{1,2}$')]
		[string[]]
		$LimitSubnet
		,
		[Parameter(Mandatory=$false, ParameterSetName='dns')]
		[Parameter(Mandatory=$false, ParameterSetName='subnet')]
		[string[]]
		$ExcludeComputerName
		,
		[Parameter(Mandatory=$false)] 
		[ValidateRange(1,100)]
		[alias('Throttle')]
		[byte]
		$MaxConcurrencyThrottle = $env:NUMBER_OF_PROCESSORS
		,
		[Parameter(Mandatory=$false)] 
		[switch]
		$PrivateOnly = $false
		,
		[Parameter(Mandatory=$false)] 
		[switch]
		$ResolveAliases = $false
		,
		[Parameter(Mandatory=$false)]
		[ValidateNotNull()]
		[Int32]
		$ParentProgressId = -1
	)
	process {

		$Device = @{}
		$ActiveDirDnsConfig = $null
		$DnsARecord = $null
		$PingAliveDevice = $null
		$IPv4SubnetConfig = @()
		$LimitSubnetNetwork = @()
		$ExcludeSubnetNetwork = @()
		$ExcludeIpAddress = @()
		$IPAddress = $null
		$DecimalAddress = $null
		$HostName = $null
		$ScanCount = 0
		$DeviceCount = 0
		$HasMetCriteria = $false

		$PingProgressId = Get-Random
		$WmiProgressId = Get-Random

		# For use with runspaces
		$ScriptBlock = $null
		$SessionState = $null
		$RunspacePool = $null
		$Runspaces = $null
		$PowerShell = $null
		$HashKey = $null

		# Fallback in case value isn't supplied or somehow missing from the environment variables
		if (-not $MaxConcurrencyThrottle) { $MaxConcurrencyThrottle = 1 }

		# Log start status
		Write-NetworkScanLog -Message 'Start Function: Find-IPv4Device' -MessageLevel Debug
		Write-NetworkScanLog -Message 'Beginning network scan' -MessageLevel Information

		switch ($PsCmdlet.ParameterSetName) {
			'dns' {
				Write-NetworkScanLog -Message "`t-DnsServer: $([String]::Join(',',$DnsServer))" -MessageLevel Information
				Write-NetworkScanLog -Message "`t-DnsDomain: $DnsDomain" -MessageLevel Information

			}
			'subnet' {
				Write-NetworkScanLog -Message "`t-Subnet: $([String]::Join(',',$Subnet))" -MessageLevel Information

				# Initialize variables that don't yet exist in this parameterset
				#$LimitSubnet = $false
			}
			'computername' {
				Write-NetworkScanLog -Message "`t-ComputerName: $([String]::Join(',',$ComputerName))" -MessageLevel Information

				# Initialize variables that don't yet exist in this parameterset
				#$ExcludeSubnet = $null
				#$LimitSubnet = $null
				#$ExcludeComputerName = $null
			}
		}
		if ($ExcludeSubnet) { Write-NetworkScanLog -Message "`t-ExcludeSubnet: $([String]::Join(',',$ExcludeSubnet))" -MessageLevel Information }
		if ($LimitSubnet) { Write-NetworkScanLog -Message "`t-LimitSubnet: $([String]::Join(',',$LimitSubnet))" -MessageLevel Information }
		if ($ExcludeComputerName) { Write-NetworkScanLog -Message "`t-ExcludeComputerName: $([String]::Join(',',$ExcludeComputerName))" -MessageLevel Information }
		Write-NetworkScanLog -Message "`t-PrivateOnly: $PrivateOnly" -MessageLevel Information
		Write-NetworkScanLog -Message "`t-MaxConcurrencyThrottle: $MaxConcurrencyThrottle" -MessageLevel Information
		Write-NetworkScanLog -Message "`t-ResolveAliases: $ResolveAliases" -MessageLevel Information



		# Get network summaries for subnet filtering if exclude\include subnets were provided
		if ($LimitSubnet) {
			$LimitSubnet | ForEach-Object {
				$LimitSubnetNetwork += (Get-NetworkSummary -Network $_)
			}
		}
		if ($ExcludeSubnet) {
			$ExcludeSubnet | ForEach-Object {
				$ExcludeSubnetNetwork += (Get-NetworkSummary -Network $_)
			}
		}

		if ($ExcludeComputerName) {

			# Add 'localhost' to the list of excluded computers if $ExcludeComputerName contains the local machine
			# Fringe case, I know...but I'm OCD so I put it in here.
			if ($ExcludeComputerName -icontains $env:COMPUTERNAME) {
				$ExcludeComputerName += 'localhost'
			}

			try {
				$ExcludeComputerName | 
				Where-Object { $_.IndexOfAny(@('/','\','[',']','"',':',';','|','<','>','+','=',',','?','*',' ','_')) -eq -1 } |
				Select-Object -Unique |
				ForEach-Object {
					# Get IP Addresses for the host name
					# Limit to IPv4 addresses
					[System.Net.Dns]::GetHostAddresses($_) | Where-Object { $_.AddressFamily -ieq 'InterNetwork' } | ForEach-Object {
						$ExcludeIpAddress += $_.IPAddressToString
					}
				}
			}
			catch {
			}
		}


		# Get list of IP addresses to scan
		#region
		switch ($PsCmdlet.ParameterSetName) {
			'dns' {

				# For automatic DNSDomain get this computer's domain setting
				if ($DnsDomain -ilike 'auto*') {
					$DnsDomain = Get-ComputerDomain -ComputerName localhost
				} 

				if ($DnsServer -ilike 'auto*') {
					$ActiveDirDnsConfig = @(Get-ActiveDirectoryDnsConfiguration -ComputerName localhost)
				} else {
					$ActiveDirDnsConfig = @(
						New-Object -TypeName psobject -Property @{ 
							ComputerName = $Computer
							Domain = $DnsDomain
							DnsServer = $DnsServer
						}
					)
				}

				$ActiveDirDnsConfig | ForEach-Object {

					Write-NetworkScanLog -Message "Querying for A records from DNS Server $($_.DnsServer) for domain '$($_.Domain)'" -MessageLevel Information

					$DnsARecord = Get-DnsARecord -DnsServerIPAddress $_.DnsServer -Domain $_.Domain

					Write-NetworkScanLog -Message "Found $($($DnsARecord | Measure-Object).Count) DNS A records" -MessageLevel Information

					$DnsARecord | Where-Object { $_.IPAddress } | ForEach-Object {

						$HasMetCriteria = $true

						if (
							# Remember PowerShell short circuits -and and -or
							$PrivateOnly -ne $true -or
							$(Test-PrivateIPAddress -IPAddress $_.IPAddress) -eq $true
						) {

							$HostName = $_.OwnerName.ToUpper()
							$IPAddress = $_.IPAddress

							# Check if the host name is in the list of excluded computer names
							if ($ExcludeComputerName) {
								$ExcludeComputerName | ForEach-Object {
									if ($HostName -ilike $_) {
										$HasMetCriteria = $false
										break
									}
								}

								# Also check that the IP address isn't in the list of IPs to exclude
								if ($ExcludeIpAddress -icontains $IPAddress) {
									$HasMetCriteria = $false
								}

							}

							# Check if the host IP is within the ranges for $LimitSubnet and $ExcludeSubnet
							if ($HasMetCriteria -and 
								(									$LimitSubnet -or $ExcludeSubnet)
							) {

								$DecimalAddress = ConvertTo-DecimalIP -IPAddress $IPAddress

								if ($LimitSubnet) {
									$LimitSubnetNetwork | ForEach-Object {
										if (($DecimalAddress -le $_.NetworkDecimal) -or ($DecimalAddress -ge $_.BroadcastDecimal)) {
											$HasMetCriteria = $false
										}
									}
								}

								if ($ExcludeSubnet) {
									$ExcludeSubnetNetwork | ForEach-Object {
										if (($DecimalAddress -ge $_.NetworkDecimal) -and ($DecimalAddress -lt $_.BroadcastDecimal)) {
											$HasMetCriteria = $false
										}
									}
								}

							}


						}
						else {
							$HasMetCriteria = $false
						}

						if ($HasMetCriteria -eq $true) {
							if (-not ($Device.Values | Where-Object { ($_.DnsRecordName -ieq $HostName) -and ($_.IPAddress -ieq $IPAddress) })) {
								$Device.Add([guid]::NewGuid(), (Get-DeviceScanObject -DnsRecordName $HostName -WmiMachineName $null -IPAddress $IPAddress -IsPingAlive $false -IsWmiAlive $false))
							} 
						}
					}
				}
			}
			'subnet' { 
				if ($Subnet -ilike 'auto*') {
					$IPv4SubnetConfig = @(Get-IPv4SubnetConfiguration -ComputerName localhost)
				} else {
					$Subnet | Select-Object -Unique | ForEach-Object { 
						$IPv4SubnetConfig += (
							Get-NetworkSummary -Network $_ | ForEach-Object {
								New-Object -TypeName psobject -Property @{ 
									IPAddress = $_.NetworkAddress
									SubnetMask = $_.Mask
								}
							}
						) 
					}
				} 

				$IPv4SubnetConfig | ForEach-Object {

					Write-NetworkScanLog -Message "Resolving addresses for network $($_.IPAddress) with mask $($_.SubnetMask)" -MessageLevel Information

					if ((($PrivateOnly -eq $true) -and ((Test-PrivateIPAddress -IPAddress $_.IPAddress) -eq $true)) -or ($PrivateOnly -eq $false)) {
						$(
							if ($_.SubnetMask -ieq '255.255.255.255') {
								@( $_.IPAddress )
							} else {
								Get-NetworkRange -IPAddress $_.IPAddress -SubnetMask $_.SubnetMask
							}
						) | ForEach-Object {
							$IPAddress = $_
							$HasMetCriteria = $true

							# Check if the host name is in the list of excluded computer names
							if ($ExcludeComputerName) {
								# Check that the IP address isn't in the list of IPs to exclude
								if ($ExcludeIpAddress -icontains $IPAddress) {
									$HasMetCriteria = $false
								}
							}

							if ($HasMetCriteria -and $ExcludeSubnet) {
								$DecimalAddress = ConvertTo-DecimalIP -IPAddress $IPAddress

								$ExcludeSubnetNetwork | ForEach-Object {
									if (($DecimalAddress -ge $_.NetworkDecimal) -and ($DecimalAddress -lt $_.BroadcastDecimal)) {
										$HasMetCriteria = $false
									}
								}
							}

							# Skip checking - this was causing major performance problems for larger subnets
							# and unless overlapping subnets were supplied there's minimal chance of duplicate addresses
							#if (-not ($Device.Values | Where-Object { ($_.IPAddress -ieq $IPAddress) })) {

							if ($HasMetCriteria -eq $true) {
								$Device.Add([guid]::NewGuid(), (Get-DeviceScanObject -DnsRecordName $null -WmiMachineName $null -IPAddress $IPAddress -IsPingAlive $false -IsWmiAlive $false)) 
							}
							#}
						}
					}
				}

			}
			'computername' {

				$ComputerName | Select-Object -Unique | ForEach-Object {

					$HostName = $_.ToUpper()
					Write-NetworkScanLog -Message "Resolving IP address for $HostName" -MessageLevel Information

					try {
						$IPAddress = $null

						[System.Net.Dns]::GetHostByName($HostName) | ForEach-Object {

							$HostName = $_.HostName.ToUpper() # Use value from results because it contains the FQDN

							$_.AddressList | Where-Object { $_.AddressFamily -ieq 'InterNetwork' } | ForEach-Object {
								if ((($PrivateOnly -eq $true) -and ((Test-PrivateIPAddress -IPAddress $_.IPAddressToString) -eq $true)) -or ($PrivateOnly -eq $false)) {
									$Device.Add([guid]::NewGuid(), (Get-DeviceScanObject -DnsRecordName $HostName -WmiMachineName $null -IPAddress $_.IPAddressToString -IsPingAlive $false -IsWmiAlive $false)) 
								}
							}
						}
					} catch {
						Write-NetworkScanLog -Message "Error resolving IP address for $($HostName): $($_.Exception.InnerException.Message)" -MessageLevel Information
					}
				}
			}
		}
		#endregion

		# Create a Session State, Create a RunspacePool, and open the RunspacePool
		$SessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
		$RunspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $MaxConcurrencyThrottle, $SessionState, $Host)
		$RunspacePool.Open()

		# Create an empty collection to hold the Runspace jobs
		$Runspaces = New-Object System.Collections.ArrayList 

		$DeviceCount = $Device.Count # Not using Measure-Object we know this starts as an empty hash
		$ScanCount = 0

		Write-NetworkScanLog -Message "Testing PING connectivity to $DeviceCount addresses" -MessageLevel Information
		Write-Progress -Activity 'Testing PING connectivity' -PercentComplete 0 -Status "Testing $DeviceCount Addresses" -Id $PingProgressId -ParentId $ParentProgressId

		<# PING CONNECTIVITY TEST #>
		#region

		$ScanCount = 0
		$ScriptBlock = {
			Param (
				[String]$ComputerName
			)
			Test-Connection -ComputerName $ComputerName -Count 3 -Quiet
		}


		# Queue up PING tests
		$Device.GetEnumerator() | ForEach-Object {

			$ScanCount++

			# Update progress
			if ($($_.Value).DnsRecordName) {
				Write-NetworkScanLog -Message "Testing PING connectivity to $($($_.Value).DnsRecordName) ($($($_.Value).IPAddress)) [$ScanCount of $DeviceCount]" -MessageLevel Verbose
			} elseif ($($_.Value).WmiMachineName) {
				Write-NetworkScanLog -Message "Testing PING connectivity to $($($_.Value).WmiMachineName) ($($($_.Value).IPAddress)) [$ScanCount of $DeviceCount]" -MessageLevel Verbose
			} else {
				Write-NetworkScanLog -Message "Testing PING connectivity to IP address $($($_.Value).IPAddress) [$ScanCount of $DeviceCount]" -MessageLevel Verbose
			}

			# Test connectivity to the machine 

			#Create the PowerShell instance and supply the scriptblock with the other parameters
			$PowerShell = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
			$PowerShell = $PowerShell.AddArgument($($_.Value).IPAddress)

			#Add the runspace into the PowerShell instance
			$PowerShell.RunspacePool = $RunspacePool

			$Runspaces.Add((
					New-Object -TypeName PsObject -Property @{
						PowerShell = $PowerShell
						Runspace = $PowerShell.BeginInvoke()
						HashKey = $_.Key
					}
				)) | Out-Null

		}

		# Reset the scan counter
		$ScanCount = 0

		# Process results as they complete
		Do {
			$Runspaces | ForEach-Object {

				If ($_.Runspace.IsCompleted) {
					try {
						$HashKey = $_.HashKey

						# This is where the output gets returned
						$_.PowerShell.EndInvoke($_.Runspace) | ForEach-Object {
							$Device[$HashKey].IsPingAlive = $_
						} 
					}
					catch { }
					finally {
						# Cleanup
						$_.PowerShell.dispose()
						$_.Runspace = $null
						$_.PowerShell = $null

						if ($Device[$HashKey].DnsRecordName) {
							Write-NetworkScanLog -Message "PING response from $($Device[$HashKey].DnsRecordName) ($($Device[$HashKey].IPAddress)): $($Device[$HashKey].IsPingAlive)" -MessageLevel Verbose
						} elseif ($Device[$HashKey].WmiMachineName) {
							Write-NetworkScanLog -Message "PING response from $($Device[$HashKey].WmiMachineName) ($($Device[$HashKey].IPAddress)): $($Device[$HashKey].IsPingAlive)" -MessageLevel Verbose
						} else {
							Write-NetworkScanLog -Message "PING response from $($Device[$HashKey].IPAddress): $($Device[$HashKey].IsPingAlive)" -MessageLevel Verbose
						}
					}
				}
			}

			# Found that in some cases an ACCESS_VIOLATION error occurs if we don't delay a little bit during each iteration 
			Start-Sleep -Milliseconds 250

			# Clean out unused runspace jobs
			$Runspaces.clone() | Where-Object { ($_.Runspace -eq $Null) } | ForEach {
				$Runspaces.remove($_)
				$ScanCount++
				Write-Progress -Activity 'Testing PING connectivity' -PercentComplete (($ScanCount / $DeviceCount)*100) -Status "$ScanCount of $DeviceCount" -Id $PingProgressId -ParentId $ParentProgressId
			}

		} while (($Runspaces | Where-Object {$_.Runspace -ne $Null} | Measure-Object).Count -gt 0)
		#endregion


		# Count how many devices responded
		$PingAliveDevice = @($Device.GetEnumerator() | Where-Object { $($_.Value).IsPingAlive -eq $true })
		$DeviceCount = $($PingAliveDevice | Measure-Object).Count

		Write-NetworkScanLog -Message 'PING connectivity test complete' -MessageLevel Verbose
		Write-Progress -Activity 'Testing PING connectivity' -PercentComplete 100 -Status "$ScanCount addresses tested, $DeviceCount replies" -Id $PingProgressId -ParentId $ParentProgressId
		Write-NetworkScanLog -Message "Testing WMI connectivity to $DeviceCount addresses" -MessageLevel Information
		Write-Progress -Activity 'Testing WMI connectivity' -PercentComplete 0 -Status "Testing $DeviceCount Addresses" -Id $WmiProgressId -ParentId $ParentProgressId


		<# WMI CONNECTIVITY TEST #>
		#region

		$ScanCount = 0
		$ScriptBlock = {
			Param (
				[Net.IPAddress]$IPAddress,
				[switch]$IncludeDomainName = $false
			)

			<#
			function Get-WMIObjectWithTimeout {
				[CmdletBinding()]
				param(
					[Parameter(Mandatory=$false)]
					[Alias('ns')]
					[ValidateNotNullOrEmpty()]
					[System.String]
					$NameSpace = 'root\CIMV2'
					,
					[Parameter(Mandatory=$true)]
					[ValidateNotNull()]
					[System.String]
					$Class
					,
					[Parameter(Mandatory=$false)]
					[ValidateNotNull()]
					[System.String[]]
					$Property = @('*')
					,
					[Parameter(Mandatory=$false)]
					[Alias('cn')]
					[ValidateNotNull()]
					[System.String]
					$ComputerName = '.'
					,
					[Parameter(Mandatory=$false)]
					[ValidateNotNull()]
					[System.String]
					$Filter
					,
					[Parameter(Mandatory=$false)]
					[Alias('timeout')]
					[ValidateRange(1,3600)]
					[Int]
					$TimeoutSeconds = 600
				)
				try {
					$WmiSearcher = [WMISearcher]''
					$Query = 'select ' + [String]::Join(',',$Property) + ' from ' + $Class

					if ($Filter) {
						$Query = "$Query where $Filter"
					}
					$WmiSearcher.Options.Timeout = [TimeSpan]::FromSeconds($TimeoutSeconds)
					$WmiSearcher.Options.ReturnImmediately = $true
					$WmiSearcher.Scope.Path = "\\$ComputerName\$NameSpace"
					$WmiSearcher.Query = $Query
					$WmiSearcher.Get() 
				}
				catch {
					Throw
				}
			}

			$Win32_ComputerSystem = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_ComputerSystem -Property Name -ComputerName $IPAddress -TimeoutSeconds 60 -ErrorAction Stop
			#>

			$Win32_ComputerSystem = Get-WMIObject -Namespace root\CIMV2 -Class Win32_ComputerSystem -Property Name -ComputerName $IPAddress -ErrorAction Stop
			$ComputerName = $Win32_ComputerSystem.Name
			Write-Output $ComputerName
		}

		$PingAliveDevice | ForEach-Object {

			$ScanCount++

			if ($($_.Value).DnsRecordName) {
				Write-NetworkScanLog -Message "Testing WMI connectivity to $($($_.Value).DnsRecordName) ($($($_.Value).IPAddress)) [$ScanCount of $DeviceCount]" -MessageLevel Verbose
			} elseif ($($_.Value).WmiMachineName) {
				Write-NetworkScanLog -Message "Testing WMI connectivity to $($($_.Value).WmiMachineName) ($($($_.Value).IPAddress)) [$ScanCount of $DeviceCount]" -MessageLevel Verbose
			} else {
				Write-NetworkScanLog -Message "Testing WMI connectivity to IP address $($($_.Value).IPAddress) [$ScanCount of $DeviceCount]" -MessageLevel Verbose
			}

			# Create the PowerShell instance and supply the scriptblock with the other parameters
			$PowerShell = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
			$PowerShell = $PowerShell.AddArgument($($_.Value).IPAddress)

			# Add the runspace into the PowerShell instance
			$PowerShell.RunspacePool = $RunspacePool

			$Runspaces.Add((
					New-Object -TypeName PsObject -Property @{
						PowerShell = $PowerShell
						Runspace = $PowerShell.BeginInvoke()
						HashKey = $_.Key
					}
				)) | Out-Null

		}

		# Reset the scan counter
		$ScanCount = 0

		# Process results
		Do {
			$Runspaces | ForEach-Object {
				If ($_.Runspace.IsCompleted) {
					try {

						$_.PowerShell.EndInvoke($_.Runspace) | ForEach-Object {
							$HostName = $_
						} # This is where the output gets returned

						$HashKey = $_.HashKey

						$Device[$HashKey].WmiMachineName = $HostName # This is where the output gets returned
						$Device[$HashKey].IsWmiAlive = $true


						# Double check that we've got a DNS Hostname. If not, get it from DNS
						# If we do have a DNS Hostname that doesn't begin with the WMI Machine Name and $ResolveAliases is true, get the machine name from DNS
						if (
							(								!$Device[$HashKey].DnsRecordName) `
							-or `
							(								(									$ResolveAliases -eq $true) -and ($Device[$HashKey].DnsRecordName.StartsWith($HostName, 'CurrentCultureIgnoreCase') -ne $true))
						) {
							try {
								[System.Net.Dns]::GetHostByName($HostName) | ForEach-Object {
									$Device[$HashKey].DnsRecordName = $_.HostName.ToUpper()
								}
							} catch {
								# Fallback to using the WMI machine name as the DNS host name in the event there's an error
								$Device[$HashKey].DnsRecordName = $HostName
							}
						}
					}
					catch {}
					finally {
						# Cleanup
						$_.PowerShell.dispose()
						$_.Runspace = $null
						$_.PowerShell = $null

						if ($Device[$HashKey].DnsRecordName) {
							Write-NetworkScanLog -Message "WMI response from $($Device[$HashKey].DnsRecordName) ($($Device[$HashKey].IPAddress)): $($Device[$HashKey].IsWmiAlive)" -MessageLevel Verbose
						} elseif ($Device[$HashKey].WmiMachineName) {
							Write-NetworkScanLog -Message "WMI response from $($Device[$HashKey].WmiMachineName) ($($Device[$HashKey].IPAddress)): $($Device[$HashKey].IsWmiAlive)" -MessageLevel Verbose
						} else {
							Write-NetworkScanLog -Message "WMI response from $($Device[$HashKey].IPAddress): $($Device[$HashKey].IsWmiAlive)" -MessageLevel Verbose
						}
					}
				}
			}

			# Found that in some cases an ACCESS_VIOLATION error occurs if we don't delay a little bit during each iteration 
			Start-Sleep -Milliseconds 250

			# Clean out unused runspace jobs
			$Runspaces.clone() | Where-Object { ($_.Runspace -eq $Null) } | ForEach {
				$Runspaces.remove($_)
				$ScanCount++
				Write-Progress -Activity 'Testing WMI connectivity' -PercentComplete (($ScanCount / $DeviceCount)*100) -Status "$ScanCount of $DeviceCount" -Id $WmiProgressId -ParentId $ParentProgressId
			}

		} while (($Runspaces | Where-Object {$_.Runspace -ne $Null} | Measure-Object).Count -gt 0)

		#endregion


		# Finally, close the runspaces
		$RunspacePool.close()

		Write-NetworkScanLog -Message 'WMI connectivity test complete' -MessageLevel Verbose
		Write-Progress -Activity 'Testing PING connectivity' -PercentComplete 100 -Status 'Complete' -Id $PingProgressId -ParentId $ParentProgressId -Completed
		Write-Progress -Activity 'Testing WMI connectivity' -PercentComplete 100 -Status 'Complete' -Id $WmiProgressId -ParentId $ParentProgressId -Completed


		# Return results
		Write-Output $Device.Values

		Write-NetworkScanLog -Message 'Network scan complete' -MessageLevel Information
		Write-NetworkScanLog -Message "`t-IP Addresses Scanned: $($($Device.Values | Measure-Object).Count)" -MessageLevel Information
		Write-NetworkScanLog -Message "`t-PING Replies: $($($Device.Values | Where-Object { $_.IsPingAlive -eq $true } | Measure-Object).Count)" -MessageLevel Information
		Write-NetworkScanLog -Message "`t-WMI Replies: $($($Device.Values | Where-Object { $_.IsWmiAlive -eq $true } | Measure-Object).Count)" -MessageLevel Information

		Write-NetworkScanLog -Message 'End Function: Find-IPv4Device' -MessageLevel Debug

		Remove-Variable -Name Device

	}

}

function Find-SqlServerService {
	<#
	.SYNOPSIS
		Synchronously look for SQL Server Services on a network.

	.DESCRIPTION
		This function executes a synchronus scan for hosts with SQL Server Services on a network.

	.PARAMETER  DnsServer
		'Automatic', or the Name or IP address of an Active Directory DNS server to query for a list of hosts to test for connectivity

		When 'Automatic' is specified the function will use WMI queries to discover the current computer's DNS server(s) to query.
		
	.PARAMETER  DnsDomain
		'Automatic' or the Active Directory domain name to use when querying DNS for a list of hosts.
		
		When 'Automatic' is specified the function will use the current computer's AD domain.
		
		'Automatic' will be used by default if DnsServer is specified but DnsDomain is not provided.

	.PARAMETER  Subnet
		'Automatic' or a comma delimited list of subnets (in CIDR notation) to scan for connectivity.
		
		When 'Automatic' is specified the function will use the current computer's IP configuration to determine subnets to scan. 
		
		A quick refresher on CIDR notation:

			BITS	SUBNET MASK			USABLE HOSTS PER SUBNET
			----	---------------		-----------------------
			/20		255.255.240.0		4094
			/21		255.255.248.0		2046 
			/22		255.255.252.0		1022
			/23		255.255.254.0		510 
			/24		255.255.255.0		254 
			/25		255.255.255.128		126 
			/26		255.255.255.192		62
			/27		255.255.255.224		30
			/28		255.255.255.240		14
			/29		255.255.255.248		6
			/30		255.255.255.252		2
			/32		255.255.255.255		1		

	.PARAMETER  ComputerName
		A comma delimited list of of computer names to test for SQL Server services.

	.PARAMETER  ExcludeSubnet
		A comma delimited list of subnets (in CIDR notation) to exclude when testing for connectivity.
		
	.PARAMETER  LimitSubnet
		A comma delimited list of subnets (in CIDR notation) to limit the scope of connectivity tests. Only hosts with IP Addresses that fall within the specified subnet(s) will be included in the results.

	.PARAMETER  ExcludeComputerName
		A comma delimited list of computer names to exclude when testing for connectivity. Wildcards are accepted.
		
		An attempt will be made to resolve the IP Address(es) for each computer in this list and those addresses will also be used when determining if a host should be included or excluded when testing for connectivity.		

	.PARAMETER  MaxConcurrencyThrottle
		Number between 1-100 to indicate how many instances to collect information from concurrently.

		If not provided then the number of logical CPUs present to your session will be used.

	.PARAMETER  PrivateOnly
		Only include hosts with private class A, B, or C IP addresses
		
	.PARAMETER  ParentProgressId
		If the caller is using Write-Progress then all progress information will be written using ParentProgressId as the ParentID		

	.EXAMPLE
		Find-SqlServerService -DNSServer automatic -DNSDomain automatic -PrivateOnly
		
		Description
		-----------
		Queries Active Directory for a list of hosts to scan for SQL Server services. The list of hosts will be restricted to private IP addresses only.

	.EXAMPLE
		Find-SqlServerService -Subnet 172.20.40.0/28
		
		Description
		-----------
		Scans all hosts in the subnet 172.20.40.0/28 for SQL Server services.
		
	.EXAMPLE
		Find-SqlServerService -Computername Server1,Server2,Server3
		
		Description
		-----------
		Scanning Server1, Server2, and Server3 for SQL Server services.

	.OUTPUTS
		System.Management.Automation.PSObject

	.NOTES

#>
	[CmdletBinding(DefaultParametersetName='dns')]
	param(
		[Parameter(
			Mandatory=$true,
			ParameterSetName='dns',
			HelpMessage='DNS Server(s)'
		)] 
		[alias('dns')]
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$|^auto$|^automatic$')]
		[string[]]
		$DnsServer = 'automatic'
		,
		[Parameter(
			Mandatory=$false,
			ParameterSetName='dns',
			HelpMessage='DNS Domain Name'
		)] 
		[alias('domain')]
		[string]
		$DnsDomain = 'automatic'
		,
		[Parameter(
			Mandatory=$true,
			ParameterSetName='subnet',
			HelpMessage='Subnet (in CIDR notation)'
		)] 
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)[\\/]\d{1,2}$|^auto$|^automatic$')]
		[string[]]
		$Subnet = 'Automatic'
		,
		[Parameter(
			Mandatory=$true,
			ParameterSetName='computername',
			HelpMessage='Computer Name(s)'
		)] 
		[alias('Computer')]
		[string[]]
		$ComputerName
		,
		[Parameter(Mandatory=$false, ParameterSetName='dns')]
		[Parameter(Mandatory=$false, ParameterSetName='subnet')]
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)[\\/]\d{1,2}$')]
		[string[]]
		$ExcludeSubnet
		,
		[Parameter(Mandatory=$false, ParameterSetName='dns')]
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)[\\/]\d{1,2}$')]
		[string[]]
		$LimitSubnet
		,
		[Parameter(Mandatory=$false, ParameterSetName='dns')]
		[Parameter(Mandatory=$false, ParameterSetName='subnet')]
		[string[]]
		$ExcludeComputerName
		,
		[Parameter(Mandatory=$false)] 
		[ValidateRange(1,100)]
		[alias('Throttle')]
		[byte]
		$MaxConcurrencyThrottle = $env:NUMBER_OF_PROCESSORS
		,
		[Parameter(Mandatory=$false)] 
		[switch]
		$PrivateOnly = $false
		,
		[Parameter(Mandatory=$false)]
		[ValidateNotNull()]
		[Int32]
		$ParentProgressId = -1
	)
	process {

		$Service = New-Object -TypeName psobject -Property @{ Service = @() }
		$IPv4Device = @()
		$ParameterHash = $null
		$ScanCount = 0
		$WmiDeviceCount = 0

		$SqlScanProgressId = Get-Random

		# For use with runspaces
		$ScriptBlock = $null
		$SessionState = $null
		$RunspacePool = $null
		$Runspaces = $null
		$Runspace = $null
		$PowerShell = $null
		$HashKey = $null

		# Fallback in case value isn't supplied or somehow missing from the environment variables
		if (-not $MaxConcurrencyThrottle) { $MaxConcurrencyThrottle = 1 }

		Write-NetworkScanLog -Message 'Start Function: Find-SqlServerService' -MessageLevel Debug

		# Build command for splatting
		$ParameterHash = @{
			MaxConcurrencyThrottle = $MaxConcurrencyThrottle
			PrivateOnly = $PrivateOnly
			ResolveAliases = $true
			ParentProgressId = $ParentProgressId
		}

		switch ($PsCmdlet.ParameterSetName) {
			'dns' {
				$ParameterHash.Add('DnsServer',$DnsServer)
				$ParameterHash.Add('DnsDomain',$DnsDomain)
				if ($ExcludeSubnet) { $ParameterHash.Add('ExcludeSubnet',$ExcludeSubnet) }
				if ($LimitSubnet) { $ParameterHash.Add('IncludeSubnet',$LimitSubnet) }
				if ($ExcludeComputerName) { $ParameterHash.Add('ExcludeComputerName',$ExcludeComputerName) }
			}
			'subnet' {
				$ParameterHash.Add('Subnet',$Subnet)
				if ($ExcludeSubnet) { $ParameterHash.Add('ExcludeSubnet',$ExcludeSubnet) }
				if ($ExcludeComputerName) { $ParameterHash.Add('ExcludeComputerName',$ExcludeComputerName) }
			}
			'computername' {
				$ParameterHash.Add('ComputerName',$ComputerName)
			}

		}

		# Scan the network to find WMI capable (i.e. Windows) devices
		$IPv4Device = (Find-IPv4Device @ParameterHash)
		$WmiDeviceCount = $($IPv4Device | Where-Object { $_.IsWmiAlive -eq $true } | Group-Object -Property DnsRecordName | Measure-Object).Count

		Write-NetworkScanLog -Message 'Beginning SQL Service discovery scan' -MessageLevel Information
		Write-Progress -Activity 'Scanning for SQL Services' -PercentComplete 0 -Status "Scanning $WmiDeviceCount Devices" -Id $SqlScanProgressId -ParentId $ParentProgressId


		# Create a Session State, Create a RunspacePool, and open the RunspacePool
		$SessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
		$RunspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $MaxConcurrencyThrottle, $SessionState, $Host)
		$RunspacePool.Open()

		# Create an empty collection to hold the Runspace jobs
		$Runspaces = New-Object System.Collections.ArrayList 


		$ScanCount = 0
		$ScriptBlock = {
			Param (
				[String]$IpAddress,
				[String]$ComputerName
			)

			# Load SMO Assemblies
			[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | ForEach-Object {
				if ($_.GetName().Version.Major -ge 10) {
					[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMOExtended') | Out-Null
					[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SQLWMIManagement') | Out-Null
				}
			}

			# Registry constants
			New-Variable -Name HKEY_LOCAL_MACHINE -Value 2147483650 -Scope Script -Option Constant

			# Variables
			$InstanceName = $null
			$IsNamedInstance = $false
			$IsClusteredInstance = $false
			$ClusterName = $null
			$ServiceIpAddress = $null
			$DomainName = $null
			$ServiceTypeName = $null
			$ServiceName = $null
			$Server = $null
			$Port = $null
			$IsDynamicPort = $false
			$ServiceInstallDate = $null
			$ServiceStartDate = $null

			$StdRegProv = $null
			$RegistryKeyRootPath = $null
			$PathName = $null
			$StartMode = $null
			$ProcessId = $null
			$ServiceState = $null
			$ServiceAccount = $null
			$Description = $null
			$RegistryKeyPath = $null
			$Parameters = $null
			$StartupParameters = $null


			$ManagedComputer = New-Object -TypeName 'Microsoft.SqlServer.Management.Smo.Wmi.ManagedComputer' -ArgumentList $ComputerName
			#$ManagedServiceTypeEnum = 'Microsoft.SqlServer.Management.Smo.Wmi.ManagedServiceType' -as [Type]

			# Try to use the WMI Managed Computer Object to find SQL Services on $ComputerName
			try {
				$ManagedComputer.Services | ForEach-Object {

					if (($_.Name).IndexOf('$') -gt 0) {
						$InstanceName = ($_.Name).Substring(($_.Name).IndexOf('$') + 1)
						$IsNamedInstance = $true
					} else {
						$InstanceName = $null
						$IsNamedInstance = $false
					}

					# Try and determine if this is a clustered server (and the FQDN & IP for the cluster if it is)
					try {
						if ($_.AdvancedProperties['CLUSTERED'].Value -eq $true) {
							$IsClusteredInstance = $true

							Get-WmiObject -Namespace root\CIMV2 -Class Win32_ComputerSystem -Property Domain -ComputerName $IpAddress -ErrorAction Stop | ForEach-Object {
								$DomainName = $_.Domain
							}

							$ClusterName = [String]::Join('.', @($_.AdvancedProperties['VSNAME'].Value, $DomainName)).ToUpper()

							[System.Net.Dns]::GetHostByName($ClusterName) | ForEach-Object {
								$_.AddressList | Where-Object { $_.AddressFamily -ieq 'InterNetwork' } | ForEach-Object {
									$ServiceIpAddress = $_.IPAddressToString
								}
							}

						} else {
							$IsClusteredInstance = $false
							$ClusterName = [String]::Empty
							$ServiceIpAddress = $IpAddress
						}
					}
					catch {
						$IsClusteredInstance = $false
						$ClusterName = [String]::Empty
						$ServiceIpAddress = $IpAddress
					}


					# Get the Friendly name for the service
					$ServiceTypeName = switch ($_.Type.value__) {
						5 { 'SQL Server Analysis Services' }
						#$ManagedServiceTypeEnum::NotificationServer { 'SQL Server Notification Services' }
						6 { 'SQL Server Reporting Services' }
						#$ManagedServiceTypeEnum::Search { 'Microsoft Search service' }
						3 { 'SQL Server FullText Search' }
						2 { 'SQL Server Agent' }
						7 { 'SQL Server Browser' }
						1 { 'SQL Server' }
						4 { 'SQL Server Integration Services' }
						9 { 'SQL Full-text Filter Daemon Launcher' }
						$null { 'Unknown' }
						default { $_.Type.ToString() }
					}


					# Get the port number for SQL Server Services
					#if ($_.Type -eq $ManagedServiceTypeEnum::SqlServer) {
					if ($ServiceTypeName -ieq 'SQL Server') {
						$ServiceName = $_.Name
						$ManagedComputer.ServerInstances | Where-Object { $_.Name -ieq $ServiceName } | ForEach-Object {

							$_.ServerProtocols | Where-Object { $_.Name -ieq 'tcp' } | ForEach-Object {

								# First get the port for all IP Addresses
								$_.IPAddresses | Where-Object { $_.Name -ieq 'ipall' } | ForEach-Object {
									$Port = $_.IPAddressProperties['TcpPort'].Value
									if (-not $Port) {
										$Port = $_.IPAddressProperties['TcpDynamicPorts'].Value
										$IsDynamicPort = $true
									} else {
										$IsDynamicPort = $false
									}
								}

								# Then check to see if the port is overridden for the specific IP Address provided
								$_.IPAddresses | Where-Object { ($_.IPAddress -ieq $ServiceIpAddress) -and ($_.IPAddressProperties['Active'].Value -eq $true) -and ($_.IPAddressProperties['Enabled'].Value -eq $true) } | ForEach-Object {
									$Port = $_.IPAddressProperties['TcpPort'].Value
									if (-not $Port) {
										$Port = $_.IPAddressProperties['TcpDynamicPorts'].Value
										$IsDynamicPort = $true
									} else {
										$IsDynamicPort = $false
									}
								}
							}

						}
					} else {
						$Port = $null
						$IsDynamicPort = $null
					}


					# Get the Service Start Date (if it's got a Process ID greater than than 0)
					if ($_.ProcessId -gt 0) {
						try {
							Get-WmiObject -Namespace root\CIMV2 -Class Win32_Process -Filter "ProcessId = '$($_.ProcessId)'" -Property CreationDate -ComputerName $IpAddress -ErrorAction Stop | ForEach-Object {
								$ServiceStartDate = $_.ConvertToDateTime($_.CreationDate)
							}
						}
						catch {
							$ServiceStartDate = $null
						}
					} else {
						$ServiceStartDate = $null
					}

					# 					# Get the Service Install Date
					# 					Get-WmiObject -Namespace root\CIMV2 -Class Win32_Service -Filter "DisplayName = '$($_.DisplayName)'" -Property CreationDate -ComputerName $IpAddress | ForEach-Object {
					# 						$ServiceStartDate = $_.CreationDate
					# 					}


					Write-Output (
						New-Object -TypeName psobject -Property @{
							ComputerName = $ComputerName
							#ClusterName = $ClusterName
							DisplayName = $_.DisplayName
							Description = $_.Description
							ComputerIpAddress = $IpAddress
							#InstanceName = $InstanceName
							IsNamedInstance = $IsNamedInstance
							IsClusteredInstance = $IsClusteredInstance
							IsDynamicPort = $IsDynamicPort
							IsHadrEnabled = $_.IsHadrEnabled
							PathName = $_.PathName
							Port = $Port
							ProcessId = $_.ProcessId
							ServerName = $(
								if ($IsNamedInstance -eq $true) {
									if ($IsClusteredInstance -eq $true) {
										[String]::Join('\', @($ClusterName, $InstanceName))
									} else {
										[String]::Join('\', @($ComputerName, $InstanceName))
									}
								} else {
									if ($IsClusteredInstance -eq $true) {
										$ClusterName
									} else {
										$ComputerName
									}
								}
							)
							ServiceIpAddress = $ServiceIpAddress
							ServiceStartDate = $ServiceStartDate
							ServiceState = $_.ServiceState.ToString()
							#ServiceType = $_.Type.ToString()
							ServiceTypeName = $ServiceTypeName
							ServiceAccount = $_.ServiceAccount
							StartMode = $_.StartMode.ToString()
							StartupParameters = $_.StartupParameters
						}
					)
				} 
			}
			catch {

				# If we get this error it's possible that this is a SQL 2000 server in which case we need to look at the registry to find installed services
				if ($_.Exception.Message -ilike '*SQL Server WMI provider is not available*') {
					try {

						# Use the WMI Registry provider to access the registry on $ComputerName
						# For info on using this WMI class see http://msdn.microsoft.com/en-us/library/windows/desktop/aa393664(v=vs.85).aspx
						$StdRegProv = Get-WmiObject -Namespace root\DEFAULT -Query "select * FROM meta_class WHERE __Class = 'StdRegProv'" -ComputerName $ComputerName -ErrorAction Stop

						# Iterate through installed instances of the Database Engine (which includes the SQL Agent)
						$($StdRegProv.GetMultiStringValue($HKEY_LOCAL_MACHINE,'SOFTWARE\Microsoft\Microsoft SQL Server','InstalledInstances')).sValue | ForEach-Object {

							if ($_ -ine 'MSSQLSERVER') {
								$InstanceName = $_
								$DisplayName = "MSSQL`$$InstanceName"
								$IsNamedInstance = $true
								$RegistryKeyRootPath = "SOFTWARE\Microsoft\Microsoft SQL Server\$($_)"
							} else {
								$InstanceName = $null
								$DisplayName = 'MSSQLSERVER'
								$IsNamedInstance = $false
								$RegistryKeyRootPath = 'SOFTWARE\Microsoft\MSSQLServer'
							}

							# Determine if instance is clustered
							$IsClusteredInstance = $null
							$ClusterName = [String]::Empty
							$ServiceIpAddress = $IpAddress

							# Get the TCP port number for SQL Server Services
							$Port = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyRootPath\MSSQLServer\SuperSocketNetLib\Tcp",'TcpDynamicPorts')).sValue
							if (-not $Port) {
								$Port = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyRootPath\MSSQLServer\SuperSocketNetLib\Tcp",'TcpPort')).sValue
								$IsDynamicPort = $false
							} else {
								$IsDynamicPort = $true
							}

							# Get Service Information
							Get-WmiObject -Namespace root\CIMV2 -Class Win32_Service `
							-Filter "DisplayName = '$DisplayName'" `
							-Property PathName,StartMode,ProcessId,State,StartName,Description -ComputerName $IpAddress -ErrorAction Stop | 
							ForEach-Object {
								$PathName = $_.PathName
								$StartMode = $_.StartMode
								$ProcessId = $_.ProcessId
								$ServiceState = $_.State
								$ServiceAccount = $_.StartName
								$Description = $_.Description
							}

							# Get the Service Start Date (if it's got a Process ID greater than than 0)
							if ($ProcessId -gt 0) {
								try {
									Get-WmiObject -Namespace root\CIMV2 -Class Win32_Process -Filter "ProcessId = '$ProcessId'" -Property CreationDate -ComputerName $IpAddress -ErrorAction Stop | ForEach-Object {
										$ServiceStartDate = $_.ConvertToDateTime($_.CreationDate)
									}
								}
								catch {
									$ServiceStartDate = $null
								}
							} else {
								$ServiceStartDate = $null
							}


							# Startup Parameters
							$RegistryKeyPath = "$RegistryKeyRootPath\MSSQLServer\Parameters"
							$Parameters = $StdRegProv.EnumValues($HKEY_LOCAL_MACHINE,$RegistryKeyPath)
							$StartupParameters = @()

							for ($i = 0; $i -lt ($Parameters.sNames | Measure-Object).Count; $i++) {
								switch ($Parameters.Types[$i]) {
									1 {
										# REG_SZ
										$StartupParameters += $($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,$RegistryKeyPath,"$($Parameters.sNames[$i])")).sValue
									}
									2 { 
										# REG_EXPAND_SZ
										$StartupParameters += $($StdRegProv.GetExpandedStringValue($HKEY_LOCAL_MACHINE,$RegistryKeyPath,"$($Parameters.sNames[$i])")).sValue
									}
									3 {
										# REG_BINARY
										$StartupParameters += [System.BitConverter]::ToString($($StdRegProv.GetBinaryValue($HKEY_LOCAL_MACHINE,$RegistryKeyPath,"$($Parameters.sNames[$i])").uValue) )
									}
									4 {
										# REG_DWORD
										$StartupParameters += $($StdRegProv.GetDWORDValue($HKEY_LOCAL_MACHINE,$RegistryKeyPath, "$($Parameters.sNames[$i])")).uValue
									}
									7 {
										# REG_MULTI_SZ
										$($StdRegProv.GetMultiStringValue($HKEY_LOCAL_MACHINE,$RegistryKeyPath,"$($Parameters.sNames[$i])")).sValue | ForEach-Object {
											$StartupParameters += $_
										} 
									}
									default { $null }
								}
							} 

							Write-Output (
								New-Object -TypeName psobject -Property @{
									ComputerName = $ComputerName
									DisplayName = $DisplayName
									Description = $Description
									ComputerIpAddress = $IpAddress
									IsNamedInstance = $IsNamedInstance
									IsClusteredInstance = $IsClusteredInstance
									IsDynamicPort = $IsDynamicPort
									IsHadrEnabled = $null # Not applicable to SQL 2000
									PathName = $PathName
									Port = $Port
									ProcessId = $ProcessId
									ServerName = $(
										if ($IsNamedInstance -eq $true) {
											if ($IsClusteredInstance -eq $true) {
												[String]::Join('\', @($ClusterName, $InstanceName))
											} else {
												[String]::Join('\', @($ComputerName, $InstanceName))
											}
										} else {
											if ($IsClusteredInstance -eq $true) {
												$ClusterName
											} else {
												$ComputerName
											}
										}
									)
									ServiceIpAddress = $ServiceIpAddress
									ServiceStartDate = $ServiceStartDate
									ServiceState = $ServiceState
									ServiceTypeName = 'SQL Server'
									ServiceAccount = $ServiceAccount
									StartMode = $StartMode
									StartupParameters = [String]::Join(';', $StartupParameters)
								}
							)



							# Now let's tackle the SQL Agent. A lot of the information is the same as the SQL Server Service
							if ($IsNamedInstance) {
								$DisplayName = "SQLAgent`$$InstanceName"
							} else {
								$DisplayName = 'SQLSERVERAGENT'
							}

							# Get Service Information
							Get-WmiObject -Namespace root\CIMV2 -Class Win32_Service `
							-Filter "DisplayName = '$DisplayName'" `
							-Property PathName,StartMode,ProcessId,State,StartName,Description -ComputerName $IpAddress -ErrorAction Stop | 
							ForEach-Object {
								$PathName = $_.PathName
								$StartMode = $_.StartMode
								$ProcessId = $_.ProcessId
								$ServiceState = $_.State
								$ServiceAccount = $_.StartName
								$Description = $_.Description
							}

							# Get the Service Start Date (if it's got a Process ID greater than than 0)
							if ($ProcessId -gt 0) {
								try {
									Get-WmiObject -Namespace root\CIMV2 -Class Win32_Process -Filter "ProcessId = '$ProcessId'" -Property CreationDate -ComputerName $IpAddress -ErrorAction Stop | ForEach-Object {
										$ServiceStartDate = $_.ConvertToDateTime($_.CreationDate)
									}
								}
								catch {
									$ServiceStartDate = $null
								}
							} else {
								$ServiceStartDate = $null
							}

							Write-Output (
								New-Object -TypeName psobject -Property @{
									ComputerName = $ComputerName
									DisplayName = $DisplayName
									Description = $Description
									ComputerIpAddress = $IpAddress
									IsNamedInstance = $IsNamedInstance
									IsClusteredInstance = $IsClusteredInstance
									IsDynamicPort = $IsDynamicPort
									IsHadrEnabled = $null # Not applicable to SQL 2000
									PathName = $PathName
									Port = $null
									ProcessId = $ProcessId
									ServerName = $(
										if ($IsNamedInstance -eq $true) {
											if ($IsClusteredInstance -eq $true) {
												[String]::Join('\', @($ClusterName, $InstanceName))
											} else {
												[String]::Join('\', @($ComputerName, $InstanceName))
											}
										} else {
											if ($IsClusteredInstance -eq $true) {
												$ClusterName
											} else {
												$ComputerName
											}
										}
									)
									ServiceIpAddress = $ServiceIpAddress
									ServiceStartDate = $ServiceStartDate
									ServiceState = $ServiceState
									ServiceTypeName = 'SQL Server Agent'
									ServiceAccount = $ServiceAccount
									StartMode = $StartMode
									StartupParameters = $null
								}
							)
						}

						# Now let's test for Analysis Services, Reporting Services, and Microsoft Search. 
						# You can't have more than one instance of each in 2000
						Get-WmiObject -Namespace root\CIMV2 -Class Win32_Service `
						-Filter "(DisplayName = 'MSSQLServerOLAPService') or (DisplayName = 'Microsoft Search') or (DisplayName = 'ReportServer')" `
						-Property DisplayName,PathName,StartMode,ProcessId,State,StartName,Description -ComputerName $IpAddress -ErrorAction Stop | 
						ForEach-Object {

							$DisplayName = $_.DisplayName
							$PathName = $_.PathName
							$StartMode = $_.StartMode
							$ProcessId = $_.ProcessId
							$ServiceState = $_.State
							$ServiceAccount = $_.StartName
							$Description = $_.Description

							# Get the Service Start Date (if it's got a Process ID greater than than 0)
							if ($ProcessId -gt 0) {
								try {
									Get-WmiObject -Namespace root\CIMV2 -Class Win32_Process -Filter "ProcessId = '$ProcessId'" -Property CreationDate -ComputerName $IpAddress -ErrorAction Stop | ForEach-Object {
										$ServiceStartDate = $_.ConvertToDateTime($_.CreationDate)
									}
								}
								catch {
									$ServiceStartDate = $null
								}
							} else {
								$ServiceStartDate = $null
							}

							Write-Output (
								New-Object -TypeName psobject -Property @{
									ComputerName = $ComputerName
									DisplayName = $DisplayName
									Description = $Description
									ComputerIpAddress = $IpAddress
									IsNamedInstance = $false # Can't have named instances of SSAS, SSRS, or Microsoft Search in SQL 2000
									IsClusteredInstance = $false # Can't cluster SSAS, SSRS, or Microsoft Search in SQL 2000
									IsDynamicPort = $null
									IsHadrEnabled = $null # Not applicable to SQL 2000
									PathName = $PathName
									Port = $null
									ProcessId = $ProcessId
									ServerName = $ComputerName
									ServiceIpAddress = $IpAddress
									ServiceStartDate = $ServiceStartDate
									ServiceState = $ServiceState
									ServiceTypeName = switch ($DisplayName) {
										'Microsoft Search' { 'Microsoft Search service' } 
										'MSSQLServerOLAPService' { 'SQL Server Analysis Services' }
										'ReportServer' { 'SQL Server Reporting Services' }
										default { 'Unknown' }
									}
									ServiceAccount = $ServiceAccount
									StartMode = $StartMode
									StartupParameters = $null
								}
							)
						} 

					}
					catch {
						throw
					}
				}
				else {
					# Something else has happened; Let the error bubble up
					throw
				}
			}
		}


		# Iterate through each machine that we could make a WMI connection to and gather information
		# Some machines may have multiple entries (b\c of multiple IP Addresses) so only use the first IP Address for each
		$IPv4Device | Where-Object { $_.IsWmiAlive -eq $true } | Group-Object -Property DnsRecordName | ForEach-Object {

			$ScanCount++
			Write-NetworkScanLog -Message "Scanning $(($_.Group[0]).DnsRecordName) at IP address $(($_.Group[0]).IPAddress) for SQL Services [Device $ScanCount of $WmiDeviceCount]" -MessageLevel Information

			#Create the PowerShell instance and supply the scriptblock with the other parameters
			$PowerShell = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
			$PowerShell = $PowerShell.AddArgument($($_.Group[0]).IPAddress)
			$PowerShell = $PowerShell.AddArgument($($_.Group[0]).DnsRecordName)

			#Add the runspace into the PowerShell instance
			$PowerShell.RunspacePool = $RunspacePool

			$Runspaces.Add((
					New-Object -TypeName PsObject -Property @{
						PowerShell = $PowerShell
						Runspace = $PowerShell.BeginInvoke()
						ComputerName = $($_.Group[0]).DnsRecordName
						IPAddress = $($_.Group[0]).IPAddress
					}
				)) | Out-Null
		}

		# Reset the scan counter
		$ScanCount = 0

		# Process results as they complete
		Do {
			foreach ($Runspace in $Runspaces) {

				If ($Runspace.Runspace.IsCompleted) {
					try {

						# This is where the output gets returned
						$Runspace.PowerShell.EndInvoke($Runspace.Runspace) | ForEach-Object {
							$Service.Service += $_

							if ($_.IsNamedInstance -eq $true) {
								Write-NetworkScanLog -Message "Found $($_.ServiceTypeName) named instance $($_.ServerName) at IP address $($_.ServiceIpAddress)" -MessageLevel Information
							} else {
								Write-NetworkScanLog -Message "Found $($_.ServiceTypeName) default instance $($_.ServerName) at IP address $($_.ServiceIpAddress)" -MessageLevel Information
							}
						}

					}
					catch {
						#if ($_.Exception.Message -ilike '*SQL Server WMI provider is not available*') {
						#	Write-NetworkScanLog -Message "ERROR: Unable to retrieve service information from $($Runspace.ComputerName) ($($Runspace.IPAddress)). The SQL Server WMI provider may need to be installed on $($Runspace.ComputerName)." -MessageLevel Information
						#} else {
						Write-NetworkScanLog -Message "ERROR: Unable to retrieve service information from $($Runspace.ComputerName) ($($Runspace.IPAddress)): $($_.Exception.Message)" -MessageLevel Information
						#} 
					}
					finally {
						# Cleanup
						$Runspace.PowerShell.dispose()
						$Runspace.Runspace = $null
						$Runspace.PowerShell = $null
					}
				}
			}

			# Found that in some cases an ACCESS_VIOLATION error occurs if we don't delay a little bit during each iteration 
			Start-Sleep -Milliseconds 250

			# Clean out unused runspace jobs
			$Runspaces.clone() | Where-Object { ($_.Runspace -eq $Null) } | ForEach {
				$Runspaces.remove($_)
				$ScanCount++
				Write-Progress -Activity 'Scanning for SQL Services' -PercentComplete (($ScanCount / $WmiDeviceCount)*100) -Status "Device $ScanCount of $WmiDeviceCount" -Id $SqlScanProgressId -ParentId $ParentProgressId
			}


		} while (($Runspaces | Where-Object {$_.Runspace -ne $Null} | Measure-Object).Count -gt 0)
		#endregion

		# Finally, close the runspaces
		$RunspacePool.close()

		Write-Progress -Activity 'Scanning for SQL Services' -PercentComplete 100 -Status 'Complete' -Id $SqlScanProgressId -ParentId $ParentProgressId -Completed

		Write-NetworkScanLog -Message 'SQL Server service discovery complete' -MessageLevel Information

		$Service.Service | Select-Object -Property ServiceTypeName -Unique | Sort-Object -Property ServiceTypeName | ForEach-Object {
			$ServiceTypeName = $_.ServiceTypeName
			Write-NetworkScanLog -Message "`t-$($ServiceTypeName) Instance Count: $(($Service.Service | Where-Object { ($_.ServiceTypeName -ieq $ServiceTypeName) } | Measure-Object).Count)" -MessageLevel Information
		}
		Write-NetworkScanLog -Message 'End Function: Find-SqlServerService' -MessageLevel Debug

		# Write output
		Write-Output $Service.Service

	}

}