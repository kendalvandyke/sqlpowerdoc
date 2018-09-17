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
				ValueFromPipeline=$true,
				ValueFromPipelineByPropertyName=$true,
		HelpMessage='What computer name would you like to target?')]
		[Alias('host')]
		[ValidateLength(3,30)]
		[string[]]$ComputerName = 'localhost'
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
							Domain       = $Domain
							DnsServer    = $_.DnsServerSearchOrder
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
				ValueFromPipeline=$true,
				ValueFromPipelineByPropertyName=$true,
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
							IPAddress  = $_.IPAddress[$pos]
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
					Get-WmiObject -Namespace root\MicrosoftDNS -Class MicrosoftDNS_AType -ComputerName $_ -Filter "ContainerName= '$Domain'" -ErrorAction Stop | 
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
				DnsRecordName  = $DnsRecordName
				WmiMachineName = $WmiMachineName
				IPAddress      = $IPAddress
				IsPingAlive    = $IsPingAlive
				IsWmiAlive     = $IsWmiAlive
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

			BITS    SUBNET MASK        USABLE HOSTS PER SUBNET
			----    ---------------    -----------------------
			/20	    255.255.240.0      4094
			/21	    255.255.248.0      2046
			/22	    255.255.252.0      1022
			/23	    255.255.254.0      510
			/24	    255.255.255.0      254
			/25	    255.255.255.128    126
			/26	    255.255.255.192    62
			/27	    255.255.255.224    30
			/28	    255.255.255.240    14
			/29	    255.255.255.248    6
			/30	    255.255.255.252    2
			/32	    255.255.255.255    1

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

			.PARAMETER  TimeoutSeconds
			Number of seconds to wait for WMI connectivity test to return before timing out.

			If not provided then 30 seconds is used as the default.

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
	[CmdletBinding()]
	param(
		[Parameter(
				Mandatory=$true,
				ParameterSetName='dns',
				ValueFromPipelineByPropertyName=$true,
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
				ValueFromPipelineByPropertyName=$true,
				HelpMessage='DNS Domain Name'
		)] 
		[alias('domain')]
		[string]
		$DnsDomain = 'automatic'
		,
		[Parameter(
				Mandatory=$true,
				ParameterSetName='subnet',
				ValueFromPipelineByPropertyName=$true,
				HelpMessage='Subnet (in CIDR notation)'
		)] 
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)[\\/]\d{1,2}$|^auto$|^automatic$')]
		[string[]]
		$Subnet = 'Automatic'
		,
		[Parameter(
				Mandatory=$true,
				ParameterSetName='computername',
				ValueFromPipelineByPropertyName=$true,
				HelpMessage='Computer Name(s)'
		)] 
		[alias('Computer')]
		[string[]]
		$ComputerName
		,
		[Parameter(
				Mandatory=$false, 
				ParameterSetName='dns',
				ValueFromPipelineByPropertyName=$true
		)]
		[Parameter(
				Mandatory=$false, 
				ParameterSetName='subnet',
				ValueFromPipelineByPropertyName=$true
		)]
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)[\\/]\d{1,2}$')]
		[string[]]
		$ExcludeSubnet
		,
		[Parameter(
				Mandatory=$false, 
				ParameterSetName='dns',
				ValueFromPipelineByPropertyName=$true
		)]
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)[\\/]\d{1,2}$')]
		[string[]]
		$LimitSubnet
		,
		[Parameter(
				Mandatory=$false, 
				ParameterSetName='dns',
				ValueFromPipelineByPropertyName=$true
		)]
		[Parameter(
				Mandatory=$false, 
				ParameterSetName='subnet',
				ValueFromPipelineByPropertyName=$true
		)]
		[string[]]
		$ExcludeComputerName
		,
		[Parameter(
				Mandatory=$false, 
				ParameterSetName='dns',
				ValueFromPipelineByPropertyName=$true
		)]
		[Parameter(
				Mandatory=$false, 
				ParameterSetName='subnet',
				ValueFromPipelineByPropertyName=$true
		)]
		[string[]]
		$IncludeComputerName
		,
		[Parameter(
				Mandatory=$false,
				ValueFromPipelineByPropertyName=$true
		)] 
		[ValidateRange(1,100)]
		[alias('Throttle')]
		[byte]
		$MaxConcurrencyThrottle = $env:NUMBER_OF_PROCESSORS
		,
		[Parameter(
				Mandatory=$false,
				ValueFromPipelineByPropertyName=$true
		)] 
		[switch]
		$PrivateOnly = $false
		,
		[Parameter(
				Mandatory=$false,
				ValueFromPipelineByPropertyName=$true
		)] 
		[switch]
		$ResolveAliases = $false
		,
		[Parameter(
				Mandatory=$false,
				ValueFromPipelineByPropertyName=$true
		)] 
		[switch]
		$SkipConnectionTest = $false
		,
		[Parameter(
				Mandatory=$false,
				ValueFromPipelineByPropertyName=$true
		)]
		[ValidateNotNull()]
		[Int32]
		$ParentProgressId = -1
		,
		[Parameter(
				Mandatory=$false,
				ValueFromPipelineByPropertyName=$true
		)] 
		[ValidateRange(1,32767)]
		[alias('Timeout')]
		[Int16]
		$TimeoutSeconds = 30
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

		Write-NetworkScanLog -Message "`t-MaxConcurrencyThrottle: $MaxConcurrencyThrottle" -MessageLevel Information
		Write-NetworkScanLog -Message "`t-PrivateOnly: $PrivateOnly" -MessageLevel Information
		Write-NetworkScanLog -Message "`t-ResolveAliases: $ResolveAliases" -MessageLevel Information
		Write-NetworkScanLog -Message "`t-SkipConnectionTest: $SkipConnectionTest" -MessageLevel Information


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
				Where-Object { $_.IndexOfAny(@('/', '\', '[', ']', '"', ':', ';', '|', '<', '>', '+', '=', ',', '?', '*', ' ', '_')) -eq -1 } |
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
							Domain       = $DnsDomain
							DnsServer    = $DnsServer
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

							# Check if the host name is in the list of included computer names
							# Unlike exclusions, this is strictly based on simple name matching, i.e. no IP address matching
							if ($IncludeComputerName) {

								$HasMetCriteria = $false

								$IncludeComputerName | ForEach-Object {
									if ($HostName -ilike $_) {
										$HasMetCriteria = $true
										break
									}
								}

							}

							# Check if the host name is in the list of excluded computer names
							# Exclusions trump inclusions
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
								($LimitSubnet -or $ExcludeSubnet)
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
									IPAddress  = $_.NetworkAddress
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

							# Check if the host name is in the list of included computer names
							# Unlike exclusions, this is strictly based on simple name matching, i.e. no IP address matching
							if ($IncludeComputerName) {

								$HasMetCriteria = $false

								$IncludeComputerName | ForEach-Object {
									if ($HostName -ilike $_) {
										$HasMetCriteria = $true
										break
									}
								}

							}

							# Check if the host name is in the list of excluded computer names
							# Exclusions trump inclusions
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

				<# Write warnings for Azure SQL Databases #>
				$ComputerName |
				Where-Object { $_ -ilike '*.database.windows.net' } |
				Select-Object -Unique |
				ForEach-Object {
					Write-NetworkScanLog -Message "Excluding $_ from discovery; Windows Azure SQL Databases cannot be discovered by WMI" -MessageLevel warning
				}

				$ComputerName | 
				Where-Object { $_ -inotlike '*.database.windows.net' } |
				Select-Object -Unique | 
				ForEach-Object {

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

		if ($SkipConnectionTest -eq $true) {
			$Device |
			ForEach-Object {
				$_.IsPingAlive = $null
				$_.IsWmiAlive = $null
			}

		} else {


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

			#region PING CONNECTIVITY TEST

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

				$null = $Runspaces.Add((
						New-Object -TypeName PsObject -Property @{
							PowerShell = $PowerShell
							Runspace   = $PowerShell.BeginInvoke()
							HashKey    = $_.Key
						}
				))

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
				$Runspaces.clone() | Where-Object { ($_.Runspace -eq $null) } | ForEach-Object {
					$Runspaces.remove($_)
					$ScanCount++
					Write-Progress -Activity 'Testing PING connectivity' -PercentComplete (($ScanCount / $DeviceCount)*100) -Status "$ScanCount of $DeviceCount" -Id $PingProgressId -ParentId $ParentProgressId
				}

			} while (($Runspaces | Where-Object {$_.Runspace -ne $null} | Measure-Object).Count -gt 0)
		
			#endregion PING CONNECTIVITY TEST


			# Count how many devices responded
			$PingAliveDevice = @($Device.GetEnumerator() | Where-Object { $($_.Value).IsPingAlive -eq $true })
			$DeviceCount = $($PingAliveDevice | Measure-Object).Count

			Write-NetworkScanLog -Message 'PING connectivity test complete' -MessageLevel Verbose
			Write-Progress -Activity 'Testing PING connectivity' -PercentComplete 100 -Status "$ScanCount addresses tested, $DeviceCount replies" -Id $PingProgressId -ParentId $ParentProgressId
			Write-NetworkScanLog -Message "Testing WMI connectivity to $DeviceCount addresses" -MessageLevel Information
			Write-Progress -Activity 'Testing WMI connectivity' -PercentComplete 0 -Status "Testing $DeviceCount Addresses" -Id $WmiProgressId -ParentId $ParentProgressId


			#region WMI CONNECTIVITY TEST

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

				$Win32_ComputerSystem = Get-WmiObject -Namespace root\CIMV2 -Class Win32_ComputerSystem -Property Name -ComputerName $IPAddress -ErrorAction Stop
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

				$null = $Runspaces.Add((
						New-Object -TypeName PsObject -Property @{
							PowerShell = $PowerShell
							Runspace   = $PowerShell.BeginInvoke()
							HashKey    = $_.Key
							StartDate  = [DateTime]::Now
						}
				))

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
								!$Device[$HashKey].DnsRecordName -or
								(
									$ResolveAliases -eq $true -and 
									$Device[$HashKey].DnsRecordName.StartsWith($HostName, 'CurrentCultureIgnoreCase') -ne $true
								)
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
					elseif ($([DateTime]::Now).Subtract($_.StartDate).TotalSeconds -gt $TimeoutSeconds) {

						$HashKey = $_.HashKey
						$_.PowerShell.Stop()
						$_.PowerShell.dispose()
						$_.Runspace = $null
						$_.PowerShell = $null

						if ($Device[$HashKey].DnsRecordName) {
							Write-NetworkScanLog -Message "Timeout waiting for WMI response from $($Device[$HashKey].DnsRecordName) ($($Device[$HashKey].IPAddress))" -MessageLevel Warning
							Write-NetworkScanLog -Message "WMI response from $($Device[$HashKey].DnsRecordName) ($($Device[$HashKey].IPAddress)): $($Device[$HashKey].IsWmiAlive)" -MessageLevel Verbose
						} elseif ($Device[$HashKey].WmiMachineName) {
							Write-NetworkScanLog -Message "Timeout waiting for WMI response from $($Device[$HashKey].WmiMachineName) ($($Device[$HashKey].IPAddress))" -MessageLevel Warning
							Write-NetworkScanLog -Message "WMI response from $($Device[$HashKey].WmiMachineName) ($($Device[$HashKey].IPAddress)): $($Device[$HashKey].IsWmiAlive)" -MessageLevel Verbose
						} else {
							Write-NetworkScanLog -Message "Timeout waiting for WMI response from $($Device[$HashKey].IPAddress): $($Device[$HashKey].IsWmiAlive)" -MessageLevel Warning
							Write-NetworkScanLog -Message "WMI response from $($Device[$HashKey].IPAddress): $($Device[$HashKey].IsWmiAlive)" -MessageLevel Verbose
						}
					}
				}

				# Found that in some cases an ACCESS_VIOLATION error occurs if we don't delay a little bit during each iteration 
				Start-Sleep -Milliseconds 250

				# Clean out unused runspace jobs
				$Runspaces.clone() | Where-Object { ($_.Runspace -eq $null) } | ForEach-Object {
					$Runspaces.remove($_)
					$ScanCount++
					Write-Progress -Activity 'Testing WMI connectivity' -PercentComplete (($ScanCount / $DeviceCount)*100) -Status "$ScanCount of $DeviceCount" -Id $WmiProgressId -ParentId $ParentProgressId
				}

			} while (($Runspaces | Where-Object {$_.Runspace -ne $null} | Measure-Object).Count -gt 0)

			#endregion WMI CONNECTIVITY TEST


			# Finally, close the runspaces
			$RunspacePool.close()

			Write-NetworkScanLog -Message 'WMI connectivity test complete' -MessageLevel Verbose
			Write-Progress -Activity 'Testing PING connectivity' -PercentComplete 100 -Status 'Complete' -Id $PingProgressId -ParentId $ParentProgressId -Completed
			Write-Progress -Activity 'Testing WMI connectivity' -PercentComplete 100 -Status 'Complete' -Id $WmiProgressId -ParentId $ParentProgressId -Completed

		}


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

			BITS    SUBNET MASK        USABLE HOSTS PER SUBNET
			----    ---------------    -----------------------
			/20	    255.255.240.0      4094
			/21	    255.255.248.0      2046
			/22	    255.255.252.0      1022
			/23	    255.255.254.0      510
			/24	    255.255.255.0      254
			/25	    255.255.255.128    126
			/26	    255.255.255.192    62
			/27	    255.255.255.224    30
			/28	    255.255.255.240    14
			/29	    255.255.255.248    6
			/30	    255.255.255.252    2
			/32	    255.255.255.255    1

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
	[CmdletBinding()]
	param(
		[Parameter(
				Mandatory=$true,
				ParameterSetName='dns',
				ValueFromPipelineByPropertyName=$true,
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
				ValueFromPipelineByPropertyName=$true,
				HelpMessage='DNS Domain Name'
		)] 
		[alias('domain')]
		[string]
		$DnsDomain = 'automatic'
		,
		[Parameter(
				Mandatory=$true,
				ParameterSetName='subnet',
				ValueFromPipelineByPropertyName=$true,
				HelpMessage='Subnet (in CIDR notation)'
		)] 
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)[\\/]\d{1,2}$|^auto$|^automatic$')]
		[string[]]
		$Subnet = 'Automatic'
		,
		[Parameter(
				Mandatory=$true,
				ParameterSetName='computername',
				ValueFromPipelineByPropertyName=$true,
				HelpMessage='Computer Name(s)'
		)] 
		[alias('Computer')]
		[string[]]
		$ComputerName
		,
		[Parameter(
				Mandatory=$false,
				ParameterSetName='dns',
				ValueFromPipelineByPropertyName=$true
		)]
		[Parameter(
				Mandatory=$false, 
				ParameterSetName='subnet',
				ValueFromPipelineByPropertyName=$true
		)]
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)[\\/]\d{1,2}$')]
		[string[]]
		$ExcludeSubnet
		,
		[Parameter(
				Mandatory=$false, 
				ParameterSetName='dns',
				ValueFromPipelineByPropertyName=$true
		)]
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)[\\/]\d{1,2}$')]
		[string[]]
		$LimitSubnet
		,
		[Parameter(
				Mandatory=$false, 
				ParameterSetName='dns',
				ValueFromPipelineByPropertyName=$true
		)]
		[Parameter(
				Mandatory=$false,
				ParameterSetName='subnet',
				ValueFromPipelineByPropertyName=$true
		)]
		[string[]]
		$ExcludeComputerName
		,
		[Parameter(
				Mandatory=$false, 
				ParameterSetName='dns',
				ValueFromPipelineByPropertyName=$true
		)]
		[Parameter(
				Mandatory=$false,
				ParameterSetName='subnet',
				ValueFromPipelineByPropertyName=$true
		)]
		[string[]]
		$IncludeComputerName
		,
		[Parameter(
				Mandatory=$false,
				ValueFromPipelineByPropertyName=$true
		)] 
		[ValidateRange(1,100)]
		[alias('Throttle')]
		[byte]
		$MaxConcurrencyThrottle = $env:NUMBER_OF_PROCESSORS
		,
		[Parameter(
				Mandatory=$false,
				ValueFromPipelineByPropertyName=$true
		)] 
		[switch]
		$PrivateOnly = $false
		,
		[Parameter(
				Mandatory=$false,
				ValueFromPipelineByPropertyName=$true
		)] 
		[switch]
		$SkipConnectionTest = $false
		,
		[Parameter(
				Mandatory=$false,
				ValueFromPipelineByPropertyName=$true
		)]
		[ValidateNotNull()]
		[Int32]
		$ParentProgressId = -1
	)
	process {

		$Service = New-Object -TypeName psobject -Property @{
			Service = @()
		}
		$IPv4Device = @()
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

		# First thing's first, use the parameters passed in to scan the network to find WMI capable (i.e. Windows) devices
		$IPv4Device = New-Object -TypeName psobject -Property $(New-Object -TypeName System.Collections.Hashtable -ArgumentList $PSBoundParameters) | 
		Find-IPv4Device -MaxConcurrencyThrottle $MaxConcurrencyThrottle -PrivateOnly:$PrivateOnly -SkipConnectionTest:$SkipConnectionTest -ParentProgressId $ParentProgressId -ResolveAliases

		$WmiDeviceCount = $(
			$IPv4Device | 
			Where-Object { 
				$SkipConnectionTest -or
				$_.IsWmiAlive -eq $true
			} | 
			Group-Object -Property DnsRecordName | 
			Measure-Object
		).Count

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
				[String]$IPAddress,
				[String]$ComputerName
			)

			# Load SMO Assemblies
			[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | ForEach-Object {
				if ($_.GetName().Version.Major -ge 10) {
					$null = [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMOExtended')
					$null = [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SQLWMIManagement')
				}
			}

			# Registry constants
			New-Variable -Name HKEY_LOCAL_MACHINE -Value 2147483650 -Scope Script -Option Constant


			# Variables
			$DomainName = $null
			$ManagedComputerServerInstanceName = $null

			$StdRegProv = $null
			$RegistryKeyInstanceNameRoot = $null
			$RegistryKeySoftwareRoot = $null
			$RegistryKeyInstanceSubkey = $null
			$RegistryKeyInstanceIdRoot = $null
			$Parameters = $null
			$StartupParameters = $null
			$ServiceCollection = @()

			$SqlService = $null
			$SqlServiceBase = $null
			$ServiceFilter = [String]::Empty
			$IncludeService = [String]::Empty
			$ExcludeService = [String]::Empty
			$ServicePropertyList = @('Name', 'DisplayName', 'PathName', 'StartMode', 'ProcessId', 'State', 'StartName', 'Description')
			$InstanceId = [String]::Empty
			$HasWow6432Node = $null
			$EditionDescriptor = ' (64 Bit)'


			$ManagedComputer = New-Object -TypeName 'Microsoft.SqlServer.Management.Smo.Wmi.ManagedComputer' -ArgumentList $ComputerName
			#$ManagedServiceTypeEnum = 'Microsoft.SqlServer.Management.Smo.Wmi.ManagedServiceType' -as [Type]

			function Get-SqlServiceObject
			{
				Write-Output (
					New-Object -TypeName psobject -Property @{
						ComputerName        = $null
						DisplayName         = $null
						Description         = $null
						Edition             = $null
						ComputerIpAddress   = $null
						InstanceName        = $null
						IsNamedInstance     = $null
						IsClusteredInstance = $null
						IsDynamicPort       = $null
						IsHadrEnabled       = $null
						PathName            = $null
						Port                = $null
						ProcessId           = $null
						ServerName          = $null
						ServiceIpAddress    = $null
						ServicePackLevel    = $null
						ServiceProtocols    = $null
						ServiceStartDate    = $null
						ServiceState        = $null
						ServiceTypeName     = $null
						ServiceAccount      = $null
						StartMode           = $null
						StartupParameters   = $null
						Version             = $null
						VirtualServerName   = $null
					}
				)
			}

			function Get-RegistryKeyValueEnum
			{
				[CmdletBinding()]
				[OutputType([System.Collections.Hashtable])]
				Param
				(
					[Parameter(Mandatory=$true, Position=0)]
					[System.Management.ManagementClass]
					$StdRegProv,
					[Parameter(Mandatory=$true, Position=1)]
					[uint32]
					$RootKey,
					[Parameter(Mandatory=$true, Position=2)]
					[string]
					$SubKey,
					[Parameter(Mandatory=$false, Position=3)]
					[string[]]
					$Exclude = @()
				)

				Begin
				{
					$KeyValue = @{}
					$Parameters = $null 
					$i = 0
					$ValueName = $null
					$Value = $null
				}
				Process
				{
					$Parameters = $StdRegProv.EnumValues($RootKey,$SubKey)

					for ($i = 0; $i -lt ($Parameters.sNames | Measure-Object).Count; $i++) {

						$ValueName = '{0}' -f $Parameters.sNames[$i]
                        
						if ($Exclude -inotcontains $ValueName) {
                            
							switch ($Parameters.Types[$i]) {
								1 {
									# REG_SZ
									$Value = $($StdRegProv.GetStringValue($RootKey,$SubKey,$ValueName)).sValue
								}
								2 { 
									# REG_EXPAND_SZ
									$Value = $($StdRegProv.GetExpandedStringValue($RootKey,$SubKey,$ValueName)).sValue
								}
								3 {
									# REG_BINARY
									$Value = [System.BitConverter]::ToString($($StdRegProv.GetBinaryValue($RootKey,$SubKey,$ValueName).uValue) )
								}
								4 {
									# REG_DWORD
									$Value = $($StdRegProv.GetDWORDValue($RootKey,$SubKey,$ValueName)).uValue
								}
								7 {
									# REG_MULTI_SZ
									$Value = $($StdRegProv.GetMultiStringValue($RootKey,$SubKey,$ValueName)).sValue | ForEach-Object {
										$_
									} 
								}
								default { $Value = $null }
							}

							$KeyValue.Add($ValueName, $Value)

						}
					}
                    
				}
				End
				{
					Write-Output $KeyValue
					Remove-Variable -Name KeyValue, Parameters, i, ValueName, Value
				}
			}

			# Determine an IP Address & port the SQL Server service is listening on (if TCP/IP enabled)
			function Set-SqlServerServiceIpAddress
			{
				[CmdletBinding()]
				Param
				(
					[Parameter(Mandatory=$true,Position=0)]
					[PSObject]
					$SqlService
				)
				Begin
				{
				}
				Process
				{
					$SqlService.ServiceProtocols | 
					Where-Object { 
						$_.Name -ieq 'tcp' -and
						$_.IsEnabled -eq $true
					} |
					ForEach-Object {

						# If listening on all IPs then get the "IPAll" port info
						# Otherwise get port info for an IP that's enabled and active

						if ($_.ProtocolProperties['ListenOnAllIPs'] -eq $true) {

							$_.IPAddresses | 
							Where-Object { $_.Name -ieq 'ipall' } | 
							ForEach-Object {
								$SqlService.Port = $_.IPAddressProperties['TcpPort']
								if (-not $SqlService.Port) {
									$SqlService.Port = $_.IPAddressProperties['TcpDynamicPorts']
									$SqlService.IsDynamicPort = $true
								} else {
									$SqlService.IsDynamicPort = $false
								}
							}

						} else {

							# Start with 127.0.0.1 first in case that's the only IP that's enabled
							$_.IPAddresses | 
							Where-Object {
								$_.IPAddressProperties['Active'] -eq $true -and 
								$_.IPAddressProperties['Enabled'] -eq $true -and
								$_.IPAddressFamily -ieq 'InterNetwork' -and
								$_.IPAddress -ieq '127.0.0.1'
							} | 
							Select-Object -First 1 | 
							ForEach-Object {

								$SqlService.ServiceIpAddress = $_.IPAddress
								$SqlService.Port = $_.IPAddressProperties['TcpPort']

								if (-not $SqlService.Port) {
									$SqlService.Port = $_.IPAddressProperties['TcpDynamicPorts']
									$SqlService.IsDynamicPort = $true
								} else {
									$SqlService.IsDynamicPort = $false
								}
							}

							# Now try and see if there's a non-loopback IP enabled
							$_.IPAddresses | 
							Where-Object {
								$_.IPAddressProperties['Active'] -eq $true -and 
								$_.IPAddressProperties['Enabled'] -eq $true -and
								$_.IPAddressFamily -ieq 'InterNetwork' -and
								$_.IPAddress -ine '127.0.0.1'
							} | 
							Select-Object -First 1 | 
							ForEach-Object {

								$SqlService.ServiceIpAddress = $_.IPAddress
								$SqlService.Port = $_.IPAddressProperties['TcpPort']

								if (-not $SqlService.Port) {
									$SqlService.Port = $_.IPAddressProperties['TcpDynamicPorts']
									$SqlService.IsDynamicPort = $true
								} else {
									$SqlService.IsDynamicPort = $false
								}
							} 
						}
					}
				}
				End
				{
				}
			}

			function Get-SqlServiceServername
			{
				[CmdletBinding()]
				[OutputType([string])]
				Param
				(
					# Param1 help description
					[Parameter(Mandatory=$true, Position=0)]
					[PSObject]
					$SqlService
				)

				Begin
				{
					$NetbiosComputerName = $SqlService.ComputerName.Split('.')[0].PadRight(15,' ').Substring(0,15).Trim()
				}
				Process
				{
					if ($SqlService.IsNamedInstance -eq $true) {
						if ($SqlService.IsClusteredInstance -eq $true) {
							if ($SqlService.ServiceTypeName -ine 'SQL Server Analysis Services') {
								[String]::Join('\', @($SqlService.VirtualServerName, $SqlService.InstanceName))
							} else {
								# Connections to SSAS on a cluster do not use the instance name - only the cluster name
								$SqlService.VirtualServerName
							}
						} else {
							if (
								$SqlService.ServiceTypeName -ine 'SQL Server' -or
								$(
									$SqlService.ServiceProtocols | Where-Object {
										$_.Name -ine 'sm' -and
										$_.IsEnabled -eq $true
									}
								)
							) {
								# Another protocol besides shared memory is enabled or the service isn't SQL Server\SQL Server Agent; use the FQDN
								[String]::Join('\', @($SqlService.ComputerName, $SqlService.InstanceName))
							} else {
								# Shared memory is the only protocol enabled; use the NETBIOS name
								[String]::Join('\', @($NetbiosComputerName, $SqlService.InstanceName))
							}
						}
					} else {
						if ($SqlService.IsClusteredInstance -eq $true) {
							$SqlService.VirtualServerName
						} else {
							if (
								$SqlService.ServiceTypeName -ine 'SQL Server' -or
								$(
									$SqlService.ServiceProtocols | Where-Object {
										$_.Name -ine 'sm' -and
										$_.IsEnabled -eq $true
									}
								)
							) {
								# Another protocol besides shared memory is enabled or the service isn't SQL Server\SQL Server Agent; use the FQDN
								$SqlService.ComputerName
							} else {
								# Shared memory is the only protocol enabled; use the NETBIOS name
								$NetbiosComputerName
							}
						}
					}

				}
				End
				{
					Remove-Variable -Name NetbiosComputerName
				}
			}


			#region WMI Managed Computer Object

			# Try to use the WMI Managed Computer Object to find SQL Services on $ComputerName
			try {
				$ManagedComputer.Services | 
				ForEach-Object {

					$SqlService = Get-SqlServiceObject
					$SqlService.ComputerName = $ComputerName
					$SqlService.DisplayName = $_.DisplayName
					$SqlService.Description = $_.Description
					$SqlService.Edition = $_.AdvancedProperties['SKUNAME'].Value
					$SqlService.ComputerIpAddress = $IPAddress
					$SqlService.IsHadrEnabled = $_.IsHadrEnabled
					$SqlService.PathName = $_.PathName
					$SqlService.ProcessId = $_.ProcessId
					$SqlService.ServiceIpAddress = $null
					$SqlService.ServicePackLevel = $_.AdvancedProperties['SPLEVEL'].Value
					$SqlService.ServiceState = $_.ServiceState.ToString()
					$SqlService.ServiceAccount = $_.ServiceAccount
					$SqlService.StartMode = $_.StartMode.ToString()
					$SqlService.StartupParameters = $_.StartupParameters
					$SqlService.Version = $_.AdvancedProperties['VERSION'].Value


					if (($_.Name).IndexOf('$') -gt 0) {
						$SqlService.InstanceName = ($_.Name).Substring(($_.Name).IndexOf('$') + 1)
						$SqlService.IsNamedInstance = $true
						$ManagedComputerServerInstanceName = $SqlService.InstanceName
					} else {
						$SqlService.InstanceName = 'MSSQLSERVER'
						$SqlService.IsNamedInstance = $false
						$ManagedComputerServerInstanceName = $_.Name
					}


					# Get the Friendly name for the service
					$SqlService.ServiceTypeName = switch ($_.Type.value__) {
						1 { 'SQL Server' }
						2 { 'SQL Server Agent' }
						3 { 'SQL Server FullText Search' }
						4 { 'SQL Server Integration Services' }
						5 { 'SQL Server Analysis Services' }
						6 { 'SQL Server Reporting Services' }
						7 { 'SQL Server Browser' }
						8 { 'SQL Server Notification Services' }
						9 { 'SQL Full-text Filter Daemon Launcher' }
                        12 { 'SQL Server Launchpad' }
                        $null { 'Unknown' }
                        default { 'Unknown' }
					}


					# Determine if this is a clustered server (and the FQDN for the cluster if it is)
					# If it is a clustered instance don't worry about resolving the service IP address - let the client making the connection handle that
					# ...because you can get into some complicated scenarios with multi-subnet clusters that are best left to letting the client resolve
					# Also, SQL Browser service cannot be clustered but SMO shows that it is so explicitly exclude it
					if ($_.AdvancedProperties['CLUSTERED'].Value -eq $true -and $SqlService.ServiceTypeName -ine 'SQL Server Browser') {

						$SqlService.VirtualServerName = $_.AdvancedProperties['VSNAME'].Value

						if (-not [String]::IsNullOrEmpty($SqlService.VirtualServerName)) {
							$SqlService.IsClusteredInstance = $true
							$SqlService.ServiceIpAddress = $null

							try {
								if ([String]::IsNullOrEmpty($DomainName)) {
									Get-WmiObject -Namespace root\CIMV2 -Class Win32_ComputerSystem -Property Domain -ComputerName $IPAddress -ErrorAction Stop | 
									ForEach-Object {
										$DomainName = $_.Domain
									}
								}
								$SqlService.VirtualServerName = [String]::Join('.', @($SqlService.VirtualServerName, $DomainName)).ToUpper()
							}
							catch {
								# Unable to resolve the domain name for the clustered service so just use the virtual server name from the registry
								$SqlService.VirtualServerName = $_.AdvancedProperties['VSNAME'].Value
							}

						} else {
							$SqlService.VirtualServerName = [String]::Empty
						}

					} else {
						$SqlService.IsClusteredInstance = $false
						$SqlService.VirtualServerName = $null
					}


					# Gather protocol details for SQL Server service
					if ($SqlService.ServiceTypeName -ieq 'SQL Server') {

						# Gather protocol details
						$SqlService.ServiceProtocols = $ManagedComputer.ServerInstances | 
						Where-Object { $_.Name -ieq $ManagedComputerServerInstanceName } | 
						ForEach-Object {
							$_.ServerProtocols | 
							Where-Object { -not [String]::IsNullOrEmpty($_.Name) } | 
							ForEach-Object {
								New-Object -TypeName PSObject -Property @{
									Name               = $_.Name
									DisplayName        = $_.DisplayName
									IsEnabled          = $_.IsEnabled
									IPAddresses        = $_.IPAddresses |
									Where-Object { $_.IPAddress } |
									ForEach-Object {
										New-Object -TypeName PSObject -Property @{
											Name                = $_.Name
											IpAddress           = $_.IPAddress.ToString()
											IpAddressFamily     = $_.IPAddress.AddressFamily.ToString()
											IPAddressProperties = $(
												$IpAddressProperty = @{}
												$_.IPAddressProperties |
												Where-Object {
													-not [String]::IsNullOrEmpty($_.Name)
												} |
												ForEach-Object { 
													$IpAddressProperty.Add($_.Name, $_.Value)
												}
												Write-Output $IpAddressProperty
											)
										}
									}
									ProtocolProperties = $(
										$ProtocolProperty = @{}
										$_.ProtocolProperties |
										Where-Object { 
											-not [String]::IsNullOrEmpty($_.Name) -and
											$_.Name -ine 'Enabled' 
										} |
										ForEach-Object {
											$ProtocolProperty.Add($_.Name, $_.Value)
										}
										Write-Output $ProtocolProperty
									)
								}
							}
						}


						# Figure out the IP address and port that SQL Server is listening on
						Set-SqlServerServiceIpAddress -SqlService $SqlService

					}


					# Get the Service Start Date (if it's got a Process ID greater than than 0)
					if ($_.ProcessId -gt 0) {
						try {
							Get-WmiObject -Namespace root\CIMV2 -Class Win32_Process -Filter "ProcessId = '$($_.ProcessId)'" -Property CreationDate -ComputerName $IPAddress -ErrorAction Stop | 
							ForEach-Object {
								$SqlService.ServiceStartDate = $_.ConvertToDateTime($_.CreationDate)
							}
						}
						catch {
							$SqlService.ServiceStartDate = $null
						}
					} else {
						$SqlService.ServiceStartDate = $null
					}


					# Finally, figure out the servername for this service
					$SqlService.ServerName = Get-SqlServiceServername -SqlService $SqlService


					Write-Output -InputObject $SqlService

					# Add the service to the discovered service collection
					$ServiceCollection += $SqlService.DisplayName

				} 
			}
			catch {

				# If we get this error it's possible that this is a SQL 2000 server in which case we need to look at the registry to find installed services
				if ($_.Exception.Message -inotlike '*SQL Server WMI provider is not available*') {
					# Something else has happened; Let the error bubble up
					throw
				}
			}
			#endregion


			#region SQL Server Services (WMI & registry)

			# Now look through the machine's registry for instances that were not found by the ManagedComputer object
			# This can happen for a few reasons:
			#   - Using a lower version of SMO on the machine where the script is running than the version of SQL Server on the target server
			#   - The target server is running one or more SQL 2000 instances
			# This is not the preferred way to get service information because it assumes things about registry paths...but it's better than nothing
			try {

				$SqlServiceBase = Get-SqlServiceObject
				$SqlServiceBase.ComputerName = $ComputerName
				$SqlServiceBase.ComputerIpAddress = $IPAddress
				$SqlServiceBase.IsHadrEnabled = $null
				$SqlServiceBase.ServiceIpAddress = $null
				$SqlServiceBase.ServiceProtocols = $null
				$SqlServiceBase.Version = $null


				# Use the WMI Registry provider to access the registry on $ComputerName
				# For info on using this WMI class see http://msdn.microsoft.com/en-us/library/windows/desktop/aa393664(v=vs.85).aspx
				$StdRegProv = Get-WmiObject -Namespace root\DEFAULT -Query "select * FROM meta_class WHERE __Class = 'StdRegProv'" -ComputerName $ComputerName -ErrorAction Stop


				# Check if there's a Wow6432Node node in the registry. Use this later for determining 32-bit vs 64-bit versions of SQL Server
				# Return code of 0 means it DOES exist
				if (($StdRegProv.EnumKey($HKEY_LOCAL_MACHINE,'SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server')).ReturnValue -eq 0) {
					$HasWow6432Node = $true
				} else {
					$HasWow6432Node = $false
				}


				# Let's use WMI and the registry to gather service information
				# A few rules:
				# - There can only be one instance of SQL Browser and Integration Services in 2005 and higher (and it cannot be a named instance)
				# - There can be only one instance of Analysis Services, Reporting Services, or Microsoft Search in 2000 (and it cannot be a named instance)
				# - There can be multiple instances of Analysis Services, Reporting Services, and Fulltext Services in 2005 and higher
				# - Only SSAS in SQL 2012 and higher can be clustered
                
				$IncludeService = @(
					'(DisplayName="MSSQLSERVER")', 
					'(DisplayName LIKE "MSSQL$%")', 
					'(DisplayName="SQLSERVERAGENT")', 
					'(DisplayName LIKE "SQLAgent$%")', 
					'(DisplayName="MSSQLServerOLAPService")', 
					'(DisplayName="Microsoft Search")', 
					'(DisplayName="ReportServer")', 
					'(DisplayName="SQL Server Browser")', 
					'(DisplayName LIKE "NS$%")', 
					'(DisplayName LIKE "SQL Server Integration Services%")', 
					'(DisplayName LIKE "SQL Server (%)")', 
					'(DisplayName LIKE "SQL Server Agent (%)")', 
					'(DisplayName LIKE "SQL Server Analysis Services (%)")', 
					'(DisplayName LIKE "SQL Server FullText Search (%)")', 
					'(DisplayName LIKE "SQL Full-text Filter Daemon Launcher (%)")', 
					'(DisplayName LIKE "SQL Server Reporting Services (%)")'
				) -join ' OR '

				# Build the WMI filter parameter based on services to find and services already found
				if ([String]::IsNullOrEmpty($ServiceCollection) -eq $true) {
					$ServiceFilter = $IncludeService
				} else {
					$ExcludeService = $($ServiceCollection | ForEach-Object { 'DisplayName != "{0}"' -f $_ }) -join ' AND '
					$ServiceFilter = '({0}) AND ({1})' -f $IncludeService, $ExcludeService
				}

				Get-WmiObject -Namespace root\CIMV2 -Class Win32_Service -Filter $ServiceFilter -Property $ServicePropertyList -ComputerName $IPAddress -ErrorAction Stop | 
				ForEach-Object {

					# Copy the base object first
					$SqlService = $SqlServiceBase.psobject.Copy()

					$SqlService.ServerName = $ComputerName
					$SqlService.DisplayName = $_.DisplayName
					$SqlService.PathName = $_.PathName
					$SqlService.StartMode = $_.StartMode
					$SqlService.ProcessId = $_.ProcessId
					$SqlService.ServiceState = $_.State
					$SqlService.ServiceAccount = $_.StartName
					$SqlService.Description = $_.Description

					<# Assume the instance is not clustered. We'll figure this out for sure in just a bit #>
					$SqlService.IsClusteredInstance = $false
					$SqlService.VirtualServerName = [String]::Empty


					# Build registry key paths for the instance
					if (($_.Name).IndexOf('$') -gt 0) {
						$SqlService.InstanceName = ($_.Name).Substring(($_.Name).IndexOf('$') + 1)
						$SqlService.IsNamedInstance = $true
						$RegistryKeyInstanceSubkey = 'Microsoft SQL Server\{0}' -f $SqlService.InstanceName
					} else {
						$SqlService.InstanceName = 'MSSQLSERVER'
						$SqlService.IsNamedInstance = $false
						$RegistryKeyInstanceSubkey = 'MSSQLServer' -f $SqlService.InstanceName
					}
					$RegistryKeySoftwareRoot = 'SOFTWARE'
					$RegistryKeyInstanceNameRoot = '{0}\Microsoft\{1}' -f $RegistryKeySoftwareRoot, $RegistryKeyInstanceSubkey


					# Get the Friendly name for the service based on the service name
					$SqlService.ServiceTypeName = switch -Wildcard ($_.Name) {
						'MSSQLSERVER' { 'SQL Server' }
						'MSSQL$*' { 'SQL Server' }
						'SQLSERVERAGENT' { 'SQL Server Agent' }
						'SQLAgent$*' { 'SQL Server Agent' }
						'MSSEARCH' { 'Microsoft Search service' }
						'msftesql' { 'SQL Server FullText Search' }
						'msftesql$*' { 'SQL Server FullText Search' }
						'MsDtsServer*' { 'SQL Server Integration Services' }
						'MSSQLServerOLAPService' { 'SQL Server Analysis Services' }
						'MSOLAP$*' { 'SQL Server Analysis Services' }
						'ReportServer' { 'SQL Server Reporting Services' }
						'ReportServer$*' { 'SQL Server Reporting Services' }
						'SQLBrowser' { 'SQL Server Browser' }
						'NS$*' { 'SQL Server Notification Services' }
						'MSSQLFDLauncher' { 'SQL Full-text Filter Daemon Launcher' }
						'MSSQLFDLauncher$*' { 'SQL Full-text Filter Daemon Launcher' }
						default { 'Unknown' }
					}


					# Set defaults that apply to all services
					$SqlService.IsDynamicPort = $null
					$SqlService.Port = $null
					$SqlService.ServerName = $ComputerName

					<# Gather details specific to SQL Server service #>
					if ($SqlService.ServiceTypeName -ieq 'SQL Server') {
                        
						if ($SqlService.DisplayName -ieq 'MSSQLSERVER' -or $SqlService.DisplayName -ilike 'MSSQL$*') {
                            
							# This condition means <= SQL 2000

							# Determine if this is a clustered server (and the FQDN if it is)
							# If it is a clustered instance don't worry about resolving the service IP address - let the client making the connection handle that
							$SqlService.VirtualServerName = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceNameRoot\Cluster", 'ClusterName')).sValue

							if (-not [String]::IsNullOrEmpty($SqlService.VirtualServerName)) {
								$SqlService.IsClusteredInstance = $true
								$SqlService.ServiceIpAddress = $null

								try {
									if ([String]::IsNullOrEmpty($DomainName)) {
										Get-WmiObject -Namespace root\CIMV2 -Class Win32_ComputerSystem -Property Domain -ComputerName $IPAddress -ErrorAction Stop | 
										ForEach-Object {
											$DomainName = $_.Domain
										}
									}
									$SqlService.VirtualServerName = [String]::Join('.', @($SqlService.VirtualServerName, $DomainName)).ToUpper()
								}
								catch {
									# Unable to resolve the domain name for the clustered service so just use the virtual server name from the registry
									$SqlService.VirtualServerName = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceNameRoot\Cluster", 'ClusterName')).sValue
								}

							} else {
								$SqlService.IsClusteredInstance = $false
								$SqlService.VirtualServerName = [String]::Empty
								$SqlService.ServiceIpAddress = $IPAddress
							}


							$SqlService.Edition = $null # Not exposed in SQL 2000
							$SqlService.Version = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceNameRoot\MSSQLServer\CurrentVersion", 'CurrentVersion')).sValue
							$SqlService.ServicePackLevel = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceNameRoot\MSSQLServer\CurrentVersion", 'CSDVersion')).sValue

							# Get the TCP port number for SQL Server Services
							$SqlService.Port = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceNameRoot\MSSQLServer\SuperSocketNetLib\Tcp",'TcpDynamicPorts')).sValue
							if (-not $SqlService.Port) {
								$SqlService.Port = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceNameRoot\MSSQLServer\SuperSocketNetLib\Tcp",'TcpPort')).sValue
								$SqlService.IsDynamicPort = $false
							} else {
								$SqlService.IsDynamicPort = $true
							}


							# Gather protocol details
							$SqlService.ServiceProtocols = $($StdRegProv.GetMultiStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceNameRoot\MSSQLServer\SuperSocketNetLib",'ProtocolList')).sValue | 
							ForEach-Object {                                
								New-Object -TypeName PSObject -Property @{
									Name               = $_
									DisplayName        = switch ($_) {
										'adsp' { 'Apple Talk' }
										'bv' { 'Banyan Vines' }
										'np' { 'Named Pipes' }
										'rpc' { 'Multiprotocol' }
										'spx' { 'NWLink IPX/SPX' }
										'tcp' { 'TCP/IP' }
										'via' { 'VIA' }
										default { $_ }
									}
									IsEnabled          = $true
									ProtocolProperties = Get-RegistryKeyValueEnum -StdRegProv $StdRegProv -RootKey $HKEY_LOCAL_MACHINE -SubKey "$RegistryKeyInstanceNameRoot\MSSQLServer\SuperSocketNetLib\$_"
								}                                
							}


							# If $ComputerName is the local host then use the loopback IP for connectivity, otherwise use $IpAddress
							if (
								$ComputerName -ieq $env:COMPUTERNAME -or
								$ComputerName.StartsWith([String]::Concat($env:COMPUTERNAME, '.'), [System.StringComparison]::InvariantCultureIgnoreCase)
							) {
								$SqlServiceBase.ServiceIpAddress = '127.0.0.1'
							} else {
								$SqlServiceBase.ServiceIpAddress = $IPAddress
							}


						} else {

							# This condition means >= SQL 2005

							# Figure out where to get instance details from the registry
							$InstanceId = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeySoftwareRoot\Microsoft\Microsoft SQL Server\Instance Names\SQL", $SqlService.InstanceName)).sValue
							if ([String]::IsNullOrEmpty($InstanceId)) {
								$RegistryKeySoftwareRoot = 'SOFTWARE\Wow6432Node'
								$RegistryKeyInstanceNameRoot = '{0}\Microsoft\{1}' -f $RegistryKeySoftwareRoot, $RegistryKeyInstanceSubkey
								$InstanceId = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeySoftwareRoot\Microsoft\Microsoft SQL Server\Instance Names\SQL", $SqlService.InstanceName)).sValue
								$EditionDescriptor = [String]::Empty
							} else {
								# If there's a Wow6432Node in the registry then this instance is 64-bit; If not then it's 32-bit
								if ($HasWow6432Node -eq $true) {
									$EditionDescriptor = ' (64-bit)'
								} else {
									$EditionDescriptor = [String]::Empty
								}
							}
							$RegistryKeyInstanceIdRoot = '{0}\Microsoft\Microsoft SQL Server\{1}' -f $RegistryKeySoftwareRoot, $InstanceId


							# Determine if this is a clustered server (and the FQDN if it is)
							# If it is a clustered instance don't worry about resolving the service IP address - let the client making the connection handle that
							# ...because you can get into some complicated scenarios with multi-subnet clusters that are best left to letting the client resolve
							$SqlService.VirtualServerName = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\Cluster", 'ClusterName')).sValue
							if (-not [String]::IsNullOrEmpty($SqlService.VirtualServerName)) {
								$SqlService.IsClusteredInstance = $true
								$SqlService.ServiceIpAddress = $null

								try {
									if ([String]::IsNullOrEmpty($DomainName)) {
										Get-WmiObject -Namespace root\CIMV2 -Class Win32_ComputerSystem -Property Domain -ComputerName $IPAddress -ErrorAction Stop | 
										ForEach-Object {
											$DomainName = $_.Domain
										}
									}
									$SqlService.VirtualServerName = [String]::Join('.', @($SqlService.VirtualServerName, $DomainName)).ToUpper()
								}
								catch {
									# Unable to resolve the domain name for the clustered service so just use the virtual server name from the registry
									$SqlService.VirtualServerName = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\Cluster", 'ClusterName')).sValue
								}

							} else {
								$SqlService.IsClusteredInstance = $false
								$SqlService.VirtualServerName = [String]::Empty
							}



							# Get other details from registry
							$SqlService.IsHadrEnabled = switch (($StdRegProv.GetDWORDValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\MSSQLServer\HADR", 'HADR_Enabled')).uValue) {
								0 { $false }
								1 { $true }
								default { $null }
							}
							$SqlService.Edition = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\Setup", 'Edition')).sValue
							if (-not [String]::IsNullOrEmpty($SqlService.Edition) -and $SqlService.Edition.IndexOf($EditionDescriptor) -lt 0) {
								$SqlService.Edition = '{0}{1}' -f $SqlService.Edition, $EditionDescriptor
							}
							$SqlService.Version = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\Setup", 'Version')).sValue
							$SqlService.ServicePackLevel = ($StdRegProv.GetDWORDValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\Setup", 'SP')).uValue


							# Gather protocol details
							# 3 protocols for SQL 2005 and up: Shared Memory, Named Pipes, & TCP/IP (although Via is in the registry if you care to look for yourself)
							$SqlService.ServiceProtocols = $(

								# Named Pipes
								New-Object -TypeName PSObject -Property @{
									Name               = 'Np'
									DisplayName        = 'Named Pipes'
									IsEnabled          = [Boolean]($StdRegProv.GetDWORDValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\MSSQLServer\SuperSocketNetLib\Np", 'Enabled')).uValue
									IPAddresses        = $null
									ProtocolProperties = Get-RegistryKeyValueEnum -StdRegProv $StdRegProv `
									-RootKey $HKEY_LOCAL_MACHINE `
									-SubKey "$RegistryKeyInstanceIdRoot\MSSQLServer\SuperSocketNetLib\Np" `
									-Exclude DisplayName, Enabled
								}

								# Shared Memory
								New-Object -TypeName PSObject -Property @{
									Name               = 'Sm'
									DisplayName        = 'Shared Memory'
									IsEnabled          = [Boolean]($StdRegProv.GetDWORDValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\MSSQLServer\SuperSocketNetLib\Sm", 'Enabled')).uValue
									IPAddresses        = $null
									ProtocolProperties = @{}
								}

								# TCP/IP
								New-Object -TypeName PSObject -Property @{
									Name               = 'Tcp'
									DisplayName        = 'TCP/IP'
									IsEnabled          = [Boolean]($StdRegProv.GetDWORDValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\MSSQLServer\SuperSocketNetLib\Tcp", 'Enabled')).uValue
									IPAddresses        = $(
										($StdRegProv.EnumKey($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\MSSQLServer\SuperSocketNetLib\Tcp")).sNames | 
										Where-Object {
											$_ -match '^IP(\d{1}|All)$'
										} |
										ForEach-Object {
											New-Object -TypeName PSObject -Property @{
												Name                = $_
												IPAddressProperties = Get-RegistryKeyValueEnum -StdRegProv $StdRegProv `
												-RootKey $HKEY_LOCAL_MACHINE `
												-SubKey "$RegistryKeyInstanceIdRoot\MSSQLServer\SuperSocketNetLib\Tcp\$_" `
												-Exclude DisplayName, Enabled
											} |
											Select-Object -Property Name, IPAddressProperties, @{
												Name       = 'IpAddress'
												Expression = {
													if ([String]::IsNullOrEmpty($_.IPAddressProperties['IpAddress'])) {
														'0.0.0.0'
													} else {
														$_.IPAddressProperties.IpAddress
													}
												}
											} | 
											Select-Object -Property Name, IpAddress, IPAddressProperties, @{
												Name       = 'IpAddressFamily'
												Expression = {
													([System.Net.IPAddress]$_.IpAddress).AddressFamily
												}
											}
										}
									)
									ProtocolProperties = Get-RegistryKeyValueEnum -StdRegProv $StdRegProv `
									-RootKey $HKEY_LOCAL_MACHINE `
									-SubKey "$RegistryKeyInstanceIdRoot\MSSQLServer\SuperSocketNetLib\Tcp\" `
									-Exclude DisplayName, Enabled
								}
							)


							# Figure out the IP address and port that SQL Server is listening on
							Set-SqlServerServiceIpAddress -SqlService $SqlService

						}

						# Get startup parameters
						$SqlService.StartupParameters = $(
							(Get-RegistryKeyValueEnum -StdRegProv $StdRegProv `
								-RootKey $HKEY_LOCAL_MACHINE `
								-SubKey "$RegistryKeyInstanceIdRoot\MSSQLServer\Parameters"
							).Values | Sort-Object
						) -join ';'
					}


					<# Gather details specific to SQL Server Agent service #>
					if ($SqlService.ServiceTypeName -ieq 'SQL Server Agent') {
						if ($SqlService.DisplayName -ieq 'SQLSERVERAGENT' -or $SqlService.DisplayName -ilike 'SQLAgent$*') {

							# Determine if this is a clustered server (and the FQDN if it is)
							# If it is a clustered instance don't worry about resolving the service IP address - let the client making the connection handle that
							$SqlService.VirtualServerName = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceNameRoot\Cluster", 'ClusterName')).sValue

							if (-not [String]::IsNullOrEmpty($SqlService.VirtualServerName)) {
								$SqlService.IsClusteredInstance = $true
								$SqlService.ServiceIpAddress = $null

								try {
									if ([String]::IsNullOrEmpty($DomainName)) {
										Get-WmiObject -Namespace root\CIMV2 -Class Win32_ComputerSystem -Property Domain -ComputerName $IPAddress -ErrorAction Stop | 
										ForEach-Object {
											$DomainName = $_.Domain
										}
									}
									$SqlService.VirtualServerName = [String]::Join('.', @($SqlService.VirtualServerName, $DomainName)).ToUpper()
								}
								catch {
									# Unable to resolve the domain name for the clustered service so just use the virtual server name from the registry
									$SqlService.VirtualServerName = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceNameRoot\Cluster", 'ClusterName')).sValue
								}

							} else {
								$SqlService.IsClusteredInstance = $false
								$SqlService.VirtualServerName = [String]::Empty
							}

						} else {

							# Figure out where to get instance details from the registry
							$InstanceId = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeySoftwareRoot\Microsoft\Microsoft SQL Server\Instance Names\SQL", $SqlService.InstanceName)).sValue
							if ([String]::IsNullOrEmpty($InstanceId)) {
								$RegistryKeySoftwareRoot = 'SOFTWARE\Wow6432Node'
								$RegistryKeyInstanceNameRoot = '{0}\Microsoft\{1}' -f $RegistryKeySoftwareRoot, $RegistryKeyInstanceSubkey
								$InstanceId = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeySoftwareRoot\Microsoft\Microsoft SQL Server\Instance Names\SQL", $SqlService.InstanceName)).sValue
								$EditionDescriptor = [String]::Empty
							} else {
								# If there's a Wow6432Node in the registry then this instance is 64-bit; If not then it's 32-bit
								if ($HasWow6432Node -eq $true) {
									$EditionDescriptor = ' (64-bit)'
								} else {
									$EditionDescriptor = [String]::Empty
								}
							}
							$RegistryKeyInstanceIdRoot = '{0}\Microsoft\Microsoft SQL Server\{1}' -f $RegistryKeySoftwareRoot, $InstanceId


							# Determine if this is a clustered server (and the FQDN if it is)
							# If it is a clustered instance don't worry about resolving the service IP address - let the client making the connection handle that
							# ...because you can get into some complicated scenarios with multi-subnet clusters that are best left to letting the client resolve
							$SqlService.VirtualServerName = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\Cluster", 'ClusterName')).sValue
							if (-not [String]::IsNullOrEmpty($SqlService.VirtualServerName)) {
								$SqlService.IsClusteredInstance = $true
								$SqlService.ServiceIpAddress = $null

								try {
									if ([String]::IsNullOrEmpty($DomainName)) {
										Get-WmiObject -Namespace root\CIMV2 -Class Win32_ComputerSystem -Property Domain -ComputerName $IPAddress -ErrorAction Stop | 
										ForEach-Object {
											$DomainName = $_.Domain
										}
									}
									$SqlService.VirtualServerName = [String]::Join('.', @($SqlService.VirtualServerName, $DomainName)).ToUpper()
								}
								catch {
									# Unable to resolve the domain name for the clustered service so just use the virtual server name from the registry
									$SqlService.VirtualServerName = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\Cluster", 'ClusterName')).sValue
								}

							} else {
								$SqlService.IsClusteredInstance = $false
								$SqlService.VirtualServerName = [String]::Empty
							}

						}
					}


					<# Gather details specific to SSAS service #>
					if ($SqlService.ServiceTypeName -ieq 'SQL Server Analysis Services') {
						if ($SqlService.DisplayName -ieq 'MSSQLServerOLAPService') {
							## These are not the right registry locations for SQL 2000 - verified in a VM with only SSAS installed. Are they available somewhere else?
							#$SqlService.Edition = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeySoftwareRoot\Microsoft\MSSQLServer\MSSQLServer\Setup", 'Edition')).sValue
							#$SqlService.Version = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeySoftwareRoot\Microsoft\MSSQLServer\MSSQLServer\CurrentVersion", 'CurrentVersion')).sValue
							#$SqlService.ServicePackLevel = ($StdRegProv.GetDWORDValue($HKEY_LOCAL_MACHINE,"$RegistryKeySoftwareRoot\Microsoft\OLAP Server\CurrentVersion", 'CSDVersion')).uValue

							$SqlService.Edition = $null
							$SqlService.Version = $null
							$SqlService.ServicePackLevel = $null

						} else {

							# Figure out where to get instance details from the registry
							$InstanceId = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeySoftwareRoot\Microsoft\Microsoft SQL Server\Instance Names\OLAP", $SqlService.InstanceName)).sValue
							if ([String]::IsNullOrEmpty($InstanceId)) {
								$RegistryKeySoftwareRoot = 'SOFTWARE\Wow6432Node'
								$RegistryKeyInstanceNameRoot = '{0}\Microsoft\{1}' -f $RegistryKeySoftwareRoot, $RegistryKeyInstanceSubkey
								$InstanceId = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeySoftwareRoot\Microsoft\Microsoft SQL Server\Instance Names\OLAP", $SqlService.InstanceName)).sValue
								$EditionDescriptor = [String]::Empty
							} else {
								# If there's a Wow6432Node in the registry then this instance is 64-bit; If not then it's 32-bit
								if ($HasWow6432Node -eq $true) {
									$EditionDescriptor = ' (64-bit)'
								} else {
									$EditionDescriptor = [String]::Empty
								}
							}
							$RegistryKeyInstanceIdRoot = '{0}\Microsoft\Microsoft SQL Server\{1}' -f $RegistryKeySoftwareRoot, $InstanceId


							# Determine if this is a clustered server (and the FQDN if it is)
							# If it is a clustered instance don't worry about resolving the service IP address - let the client making the connection handle that
							# ...because you can get into some complicated scenarios with multi-subnet clusters that are best left to letting the client resolve
							$SqlService.VirtualServerName = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\Cluster", 'ClusterName')).sValue
							if (-not [String]::IsNullOrEmpty($SqlService.VirtualServerName)) {
								$SqlService.IsClusteredInstance = $true
								$SqlService.ServiceIpAddress = $null

								try {
									if ([String]::IsNullOrEmpty($DomainName)) {
										Get-WmiObject -Namespace root\CIMV2 -Class Win32_ComputerSystem -Property Domain -ComputerName $IPAddress -ErrorAction Stop | 
										ForEach-Object {
											$DomainName = $_.Domain
										}
									}
									$SqlService.VirtualServerName = [String]::Join('.', @($SqlService.VirtualServerName, $DomainName)).ToUpper()
								}
								catch {
									# Unable to resolve the domain name for the clustered service so just use the virtual server name from the registry
									$SqlService.VirtualServerName = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\Cluster", 'ClusterName')).sValue
								}

							} else {
								$SqlService.IsClusteredInstance = $false
								$SqlService.VirtualServerName = [String]::Empty
							}


							$SqlService.Edition = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\Setup", 'Edition')).sValue
							if (-not [String]::IsNullOrEmpty($SqlService.Edition) -and $SqlService.Edition.IndexOf($EditionDescriptor) -lt 0) {
								$SqlService.Edition = '{0}{1}' -f $SqlService.Edition, $EditionDescriptor
							}
							$SqlService.Version = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\Setup", 'Version')).sValue
							$SqlService.ServicePackLevel = ($StdRegProv.GetDWORDValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\Setup", 'SP')).uValue
						}
					}


					<# Gather details specific to Reporting Services service #>
					if ($SqlService.ServiceTypeName -ieq 'SQL Server Reporting Services') {
						if ($SqlService.DisplayName -ieq 'ReportServer') {
							$SqlService.Edition = switch (($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeySoftwareRoot\Microsoft\MSSQLServer\MSSQLServer\Setup", 'Edition')).sValue) {
								'{2879CA50-1599-4F4B-B9EC-1110C1094C16}' { 'Developer Edition' }
								'{B19FEFE7-069D-4FC4-8FDF-19661EAB6CE4}' { 'Standard Edition' }
								'{7C93251A-BFB4-4EB8-A57C-81B875BB12E4}' { 'Evaluation Edition' }
								'{33FE9EED-1976-4A51-A7AF-332D9BBB9400}' { 'Enterprise Edition' }
							}
							$SqlService.Version = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeySoftwareRoot\Microsoft\MSSQLServer\MSSQLServer\CurrentVersion", 'CurrentVersion')).sValue
						} else {

							# Figure out where to get instance details from the registry
							$InstanceId = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeySoftwareRoot\Microsoft\Microsoft SQL Server\Instance Names\RS", $SqlService.InstanceName)).sValue
							if ([String]::IsNullOrEmpty($InstanceId)) {
								$RegistryKeySoftwareRoot = 'SOFTWARE\Wow6432Node'
								$RegistryKeyInstanceNameRoot = '{0}\Microsoft\{1}' -f $RegistryKeySoftwareRoot, $RegistryKeyInstanceSubkey
								$InstanceId = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeySoftwareRoot\Microsoft\Microsoft SQL Server\Instance Names\RS", $SqlService.InstanceName)).sValue
								$EditionDescriptor = [String]::Empty
							} else {
								# If there's a Wow6432Node in the registry then this instance is 64-bit; If not then it's 32-bit
								if ($HasWow6432Node -eq $true) {
									$EditionDescriptor = ' (64-bit)'
								} else {
									$EditionDescriptor = [String]::Empty
								}
							}
							$RegistryKeyInstanceIdRoot = '{0}\Microsoft\Microsoft SQL Server\{1}' -f $RegistryKeySoftwareRoot, $InstanceId

							$SqlService.Edition = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\Setup", 'Edition')).sValue
							if (-not [String]::IsNullOrEmpty($SqlService.Edition) -and $SqlService.Edition.IndexOf($EditionDescriptor) -lt 0) {
								$SqlService.Edition = '{0}{1}' -f $SqlService.Edition, $EditionDescriptor
							}
							$SqlService.Version = ($StdRegProv.GetStringValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\Setup", 'Version')).sValue
							$SqlService.ServicePackLevel = ($StdRegProv.GetDWORDValue($HKEY_LOCAL_MACHINE,"$RegistryKeyInstanceIdRoot\Setup", 'SP')).uValue
						}
					}



					# Get the Service Start Date (if it's got a Process ID greater than than 0)
					if ($SqlService.ProcessId -gt 0) {
						try {
							Get-WmiObject -Namespace root\CIMV2 -Class Win32_Process -Filter "ProcessId = '$($SqlService.ProcessId)'" -Property CreationDate -ComputerName $IPAddress -ErrorAction Stop | 
							ForEach-Object {
								$SqlService.ServiceStartDate = $_.ConvertToDateTime($_.CreationDate)
							}
						}
						catch {
							$SqlService.ServiceStartDate = $null
						}
					} else {
						$SqlService.ServiceStartDate = $null
					}


					# Finally, figure out the servername for this service
					$SqlService.ServerName = Get-SqlServiceServername -SqlService $SqlService

					Write-Output -InputObject $SqlService

				}
				#endregion


			}
			catch {
				throw
			}

		}

		# Iterate through each machine that we could make a WMI connection to and gather information
		# Some machines may have multiple entries (b\c of multiple IP Addresses) so only use the first IP Address for each
		$IPv4Device | 
		Where-Object { 
			$SkipConnectionTest -or
			$_.IsWmiAlive -eq $true 
		} | 
		Group-Object -Property DnsRecordName | 
		ForEach-Object {

			$ScanCount++
			Write-NetworkScanLog -Message "Scanning $(($_.Group[0]).DnsRecordName) at IP address $(($_.Group[0]).IPAddress) for SQL Services [Device $ScanCount of $WmiDeviceCount]" -MessageLevel Information

			#Create the PowerShell instance and supply the scriptblock with the other parameters
			$PowerShell = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
			$PowerShell = $PowerShell.AddArgument($($_.Group[0]).IPAddress)
			$PowerShell = $PowerShell.AddArgument($($_.Group[0]).DnsRecordName)

			#Add the runspace into the PowerShell instance
			$PowerShell.RunspacePool = $RunspacePool

			$null = $Runspaces.Add((
					New-Object -TypeName PsObject -Property @{
						PowerShell   = $PowerShell
						Runspace     = $PowerShell.BeginInvoke()
						ComputerName = $($_.Group[0]).DnsRecordName
						IPAddress    = $($_.Group[0]).IPAddress
					}
			))
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

							if (-not [String]::IsNullOrEmpty($_.ServiceIpAddress)) {
							} else {
							}


							if ($_.IsNamedInstance -eq $true) {
								if ([String]::IsNullOrEmpty($_.ServiceIpAddress)) {
									Write-NetworkScanLog -Message "Found $($_.ServiceTypeName) named instance $($_.ServerName)" -MessageLevel Information
								} else {
									Write-NetworkScanLog -Message "Found $($_.ServiceTypeName) named instance $($_.ServerName) on IP address $($_.ServiceIpAddress)" -MessageLevel Information
								}
							} else {
								if ([String]::IsNullOrEmpty($_.ServiceIpAddress)) {
									Write-NetworkScanLog -Message "Found $($_.ServiceTypeName) default instance $($_.ServerName)" -MessageLevel Information
								} else {
									Write-NetworkScanLog -Message "Found $($_.ServiceTypeName) default instance $($_.ServerName) on IP address $($_.ServiceIpAddress)" -MessageLevel Information
								}
							}
						}

					}
					catch {
						Write-NetworkScanLog -Message "ERROR: Unable to retrieve service information from $($Runspace.ComputerName) ($($Runspace.IPAddress)): $($_.Exception.Message)" -MessageLevel Information
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
			$Runspaces.clone() | Where-Object { ($_.Runspace -eq $null) } | ForEach-Object {
				$Runspaces.remove($_)
				$ScanCount++
				Write-Progress -Activity 'Scanning for SQL Services' -PercentComplete (($ScanCount / $WmiDeviceCount)*100) -Status "Device $ScanCount of $WmiDeviceCount" -Id $SqlScanProgressId -ParentId $ParentProgressId
			}


		} while (($Runspaces | Where-Object {$_.Runspace -ne $null} | Measure-Object).Count -gt 0)
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