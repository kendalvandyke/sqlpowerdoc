######################
# CONSTANTS
######################
#New-Variable -Name Delimiter -Value ', ' -Scope Script -Option Constant
New-Variable -Name Delimiter -Value "`n`r" -Scope Script -Option Constant
New-Variable -Name ScanErrorThreshold -Value 3 -Scope Script -Option Constant
New-Variable -Name HotFixRegistryPath -Value 'HKCU:Software\Powershell\WindowsInventory\HotFixCache' -Scope Script -Option Constant

New-Variable -Name XlNumFmtDate -Value '[$-409]mm/dd/yyyy h:mm:ss AM/PM;@' -Scope Script -Option Constant
New-Variable -Name XlNumFmtTime -Value '[$-409]h:mm:ss AM/PM;@' -Scope Script -Option Constant
New-Variable -Name XlNumFmtText -Value '@' -Scope Script -Option Constant
New-Variable -Name XlNumFmtNumberGeneral -Value '0;@' -Scope Script -Option Constant
New-Variable -Name XlNumFmtNumberS0 -Value '#,##0;@' -Scope Script -Option Constant
New-Variable -Name XlNumFmtNumberS2 -Value '#,##0.00;@' -Scope Script -Option Constant
New-Variable -Name XlNumFmtNumberS3 -Value '#,##0.000;@' -Scope Script -Option Constant



##########################
# PRIVATE SCRIPT VARIABLES
##########################
New-Variable -Name LogQueue -Value $null -Scope Script -Visibility Private
$script:LogQueue = [System.Collections.Queue]::Synchronized((New-Object -TypeName System.Collections.Queue))


######################
# PRIVATE FUNCTIONS
######################
function Remove-ComObject { 
	[CmdletBinding()] 
	param() 
	end {
		Start-Sleep -Milliseconds 500 
		[Management.Automation.ScopedItemOptions]$scopedOpt = 'ReadOnly, Constant' 
		Get-Variable -Scope 1 | Where-Object { 
			$_.Value.pstypenames -contains 'System.__ComObject' -and -not ($scopedOpt -band $_.Options) 
		} | ForEach-Object {
			$_ | Remove-Variable -Scope 1 -Verbose:([Bool]$PSBoundParameters['Verbose'].IsPresent) 
		}
		[System.GC]::Collect() 
	} 

	<# 
 .Synopsis 
     Releases all <__ComObject> objects in the caller scope. 
 .Description 
     Releases all <__ComObject> objects in the caller scope, except for those that are Read-Only or Constant. 
 .Example 
     Remove-ComObject -Verbose 
     Description 
     =========== 
     Releases <__ComObject> objects in the caller scope and displays the released COM objects' variable names. 
.Inputs 
     None 
 .Outputs 
     None 
 .Notes 
     Name:      Remove-ComObject 
     Author:    Robert Robelo 
     LastEdit:  01/13/2010 19:14 
 .LINK
	 http://gallery.technet.microsoft.com/scriptcenter/d16d0c29-78a0-4d8d-9014-d66d57f51f63
 
 #> 
}


# Wrapper function for logging
function Write-WindowsInventoryLog {
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
		Throw
	}
}

# Wrapper function for logging
function Get-WindowsInventoryLog {
	if ((Test-Path -Path 'function:Get-LogFile') -eq $true) {
		Get-LogFile
	} else {
		Write-Output $([System.IO.Path]::ChangeExtension([IO.Path]::GetTempFileName(),'txt')).ToString()
	}
}

# Wrapper function for logging
function Get-WindowsInventoryLoggingPreference {
	if ((Test-Path -Path 'function:Get-LoggingPreference') -eq $true) {
		Get-LoggingPreference
	} else {
		Write-Output 'none'
	} 
}

# Helper function for Adding a registry path
function Add-RegistryPath {
	param(
		[Parameter(Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[string]
		$Path
	)
	process {
		$ParentPath = (Split-Path -Path $Path -Parent)
		$Leaf = (Split-Path -Path $Path -Leaf)

		if ((Test-Path -Path $ParentPath) -ne $true) {
			Add-RegistryPath -Path $ParentPath
		}

		(			New-Item -Path $Path) | Out-Null
	}
}


function Set-WindowsInventoryLogQueue {
	[CmdletBinding()]
	param(
		[Parameter(Position=0, Mandatory=$true)]
		[ValidateNotNull()]
		[System.Collections.Queue]
		$Queue
	)
	try {
		if ($Queue.IsSynchronized) {
			$script:LogQueue = $Queue
		} else {
			$script:LogQueue = [System.Collections.Queue]::Synchronized($Queue)
		}
	}
	catch {
		throw
	}
}



######################
# PUBLIC FUNCTIONS
######################


# Main function for collecting inventory information
function Get-WindowsInventory {
	<#
	.SYNOPSIS
		Collects comprehensive information about hosts running Microsoft Windows.

	.DESCRIPTION
		The Get-WindowsInventory function leverages the NetworkScan module along with Windows Management Instrumentation (WMI) to scan for and collect comprehensive information about hosts running a Windows Operating System.
		
		Get-WindowsInventory can find, verify, and collect information by Computer Name, Subnet Scan, or Active Directory DNS query.
		
		Get-WindowsInventory collects information from Windows 2000 or higher.
						
	.PARAMETER  DnsServer
		'Automatic', or the Name or IP address of an Active Directory DNS server to query for a list of hosts to inventory.
		
		When 'Automatic' is specified the function will use WMI queries to discover the current computer's DNS server(s) to query.

	.PARAMETER  DnsDomain
		'Automatic' or the Active Directory domain name to use when querying DNS for a list of hosts.
		
		When 'Automatic' is specified the function will use the current computer's AD domain.
		
		'Automatic' will be used by default if DnsServer is specified but DnsDomain is not provided.
		
	.PARAMETER  Subnet
		'Automatic' or a comma delimited list of subnets (in CIDR notation) to scan for hosts to inventory.
		
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
		A comma delimited list of computer names to inventory.
	
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
		Restrict inventory to instances on private class A, B, or C IP addresses

	.PARAMETER  AdditionalData
		A comma delimited list of additional data to collect as part of the Inventory.
		
		Valid values include: AdditionalHardware, All, BIOS, DesktopSessions, EventLog, FullyQualifiedDomainName, InstalledApplications, InstalledPatches, IPRoutes, LastLoggedOnUser, LocalGroups, LocalUserAccounts, None, PowerPlans, Printers, PrintSpoolerLocation, Processes, ProductKeys, RegistrySize, Services, Shares, StartupCommands, WindowsComponents
		
		Use "None" to bypass collecting all additional information.
		
		The default value is "All"

	.PARAMETER  ParentProgressId
		If the caller is using Write-Progress then all progress information will be written using ParentProgressId as the ParentID		

	.EXAMPLE
		Get-WindowsInventory -DNSServer automatic -DNSDomain automatic -PrivateOnly
		
		Description
		-----------
		Collect an inventory by querying Active Directory for a list of hosts to scan for Windows machines. The list of hosts will be restricted to private IP addresses only.
		
	.EXAMPLE
		Get-WindowsInventory -Subnet 172.20.40.0/28
		
		Description
		-----------
		Collect an inventory by scanning all hosts in the subnet 172.20.40.0/28 for Windows machines.
		
	.EXAMPLE
		Get-WindowsInventory -Computername Server1,Server2,Server3
		
		Description
		-----------
		Collect an inventory by scanning Server1, Server2, and Server3 for Windows machines.
		
	.EXAMPLE
		Get-WindowsInventory -Computername $env:COMPUTERNAME -AdditionalData None
		
		Description
		-----------
		Collect an inventory by scanning the local machine for Windows machines.
		
		Do not collect any data beyond the core set of information.
				

	.OUTPUTS
		System.Management.Automation.PSObject

	.NOTES

	.LINK
		Export-WindowsInventoryToExcel

#>
	[cmdletBinding(DefaultParametersetName='dns')]
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
		$Subnet = 'automatic'
		,
		[Parameter(
			Mandatory=$true,
			ParameterSetName='computername',
			HelpMessage='Computer Name(s)'
		)] 
		[alias('computer')]
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
		, [Parameter(Mandatory=$false)] 
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
		[alias('data')]
		[ValidateSet('AdditionalHardware','All','BIOS','DesktopSessions','EventLog','FullyQualifiedDomainName','InstalledApplications','InstalledPatches','IPRoutes', `
			'LastLoggedOnUser','LocalGroups','LocalUserAccounts','None','PowerPlans','Printers','PrintSpoolerLocation','Processes', `
			'ProductKeys','RegistrySize','Services','Shares','StartupCommands','WindowsComponents')]
		[string[]]
		$AdditionalData = @('All')
		,
		[Parameter(Mandatory=$false)]
		[ValidateNotNull()]
		[Int32]
		$ParentProgressId = -1
	)
	process {

		$Inventory = New-Object -TypeName psobject -Property @{ 
			Machine = @()
			StartDateUTC = [DateTime]::UtcNow
			EndDateUTC = $null
			ScanSuccessCount = 0
			ScanFailCount = 0
		} | Add-Member -MemberType ScriptProperty -Name ScanCount -Value {
			$this.ScanFailCount + $this.ScanSuccessCount
		} -PassThru

		$WmiDevice = @()
		$ParameterHash = $null
		$CurrentScanCount = 0
		$TotalScanCount = 0
		$MasterProgressId = Get-Random
		$ScanProgressId = Get-Random

		# For use with runspaces
		$ScriptBlock = $null
		$SessionState = $null
		$RunspacePool = $null
		$Runspaces = $null
		$PowerShell = $null
		$HashKey = $null

		# Fallback in case value isn't supplied or somehow missing from the environment variables
		if (-not $MaxConcurrencyThrottle) { $MaxConcurrencyThrottle = 1 }

		Write-WindowsInventoryLog -Message "Start Function: $($MyInvocation.InvocationName)" -MessageLevel Debug
		Write-Progress -Activity 'Windows Inventory' -PercentComplete 0 -Status 'Discovering Windows Machines' -Id $MasterProgressId -ParentId $ParentProgressId

		# Build command for splatting
		$ParameterHash = @{
			MaxConcurrencyThrottle = $MaxConcurrencyThrottle
			PrivateOnly = $PrivateOnly
			ParentProgressId = $MasterProgressId
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
		# Some devices may have multiple IP Addresses so only use the first WMI-capable address for each
		Find-IPv4Device @ParameterHash | Where-Object { $_.IsWmiAlive -eq $true } | Group-Object -Property WmiMachineName | ForEach-Object {
			$WmiDevice += $_.Group[0]
		}
		$TotalScanCount = $($WmiDevice | Measure-Object).Count

		Write-WindowsInventoryLog -Message 'Beginning machine scan' -MessageLevel Information
		Write-Progress -Activity 'Windows Inventory' -PercentComplete 50 -Status 'Collecting Windows Machine Information' -Id $MasterProgressId -ParentId $ParentProgressId
		Write-Progress -Activity 'Scanning Machines' -PercentComplete 0 -Status "Scanning $TotalScanCount machines" -Id $ScanProgressId -ParentId $MasterProgressId

		# Create a Session State, Create a RunspacePool, and open the RunspacePool
		$SessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
		$RunspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $MaxConcurrencyThrottle, $SessionState, $Host)
		$RunspacePool.Open()

		# Create an empty collection to hold the Runspace jobs
		$Runspaces = New-Object System.Collections.ArrayList 

		$CurrentScanCount = 0
		$ScriptBlock = {
			Param (
				[String]$LogPath,
				[String]$LoggingPreference,
				[System.Collections.Queue]$LogQueue,
				[String]$ComputerName,
				[String[]]$AdditionalData,
				[Int]$StopAtErrorCount
			)
			Get-Module -ListAvailable | Where-Object { $_.Name -ieq 'rds-manager' } | ForEach-Object {
				Import-Module -Name RDS-Manager
			}

			Import-Module -Name LogHelper, WindowsMachineInformation
			Set-LogQueue -Queue $LogQueue
			Set-LogFile -Path $LogPath
			Set-LoggingPreference -Preference $LoggingPreference
			Get-WindowsMachineInformation -AdditionalData $AdditionalData -ComputerName $ComputerName -StopAtErrorCount $StopAtErrorCount
			Remove-Module -Name LogHelper, WindowsMachineInformation

			Get-Module | Where-Object { $_.Name -ieq 'rds-manager' } | ForEach-Object {
				Remove-Module -Name RDS-Manager
			}

			[System.GC]::Collect()

		}

		# Iterate through each machine that we could make a WMI connection to and gather information
		$WmiDevice | ForEach-Object {

			$CurrentScanCount++
			Write-WindowsInventoryLog -Message "Scanning $($_.WmiMachineName) on IP address $($_.IPAddress) [Machine $CurrentScanCount of $TotalScanCount]" -MessageLevel Information

			#Create the PowerShell instance and supply the scriptblock with the other parameters
			$PowerShell = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
			$PowerShell = $PowerShell.AddArgument($(Get-WindowsInventoryLog))
			$PowerShell = $PowerShell.AddArgument($(Get-WindowsInventoryLoggingPreference))
			$PowerShell = $PowerShell.AddArgument($script:LogQueue)
			$PowerShell = $PowerShell.AddArgument($_.IPAddress)
			$PowerShell = $PowerShell.AddArgument($AdditionalData)
			$PowerShell = $PowerShell.AddArgument($ScanErrorThreshold)

			#Add the runspace into the PowerShell instance
			$PowerShell.RunspacePool = $RunspacePool

			$Runspaces.Add((
					New-Object -TypeName PsObject -Property @{
						PowerShell = $PowerShell
						Runspace = $PowerShell.BeginInvoke()
						DeviceInfo = $_
					}
				)) | Out-Null
		}


		# Reset the current scan count
		$CurrentScanCount = 0

		# Process results as they complete until they are all complete
		Do {
			$Runspaces | ForEach-Object {

				If ($_.Runspace.IsCompleted) {

					$DeviceInfo = $_.DeviceInfo

					try {

						# This is where the output gets returned
						$_.PowerShell.EndInvoke($_.Runspace) | ForEach-Object {

							# $_ Represents a scanned machine
							if ($_.ScanErrorCount -lt $ScanErrorThreshold) {
								$Inventory.Machine += $_
								$Inventory.ScanSuccessCount++
								Write-WindowsInventoryLog -Message "Scanned $($DeviceInfo.WmiMachineName) at IP address $($DeviceInfo.IPAddress) with $($_.ScanErrorCount) errors" -MessageLevel Information
							} else {
								$Inventory.ScanFailCount++
								Write-WindowsInventoryLog -Message "Failed to scan $($DeviceInfo.WmiMachineName) at IP address $($DeviceInfo.IPAddress) -  $($_.ScanErrorCount) errors" -MessageLevel Error
							}

						} 
					}
					catch {
						$Inventory.ScanFailCount++
						Write-WindowsInventoryLog -Message "An unrecoverable error was encountered while attempting to retrieve machine information from $($DeviceInfo.WmiMachineName) at IP address $($DeviceInfo.IPAddress)" -MessageLevel Error
					}
					finally {
						# Cleanup
						$_.PowerShell.dispose()
						$_.Runspace = $null
						$_.PowerShell = $null
					}
				}
			}

			# Found that in some cases an ACCESS_VIOLATION error occurs if we don't delay a little bit during each iteration 
			Start-Sleep -Milliseconds 250

			# Clean out unused runspace jobs
			$Runspaces.clone() | Where-Object { ($_.Runspace -eq $Null) } | ForEach {
				$Runspaces.remove($_)
				$CurrentScanCount++
				Write-Progress -Activity 'Scanning Machines' -PercentComplete (($CurrentScanCount / $TotalScanCount)*100) -Status "Scanned $CurrentScanCount of $TotalScanCount Machine(s)" -Id $ScanProgressId -ParentId $MasterProgressId
			}

		} while (($Runspaces | Where-Object {$_.Runspace -ne $Null} | Measure-Object).Count -gt 0)

		# Finally, close the runspaces
		$RunspacePool.close()

		# Record the scan end date
		$Inventory.EndDateUTC = [DateTime]::UtcNow

		Write-Progress -Activity 'Scanning Machines' -PercentComplete 100 -Status 'Scan Complete' -Id $ScanProgressId -ParentId $MasterProgressId -Completed
		Write-Progress -Activity 'Windows Inventory' -PercentComplete 100 -Status 'Inventory Complete' -Id $MasterProgressId -ParentId $ParentProgressId -Completed
		Write-WindowsInventoryLog -Message "Machine scan complete (Success: $($Inventory.ScanSuccessCount); Failure: $($Inventory.ScanFailCount))" -MessageLevel Information
		Write-WindowsInventoryLog -Message "End Function: $($MyInvocation.InvocationName)" -MessageLevel Debug 

		# Output the results
		Write-Output $Inventory

		Remove-Variable -Name Inventory, WmiDevice, ParameterHash, CurrentScanCount, TotalScanCount, ScanProgressId


	}
}

# Helper function to get HotFix titles
function Get-HotFixTitle {
	param(
		[Parameter(Mandatory=$true)]
		[ValidatePattern('^(KB)?\d+$')]
		[String[]]
		$HotfixId
		,
		[Parameter(Mandatory=$false)]
		[alias("Cache")]
		[switch]
		$CacheInRegistry = $false
		,
		[Parameter(Mandatory=$false)]
		[ValidateNotNull()]
		[Int32]
		$ParentProgressId = -1
	)
	process {

		$WebClient = $null
		$Url = $null
		$Title = $null
		$ProgressId = Get-Random
		$HotfixCounter = 0
		$HotFix = @{}

		Write-WindowsInventoryLog -Message "Start Function: $($MyInvocation.InvocationName)" -MessageLevel Debug
		Write-Progress -Activity 'Retrieving Hotfix Titles' -PercentComplete 0 -Status 'Beginning retrieval' -Id $ProgressId -ParentId $ParentProgressId
		Write-WindowsInventoryLog -Message 'Retrieving Hotfix Titles' -MessageLevel Information

		$HotfixId | ForEach-Object {

			$HotfixCounter++
			Write-Progress -Activity 'Retrieving Hotfix Titles' -PercentComplete (($HotfixCounter / $HotfixId.Count)*100) -Status "$($_.ToUpper()) ($HotfixCounter of $($HotfixId.Count))" -Id $ProgressId -ParentId $ParentProgressId

			$Title = $null

			# Try and get the title from the $HotFix hash first
			$Title = $HotFix[$($_.ToUpper())]

			# If not found in the $HotFix hash then try to get it from the registry cache
			if (-not $Title) {

				$Title = (Get-ItemProperty -Path $HotFixRegistryPath -Name ($_.ToUpper()) -ErrorAction SilentlyContinue).($_.ToUpper())

				# If found in the registry cache add to the $HotFix hash
				if ($Title) {
					$HotFix[($_.ToUpper())] = $Title
					Write-WindowsInventoryLog -Message "HotFix ID $($_.ToUpper()) found in registry cache" -MessageLevel Verbose
				}
			}

			# If not found in registry or the $HotFix hash then go to the intertubez and get it
			if (-not $Title) {

				if (-not $WebClient) {
					$WebClient = New-Object -TypeName System.Net.WebClient 
				} 
				$Url = [String]::Concat('http://support.microsoft.com/kb/', $_.ToUpper().Replace('KB',''))

				Write-WindowsInventoryLog -Message "HotFix ID $($_.ToUpper()) not found in cache. Retrieving title from $Url" -MessageLevel Verbose

				try {
					$Title = (($WebClient.DownloadString($Url)) | Select-String -Pattern '<title>\s*(.*)\s*</title>').Matches[0].Groups[1].Value
				}
				catch {
					# Do nothing
				}

				# If we found a title then add it to the $HotFix hash 
				if ($Title) {
					$HotFix[($_.ToUpper())] = $Title

					if ($CacheInRegistry) {
						try {
							if ((Test-Path -Path $HotFixRegistryPath) -ne $true) {
								Add-RegistryPath -Path $HotFixRegistryPath
							}

							# Cache in the registry to avoid hitting the web on future requests
							New-ItemProperty -Name ($_.ToUpper()) -Value $Title -Path $HotFixRegistryPath -PropertyType String -Force | Out-Null
						}
						catch {
							# Do nothing
						}

					} 
				}
			}
		}

		Write-Output $HotFix

		Write-WindowsInventoryLog -Message "End Function: $($MyInvocation.InvocationName)" -MessageLevel Debug
		Write-Progress -Activity 'Retrieving Hotfix Titles' -PercentComplete 100 -Status 'Retrieved $($HotfixId.Count) Hotfix titles' -Id $ProgressId -ParentId $ParentProgressId -Completed

	}
}

function Export-WindowsInventoryToExcel {
	<#
	.SYNOPSIS
		Writes an Excel file containing the information from a Windows Inventory.

	.DESCRIPTION
		The Export-WindowsInventoryToExcel function uses COM Interop to write an Excel file containing the Windows Inventory information returned by Get-WindowsInventory.
		
		Microsoft Excel 2007 or higher must be installed in order to write the Excel file.
		
	.PARAMETER  WindowsInventory
		A Windows Inventory object returned by Get-WindowsInventory.
		
	.PARAMETER  Path
		Specifies the path where the Excel file will be written. This is a fully qualified path to a .XLSX file.
		
		If not specified then the file is named "Windows Inventory - [Year][Month][Day][Hour][Minute].xlsx" and is written to your "My Documents" folder.

	.PARAMETER  ColorTheme
		An Office Theme Color to apply to each worksheet. If not specified or if an unknown theme color is provided the default "Office" theme colors will be used.
		
		Office 2013 theme colors include: Aspect, Blue Green, Blue II, Blue Warm, Blue, Grayscale, Green Yellow, Green, Marquee, Median, Office, Office 2007 - 2010, Orange Red, Orange, Paper, Red Orange, Red Violet, Red, Slipstream, Violet II, Violet, Yellow Orange, Yellow
		
		Office 2010 theme colors include: Adjacency, Angles, Apex, Apothecary, Aspect, Austin, Black Tie, Civic, Clarity, Composite, Concourse, Couture, Elemental, Equity, Essential, Executive, Flow, Foundry, Grayscale, Grid, Hardcover, Horizon, Median, Metro, Module, Newsprint, Office, Opulent, Oriel, Origin, Paper, Perspective, Pushpin, Slipstream, Solstice, Technic, Thatch, Trek, Urban, Verve, Waveform

		Office 2007 theme colors include: Apex, Aspect, Civic, Concourse, Equity, Flow, Foundry, Grayscale, Median, Metro, Module, Office, Opulent, Oriel, Origin, Paper, Solstice, Technic, Trek, Urban, Verve
		
	.PARAMETER  ColorScheme
		The color theme to apply to each worksheet. Valid values are "Light", "Medium", and "Dark". 
		
		If not specified then "Medium" is used as the default value .
		
	.PARAMETER  ParentProgressId
		If the caller is using Write-Progress then all progress information will be written using ParentProgressId as the ParentID			
		
	.EXAMPLE
		Export-WindowsInventoryToExcel -WindowsInventory $Inventory 
		
		Description
		-----------
		Write a Windows inventory using $Inventory.

		The Excel workbook will be written to your "My Documents" folder.
		
		The Office color theme and Medium color scheme will be used by default.
		
	.EXAMPLE
		Export-WindowsInventoryToExcel -WindowsInventory $Inventory -Path 'C:\Windows Inventory.xlsx'
		
		Description
		-----------
		Write a Windows inventory using $Inventory.

		The Excel workbook will be written to your C:\Windows Inventory.xlsx.
		
		The Office color theme and Medium color scheme will be used by default.

	.EXAMPLE
		Export-WindowsInventoryToExcel -WindowsInventory $Inventory -ColorTheme Blue -ColorScheme Dark
		
		Description
		-----------
		Write a Windows inventory using $Inventory.

		The Excel workbook will be written to your "My Documents" folder.
		
		The Blue color theme and Dark color scheme will be used.
	
	.NOTES
		Blue and Green are nice looking Color Themes for Office 2013

		Waveform is a nice looking Color Theme for Office 2010

	.LINK
		Get-WindowsInventory

#>
	[cmdletBinding()]
	param(
		[Parameter(Mandatory=$true, ValueFromPipeline=$true)]
		[PSCustomObject]
		$WindowsInventory
		,
		[Parameter(Mandatory=$false)] 
		[alias('File')]
		[ValidateNotNullOrEmpty()]
		[string]
		$Path = (Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)) -ChildPath ("Windows Inventory - " + (Get-Date -Format "yyyy-MM-dd-HH-mm") + ".xlsx"))
		,
		[Parameter(Mandatory=$false)] 
		[alias('Theme')]
		[string]
		$ColorTheme = 'Office'
		,
		[Parameter(Mandatory=$false)] 
		[ValidateSet('Light','Medium','Dark')]
		[string]
		$ColorScheme = 'Medium'
		,
		[Parameter(Mandatory=$false)]
		[ValidateNotNull()]
		[Int32]
		$ParentProgressId = -1
	)
	begin {
		# Add the Excel Interop assembly
		Add-Type -AssemblyName Microsoft.Office.Interop.Excel 
	}
	process {

		$Excel = New-Object -Com Excel.Application
		$Workbook = $null
		$Worksheet = $null
		$WorksheetNumber = 0
		$Range = $null
		$WorksheetData = $null
		$Row = 0
		$RowCount = 0
		$Col = 0
		$ColumnCount = 0
		$MissingType = [System.Type]::Missing
		$HotFix = @{}
		$HotFixId = @()


		$ColorThemePathPattern = $null
		$ColorThemePath = $null

		# Used to hold all of the formatting to be applied at the end
		$WorksheetCount = 23
		$WorksheetFormat = @{}


		# Excel 2010 Enumerations Reference: http://msdn.microsoft.com/en-us/library/ff838815.aspx
		$XlSortOrder = 'Microsoft.Office.Interop.Excel.XlSortOrder' -as [Type]
		$XlYesNoGuess = 'Microsoft.Office.Interop.Excel.XlYesNoGuess' -as [Type]
		$XlHAlign = 'Microsoft.Office.Interop.Excel.XlHAlign' -as [Type]
		$XlVAlign = 'Microsoft.Office.Interop.Excel.XlVAlign' -as [Type]
		$XlListObjectSourceType = 'Microsoft.Office.Interop.Excel.XlListObjectSourceType' -as [Type]
		$XlThemeColor = 'Microsoft.Office.Interop.Excel.XlThemeColor' -as [Type]

		$OverviewTabColor = $XlThemeColor::xlThemeColorLight1 # Oppose of what you'd think - this makes the tab black
		$HardwareTabColor = $XlThemeColor::xlThemeColorAccent1
		$OperatingSystemTabColor = $XlThemeColor::xlThemeColorAccent2
		$SoftwareTabColor = $XlThemeColor::xlThemeColorAccent3

		$TableStyle = switch ($ColorScheme) {
			'light' { 'TableStyleLight8' }
			'medium' { 'TableStyleMedium15' }
			'dark' { 'TableStyleDark1' }
		}

		$ProgressId = Get-Random
		$ProgressActivity = 'Export-WindowsInventoryToExcel'
		$ProgressStatus = 'Beginning output to Excel'


		Write-WindowsInventoryLog -Message "Start Function: $($MyInvocation.InvocationName)" -MessageLevel Debug
		Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Information
		Write-Progress -Activity $ProgressActivity -PercentComplete 0 -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId


		# Get Hotfix titles
		$HotFixId = ($WindowsInventory.Machine | ForEach-Object { $_.Software.Patches } | ForEach-Object { $_.HotFixID }) | Select-Object -Unique | Where-Object { $_ -imatch '^(KB)?\d+$' }
		if (($HotFixId | Measure-Object).Count -gt 0) {
			$HotFix = Get-HotFixTitle -HotfixId $HotFixId -ParentProgressId $ParentProgressId -CacheInRegistry
		}

		Write-WindowsInventoryLog -Message 'Beginning output to Excel' -MessageLevel Information

		#region

		# Write to Excel
		$Excel = New-Object -Com Excel.Application

		# Hide the Excel instance (is this necessary?)
		$Excel.visible = $false

		# Turn off screen updating
		$Excel.ScreenUpdating = $false

		# Turn off automatic calculations
		#$Excel.Calculation = [Microsoft.Office.Interop.Excel.XlCalculation]::xlCalculationManual

		# Add a workbook
		$Workbook = $Excel.Workbooks.Add()
		$Workbook.Title = 'Windows Machine Inventory'


		# Try to load the theme specified by $ColorTheme
		# The default theme is called "Office". If that's what was specified then skip over this stuff - it's already loaded
		if ($ColorTheme -ine 'office') {

			$ColorThemePathPattern = [String]::Join([System.IO.Path]::DirectorySeparatorChar, @([system.IO.Path]::GetDirectoryName($Excel.Path), 'Document Themes *','Theme Colors',[System.IO.Path]::ChangeExtension($ColorTheme,'xml')))
			$ColorThemePath = $null

			Get-ChildItem -Path $ColorThemePathPattern | ForEach-Object {
				$ColorThemePath = $_.FullName
			}

			if ($ColorThemePath) {
				$Workbook.Theme.ThemeColorScheme.Load($ColorThemePath)
			} else {
				Write-WindowsInventoryLog -Message "Unable to find a theme named ""$ColorTheme"", using default Excel theme instead" -MessageLevel Warning
			}

		}


		# Add enough worksheets to get us to 23
		$Excel.Worksheets.Add($MissingType, $Excel.Worksheets.Item($Excel.Worksheets.Count), $WorksheetCount - $Excel.Worksheets.Count, $Excel.Worksheets.Item(1).Type) | Out-Null
		$WorksheetNumber = 1

		try {

			##### Hardware:

			# Worksheet 1: Overview - Operating System, Computer Type, Chassis Type, Timezones, etc.
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Overview"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Overview'
			$Worksheet.Tab.ThemeColor = $OverviewTabColor

			$RowCount = ($WindowsInventory.Machine | Measure-Object).Count + 1
			$ColumnCount = 17
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Scan Date (UTC)'
			$WorksheetData[0,$Col++] = 'Manufacturer'
			$WorksheetData[0,$Col++] = 'Product Name'
			$WorksheetData[0,$Col++] = 'Product ID'
			$WorksheetData[0,$Col++] = 'Product Version'
			$WorksheetData[0,$Col++] = 'Operating System'
			$WorksheetData[0,$Col++] = 'Version'
			$WorksheetData[0,$Col++] = 'Service Pack'
			$WorksheetData[0,$Col++] = 'Install Date (UTC)'
			$WorksheetData[0,$Col++] = 'Domain'
			$WorksheetData[0,$Col++] = 'Role'
			$WorksheetData[0,$Col++] = 'Physical Processors'
			$WorksheetData[0,$Col++] = 'Processor Cores'
			$WorksheetData[0,$Col++] = 'Logical Processors'
			$WorksheetData[0,$Col++] = 'Physical Memory (MB)'
			$WorksheetData[0,$Col++] = 'Logical Drives'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$Col = 0
				$WorksheetData[$Row,$Col++] = $_.OperatingSystem.Settings.ComputerSystem.Name
				$WorksheetData[$Row,$Col++] = $_.ScanDateUTC
				$WorksheetData[$Row,$Col++] = $_.OperatingSystem.Settings.ComputerSystemProduct.Manufacturer
				$WorksheetData[$Row,$Col++] = $_.OperatingSystem.Settings.ComputerSystemProduct.Name
				$WorksheetData[$Row,$Col++] = $_.OperatingSystem.Settings.ComputerSystemProduct.IdentifyingNumber
				$WorksheetData[$Row,$Col++] = $_.OperatingSystem.Settings.ComputerSystemProduct.Version
				$WorksheetData[$Row,$Col++] = $_.OperatingSystem.Settings.OperatingSystem.Name
				$WorksheetData[$Row,$Col++] = $_.OperatingSystem.Settings.OperatingSystem.Version
				$WorksheetData[$Row,$Col++] = $_.OperatingSystem.Settings.OperatingSystem.ServicePack
				$WorksheetData[$Row,$Col++] = $_.OperatingSystem.Settings.OperatingSystem.InstallDateUTC
				$WorksheetData[$Row,$Col++] = $_.OperatingSystem.Settings.ComputerSystem.Domain
				$WorksheetData[$Row,$Col++] = $_.OperatingSystem.Settings.ComputerSystem.ComputerRole
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.NumberofPhysicalProcessors
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.NumberOfCores
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.NumberOfLogicalProcessors
				$WorksheetData[$Row,$Col++] = "{0:N2}" -f ($_.OperatingSystem.Settings.ComputerSystem.TotalPhysicalMemoryBytes / 1MB)
				$WorksheetData[$Row,$Col++] = (
					$_.Hardware.Storage.DiskDrive | ForEach-Object {
						$_.Partitions | ForEach-Object {
							$_.LogicalDisks | Select-Object -Property Caption, SizeBytes
						}
					} | Sort-Object -Property Caption | ForEach-Object {
						'{0} ({1:N2} GB); ' -f $_.Caption, ($_.SizeBytes / 1GB)
					}
				)

				$Row++
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'B2'
					ColumnFormat = @(
						@{ColumnNumber = 2; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtText},
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtText},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 14; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 16; NumberFormat = $XlNumFmtNumberS2}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 2: Physical Memory
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Physical Memory"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Physical Memory'
			$Worksheet.Tab.ThemeColor = $HardwareTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.Hardware.MotherboardControllerAndPort.PhysicalMemory }) | Measure-Object).Count + 1
			$ColumnCount = 12
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Bank Label'
			$WorksheetData[0,$Col++] = 'Capacity (MB)'
			$WorksheetData[0,$Col++] = 'Device Locator'
			$WorksheetData[0,$Col++] = 'Form Factor'
			$WorksheetData[0,$Col++] = 'Hot Swappable'
			$WorksheetData[0,$Col++] = 'Manufacturer'
			$WorksheetData[0,$Col++] = 'Memory Type'
			$WorksheetData[0,$Col++] = 'Part Number'
			$WorksheetData[0,$Col++] = 'Serial Number'
			$WorksheetData[0,$Col++] = 'Speed (ns)'
			$WorksheetData[0,$Col++] = 'Type Detail'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				$_.Hardware.MotherboardControllerAndPort.PhysicalMemory | Where-Object { $_.CapacityBytes } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.BankLabel
					$WorksheetData[$Row,$Col++] = "{0:N2}" -f ($_.CapacityBytes / 1MB)
					$WorksheetData[$Row,$Col++] = $_.DeviceLocator
					$WorksheetData[$Row,$Col++] = $_.FormFactor
					$WorksheetData[$Row,$Col++] = $_.HotSwappable
					$WorksheetData[$Row,$Col++] = $_.Manufacturer
					$WorksheetData[$Row,$Col++] = $_.MemoryType
					$WorksheetData[$Row,$Col++] = $_.PartNumber
					$WorksheetData[$Row,$Col++] = $_.SerialNumber
					$WorksheetData[$Row,$Col++] = $_.Speed
					$WorksheetData[$Row,$Col++] = $_.TypeDetail
					$Row++
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(4), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 3; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 11; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++ 
			#endregion


			# Worksheet 3: Processor
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Processor"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Processor'
			$Worksheet.Tab.ThemeColor = $HardwareTabColor

			$RowCount = ($WindowsInventory.Machine | Measure-Object).Count + 1
			$ColumnCount = 16
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Processor Name'
			$WorksheetData[0,$Col++] = 'Description'
			$WorksheetData[0,$Col++] = 'Family'
			$WorksheetData[0,$Col++] = 'Data Width'
			$WorksheetData[0,$Col++] = 'External Clock Frequency (MHz)'
			$WorksheetData[0,$Col++] = 'Hyperthreading?'
			$WorksheetData[0,$Col++] = 'L2 Cache Size (MB)'
			$WorksheetData[0,$Col++] = 'L2 Cache Speed (MHz)'
			$WorksheetData[0,$Col++] = 'L3 Cache Size (MB)'
			$WorksheetData[0,$Col++] = 'L3 Cache Speed (MHz)'
			$WorksheetData[0,$Col++] = 'Manufacturer'
			$WorksheetData[0,$Col++] = 'Max Clock Speed (MHz)'
			$WorksheetData[0,$Col++] = 'Cores'
			$WorksheetData[0,$Col++] = 'Logical Processors'
			$WorksheetData[0,$Col++] = 'Physical Processors'

			$Row = 1
			$WindowsInventory.Machine | Where-Object { $_.Hardware.MotherboardControllerAndPort.Processor.Name } | ForEach-Object {

				$Col = 0
				$WorksheetData[$Row,$Col++] = $_.OperatingSystem.Settings.ComputerSystem.Name
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.Name
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.Description
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.Family
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.DataWidth
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.ExtClockMHz
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.Hyperthreading
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.L2CacheSizeKB
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.L2CacheSpeedMHz
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.L3CacheSizeKB
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.L3CacheSpeedMHz
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.Manufacturer
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.MaxClockSpeedMHz
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.NumberOfCores
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.NumberOfLogicalProcessors
				$WorksheetData[$Row,$Col++] = $_.Hardware.MotherboardControllerAndPort.Processor.NumberofPhysicalProcessors
				$Row++

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 11; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 14; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 16; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 4: Network Adapaters
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Network Adapaters"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Network Adapter'
			$Worksheet.Tab.ThemeColor = $HardwareTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.Hardware.NetworkAdapter }) | Measure-Object).Count + 1
			$ColumnCount = 13
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Description'
			$WorksheetData[0,$Col++] = 'MAC Address'
			$WorksheetData[0,$Col++] = 'DNS Host Name'
			$WorksheetData[0,$Col++] = 'DHCP Enabled'
			$WorksheetData[0,$Col++] = 'DHCP Server'
			$WorksheetData[0,$Col++] = 'DNS Domain'
			$WorksheetData[0,$Col++] = 'WINS Primary'
			$WorksheetData[0,$Col++] = 'WINS Secondary'
			$WorksheetData[0,$Col++] = 'IP Address(es)'
			$WorksheetData[0,$Col++] = 'IP Subnet(s)'
			$WorksheetData[0,$Col++] = 'Default Gateway'
			$WorksheetData[0,$Col++] = 'DNS Server(s)'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				$_.Hardware.NetworkAdapter | Where-Object { $_.MACAddress } | ForEach-Object {

					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.Description
					$WorksheetData[$Row,$Col++] = $_.MACAddress
					$WorksheetData[$Row,$Col++] = $_.DNSHostName
					$WorksheetData[$Row,$Col++] = $_.DHCPEnabled
					$WorksheetData[$Row,$Col++] = $_.DHCPServer
					$WorksheetData[$Row,$Col++] = $_.DNS
					$WorksheetData[$Row,$Col++] = $_.WINSPrimaryServer
					$WorksheetData[$Row,$Col++] = $_.WINSSecondaryServer
					$WorksheetData[$Row,$Col++] = $_.IPAddress -join $Delimiter
					$WorksheetData[$Row,$Col++] = $_.IPSubnet -join $Delimiter
					$WorksheetData[$Row,$Col++] = $_.DefaultIPGateway -join $Delimiter
					$WorksheetData[$Row,$Col++] = $_.DnsServerSearchOrder -join $Delimiter
					$Row++
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 5: Printers
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Printers"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Printer'
			$Worksheet.Tab.ThemeColor = $HardwareTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.Hardware.Printer }) | Measure-Object).Count + 1
			$ColumnCount = 4
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Printer Name'
			$WorksheetData[0,$Col++] = 'Driver Name'
			$WorksheetData[0,$Col++] = 'Port Name'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				$_.Hardware.Printer | Where-Object { $_.Name } | ForEach-Object {

					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.Name
					$WorksheetData[$Row,$Col++] = $_.DriverName
					$WorksheetData[$Row,$Col++] = $_.PortName
					$Row++

				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'B2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 6: CD\DVD Drives
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): CD\DVD Drives"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'CD & DVD Drive'
			$Worksheet.Tab.ThemeColor = $HardwareTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.Hardware.Storage.CDROMDrive }) | Measure-Object).Count + 1
			$ColumnCount = 4
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Drive'
			$WorksheetData[0,$Col++] = 'Manufacturer'
			$WorksheetData[0,$Col++] = 'Name'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				$_.Hardware.Storage.CDROMDrive | Where-Object { $_.Drive } | ForEach-Object {

					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.Drive
					$WorksheetData[$Row,$Col++] = $_.Manufacturer
					$WorksheetData[$Row,$Col++] = $_.Name
					$Row++

				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 7: Disk Drives
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Disk Drives"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Disk Drive'
			$Worksheet.Tab.ThemeColor = $HardwareTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.Hardware.Storage.DiskDrive | ForEach-Object { $_.Partitions | ForEach-Object { $_.LogicalDisks } } }) | Measure-Object).Count + 1 
			$ColumnCount = 14
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Logical Disk Caption'
			$WorksheetData[0,$Col++] = 'Logical Disk File System'
			$WorksheetData[0,$Col++] = 'Logical Disk Size (GB)'
			$WorksheetData[0,$Col++] = 'Logical Disk Free Space (GB)'
			$WorksheetData[0,$Col++] = 'Logical Disk Volume Name'
			$WorksheetData[0,$Col++] = 'Drive Device ID'
			$WorksheetData[0,$Col++] = 'Drive Caption'
			$WorksheetData[0,$Col++] = 'Drive Interface Type'
			$WorksheetData[0,$Col++] = 'Drive Size (GB)'
			$WorksheetData[0,$Col++] = 'Partition Caption'
			$WorksheetData[0,$Col++] = 'Partition Starting Offset (B)'
			$WorksheetData[0,$Col++] = 'Partition Block Size (B)'
			$WorksheetData[0,$Col++] = 'Logical Disk Allocation Unit Size (B)'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				foreach ($DiskDrive in $_.Hardware.Storage.DiskDrive) {
					foreach ($Partition in $DiskDrive.Partitions) {
						$Partition.LogicalDisks | Where-Object { $_.SizeBytes } | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $MachineName
							$WorksheetData[$Row,$Col++] = $_.Caption
							$WorksheetData[$Row,$Col++] = $_.FileSystem
							$WorksheetData[$Row,$Col++] = "{0:N2}" -f ($_.SizeBytes / 1GB)
							$WorksheetData[$Row,$Col++] = "{0:N2}" -f ($_.FreeSpaceBytes / 1GB)
							$WorksheetData[$Row,$Col++] = $_.VolumeName
							$WorksheetData[$Row,$Col++] = $DiskDrive.DeviceID
							$WorksheetData[$Row,$Col++] = $DiskDrive.Caption
							$WorksheetData[$Row,$Col++] = $DiskDrive.Interfacetype
							$WorksheetData[$Row,$Col++] = "{0:N2}" -f ($DiskDrive.SizeBytes / 1GB)
							$WorksheetData[$Row,$Col++] = $Partition.Caption
							$WorksheetData[$Row,$Col++] = $Partition.StartingOffsetBytes
							$WorksheetData[$Row,$Col++] = $Partition.BlockSizeBytes
							$WorksheetData[$Row,$Col++] = $_.AllocationUnitSizeBytes
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 12; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 14; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 8: Tape Drives
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tape Drives"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tape Drive'
			$Worksheet.Tab.ThemeColor = $HardwareTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.Hardware.Storage.TapeDrive | Where-Object { $_.Name } }) | Measure-Object).Count + 1
			$ColumnCount = 4
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Name'
			$WorksheetData[0,$Col++] = 'Manufacturer'
			$WorksheetData[0,$Col++] = 'Description'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				$_.Hardware.Storage.TapeDrive | Where-Object { $_.Name } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.Name
					$WorksheetData[$Row,$Col++] = $_.Manufacturer
					$WorksheetData[$Row,$Col++] = $_.Description
					$Row++
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 9: Video Cards
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Video Cards"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Video Adapter'
			$Worksheet.Tab.ThemeColor = $HardwareTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.Hardware.VideoAndMonitor.VideoController }) | Measure-Object).Count + 1
			$ColumnCount = 4
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Name'
			$WorksheetData[0,$Col++] = 'RAM (MB)'
			$WorksheetData[0,$Col++] = 'Compatibility'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				$_.Hardware.VideoAndMonitor.VideoController | Where-Object { $_.Name } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.Name
					$WorksheetData[$Row,$Col++] = "{0:N2}" -f ($_.AdapterRAMBytes / 1MB)
					$WorksheetData[$Row,$Col++] = $_.AdapterCompatibility
					$Row++
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtNumberS2}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			##### Operating System

			#Worksheet 10: Event Logs
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Event Logs"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Event Log'
			$Worksheet.Tab.ThemeColor = $OperatingSystemTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.OperatingSystem.EventLog }) | Measure-Object).Count + 1
			$ColumnCount = 5
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Log Name'
			$WorksheetData[0,$Col++] = 'File Name'
			$WorksheetData[0,$Col++] = 'Max File Size (MB)'
			$WorksheetData[0,$Col++] = 'Overwrite Policy'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				$_.OperatingSystem.EventLog | Where-Object { $_.LogName } | ForEach-Object {

					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.LogName
					$WorksheetData[$Row,$Col++] = $_.FileName
					$WorksheetData[$Row,$Col++] = "{0:N2}" -f ($_.MaxFileSizeBytes / 1MB)
					$WorksheetData[$Row,$Col++] = $_.OverwritePolicy
					$Row++

				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtNumberS2}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			#Worksheet 11: IPv4 Routes
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): IPv4 Routes"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'IPV4 Route'
			$Worksheet.Tab.ThemeColor = $OperatingSystemTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.OperatingSystem.Network.IPV4RouteTable }) | Measure-Object).Count + 1
			$ColumnCount = 9
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Destination'
			$WorksheetData[0,$Col++] = 'Mask'
			$WorksheetData[0,$Col++] = 'Next Hop'
			$WorksheetData[0,$Col++] = 'Metric 1'
			$WorksheetData[0,$Col++] = 'Metric 2'
			$WorksheetData[0,$Col++] = 'Metric 3'
			$WorksheetData[0,$Col++] = 'Metric 4'
			$WorksheetData[0,$Col++] = 'Metric 5'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				$_.OperatingSystem.Network.IPV4RouteTable | Where-Object { $_.Destination } | ForEach-Object {

					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.Destination
					$WorksheetData[$Row,$Col++] = $_.Mask
					$WorksheetData[$Row,$Col++] = $_.NextHop
					$WorksheetData[$Row,$Col++] = $_.Metric1
					$WorksheetData[$Row,$Col++] = $_.Metric2
					$WorksheetData[$Row,$Col++] = $_.Metric3
					$WorksheetData[$Row,$Col++] = $_.Metric4
					$WorksheetData[$Row,$Col++] = $_.Metric5
					$Row++

				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null 

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'D2'
					ColumnFormat = @(
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			#Worksheet 12: Page Files
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Page Files"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Page File'
			$Worksheet.Tab.ThemeColor = $OperatingSystemTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.OperatingSystem.PageFile }) | Measure-Object).Count + 1
			$ColumnCount = 8
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Drive'
			$WorksheetData[0,$Col++] = 'Initial Size (MB)'
			$WorksheetData[0,$Col++] = 'Max Size (MB)'
			$WorksheetData[0,$Col++] = 'Current Size (MB)'
			$WorksheetData[0,$Col++] = 'Peak Size (MB)'
			$WorksheetData[0,$Col++] = 'Temporary'
			$WorksheetData[0,$Col++] = 'Auto Managed'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				$_.OperatingSystem.PageFile | Where-Object { $_.Drive } | ForEach-Object {

					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.Drive
					$WorksheetData[$Row,$Col++] = $_.InitialSizeMB
					$WorksheetData[$Row,$Col++] = $_.MaximumSizeMB
					$WorksheetData[$Row,$Col++] = $_.CurrentSizeMB
					$WorksheetData[$Row,$Col++] = $_.PeakSizeMB
					$WorksheetData[$Row,$Col++] = $_.IsTemporary # Sometimes this is coming back as blank (unknown?)
					$WorksheetData[$Row,$Col++] = $_.IsAutoManaged
					$Row++

				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 3; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			#Worksheet 13: Registry
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Registry"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Registry'
			$Worksheet.Tab.ThemeColor = $OperatingSystemTabColor

			$RowCount = ($WindowsInventory.Machine | Measure-Object).Count + 1
			$ColumnCount = 3
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Current Size (MB)'
			$WorksheetData[0,$Col++] = 'Maximum Size (MB)'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$Col = 0
				$WorksheetData[$Row,$Col++] = $_.OperatingSystem.Settings.ComputerSystem.Name
				$WorksheetData[$Row,$Col++] = $_.OperatingSystem.Registry.CurrentSizeMB
				$WorksheetData[$Row,$Col++] = $_.OperatingSystem.Registry.MaximumSizeMB
				$Row++

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'B2'
					ColumnFormat = @(
						@{ColumnNumber = 2; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 3; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			#Worksheet 14: Power Plans
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Power Plans"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Power Plans'
			$Worksheet.Tab.ThemeColor = $OperatingSystemTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.OperatingSystem.Settings.PowerPlan }) | Measure-Object).Count + 1
			$ColumnCount = 5
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Operating System'
			$WorksheetData[0,$Col++] = 'Plan Name'
			$WorksheetData[0,$Col++] = 'Active'
			$WorksheetData[0,$Col++] = 'Description'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name
				$OperatingSystem = $_.OperatingSystem.Settings.OperatingSystem.Name

				$_.OperatingSystem.Settings.PowerPlan | Where-Object { $_.PlanName } | ForEach-Object {

					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $OperatingSystem
					$WorksheetData[$Row,$Col++] = $_.PlanName
					$WorksheetData[$Row,$Col++] = $_.IsActive
					$WorksheetData[$Row,$Col++] = $_.Description
					$Row++

				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'D2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			#Worksheet 15: Startup Commands
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Startup Commmands"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Startup Commands'
			$Worksheet.Tab.ThemeColor = $OperatingSystemTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.OperatingSystem.Settings.StartupCommands }) | Measure-Object).Count + 1
			$ColumnCount = 4
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Name'
			$WorksheetData[0,$Col++] = 'Command'
			$WorksheetData[0,$Col++] = 'User'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				$_.OperatingSystem.Settings.StartupCommands | Where-Object { $_.Name } | ForEach-Object {

					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.Name
					$WorksheetData[$Row,$Col++] = $_.Command
					$WorksheetData[$Row,$Col++] = $_.User
					$Row++

				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(4), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			#Worksheet 16: Running Processes
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Running Processes"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Processes'
			$Worksheet.Tab.ThemeColor = $OperatingSystemTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.OperatingSystem.RunningProcesses }) | Measure-Object).Count + 1
			$ColumnCount = 5
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Process Name'
			$WorksheetData[0,$Col++] = 'Executable Path'
			$WorksheetData[0,$Col++] = 'Domain'
			$WorksheetData[0,$Col++] = 'Username'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				$_.OperatingSystem.RunningProcesses | Where-Object { $_.Caption } | ForEach-Object {

					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.Caption
					$WorksheetData[$Row,$Col++] = $_.ExecutablePath
					$WorksheetData[$Row,$Col++] = $_.OwnerDomain
					$WorksheetData[$Row,$Col++] = $_.OwnerUser
					$Row++

				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			#Worksheet 17: Services
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Services"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Services'
			$Worksheet.Tab.ThemeColor = $OperatingSystemTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.OperatingSystem.Services }) | Measure-Object).Count + 1
			$ColumnCount = 6
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Caption'
			$WorksheetData[0,$Col++] = 'Path'
			$WorksheetData[0,$Col++] = 'Started'
			$WorksheetData[0,$Col++] = 'Start Mode'
			$WorksheetData[0,$Col++] = 'Start As'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				$_.OperatingSystem.Services | Where-Object { $_.Caption } | ForEach-Object {

					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.Caption
					$WorksheetData[$Row,$Col++] = $_.PathName
					$WorksheetData[$Row,$Col++] = $_.Started
					$WorksheetData[$Row,$Col++] = $_.StartMode
					$WorksheetData[$Row,$Col++] = $_.StartAs
					$Row++

				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			#Worksheet 18: Shares
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Shares"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Shares'
			$Worksheet.Tab.ThemeColor = $OperatingSystemTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.OperatingSystem.Shares }) | Measure-Object).Count + 1
			$ColumnCount = 5
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Type'
			$WorksheetData[0,$Col++] = 'Share Name'
			$WorksheetData[0,$Col++] = 'Path'
			$WorksheetData[0,$Col++] = 'Description'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				$_.OperatingSystem.Shares | Where-Object { $_.ShareType } | ForEach-Object {

					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.ShareType
					$WorksheetData[$Row,$Col++] = $_.Name
					$WorksheetData[$Row,$Col++] = $_.Path
					$WorksheetData[$Row,$Col++] = $_.Description
					$Row++

				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'D2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			#Worksheet 19: Local Groups
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Local Groups"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Local Groups'
			$Worksheet.Tab.ThemeColor = $OperatingSystemTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.OperatingSystem.Users.LocalGroups }) | Measure-Object).Count + 1
			$ColumnCount = 3
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Group Name'
			$WorksheetData[0,$Col++] = 'User Name'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name
				$GroupName = $null

				#$_.OperatingSystem.Users.LocalGroups | Sort-Object -Property Name | ForEach-Object {
				$_.OperatingSystem.Users.LocalGroups | Where-Object { $_.Name } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.Name
					$WorksheetData[$Row,$Col++] = $_.Members -join $Delimiter
					$Row++
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			#Worksheet 20: Local Users
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Local Users"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Local Users'
			$Worksheet.Tab.ThemeColor = $OperatingSystemTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.OperatingSystem.Users.LocalUsers }) | Measure-Object).Count + 1
			$ColumnCount = 9
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Username'
			$WorksheetData[0,$Col++] = 'Full Name'
			$WorksheetData[0,$Col++] = 'Disabled'
			$WorksheetData[0,$Col++] = 'Locked Out'
			$WorksheetData[0,$Col++] = 'Password Changeable'
			$WorksheetData[0,$Col++] = 'Password Expires'
			$WorksheetData[0,$Col++] = 'Password Required'
			$WorksheetData[0,$Col++] = 'Description'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				$_.OperatingSystem.Users.LocalUsers | Where-Object { $_.UserName } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.UserName
					$WorksheetData[$Row,$Col++] = $_.FullName
					$WorksheetData[$Row,$Col++] = $_.Disabled
					$WorksheetData[$Row,$Col++] = $_.Lockout
					$WorksheetData[$Row,$Col++] = $_.PasswordChangeable
					$WorksheetData[$Row,$Col++] = $_.PasswordExpires
					$WorksheetData[$Row,$Col++] = $_.PasswordRequired
					$WorksheetData[$Row,$Col++] = $_.Description
					$Row++
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			#Worksheet 21: Desktop Sessions
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Desktop Sessions"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Desktop Sessions'
			$Worksheet.Tab.ThemeColor = $OperatingSystemTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.OperatingSystem.Users.DesktopSessions }) | Measure-Object).Count + 1
			$ColumnCount = 10
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'User'
			$WorksheetData[0,$Col++] = 'Logon Date (UTC)'
			$WorksheetData[0,$Col++] = 'Protocol'
			$WorksheetData[0,$Col++] = 'Host'
			$WorksheetData[0,$Col++] = 'State'
			$WorksheetData[0,$Col++] = 'Idle Time'
			$WorksheetData[0,$Col++] = 'Client'
			$WorksheetData[0,$Col++] = 'Session'
			$WorksheetData[0,$Col++] = 'SessionID'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				$_.OperatingSystem.Users.DesktopSessions | Where-Object { $_.User } | ForEach-Object {

					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.User
					$WorksheetData[$Row,$Col++] = $_.LogonTimeUTC
					$WorksheetData[$Row,$Col++] = $_.ProtocolType
					$WorksheetData[$Row,$Col++] = $_.Host
					$WorksheetData[$Row,$Col++] = $_.State
					$WorksheetData[$Row,$Col++] = $_.IdleTime
					$WorksheetData[$Row,$Col++] = $_.Client
					$WorksheetData[$Row,$Col++] = $_.Session
					$WorksheetData[$Row,$Col++] = $_.SessionID
					$Row++

				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 3; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion



			##### Software
			#Worksheet 22: Installed Applications
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Installed Applications"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Applications'
			$Worksheet.Tab.ThemeColor = $SoftwareTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.Software.InstalledApplications }) | Measure-Object).Count + 1
			$ColumnCount = 5
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'Product Name'
			$WorksheetData[0,$Col++] = 'Vendor'
			$WorksheetData[0,$Col++] = 'Version'
			$WorksheetData[0,$Col++] = 'Install Date (UTC)'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				$_.Software.InstalledApplications | Where-Object { $_.ProductName } | ForEach-Object {

					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.ProductName
					$WorksheetData[$Row,$Col++] = $_.Vendor
					$WorksheetData[$Row,$Col++] = $_.Version
					$WorksheetData[$Row,$Col++] = $_.InstallDateUTC
					$Row++

				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			#Worksheet 23: Patches
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Patches"
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Patches'
			$Worksheet.Tab.ThemeColor = $SoftwareTabColor

			$RowCount = (($WindowsInventory.Machine | ForEach-Object { $_.Software.Patches }) | Measure-Object).Count + 1
			$ColumnCount = 6
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Machine Name'
			$WorksheetData[0,$Col++] = 'HotFix ID'
			$WorksheetData[0,$Col++] = 'Install Date (UTC)'
			$WorksheetData[0,$Col++] = 'Installed By'
			$WorksheetData[0,$Col++] = 'Type'
			$WorksheetData[0,$Col++] = 'Title'

			$Row = 1
			$WindowsInventory.Machine | ForEach-Object {

				$MachineName = $_.OperatingSystem.Settings.ComputerSystem.Name

				$_.Software.Patches | Where-Object { $_.HotFixID } | ForEach-Object {

					$Col = 0
					$WorksheetData[$Row,$Col++] = $MachineName
					$WorksheetData[$Row,$Col++] = $_.HotFixID
					$WorksheetData[$Row,$Col++] = $_.InstallDateUTC
					$WorksheetData[$Row,$Col++] = $_.InstalledBy
					$WorksheetData[$Row,$Col++] = $_.Description

					if ($_.HotFixID -imatch '^(KB)?\d+$') {
						$WorksheetData[$Row,$Col++] = $HotFix[($_.HotFixID).ToUpper()]
					} else {
						$WorksheetData[$Row,$Col++] = [String]::Empty
					} 

					$Row++
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 2; NumberFormat = $XlNumFmtText}
						@{ColumnNumber = 3; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Apply formatting to every worksheet
			# Work backwards so that the first sheet is active when the workbook is saved
			$ProgressStatus = 'Applying formatting to all worksheets'
			Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			for ($WorksheetNumber = $WorksheetCount; $WorksheetNumber -ge 1; $WorksheetNumber--) {

				$ProgressStatus = "Applying formatting to Worksheet #$($WorksheetNumber)"
				Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Verbose
				Write-Progress -Activity $ProgressActivity -PercentComplete (((($WorksheetCount * 2) - $WorksheetNumber + 1) / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId

				$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)

				# Switch to the worksheet
				$Worksheet.Activate() | Out-Null

				# Bold the header row
				$Duration = (Measure-Command {
						$Worksheet.Rows.Item(1).Font.Bold = $WorksheetFormat[$WorksheetNumber].BoldFirstRow
					}).TotalMilliseconds
				Write-WindowsInventoryLog -Message "Bold Header Row Duration (ms): $Duration" -MessageLevel Debug

				# Bold the 1st column
				$Duration = (Measure-Command {
						$Worksheet.Columns.Item(1).Font.Bold = $WorksheetFormat[$WorksheetNumber].BoldFirstColumn
					}).TotalMilliseconds
				Write-WindowsInventoryLog -Message "Bold 1st Column Duration (ms): $Duration" -MessageLevel Debug

				# Freeze View
				$Duration = (Measure-Command {
						$Worksheet.Range($WorksheetFormat[$WorksheetNumber].FreezeAtCell).Select() | Out-Null
						$Worksheet.Application.ActiveWindow.FreezePanes = $true 
					}).TotalMilliseconds
				Write-WindowsInventoryLog -Message "Freeze View Duration (ms): $Duration" -MessageLevel Debug


				# Apply Column formatting
				$Duration = (Measure-Command {
						$WorksheetFormat[$WorksheetNumber].ColumnFormat | ForEach-Object {
							$Worksheet.Columns.Item($_.ColumnNumber).NumberFormat = $_.NumberFormat
						}
					}).TotalMilliseconds
				Write-WindowsInventoryLog -Message "Apply Column formatting Duration (ms): $Duration" -MessageLevel Debug

				# Apply Row formatting
				$Duration = (Measure-Command {
						$WorksheetFormat[$WorksheetNumber].RowFormat | ForEach-Object {
							$Worksheet.Rows.Item($_.RowNumber).NumberFormat = $_.NumberFormat
						}
					}).TotalMilliseconds
				Write-WindowsInventoryLog -Message "Apply Row formatting Duration (ms): $Duration" -MessageLevel Debug

				# Update worksheet values so row and column formatting apply
				$Duration = (Measure-Command {
						try {
							$Worksheet.UsedRange.Value2 = $Worksheet.UsedRange.Value2
						} catch {
							# Sometimes trying to set the entire worksheet's value to itself will result in the following exception:
							# 	"Not enough storage is available to complete this operation. 0x8007000E (E_OUTOFMEMORY))"
							# See http://support.microsoft.com/kb/313275 for more information
							# When this happens the workaround is to try doing the work in smaller chunks
							# ...so we'll try to update the column\row values that have specific formatting one at a time instead of the entire worksheet at once
							$WorksheetFormat[$WorksheetNumber].ColumnFormat | ForEach-Object {
								$Worksheet.Columns.Item($_.ColumnNumber).Value2 = $Worksheet.Columns.Item($_.ColumnNumber).Value2
							}
							$WorksheetFormat[$WorksheetNumber].RowFormat | ForEach-Object {
								$Worksheet.Rows.Item($_.RowNumber).Value2 = $Worksheet.Rows.Item($_.RowNumber).Value2
							}
						}
					}).TotalMilliseconds
				Write-WindowsInventoryLog -Message "Apply Row and Column formatting - Update Values (ms): $Duration" -MessageLevel Debug


				# Apply table formatting
				$Duration = (Measure-Command {
						$ListObject = $Worksheet.ListObjects.Add($XlListObjectSourceType::xlSrcRange, $Worksheet.UsedRange, $null, $XlYesNoGuess::xlYes, $null) 
						$ListObject.Name = "Table $WorksheetNumber"
						$ListObject.TableStyle = $TableStyle
						$ListObject.ShowTableStyleFirstColumn = $WorksheetFormat[$WorksheetNumber].BoldFirstColumn # Put a background color behind the 1st column
						$ListObject.ShowAutoFilter = $WorksheetFormat[$WorksheetNumber].AutoFilter
					}).TotalMilliseconds
				Write-WindowsInventoryLog -Message "Apply table formatting Duration (ms): $Duration" -MessageLevel Debug

				# Zoom back to 80%
				$Duration = (Measure-Command {
						$Worksheet.Application.ActiveWindow.Zoom = 80
					}).TotalMilliseconds
				Write-WindowsInventoryLog -Message "Zoom to 80% Duration (ms): $Duration" -MessageLevel Debug

				# Adjust the column widths to 250 before autofitting contents
				# This allows longer lines of text to remain on one line
				$Duration = (Measure-Command {
						$Worksheet.UsedRange.EntireColumn.ColumnWidth = 250
					}).TotalMilliseconds
				Write-WindowsInventoryLog -Message "Change column width Duration (ms): $Duration" -MessageLevel Debug

				# Wrap text
				$Duration = (Measure-Command {
						$Worksheet.UsedRange.WrapText = $true
					}).TotalMilliseconds
				Write-WindowsInventoryLog -Message "Wrap text Duration (ms): $Duration" -MessageLevel Debug

				# Autofit column and row contents
				$Duration = (Measure-Command {
						$Worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null
						$Worksheet.UsedRange.EntireRow.AutoFit() | Out-Null
					}).TotalMilliseconds
				Write-WindowsInventoryLog -Message "Autofit contents Duration (ms): $Duration" -MessageLevel Debug

				# Left align contents
				$Duration = (Measure-Command {
						$Worksheet.UsedRange.EntireColumn.HorizontalAlignment = $XlHAlign::xlHAlignLeft
					}).TotalMilliseconds
				Write-WindowsInventoryLog -Message "Left align contents Duration (ms): $Duration" -MessageLevel Debug

				# Vertical align contents
				$Duration = (Measure-Command {
						$Worksheet.UsedRange.EntireColumn.VerticalAlignment = $XlVAlign::xlVAlignTop
					}).TotalMilliseconds
				Write-WindowsInventoryLog -Message "Vertical align contents Duration (ms): $Duration" -MessageLevel Debug

				# Put the selection back to the upper left cell
				$Duration = (Measure-Command {
						$Worksheet.Range('A1').Select() | Out-Null
					}).TotalMilliseconds
				Write-WindowsInventoryLog -Message "Reset selection Duration (ms): $Duration" -MessageLevel Debug

			}

		}
		catch {
			throw
		}
		finally {
			# Save and quit Excel
			$Worksheet.Application.DisplayAlerts = $false
			$Workbook.SaveAs($Path)
			$Workbook.Saved = $true

			# Turn on screen updating
			$Excel.ScreenUpdating = $true

			# Turn on automatic calculations
			#$Excel.Calculation = [Microsoft.Office.Interop.Excel.XlCalculation]::xlCalculationAutomatic

			$Excel.Quit()
		}

		#endregion


		$ProgressStatus = 'Output to Excel complete'
		Write-Progress -Activity $ProgressActivity -PercentComplete 100 -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId -Completed
		Write-WindowsInventoryLog -Message $ProgressStatus -MessageLevel Information
		Write-WindowsInventoryLog -Message "End Function: $($MyInvocation.InvocationName)" -MessageLevel Debug


		Remove-Variable -Name Excel
		Remove-Variable -Name Workbook
		Remove-Variable -Name Worksheet
		Remove-Variable -Name WorksheetNumber
		Remove-Variable -Name Range
		Remove-Variable -Name WorksheetData
		Remove-Variable -Name Row
		Remove-Variable -Name RowCount
		Remove-Variable -Name Col
		Remove-Variable -Name ColumnCount
		Remove-Variable -Name MissingType
		Remove-Variable -Name HotFix
		Remove-Variable -Name HotFixId

		Remove-Variable -Name ColorThemePathPattern
		Remove-Variable -Name ColorThemePath

		Remove-Variable -Name WorksheetFormat

		Remove-Variable -Name XlSortOrder
		Remove-Variable -Name XlYesNoGuess
		Remove-Variable -Name XlListObjectSourceType
		Remove-Variable -Name XlThemeColor

		Remove-Variable -Name OverviewTabColor
		Remove-Variable -Name HardwareTabColor
		Remove-Variable -Name OperatingSystemTabColor
		Remove-Variable -Name SoftwareTabColor

		Remove-Variable -Name TableStyle


		# Release all lingering COM objects
		Remove-ComObject

	}
}
