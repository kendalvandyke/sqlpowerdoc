<#
TODO:
- Support for converting UTC times to local timezone?
#>


######################
# CONSTANTS
######################
#New-Variable -Name Delimiter -Value ', ' -Scope Script -Option Constant
New-Variable -Name Delimiter -Value "`n`r" -Scope Script -Option Constant
New-Variable -Name ScanErrorThreshold -Value 3 -Scope Script -Option Constant

New-Variable -Name HighPriority -Value 10 -Scope Script -Option Constant
New-Variable -Name MediumPriority -Value 20 -Scope Script -Option Constant
New-Variable -Name LowPriority -Value 30 -Scope Script -Option Constant
New-Variable -Name NoPriority -Value 40 -Scope Script -Option Constant

New-Variable -Name CatPerformance -Value 'Performance' -Scope Script -Option Constant
New-Variable -Name CatReliability -Value 'Reliability' -Scope Script -Option Constant
New-Variable -Name CatSecurity -Value 'Security' -Scope Script -Option Constant
New-Variable -Name CatAvailability -Value 'Availability' -Scope Script -Option Constant
New-Variable -Name CatRecovery -Value 'Recovery' -Scope Script -Option Constant
New-Variable -Name CatInformation -Value 'Information' -Scope Script -Option Constant

New-Variable -Name XlNumFmtDate -Value '[$-409]mm/dd/yyyy h:mm:ss AM/PM;@' -Scope Script -Option Constant
New-Variable -Name XlNumFmtTime -Value '[$-409]h:mm:ss AM/PM;@' -Scope Script -Option Constant
New-Variable -Name XlNumFmtText -Value '@' -Scope Script -Option Constant
New-Variable -Name XlNumFmtNumberGeneral -Value '0;@' -Scope Script -Option Constant
New-Variable -Name XlNumFmtNumberS0 -Value '#,##0;@' -Scope Script -Option Constant
New-Variable -Name XlNumFmtNumberS2 -Value '#,##0.00;@' -Scope Script -Option Constant
New-Variable -Name XlNumFmtNumberS3 -Value '#,##0.000;@' -Scope Script -Option Constant


######################
# SCRIPT VARIABLES
######################
New-Object -TypeName System.Version -ArgumentList '1.0.0.0' | New-Variable -Name ModuleVersion -Scope Script -Option Constant -Visibility Private

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
		[GC]::Collect() 
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


# Based on http://poshcode.org/4050
function ConvertTo-GzCliXml {
	param(
		[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		[ValidateNotNullOrEmpty()]
		[PSObject[]]$InputObject
	)
	begin {
		$type = [System.Management.Automation.PSObject].Assembly.GetType('System.Management.Automation.Serializer')
		$ctor = $type.GetConstructor('instance,nonpublic', $null, @([System.Xml.XmlWriter]), $null)
		$MemoryStream = New-Object -TypeName System.IO.MemoryStream
		$GZipStream = New-Object -TypeName System.IO.Compression.GZipStream($MemoryStream, [System.IO.Compression.CompressionMode]::Compress, $true)
		$BufferedStream = New-Object -TypeName System.IO.BufferedStream($GZipStream, 8192)
		$xw = [System.Xml.XmlTextWriter]::Create($BufferedStream)
		$serializer = $ctor.Invoke($xw)
		$method = $type.GetMethod('Serialize', 'nonpublic,instance', $null, [type[]]@([object]), $null)
		$done = $type.GetMethod('Done', [System.Reflection.BindingFlags]'nonpublic,instance')
	}
	process {
		try {
			[void]$method.Invoke($serializer, $InputObject)
		} catch {
			write-warning "Could not serialize $($InputObject.gettype()): $_"
		}
	}
	end {
		[void]$done.invoke($serializer, @())
		$xw.Close()
		$BufferedStream.Flush()
		$BufferedStream.Dispose()
		$GZipStream.Dispose()

		# ToArray() creates a copy of the data in memory, but...
		# GetBuffer() can be double the length (or more) of ToArray() as $InputObject gets larger
		# See http://msdn.microsoft.com/en-us/library/system.io.memorystream.getbuffer.aspx
		#$MemoryStream.GetBuffer()
		$MemoryStream.ToArray()

		$MemoryStream.Dispose()

		# Cleanup
		Remove-Variable -Name type, ctor, MemoryStream, GZipStream, xw, serializer, method, done 
	}
}

# Based on http://poshcode.org/4051
function ConvertFrom-GzCliXml {
	param(
		[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		[ValidateNotNullOrEmpty()]
		[Byte[]]$InputObject
	)
	begin
	{
	}
	process
	{
	}
	end
	{
		$type = [System.Management.Automation.PSObject].Assembly.GetType('System.Management.Automation.Deserializer')
		$ctor = $type.GetConstructor('instance,nonpublic', $null, @([xml.xmlreader]), $null)
		$MemoryStream = New-Object -TypeName System.IO.MemoryStream -ArgumentList (,$InputObject)
		$GZipStream = New-Object -TypeName System.IO.Compression.GZipStream($MemoryStream, [System.IO.Compression.CompressionMode]::Decompress, $false)
		$xr = [System.Xml.XmlTextReader]::Create($GZipStream)
		$deserializer = $ctor.Invoke($xr)
		$method = @($type.GetMethods('nonpublic,instance') | Where-Object {$_.Name -like "Deserialize"})[1]
		$done = $type.GetMethod('Done', [System.Reflection.BindingFlags]'nonpublic,instance')
		while (!$done.Invoke($deserializer, @()))
		{
			try {
				$method.Invoke($deserializer, "")
			} catch {
				Write-Warning "Could not deserialize ${string}: $_"
			}
		}
		$xr.Close()
		$GZipStream.Dispose()
		$MemoryStream.Dispose()

		# Cleanup
		Remove-Variable -Name type, ctor, MemoryStream, GZipStream, xr, deserializer, method, done 
	}
}

# Based on http://poshcode.org/4050
function Export-GzCliXml {
	param(
		[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		[ValidateNotNullOrEmpty()]
		[PSObject[]]$InputObject
		,
		[Parameter(Position=1, Mandatory=$true, ValueFromPipeline=$false)]
		[ValidateNotNullOrEmpty()]
		[String]$Path
	)
	begin {
		$type = [System.Management.Automation.PSObject].Assembly.GetType('System.Management.Automation.Serializer')
		$ctor = $type.GetConstructor('instance,nonpublic', $null, @([System.Xml.XmlWriter]), $null)
		$FileStream = New-Object -Typename System.IO.FileStream($Path, [System.IO.FileMode]::Create)
		$GZipStream = New-Object -TypeName System.IO.Compression.GZipStream($FileStream, [System.IO.Compression.CompressionMode]::Compress, $true)
		$BufferedStream = New-Object -TypeName System.IO.BufferedStream($GZipStream, 8192)
		$xw = [System.Xml.XmlTextWriter]::Create($BufferedStream)
		$serializer = $ctor.Invoke($xw)
		$method = $type.GetMethod('Serialize', 'nonpublic,instance', $null, [type[]]@([object]), $null)
		$done = $type.GetMethod('Done', [System.Reflection.BindingFlags]'nonpublic,instance')
	}
	process {
		try {
			[void]$method.Invoke($serializer, $InputObject)
		} catch {
			write-warning "Could not serialize $($InputObject.gettype()): $_"
		}
	}
	end {
		[void]$done.invoke($serializer, @())
		$xw.Close()
		$BufferedStream.Flush()
		$BufferedStream.Dispose()
		$GZipStream.Dispose()
		$FileStream.Dispose()

		# Cleanup
		Remove-Variable -Name type, ctor, FileStream, GZipStream, xw, serializer, method, done 
	}
}

# Based on http://poshcode.org/4051
function Import-GzCliXml {
	param(
		[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		[ValidateNotNullOrEmpty()]
		[String[]]$Path
	)
	begin
	{
		$type = [System.Management.Automation.PSObject].Assembly.GetType('System.Management.Automation.Deserializer')
		$ctor = $type.GetConstructor('instance,nonpublic', $null, @([xml.xmlreader]), $null)
	}
	process
	{
		foreach ($InputObject in $Path) {
			$FileStream = New-Object -Typename System.IO.FileStream($InputObject, [System.IO.FileMode]::Open)
			$GZipStream = New-Object -TypeName System.IO.Compression.GZipStream($FileStream, [System.IO.Compression.CompressionMode]::Decompress, $false)
			$xr = [System.Xml.XmlTextReader]::Create($GZipStream)
			$deserializer = $ctor.Invoke($xr)
			$method = @($type.GetMethods('nonpublic,instance') | Where-Object {$_.Name -like "Deserialize"})[1]
			$done = $type.GetMethod('Done', [System.Reflection.BindingFlags]'nonpublic,instance')
			while (!$done.Invoke($deserializer, @()))
			{
				try {
					$method.Invoke($deserializer, "")
				} catch {
					Write-Warning "Could not deserialize ${string}: $_"
				}
			}
			$xr.Close()
			$GZipStream.Dispose()
			$FileStream.Dispose()
		}
	}
	end
	{
		# Cleanup
		Remove-Variable -Name type, ctor, FileStream, GZipStream, xr, deserializer, method, done 
	}
}



# Wrapper function for logging
function Write-SqlServerInventoryLog {
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
function Get-SqlServerInventoryLog {
	if ((Test-Path -Path 'function:Get-LogFile') -eq $true) {
		Get-LogFile
	} else {
		Write-Output "$env:temp\PowerShellLog.txt"
	}
}

# Wrapper function for logging
function Get-SqlServerInventoryLoggingPreference {
	if ((Test-Path -Path 'function:Get-LoggingPreference') -eq $true) {
		Get-LoggingPreference
	} else {
		Write-Output 'none'
	} 
}

# Wrapper function for logging to support runspaces
function Set-SqlServerInventoryLogQueue {
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

# Recursively traverse Server configuration information to find sys.configurations values
function Get-ServerConfigurationItem([PSObject]$ServerConfigurationInformation) {

	$ServerConfigurationInformation | Where-Object {
		(			$_.PSObject.Properties | Where-Object {
				@('DefaultValue','RunningValue','ConfiguredValue','ConfigurationName') -icontains $_.Name
			} | Measure-Object).Count -eq 4
	} | ForEach-Object {
		$ServerConfigurationInformation
	}

	$ServerConfigurationInformation.PSObject.Properties | Where-Object { $_.MemberType -ieq 'NoteProperty' } | ForEach-Object {
		Get-ServerConfigurationItem -ServerConfigurationInformation $_.Value
	}

}

function Get-PriorityValue {
	[CmdletBinding()]
	param(
		[Parameter(Position=0, Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[Int]
		$Priority
	)
	try {
		Write-Output $(
			switch ($Priority) {
				$HighPriority { 'High' }
				$MediumPriority { 'Medium' }
				$LowPriority { 'Low' }
				$NoPriority { 'None' }
				default { 'Unknown' }
			}
		)
	}
	catch {
		Throw
	}
}

function Test-InventoryIsCompressed {
	[CmdletBinding()]
	param(
		[Parameter(Position=0, Mandatory=$true)]
		[AllowNull()]
		[Object]$InputObject
	)
	begin {
	}
	process {

		if (
			$InputObject.GetType().BaseType.Name -ieq 'Array' -and
			$InputObject[0].GetType().Name -ieq 'Byte'
		) {
			Write-Output $true
		} else {
			Write-Output $false
		}
	}
	end {

	}
}

function Get-AssessmentFinding {
	[CmdletBinding()]
	[OutputType([PsObject])]
	Param
	(
		[Parameter(Mandatory=$true)]
		[String]
		$ServerName
		,
		[Parameter(Mandatory=$false)]
		[AllowNull()]
		[String]
		$DatabaseName
		,
		[Parameter(Mandatory=$true)]
		[Int]
		$Priority
		,
		[Parameter(Mandatory=$true)]
		[String]
		$Category
		,
		[Parameter(Mandatory=$true)]
		[String]
		$Description
		,
		[Parameter(Mandatory=$true)]
		[String]
		$Details
		,
		[Parameter(Mandatory=$false)]
		[AllowNull()]
		[String]
		$URL
	)
	Process
	{
		New-Object -TypeName PSObject -Property @{
			ServerName = $ServerName
			DatabaseName = $DatabaseName
			Priority = $Priority
			PriorityValue = [String](Get-PriorityValue -Priority $Priority)
			Category = $Category
			Description = $Description
			Details = $Details
			URL = $URL
		}
	}
}


######################
# PUBLIC FUNCTIONS
######################

function Export-SqlServerInventoryToGzClixml {
	<#
	.SYNOPSIS
		Writes a GZip compressed representation of a SQL Server Inventory returned by Get-SqlServerInventory to disk.

	.DESCRIPTION
		Uses System.IO.Compression.GZipStream to compress a SQL Server Inventory object that was returned by Get-SqlServerInventory and write it to disk.

	.PARAMETER  SqlServerInventory
		A SQL Server Inventory object returned by Get-SqlServerInventory.

	.PARAMETER  Path
		Specifies the path where the file will be written.

	.EXAMPLE
		Export-SqlServerInventoryToGzClixml -SqlServerInventory $Inventory -Path 'C:\SqlServerInventory.xml.gz'
		
	.LINK
		Import-SqlServerInventoryFromGzClixml
		Get-SqlServerInventory
#>
	[CmdletBinding()]
	param(
		[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		[ValidateNotNullOrEmpty()]
		[PSObject]$SqlServerInventory
		,
		[Parameter(Position=1, Mandatory=$true, ValueFromPipeline=$false)]
		[ValidateNotNullOrEmpty()]
		[String]$Path
	)
	begin {
	}
	process {

		$Inventory = $SqlServerInventory.psobject.Copy()
		#$Inventory.DatabaseServer = $SqlServerInventory.DatabaseServer.psobject.Copy()

		$Inventory.DatabaseServer = $Inventory.DatabaseServer | ForEach-Object {

			$DatabaseServer = $_.psobject.Copy()

			# Remove the reference from each DatabaseServer to its Windows machine (if one exists)
			# When serializing $SqlServerInventory it gets duplicated and we don't need that.
			# The reference is restored by Import-SqlServerInventoryToGzClixml
			if ($_.Machine) {
				$DatabaseServer.Machine = $null
			}

			$DatabaseServer.Server = $_.Server.psobject.Copy()
			#$DatabaseServer.Databases = $_.Server.Databases.psobject.Copy()

			$DatabaseServer.Server.Databases = $DatabaseServer.Server.Databases | ForEach-Object {
				$Database = $_.psobject.Copy()

				# In some cases there can be a TON of object permissions
				if ($($_.Properties.Permissions | Measure-Object).Count -gt 0) {
					$Database.Properties = $Database.Properties.psobject.Copy()
					$Database.Properties.Permissions = [System.Convert]::ToBase64String($($_.Properties.Permissions | ConvertTo-GzCliXml))
				}

				if ($($_.Tables | Measure-Object).Count -gt 0) {
					$Database.Tables = [System.Convert]::ToBase64String($($_.Tables | ConvertTo-GzCliXml))
				}

				if ($($_.Views | Measure-Object).Count -gt 0) {
					$Database.Views = [System.Convert]::ToBase64String($($_.Views | ConvertTo-GzCliXml))
				}

				# Programmability
				#$Database.Programmability = [System.Convert]::ToBase64String($($_.Programmability | ConvertTo-GzCliXml))
				$Database.Programmability = $_.Programmability.psobject.Copy()

				if ($($_.Programmability.StoredProcedures | Measure-Object).Count -gt 0) {
					$Database.Programmability.StoredProcedures = [System.Convert]::ToBase64String($($_.Programmability.StoredProcedures | ConvertTo-GzCliXml))
				}
				if ($($_.Programmability.ExtendedStoredProcedures | Measure-Object).Count -gt 0) {
					$Database.Programmability.ExtendedStoredProcedures = [System.Convert]::ToBase64String($($_.Programmability.ExtendedStoredProcedures | ConvertTo-GzCliXml))
				}
				if ($($_.Programmability.Functions | Measure-Object).Count -gt 0) {
					$Database.Programmability.Functions = [System.Convert]::ToBase64String($($_.Programmability.Functions | ConvertTo-GzCliXml))
				}
				if ($($_.Programmability.DatabaseTriggers | Measure-Object).Count -gt 0) {
					$Database.Programmability.DatabaseTriggers = [System.Convert]::ToBase64String($($_.Programmability.DatabaseTriggers | ConvertTo-GzCliXml))
				}
				if ($($_.Programmability.Assemblies | Measure-Object).Count -gt 0) {
					$Database.Programmability.Assemblies = [System.Convert]::ToBase64String($($_.Programmability.Assemblies | ConvertTo-GzCliXml))
				} 
				if ($($_.Programmability.Rules | Measure-Object).Count -gt 0) {
					$Database.Programmability.Rules = [System.Convert]::ToBase64String($($_.Programmability.Rules | ConvertTo-GzCliXml))
				}
				if ($($_.Programmability.Defaults | Measure-Object).Count -gt 0) {
					$Database.Programmability.Defaults = [System.Convert]::ToBase64String($($_.Programmability.Defaults | ConvertTo-GzCliXml))
				}
				if ($($_.Programmability.PlanGuides | Measure-Object).Count -gt 0) {
					$Database.Programmability.PlanGuides = [System.Convert]::ToBase64String($($_.Programmability.PlanGuides | ConvertTo-GzCliXml))
				}
				if ($($_.Programmability.Sequences | Measure-Object).Count -gt 0) {
					$Database.Programmability.Sequences = [System.Convert]::ToBase64String($($_.Programmability.Sequences | ConvertTo-GzCliXml))
				} 


				# Programmability - Types
				$Database.Programmability.Types = $_.Programmability.Types.psobject.Copy()

				if ($($_.Programmability.Types.UserDefinedAggregates | Measure-Object).Count -gt 0) {
					$Database.Programmability.Types.UserDefinedAggregates = [System.Convert]::ToBase64String($($_.Programmability.Types.UserDefinedAggregates | ConvertTo-GzCliXml))
				}
				if ($($_.Programmability.Types.UserDefinedDataTypes | Measure-Object).Count -gt 0) {
					$Database.Programmability.Types.UserDefinedDataTypes = [System.Convert]::ToBase64String($($_.Programmability.Types.UserDefinedDataTypes | ConvertTo-GzCliXml))
				}
				if ($($_.Programmability.Types.UserDefinedTableTypes | Measure-Object).Count -gt 0) {
					$Database.Programmability.Types.UserDefinedTableTypes = [System.Convert]::ToBase64String($($_.Programmability.Types.UserDefinedTableTypes | ConvertTo-GzCliXml))
				}
				if ($($_.Programmability.Types.UserDefinedTypes | Measure-Object).Count -gt 0) {
					$Database.Programmability.Types.UserDefinedTypes = [System.Convert]::ToBase64String($($_.Programmability.Types.UserDefinedTypes | ConvertTo-GzCliXml))
				}
				if ($($_.Programmability.Types.XmlSchemaCollections | Measure-Object).Count -gt 0) {
					$Database.Programmability.Types.XmlSchemaCollections = [System.Convert]::ToBase64String($($_.Programmability.Types.XmlSchemaCollections | ConvertTo-GzCliXml))
				}


				$Database.ServiceBroker = [System.Convert]::ToBase64String($($_.ServiceBroker | ConvertTo-GzCliXml))

				Write-Output $Database
			}

			if ($($_.Server.Management.SQLTrace | Measure-Object).count -gt 0) {
				$DatabaseServer.Server.Management = $_.Server.Management.psobject.Copy()
				$DatabaseServer.Server.Management.SQLTrace = [System.Convert]::ToBase64String($($_.Server.Management.SQLTrace | convertto-gzclixml)) #get-compressedpsobject -inputobject $_.Server.Management.SQLTrace
			}

			<#
			# In testing I found that there is no advantage to compressing Agent Jobs
			# Export time and file size were WORSE when compressing Agent Jobs
			if ($($_.Agent.Jobs | Measure-Object).count -gt 0) {
				$DatabaseServer.Agent = $_.Agent.psobject.Copy()
				$DatabaseServer.Agent.Jobs = [System.Convert]::ToBase64String($($_.Agent.Jobs | convertto-gzclixml))
			}
			#>

			# Compress the database server and convert it to a Base64 string
			[System.Convert]::ToBase64String($($DatabaseServer | ConvertTo-GzCliXml))

		}

		# Export to disk
		$Inventory | Export-GzCliXml -Path $Path
	}
	end {
		Remove-Variable -Name Inventory
	}
}

function Import-SqlServerInventoryFromGzClixml {
	<#
	.SYNOPSIS
		Imports a GZip compressed SQL Server Inventory that was written to disk by Export-SqlServerInventoryToGzClixml

	.DESCRIPTION
		Uses System.IO.Compression.GZipStream to expand a SQL Server Inventory object that was compressed and written to disk by Export-SqlServerInventoryToGzClixml.

	.PARAMETER  Path
		Fully qualified path to the GZip compressed SQL Server Inventory written to disk by Export-SqlServerInventoryToGzClixml

	.EXAMPLE
		Import-SqlServerInventoryFromGzClixml -Path 'C:\SqlServerInventory.xml.gz'
		
	.LINK
		Export-SqlServerInventoryToGzClixml
		Get-SqlServerInventory
#>
	[CmdletBinding()]
	param(
		[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		[ValidateNotNullOrEmpty()]
		[String[]]$Path
	)
	begin {
	}
	process {
		foreach ($InputObject in $Path) {

			$Inventory = Import-GzCliXml -Path $InputObject

			# Expand the database server that was stored as a Base64 string representation of the compressed clixml
			$Inventory.DatabaseServer = $Inventory.DatabaseServer | ForEach-Object {
				ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_))
			}

			$Inventory.DatabaseServer | ForEach-Object {

				$_.Server.Databases | ForEach-Object {

					if (
						$_.Properties.Permissions.Length -gt 0 -and
						$_.Properties.Permissions.GetType().Name -ieq 'String'
					) {
						$_.Properties.Permissions = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Properties.Permissions))
					}

					if (
						$_.Tables.Length -gt 0 -and
						$_.Tables.GetType().Name -ieq 'String'
					) {
						$_.Tables = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Tables))
					}

					if (
						$_.Views.Length -gt 0 -and
						$_.Views.GetType().Name -ieq 'String'
					) {
						$_.Views = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Views)) 
					}


					if (
						$_.Programmability.StoredProcedures.Length -gt 0 -and
						$_.Programmability.StoredProcedures.GetType().Name -ieq 'String'
					) {
						$_.Programmability.StoredProcedures = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Programmability.StoredProcedures))
					}
					if (
						$_.Programmability.ExtendedStoredProcedures.Length -gt 0 -and
						$_.Programmability.ExtendedStoredProcedures.GetType().Name -ieq 'String'
					) {
						$_.Programmability.ExtendedStoredProcedures = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Programmability.ExtendedStoredProcedures))
					}
					if (
						$_.Programmability.Functions.Length -gt 0 -and
						$_.Programmability.Functions.GetType().Name -ieq 'String'
					) {
						$_.Programmability.Functions = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Programmability.Functions))
					}
					if (
						$_.Programmability.DatabaseTriggers.Length -gt 0 -and
						$_.Programmability.DatabaseTriggers.GetType().Name -ieq 'String'
					) {
						$_.Programmability.DatabaseTriggers = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Programmability.DatabaseTriggers))
					}
					if (
						$_.Programmability.Assemblies.Length -gt 0 -and
						$_.Programmability.Assemblies.GetType().Name -ieq 'String'
					) {
						$_.Programmability.Assemblies = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Programmability.Assemblies))
					}
					if (
						$_.Programmability.Rules.Length -gt 0 -and
						$_.Programmability.Rules.GetType().Name -ieq 'String'
					) {
						$_.Programmability.Rules = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Programmability.Rules))
					}
					if (
						$_.Programmability.Defaults.Length -gt 0 -and
						$_.Programmability.Defaults.GetType().Name -ieq 'String'
					) {
						$_.Programmability.Defaults = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Programmability.Defaults))
					}
					if (
						$_.Programmability.PlanGuides.Length -gt 0 -and
						$_.Programmability.PlanGuides.GetType().Name -ieq 'String'
					) {
						$_.Programmability.PlanGuides = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Programmability.PlanGuides))
					}
					if (
						$_.Programmability.Sequences.Length -gt 0 -and
						$_.Programmability.Sequences.GetType().Name -ieq 'String'
					) {
						$_.Programmability.Sequences = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Programmability.Sequences))
					}
					if (
						$_.Programmability.Types.UserDefinedAggregates.Length -gt 0 -and
						$_.Programmability.Types.UserDefinedAggregates.GetType().Name -ieq 'String'
					) {
						$_.Programmability.Types.UserDefinedAggregates = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Programmability.Types.UserDefinedAggregates))
					}
					if (
						$_.Programmability.Types.UserDefinedDataTypes.Length -gt 0 -and
						$_.Programmability.Types.UserDefinedDataTypes.GetType().Name -ieq 'String'
					) {
						$_.Programmability.Types.UserDefinedDataTypes = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Programmability.Types.UserDefinedDataTypes))
					}
					if (
						$_.Programmability.Types.UserDefinedTableTypes.Length -gt 0 -and
						$_.Programmability.Types.UserDefinedTableTypes.GetType().Name -ieq 'String'
					) {
						$_.Programmability.Types.UserDefinedTableTypes = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Programmability.Types.UserDefinedTableTypes))
					}
					if (
						$_.Programmability.Types.UserDefinedTypes.Length -gt 0 -and
						$_.Programmability.Types.UserDefinedTypes.GetType().Name -ieq 'String'
					) {
						$_.Programmability.Types.UserDefinedTypes = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Programmability.Types.UserDefinedTypes))
					}
					if (
						$_.Programmability.Types.XmlSchemaCollections.Length -gt 0 -and
						$_.Programmability.Types.XmlSchemaCollections.GetType().Name -ieq 'String'
					) {
						$_.Programmability.Types.XmlSchemaCollections = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Programmability.Types.XmlSchemaCollections))
					}

					if (
						$_.ServiceBroker.Length -gt 0 -and
						$_.ServiceBroker.GetType().Name -ieq 'String'
					) {
						$_.ServiceBroker = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.ServiceBroker))
					}

				}

				if (
					$_.Server.Management.SQLTrace.Length -gt 0 -and
					$_.Server.Management.SQLTrace.GetType().Name -ieq 'String'
				) {
					$_.Server.Management.SQLTrace = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Server.Management.SQLTrace))
				}

				if (
					$_.Agent.Jobs.Length -gt 0 -and
					$_.Agent.Jobs.GetType().Name -ieq 'String'
				) {
					$_.Agent.Jobs = ConvertFrom-GzCliXml -InputObject $([System.Convert]::FromBase64String($_.Agent.Jobs))
				}

			}

			# Create a reference from each DatabaseServer to its Windows machine
			# This reference was removed by ConvertTo-GzSqlServerInventory
			foreach ($Machine in ($Inventory.WindowsInventory.Machine)) {
				$Inventory.Service | Where-Object { 
					$_.ComputerName -ieq $Machine.OperatingSystem.Settings.ComputerSystem.FullyQualifiedDomainName -and 
					$_.ServiceTypeName -ieq 'sql server'
				} | ForEach-Object {
					$InventoryServiceId = $_.InventoryServiceId
					$Inventory.DatabaseServer | Where-Object { $_.InventoryServiceId -ieq $InventoryServiceId } | ForEach-Object {
						$_ | Add-Member -MemberType NoteProperty -Name Machine -Value $Machine -Force
					} 
				}
			}

			Write-Output $Inventory
		}

	}
	end {
		Remove-Variable -Name Inventory
	}
}

function Get-SqlServerInventory {
	<#
	.SYNOPSIS
		Collects comprehensive information about SQL Server instances and their underlying Windows Operating System.

	.DESCRIPTION
		The Get-SqlServerInventory function leverages the NetworkScan, SqlServerDatabaseEngine, and WindowsInventory modules along with SQL Server Shared Management Objects (SMO) and Windows Management Instrumentation (WMI) to scan for and collect comprehensive information about SQL Server instances and their underlying Windows Operating System.
		
		Get-SqlServerInventory can find, verify, and collect information by Computer Name, Subnet Scan, or Active Directory DNS query.
		
		Get-SqlServerInventory collects information from SQL Server 2000 or higher and Windows Azure SQL Database (if using SMO 2008 or higher).
		
		This function works best when using a version of SMO that matches or is higher than the highest version of each SQL Server instance information is being collected from.
		
		The latest version of SMO can be downloaded from http://www.microsoft.com/en-us/download/details.aspx?id=29065
		Note that SMO also requires the Microsoft SQL Server System CLR Types which can be downloaded from the same page
				
	.PARAMETER  DnsServer
		'Automatic', or the Name or IP address of an Active Directory DNS server to query for a list of hosts to inventory (if there is an instance of SQL Server installed).
		
		When 'Automatic' is specified the function will use WMI queries to discover the current computer's DNS server(s) to query.

	.PARAMETER  DnsDomain
		'Automatic' or the Active Directory domain name to use when querying DNS for a list of hosts.
		
		When 'Automatic' is specified the function will use the current computer's AD domain.
		
		'Automatic' will be used by default if DnsServer is specified but DnsDomain is not provided.
		
	.PARAMETER  Subnet
		'Automatic' or a comma delimited list of subnets (in CIDR notation) to scan for SQL Server instances.
		
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

	.PARAMETER  Username
		SQL Server username to use when connecting to instances. 
		
		Windows authentication will be used to connect if this parameter is not provided.

	.PARAMETER  Password
		SQL Server password to use when connecting to instances.

	.PARAMETER  MaxConcurrencyThrottle
		Number between 1-100 to indicate how many instances to collect information from concurrently.

		If not provided then the number of logical CPUs present to your session will be used.

	.PARAMETER  PrivateOnly
		Restrict inventory to instances on private class A, B, or C IP addresses

	.PARAMETER  ParentProgressId
		If the caller is using Write-Progress then all progress information will be written using ParentProgressId as the ParentID		

	.PARAMETER  IncludeDatabaseObjectPermissions
		Includes database object level permissions (System object permissions included only if -IncludeDatabaseSystemObjects is also provided)

	.PARAMETER  IncludeDatabaseObjectInformation
		Includes database object information (System objects included only if -IncludeDatabaseSystemObjects is also provided)

	.PARAMETER  IncludeDatabaseSystemObjects
		Include system objects when retrieving database object information. 
		
		This has no effect if neither -IncludeDatabaseObjectInformation nor -IncludeDatabaseObjectPermissions are specified.


	.EXAMPLE
		Get-SqlServerInventory -DNSServer automatic -DNSDomain automatic -PrivateOnly
		
		Description
		-----------
		Collect an inventory by querying Active Directory for a list of hosts to scan for SQL Server instances. The list of hosts will be restricted to private IP addresses only.
		
		Windows Authentication will be used to connect to each instance.
		
		Database objects will NOT be included in the results.

	.EXAMPLE
		Get-SqlServerInventory -Subnet 172.20.40.0/28 -Username sa -Password BetterNotBeBlank
		
		Description
		-----------
		Collect an inventory by scanning all hosts in the subnet 172.20.40.0/28 for SQL Server instances.
		
		SQL authentication (username = "sa", password = "BetterNotBeBlank") will be used to connect to the instance.
		
		Database objects will NOT be included in the results.

	.EXAMPLE
		Get-SqlServerInventory -Computername Server1,Server2,Server3
		
		Description
		-----------
		Collect an inventory by scanning Server1, Server2, and Server3 for SQL Server instances.
		
		Windows Authentication will be used to connect to the instance.
		
		Database objects will NOT be included in the results.


	.EXAMPLE
		Get-SqlServerInventory -Computername $env:COMPUTERNAME -IncludeDatabaseObjectInformation
		
		Description
		-----------
		Collect an inventory by scanning the local machine for SQL Server instances.
		
		Windows Authentication will be used to connect to the instance.
		
		Database objects (EXCLUDING system objects) will be included in the results.

	.EXAMPLE
		Get-SqlServerInventory -Computername $env:COMPUTERNAME -IncludeDatabaseObjectInformation -IncludeDatabaseSystemObjects

		Description
		-----------
		Collect an inventory by scanning the local machine for SQL Server instances.
		
		Windows Authentication will be used to connect to the instance.
		
		Database objects (INCLUDING system objects) will be included in the results.
		

	.OUTPUTS
		System.Management.Automation.PSObject

	.NOTES

	.LINK
		Export-SqlServerInventoryDatabaseEngineConfigToExcel

#>
	[cmdletBinding(DefaultParametersetName='computername_WindowsAuthentication')]
	param(
		[Parameter(Mandatory=$true, ParameterSetName='dns_SQLAuthentication', HelpMessage='DNS Server(s)')]
		[Parameter(Mandatory=$true, ParameterSetName='dns_WindowsAuthentication', HelpMessage='DNS Server(s)')]
		[alias('dns')]
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$|^auto$|^automatic$')]
		[string[]]
		$DnsServer = 'automatic'
		,
		[Parameter(Mandatory=$false, ParameterSetName='dns_SQLAuthentication', HelpMessage='DNS Domain Name')] 
		[Parameter(Mandatory=$false, ParameterSetName='dns_WindowsAuthentication', HelpMessage='DNS Domain Name')] 
		[alias('domain')]
		[string]
		$DnsDomain = 'automatic'
		,
		[Parameter(Mandatory=$true, ParameterSetName='subnet_SQLAuthentication', HelpMessage='Subnet (in CIDR notation)')] 
		[Parameter(Mandatory=$true, ParameterSetName='subnet_WindowsAuthentication', HelpMessage='Subnet (in CIDR notation)')] 
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)[\\/]\d{1,2}$|^auto$|^automatic$')]
		[string[]]
		$Subnet = 'automatic'
		,
		[Parameter(Mandatory=$true, ParameterSetName='computername_SQLAuthentication', HelpMessage='Computer Name(s)')] 
		[Parameter(Mandatory=$true, ParameterSetName='computername_WindowsAuthentication', HelpMessage='Computer Name(s)')] 
		[alias('computer')]
		[string[]]
		$ComputerName
		,
		[Parameter(Mandatory=$false, ParameterSetName='dns_SQLAuthentication')]
		[Parameter(Mandatory=$false, ParameterSetName='dns_WindowsAuthentication')]
		[Parameter(Mandatory=$false, ParameterSetName='subnet_SQLAuthentication')]
		[Parameter(Mandatory=$false, ParameterSetName='subnet_WindowsAuthentication')]
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)[\\/]\d{1,2}$')]
		[string[]]
		$ExcludeSubnet
		,
		[Parameter(Mandatory=$false, ParameterSetName='dns_SQLAuthentication')]
		[Parameter(Mandatory=$false, ParameterSetName='dns_WindowsAuthentication')]
		[ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)[\\/]\d{1,2}$')]
		[string[]]
		$LimitSubnet
		,
		[Parameter(Mandatory=$false, ParameterSetName='dns_SQLAuthentication')]
		[Parameter(Mandatory=$false, ParameterSetName='dns_WindowsAuthentication')]
		[Parameter(Mandatory=$false, ParameterSetName='subnet_SQLAuthentication')]
		[Parameter(Mandatory=$false, ParameterSetName='subnet_WindowsAuthentication')]
		[string[]]
		$ExcludeComputerName
		,
		[Parameter(Mandatory=$true, ParameterSetName='dns_SQLAuthentication') ]
		[Parameter(Mandatory=$true, ParameterSetName='subnet_SQLAuthentication') ]
		[Parameter(Mandatory=$true, ParameterSetName='computername_SQLAuthentication') ]
		[ValidateNotNull()]
		[System.String]
		$Username
		,
		[Parameter(Mandatory=$true, ParameterSetName='dns_SQLAuthentication') ]
		[Parameter(Mandatory=$true, ParameterSetName='subnet_SQLAuthentication') ]
		[Parameter(Mandatory=$true, ParameterSetName='computername_SQLAuthentication') ]
		[ValidateNotNull()]
		[System.String]
		$Password
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
		,
		[Parameter(Mandatory=$false)]
		[switch]
		$IncludeDatabaseObjectPermissions = $false
		,
		[Parameter(Mandatory=$false)]
		[alias('IncludeDbObjects','IncludeDatabaseObjects')]
		[switch]
		$IncludeDatabaseObjectInformation = $false
		,
		[Parameter(Mandatory=$false)]
		[switch]
		$IncludeDatabaseSystemObjects = $false
	)
	process {

		$Inventory = New-Object -TypeName psobject -Property @{ 
			#Machine = @()
			WindowsInventory = $null
			Service = @()
			DatabaseServer = @()
			## Eventually these can be added later
			#	IntegrationServer = @()
			#	AnalysisServer = @()
			#	ReportServer = @()
			Version = $ModuleVersion
			StartDateUTC = [DateTime]::UtcNow
			EndDateUTC = $null
			DatabaseServerScanSuccessCount = 0
			DatabaseServerScanErrorCount = 0
			#MachineScanSuccessCount = 0
			#MachineScanErrorCount = 0
		} | Add-Member -MemberType ScriptProperty -Name DatabaseServerScanCount -Value {
			$this.DatabaseServerScanSuccessCount + $this.DatabaseServerScanErrorCount
		} -PassThru


		$SqlServerService = @()
		$ParameterHash = $null
		$Machine = $null
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

		Write-SqlServerInventoryLog -Message "Start Function: $($MyInvocation.InvocationName)" -MessageLevel Debug
		Write-Progress -Activity 'SQL Server Inventory' -PercentComplete 0 -Status 'Discovering SQL Server Instances' -Id $MasterProgressId -ParentId $ParentProgressId

		# Build command for splatting
		$ParameterHash = @{
			MaxConcurrencyThrottle = $MaxConcurrencyThrottle
			PrivateOnly = $PrivateOnly
			ParentProgressId = $MasterProgressId
		}

		switch ($PsCmdlet.ParameterSetName) {
			'dns_SQLAuthentication' {
				$ParameterHash.Add('DnsServer',$DnsServer)
				$ParameterHash.Add('DnsDomain',$DnsDomain)
				if ($ExcludeSubnet) { $ParameterHash.Add('ExcludeSubnet',$ExcludeSubnet) }
				if ($LimitSubnet) { $ParameterHash.Add('IncludeSubnet',$LimitSubnet) }
				if ($ExcludeComputerName) { $ParameterHash.Add('ExcludeComputerName',$ExcludeComputerName) }
				$TotalScanCount = 1
			}
			'dns_WindowsAuthentication' {
				$ParameterHash.Add('DnsServer',$DnsServer)
				$ParameterHash.Add('DnsDomain',$DnsDomain)
				if ($ExcludeSubnet) { $ParameterHash.Add('ExcludeSubnet',$ExcludeSubnet) }
				if ($LimitSubnet) { $ParameterHash.Add('IncludeSubnet',$LimitSubnet) }
				if ($ExcludeComputerName) { $ParameterHash.Add('ExcludeComputerName',$ExcludeComputerName) }
				$TotalScanCount = 1
			}
			'subnet_SQLAuthentication' {
				$ParameterHash.Add('Subnet',$Subnet)
				if ($ExcludeSubnet) { $ParameterHash.Add('ExcludeSubnet',$ExcludeSubnet) }
				if ($ExcludeComputerName) { $ParameterHash.Add('ExcludeComputerName',$ExcludeComputerName) }
				$TotalScanCount = 1
			}
			'subnet_WindowsAuthentication' {
				$ParameterHash.Add('Subnet',$Subnet)
				if ($ExcludeSubnet) { $ParameterHash.Add('ExcludeSubnet',$ExcludeSubnet) }
				if ($ExcludeComputerName) { $ParameterHash.Add('ExcludeComputerName',$ExcludeComputerName) }
				$TotalScanCount = 1
			}
			'computername_SQLAuthentication' {
				# Don't bother with Windows Azure SQL Databases
				# As of 2013/03/27 there is no way to discover WASD services with SMO & WMI
				# We'll just assume it's running and try to connect later
				#$ParameterHash.Add('ComputerName',$ComputerName)
				$ParameterHash.Add('ComputerName',$($ComputerName | Where-Object { $_ -inotlike '*.database.windows.net' }))
				$TotalScanCount = $($ParameterHash.ComputerName | Measure-Object).Count
			}
			'computername_WindowsAuthentication' {
				$ParameterHash.Add('ComputerName',$($ComputerName | Where-Object { $_ -inotlike '*.database.windows.net' }))
				$TotalScanCount = $($ParameterHash.ComputerName | Measure-Object).Count
			}
		}

		if ($TotalScanCount -gt 0) {
			# Scan the network to find SQL Server Services
			# Some devices may have multiple IP Addresses so only use the first WMI-capable address for each
			Find-SqlServerService @ParameterHash | ForEach-Object {
				# Add a GUID called InventoryServiceId to each object returned. We'll use this later 
				$Inventory.Service += $_ | Add-Member -MemberType NoteProperty -Name InventoryServiceId -Value $([System.Guid]::NewGuid()) -Force -PassThru
			}

			# Update $TotalScanCount to reflect how many services were actually found 
			$TotalScanCount = $($Inventory.Service | Where-Object { ($_.ServiceTypeName -ieq 'sql server') -and ($_.ServiceState -ieq 'running') } | Measure-Object).Count
		}

		# Add Windows Azure SQL Databases to $TotalScanCount
		if ($PsCmdlet.ParameterSetName -ieq 'computername_SQLAuthentication') {
			$TotalScanCount += $($ComputerName | Where-Object { $_ -ilike '*.database.windows.net' } | Select-Object -Unique | Measure-Object).Count
		}

		Write-SqlServerInventoryLog -Message "Beginning scan of $TotalScanCount instance(s)" -MessageLevel Information
		Write-Progress -Activity 'SQL Server Inventory' -PercentComplete 33 -Status 'Collecting SQL Server Instance Information' -Id $MasterProgressId
		Write-Progress -Activity 'Scanning Instances' -PercentComplete 0 -Status "Scanning $TotalScanCount instance(s)" -Id $ScanProgressId -ParentId $MasterProgressId


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
				[System.Collections.Hashtable]$SplatHash
			)
			Import-Module -Name LogHelper, SqlServerDatabaseEngineInformation
			Set-LogQueue -Queue $LogQueue
			Set-LogFile -Path $LogPath
			Set-LoggingPreference -Preference $LoggingPreference
			Get-SqlServerDatabaseEngineInformation @SplatHash
			Remove-Module -Name LogHelper, SqlServerDatabaseEngineInformation

			#region
			# Start-Sleep -Seconds 2
			# Write-Output (
			# 	New-Object -TypeName psobject -Property @{
			# 		Server = @{
			# 			Configuration = $null
			# 			Databases = @()
			# 			ServerObjects = @{
			# 				StartupProcedures = @()
			# 				LinkedServers = @()
			# 			}
			# 			Security = $null
			# 			Service = $null
			# 		}
			# 		Agent = @{
			# 			Configuration = $null
			# 			Service = $null
			# 			Jobs = @()
			# 		}
			# 		ScanDateUTC = [DateTime]::UtcNow
			# 		ScanErrorCount = 0
			# 	}
			# )
			#endregion

		}


		# Iterate through each SQL Server instance that's running and gather information
		$Inventory.Service | Where-Object { ($_.ServiceTypeName -ieq 'sql server') -and ($_.ServiceState -ieq 'running') } | ForEach-Object {

			$CurrentScanCount++
			Write-SqlServerInventoryLog -Message "Gathering information from $($_.ServerName) at $($_.ServiceIpAddress) [Instance $CurrentScanCount of $TotalScanCount]" -MessageLevel Information

			# Build command for splatting
			$ParameterHash = @{
				InstanceName = $_.ServerName
				StopAtErrorCount = $ScanErrorThreshold
				Port = if ($_.Port -eq 1433) { $null } else { $_.Port }
				IncludeDatabaseObjectPermissions = $IncludeDatabaseObjectPermissions
				IncludeDatabaseObjectInformation = $IncludeDatabaseObjectInformation
				IncludeDatabaseSystemObjects = $IncludeDatabaseSystemObjects
			}

			switch ($PsCmdlet.ParameterSetName) {
				'dns_SQLAuthentication' {
					$ParameterHash.Add('Username',$Username)
					$ParameterHash.Add('Password',$Password)
				}
				'subnet_SQLAuthentication' {
					$ParameterHash.Add('Username',$Username)
					$ParameterHash.Add('Password',$Password)
				}
				'computername_SQLAuthentication' {
					$ParameterHash.Add('Username',$Username)
					$ParameterHash.Add('Password',$Password)
				}
			}


			#Create the PowerShell instance and supply the scriptblock with the other parameters
			$PowerShell = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
			$PowerShell = $PowerShell.AddArgument($(Get-SqlServerInventoryLog))
			$PowerShell = $PowerShell.AddArgument($(Get-SqlServerInventoryLoggingPreference))
			$PowerShell = $PowerShell.AddArgument($script:LogQueue)
			$PowerShell = $PowerShell.AddArgument($ParameterHash)

			#Add the runspace into the PowerShell instance
			$PowerShell.RunspacePool = $RunspacePool

			$Runspaces.Add((
					New-Object -TypeName PsObject -Property @{
						PowerShell = $PowerShell
						Runspace = $PowerShell.BeginInvoke()
						ServiceInfo = $_
					}
				)) | Out-Null

		}


		# Scan for Windows Azure SQL Database
		if ($PsCmdlet.ParameterSetName -ieq 'computername_SQLAuthentication') {

			# Scan for Windows Azure SQL Database
			$ComputerName | Where-Object { $_ -ilike '*.database.windows.net' } | Select-Object -Unique | ForEach-Object {

				$CurrentScanCount++
				Write-SqlServerInventoryLog -Message "Gathering information from $($_) [Instance $CurrentScanCount of $TotalScanCount]" -MessageLevel Information

				# Build command for splatting
				$ParameterHash = @{
					InstanceName = $_
					StopAtErrorCount = $ScanErrorThreshold
					Port = $null
					IncludeDatabaseObjectInformation = $IncludeDatabaseObjectInformation
					IncludeDatabaseSystemObjects = $IncludeDatabaseSystemObjects
					Username = $Username
					Password = $Password
				}

				# Create the PowerShell instance and supply the scriptblock with the other parameters
				$PowerShell = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
				$PowerShell = $PowerShell.AddArgument($(Get-SqlServerInventoryLog))
				$PowerShell = $PowerShell.AddArgument($(Get-SqlServerInventoryLoggingPreference))
				$PowerShell = $PowerShell.AddArgument($script:LogQueue)
				$PowerShell = $PowerShell.AddArgument($ParameterHash)

				# Add the runspace into the PowerShell instance
				# Simulate that there was a service even though WASD was excluded from the Services scan
				$PowerShell.RunspacePool = $RunspacePool

				$Runspaces.Add((
						New-Object -TypeName PsObject -Property @{
							PowerShell = $PowerShell
							Runspace = $PowerShell.BeginInvoke()
							ServiceInfo = New-Object -TypeName PSObject -Property @{ 
								InventoryServiceId = [System.Guid]::NewGuid()
							}
						}
					)) | Out-Null
			}
		} 


		# Reset the scan counter
		$CurrentScanCount = 0

		# Process results as they complete until they are all complete
		Do {
			$Runspaces | ForEach-Object {

				If ($_.Runspace.IsCompleted) {

					$ServiceInfo = $_.ServiceInfo

					try {
						# This is where the output gets returned
						$_.PowerShell.EndInvoke($_.Runspace) | ForEach-Object {
							if ($_.ScanErrorCount -lt $ScanErrorThreshold) {

								# Add the InventoryServiceId to the object
								$Inventory.DatabaseServer += $_ | Add-Member -MemberType NoteProperty -Name InventoryServiceId -Value $ServiceInfo.InventoryServiceId -Force -PassThru

								$Inventory.DatabaseServerScanSuccessCount++
								Write-SqlServerInventoryLog -Message "Scanned $($ServiceInfo.ServerName) with $($_.ScanErrorCount) errors" -MessageLevel Information

							} else {
								$Inventory.DatabaseServerScanErrorCount++
								Write-SqlServerInventoryLog -Message "Failed to scan $($ServiceInfo.ServerName) -  $($_.ScanErrorCount) errors" -MessageLevel Error
							}
						} 
					}
					catch {
						$Inventory.DatabaseServerScanErrorCount++
						Write-SqlServerInventoryLog -Message "An unrecoverable error was encountered while attempting to retrieve information from $($ServiceInfo.ServerName)" -MessageLevel Error
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
				Write-Progress -Activity 'Scanning Instances' -PercentComplete (($CurrentScanCount / $TotalScanCount)*100) -Status "Scanned $CurrentScanCount of $TotalScanCount Instance(s)" -Id $ScanProgressId -ParentId $MasterProgressId
			}

		} while (($Runspaces | Where-Object {$_.Runspace -ne $Null} | Measure-Object).Count -gt 0)

		# Finally, close the runspaces
		$RunspacePool.close()

		Write-Progress -Activity 'Scanning Instances' -PercentComplete 100 -Status 'Scan Complete' -Id $ScanProgressId -ParentId $MasterProgressId
		Write-SqlServerInventoryLog -Message "Instance scan complete (Success: $($Inventory.DatabaseServerScanSuccessCount); Errors: $($Inventory.DatabaseServerScanErrorCount))" -MessageLevel Information


		# Now collect information about Windows for each distinct machine we found a running SQL Service on
		# Windows Azure SQL Databases will NOT be included in this list since they weren't included in the Services Scan
		#region

		if ($TotalScanCount -gt 0) {

			Write-Progress -Activity 'SQL Server Inventory' -PercentComplete 66 -Status 'Collecting Windows Machine Information' -Id $MasterProgressId -ParentId $ParentProgressId

			# Double check that there are machines to scan
			# We might not have any if only Windows Azure SQL Databases were included in the inventory
			if ($($Inventory.Service | Measure-Object).Count -gt 0) {

				# Build command for splatting
				$ParameterHash = @{
					MaxConcurrencyThrottle = $MaxConcurrencyThrottle
					PrivateOnly = $PrivateOnly
					ParentProgressId = $MasterProgressId
					ComputerName = @($Inventory.Service | Select-Object -Property ComputerName -Unique | ForEach-Object { $_.ComputerName })
					AdditionalData = @('All')
				}
				
				$Inventory.WindowsInventory = Get-WindowsInventory @ParameterHash 

			}

			# Create a reference from each DatabaseServer to its Windows machine
			foreach ($Machine in ($Inventory.WindowsInventory.Machine)) {
				$Inventory.Service | Where-Object { 
					$_.ComputerName -ieq $Machine.OperatingSystem.Settings.ComputerSystem.FullyQualifiedDomainName -and 
					$_.ServiceTypeName -ieq 'sql server'
				} | ForEach-Object {
					$InventoryServiceId = $_.InventoryServiceId
					$Inventory.DatabaseServer | Where-Object { $_.InventoryServiceId -ieq $InventoryServiceId } | ForEach-Object {
						$_ | Add-Member -MemberType NoteProperty -Name Machine -Value $Machine -Force
					} 
				}
			}

		}
		#endregion


		# Record the scan end date
		$Inventory.EndDateUTC = [DateTime]::UtcNow

		Write-SqlServerInventoryLog -Message "End Function: $($MyInvocation.InvocationName)" -MessageLevel Debug

		Write-Progress -Activity 'Scanning Instances' -PercentComplete 100 -Status 'Scan Complete' -Id $ScanProgressId -ParentId $MasterProgressId -Completed 
		Write-Progress -Activity 'SQL Server Inventory' -PercentComplete 100 -Status 'Inventory Complete' -Id $MasterProgressId -ParentId $ParentProgressId -Completed

		# Return a compressed version of the inventory
		#Write-Output $(Get-CompressedPsObject -InputObject $Inventory)

		Write-Output $Inventory

		Remove-Variable -Name Inventory, SqlServerService, ParameterHash, CurrentScanCount, TotalScanCount, ScanProgressId, MasterProgressId

	}
}

function Get-SqlServerInventoryDatabaseEngineAssessment {
	[cmdletBinding()]
	param(
		[Parameter(Mandatory=$true, ValueFromPipeline=$true)]
		[PSCustomObject]
		$SqlServerInventory
		,
		[Parameter(Mandatory=$false)]
		[ValidateNotNull()]
		[Int32]
		$ParentProgressId = -1
	)
	begin {
		$NullDatabaseName = $null
	}
	process {

		# Only look at standalone servers; i.e. ignore Windows Azure SQL Database (for now)
		#region
		$SqlServerInventory.DatabaseServer | Where-Object { $_.Server.Configuration.General.ServerType -ieq 'standalone' } | ForEach-Object {

			$DatabaseServer = $_
			$ServerName = $DatabaseServer.Server.Configuration.General.Name

			$ServerServiceAccount = $_.Server.Service.ServiceAccount
			$AgentServiceAccount = $_.Agent.Service.ServiceAccount

			$ServerVersion = $DatabaseServer.Server.Configuration.General.Version
			$HelpUrlModifier = switch -wildcard ($DatabaseServer.Server.Configuration.General.Version) {
				'11.*' { '(v=sql.110)' } # SQL 2012
				'10.5' { '(v=sql.105)' } # SQL 2008 R2
				'10.*' { '(v=sql.100)' } # SQL 2008
				'9.*' { '(v=sql.90)' } # SQL 2005
				default { [String]::Empty } # Everything else
			}
			$ScanDateLocal = $DatabaseServer.ScanDateUTC.ToLocalTime()

			$DatabaseServer.Server.Security.Logins | Where-Object { $_.Sid -eq [System.BitConverter]::ToString(0x01) } | ForEach-Object { $SaLogin = $_.Name }
			if (-not $SaLogin) { $SaLogin = 'sa' }

			if ($DatabaseServer.Machine.OperatingSystem.Settings.OperatingSystem.WindowsDirectory) {
				$SystemDrive = [System.IO.Path]::GetPathRoot($DatabaseServer.Machine.OperatingSystem.Settings.OperatingSystem.WindowsDirectory)
			} else {
				# Assume System Drive is on C:
				$SystemDrive = 'C:\'
			}

			# Use the processor count that SQL Server sees
			#$CpuCount = $_.Machine.Hardware.MotherboardControllerAndPort.Processor.NumberOfLogicalProcessors
			$CpuCount = $DatabaseServer.Server.Configuration.General.ProcessorCount

			$ServerCollation = $DatabaseServer.Server.Configuration.General.ServerCollation

			$DatabaseServer.Server.Databases | ForEach-Object {
				if ($_.Name -ieq 'model') {
					$ModelDbCompatLevel = $_.Properties.Options.CompatibilityLevel
				}
				if ($_.Name -ieq 'tempdb') {
					$TempDbCollation = $_.Properties.General.Maintenance.Collation
				}
			}


			# Root paths for defaults
			$DefaultDataPathRoot = if (
				-not [String]::IsNullOrEmpty($DatabaseServer.Server.Configuration.DatabaseSettings.DataPath) -and
				-not (
					[System.IO.Path]::GetInvalidPathChars() | Where-Object {
						$DatabaseServer.Server.Configuration.DatabaseSettings.DataPath.Contains($_)
					}
				) -and
				[System.IO.Path]::IsPathRooted($DatabaseServer.Server.Configuration.DatabaseSettings.DataPath) -eq $true
			) {
				[System.IO.Path]::GetPathRoot($DatabaseServer.Server.Configuration.DatabaseSettings.DataPath)
			} else {
				$null
			}

			$DefaultLogPathRoot = if (
				-not [String]::IsNullOrEmpty($DatabaseServer.Server.Configuration.DatabaseSettings.LogPath) -and
				-not (
					[System.IO.Path]::GetInvalidPathChars() | Where-Object {
						$DatabaseServer.Server.Configuration.DatabaseSettings.LogPath.Contains($_)
					}
				) -and
				[System.IO.Path]::IsPathRooted($DatabaseServer.Server.Configuration.DatabaseSettings.LogPath) -eq $true
			) {
				[System.IO.Path]::GetPathRoot($DatabaseServer.Server.Configuration.DatabaseSettings.LogPath)
			} else {
				$null
			}

			$DefaultBackupPathRoot = if (
				-not [String]::IsNullOrEmpty($DatabaseServer.Server.Configuration.DatabaseSettings.BackupPath) -and
				-not (
					[System.IO.Path]::GetInvalidPathChars() | Where-Object {
						$DatabaseServer.Server.Configuration.DatabaseSettings.BackupPath.Contains($_)
					}
				) -and
				[System.IO.Path]::IsPathRooted($DatabaseServer.Server.Configuration.DatabaseSettings.BackupPath) -eq $true
			) {
				[System.IO.Path]::GetPathRoot($DatabaseServer.Server.Configuration.DatabaseSettings.BackupPath)
			} else {
				$null
			}



			######################
			# Server Checks
			######################

			# Resource Governor Enabled
			#region
			if ($DatabaseServer.Server.Management.ResourceGovernor.Enabled -eq $true) {
				$Details = "Resource Governor is enabled. Queries may be throttled. Make sure you understand how the Classifier Function is configured."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $LowPriority `
				-Category $CatPerformance `
				-Description 'Resource Governor Enabled' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb933866', $HelpUrlModifier, '.aspx')))
			}
			#endregion


			# Sysadmins
			#region
			$DatabaseServer.Server.Security.ServerRoles | Where-Object { $_.Name -ieq 'sysadmin' } | ForEach-Object {
				$_.Member | Where-Object { $_ -ine $SaLogin } | ForEach-Object {

					$Details = "Login [$($_)] is a sysadmin - meaning they can perform any activity in the server, whether they mean to or not!"

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $NullDatabaseName `
					-Priority $HighPriority `
					-Category $CatSecurity `
					-Description 'Sysadmins' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb933866', $HelpUrlModifier, '.aspx')))
				}
			}
			#endregion 


			# Security Admins
			#region
			$DatabaseServer.Server.Security.ServerRoles | Where-Object { $_.Name -ieq 'securityadmin' } | ForEach-Object {
				$_.Member | Where-Object {
					[String]$_ -ine [String]::Empty -and
					$_ -ine $SaLogin
				} | ForEach-Object {

					$Details = "Login [$($_)] is a security admin - meaning they can give themselves permission to do anything in SQL Server. This should be treated as equivalent to the sysadmin role."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $NullDatabaseName `
					-Priority $HighPriority `
					-Category $CatSecurity `
					-Description 'Security Admins' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb933866', $HelpUrlModifier, '.aspx')))
				}
			}
			#endregion


			# Jobs owned by users (i.e. not owned by "sa")
			#region
			$DatabaseServer.Agent.Jobs | Where-Object {
				$( $_ | Measure-Object).Count -gt 0 -and
				$_.General.Owner -ine $SaLogin
			} | ForEach-Object {
				$Details = "Job [$($_.General.Name)] is owned by [$($_.General.Owner)] - if their login is disabled or not available due to Active Directory problems, the job will stop working. Consider changing the login to [$SaLogin]"

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $LowPriority `
				-Category $CatSecurity `
				-Description 'Jobs Owned By Users' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms188745', $HelpUrlModifier, '.aspx')))
			}
			#endregion


			# Login Has a Blank Password
			#region
			$DatabaseServer.Server.Security.Logins | Where-Object {
				$_.HasBlankPassword -eq $true
			} | ForEach-Object {
				$Details = "Login [$($_.Name)] has a blank password, meaning anyone can log in as this account without needing credentials. This presents a serious security risk - especially if the login has elevated priveleges."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $HighPriority `
				-Category $CatSecurity `
				-Description 'Login Has a Blank Password' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms189828', $HelpUrlModifier, '.aspx')))
			}
			#endregion


			# Password Matches Login Name
			#region
			$DatabaseServer.Server.Security.Logins | Where-Object {
				$_.HasNameAsPassword -eq $true
			} | ForEach-Object {
				$Details = "The password for login [$($_.Name)] is the same as the login name, meaning anyone can log in as this account with minimal effort. This presents a serious security risk - especially if the login has elevated priveleges."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $HighPriority `
				-Category $CatSecurity `
				-Description 'Password Matches Login Name' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms189828', $HelpUrlModifier, '.aspx')))
			}
			#endregion


			# Startup Procedures
			#region
			$DatabaseServer.Server.ServerObjects.StartupProcedures | Where-Object { $_.Name } | ForEach-Object {

				$Details = "Stored procedure [master].[$($_.Schema)].[$($_.Name)] runs automatically when SQL Server starts up. You should understand exactly what this stored procedure does as it could pose a security risk."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $MediumPriority `
				-Category $CatSecurity `
				-Description 'Stored Procedure Runs at Startup' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms181720', $HelpUrlModifier, '.aspx')))
			}
			#endregion


			# Server Audits
			#region
			$DatabaseServer.Server.Security.Audits | Where-Object { $_.General.ID } | ForEach-Object {

				$Details = "SQL Server built-in audit functionality is being used by server audit: $($_.General.AuditName)"

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $NoPriority `
				-Category $CatSecurity `
				-Description 'Server Audits Running' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms181720', $HelpUrlModifier, '.aspx')))
			}
			#endregion


			# Endpoints
			#region
			$DatabaseServer.Server.ServerObjects.Endpoints | Where-Object { 
				$_.ID -and 
				$_.EndpointType -ine 'tsql'
			} | ForEach-Object {

				$Details = "SQL Server endpoints are configured. These can be used for database mirroring or Service Broker, but if you do not need them, avoid leaving them enabled.  Endpoint name: $($_.Name)"

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $LowPriority `
				-Category $CatSecurity `
				-Description 'Surface Area' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms181586', $HelpUrlModifier, '.aspx')))
			}
			#endregion


			# Server Triggers Enabled
			#region
			$DatabaseServer.Server.ServerObjects.Triggers | Where-Object {
				$_.ID -and
				$_.IsEnabled -eq $true -and
				$_.IsSystemObject -ne $true
			} | ForEach-Object {

				$Details = "Server Trigger [$($_.Name)] is enabled, so it runs every time someone logs in.  Make sure you understand what that trigger is doing - the less work it does, the better."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $MediumPriority `
				-Category $CatPerformance `
				-Description 'Server Triggers Enabled' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms189799', $HelpUrlModifier, '.aspx')))
			}
			#endregion 


			# Not All Alerts Configured
			#region
			if (
				$DatabaseServer.Agent.Alerts -and
				$(
					$DatabaseServer.Agent.Alerts | Where-Object { 
						$_.General.ID -and
						$_.General.Definition.Severity -ge 19 -and
						$_.General.Definition.Severity -le 25
					} | Measure-Object
				).Count -lt 7
			) {

				$Details = 'Not all SQL Server Agent alerts have been configured. Agent alerts are a free, easy way to get notified of server problems even before monitoring systems pick them up.'

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $MediumPriority `
				-Category $CatReliability `
				-Description 'Not All Alerts Configured' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms189531', $HelpUrlModifier, '.aspx')))
			}
			#endregion


			# Servers with Non-Default Config Values
			#region
			Get-ServerConfigurationItem -ServerConfigurationInformation $DatabaseServer.Server.Configuration | Where-Object {
				(
					@(
						#[String]::Empty
						<#
						'access check cache bucket count','access check cache quota','affinity64 I/O mask','affinity64 mask',
						'backup compression default', 'common criteria compliance enabled', 'contained database authentication',
						'EKM provider enabled'
						#>
						'xp_cmdshell'
					) -inotcontains $_.ConfigurationName
				) -and
				$_.RunningValue -ne $_.DefaultValue -and
				(
					# In some cases 'min server memory (MB)' can have a default of 0 but a running value of 16; OK to ignore this
					$_.ConfigurationName -ine 'min server memory (MB)' -or 
					(
						$_.ConfigurationName -ieq 'min server memory (MB)' -and
						$_.DefaultValue -eq 0 -and
						$_.RunningValue -ne 16 -and
						$_.RunningValue -ne 0 
					)
				)
			} | ForEach-Object {

				$Details = "The sp_configure option '$($_.ConfigurationName)' has been changed. Its default value is $($_.DefaultValue) and it has been set to $($_.RunningValue)"

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $LowPriority `
				-Category $CatInformation `
				-Description 'Non-Default Server Config' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms189631', $HelpUrlModifier, '.aspx')))
			}
			#endregion


			# xp_cmdshell Enabled
			#region
			if ($DatabaseServer.Server.Configuration.Advanced.Miscellaneous.XPCmdShellEnabled.RunningValue -eq $true) {

				$Details = "xp_cmdshell is enabled. This allows a command to be passed to the operating system for execution and can be a potential security risk. Unless you know you need this feature it is recommended that you disable it. If you are using this feature consider enabling\disabling on demand or using a proxy account to mitigate the potential risks."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $MediumPriority `
				-Category $CatSecurity `
				-Description 'xp_cmdshell Enabled' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms175046', $HelpUrlModifier, '.aspx')))
			}
			#endregion


			# Server public Permissions
			# VIEW ANY DATABASE is granted by default so go ahead and let this one slide
			#region
			$DatabaseServer.Server.Configuration.Permissions | Where-Object {
				@('Grant','Grant With Grant') -icontains $_.PermissionState -and
				$_.Grantee -ieq 'public' -and
				$_.GranteeType -ieq 'Server Role' -and
				$_.PermissionType -ine 'VIEW ANY DATABASE'
			} | ForEach-Object {

				$Details = "The public server role has been granted the $($_.PermissionType) server permission. Because every server login is a member of the public server role every login will inherit this permission."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $MediumPriority `
				-Category $CatSecurity `
				-Description 'Server public Permissions' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/cc645930', $HelpUrlModifier, '.aspx'))) 
			}
			#endregion


			# Linked Servers
			#region
			$DatabaseServer.Server.ServerObjects.LinkedServers | Where-Object { $_.General.Name } | ForEach-Object {

				$Details = if ($_.Security.LocalLogin -ieq 'sa') {
					"$($_.General.Name) is configured as a linked server. It's connecting as SA, meaning any user who queries it will get more permission than you probably want them to have."
				} else {
					"$($_.General.Name) is configured as a linked server. Check its security configuration to make sure it isn't connecting with SA or some other administrative login, because any user who queries it might get more permission than you probably want them to have."
				}

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $LowPriority `
				-Category $CatSecurity `
				-Description 'Linked Server Configured' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms188279', $HelpUrlModifier, '.aspx')))
			}
			#endregion


			# No Operators Configured/Enabled
			#region
			$DatabaseServer.Agent.Alerts | Where-Object {
				$_.General.ID -and
				$_.General.IsEnabled -and
				(
					$($_.Response.NotifyOperators | Measure-Object).Count -eq 0 -or
					(
						$_.Response.NotifyOperators | Where-Object {
							$_.OperatorId -and
							(
								(
									$_.UseNetSend -eq $true -and 
									$_.HasNetSend -ne $true
								) -or 
								$_.UseNetSend -ne $true
							) -and
							(
								(
									$_.UseEmail -eq $true -and 
									$_.HasEmail -ne $true
								) -or 
								$_.UseEmail -ne $true
							) -and
							(
								(
									$_.UsePager -eq $true -and 
									$_.HasPager -ne $true
								) -or 
								$_.UsePager -ne $true
							)
						} | Measure-Object
					).Count -gt 0
				)
			} | ForEach-Object {

				$Details = 'SQL Server Agent alerts have been configured but they either do not notify anyone or they do not take any action. Agent alerts are a free, easy way to get notified of server problems even before monitoring systems pick them up.'

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $LowPriority `
				-Category $CatReliability `
				-Description 'Alerts Configured without Follow Up' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms178616', $HelpUrlModifier, '.aspx')))
			}
			#endregion


			# No Alerts for Corruption
			# Is this really right? sp_Blitz makes it possible for an alert to be set up 
			# for 823 but not 824 and 825 and the check would pass
			# ...but don't we want to alert when all three aren't set up?
			#region
			if (
				$DatabaseServer.Agent.Alerts -and
				$(
					$_.Agent.Alerts | Where-Object { 
						$_.General.ID -and
						@(823,824,825) -contains $_.General.Definition.ErrorNumber
					} | Measure-Object
				).Count -lt 3
			) {

				$Details = 'SQL Server Agent alerts do not exist for errors 823, 824, and 825. These three errors can give you notification about early hardware failure. Enabling them can prevent you a lot of heartbreak.'

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $MediumPriority `
				-Category $CatReliability `
				-Description 'No Alerts for Corruption' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms178616', $HelpUrlModifier, '.aspx'))) 
			}
			#endregion


			# No Alerts for Sev 19-25
			#region
			if (
				$DatabaseServer.Agent.Alerts -and 
				$(
					$_.Agent.Alerts | Where-Object { 
						$_.General.ID -and
						$_.General.Definition.Severity -ge 19 -and
						$_.General.Definition.Severity -le 25
					} | Measure-Object
				).Count -eq 0
			) {

				$Details = 'SQL Server Agent alerts do not exist for severity levels 19 through 25. These are some very severe SQL Server errors. Knowing that these are happening may let you recover from errors faster.'

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $HighPriority `
				-Category $CatReliability `
				-Description 'No Alerts for Sev 19-25' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms178616', $HelpUrlModifier, '.aspx'))) 
			}
			#endregion


			# Alerts Disabled
			#region
			$DatabaseServer.Agent.Alerts | Where-Object {
				$_.General.ID -and
				$_.General.IsEnabled -ne $true
			} | ForEach-Object {

				$Details = "The following alert is disabled, please review and enable if desired: $($_.General.Name)"

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $LowPriority `
				-Category $CatReliability `
				-Description 'Alerts Disabled' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms178616', $HelpUrlModifier, '.aspx'))) 
			}
			#endregion


			# No Operators Configured/Enabled
			#region
			if (
				$DatabaseServer.Agent.Operators -and
				$(
					$_.Agent.Operators | Where-Object { 
						$_.General.ID -and
						$_.General.IsEnabled
					} | Measure-Object
				).Count -eq 0
			) {

				$Details = 'No SQL Server Agent operators (emails) have been configured.'

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $LowPriority `
				-Category $CatReliability `
				-Description 'No Operators Configured/Enabled' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms186747', $HelpUrlModifier, '.aspx'))) 
			}
			#endregion


			# Max Memory Set Too High
			#region
			if ($DatabaseServer.Server.Configuration.Memory.MaxServerMemoryMB.RunningValue -gt $DatabaseServer.Server.Configuration.General.MemoryMB) {

				$Details = "SQL Server max memory is set to $($DatabaseServer.Server.Configuration.Memory.MaxServerMemoryMB.RunningValue) megabytes, but the server only has $($DatabaseServer.Server.Configuration.General.MemoryMB) megabytes.  SQL Server may drain the system dry of memory, and under certain conditions, this can cause Windows to swap to disk."

				# This could be reliability as well - server may BSOD and\or restart due to lack of memory!
				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $HighPriority `
				-Category $CatPerformance `
				-Description 'Max Memory Set Too High' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms178067', $HelpUrlModifier, '.aspx'))) 
			}
			#endregion


			# Memory Dangerously Low
			#region
			if ($( ($DatabaseServer.Server.Configuration.General.MemoryMB * 1KB) - $DatabaseServer.Server.Configuration.General.MemoryInUseKB) -lt 262144) {

				$Details = "Although available memory is $((($_.Server.Configuration.General.MemoryMB * 1KB) - $_.Server.Configuration.General.MemoryInUseKB) / 1KB) MB, only $($_.Server.Configuration.General.MemoryMB) MB of memory are present.  As the server runs out of memory, there is danger of swapping to disk, which will kill performance."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $HighPriority `
				-Category $CatPerformance `
				-Description 'Memory Dangerously Low' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms178067', $HelpUrlModifier, '.aspx'))) 
			}
			#endregion


			# Cluster Node
			#region
			if ($DatabaseServer.Server.Configuration.General.IsClustered -eq $true) {

				$Details = 'This is a node in a cluster.'

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $NoPriority `
				-Category $CatAvailability `
				-Description 'Cluster Node' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms189134', $HelpUrlModifier, '.aspx'))) 
			}
			#endregion


			# SQL Agent Job Runs at Startup
			#region
			$DatabaseServer.Agent.Jobs | Where-Object {
				(
					$_.Schedules | Where-Object { 
						$_.Id -and
						$_.Description -ieq 'Start automatically when SQL Server Agent starts'
					} | Measure-Object
				).Count -gt 0
			} | ForEach-Object {

				$Details = "Job [$($_.General.Name)] runs automatically when SQL Server Agent starts up. Make sure you know exactly what this job is doing, because it could pose a security risk."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $MediumPriority `
				-Category $CatSecurity `
				-Description 'SQL Agent Job Runs at Startup' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms178560', $HelpUrlModifier, '.aspx')))
			}
			#endregion


			# Unusual SQL Server Edition
			#region
			if (
				$DatabaseServer.Server.Configuration.General.Edition -and
				$DatabaseServer.Server.Configuration.General.Edition -inotlike '*Standard*' -and
				$DatabaseServer.Server.Configuration.General.Edition -inotlike '*Enterprise*' -and
				$DatabaseServer.Server.Configuration.General.Edition -inotlike '*Developer*'
			) {

				$Details = "This server is using $($_.Server.Configuration.General.Edition) edition, which is capped at low amounts of CPU and memory."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $LowPriority `
				-Category $CatPerformance `
				-Description 'Unusual SQL Server Edition' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms178560', $HelpUrlModifier, '.aspx')))
			}
			#endregion


			# No failsafe operator configured
			#region
			if (
				$DatabaseServer.Agent.Configuration -and
				-not $DatabaseServer.Agent.Configuration.AlertSystem.FailSafeOperator.Operator
			) {

				$Details = 'No failsafe operator is configured on this server. This is a good idea to set up in case there are issues with the [MSDB] database that prevent alerting.'

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $MediumPriority `
				-Category $CatReliability `
				-Description 'No failsafe operator configured' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms175514', $HelpUrlModifier, '.aspx'))) 
			}
			#endregion


			# Global Trace Flags
			#region
			$DatabaseServer.Server.Management.TraceFlags | Where-Object { 
				$_.TraceFlag -and
				$_.IsGlobal
			} | ForEach-Object {

				$Details = "Trace flag $($_.TraceFlag) is enabled globally."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $NoPriority `
				-Category $CatInformation `
				-Description 'TraceFlag On' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms188396', $HelpUrlModifier, '.aspx'))) 
			}
			#endregion


			# Shrink Database Job
			#region
			$DatabaseServer.Agent.Jobs | ForEach-Object {
				$JobName = $_.General.Name

				$_.Steps | Where-Object { 
					$_.Id -and
					(
						$_.General.Command -ilike '*SHRINKDATABASE*' -or
						$_.General.Command -ilike '*SHRINKFILE*'
					)
				} | ForEach-Object {

					$Details = "In the [$JobName] job, step [$($_.General.StepName)] has SHRINKDATABASE or SHRINKFILE, which may be causing database fragmentation."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $NullDatabaseName `
					-Priority $MediumPriority `
					-Category $CatPerformance `
					-Description 'Shrink Database Job' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms189493', $HelpUrlModifier, '.aspx')))
				}
			}
			#endregion


			# Non-Active Server Config
			#region
			Get-ServerConfigurationItem -ServerConfigurationInformation $DatabaseServer.Server.Configuration | Where-Object {
				@(
					[String]::Empty <#,
					'access check cache bucket count','access check cache quota','affinity64 I/O mask','affinity64 mask',
					'backup compression default', 'common criteria compliance enabled', 'contained database authentication',
					'EKM provider enabled'#>
				) -inotcontains $_.ConfigurationName -and
				$_.RunningValue -ne $_.ConfiguredValue
			} | ForEach-Object {

				$Details = "The sp_configure option '$($_.ConfigurationName)' ($($_.FriendlyName)) isn't running under its set value. Its set value is $($_.ConfiguredValue) and its running value is $($_.RunningValue). When someone does a RECONFIGURE or restarts the instance, this setting will start taking effect."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $HighPriority `
				-Category $CatReliability `
				-Description 'Non-Active Server Config' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms188787', $HelpUrlModifier, '.aspx')))
			}
			#endregion


			# @@Servername not set
			#region
			$DatabaseServer.Server.Configuration.General | Where-Object {
				[String]::IsNullOrEmpty($_.GlobalName) -eq $true
				#$_.GlobalName -eq $null -or
				#$_.GlobalName.Length -eq 0
			} | ForEach-Object {

				$Details = "@@Servername variable is null. Correct by executing ""sp_addserver '$($_.Name)', local"""

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $LowPriority `
				-Category $CatInformation `
				-Description '@@Servername not set' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms174411', $HelpUrlModifier, '.aspx'))) 
			}
			#endregion


			# Suboptimal Operating System Power Plan
			#region
			$DatabaseServer.Machine.OperatingSystem.Settings.PowerPlan | Where-Object { 
				$_.IsActive -eq $true -and
				$_.PlanName -ine 'High performance'
			} | ForEach-Object { 

				$Details = "The Windows Power Plan is set to ""$($_.PlanName)"" which can result in decreased performance. Consider switching to the ""High performance"" plan instead."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $LowPriority `
				-Category $CatPerformance `
				-Description 'Non-Optimal Windows Power Plan' `
				-Details $Details `
				-URL 'http://technet.microsoft.com/en-us/library/dd744398.aspx'

			}
			#endregion


			# Increase or Disable Blocked Process Threshold
			#region
			if (
				$_.Server.Configuration.Advanced.Miscellaneous.BlockedProcessThresholdSeconds.RunningValue -gt 0 -and
				$_.Server.Configuration.Advanced.Miscellaneous.BlockedProcessThresholdSeconds.RunningValue -le 4
			) {

				$Details = "The sp_configure option 'blocked process threshold' is set to $($_.Server.Configuration.Advanced.Miscellaneous.BlockedProcessThresholdSeconds.RunningValue). Values 1 to 4 can cause the deadlock monitor to run constantly and should only be used for troubleshooting (and especially not used long term in a production environment)."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $HighPriority `
				-Category $CatPerformance `
				-Description 'Increase or Disable Blocked Process Threshold' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb402879', $HelpUrlModifier, '.aspx'))) 
			}
			#endregion


			# Keep the Locks Configuration Option Default Value
			#region
			if (
				-not [String]::IsNullOrEmpty($_.Server.Configuration.Advanced.Parallelism.Locks.RunningValue) -and
				$_.Server.Configuration.Advanced.Parallelism.Locks.RunningValue -ne 0
			) {

				$Details = "The sp_configure option 'Locks' is set to $($_.Server.Configuration.Advanced.Parallelism.Locks.RunningValue). Nonzero values can cause batch jobs to stop and ""out of locks"" error messages to be generated if the set value is exceeded."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $HighPriority `
				-Category $CatPerformance `
				-Description 'Keep the Locks Configuration Option Default Value' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb402936', $HelpUrlModifier, '.aspx'))) 
			}
			#endregion


			# Lightweight Pooling Enabled
			#region
			if ($_.Server.Configuration.Processor.UseWindowsFibers.RunningValue -eq $true) {

				$Details = "The sp_configure option 'lightweightpooling' is enabled. When enabled, SQL Server uses fiber mode scheduling. This setting is only intended for specific circumstances and rarely improced performance or scalability on a typical system."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $HighPriority `
				-Category $CatPerformance `
				-Description 'Lightweight Pooling Enabled' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb402857', $HelpUrlModifier, '.aspx'))) 
			}
			#endregion


			# Max Degree Of Parallelism set too high
			#region
			if ($_.Server.Configuration.Advanced.Parallelism.MaxDegreeOfParallelism.RunningValue -gt 8) {

				$Details = "The sp_configure option 'max degree of parallelism' is set to $($_.Server.Configuration.Advanced.Parallelism.MaxDegreeOfParallelism.RunningValue). Values higher than 8 may lead to unexpected resource consumption and result in decreased performance."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $HighPriority `
				-Category $CatPerformance `
				-Description 'Max Degree Of Parallelism' `
				-Details $Details `
				-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb402932', $HelpUrlModifier, '.aspx'))) 
			}
			#endregion


			# SQL Server service account is a member of the Administrators group
			#region
			if (
				$DatabaseServer.Machine.OperatingSystem.Users.LocalGroups | Where-Object { 
					$_.Name -ieq 'Administrators' -and
					$_.Members -icontains $ServerServiceAccount
				}
			) {

				$Details = "The SQL Server Database Engine service account is a member of the machine's Administrators group. This is more access than necessary for the service to work and increases the security risks if the account is compromised."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $MediumPriority `
				-Category $CatSecurity `
				-Description 'Elevated Service Account Permissions' `
				-Details $Details `
				-URL 'http://support.microsoft.com/kb/2160720'

			}
			#endregion


			# SQL Agent service account is a member of the Administrators group
			#region
			if (
				$DatabaseServer.Machine.OperatingSystem.Users.LocalGroups | Where-Object { 
					$_.Name -ieq 'Administrators' -and
					$_.Members -icontains $AgentServiceAccount
				}
			) {

				$Details = "The SQL Server Agent service account is a member of the machine's Administrators group. This is more access than necessary for the service to work and increases the security risks if the account is compromised."

				Get-AssessmentFinding -ServerName $ServerName `
				-DatabaseName $NullDatabaseName `
				-Priority $MediumPriority `
				-Category $CatSecurity `
				-Description 'Elevated Service Account Permissions' `
				-Details $Details `
				-URL 'http://support.microsoft.com/kb/2160720'

			}
			#endregion


			# Check Default Data, Log, and Backup paths for configuration issues
			if (-not [String]::IsNullOrEmpty($DefaultDataPathRoot)) {

				# Default data file drive using C:
				#region
				if ($DefaultDataPathRoot -ilike 'C:*') {

					$Details = "The Default Data Path is configured to use the C: drive. This can result in newly created databases with data files on the C: drive which runs the risk of crashing the server when it runs out of space."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $NullDatabaseName `
					-Priority $MediumPriority `
					-Category $CatReliability `
					-Description 'Default Path Configuration' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Default data file drive is the same as the default log file drive
				#region
				if ($DefaultDataPathRoot -ieq $DefaultLogPathRoot) {

					$Details = "The Default Log Path is configured for the same drive ($DefaultDataPathRoot) as the Default Data Path. This may result in multiple file types on the same drive when new databases are created which can negatively impact performance."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $NullDatabaseName `
					-Priority $MediumPriority `
					-Category $CatPerformance `
					-Description 'Default Path Configuration' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Default data file drive is the same as the default backup drive
				#region
				if ($DefaultDataPathRoot -ieq $DefaultBackupPathRoot) {

					$Details = "The Default Backup Path is configured for the same drive ($DefaultDataPathRoot) as the Default Data Path. When backups to disk are performed the backup file may be located on the same drive as the log file(s) being backed up. This is not a good idea."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $NullDatabaseName `
					-Priority $MediumPriority `
					-Category $CatRecovery `
					-Description 'Default Path Configuration' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))

				}
				#endregion

				# Default data file drive not found on server
				#region
				if (
					$DefaultDataPathRoot -ilike '[a-z][:]\*' -and
					$DatabaseServer.Machine.Hardware.Storage.DiskDrive.DeviceID -and
					-not (
						$DatabaseServer.Machine.Hardware.Storage.DiskDrive | Where-Object {
							$_.Partitions | Where-Object {
								$_.LogicalDisks | Where-Object {
									$DefaultDataPathRoot -ieq [String]::Concat($_.Caption, [System.IO.Path]::DirectorySeparatorChar)
								}
							}
						}
					)
				) {

					$Details = "The Default Data Path is set to a drive ($DefaultDataPathRoot) that was not found on the server. This can cause problems with creating new databases and may even prevent Cumulative Updates from installing."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $NullDatabaseName `
					-Priority $LowPriority `
					-Category $CatReliability `
					-Description 'Default Path Configuration' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))

				}
				#endregion
			}

			if (-not [String]::IsNullOrEmpty($DefaultLogPathRoot)) {

				# Default log file drive using C:
				#region
				if ($DefaultLogPathRoot -ilike 'C:*') {

					$Details = "The Default Log Path is configured to use the C: drive. This can result in newly created databases with log files on the C: drive which runs the risk of crashing the server when it runs out of space."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $NullDatabaseName `
					-Priority $MediumPriority `
					-Category $CatReliability `
					-Description 'Default Path Configuration' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Default log file drive is the same as the default backup drive
				#region
				if ($DefaultLogPathRoot -ieq $DefaultBackupPathRoot) {

					$Details = "The Default Backup Path is configured for the same drive ($DefaultLogPathRoot) as the Default Log Path. When backups to disk are performed the backup file may be located on the same drive as the log file(s) being backed up. This is not a good idea."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $NullDatabaseName `
					-Priority $MediumPriority `
					-Category $CatRecovery `
					-Description 'Default Path Configuration' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))

				}
				#endregion

				# Default log file drive not found on server
				#region
				if (
					$DefaultLogPathRoot -ilike '[a-z][:]\*' -and
					$DatabaseServer.Machine.Hardware.Storage.DiskDrive.DeviceID -and
					-not (
						$DatabaseServer.Machine.Hardware.Storage.DiskDrive | Where-Object {
							$_.Partitions | Where-Object {
								$_.LogicalDisks | Where-Object {
									$DefaultLogPathRoot -ieq [String]::Concat($_.Caption, [System.IO.Path]::DirectorySeparatorChar)
								}
							}
						}
					)
				) {

					$Details = "The Default Log Path is set to a drive ($DefaultLogPathRoot) that was not found on the server. This can cause problems with creating new databases and may even prevent Cumulative Updates from installing."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $NullDatabaseName `
					-Priority $LowPriority `
					-Category $CatReliability `
					-Description 'Default Path Configuration' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))

				}
				#endregion
			}

			if (-not [String]::IsNullOrEmpty($DefaultBackupPathRoot)) {

				# Default backup file drive using C:
				#region
				if ($DefaultBackupPathRoot -ilike 'C:*') {

					$Details = "The Default Backup Path is configured to use the C: drive. This can result in backups on the C: drive which runs the risk of crashing the server when it runs out of space."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $NullDatabaseName `
					-Priority $MediumPriority `
					-Category $CatReliability `
					-Description 'Default Path Configuration' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Default backup file drive not found on server
				#region
				if (
					$DefaultBackupPathRoot -ilike '[a-z][:]\*' -and
					$DatabaseServer.Machine.Hardware.Storage.DiskDrive.DeviceID -and
					-not (
						$DatabaseServer.Machine.Hardware.Storage.DiskDrive | Where-Object {
							$_.Partitions | Where-Object {
								$_.LogicalDisks | Where-Object {
									$DefaultBackupPathRoot -ieq [String]::Concat($_.Caption, [System.IO.Path]::DirectorySeparatorChar)
								}
							}
						}
					)
				) {
					$Details = "The Default Backup Path is set to a drive ($DefaultBackupPathRoot) that was not found on the server. This can cause problems with creating new databases and may even prevent Cumulative Updates from installing."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $NullDatabaseName `
					-Priority $LowPriority `
					-Category $CatReliability `
					-Description 'Default Path Configuration' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))
				}
				#endregion
			}


			# Check Default Paths for validity
			if (-not [String]::IsNullOrEmpty($DatabaseServer.Server.Configuration.General.Name)) {

				# Default Data Path
				#region
				if (
					[String]::IsNullOrEmpty($DatabaseServer.Server.Configuration.DatabaseSettings.DataPath) -or 
					[String]::IsNullOrEmpty($DatabaseServer.Server.Configuration.DatabaseSettings.DataPath.Replace(' ', [String]::Empty))
				) {

					$Details = "The Default Data Path has not been defined. This can cause problems with creating new databases and may even prevent Cumulative Updates from installing."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $NullDatabaseName `
					-Priority $LowPriority `
					-Category $CatReliability `
					-Description 'Default Path Configuration' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))

				} else {

					if ([String]::IsNullOrEmpty($DefaultDataPathRoot)) {

						$Details = "The Default Data Path points to an invalid location. This can cause problems with creating new databases and may even prevent Cumulative Updates from installing."

						Get-AssessmentFinding -ServerName $ServerName `
						-DatabaseName $NullDatabaseName `
						-Priority $LowPriority `
						-Category $CatReliability `
						-Description 'Default Path Configuration' `
						-Details $Details `
						-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))

					} elseif ($DatabaseServer.Server.Configuration.DatabaseSettings.DataPath -ilike '\\[a-z]*') {

						$Details = "The Default Data Path is a UNC share. Although supported in 2008 R2 and higher placing data files on a UNC share may have significant performance drawbacks."

						Get-AssessmentFinding -ServerName $ServerName `
						-DatabaseName $NullDatabaseName `
						-Priority $LowPriority `
						-Category $CatPerformance `
						-Description 'Default Path Configuration' `
						-Details $Details `
						-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))

					}
				}
				#endregion

				# Default Log Path
				#region
				if (
					[String]::IsNullOrEmpty($DatabaseServer.Server.Configuration.DatabaseSettings.LogPath) -or 
					[String]::IsNullOrEmpty($DatabaseServer.Server.Configuration.DatabaseSettings.LogPath.Replace(' ', [String]::Empty))
				) {

					$Details = "The Default Log Path has not been defined. This can cause problems with creating new databases and may even prevent Cumulative Updates from installing."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $NullDatabaseName `
					-Priority $LowPriority `
					-Category $CatReliability `
					-Description 'Default Path Configuration' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))

				} else {
					if ([String]::IsNullOrEmpty($DefaultLogPathRoot)) {

						$Details = "The Default Log Path points to an invalid location. This can cause problems with creating new databases and may even prevent Cumulative Updates from installing."

						Get-AssessmentFinding -ServerName $ServerName `
						-DatabaseName $NullDatabaseName `
						-Priority $LowPriority `
						-Category $CatReliability `
						-Description 'Default Path Configuration' `
						-Details $Details `
						-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))

					} elseif ($DatabaseServer.Server.Configuration.DatabaseSettings.LogPath -ilike '\\[a-z]*') {

						$Details = "The Default Log Path is a UNC share. Although supported in 2008 R2 and higher placing log files on a UNC share may have significant performance drawbacks."

						Get-AssessmentFinding -ServerName $ServerName `
						-DatabaseName $NullDatabaseName `
						-Priority $LowPriority `
						-Category $CatPerformance `
						-Description 'Default Path Configuration' `
						-Details $Details `
						-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))

					}
				}
				#endregion

				# Default Backup Path
				#region
				if (
					[String]::IsNullOrEmpty($DatabaseServer.Server.Configuration.DatabaseSettings.BackupPath) -or 
					[String]::IsNullOrEmpty($DatabaseServer.Server.Configuration.DatabaseSettings.BackupPath.Replace(' ', [String]::Empty))
				) {

					$Details = "The Default Backup Path has not been defined. This can cause problems with backing up databases under certain conditions."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $NullDatabaseName `
					-Priority $LowPriority `
					-Category $CatReliability `
					-Description 'Default Path Configuration' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))

				} else {
					if ([String]::IsNullOrEmpty($DefaultBackupPathRoot)) {

						$Details = "The Default Backup Path points to an invalid location. This can cause problems with backing up databases under certain conditions."

						Get-AssessmentFinding -ServerName $ServerName `
						-DatabaseName $NullDatabaseName `
						-Priority $LowPriority `
						-Category $CatReliability `
						-Description 'Default Path Configuration' `
						-Details $Details `
						-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))

					} 
				}
				#endregion

			}


			######################
			# Disk Checks
			######################

			$DatabaseServer.Machine.Hardware.Storage.DiskDrive | ForEach-Object {
				foreach ($Partition in $_.Partitions) {
					foreach ($LogicalDisk in $Partition.LogicalDisks) { 

						$PartitionRoot = [String]::Concat($LogicalDisk.Caption, [System.IO.Path]::DirectorySeparatorChar)

						# Do further analysis if any database files exist on the current drive
						if (
							$DatabaseServer.Server.Databases | Where-Object {
								$_.Properties.Files.DatabaseFiles | Where-Object { 
									$_.ID -and
									[System.IO.Path]::GetPathRoot($_.Path) -ieq $PartitionRoot
								}

							}
						) {

							# Nonaligned partitions
							#region
							if ([int]($Partition.StartingOffsetBytes / $LogicalDisk.AllocationUnitSizeBytes) -ne ($Partition.StartingOffsetBytes / $LogicalDisk.AllocationUnitSizeBytes)) {

								$Details = "Drive $($LogicalDisk.Caption) contains SQL Server files and has a nonaligned partition. If the drive is part of a RAID configuration this can have a significant performance impact when partition writes overlap stripe boundaries."

								Get-AssessmentFinding -ServerName $ServerName `
								-DatabaseName $NullDatabaseName `
								-Priority $LowPriority `
								-Category $CatPerformance `
								-Description 'Nonaligned Storage Partition' `
								-Details $Details `
								-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd758814', $HelpUrlModifier, '.aspx'))) 

							}
							#endregion


							# Partition Allocation Unit Size
							#region
							if ($LogicalDisk.AllocationUnitSizeBytes -ne 65536) {

								$Details = "Drive $($LogicalDisk.Caption) contains SQL Server files and has a $($LogicalDisk.AllocationUnitSizeBytes / 1KB) KB allocation unit size. While your mileage may vary, it is generally recommended that a 64 KB allocation unit size is optimal for SQL Server I/O."

								Get-AssessmentFinding -ServerName $ServerName `
								-DatabaseName $NullDatabaseName `
								-Priority $LowPriority `
								-Category $CatPerformance `
								-Description 'Partition Allocation Unit Size' `
								-Details $Details `
								-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd758814', $HelpUrlModifier, '.aspx'))) 

							}
							#endregion

						}

						# Multiple File Types On Same Drive
						#region
						if (
							$(
								$DatabaseServer.Server.Databases | ForEach-Object {
									$_.Properties.Files.DatabaseFiles | Where-Object { 
										$_.ID -and
										[System.IO.Path]::GetPathRoot($_.Path) -ieq $PartitionRoot
									}
								} | Group-Object -Property FileType | Measure-Object
							).Count -gt 1
						) {

							$Details = "Drive $($LogicalDisk.Caption) contains multiple types of SQL Server files. This can negatively affect performance since each file type has different I/O characteristics."

							Get-AssessmentFinding -ServerName $ServerName `
							-DatabaseName $NullDatabaseName `
							-Priority $LowPriority `
							-Category $CatPerformance `
							-Description 'Multiple File Types On Same Drive' `
							-Details $Details `
							-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms179316', $HelpUrlModifier, '.aspx'))) 

						}
						#endregion

						# Default data path root is on a drive that holds non-data files
						#region
						if (
							$DefaultDataPathRoot -ieq $PartitionRoot -and
							(
								$DatabaseServer.Server.Databases | Where-Object {
									$_.Properties.Files.DatabaseFiles | Where-Object { 
										$_.ID -and
										[System.IO.Path]::GetPathRoot($_.Path) -ieq $PartitionRoot -and
										$_.FileType -ine 'Rows Data'
									}
								}
							)
						) {

							$Details = "The Default Data Path is configured for a drive ($DefaultDataPathRoot) that currently holds database files of another type. This may result in multiple file types on the same drive when new databases are created which can negatively impact performance."

							Get-AssessmentFinding -ServerName $ServerName `
							-DatabaseName $NullDatabaseName `
							-Priority $MediumPriority `
							-Category $CatPerformance `
							-Description 'Default Path Configuration' `
							-Details $Details `
							-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))

						}
						#endregion

						# Default log path root is on a drive that holds non-log files
						#region
						if (
							$DefaultLogPathRoot -ieq $PartitionRoot -and
							(
								$DatabaseServer.Server.Databases | Where-Object {
									$_.Properties.Files.DatabaseFiles | Where-Object { 
										$_.ID -and
										[System.IO.Path]::GetPathRoot($_.Path) -ieq $PartitionRoot -and
										$_.FileType -ine 'Log'
									}
								}
							)
						) {

							$Details = "The Default Log Path is configured for a drive ($DefaultDataPathRoot) that currently holds database files of another type. This may result in multiple file types on the same drive when new databases are created which can negatively impact performance."

							Get-AssessmentFinding -ServerName $ServerName `
							-DatabaseName $NullDatabaseName `
							-Priority $MediumPriority `
							-Category $CatPerformance `
							-Description 'Default Path Configuration' `
							-Details $Details `
							-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))

						}
						#endregion

						# Default backup path root is on a drive that holds database files
						#region
						if (
							$DefaultBackupPathRoot -ieq $PartitionRoot -and
							(
								$DatabaseServer.Server.Databases | Where-Object {
									$_.Properties.Files.DatabaseFiles | Where-Object { 
										$_.ID -and
										[System.IO.Path]::GetPathRoot($_.Path) -ieq $PartitionRoot
									}
								}
							)
						) {

							$Details = "The Default Backup Path is configured for a drive ($DefaultDataPathRoot) that currently holds database files. If a database is backed up without specifying a file path the backup file may be located on the same drive as the file(s) being backed up. This is not a good idea."

							Get-AssessmentFinding -ServerName $ServerName `
							-DatabaseName $NullDatabaseName `
							-Priority $MediumPriority `
							-Category $CatPerformance `
							-Description 'Default Path Configuration' `
							-Details $Details `
							-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/dd206993', $HelpUrlModifier, '.aspx')))

						}
						#endregion

					}
				}
			}




			######################
			# Database Checks
			######################

			#region
			$DatabaseServer.Server.Databases | ForEach-Object {

				$DatabaseName = $_.Name

				# DB has never had a FULL backup
				# Exclude tempdb and databases that aren't in "normal" status
				#region
				if (
					@('tempdb') -notcontains $_.Name -and
					$_.Properties.General.Database.Status -ieq 'normal' -and
					$_.Properties.General.Database.IsDatabaseSnapshot -ne $true -and
					$_.Properties.General.Backup.LastFullBackupDate -eq $null 
				) {

					$Details = "Database has never been backed up."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $HighPriority `
					-Category $CatRecovery `
					-Description 'Backup' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms186865', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Databases that have had a FULL backup but not a DIFFERENTIAL
				# This is interesting, but not really a major concern if it's not happening
				# Exclude master, model, tempdb, and msdb
				#region
				# 			if (
				# 				@('master','model','msdb','tempdb') -notcontains $_.Name -and
				# 				$_.Properties.General.Database.Status -ieq 'normal' -and
				# 				$_.Properties.General.Database.IsDatabaseSnapshot -ne $true -and
				# 				$_.Properties.General.Backup.LastFullBackupDate -ne $null -and
				# 				$_.Properties.General.Backup.LastDifferentialBackupDate -eq $null
				# 			) { 
				# 				Write-Output $_.Name 
				# 			}
				#endregion

				# No FULL backup in the week prior to the scan date
				# Exclude tempdb
				#region
				if (
					$_.Name -ine 'tempdb' -and
					$_.Properties.General.Database.Status -ieq 'normal' -and
					$_.Properties.General.Database.IsDatabaseSnapshot -ne $true -and
					$_.Properties.General.Backup.LastFullBackupDate -ne $null -and
					$_.Properties.General.Backup.LastFullBackupDate.CompareTo($($ScanDateLocal).AddDays(-7)) -le 0
				) {

					$Details = "Database last backed up: $($_.Properties.General.Backup.LastFullBackupDate.ToString('G')), which is more than a week prior to the inventory scan on $($ScanDateLocal.ToString('G'))."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $HighPriority `
					-Category $CatRecovery `
					-Description 'Backup' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms186865', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# FULL or BULKLOGGED recovery model but never had a log backup
				# Exclude tempdb and model
				#region
				if (
					@('model','tempdb') -notcontains $_.Name -and
					$_.Properties.General.Database.Status -ieq 'normal' -and
					$_.Properties.General.Database.IsDatabaseSnapshot -ne $true -and
					@('FULL','BULKLOGGED') -contains $_.Properties.Options.RecoveryModel -and
					$_.Properties.General.Backup.LastLogBackupDate -eq $null
				) {

					$Details = "Database is in the $($_.Properties.Options.RecoveryModel) recovery mode but has never had a log backup."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $HighPriority `
					-Category $CatRecovery `
					-Description 'Full Recovery Mode w/o Log Backups' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms186865', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# FULL or BULKLOGGED recovery model that without a log backup in the week prior to the scan date
				# Exclude tempdb
				#region
				if (
					$_.Name -ine 'tempdb' -and
					$_.Properties.General.Database.Status -ieq 'normal' -and
					$_.Properties.General.Database.IsDatabaseSnapshot -ne $true -and
					@('FULL','BULKLOGGED') -contains $_.Properties.Options.RecoveryModel -and
					$_.Properties.General.Backup.LastLogBackupDate -ne $null -and
					$_.Properties.General.Backup.LastLogBackupDate.CompareTo($($ScanDateLocal).AddDays(-7)) -le 0
				) {

					$Details = "Database is in $($_.Properties.Options.RecoveryModel) recovery mode but has not had a log backup in the last week prior to the inventory scan on $($ScanDateLocal.ToString('G'))."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $HighPriority `
					-Category $CatRecovery `
					-Description 'Full Recovery Mode w/o Log Backups' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms186865', $HelpUrlModifier, '.aspx')))
				} 
				#endregion

				# Autoclose Enabled
				#region
				if ($_.Properties.Options.OtherOptions.Automatic.AutoClose -eq $true) {

					$Details = "Database has auto-close enabled.  This setting can dramatically decrease performance."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $HighPriority `
					-Category $CatPerformance `
					-Description 'Auto-Close Enabled' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms135094', $HelpUrlModifier, '.aspx')))
				}
				#endregion 

				# Autoshrink Enabled
				#region
				if ($_.Properties.Options.OtherOptions.Automatic.AutoShrink -eq $true) {

					$Details = "Database has auto-shrink enabled.  This setting can dramatically decrease performance."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $HighPriority `
					-Category $CatPerformance `
					-Description 'Auto-Shrink Enabled' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms136209', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Page Verification Not Optimal (all versions)
				#region
				if (
					$_.Name -ine 'tempdb' -and
					$_.Properties.Options.OtherOptions.Recovery.PageVerify -ine 'CHECKSUM'
				) {

					$Details = "Database has $($_.Properties.Options.OtherOptions.Recovery.PageVerify) for page verification.  SQL Server may have a harder time recognizing and recovering from storage corruption.  Consider using CHECKSUM instead."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $MediumPriority `
					-Category $CatReliability `
					-Description 'Page Verification Not Optimal'`
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb402873', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Auto-Create Stats Disabled
				#region
				if ($_.Properties.Options.OtherOptions.Automatic.AutoCreateStatistics -ne $true) {

					$Details = "Database has auto-create-stats disabled. SQL Server uses statistics to build better execution plans, and without the ability to automatically create more, performance may suffer."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $LowPriority `
					-Category $CatPerformance `
					-Description 'Auto-Create Stats Disabled' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb522682', $HelpUrlModifier, '.aspx')))
				}
				#endregion 

				# Auto-Update Stats Disabled
				#region
				if ($_.Properties.Options.OtherOptions.Automatic.AutoUpdateStatistics -ne $true) {

					$Details = "Database has auto-update-stats disabled. SQL Server uses statistics to build better execution plans, and without the ability to automatically create more, performance may suffer."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $LowPriority `
					-Category $CatPerformance `
					-Description 'Auto-Update Stats Disabled' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb522682', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Auto-Update Stats Async Enabled
				#region
				if ($_.Properties.Options.OtherOptions.Automatic.AutoUpdateStatisticsAsync -eq $true) {

					$Details = "Database has auto-update-stats-async enabled. When SQL Server gets a query for a table with out-of-date statistics, it will run the query with the stats it has - while updating stats to make later queries better. The initial run of the query may suffer, though."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $LowPriority `
					-Category $CatPerformance `
					-Description 'Auto-Update Stats Enabled' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb522682', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Forced Parameterization
				#region
				if ($_.Properties.Options.OtherOptions.Miscellaneous.Parameterization -ieq 'forced') {

					$Details = "Database has forced parameterization enabled. SQL Server will aggressively reuse query execution plans even if the applications do not parameterize their queries.  This can be a performance booster with some programming languages, or it may use universally bad execution plans when better alternatives are available for certain parameters."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $LowPriority `
					-Category $CatPerformance `
					-Description 'Forced Parameterization On' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb522682', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Replication in use
				#region
				if (
					$_.Properties.General.Database.ReplicationOptions -and
					$_.Properties.General.Database.ReplicationOptions -ine 'none'
				) {

					$Details = "Database is a replication publisher, subscriber, or distributor."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $NoPriority `
					-Category $CatInformation `
					-Description 'Replication In Use' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms151198', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Date Correlation Enabled
				#region
				if ($_.Properties.Options.OtherOptions.Miscellaneous.DateCorrelationOptimization -eq $true) {

					$Details = "Database has date correlation enabled.  This is not a default setting, and it has some performance overhead.  It tells SQL Server that date fields in two tables are related, and SQL Server maintains statistics showing that relation."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $LowPriority `
					-Category $CatPerformance `
					-Description 'Date Correlation On' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb522682', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Transparent Database Encryption Enabled
				#region
				if ($_.Properties.Options.OtherOptions.State.EncryptionEnabled -eq $true) {

					$Details = "Database has Transparent Data Encryption enabled.  Make absolutely sure you have backed up the certificate and private key, or else you will not be able to restore this database."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $LowPriority `
					-Category $CatSecurity `
					-Description 'Database Encrypted' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb934049', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Database files on the Windows system drive
				#region

				# System Database on C Drive
				#region
				if (
					@('master','model','msdb') -icontains $_.Name -and
					(
						$_.Properties.Files.DatabaseFiles | Where-Object { 
							$_.ID -and
							[System.IO.Path]::GetPathRoot($_.Path) -ieq $SystemDrive
						} | Measure-Object
					).Count -gt 0
				) {

					$Details = "System database has a file on the C: drive. Putting system databases on the C: drive runs the risk of crashing the server when it runs out of space."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $MediumPriority `
					-Category $CatReliability `
					-Description 'System Database on C: Drive' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms178028', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# TempDB on C Drive
				#region
				if (
					@('tempdb') -icontains $_.Name -and
					(
						$_.Properties.Files.DatabaseFiles | Where-Object { 
							$_.ID -and
							[System.IO.Path]::GetPathRoot($_.Path) -ieq $SystemDrive
						} | Measure-Object
					).Count -gt 0
				) {

					$Details = if (
						$( $_.Properties.Files.DatabaseFiles | Where-Object { $_.Growth -gt 0 } | Measure-Object).Count -gt 0
					) {
						"The tempdb database has files on the C: drive. TempDB frequently grows unpredictably, putting your server at risk of running out of C: drive space and crashing hard. C: is also often much slower than other drives, so performance may be suffering."
					} else {
						"The tempdb database has files on the C: drive. TempDB is not set to Autogrow, hopefully it is big enough. C: is also often much slower than other drives, so performance may be suffering."
					}

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $MediumPriority `
					-Category $CatPerformance `
					-Description 'TempDB on C: Drive' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms178028', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# User Databases on C Drive
				#region
				if (
					@('master','model','msdb','tempdb') -notcontains $_.Name -and
					(
						$_.Properties.Files.DatabaseFiles | Where-Object { 
							$_.ID -and
							[System.IO.Path]::GetPathRoot($_.Path) -ieq $SystemDrive
						} | Measure-Object
					).Count -gt 0
				) {

					$Details = "Database has a file on the C: drive. Putting databases on the C: drive runs the risk of crashing the server when it runs out of space."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $MediumPriority `
					-Category $CatReliability `
					-Description 'User Databases on C: Drive' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms178028', $HelpUrlModifier, '.aspx')))
				} 
				#endregion

				#endregion

				# Tempdb checks
				#region
				if ($_.name -ieq 'tempdb') {

					$FileCount = ($_.Properties.Files.DatabaseFiles | Where-Object { $_.FileType -ieq 'rows data' } | Measure-Object).Count

					# Tempdb file counts not optimal
					if (
						(
							$CpuCount -le 8 -and 
							$CpuCount -ne $FileCount
						) -or
						(
							$CpuCount -gt 8 -and 
							$FileCount -lt 8
						)
					) {

						$Details = "TempDB is configured with $($FileCount) data file(s). SQL Server may experience SGAM contention when tempdb is heavily used; more data files may alleviate the contention."

						Get-AssessmentFinding -ServerName $ServerName `
						-DatabaseName $DatabaseName `
						-Priority $MediumPriority `
						-Category $CatPerformance `
						-Description 'TempDB Data File Count Not Optimal' `
						-Details $Details `
						-URL 'http://support.microsoft.com/kb/328551'
						#-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb522469', $HelpUrlModifier, '.aspx')))
					}
				}
				#endregion

				# Multiple Log Files on One Drive
				#region
				if (@('tempdb') -notcontains $_.Name) {

					$_.Properties.Files.DatabaseFiles | Where-Object { 
						$_.ID -and 
						$_.FileType -ieq 'log' 
					} | Group-Object -Property { [System.IO.Path]::GetPathRoot($_.Path)} | Where-Object { 
						$_.Count -gt 1
					} | ForEach-Object {

						$Details = "Database has multiple log files on the $($_.Name) drive. This is not a performance booster because log file access is sequential, not parallel."

						Get-AssessmentFinding -ServerName $ServerName `
						-DatabaseName $DatabaseName `
						-Priority $MediumPriority `
						-Category $CatPerformance `
						-Description 'Multiple Log Files on One Drive' `
						-Details $Details `
						-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb522469', $HelpUrlModifier, '.aspx')))
					}
				}
				#endregion

				# Uneven File Growth Settings in One Filegroup
				#region
				$_.Properties.Files.DatabaseFiles | Where-Object {
					$_.ID -and
					$_.FileType -ieq 'rows data'
				} | Group-Object -Property Filegroup | Where-Object {
					$_.Count -gt 1 -and
					(
						$($_.Group | Group-Object -Property Growth | Measure-Object).Count -gt 1 -or
						$($_.Group | Group-Object -Property GrowthType | Measure-Object).Count -gt 1
					)
				} | Select-Object -First 1 | ForEach-Object {

					$Details = "Database has multiple data files in one filegroup, but they are not all set up to grow in identical amounts.  This can lead to uneven file activity inside the filegroup."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $MediumPriority `
					-Category $CatPerformance `
					-Description 'Uneven File Growth Settings in One Filegroup' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb522469', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Database Owner <> SA
				#region
				if ($_.Properties.General.Database.Owner -ine $SaLogin) {

					$Details = "Database is owned by [$($_.Properties.General.Database.Owner)]. Several non-obvious things may not work as expected when the owner is not SA. Consider changing the owner to the SA account"

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $LowPriority `
					-Category $CatSecurity `
					-Description 'Database Owner <> SA' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms187359', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Database Collation Mismatch
				#region
				if (
					@('ReportServer','ReportServerTempDB') -notcontains $_.Name -and
					$ServerCollation -ine $_.Properties.General.Maintenance.Collation
				) {

					$Details = "Database collation $($_.Properties.General.Maintenance.Collation) does not match Server collation $ServerCollation"

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $LowPriority `
					-Category $CatReliability `
					-Description 'Collation Mismatch' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms174269', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# File growth set to percent
				#region
				if (
					(
						$_.Properties.Files.DatabaseFiles | Where-Object { 
							$_.ID -and
							$_.GrowthType -ieq 'percent'
						} | Measure-Object
					).Count -gt 0
				) {

					$Details = "Database is using percent filegrowth settings. This can lead to out of control filegrowth."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $MediumPriority `
					-Category $CatPerformance `
					-Description 'File growth set to percent' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb522469', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Old Compatibility Level
				#region
				if (
					@('model') -notcontains $_.Name -and
					$ModelDbCompatLevel -ne $_.Properties.Options.CompatibilityLevel
				) {

					$Details = "Database is configured for compatibility level ""$($_.Properties.Options.CompatibilityLevel)"", which may cause unwanted results when trying to run queries that have newer T-SQL features."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $MediumPriority `
					-Category $CatPerformance `
					-Description 'Old Compatibility Level' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb510680', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Last good DBCC CHECKDB over 2 weeks old
				#region
				if (
					$_.Properties.General.Database.Status -ieq 'normal' -and
					$_.Properties.General.Database.IsDatabaseSnapshot -ne $true -and
					$_.Properties.General.Database.LastKnownGoodDbccDate -ne $null -and
					$_.Properties.General.Database.LastKnownGoodDbccDate.CompareTo($($ScanDateLocal).AddDays(-14)) -le 0
				) {

					$Details = "Database last had a successful DBCC CHECKDB run on $($_.Properties.General.Database.LastKnownGoodDbccDate.ToString('G')) which is more than two weeks prior to the inventory scan on $($ScanDateLocal.ToString('G')). This check should be run regularly to catch any database corruption as soon as possible. Note: you can restore a backup of a busy production database to a test server and run DBCC CHECKDB against that to minimize impact. If you do that, you can ignore this warning."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $MediumPriority `
					-Category $CatReliability `
					-Description 'Last good DBCC CHECKDB over 2 weeks old' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms176064', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# DBCC CHECKDB never run sucessfully
				#region
				if (
					$_.Properties.General.Database.Status -ieq 'normal' -and
					$_.Properties.General.Database.IsDatabaseSnapshot -ne $true -and
					$_.Properties.General.Database.LastKnownGoodDbccDate -eq $null
				) {

					$Details = "Database has never had a successful DBCC CHECKDB. This check should be run regularly to catch any database corruption as soon as possible. Note: you can restore a backup of a busy production database to a test server and run DBCC CHECKDB against that to minimize impact. If you do that, you can ignore this warning."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $HighPriority `
					-Category $CatReliability `
					-Description 'Last good DBCC CHECKDB over 2 weeks old' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms176064', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# High VLF Count
				#region
				$_.Properties.Files.DatabaseFiles | Where-Object {
					$_.ID -and
					$_.FileType -ieq 'log'
				} | Measure-Object -Property VlfCount -Sum | Where-Object {
					$_.Sum -gt 100
				} | ForEach-Object {

					$Details = "Database has $($_.Sum) virtual log files (VLFs). This may be slowing down startup, restores, and even inserts/updates/deletes."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $MediumPriority `
					-Category $CatPerformance `
					-Description 'High VLF Count' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms190925', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Transaction Log Larger than Data File
				#region
				if ($_.Properties.General.Database.IsDatabaseSnapshot -ne $true) {

					$DatabaseFiles = $_.Properties.Files.DatabaseFiles

					$DatabaseFiles | Where-Object { 
						$_.ID -and
						$_.FileType -ieq 'log' -and
						$_.SizeKB/1MB -gt 1
					} | ForEach-Object {
						$SizeKB = $_.SizeKB

						$DatabaseFiles | Where-Object { 
							$_.ID -and
							$_.FileType -ine 'log' -and
							$SizeKB -gt $_.SizeKB
						}
					} | Select-Object -First 1 | ForEach-Object {

						$Details = "Database has a transaction log file larger than a data file. This may indicate that transaction log backups are not being performed or not performed often enough."

						Get-AssessmentFinding -ServerName $ServerName `
						-DatabaseName $DatabaseName `
						-Priority $MediumPriority `
						-Category $CatReliability `
						-Description 'Transaction Log Larger than Data File' `
						-Details $Details `
						-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms190925', $HelpUrlModifier, '.aspx')))
					}
				}
				#endregion

				# Database collation different than tempdb collation
				#region
				if (
					@('master','model','msdb','tempdb','ReportServer','ReportServerTempDB') -notcontains $_.Name -and
					$TempDbCollation -ine $_.Properties.General.Maintenance.Collation
				) {

					$Details = "Database collation $($_.Properties.General.Maintenance.Collation) is different than TempDB collation $TempDbCollation. Collation differences between user databases and tempdb can cause conflicts especially when comparing string values"

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $MediumPriority `
					-Category $CatReliability `
					-Description 'Collation Mismatch' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms174269', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Database Snapshot Online
				#region
				if ($_.Properties.General.Database.IsDatabaseSnapshot -eq $true) {

					$Details = "Database is a snapshot of [$($_.Properties.General.Database.DatabaseSnapshotBaseName)]. Make sure you have enough drive space to maintain the snapshot as the original database grows."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $MediumPriority `
					-Category $CatReliability `
					-Description 'Database Snapshot Online' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms175158', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Max File Size Set
				#region
				$_.Properties.Files.DatabaseFiles | Where-Object {
					$_.ID -and
					$_.MaxSizeKB -ne 268435456 -and
					$_.MaxSizeKB -ne 2147483648 -and
					$_.MaxSizeKB -ne -1
				} | ForEach-Object {

					$Details = "Database file [$($_.LogicalName)] has a max file size set to $($_.MaxSizeKB / 1KB) MB. If it runs out of space, the database will stop working even though there may be drive space available."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $MediumPriority `
					-Category $CatReliability `
					-Description 'Max File Size Set' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb522469', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Plan Guides Enabled
				#region
				if ($_.Programmability.PlanGuides | Where-Object { $_.ID }) {

					$Details = "Query plan guides are in use. Plan guides associate hints and actual plans with queries without having to modify the query code itself. If your query performance won't improve no matter what you try it may be because of a frozen plan. Check sys.plan_guides to review the plan guides that are in place in this database."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $LowPriority `
					-Category $CatPerformance `
					-Description 'Plan Guides Enabled' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms187032', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Tables in the master database
				#region
				if (
					$_.Name -ieq 'master' -and
					$_.Tables | Where-Object { 
						$_.ID -and
						$_.Properties.General.Description.IsSystemObject -ne $true
					}
				) {

					$Details = "The Master database contains user created tables. Tables in this database may not be restored in the event of a disaster."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $LowPriority `
					-Category $CatRecovery `
					-Description 'User Tables In The Master Database' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms187837', $HelpUrlModifier, '.aspx')))
				} 
				#endregion

				# Tables in the msdb database
				#region
				if (
					$_.Name -ieq 'master' -and
					$_.Tables | Where-Object { 
						$_.ID -and
						$_.Properties.General.Description.IsSystemObject -ne $true
					}
				) {

					$Details = "The MSDB database contains user created tables. Tables in this database may not be restored in the event of a disaster."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $LowPriority `
					-Category $CatRecovery `
					-Description 'User Tables In The MSDB Database' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms187112', $HelpUrlModifier, '.aspx')))
				} 
				#endregion

				# Tables in the model database
				#region
				if (
					$_.Name -ieq 'model' -and
					$_.Tables | Where-Object { 
						$_.ID -and
						$_.Properties.General.Description.IsSystemObject -ne $true
					}
				) {

					$Details = "The Model database contains user created tables. Any new database created on this instance will contain a copy of these tables."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $LowPriority `
					-Category $CatReliability `
					-Description 'User Tables In The Model Database' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms186388', $HelpUrlModifier, '.aspx')))
				} 
				#endregion

				# Offline files
				#region
				if (
					$_.Properties.Files.DatabaseFiles | Where-Object { 
						$_.IsOffline -eq $true
					}
				) {

					$Details = "Database has file(s) that are offline"

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $HighPriority `
					-Category $CatAvailability `
					-Description 'Offline Files' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb522469', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Guest Permissions on User Databases
				#region
				if (
					$_.Properties.General.Database.IsSystemObject -ne $true -and
					(
						$_.Properties.Permissions | Where-Object {
							@('Grant','Grant With Grant') -icontains $_.PermissionState -and
							$_.PermissionType -ieq 'connect' -and
							$_.Grantee -ieq 'guest' -and
							$_.GranteeType -ieq 'user'
						}
					)
				) {

					$Details = "The guest user has permission to access the database. The guest user can't be dropped, but you should revoke the user's permission to access the database if it's not required."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $MediumPriority `
					-Category $CatSecurity `
					-Description 'Guest Permissions on User Database' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/bb402861', $HelpUrlModifier, '.aspx')))

				}
				#endregion

				# Triggers on Tables
				#region
				if ($_.Tables | Where-Object { $_.Triggers | Where-Object { $_.ID } }) { 

					$Details = "There are tables with triggers in the database. Triggers have a reputation for being tricky to track and troubleshoot; be sure to review them to ensure they aren't having a negative performance impact."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $MediumPriority `
					-Category $CatPerformance `
					-Description 'Triggers On Tables' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms178110', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Granular Table checks
				$_.Tables | Where-Object { $_.ID } | ForEach-Object {

					$SchemaName = $_.Properties.General.Description.Schema
					$ObjectName = $_.Properties.General.Description.Name

					# Disabled Indexes
					#region
					$_.Indexes | Where-Object { $_.General.IsDisabled } | ForEach-Object {

						$Details = "Index $($_.General.Name) on table [$($SchemaName)].[$($ObjectName)] is disabled. Enable the index if it is needed or consider removing it if it's not."

						Get-AssessmentFinding -ServerName $ServerName `
						-DatabaseName $DatabaseName `
						-Priority $MediumPriority `
						-Category $CatPerformance `
						-Description 'Disabled Indexes' `
						-Details $Details `
						-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms177406', $HelpUrlModifier, '.aspx')))

					}
					#endregion


					# Leftover hypothetical indexes
					#region
					$_.Indexes | Where-Object { $_.General.IsHypothetical } | ForEach-Object {

						$Details = "Index $($_.General.Name) on table [$($SchemaName)].[$($ObjectName)] is hypothetical. This is often an artifact from running the Index Tuning Wizard or Database Tuning Advisor. The index does nothing to help performance and is safe to remove."

						Get-AssessmentFinding -ServerName $ServerName `
						-DatabaseName $DatabaseName `
						-Priority $LowPriority `
						-Category $CatPerformance `
						-Description 'Leftover Hypothetical Indexes' `
						-Details $Details `
						-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms190172', $HelpUrlModifier, '.aspx')))

					}
					#endregion


					# Table Foreign Keys Not Trusted
					#region
					$_.ForeignKeys | Where-Object {
						$_.IsChecked -ne $true -and
						$_.IsNotForReplication -ne $true -and
						$_.IsEnabled -eq $true
					} | ForEach-Object {

						$Details = "Foreign Key $($_.Name) on table [$($SchemaName)].[$($ObjectName)] is enabled but not checked. Alter the table with CHECK CONSTRAINT to enable and check the constraint. When enabled AND checked the query optimizer can produce more efficient query plans."

						Get-AssessmentFinding -ServerName $ServerName `
						-DatabaseName $DatabaseName `
						-Priority $MediumPriority `
						-Category $CatPerformance `
						-Description 'Foreign Keys Not Trusted' `
						-Details $Details `
						-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms175464', $HelpUrlModifier, '.aspx')))
					}
					#endregion

					# Table Check Constraints Not Trusted
					#region
					$_.Checks | Where-Object {
						$_.IsChecked -ne $true -and
						$_.IsNotForReplication -ne $true -and
						$_.IsEnabled -eq $true
					} | ForEach-Object {

						$Details = "Check Constraint $($_.Name) on table [$($SchemaName)].[$($ObjectName)] is enabled but not checked. Alter the table with CHECK CONSTRAINT to enable and check the constraint. When enabled AND checked the query optimizer can produce more efficient query plans."

						Get-AssessmentFinding -ServerName $ServerName `
						-DatabaseName $DatabaseName `
						-Priority $MediumPriority `
						-Category $CatPerformance `
						-Description 'Check Constraints Not Trusted' `
						-Details $Details `
						-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms188258', $HelpUrlModifier, '.aspx')))
					}
					#endregion

				}

				# Stored Procedure WITH RECOMPILE
				#region
				$_.Programmability.StoredProcedures | Where-Object {
					$_.Definition -ilike '*WITH RECOMPILE*'
				} | ForEach-Object {

					$Details = "Stored Procedure [$($_.Properties.General.Description.Schema)].[$($_.Properties.General.Description.Name)] has WITH RECOMPILE in the code. This will cause the procedure to be recompiled on each execution which may lead to increased CPU usage if it is called frequently."

					Get-AssessmentFinding -ServerName $ServerName `
					-DatabaseName $DatabaseName `
					-Priority $LowPriority `
					-Category $CatPerformance `
					-Description 'Stored Procedure WITH RECOMPILE' `
					-Details $Details `
					-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms187926', $HelpUrlModifier, '.aspx')))
				}
				#endregion

				# Elevated Database Permissions
				#region
				$_.Security.DatabaseRole | Where-Object {
					@('db_owner','db_accessAdmin','db_securityadmin','db_ddladmin') -icontains $_.Name -and
					-not [String]::IsNullOrEmpty($_.Member)
				} | ForEach-Object {

					$RoleName = $_.Name

					$_.Member | Where-Object {
						$_ -ine 'dbo' 
					} | ForEach-Object {

						$Details = "User [$_] is a member of the $RoleName fixed database role. This role grants the user rights to do things besides reading and writing data - whether they mean to or not!"

						Get-AssessmentFinding -ServerName $ServerName `
						-DatabaseName $DatabaseName `
						-Priority $MediumPriority `
						-Category $CatSecurity `
						-Description 'Elevated Database Permissions' `
						-Details $Details `
						-URL $([String]::Concat(@('http://msdn.microsoft.com/en-us/library/ms189121', $HelpUrlModifier, '.aspx'))) 
					}
				}
				#endregion

			}
			#endregion

		}
		#endregion


		#region
		# # Page Verification Not Optimal (2000)
		# $SqlServerInventory.DatabaseServer | Where-Object {
		#     ($_.Server.Configuration.General.Version -ilike '8.*')
		# } | ForEach-Object {
		# 	$ServerName = $_.Server.Configuration.General.Name
		#     $_.Server.Databases | Where-Object {
		#         ($_.Properties.Options.OtherOptions.Recovery.PageVerify -ine 'CHECKSUM')
		#     } | Sort-Object -Property $_.Name | ForEach-Object {
		#         Write-Output $_.Name
		#     }
		# }
		# 
		# # Page Verification Not Optimal (2005+)
		# $SqlServerInventory.DatabaseServer | Where-Object {
		#     ($_.Server.Configuration.General.Version -inotlike '8.*')
		# } | Sort-Object -Property $_.ServerName | ForEach-Object {
		#     $_.Server.Databases | Where-Object {
		#         ($_.Name -ine 'tempdb') -and
		#         #($_.Properties.Options.OtherOptions.Recovery.PageVerify -ieq 'NONE') -or
		#         #($_.Properties.Options.OtherOptions.Recovery.PageVerify -ieq 'TORN_PAGE_DETECTION')
		#         ($_.Properties.Options.OtherOptions.Recovery.PageVerify -ine 'CHECKSUM')
		#     } | Sort-Object -Property $_.Name | ForEach-Object {
		#         Write-Output $_.Name
		#     }
		# }
		#endregion


		## Linked Servers Configured for Replication
		#region
		# $SqlServerInventory.DatabaseServer | Where-Object {
		#     ($_.Options.Distributor -eq $true) -or
		#     ($_.Options.DistPublisher -eq $true) -or
		#     ($_.Options.Publisher -eq $true) -or
		#     ($_.Options.Subscriber -eq $true)
		# } | Sort-Object -Property $_.ServerName | ForEach-Object {
		#     Write-Output $_.Name
		# }
		#endregion

	}
	end {
		Remove-Variable -Name NullDatabaseName
	}
}

function Export-SqlServerInventoryToExcel {
	<#
	.SYNOPSIS
		Writes an Excel file containing the information in a SQL Server Inventory.

	.DESCRIPTION
		The Export-SqlServerInventoryToExcel function uses COM Interop to write Excel files containing the information in a SQL Server Inventory returned by Get-SqlServerInventory.
		
		Although the SQL Server Shared Management Objects (SMO) libraries are required to perform an inventory they are NOT required to write the Excel files.
		
		Microsoft Excel 2007 or higher must be installed in order to write the Excel files.
		
	.PARAMETER  SqlServerInventory
		A SQL Server Inventory object returned by Get-SqlServerInventory.
		
	.PARAMETER  DirectoryPath
		Specifies the directory path where the Excel files will be written.
		
		If not specified then the default is your "My Documents" folder.
		
	.PARAMETER  BaseFilename
		Specifies the base name to be used for each Excel file that is written.
		
		If not specified then the default is "SQL Server Inventory".

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
		Export-SqlServerInventoryToExcel -SqlServerInventory $Inventory 
		
		Description
		-----------
		Write a SQL Server inventory using the SQL Server Inventory contained in $Inventory.

		The Excel workbooks will be written to your "My Documents" folder.
		
		The Office color theme and Medium color scheme will be used by default.
		
	.EXAMPLE
		Export-SqlServerInventoryToExcel -SqlServerInventory $Inventory -DirectoryPath 'C:\DB Engine Inventory.xlsx'
		
		Description
		-----------
		Write a SQL Server inventory using the SQL Server Inventory contained in $Inventory.

		The Excel workbooks will be written to the "C:\DB Engine Inventory\" directory.
		
		The Office color theme and Medium color scheme will be used by default.

	.EXAMPLE
		Export-SqlServerInventoryToExcel -SqlServerInventory $Inventory -ColorTheme Blue -ColorScheme Dark
		
		Description
		-----------
		Write a SQL Server inventory using the SQL Server Inventory contained in $Inventory.

		The Excel workbooks will be written to your "My Documents" folder.
		
		The Blue color theme and Dark color scheme will be used.
	
	.NOTES
		Blue and Green are nice looking Color Themes for Office 2013

		Waveform is a nice looking Color Theme for Office 2010

	.LINK
		Get-SqlServerInventory
		Get-SqlServerInventoryDatabaseEngineAssessment
		Export-SqlServerInventoryWindowsInventoryToExcel
		Export-SqlServerInventoryDatabaseEngineConfigToExcel
		Export-SqlServerInventoryDatabaseEngineAssessmentToExcel
#>
	[cmdletBinding()]
	param(
		[Parameter(Mandatory=$true)]
		[PSCustomObject]
		$SqlServerInventory
		,
		[Parameter(Mandatory=$true)] 
		[ValidateNotNullOrEmpty()]
		[string]
		$DirectoryPath
		,
		[Parameter(Mandatory=$false)] 
		[AllowNull()]
		[string]
		$BaseFilename = 'SQL Server Inventory'
		,
		[Parameter(Mandatory=$false)] 
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
		$ProgressId = Get-Random
		$ProgressActivity = 'Writing SQL Server Inventory To Excel'
		$TaskCount = 4
		$TaskNumber = 0
		$DbEngineConfigPath = Join-Path -Path $DirectoryPath -ChildPath ([System.IO.Path]::GetFileNameWithoutExtension($BaseFilename) + ' - Database Engine Config.xlsx')
		$DbEngineDbObjectsPath = Join-Path -Path $DirectoryPath -ChildPath ([System.IO.Path]::GetFileNameWithoutExtension($BaseFilename) + ' - Database Engine Db Objects.xlsx')
		$DbEngineAssessmentPath = Join-Path -Path $DirectoryPath -ChildPath ([System.IO.Path]::GetFileNameWithoutExtension($BaseFilename) + ' - Database Engine Assessment.xlsx')
		$WindowsInventoryPath = Join-Path -Path $DirectoryPath -ChildPath ([System.IO.Path]::GetFileNameWithoutExtension($BaseFilename) + ' - Windows.xlsx')
	}
	process {
		if ($SqlServerInventory.DatabaseServerScanSuccessCount -gt 0) {

			# Database Engine Config
			$ProgressStatus = 'Writing Database Engine Config To Excel'
			$TaskNumber++
			Write-Progress -Activity $ProgressActivity -PercentComplete $((($TaskNumber - 0) / $TaskCount) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			Export-SqlServerInventoryDatabaseEngineConfigToExcel -SqlServerInventory $SqlServerInventory -Path $DbEngineConfigPath -ColorTheme $ColorTheme -ColorScheme $ColorScheme -ParentProgressId $ProgressId

			# Database Engine Database Objects
			if ($SqlServerInventory.DatabaseServer | Where-Object { $_.HasDatabaseObjectInformation -eq $true }) {
				$ProgressStatus = 'Writing Database Engine Database Objects To Excel'
				$TaskNumber++
				Write-Progress -Activity $ProgressActivity -PercentComplete $((($TaskNumber - 0) / $TaskCount) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
				Export-SqlServerInventoryDatabaseEngineDbObjectsToExcel -SqlServerInventory $SqlServerInventory -Path $DbEngineDbObjectsPath -ColorTheme $ColorTheme -ColorScheme $ColorScheme -ParentProgressId $ProgressId
			}

			# Database Engine Assessment
			$DbEngineAssessment = Get-SqlServerInventoryDatabaseEngineAssessment -SqlServerInventory $SqlServerInventory
			$ProgressStatus = 'Writing Database Engine Assessment To Excel'
			$TaskNumber++
			Write-Progress -Activity $ProgressActivity -PercentComplete $((($TaskNumber - 0) / $TaskCount) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			Export-SqlServerInventoryDatabaseEngineAssessmentToExcel -DatabaseEngineAssessment $DbEngineAssessment -Path $DbEngineAssessmentPath -ColorTheme $ColorTheme -ColorScheme $ColorScheme -ParentProgressId $ProgressId

			# Windows Inventory
			$ProgressStatus = 'Writing Windows Inventory To Excel'
			$TaskNumber++
			Write-Progress -Activity $ProgressActivity -PercentComplete $((($TaskNumber - 0) / $TaskCount) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			Export-SqlServerInventoryWindowsInventoryToExcel -SqlServerInventory $SqlServerInventory -Path $WindowsInventoryPath -ColorTheme $ColorTheme -ColorScheme $ColorScheme -ParentProgressId $ProgressId

		} else {
			Write-LogMessage -Message 'No SQL Server instances contained in the inventory!' -MessageLevel Warning
		} 
	}
	end {
		Remove-Variable -Name ProgressId, ProgressActivity, TaskCount, TaskNumber, DbEngineConfigPath, DbEngineDbObjectsPath, DbEngineAssessmentPath, WindowsInventoryPath
	}
}

function Export-SqlServerInventoryWindowsInventoryToExcel {
	<#
	.SYNOPSIS
		Writes an Excel file containing the Windows Inventory information in a SQL Server Inventory.

	.DESCRIPTION
		The Export-SqlServerInventoryWindowsInventoryToExcel function uses COM Interop to write an Excel file containing the Windows Inventory information returned by Get-SqlServerInventory.
		
		This function is a wrapper to Export-SqlServerInventoryToExcel in the WindowsInventory module.
		
		Microsoft Excel 2007 or higher must be installed in order to write the Excel file.
		
	.PARAMETER  SqlServerInventory
		A SQL Server Inventory object returned by Get-SqlServerInventory.
		
	.PARAMETER  Path
		Specifies the path where the Excel file will be written. This is a fully qualified path to a .XLSX file.
		
		If not specified then the file is named "SQL Server Inventory - [Year][Month][Day][Hour][Minute] - Windows.xlsx" and is written to your "My Documents" folder.

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
		Export-SqlServerInventoryWindowsInventoryToExcel -SqlServerInventory $Inventory 
		
		Description
		-----------
		Write a Windows inventory using $Inventory.

		The Excel workbook will be written to your "My Documents" folder.
		
		The Office color theme and Medium color scheme will be used by default.
		
	.EXAMPLE
		Export-SqlServerInventoryWindowsInventoryToExcel -SqlServerInventory $Inventory -Path 'C:\Windows Inventory.xlsx'
		
		Description
		-----------
		Write a Windows inventory using $Inventory.

		The Excel workbook will be written to your C:\Windows Inventory.xlsx.
		
		The Office color theme and Medium color scheme will be used by default.

	.EXAMPLE
		Export-SqlServerInventoryWindowsInventoryToExcel -SqlServerInventory $Inventory -ColorTheme Blue -ColorScheme Dark
		
		Description
		-----------
		Write a Windows inventory using $Inventory.

		The Excel workbook will be written to your "My Documents" folder.
		
		The Blue color theme and Dark color scheme will be used.
	
	.NOTES
		Blue and Green are nice looking Color Themes for Office 2013

		Waveform is a nice looking Color Theme for Office 2010

	.LINK
		Get-SqlServerInventory

#>
	[cmdletBinding()]
	param(
		[Parameter(Mandatory=$true, ValueFromPipeline=$true)]
		[PSCustomObject]
		$SqlServerInventory
		,
		[Parameter(Mandatory=$false)] 
		[ValidateNotNullOrEmpty()]
		[string]
		$Path = (Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)) -ChildPath ('SQL Server Inventory - ' + (Get-Date -Format 'yyyy-MM-dd-HH-mm') + ' - Windows'), 'xlsx')
		,
		[Parameter(Mandatory=$false)] 
		[alias('theme')]
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
	process {
		Export-WindowsInventoryToExcel -WindowsInventory $SqlServerInventory.WindowsInventory -Path $Path -ColorTheme $ColorTheme -ColorScheme $ColorScheme -ParentProgressId $ParentProgressId
	}
}

function Export-SqlServerInventoryDatabaseEngineConfigToExcel {
	<#
	.SYNOPSIS
		Writes an Excel file containing the Database Engine information in a SQL Server Inventory.

	.DESCRIPTION
		The Export-SqlServerInventoryDatabaseEngineConfigToExcel function uses COM Interop to write an Excel file containing the Database Engine information in a SQL Server Inventory returned by Get-SqlServerInventory.
		
		Although the SQL Server Shared Management Objects (SMO) libraries are required to perform an inventory they are NOT required to write the Excel file.
		
		Microsoft Excel 2007 or higher must be installed in order to write the Excel file.
		
	.PARAMETER  SqlServerInventory
		A SQL Server Inventory object returned by Get-SqlServerInventory.
		
	.PARAMETER  Path
		Specifies the path where the Excel file will be written. This is a fully qualified path to a .XLSX file.
		
		If not specified then the file is named "SQL Server Inventory - [Year][Month][Day][Hour][Minute] - Database Engine.xlsx" and is written to your "My Documents" folder.

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
		Export-SqlServerInventoryDatabaseEngineConfigToExcel -SqlServerInventory $Inventory 
		
		Description
		-----------
		Write a Database Engine inventory using the SQL Server Inventory contained in $Inventory.

		The Excel workbook will be written to your "My Documents" folder.
		
		The Office color theme and Medium color scheme will be used by default.
		
	.EXAMPLE
		Export-SqlServerInventoryDatabaseEngineConfigToExcel -SqlServerInventory $Inventory -Path 'C:\DB Engine Inventory.xlsx'
		
		Description
		-----------
		Write a Database Engine inventory using the SQL Server Inventory contained in $Inventory.

		The Excel workbook will be written to your C:\DB Engine Inventory.xlsx.
		
		The Office color theme and Medium color scheme will be used by default.

	.EXAMPLE
		Export-SqlServerInventoryDatabaseEngineConfigToExcel -SqlServerInventory $Inventory -ColorTheme Blue -ColorScheme Dark
		
		Description
		-----------
		Write a Database Engine inventory using the SQL Server Inventory contained in $Inventory.

		The Excel workbook will be written to your "My Documents" folder.
		
		The Blue color theme and Dark color scheme will be used.
	
	.NOTES
		Blue and Green are nice looking Color Themes for Office 2013

		Waveform is a nice looking Color Theme for Office 2010

	.LINK
		Get-SqlServerInventory

#>
	[cmdletBinding()]
	param(
		[Parameter(Mandatory=$true, ValueFromPipeline=$true)]
		[PSCustomObject]
		$SqlServerInventory
		,
		[Parameter(Mandatory=$false)] 
		[ValidateNotNullOrEmpty()]
		[string]
		$Path = [System.IO.Path]::ChangeExtension((Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)) -ChildPath ('SQL Server Inventory - ' + (Get-Date -Format 'yyyy-MM-dd-HH-mm') + ' - Database Engine Config')), 'xlsx')
		,
		[Parameter(Mandatory=$false)] 
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


		<#
	.LINK
		"Export Windows PowerShell Data to Excel" (http://technet.microsoft.com/en-us/query/dd297620)
		
	.LINK
		"Microsoft.Office.Interop.Excel Namespace" (http://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.excel(v=office.14).aspx)
		
	.LINK
		"Excel Object Model Reference" (http://msdn.microsoft.com/en-us/library/ff846392.aspx)
		
	.LINK
		"Excel 2010 Enumerations" (http://msdn.microsoft.com/en-us/library/ff838815.aspx)
		
	.LINK
		"TableStyle Cheat Sheet in French/English" (http://msdn.microsoft.com/fr-fr/library/documentformat.openxml.spreadsheet.tablestyle.aspx)
		
	.LINK
		"Color Palette and the 56 Excel ColorIndex Colors" (http://dmcritchie.mvps.org/excel/colors.htm)
		
	.LINK
		"Adding Color to Excel 2007 Worksheets by Using the ColorIndex Property" (http://msdn.microsoft.com/en-us/library/cc296089(v=office.12).aspx)
		
	.LINK
		"XlRgbColor Enumeration" (http://msdn.microsoft.com/en-us/library/ff197459.aspx)
#>



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

		$TabCharLength = 4
		$IndentString_1 = [String]::Empty.PadLeft($TabCharLength * 1)
		$IndentString__2 = [String]::Empty.PadLeft($TabCharLength * 2)
		$IndentString___3 = [String]::Empty.PadLeft($TabCharLength * 3)

		$ComputerName = $null
		$ServerName = $null
		$ProductName = $null
		$DatabaseName = $null
		$FileGroupName = $null

		$ColorThemePathPattern = $null
		$ColorThemePath = $null

		# Used to hold all of the formatting to be applied at the end
		$WorksheetCount = 55
		$WorksheetFormat = @{}

		$XlSortOrder = 'Microsoft.Office.Interop.Excel.XlSortOrder' -as [Type]
		$XlYesNoGuess = 'Microsoft.Office.Interop.Excel.XlYesNoGuess' -as [Type]
		$XlHAlign = 'Microsoft.Office.Interop.Excel.XlHAlign' -as [Type]
		$XlVAlign = 'Microsoft.Office.Interop.Excel.XlVAlign' -as [Type]
		$XlListObjectSourceType = 'Microsoft.Office.Interop.Excel.XlListObjectSourceType' -as [Type]
		$XlThemeColor = 'Microsoft.Office.Interop.Excel.XlThemeColor' -as [Type]

		$OverviewTabColor = $XlThemeColor::xlThemeColorDark1
		$ServicesTabColor = $XlThemeColor::xlThemeColorLight1
		# 		$ServerTabColor = $XlThemeColor::xlThemeColorAccent2
		# 		$DatabaseTabColor = $XlThemeColor::xlThemeColorAccent4
		# 		$SecurityTabColor = $XlThemeColor::xlThemeColorAccent3
		# 		$ServerObjectsTabColor = $XlThemeColor::xlThemeColorAccent5
		# 		$ManagementTabColor = $XlThemeColor::xlThemeColorAccent2
		# 		$AgentTabColor = $XlThemeColor::xlThemeColorAccent6

		$ServerTabColor = $XlThemeColor::xlThemeColorAccent1
		$DatabaseTabColor = $XlThemeColor::xlThemeColorAccent2
		$SecurityTabColor = $XlThemeColor::xlThemeColorAccent3
		$ServerObjectsTabColor = $XlThemeColor::xlThemeColorAccent4
		$ManagementTabColor = $XlThemeColor::xlThemeColorAccent5
		$AgentTabColor = $XlThemeColor::xlThemeColorAccent6

		$TableStyle = switch ($ColorScheme) {
			'light' { 'TableStyleLight8' }
			'medium' { 'TableStyleMedium15' }
			'dark' { 'TableStyleDark1' }
		}

		$ProgressId = Get-Random
		$ProgressActivity = 'Export-SqlServerInventoryDatabaseEngineConfigToExcel'
		$ProgressStatus = 'Beginning output to Excel'

		Write-SqlServerInventoryLog -Message "Start Function: $($MyInvocation.InvocationName)" -MessageLevel Debug
		Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Information
		Write-Progress -Activity $ProgressActivity -PercentComplete 0 -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId


		#region

		# Hide the Excel instance (is this necessary?)
		$Excel.visible = $false

		# Turn off screen updating
		$Excel.ScreenUpdating = $false

		# Turn off automatic calculations
		#$Excel.Calculation = [Microsoft.Office.Interop.Excel.XlCalculation]::xlCalculationManual

		# Add a workbook
		$Workbook = $Excel.Workbooks.Add()
		$Workbook.Title = 'SQL Server Inventory - Database Engine'

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
				Write-SqlServerInventoryLog -Message "Unable to find a theme named ""$ColorTheme"", using default Excel theme instead" -MessageLevel Warning
			}

		}


		# Add enough worksheets to get us to $WorksheetCount
		$Excel.Worksheets.Add($MissingType, $Excel.Worksheets.Item($Excel.Worksheets.Count), $WorksheetCount - $Excel.Worksheets.Count, $Excel.Worksheets.Item(1).Type) | Out-Null
		$WorksheetNumber = 1

		try {

			# Worksheet 1: Services
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Services"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Services'
			#$Worksheet.Tab.Color = $ServicesTabColor
			$Worksheet.Tab.ThemeColor = $ServicesTabColor

			$RowCount = ($SqlServerInventory.Service | Measure-Object).Count + 1
			$ColumnCount = 15
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Computer Name'
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Service Type'
			$WorksheetData[0,$Col++] = 'Service IP Address'
			$WorksheetData[0,$Col++] = 'Service Port'
			$WorksheetData[0,$Col++] = 'Status'
			$WorksheetData[0,$Col++] = 'Process ID'
			$WorksheetData[0,$Col++] = 'Start Date'
			$WorksheetData[0,$Col++] = 'Start Mode'
			$WorksheetData[0,$Col++] = 'Service Account'
			$WorksheetData[0,$Col++] = 'Clustered'
			$WorksheetData[0,$Col++] = 'AlwaysOn'
			$WorksheetData[0,$Col++] = 'Executable Path'
			$WorksheetData[0,$Col++] = 'Startup Parameters'

			$Row = 1
			$SqlServerInventory.Service | ForEach-Object {
				$Col = 0
				$WorksheetData[$Row,$Col++] = $_.ComputerName
				$WorksheetData[$Row,$Col++] = $_.ServerName
				$WorksheetData[$Row,$Col++] = $_.DisplayName #$_.ServiceTypeName
				$WorksheetData[$Row,$Col++] = $_.ServiceIpAddress
				$WorksheetData[$Row,$Col++] = $_.Port
				$WorksheetData[$Row,$Col++] = $_.ServiceState
				$WorksheetData[$Row,$Col++] = $_.ProcessId
				$WorksheetData[$Row,$Col++] = $_.ServiceStartDate
				$WorksheetData[$Row,$Col++] = $_.StartMode
				$WorksheetData[$Row,$Col++] = $_.ServiceAccount
				$WorksheetData[$Row,$Col++] = $_.IsClusteredInstance
				$WorksheetData[$Row,$Col++] = $_.IsHadrEnabled
				$WorksheetData[$Row,$Col++] = $_.PathName
				$WorksheetData[$Row,$Col++] = $_.StartupParameters
				$Row++
			}

			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'D2'
					ColumnFormat = @(
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 2: Server Overview - Servername, Scan Date, Version, Edition, OS, etc.
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Server Overview"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Server Overview'
			#$Worksheet.Tab.Color = $OverviewTabColor
			$Worksheet.Tab.ThemeColor = $OverviewTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | Measure-Object).Count + 1
			$ColumnCount = 19
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Computer Name'
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Scan Date (UTC)'
			$WorksheetData[0,$Col++] = 'Install Date'
			$WorksheetData[0,$Col++] = 'Startup Date'
			$WorksheetData[0,$Col++] = 'Product Name'
			$WorksheetData[0,$Col++] = 'Product Edition'
			$WorksheetData[0,$Col++] = 'Level'
			$WorksheetData[0,$Col++] = 'Platform'
			$WorksheetData[0,$Col++] = 'Version'
			$WorksheetData[0,$Col++] = 'Server Type'
			$WorksheetData[0,$Col++] = 'Clustered'
			$WorksheetData[0,$Col++] = 'Logical Processors'
			$WorksheetData[0,$Col++] = 'Total Memory (MB)'
			$WorksheetData[0,$Col++] = 'Instance Memory In Use (MB)'
			$WorksheetData[0,$Col++] = 'Operating System'
			$WorksheetData[0,$Col++] = 'System Manufacturer'
			$WorksheetData[0,$Col++] = 'System Type'
			$WorksheetData[0,$Col++] = 'Power Plan'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {
				$Col = 0
				$WorksheetData[$Row,$Col++] = $_.Machine.OperatingSystem.Settings.ComputerSystem.FullyQualifiedDomainName # $_.ComputerName
				$WorksheetData[$Row,$Col++] = $_.ServerName
				$WorksheetData[$Row,$Col++] = $_.ScanDateUTC
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.InstallDate
				$WorksheetData[$Row,$Col++] = $_.Server.Service.ServiceStartDate
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Product
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Edition
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.ProductLevel
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Platform
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Version
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.ServerType
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.IsClustered
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.ProcessorCount
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.MemoryMB
				$WorksheetData[$Row,$Col++] = "{0:N2}" -f ($_.Server.Configuration.General.MemoryInUseKB / 1KB)
				$WorksheetData[$Row,$Col++] = $_.Machine.OperatingSystem.Settings.OperatingSystem.Name
				$WorksheetData[$Row,$Col++] = $_.Machine.OperatingSystem.Settings.ComputerSystemProduct.Manufacturer
				$WorksheetData[$Row,$Col++] = $_.Machine.OperatingSystem.Settings.ComputerSystemProduct.Name
				$WorksheetData[$Row,$Col++] = $_.Machine.OperatingSystem.Settings.PowerPlan | Where-Object { $_.IsActive -eq $true } | ForEach-Object { $_.PlanName }
				$Row++
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 3; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 14; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtNumberS2}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# 			# Worksheet 3: Server Configuration (VERTICAL FORMAT)
			# 			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Configuration"
			# 			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			# 			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			# 			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			# 			$Worksheet.Name = 'Server Config'
			# 			#$Worksheet.Tab.Color = $ServerTabColor
			# 			$Worksheet.Tab.ThemeColor = $ServerTabColor
			# 
			# 			#$RowCount = ($SqlServerInventory.DatabaseServer | Measure-Object).Count + 1
			# 			#$ColumnCount = 15
			# 			$ColumnCount = ($SqlServerInventory.DatabaseServer | Measure-Object).Count + 1
			# 			$RowCount = 130
			# 			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount
			# 
			# 			$Row = 0
			# 			$WorksheetData[$Row++,0] = 'Server Name'
			# 			$WorksheetData[$Row++,0] = 'General'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Product'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Edition'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Level'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Version'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Platform'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Operating System'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Language'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Total Memory (MB)'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Instance Memory In Use (MB)'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Processors'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Root Directory'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Server Collation'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Is Clustered'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'AlwaysOn AG Enabled'
			# 			#16
			# 			$WorksheetData[$Row++,0] = 'Memory'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Server Memory Options'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Use AWE'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Minimum Server Memory (MB)'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Maximum Server Memory (MB)'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Other Memory Options'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Index Creation Memory (KB)'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Min Memory Per Query (KB)'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Set Working Set Size'	# a.k.a. "Reserve physical memory for SQL Server" in SQL 2000
			# 			# 25
			# 			$WorksheetData[$Row++,0] = 'Processors'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Enable Processors'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Auto Processor Affinity Mask'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Auto IO Affinity Mask'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Processor Affinity Mask'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Processor Affinity Mask 64'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'IO Affinity Mask'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'IO Affinity Mask 64'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Threads'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Max Worker Threads'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Boost SQL Server Priority'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Use Windows Fibers'
			# 			# 37
			# 			$WorksheetData[$Row++,0] = 'Security'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Authentication Mode'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Login Auditing'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Server Proxy Account'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Proxy Account Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Proxy Account'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Options'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Common Criteria Compliance Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'C2 Audit Tracing Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Cross Database Ownership Chaining'
			# 			# 47
			# 			$WorksheetData[$Row++,0] = 'Connections'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Max Concurrent Connections'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Query Gov. Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Query Gov. Timeout (sec)'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Default Connection Options'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Interim/Deferred Constraint Checking Default'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Implicit Transactions Default'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Cursor Close On Commit Default'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Ansi Warnings Default'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Ansi Padding Default'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Arithmatic Abort Default'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Arithmatic Ignore Default'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Quoted Identifier Default'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'No Count Default'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'ANSI NULL Default On'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'ANSI NULL Default Off'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Concat Null Yields Null Default'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Numeric Round Abort Default'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Xact Abort Default'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Remote Server Connections'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Allow Remote Admin Connections'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Allow Remote Connections'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Remote Query Timeout (sec)'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Require Distributed Transactions'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'AdHoc Distributed Queries Enabled'
			# 			# 72
			# 			$WorksheetData[$Row++,0] = 'Database Settings'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Default Index Fill Factor'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Backup Options'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Default Backup Retention (days)'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Compress Backups'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Recovery Interval (mins)'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Database Default Locations'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Default Data Path'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Default Log Path'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Default Backup Path'
			# 			# 82
			# 			$WorksheetData[$Row++,0] = 'Advanced'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Containment'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Enabled Contained DBs'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'FILESTREAM'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'FILESTREAM Access Level'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Full-Text'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Full-Text Crawl Bandwidth Max'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Full-Text Crawl Bandwidth Min'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Full-Text Crawl Range Max'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Full-Text Notify Bandwidth Max'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Full-Text Notify Bandwidth Min'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Full-Text Precompute Rank'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Full-Text Protocol Handler Timeout'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Full-Text Transform Noise Words'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Miscellaneous'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Allow Triggers To Fire Others'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Blocked Process Threshold (sec)'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'CLR Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Cursor Threshold'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Database Mail XPs Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Default Full-Text Language'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Default Language'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Default Trace Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Disallow Results From Triggers'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Extensible Key Management Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Full-Text Upgrade Option'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'In Doubt Transaction Resolution'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Max Text Repl Size'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'OLE Automation Procs Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Optimize for Ad hoc Workloads'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Replication XPs Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Scan for Startup Procs'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Server Trigger Recursion Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Show Advanced Options'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'SMO & DMO XPs Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'SQL Agent XPs Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'SQL Mail XPs Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Two Digit Year Cutoff'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Web Assistant Procs Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'xp_cmdshell Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Network'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Network Packet Size (Bytes)'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Remote Login Timeout (sec)'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Parallelism'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Cost Threshold for Parallelism'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Locks'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Max Degree of Parallelism'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Query Wait (sec)'
			# 			# 130
			# 
			# 			$Col = 1
			# 			$SqlServerInventory.DatabaseServer | Sort-Object -Property $_.ServerName | ForEach-Object {
			# 				$Row = 0
			# 
			# 				# General
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.General.Name
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.General.Product
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.General.Edition
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.General.ProductLevel
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.General.Version
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.General.Platform
			# 				$WorksheetData[$Row++,$Col] = $_.Machine.OperatingSystem.Settings.OperatingSystem.Name
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.General.Language
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.General.MemoryMB
			# 				$WorksheetData[$Row++,$Col] = "{0:N2}" -f ($_.Server.Configuration.General.MemoryInUseKB / 1KB)
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.General.ProcessorCount
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.General.RootDirectory
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.General.ServerCollation
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.General.IsClustered
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.General.IsHadrEnabled
			# 
			# 				# Memory
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Memory.MinServerMemoryMB.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Memory.MaxServerMemoryMB.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Memory.IndexCreationMemoryKB.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Memory.MinMemoryPerQueryKB.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Memory.SetWorkingSetSize.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Memory.UseAweToAllocateMemory.RunningValue
			# 
			# 				# Processors
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Processor.AutoSetProcessorAffinityMask.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Processor.AutoSetIoAffinityMask.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Processor.AffinityMask.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Processor.Affinity64Mask.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Processor.AffinityIOMask.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Processor.Affinity64IOMask.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Processor.MaxWorkerThreads.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Processor.BoostSqlServerPriority.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Processor.UseWindowsFibers.RunningValue
			# 
			# 				# Security
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Security.AuthenticationMode
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Security.LoginAuditLevel
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Security.ServerProxyAccount.Enabled
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Security.ServerProxyAccount.Username
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Security.EnableCommonCriteriaCompliance.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Security.EnableC2AuditTracing.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Security.CrossDatabaseOwnershipChaining.RunningValue
			# 
			# 				# Connections
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.MaxConcurrentConnections.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.QueryGovernor.Enabled.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.QueryGovernor.TimeoutSeconds.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.DefaultOptions.InterimConstraintChecking.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.DefaultOptions.ImplicitTransactions.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.DefaultOptions.CursorCloseOnCommit.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.DefaultOptions.AnsiWarnings.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.DefaultOptions.AnsiPadding.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.DefaultOptions.ArithmeticAbort.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.DefaultOptions.ArithmeticIgnore.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.DefaultOptions.QuotedIdentifier.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.DefaultOptions.NoCount.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.DefaultOptions.AnsiNullDefaultOn.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.DefaultOptions.AnsiNullDefaultOff.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.DefaultOptions.ConcatNullYieldsNull.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.DefaultOptions.NumericRoundAbort.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.DefaultOptions.XactAbort.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.RemoteAdminConnectionsEnabled.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.AllowRemoteConnections.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.RemoteQueryTimeoutSeconds.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.RequireDistributedTransactions.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Connections.AdHocDistributedQueriesEnabled.RunningValue
			# 
			# 				# Database Settings
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.DatabaseSettings.IndexFillFactor.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.DatabaseSettings.BackupMediaRetentionDays.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.DatabaseSettings.CompressBackup.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.DatabaseSettings.RecoveryIntervalMinutes.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.DatabaseSettings.DataPath
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.DatabaseSettings.LogPath
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.DatabaseSettings.BackupPath
			# 
			# 				# Advanced
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Containment.EnableContainedDatabases.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Filestream.FilestreamAccessLevel.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.FullText.FullTextCrawlBandwidthMax.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.FullText.FullTextCrawlBandwidthMin.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.FullText.FullTextCrawlRangeMax.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.FullText.FullTextNotifyBandwidthMax.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.FullText.FullTextNotifyBandwidthMin.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.FullText.PrecomputeRank.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.FullText.ProtocolHandlerTimeout.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.FullText.TransformNoiseWords.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.AllowTriggersToFireOthers.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.BlockedProcessThresholdSeconds.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.ClrEnabled.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.CursorThreshold.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.DatabaseMailXPsEnabled.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.DefaultFullTextLanguage.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.DefaultLanguage.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.DefaultTraceEnabled.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.DisallowResultsFromTriggers.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.ExtensibleKeyManagementEnabled.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.FullTextUpgradeOption.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.InDoubtTransactionResolution.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.MaxTextReplicationSize.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.OleAutomationProceduresEnabled.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.OptimizeForAdHocWorkloads.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.ReplicationXPsEnabled.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.ScanForStartupProcs.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.ServerTriggerRecursionEnabled.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.ShowAdvancedOptions.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.SmoAndDmoXPsEnabled.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.SqlAgentXPsEnabled.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.SqlMailXPsEnabled.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.TwoDigitYearCutoff.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.WebAssistantProceduresEnabled.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Miscellaneous.XPCmdShellEnabled.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Network.NetworkPacketSize.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Network.RemoteLoginTimeoutSeconds.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Parallelism.CostThresholdForParallelism.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Parallelism.Locks.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Parallelism.MaxDegreeOfParallelism.RunningValue
			# 				$WorksheetData[$Row++,$Col] = $_.Server.Configuration.Advanced.Parallelism.QueryWaitSeconds.RunningValue
			# 
			# 				$Col++
			# 			}
			# 			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			# 			$Range.Value2 = $WorksheetData
			# 			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			# 			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			# 			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null
			# 
			# 			$WorksheetFormat.Add($WorksheetNumber, @{
			# 					BoldFirstRow = $false
			# 					BoldFirstColumn = $true
			# 					AutoFilter = $false
			# 					FreezeAtCell = 'B2'
			# 					ColumnFormat = @()
			# 					RowFormat = @(
			# 						@{RowNumber = 10; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 11; NumberFormat = $XlNumFmtNumberS2},
			# 						@{RowNumber = 12; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 20; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 21; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 23; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 24; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 30; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 31; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 32; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 33; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 35; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 49; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 51; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 70; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 75; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 76; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 78; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 89; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 90; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 91; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 92; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 93; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 95; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 99; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 101; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 102; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 103; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 110; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 120; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 124; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 125; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 127; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 128; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 129; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 130; NumberFormat = $XlNumFmtNumberGeneral}
			# 					)
			# 				})
			# 
			# 			$WorksheetNumber++
			#endregion


			# Worksheet 3: Server Configuration - General
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Configuration - General"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Server Config - General'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | Measure-Object).Count + 1
			$ColumnCount = 15
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Product'
			$WorksheetData[0,$Col++] = 'Edition'
			$WorksheetData[0,$Col++] = 'Level'
			$WorksheetData[0,$Col++] = 'Version'
			$WorksheetData[0,$Col++] = 'Platform'
			$WorksheetData[0,$Col++] = 'Operating System'
			$WorksheetData[0,$Col++] = 'Language'
			$WorksheetData[0,$Col++] = 'Total Memory (MB)'
			$WorksheetData[0,$Col++] = 'Instance Memory In Use (MB)'
			$WorksheetData[0,$Col++] = 'Processors'
			$WorksheetData[0,$Col++] = 'Root Directory'
			$WorksheetData[0,$Col++] = 'Server Collation'
			$WorksheetData[0,$Col++] = 'Is Clustered'
			$WorksheetData[0,$Col++] = 'AlwaysOn AG Enabled'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {
				$Col = 0
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Name
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Product
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Edition
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.ProductLevel
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Version
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Platform
				$WorksheetData[$Row,$Col++] = $_.Machine.OperatingSystem.Settings.OperatingSystem.Name
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Language
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.MemoryMB
				$WorksheetData[$Row,$Col++] = "{0:N2}" -f ($_.Server.Configuration.General.MemoryInUseKB / 1KB)
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.ProcessorCount
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.RootDirectory
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.ServerCollation
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.IsClustered
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.IsHadrEnabled
				$Row++
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'B2'
					ColumnFormat = @(
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 11; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 4: Server Configuration - Memory
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Configuration - Memory"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Server Config - Memory'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | Measure-Object).Count + 1
			$ColumnCount = 7
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Use AWE'
			$WorksheetData[0,$Col++] = 'Minimum Server Memory (MB)'
			$WorksheetData[0,$Col++] = 'Maximum Server Memory (MB)'
			$WorksheetData[0,$Col++] = 'Index Creation Memory (KB)'
			$WorksheetData[0,$Col++] = 'Min Memory Per Query (KB)'
			$WorksheetData[0,$Col++] = 'Set Working Set Size'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {
				$Col = 0
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Name
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Memory.UseAweToAllocateMemory.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Memory.MinServerMemoryMB.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Memory.MaxServerMemoryMB.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Memory.IndexCreationMemoryKB.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Memory.MinMemoryPerQueryKB.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Memory.SetWorkingSetSize.RunningValue
				$Row++
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'B2'
					ColumnFormat = @(
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 5: Server Configuration - Processors
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Configuration - Processors"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Server Config - Processors'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | Measure-Object).Count + 1
			$ColumnCount = 10
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Auto Processor Affinity Mask'
			$WorksheetData[0,$Col++] = 'Auto IO Affinity Mask'
			$WorksheetData[0,$Col++] = 'Processor Affinity Mask'
			$WorksheetData[0,$Col++] = 'Processor Affinity Mask 64'
			$WorksheetData[0,$Col++] = 'IO Affinity Mask'
			$WorksheetData[0,$Col++] = 'IO Affinity Mask 64'
			$WorksheetData[0,$Col++] = 'Max Worker Threads'
			$WorksheetData[0,$Col++] = 'Boost SQL Server Priority'
			$WorksheetData[0,$Col++] = 'Use Windows Fibers'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {
				$Col = 0
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Name
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Processor.AutoSetProcessorAffinityMask.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Processor.AutoSetIoAffinityMask.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Processor.AffinityMask.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Processor.Affinity64Mask.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Processor.AffinityIOMask.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Processor.Affinity64IOMask.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Processor.MaxWorkerThreads.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Processor.BoostSqlServerPriority.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Processor.UseWindowsFibers.RunningValue
				$Row++
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'B2'
					ColumnFormat = @(
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 6: Server Configuration - Security
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Configuration - Security"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Server Config - Security'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | Measure-Object).Count + 1
			$ColumnCount = 8
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Authentication Mode'
			$WorksheetData[0,$Col++] = 'Login Auditing'
			$WorksheetData[0,$Col++] = 'Proxy Account Enabled'
			$WorksheetData[0,$Col++] = 'Proxy Account'
			$WorksheetData[0,$Col++] = 'Common Criteria Compliance Enabled'
			$WorksheetData[0,$Col++] = 'C2 Audit Tracing Enabled'
			$WorksheetData[0,$Col++] = 'Cross Database Ownership Chaining'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {
				$Col = 0
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Name
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Security.AuthenticationMode
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Security.LoginAuditLevel
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Security.ServerProxyAccount.Enabled
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Security.ServerProxyAccount.Username
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Security.EnableCommonCriteriaCompliance.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Security.EnableC2AuditTracing.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Security.CrossDatabaseOwnershipChaining.RunningValue
				$Row++
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

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


			# Worksheet 7: Server Configuration - Connections
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Configuration - Connections"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Server Config - Connections'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | Measure-Object).Count + 1
			$ColumnCount = 23
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Max Concurrent Connections'
			$WorksheetData[0,$Col++] = 'Query Gov. Enabled'
			$WorksheetData[0,$Col++] = 'Query Gov. Timeout (sec)'
			$WorksheetData[0,$Col++] = 'Intermin/Deferred Constraint Checking Default'
			$WorksheetData[0,$Col++] = 'Implicit Transactions Default'
			$WorksheetData[0,$Col++] = 'Cursor Close On Commit Default'
			$WorksheetData[0,$Col++] = 'Ansi Warnings Default'
			$WorksheetData[0,$Col++] = 'Ansi Padding Default'
			$WorksheetData[0,$Col++] = 'Arithmatic Abort Default'
			$WorksheetData[0,$Col++] = 'Arithmatic Ignore Default'
			$WorksheetData[0,$Col++] = 'Quoted Identifier Default'
			$WorksheetData[0,$Col++] = 'No Count Default'
			$WorksheetData[0,$Col++] = 'ANSI NULL Default On'
			$WorksheetData[0,$Col++] = 'ANSI NULL Default Off'
			$WorksheetData[0,$Col++] = 'Concat Null Yields Null Default'
			$WorksheetData[0,$Col++] = 'Numeric Round Abort Default'
			$WorksheetData[0,$Col++] = 'Xact Abort Default'
			$WorksheetData[0,$Col++] = 'Allow Remote Admin Connections'
			$WorksheetData[0,$Col++] = 'Allow Remote Connections'
			$WorksheetData[0,$Col++] = 'Remote Query Timeout (sec)'
			$WorksheetData[0,$Col++] = 'Require Distributed Transactions'
			$WorksheetData[0,$Col++] = 'AdHoc Distributed Queries Enabled'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {
				$Col = 0
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Name
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.MaxConcurrentConnections.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.QueryGovernor.Enabled.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.QueryGovernor.TimeoutSeconds.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.DefaultOptions.InterimConstraintChecking.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.DefaultOptions.ImplicitTransactions.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.DefaultOptions.CursorCloseOnCommit.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.DefaultOptions.AnsiWarnings.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.DefaultOptions.AnsiPadding.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.DefaultOptions.ArithmeticAbort.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.DefaultOptions.ArithmeticIgnore.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.DefaultOptions.QuotedIdentifier.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.DefaultOptions.NoCount.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.DefaultOptions.AnsiNullDefaultOn.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.DefaultOptions.AnsiNullDefaultOff.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.DefaultOptions.ConcatNullYieldsNull.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.DefaultOptions.NumericRoundAbort.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.DefaultOptions.XactAbort.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.RemoteAdminConnectionsEnabled.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.AllowRemoteConnections.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.RemoteQueryTimeoutSeconds.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.RequireDistributedTransactions.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Connections.AdHocDistributedQueriesEnabled.RunningValue
				$Row++
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'B2'
					ColumnFormat = @(
						@{ColumnNumber = 2; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 22; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 8: Server Configuration - Database Settings
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Configuration - Database Settings"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Server Config - DB Settings'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | Measure-Object).Count + 1
			$ColumnCount = 8
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Default Index Fill Factor'
			$WorksheetData[0,$Col++] = 'Default Backup Retention (days)'
			$WorksheetData[0,$Col++] = 'Compress Backups'
			$WorksheetData[0,$Col++] = 'Recovery Interval (mins)'
			$WorksheetData[0,$Col++] = 'Default Data Path'
			$WorksheetData[0,$Col++] = 'Default Log Path'
			$WorksheetData[0,$Col++] = 'Default Backup Path'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {
				$Col = 0
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Name
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.DatabaseSettings.IndexFillFactor.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.DatabaseSettings.BackupMediaRetentionDays.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.DatabaseSettings.CompressBackup.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.DatabaseSettings.RecoveryIntervalMinutes.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.DatabaseSettings.DataPath
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.DatabaseSettings.LogPath
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.DatabaseSettings.BackupPath
				$Row++
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'B2'
					ColumnFormat = @(
						@{ColumnNumber = 2; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 3; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 9: Server Configuration - Advanced
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Configuration - Advanced"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Server Config - Advanced'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | Measure-Object).Count + 1
			$ColumnCount = 42
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Enabled Contained DBs'
			$WorksheetData[0,$Col++] = 'FILESTREAM Access Level'
			$WorksheetData[0,$Col++] = 'Full-Text Crawl Bandwidth Max'
			$WorksheetData[0,$Col++] = 'Full-Text Crawl Bandwidth Min'
			$WorksheetData[0,$Col++] = 'Full-Text Crawl Range Max'
			$WorksheetData[0,$Col++] = 'Full-Text Notify Bandwidth Max'
			$WorksheetData[0,$Col++] = 'Full-Text Notify Bandwidth Min'
			$WorksheetData[0,$Col++] = 'Full-Text Precompute Rank'
			$WorksheetData[0,$Col++] = 'Full-Text Protocol Handler Timeout'
			$WorksheetData[0,$Col++] = 'Full-Text Transform Noise Words'
			$WorksheetData[0,$Col++] = 'Allow Triggers To Fire Others'
			$WorksheetData[0,$Col++] = 'Blocked Process Threshold (sec)'
			$WorksheetData[0,$Col++] = 'CLR Enabled'
			$WorksheetData[0,$Col++] = 'Cursor Threshold'
			$WorksheetData[0,$Col++] = 'Database Mail XPs Enabled'
			$WorksheetData[0,$Col++] = 'Default Full-Text Language'
			$WorksheetData[0,$Col++] = 'Default Language'
			$WorksheetData[0,$Col++] = 'Default Trace Enabled'
			$WorksheetData[0,$Col++] = 'Disallow Results From Triggers'
			$WorksheetData[0,$Col++] = 'Extensible Key Management Enabled'
			$WorksheetData[0,$Col++] = 'Full-Text Upgrade Option'
			$WorksheetData[0,$Col++] = 'In Doubt Transaction Resolution'
			$WorksheetData[0,$Col++] = 'Max Text Repl Size'
			$WorksheetData[0,$Col++] = 'OLE Automation Procs Enabled'
			$WorksheetData[0,$Col++] = 'Optimize for Ad hoc Workloads'
			$WorksheetData[0,$Col++] = 'Replication XPs Enabled'
			$WorksheetData[0,$Col++] = 'Scan for Startup Procs'
			$WorksheetData[0,$Col++] = 'Server Trigger Recursion Enabled'
			$WorksheetData[0,$Col++] = 'Show Advanced Options'
			$WorksheetData[0,$Col++] = 'SMO & DMO XPs Enabled'
			$WorksheetData[0,$Col++] = 'SQL Agent XPs Enabled'
			$WorksheetData[0,$Col++] = 'SQL Mail XPs Enabled'
			$WorksheetData[0,$Col++] = 'Two Digit Year Cutoff'
			$WorksheetData[0,$Col++] = 'Web Assistant Procs Enabled'
			$WorksheetData[0,$Col++] = 'xp_cmdshell Enabled'
			$WorksheetData[0,$Col++] = 'Network Packet Size (Bytes)'
			$WorksheetData[0,$Col++] = 'Remote Login Timeout (sec)'
			$WorksheetData[0,$Col++] = 'Cost Threshold for Parallelism'
			$WorksheetData[0,$Col++] = 'Locks'
			$WorksheetData[0,$Col++] = 'Max Degree of Parallelism'
			$WorksheetData[0,$Col++] = 'Query Wait (sec)'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {
				$Col = 0
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Name
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Containment.EnableContainedDatabases.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Filestream.FilestreamAccessLevel.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.FullText.FullTextCrawlBandwidthMax.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.FullText.FullTextCrawlBandwidthMin.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.FullText.FullTextCrawlRangeMax.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.FullText.FullTextNotifyBandwidthMax.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.FullText.FullTextNotifyBandwidthMin.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.FullText.PrecomputeRank.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.FullText.ProtocolHandlerTimeout.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.FullText.TransformNoiseWords.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.AllowTriggersToFireOthers.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.BlockedProcessThresholdSeconds.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.ClrEnabled.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.CursorThreshold.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.DatabaseMailXPsEnabled.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.DefaultFullTextLanguage.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.DefaultLanguage.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.DefaultTraceEnabled.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.DisallowResultsFromTriggers.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.ExtensibleKeyManagementEnabled.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.FullTextUpgradeOption.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.InDoubtTransactionResolution.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.MaxTextReplicationSize.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.OleAutomationProceduresEnabled.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.OptimizeForAdHocWorkloads.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.ReplicationXPsEnabled.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.ScanForStartupProcs.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.ServerTriggerRecursionEnabled.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.ShowAdvancedOptions.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.SmoAndDmoXPsEnabled.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.SqlAgentXPsEnabled.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.SqlMailXPsEnabled.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.TwoDigitYearCutoff.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.WebAssistantProceduresEnabled.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Miscellaneous.XPCmdShellEnabled.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Network.NetworkPacketSize.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Network.RemoteLoginTimeoutSeconds.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Parallelism.CostThresholdForParallelism.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Parallelism.Locks.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Parallelism.MaxDegreeOfParallelism.RunningValue
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.Advanced.Parallelism.QueryWaitSeconds.RunningValue

				$Row++
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'B2'
					ColumnFormat = @(
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 17; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 18; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 24; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 34; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 37; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 38; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 39; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 40; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 41; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 42; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 10: Server Configuration - Clustering
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Configuration - Clustering"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Server Config - Clustering'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | Measure-Object).Count + 1
			$ColumnCount = 5
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Clustered'
			$WorksheetData[0,$Col++] = 'Cluster Members'
			$WorksheetData[0,$Col++] = 'Current Owner(s)'
			$WorksheetData[0,$Col++] = 'Shared Drives'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {
				$Col = 0
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Name
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.HighAvailability.FailoverCluster.IsClusteredInstance
				$WorksheetData[$Row,$Col++] = (
					$_.Server.Configuration.HighAvailability.FailoverCluster.Member | ForEach-Object {
						'{0}{1}' -f $_.Name, $(if ($_.Status -ne $null) { '(' + $_.Status + ')' } else { [String]::Empty })
					}
				) -join $Delimiter
				$WorksheetData[$Row,$Col++] = (
					$_.Server.Configuration.HighAvailability.FailoverCluster.Member | Where-Object { 
						$_.IsCurrentOwner -eq $true 
					} | ForEach-Object {
						$_.Name
					} | Sort-Object
				) -join $Delimiter
				$WorksheetData[$Row,$Col++] = ($_.Server.Configuration.HighAvailability.FailoverCluster.SharedDrive | Sort-Object ) -join $Delimiter

				$Row++
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

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


			# Worksheet 11: Server Configuration - AlwaysOn
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Configuration - AlwaysOn"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Server Config - AlwaysOn'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | Measure-Object).Count + 1
			$ColumnCount = 7
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'AlwaysOn Enabled'
			$WorksheetData[0,$Col++] = 'AlwaysOn Manager Status'
			$WorksheetData[0,$Col++] = 'Windows Failover Cluster Name'
			$WorksheetData[0,$Col++] = 'Quorum Type'
			$WorksheetData[0,$Col++] = 'Quorum State'
			$WorksheetData[0,$Col++] = 'Cluster Members'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {
				$Col = 0
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Name
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.HighAvailability.AlwaysOn.IsAlwaysOnEnabled
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.HighAvailability.AlwaysOn.AlwaysOnManagerStatus
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.HighAvailability.AlwaysOn.WindowsFailoverCluster.Name
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.HighAvailability.AlwaysOn.WindowsFailoverCluster.QuorumType
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.HighAvailability.AlwaysOn.WindowsFailoverCluster.QuorumState
				$WorksheetData[$Row,$Col++] = (
					$_.Server.Configuration.HighAvailability.AlwaysOn.WindowsFailoverCluster.Member | Where-Object { $_.Name } | ForEach-Object {
						'{0} ({1}, {2}, {3} vote{4})' -f @(
							$_.Name,
							$_.MemberType,
							$_.MemberState,
							$_.NumberOfQuorumVotes,
							$(if ($_.NumberOfQuorumVotes -eq 1) { [String]::Empty } else { 's' }) 
						)
					}
				) -join $Delimiter

				$Row++
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

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


			# Worksheet 12: Server Configuration - Permissions
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Configuration - Permissions"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Server Config - Permissions'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Configuration.Permissions | Where-Object { $_.PermissionType } } | Measure-Object).Count + 1
			$ColumnCount = 11
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Object Type'
			$WorksheetData[0,$Col++] = 'Object Schema'
			$WorksheetData[0,$Col++] = 'Object Name'
			$WorksheetData[0,$Col++] = 'Action'
			$WorksheetData[0,$Col++] = 'Permission'
			$WorksheetData[0,$Col++] = 'Column'
			$WorksheetData[0,$Col++] = 'Granted To'
			$WorksheetData[0,$Col++] = 'Grantee Type'
			$WorksheetData[0,$Col++] = 'Granted By'
			$WorksheetData[0,$Col++] = 'Grantor Type'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Configuration.Permissions | Where-Object { $_.PermissionType } | 
				Sort-Object -Property ObjectClass, ObjectName, PermissionState, PermissionType, Grantee | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.ObjectClass
					$WorksheetData[$Row,$Col++] = $_.ObjectSchema
					$WorksheetData[$Row,$Col++] = $_.ObjectName
					$WorksheetData[$Row,$Col++] = $_.PermissionState
					$WorksheetData[$Row,$Col++] = $_.PermissionType
					$WorksheetData[$Row,$Col++] = $_.ColumnName
					$WorksheetData[$Row,$Col++] = $_.Grantee
					$WorksheetData[$Row,$Col++] = $_.GranteeType
					$WorksheetData[$Row,$Col++] = $_.Grantor
					$WorksheetData[$Row,$Col++] = $_.GrantorType
					$Row++
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 13: Server Security - Logins
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Security - Logins"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Security - Server Logins'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $SecurityTabColor #$ServerTabColor

			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Security.Logins | Where-Object { $_.Sid } }) | Measure-Object).Count + 1
			$ColumnCount = 20
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Login Name'
			$WorksheetData[0,$Col++] = 'Create Date'
			$WorksheetData[0,$Col++] = 'Type'
			$WorksheetData[0,$Col++] = 'Default Database'
			$WorksheetData[0,$Col++] = 'Password Expiration Enabled'
			$WorksheetData[0,$Col++] = 'Password Hash Algorithm'
			$WorksheetData[0,$Col++] = 'Password Policy Enforced'
			$WorksheetData[0,$Col++] = 'Credential'
			$WorksheetData[0,$Col++] = 'Certificate'
			$WorksheetData[0,$Col++] = 'Asymmetric Key' 
			$WorksheetData[0,$Col++] = 'Windows Login Access Type'
			$WorksheetData[0,$Col++] = 'Has Access'
			$WorksheetData[0,$Col++] = 'Is Disabled'
			$WorksheetData[0,$Col++] = 'Is Locked'
			$WorksheetData[0,$Col++] = 'Deny Windows Login'
			$WorksheetData[0,$Col++] = 'Is Password Expired'
			$WorksheetData[0,$Col++] = 'Is System Object'
			$WorksheetData[0,$Col++] = 'Password Is Blank'
			$WorksheetData[0,$Col++] = 'Password Is Login Name'


			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Security.Logins | Where-Object { $_.Sid } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.Name
					$WorksheetData[$Row,$Col++] = $_.CreateDate
					$WorksheetData[$Row,$Col++] = $_.LoginType
					$WorksheetData[$Row,$Col++] = $_.DefaultDatabase
					$WorksheetData[$Row,$Col++] = $_.PasswordExpirationEnabled
					$WorksheetData[$Row,$Col++] = $_.PasswordHashAlgorithm
					$WorksheetData[$Row,$Col++] = $_.PasswordPolicyEnforced
					$WorksheetData[$Row,$Col++] = $_.Credential
					$WorksheetData[$Row,$Col++] = $_.Certificate
					$WorksheetData[$Row,$Col++] = $_.AsymmetricKey
					$WorksheetData[$Row,$Col++] = $_.WindowsLoginAccessType
					$WorksheetData[$Row,$Col++] = $_.HasAccess
					$WorksheetData[$Row,$Col++] = $_.IsDisabled
					$WorksheetData[$Row,$Col++] = $_.IsLocked
					$WorksheetData[$Row,$Col++] = $_.DenyWindowsLogin
					$WorksheetData[$Row,$Col++] = $_.IsPasswordExpired
					$WorksheetData[$Row,$Col++] = $_.IsSystemObject
					$WorksheetData[$Row,$Col++] = $_.HasBlankPassword
					$WorksheetData[$Row,$Col++] = $_.HasNameAsPassword
					$Row++
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 3; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 14: Server Security - Server Roles
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Security - Server Roles"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Security - Server Roles'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $SecurityTabColor #$ServerTabColor

			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Security.ServerRoles | Where-Object { $_.Name } }) | Measure-Object).Count + 1
			$ColumnCount = 8
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Role Name'
			$WorksheetData[0,$Col++] = 'Create Date'
			$WorksheetData[0,$Col++] = 'Modify Date'
			$WorksheetData[0,$Col++] = 'Owner'
			$WorksheetData[0,$Col++] = 'IsFixedRole'
			$WorksheetData[0,$Col++] = 'Members'
			$WorksheetData[0,$Col++] = 'Member Of'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Security.ServerRoles | Where-Object { $_.Name } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.Name
					$WorksheetData[$Row,$Col++] = $_.CreateDate
					$WorksheetData[$Row,$Col++] = $_.DateModified
					$WorksheetData[$Row,$Col++] = $_.Owner
					$WorksheetData[$Row,$Col++] = $_.IsFixedRole
					$WorksheetData[$Row,$Col++] = ($_.Member | Sort-Object) -join $Delimiter # PoSH is more forgiving here than [String]::Join if $_.Member is $NULL
					$WorksheetData[$Row,$Col++] = ($_.MemberOf | Sort-Object) -join $Delimiter # PoSH is more forgiving here than [String]::Join if $_.MemberOf is NUL

					#$WorksheetData[$Row,$Col++] = if (@($_.Member).Count -gt 0) { [String]::Join($Delimiter, ($_.Member | Sort-Object)) } else { [String]::Empty }
					#$WorksheetData[$Row,$Col++] = if (@($_.MemberOf).Count -gt 0) { [String]::Join($Delimiter, ($_.MemberOf | Sort-Object)) } else { [String]::Empty }
					$Row++
				}

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 3; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 15: Server Security - Credentials
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Security - Credentials"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Security - Credentials'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $SecurityTabColor #$ServerTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Security.Credentials }) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Security.Credentials | Where-Object { $_.ID } }) | Measure-Object).Count + 1
			$ColumnCount = 7
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Credential Name'
			$WorksheetData[0,$Col++] = 'Identity'
			$WorksheetData[0,$Col++] = 'Encryption Provider'
			$WorksheetData[0,$Col++] = 'Mapped Class Type'
			$WorksheetData[0,$Col++] = 'Create Date'
			$WorksheetData[0,$Col++] = 'Modify Date'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Security.Credentials | Where-Object { $_.ID } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.Name
					$WorksheetData[$Row,$Col++] = $_.Identity
					$WorksheetData[$Row,$Col++] = $_.ProviderName
					$WorksheetData[$Row,$Col++] = $_.MappedClassType
					$WorksheetData[$Row,$Col++] = $_.CreateDate
					$WorksheetData[$Row,$Col++] = $_.DateModified
					$Row++
				}

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 16: Server Security - Audits
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Security - Audits"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Security - Audits'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $SecurityTabColor #$ServerTabColor

			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Security.Audits | Where-Object { $_.General.ID } }) | Measure-Object).Count + 1
			$ColumnCount = 16
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Audit Name'
			$WorksheetData[0,$Col++] = 'Enabled'
			$WorksheetData[0,$Col++] = 'Queue Delay (Milliseconds)'
			$WorksheetData[0,$Col++] = 'On Audit Log Failure'
			$WorksheetData[0,$Col++] = 'Audit Destination'
			$WorksheetData[0,$Col++] = 'File Path'
			$WorksheetData[0,$Col++] = 'File Name'
			$WorksheetData[0,$Col++] = 'Max Rollover Files'
			$WorksheetData[0,$Col++] = 'Max Files'
			$WorksheetData[0,$Col++] = 'Max File Size'
			$WorksheetData[0,$Col++] = 'Max File Size Unit'
			$WorksheetData[0,$Col++] = 'Reserve Disk Space'
			$WorksheetData[0,$Col++] = 'Filter'
			$WorksheetData[0,$Col++] = 'Create Date'
			$WorksheetData[0,$Col++] = 'Modify Date'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Security.Audits | Where-Object { $_.General.ID } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.General.AuditName
					$WorksheetData[$Row,$Col++] = $_.General.Enabled
					$WorksheetData[$Row,$Col++] = $_.General.QueueDelayMilliseconds
					$WorksheetData[$Row,$Col++] = $_.General.OnAuditLogFailure
					$WorksheetData[$Row,$Col++] = $_.General.AuditDestination
					$WorksheetData[$Row,$Col++] = $_.General.FilePath
					$WorksheetData[$Row,$Col++] = $_.General.FileName
					$WorksheetData[$Row,$Col++] = $_.General.MaximumRolloverFiles
					$WorksheetData[$Row,$Col++] = $_.General.MaximumFiles
					$WorksheetData[$Row,$Col++] = $_.General.MaximumFileSize
					$WorksheetData[$Row,$Col++] = $_.General.MaximumFileSizeUnit
					$WorksheetData[$Row,$Col++] = $_.General.ReserveDiskSpace
					$WorksheetData[$Row,$Col++] = $_.Filter
					$WorksheetData[$Row,$Col++] = $_.General.CreateDate
					$WorksheetData[$Row,$Col++] = $_.General.DateModified
					$Row++
				}

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 11; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 16; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 17: Server Security - Audit Specifications
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Security - Audit Specifications"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Security - Audit Specifications'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $SecurityTabColor #$ServerTabColor

			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Security.ServerAuditSpecifications | ForEach-Object { $_.Actions | Where-Object { $_.Action } } }) | Measure-Object).Count + 1
			$ColumnCount = 8
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Audit Name'
			$WorksheetData[0,$Col++] = 'Audit Specification Name'
			$WorksheetData[0,$Col++] = 'Audit Action Type'
			$WorksheetData[0,$Col++] = 'Object Class'
			$WorksheetData[0,$Col++] = 'Object Schema'
			$WorksheetData[0,$Col++] = 'Object Name'
			$WorksheetData[0,$Col++] = 'Principal Name'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Security.ServerAuditSpecifications | Where-Object { $_.ID } | ForEach-Object {
					#$_.Server.Security.ServerAuditSpecifications | Where-Object { ($_ | Measure-Object).Count -gt 0 } | ForEach-Object {

					$AuditName = $_.AuditName
					$AuditSpecificationName = $_.Name

					$_.Actions | Where-Object { $_.Action } | ForEach-Object {
						#$_.Actions | Where-Object { ($_ | Measure-Object).Count -gt 0 } | ForEach-Object {

						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName # $_.Actions
						$WorksheetData[$Row,$Col++] = $AuditName
						$WorksheetData[$Row,$Col++] = $AuditSpecificationName
						$WorksheetData[$Row,$Col++] = $_.Action
						$WorksheetData[$Row,$Col++] = $_.ObjectClass
						$WorksheetData[$Row,$Col++] = $_.ObjectSchema
						$WorksheetData[$Row,$Col++] = $_.ObjectName
						$WorksheetData[$Row,$Col++] = $_.Principal
						$Row++
					}
				}

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 18: Server Objects - Endpoints
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Objects - Endpoints"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'SVR Objects - Endpoints'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor #$ServerTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.ServerObjects.Endpoints }) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.ServerObjects.Endpoints | Where-Object { $_.ID } }) | Measure-Object).Count + 1
			$ColumnCount = 8
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Endpoint Name'
			$WorksheetData[0,$Col++] = 'Owner'
			$WorksheetData[0,$Col++] = 'Type'
			$WorksheetData[0,$Col++] = 'Protocol'
			$WorksheetData[0,$Col++] = 'Status'
			$WorksheetData[0,$Col++] = 'Is Admin Endpoint'
			$WorksheetData[0,$Col++] = 'Is System Object'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.ServerObjects.Endpoints | Where-Object { $_.ID } | ForEach-Object {
					#$_.Server.ServerObjects.Endpoints | Where-Object { ($_ | Measure-Object).Count -gt 0 } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.Name
					$WorksheetData[$Row,$Col++] = $_.Owner
					$WorksheetData[$Row,$Col++] = $_.EndpointType
					$WorksheetData[$Row,$Col++] = $_.ProtocolType
					$WorksheetData[$Row,$Col++] = $_.EndpointState
					$WorksheetData[$Row,$Col++] = $_.IsAdminEndpoint
					$WorksheetData[$Row,$Col++] = $_.IsSystemObject
					$Row++
				}

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 3; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# 			# Worksheet 19: Server Objects - Linked Server Configuration (VERTICAL FORMAT)
			# 			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Objects - Linked Server Configuration"
			# 			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			# 			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			# 			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			# 			$Worksheet.Name = 'SVR Objects - Linked Server Cfg'
			# 			#$Worksheet.Tab.Color = $ServerTabColor
			# 			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor #$ServerTabColor
			# 
			# 			$ColumnCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.ServerObjects.LinkedServers }) | Measure-Object).Count + 1
			# 
			# 			$RowCount = 24
			# 			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount
			# 
			# 			$Row = 0
			# 			$WorksheetData[$Row++,0] = 'Server Name'
			# 			$WorksheetData[$Row++,0] = 'General'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Linked Server'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Product Name'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Provider Name'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Data Source'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Provider String'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Location'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Catalog'
			# 			#9
			# 			$WorksheetData[$Row++,0] = 'Server Options'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Collation Compatible'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Data Access'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'RPC'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'RPC Out'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Use Remote Collation'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Collation Name'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Connection Timeout (sec)'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Query Timeout (sec)'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Distributor'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Distribution Publisher'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Publisher'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Subscriber'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Lazy Schema Validation'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Enable Promotion of Distributed Transactions'
			# 			#24
			# 
			# 			$Col = 1
			# 			$SqlServerInventory.DatabaseServer | ForEach-Object {
			# 				$ServerName = $_.Server.Configuration.General.Name
			# 
			# 				$_.Server.ServerObjects.LinkedServers | Sort-Object -Property $_.Name | ForEach-Object {
			# 
			# 					$Row = 0
			# 
			# 					# General
			# 					$WorksheetData[$Row++,$Col] = $ServerName
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.General.Name
			# 					$WorksheetData[$Row++,$Col] = $_.General.ProductName
			# 					$WorksheetData[$Row++,$Col] = $_.General.ProviderName
			# 					$WorksheetData[$Row++,$Col] = $_.General.DataSource
			# 					$WorksheetData[$Row++,$Col] = $_.General.ProviderString
			# 					$WorksheetData[$Row++,$Col] = $_.General.Location
			# 					$WorksheetData[$Row++,$Col] = $_.General.Catalog
			# 
			# 					# Server Options
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.Options.CollationCompatible
			# 					$WorksheetData[$Row++,$Col] = $_.Options.DataAccess
			# 					$WorksheetData[$Row++,$Col] = $_.Options.Rpc
			# 					$WorksheetData[$Row++,$Col] = $_.Options.RpcOut
			# 					$WorksheetData[$Row++,$Col] = $_.Options.UseRemoteCollation
			# 					$WorksheetData[$Row++,$Col] = $_.Options.CollationName
			# 					$WorksheetData[$Row++,$Col] = $_.Options.ConnectTimeoutSeconds
			# 					$WorksheetData[$Row++,$Col] = $_.Options.QueryTimeoutSeconds
			# 					$WorksheetData[$Row++,$Col] = $_.Options.Distributor
			# 					$WorksheetData[$Row++,$Col] = $_.Options.DistPublisher
			# 					$WorksheetData[$Row++,$Col] = $_.Options.Publisher
			# 					$WorksheetData[$Row++,$Col] = $_.Options.Subscriber
			# 					$WorksheetData[$Row++,$Col] = $_.Options.LazySchemaValidation
			# 					$WorksheetData[$Row++,$Col] = $_.Options.IsPromotionofDistributedTransactionsForRPCEnabled
			# 
			# 					$Col++
			# 				}
			# 			}
			# 
			# 			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			# 			$Range.Value2 = $WorksheetData
			# 			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			# 			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			# 			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null
			# 
			# 			$WorksheetFormat.Add($WorksheetNumber, @{
			# 					BoldFirstRow = $false
			# 					BoldFirstColumn = $true
			# 					AutoFilter = $false
			# 					FreezeAtCell = 'B2'
			# 					ColumnFormat = @()
			# 					RowFormat = @(
			# 						@{RowNumber = 17; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 18; NumberFormat = $XlNumFmtNumberS0}
			# 					)
			# 				})
			# 
			# 			$WorksheetNumber++
			#endregion


			# Worksheet 19: Server Objects - Linked Server Configuration
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Objects - Linked Server Configuration"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'SVR Objects - Linked Svr Config'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor #$ServerTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.ServerObjects.LinkedServers }) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.ServerObjects.LinkedServers | Where-Object { $_.General.Name } }) | Measure-Object).Count + 1
			$ColumnCount = 24
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'General >>'
			$WorksheetData[0,$Col++] = 'Linked Server'
			$WorksheetData[0,$Col++] = 'Product Name'
			$WorksheetData[0,$Col++] = 'Provider Name'
			$WorksheetData[0,$Col++] = 'Data Source'
			$WorksheetData[0,$Col++] = 'Provider String'
			$WorksheetData[0,$Col++] = 'Location'
			$WorksheetData[0,$Col++] = 'Catalog'
			#9
			$WorksheetData[0,$Col++] = 'Server Options >>'
			$WorksheetData[0,$Col++] = 'Collation Compatible'
			$WorksheetData[0,$Col++] = 'Data Access'
			$WorksheetData[0,$Col++] = 'RPC'
			$WorksheetData[0,$Col++] = 'RPC Out'
			$WorksheetData[0,$Col++] = 'Use Remote Collation'
			$WorksheetData[0,$Col++] = 'Collation Name'
			$WorksheetData[0,$Col++] = 'Connection Timeout (sec)'
			$WorksheetData[0,$Col++] = 'Query Timeout (sec)'
			$WorksheetData[0,$Col++] = 'Distributor'
			$WorksheetData[0,$Col++] = 'Distribution Publisher'
			$WorksheetData[0,$Col++] = 'Publisher'
			$WorksheetData[0,$Col++] = 'Subscriber'
			$WorksheetData[0,$Col++] = 'Lazy Schema Validation'
			$WorksheetData[0,$Col++] = 'Enable Promotion of Distributed Transactions'
			#24

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {
				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.ServerObjects.LinkedServers | Where-Object { $_.General.Name } | ForEach-Object {

					$Col = 0

					# General
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $null
					$WorksheetData[$Row,$Col++] = $_.General.Name
					$WorksheetData[$Row,$Col++] = $_.General.ProductName
					$WorksheetData[$Row,$Col++] = $_.General.ProviderName
					$WorksheetData[$Row,$Col++] = $_.General.DataSource
					$WorksheetData[$Row,$Col++] = $_.General.ProviderString
					$WorksheetData[$Row,$Col++] = $_.General.Location
					$WorksheetData[$Row,$Col++] = $_.General.Catalog

					# Server Options
					$WorksheetData[$Row,$Col++] = $null
					$WorksheetData[$Row,$Col++] = $_.Options.CollationCompatible
					$WorksheetData[$Row,$Col++] = $_.Options.DataAccess
					$WorksheetData[$Row,$Col++] = $_.Options.Rpc
					$WorksheetData[$Row,$Col++] = $_.Options.RpcOut
					$WorksheetData[$Row,$Col++] = $_.Options.UseRemoteCollation
					$WorksheetData[$Row,$Col++] = $_.Options.CollationName
					$WorksheetData[$Row,$Col++] = $_.Options.ConnectTimeoutSeconds
					$WorksheetData[$Row,$Col++] = $_.Options.QueryTimeoutSeconds
					$WorksheetData[$Row,$Col++] = $_.Options.Distributor
					$WorksheetData[$Row,$Col++] = $_.Options.DistPublisher
					$WorksheetData[$Row,$Col++] = $_.Options.Publisher
					$WorksheetData[$Row,$Col++] = $_.Options.Subscriber
					$WorksheetData[$Row,$Col++] = $_.Options.LazySchemaValidation
					$WorksheetData[$Row,$Col++] = $_.Options.IsPromotionofDistributedTransactionsForRPCEnabled

					$Row++
				}
			}

			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'B2'
					ColumnFormat = @( 
						@{ColumnNumber = 17; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 18; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 20: Server Objects - Linked Server Logins
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Objects - Linked Server Logins"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'SVR Objects - Linked Svr Logins'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor #$ServerTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.ServerObjects.LinkedServers | ForEach-Object { $_.Security } } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.ServerObjects.LinkedServers | ForEach-Object { $_.Security | Where-Object { $_.LocalLogin } } } ) | Measure-Object).Count + 1
			$ColumnCount = 6
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Linked Server'
			$WorksheetData[0,$Col++] = 'Local Login'
			$WorksheetData[0,$Col++] = 'Impersonate'
			$WorksheetData[0,$Col++] = 'Remote User'
			$WorksheetData[0,$Col++] = 'Last Modified Date'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.ServerObjects.LinkedServers | ForEach-Object {

					$LinkedServerName = $_.General.Name

					# $_.Security can be null so use Measure-Object to ensure we're only writing results that have data
					$_.Security | Where-Object { $_.LocalLogin } | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $LinkedServerName
						$WorksheetData[$Row,$Col++] = $_.LocalLogin
						$WorksheetData[$Row,$Col++] = $_.Impersonate
						$WorksheetData[$Row,$Col++] = $_.RemoteUser
						$WorksheetData[$Row,$Col++] = $_.DateLastModified
						$Row++
					}


				}

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'D2'
					ColumnFormat = @(
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 21: Server Objects - Startup Procedures
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Objects - Startup Procedures"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'SVR Objects - Startup Procs'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor #$ServerTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.ServerObjects.StartupProcedures } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.ServerObjects.StartupProcedures | Where-Object { $_.Name } } ) | Measure-Object).Count + 1
			$ColumnCount = 3
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Schema'
			$WorksheetData[0,$Col++] = 'Procedure Name'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.ServerObjects.StartupProcedures | Where-Object { $_.Name } | ForEach-Object {
					#$_.Server.ServerObjects.StartupProcedures | Where-Object { ($_ | Measure-Object).Count -gt 0 } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.Schema
					$WorksheetData[$Row,$Col++] = $_.Name
					$Row++
				}

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType,S $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

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


			# Worksheet 22: Server Objects - Server Triggers
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Server Objects - Server Triggers"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'SVR Objects - Server Triggers'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor #$ServerTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.ServerObjects.Triggers } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.ServerObjects.Triggers | Where-Object { $_.ID } } ) | Measure-Object).Count + 1
			$ColumnCount = 16
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Trigger Name'
			$WorksheetData[0,$Col++] = 'Enabled'
			$WorksheetData[0,$Col++] = 'Encrypted'
			$WorksheetData[0,$Col++] = 'System Object'
			$WorksheetData[0,$Col++] = 'Events'
			$WorksheetData[0,$Col++] = 'Execution Context'
			$WorksheetData[0,$Col++] = 'Execution Login'
			$WorksheetData[0,$Col++] = 'ANSI Nulls Status'
			$WorksheetData[0,$Col++] = 'Quoted Identifier Status'
			$WorksheetData[0,$Col++] = 'Implementation Type'
			$WorksheetData[0,$Col++] = 'CLR Assembly Name'
			$WorksheetData[0,$Col++] = 'CLR Class Name'
			$WorksheetData[0,$Col++] = 'CLR Method Name'
			$WorksheetData[0,$Col++] = 'Create Date'
			$WorksheetData[0,$Col++] = 'Modified Date'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.ServerObjects.Triggers | Where-Object { $_.ID } | ForEach-Object {
					#$_.Server.ServerObjects.Triggers | Where-Object { ($_ | Measure-Object).Count -gt 0 } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.Name
					$WorksheetData[$Row,$Col++] = $_.IsEnabled
					$WorksheetData[$Row,$Col++] = $_.IsEncrypted
					$WorksheetData[$Row,$Col++] = $_.IsSystemObject
					$WorksheetData[$Row,$Col++] = $_.DdlTriggerEvents
					$WorksheetData[$Row,$Col++] = $_.ExecutionContext
					$WorksheetData[$Row,$Col++] = $_.ExecutionContextLogin
					$WorksheetData[$Row,$Col++] = $_.AnsiNullsStatus
					$WorksheetData[$Row,$Col++] = $_.QuotedIdentifierStatus
					$WorksheetData[$Row,$Col++] = $_.ImplementationType
					$WorksheetData[$Row,$Col++] = $_.AssemblyName
					$WorksheetData[$Row,$Col++] = $_.ClassName
					$WorksheetData[$Row,$Col++] = $_.MethodName
					$WorksheetData[$Row,$Col++] = $_.CreateDate
					$WorksheetData[$Row,$Col++] = $_.DateLastModified
					$Row++
				}

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtDate}
						@{ColumnNumber = 16; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 23: Management - Resource Governor
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Management - Resource Governor"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Management - Resource Governor'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ManagementTabColor #$ServerTabColor

			$RowCount = (
				(
					$SqlServerInventory.DatabaseServer | ForEach-Object {
						if (($_.Server.Management.ResourceGovernor.ResourcePools | Where-Object { $_.ID } | Measure-Object).Count -eq 0) {
							1
						} else {
							$_.Server.Management.ResourceGovernor.ResourcePools | Where-Object { $_.ID } | ForEach-Object {
								$_.WorkloadGroups | Where-Object { $_.ID } | Measure-Object | ForEach-Object {
									if ($_.Count -eq 0) { 1 } else { $_.Count }
								}
							}
						}
					}
				) | Measure-Object -Sum
			).Sum + 1
			$ColumnCount = 18
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Resource Governor Enabled'
			$WorksheetData[0,$Col++] = 'Reconfigure Pending'
			$WorksheetData[0,$Col++] = 'Classifier Function'
			$WorksheetData[0,$Col++] = 'Resource Pool Name'
			$WorksheetData[0,$Col++] = 'Min CPU %'
			$WorksheetData[0,$Col++] = 'Max CPU %'
			$WorksheetData[0,$Col++] = 'Min Memory %'
			$WorksheetData[0,$Col++] = 'Max Memory %'
			$WorksheetData[0,$Col++] = 'Resource Pool Is System Object'
			$WorksheetData[0,$Col++] = 'Workload Group Name'
			$WorksheetData[0,$Col++] = 'Importance'
			$WorksheetData[0,$Col++] = 'Max Requests'
			$WorksheetData[0,$Col++] = 'CPU Time (sec)'
			$WorksheetData[0,$Col++] = 'Memory Grant %'
			$WorksheetData[0,$Col++] = 'Grant Timeout (sec)'
			$WorksheetData[0,$Col++] = 'Degree of Parallelism'
			$WorksheetData[0,$Col++] = 'Workload Group Is System Object'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name
				$ResourceGovernor = $_.Server.Management.ResourceGovernor

				if (($ResourceGovernor.ResourcePools | Where-Object { $_.ID } | Measure-Object).Count -eq 0) {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $ResourceGovernor.Enabled
					$WorksheetData[$Row,$Col++] = $ResourceGovernor.ReconfigurePending
					$WorksheetData[$Row,$Col++] = $ResourceGovernor.ClassifierFunction
					$Row++
				} else {
					$ResourceGovernor.ResourcePools | Where-Object { $_.ID } | ForEach-Object {
						$ResourcePool = $_

						if (($ResourcePool.WorkloadGroups | Where-Object { $_.ID } | Measure-Object).Count -eq 0) {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $ResourceGovernor.Enabled
							$WorksheetData[$Row,$Col++] = $ResourceGovernor.ReconfigurePending
							$WorksheetData[$Row,$Col++] = $ResourceGovernor.ClassifierFunction
							$WorksheetData[$Row,$Col++] = $ResourcePool.Name
							$WorksheetData[$Row,$Col++] = $ResourcePool.MinimumCpuPercentage
							$WorksheetData[$Row,$Col++] = $ResourcePool.MaximumCpuPercentage
							$WorksheetData[$Row,$Col++] = $ResourcePool.MinimumMemoryPercentage
							$WorksheetData[$Row,$Col++] = $ResourcePool.MaximumMemoryPercentage
							$WorksheetData[$Row,$Col++] = $ResourcePool.IsSystemObject
							$Row++
						} else {
							$ResourcePool.WorkloadGroups | Where-Object { $_.ID } | ForEach-Object {
								$Col = 0
								$WorksheetData[$Row,$Col++] = $ServerName
								$WorksheetData[$Row,$Col++] = $ResourceGovernor.Enabled
								$WorksheetData[$Row,$Col++] = $ResourceGovernor.ReconfigurePending
								$WorksheetData[$Row,$Col++] = $ResourceGovernor.ClassifierFunction
								$WorksheetData[$Row,$Col++] = $ResourcePool.Name
								$WorksheetData[$Row,$Col++] = $ResourcePool.MinimumCpuPercentage
								$WorksheetData[$Row,$Col++] = $ResourcePool.MaximumCpuPercentage
								$WorksheetData[$Row,$Col++] = $ResourcePool.MinimumMemoryPercentage
								$WorksheetData[$Row,$Col++] = $ResourcePool.MaximumMemoryPercentage
								$WorksheetData[$Row,$Col++] = $ResourcePool.IsSystemObject
								$WorksheetData[$Row,$Col++] = $_.Name
								$WorksheetData[$Row,$Col++] = $_.Importance
								$WorksheetData[$Row,$Col++] = $_.GroupMaximumRequests
								$WorksheetData[$Row,$Col++] = $_.RequestMaximumCpuTimeSeconds
								$WorksheetData[$Row,$Col++] = $_.RequestMaximumMemoryGrantPercentage
								$WorksheetData[$Row,$Col++] = $_.RequestMemoryGrantTimeoutSeconds
								$WorksheetData[$Row,$Col++] = $_.MaximumDegreeOfParallelism
								$WorksheetData[$Row,$Col++] = $_.IsSystemObject
								$Row++
							}
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType,S $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(5), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(11), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'B2'
					ColumnFormat = @(
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 14; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 16; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 17; NumberFormat = $XlNumFmtNumberS0} 
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 24: Management - SQL Trace Config
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Management - SQL Trace"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Management - SQL Trace'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ManagementTabColor #$ServerTabColor

			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Management.SQLTrace | Where-Object { $_.ID } } ) | Measure-Object).Count + 1
			$ColumnCount = 18
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Trace ID'
			$WorksheetData[0,$Col++] = 'Status'
			$WorksheetData[0,$Col++] = 'Path'
			$WorksheetData[0,$Col++] = 'Max Size (Mb)'
			$WorksheetData[0,$Col++] = 'Stop Time'
			$WorksheetData[0,$Col++] = 'Max File Count'
			$WorksheetData[0,$Col++] = 'Rowset Trace'
			$WorksheetData[0,$Col++] = 'Rollover Enabled'
			$WorksheetData[0,$Col++] = 'Shutdown Enabled'
			$WorksheetData[0,$Col++] = 'Default Trace'
			$WorksheetData[0,$Col++] = 'Buffer Count'
			$WorksheetData[0,$Col++] = 'Buffer Size (Kb)'
			$WorksheetData[0,$Col++] = 'File Position'
			$WorksheetData[0,$Col++] = 'Start Time'
			$WorksheetData[0,$Col++] = 'Last Event Time'
			$WorksheetData[0,$Col++] = 'Event Count'
			$WorksheetData[0,$Col++] = 'Dropped Event Count'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Management.SQLTrace | Where-Object { $_.ID } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.ID
					$WorksheetData[$Row,$Col++] = $_.Status
					$WorksheetData[$Row,$Col++] = $_.Path
					$WorksheetData[$Row,$Col++] = $_.MaxSizeMb
					$WorksheetData[$Row,$Col++] = $_.StopTime
					$WorksheetData[$Row,$Col++] = $_.MaxFileCount
					$WorksheetData[$Row,$Col++] = $_.IsRowsetTrace
					$WorksheetData[$Row,$Col++] = $_.RolloverEnabled
					$WorksheetData[$Row,$Col++] = $_.ShutdownEnabled
					$WorksheetData[$Row,$Col++] = $_.IsDefaultTrace
					$WorksheetData[$Row,$Col++] = $_.BufferCount
					$WorksheetData[$Row,$Col++] = $_.BufferSizeKb
					$WorksheetData[$Row,$Col++] = $_.FilePosition
					$WorksheetData[$Row,$Col++] = $_.StartTime
					$WorksheetData[$Row,$Col++] = $_.LastEventTime
					$WorksheetData[$Row,$Col++] = $_.EventCount
					$WorksheetData[$Row,$Col++] = $_.DroppedEventCount
					$Row++
				}

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 2; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate}
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 12; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 16; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 17; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 18; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 25: Server Objects - SQL Trace Events
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Management - SQL Trace Events"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Management - SQL Trace Events'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ManagementTabColor #$ServerTabColor

			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Management.SQLTrace | ForEach-Object { $_.Events } | Where-Object { $_.CategoryID } } ) | Measure-Object).Count + 1
			$ColumnCount = 6
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Trace ID'
			$WorksheetData[0,$Col++] = 'Category'
			$WorksheetData[0,$Col++] = 'Type'
			$WorksheetData[0,$Col++] = 'Event'
			$WorksheetData[0,$Col++] = 'Column'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Management.SQLTrace | Where-Object { $_.ID } | Sort-Object -Property $_.ID | ForEach-Object {

					$TraceID = $_.ID

					$_.Events | Where-Object { $_.CategoryID } | Sort-Object -Property CategoryID, EventID, ColumnID | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $TraceID
						$WorksheetData[$Row,$Col++] = $_.Category
						$WorksheetData[$Row,$Col++] = $_.Type
						$WorksheetData[$Row,$Col++] = $_.Event
						$WorksheetData[$Row,$Col++] = $_.Column
						$Row++
					}
				}

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 2; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 26: Server Objects - SQL Trace Filters
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Management - SQL Trace Filters"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Management - SQL Trace Filters'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ManagementTabColor #$ServerTabColor

			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Management.SQLTrace | ForEach-Object { $_.Filters } | Where-Object { $_.ColumnID } } ) | Measure-Object).Count + 1
			$ColumnCount = 6
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Trace ID'
			$WorksheetData[0,$Col++] = 'Logical Operator'
			$WorksheetData[0,$Col++] = 'Column'
			$WorksheetData[0,$Col++] = 'Criteria'
			$WorksheetData[0,$Col++] = 'Value'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Management.SQLTrace | Where-Object { $_.ID } | Sort-Object -Property $_.ID | ForEach-Object {

					$TraceID = $_.ID

					$_.Filters | Where-Object { $_.ColumnID } | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $TraceID
						$WorksheetData[$Row,$Col++] = $_.LogicalOperator
						$WorksheetData[$Row,$Col++] = $_.Column
						$WorksheetData[$Row,$Col++] = $_.ComparisonOperator
						$WorksheetData[$Row,$Col++] = $_.FilterValue
						$Row++
					}
				}

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(4), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 2; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 27: Server Objects - Trace Flags
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Management - Trace Flags"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Management - Trace Flags'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ManagementTabColor #$ServerTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Management.TraceFlags } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Management.TraceFlags | Where-Object { $_.TraceFlag } } ) | Measure-Object).Count + 1
			$ColumnCount = 5
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Trace Flag'
			$WorksheetData[0,$Col++] = 'Status'
			$WorksheetData[0,$Col++] = 'Global'
			$WorksheetData[0,$Col++] = 'Session'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Management.TraceFlags | Where-Object { $_.TraceFlag } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.TraceFlag
					$WorksheetData[$Row,$Col++] = $_.Status
					$WorksheetData[$Row,$Col++] = $_.IsGlobal
					$WorksheetData[$Row,$Col++] = $_.IsSession
					$Row++
				}

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

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


			# Worksheet 28: Management - Database Mail Accounts
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Management - Database Mail Accounts"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Management - DB Mail Accounts'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ManagementTabColor #$ServerTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Management.DatabaseMail.Accounts } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Management.DatabaseMail.Accounts | Where-Object { $_.ID } } ) | Measure-Object).Count + 1
			$ColumnCount = 12
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Account Name'
			$WorksheetData[0,$Col++] = 'Description'
			$WorksheetData[0,$Col++] = 'E-mail Address'
			$WorksheetData[0,$Col++] = 'Display Name'
			$WorksheetData[0,$Col++] = 'Reply E-mail'
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Port Number'
			$WorksheetData[0,$Col++] = 'SSL Required'
			$WorksheetData[0,$Col++] = 'Authentication Type'
			$WorksheetData[0,$Col++] = 'User Name'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Management.DatabaseMail.Accounts | Where-Object { $_.ID } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.AccountName
					$WorksheetData[$Row,$Col++] = $_.Description
					$WorksheetData[$Row,$Col++] = $_.OutgoingSmtpServer.EmailAddress
					$WorksheetData[$Row,$Col++] = $_.OutgoingSmtpServer.DisplayName
					$WorksheetData[$Row,$Col++] = $_.OutgoingSmtpServer.ReplyToAddress
					$WorksheetData[$Row,$Col++] = $_.OutgoingSmtpServer.ServerName
					$WorksheetData[$Row,$Col++] = $_.OutgoingSmtpServer.PortNumber
					$WorksheetData[$Row,$Col++] = $_.OutgoingSmtpServer.SslConnectionRequired
					$WorksheetData[$Row,$Col++] = $_.SmtpAuthentication.AuthenticationType
					$WorksheetData[$Row,$Col++] = $_.SmtpAuthentication.UserName
					$Row++
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType,S $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 29: Database Mail Profiles
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Management - Database Mail Profiles"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Management - DB Mail Profiles'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ManagementTabColor #$ServerTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Management.DatabaseMail.Profiles } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Management.DatabaseMail.Profiles | Where-Object { $_.ID } } ) | Measure-Object).Count + 1
			$ColumnCount = 4
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Profile Name'
			$WorksheetData[0,$Col++] = 'Description'
			$WorksheetData[0,$Col++] = 'Accounts'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Management.DatabaseMail.Profiles | Where-Object { $_.ID } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.ProfileName
					$WorksheetData[$Row,$Col++] = $_.Description
					$WorksheetData[$Row,$Col++] = ($_.Accounts | Sort-Object) -join $Delimiter # PoSH is more forgiving here than [String]::Join if $_.Accounts is $NULL
					$Row++
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType,S $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

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


			# Worksheet 30: Database Mail Security
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Management - Database Mail Security"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Management - DB Mail Security'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ManagementTabColor #$ServerTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Management.DatabaseMail.ProfileSecurity } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Management.DatabaseMail.ProfileSecurity | Where-Object { $_.ProfileId } } ) | Measure-Object).Count + 1
			$ColumnCount = 5
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Profile Name'
			$WorksheetData[0,$Col++] = 'Principal Name'
			$WorksheetData[0,$Col++] = 'Is Default'
			$WorksheetData[0,$Col++] = 'Is Public'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Management.DatabaseMail.ProfileSecurity | Where-Object { $_.ProfileId } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.ProfileName
					$WorksheetData[$Row,$Col++] = $_.PrincipalName
					$WorksheetData[$Row,$Col++] = $_.IsDefault
					$WorksheetData[$Row,$Col++] = $_.IsPublic
					$Row++
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType,S $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
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


			# Worksheet 31: Database Mail Configuration
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Management - Database Mail Configuration"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Management - DB Mail Config'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ManagementTabColor #$ServerTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Management.DatabaseMail.ConfigurationValues } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Management.DatabaseMail.ConfigurationValues | Where-Object { $_.AccountRetryAttempts } } ) | Measure-Object).Count + 1
			$ColumnCount = 9
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Mail XPs Enabled'
			$WorksheetData[0,$Col++] = 'Account Retry Attempts'
			$WorksheetData[0,$Col++] = 'Account Retry Delay (seconds)'
			$WorksheetData[0,$Col++] = 'Maximum File Size (Bytes)'
			$WorksheetData[0,$Col++] = 'Prohibited Attachment File Extensions'
			$WorksheetData[0,$Col++] = 'Database Mail Executable Minimum Lifetime (seconds)'
			$WorksheetData[0,$Col++] = 'Logging Level'
			$WorksheetData[0,$Col++] = 'Default Attachment Encoding'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name
				$DatabaseMailXpsEnabled = $_.Server.Configuration.Advanced.Miscellaneous.DatabaseMailXPsEnabled.RunningValue

				$_.Server.Management.DatabaseMail.ConfigurationValues | Where-Object { $_.AccountRetryAttempts } | ForEach-Object {
					#$_.Server.Management.DatabaseMail.ConfigurationValues | Where-Object { ($_ | Measure-Object).Count -gt 0 } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $DatabaseMailXpsEnabled
					$WorksheetData[$Row,$Col++] = $_.AccountRetryAttempts
					$WorksheetData[$Row,$Col++] = $_.AccountRetryDelaySeconds
					$WorksheetData[$Row,$Col++] = $_.MaxFileSizeBytes
					$WorksheetData[$Row,$Col++] = $_.ProhibitedExtensions
					$WorksheetData[$Row,$Col++] = $_.DatabaseMailExeMinimumLifeTimeSeconds
					$WorksheetData[$Row,$Col++] = $_.LoggingLevel
					$WorksheetData[$Row,$Col++] = $_.DefaultAttachmentEncoding
					$Row++
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType,S $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'B2'
					ColumnFormat = @(
						@{ColumnNumber = 3; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 32: Database Overview
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Overview"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Database Overview'
			#$Worksheet.Tab.Color = $OverviewTabColor
			$Worksheet.Tab.ThemeColor = $OverviewTabColor

			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases }) | Measure-Object).Count + 1
			$ColumnCount = 17
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Product Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Date Created'
			$WorksheetData[0,$Col++] = 'Status'
			$WorksheetData[0,$Col++] = 'Owner'
			$WorksheetData[0,$Col++] = 'Compatibility Level'
			$WorksheetData[0,$Col++] = 'Recovery Model'
			$WorksheetData[0,$Col++] = 'Data File Count'
			$WorksheetData[0,$Col++] = 'Log File Count'
			$WorksheetData[0,$Col++] = 'Data File Size (MB)'
			$WorksheetData[0,$Col++] = 'Log File Size (MB)'
			$WorksheetData[0,$Col++] = 'Available Data Space (MB)'
			$WorksheetData[0,$Col++] = 'Last Known DBCC Date'
			$WorksheetData[0,$Col++] = 'Last Full Backup'
			$WorksheetData[0,$Col++] = 'Last Diff Backup'
			$WorksheetData[0,$Col++] = 'Last Log Backup'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name
				$ProductName = $_.Server.Configuration.General.Product

				$_.Server.Databases | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $ProductName
					$WorksheetData[$Row,$Col++] = $_.Name
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.DateCreated
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.Status
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.Owner
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.CompatibilityLevel
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.RecoveryModel
					$WorksheetData[$Row,$Col++] = ($_.Properties.Files.DatabaseFiles | Where-Object { $_.FileType -ieq 'rows data' } | Measure-Object).Count
					$WorksheetData[$Row,$Col++] = ($_.Properties.Files.DatabaseFiles | Where-Object { $_.FileType -ieq 'Log' } | Measure-Object).Count
					$WorksheetData[$Row,$Col++] = "{0:N2}" -f (($_.Properties.Files.DatabaseFiles | Where-Object { $_.FileType -ieq 'rows data' } | Measure-Object -Property SizeKB -Sum).Sum / 1KB)
					$WorksheetData[$Row,$Col++] = "{0:N2}" -f (($_.Properties.Files.DatabaseFiles | Where-Object { $_.FileType -ieq 'log' } | Measure-Object -Property SizeKB -Sum).Sum / 1KB)
					$WorksheetData[$Row,$Col++] = "{0:N2}" -f (($_.Properties.Files.DatabaseFiles | Where-Object { $_.FileType -ieq 'rows data' } | Measure-Object -Property AvailableSpaceKB -Sum).Sum / 1KB)
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.LastKnownGoodDbccDate
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Backup.LastFullBackupDate
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Backup.LastDifferentialBackupDate
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Backup.LastLogBackupDate
					$Row++
				}

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType,S $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'D2'
					ColumnFormat = @(
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 11; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 12; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 14; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 16; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 17; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion 


			# 			# Worksheet 33: Database Configuration (VERTICAL FORMAT)
			# 			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Configuration"
			# 			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			# 			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			# 			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			# 			$Worksheet.Name = 'DB Config'
			# 			#$Worksheet.Tab.Color = $DatabaseTabColor
			# 			$Worksheet.Tab.ThemeColor = $DatabaseTabColor
			# 
			# 			#$RowCount = ($SqlServerInventory.DatabaseServer | Measure-Object).Count + 1
			# 			#$ColumnCount = 15
			# 			$ColumnCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases }) | Measure-Object).Count + 1
			# 			$RowCount = 101
			# 			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount
			# 
			# 			$Row = 0
			# 			$WorksheetData[$Row++,0] = 'Server Name'
			# 			$WorksheetData[$Row++,0] = 'Database Name'
			# 			# 2
			# 			$WorksheetData[$Row++,0] = 'General'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Backup'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Last Full Backup Date'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Last Differential Backup Date'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Last Log Backup Date'
			# 			# 7
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Database'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Database Snapshot Base Name'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Date Created'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Last Known Good DBCC Date'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Is AlwaysOn AG Member'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Is Database Snapshot'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Is Database Snapshot Base'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Full-Text Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Is Mail Host'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Is Management Data Warehouse'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Mirroring Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Is System Object'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Number Of Users'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Owner'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Size (MB)'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Space Available (KB)'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Status'
			# 			# 24
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Maintenance'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Collation'
			# 			# 26
			# 			$WorksheetData[$Row++,0] = 'Options'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Collation'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Recovery Model'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Compatibility Level'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Containment Type'
			# 			# 31
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Other Options'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Automatic' 
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Auto Close'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Auto Create Statistics'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Auto Shrink'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Auto Update Statistics'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Auto UpdateStatistics Async'
			# 			# 38
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Containment'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Default Full-Text Language'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Default Language'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Nested Triggers Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Transform Noise Words'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Two Digit Year Cutoff'
			# 			# 44
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Cursor'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Close Cursor on Commit Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Local Cursor Default'
			# 			# 47
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'FILESTREAM'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'FILESTREAM DirectoryName'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'FILESTREAM Non-Transacted Access'
			# 			# 50
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Miscellaneous'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Allow Snapshot Isolation'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'ANSI NULL Default'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'ANSI NULLS Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'ANSI Padding Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'ANSI Warnings Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Arithmetic Abort Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Concatenate Null Yields Null'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Cross-Database Ownership Chaining Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Date Correlation Optimization Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Is Read Committed Snapshot On'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Numeric Round-Abort Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Parameterization'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Quoted Identifiers Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Recursive Triggers Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Trustworthy'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'VarDecimal Storage Format Enabled'
			# 			# 67
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Recovery'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Page Verify'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Target Recovery Time (sec)'
			# 			# 70
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Service Broker'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Broker Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Honor Broker Priority'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Service Broker Identifier'
			# 			# 74
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'State'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Database Read-Only'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Encryption Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString___3 + 'Restrict Access'
			# 			# 78
			# 			$WorksheetData[$Row++,0] = 'AlwaysOn'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Availability Database Synchronization State'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Availability Group Name'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'IsAvailability Group Member'
			# 			# 82
			# 			$WorksheetData[$Row++,0] = 'Change Tracking'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Retention Period'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Retention Period Units'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Auto Cleanup'
			# 			# 87
			# 			$WorksheetData[$Row++,0] = 'Mirroring'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Failover Log Sequence Number'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Mirroring ID'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Mirroring Partner'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Mirroring Partner Instance'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Redo Queue Max Size (KB)'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Role Sequence'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Safety Level'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Safety Sequence'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Status'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Timeout (sec)'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Mirroring Witness'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Mirroring Witness Status'
			# 			# 101
			# 
			# 
			# 			$Col = 1
			# 			$SqlServerInventory.DatabaseServer | Sort-Object -Property $_.ServerName | ForEach-Object {
			# 
			# 				$ServerName = $_.Server.Configuration.General.Name
			# 
			# 				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
			# 
			# 					$Row = 0
			# 
			# 					$WorksheetData[$Row++,$Col] = $ServerName
			# 					$WorksheetData[$Row++,$Col] = $_.Name
			# 					$WorksheetData[$Row++,$Col] = $null
			# 
			# 					# General - Backup
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Backup.LastFullBackupDate
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Backup.LastDifferentialBackupDate
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Backup.LastLogBackupDate
			# 
			# 					# General - Database
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Database.DatabaseSnapshotBaseName
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Database.DateCreated
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Database.LastKnownGoodDbccDate
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Database.IsAvailabilityGroupMember
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Database.IsDatabaseSnapshot
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Database.IsDatabaseSnapshotBase
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Database.IsFullTextEnabled
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Database.IsMailHost
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Database.IsManagementDataWarehouse
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Database.IsMirroringEnabled
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Database.IsSystemObject
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Database.NumberOfUsers
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Database.Owner
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Database.SizeMB
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Database.SpaceAvailableKB
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Database.Status
			# 
			# 					# General - Maintenance
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.General.Maintenance.Collation
			# 
			# 					# Options
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.Collation
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.RecoveryModel
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.CompatibilityLevel
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.ContainmentType
			# 
			# 					# Options - Other Options
			# 					$WorksheetData[$Row++,$Col] = $null
			# 
			# 					# Options - Other Options - Automatic
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Automatic.AutoClose
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Automatic.AutoCreateStatistics
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Automatic.AutoShrink
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Automatic.AutoUpdateStatistics
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Automatic.AutoUpdateStatisticsAsync
			# 
			# 					# Options - Other Options - Containment
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Containment.DefaultFullTextLanguage
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Containment.DefaultLanguage
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Containment.NestedTriggersEnabled
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Containment.TransformNoiseWords
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Containment.TwoDigitYearCutoff
			# 
			# 
			# 					# Options - Other Options - Cursor
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Cursor.CloseCursorsOnCommitEnabled
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Cursor.LocalCursorsDefault
			# 
			# 					# Options - Other Options - FILESTREAM
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Filestream.FilestreamDirectoryName
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Filestream.FilestreamNonTransactedAccess
			# 
			# 					# Options - Other Options - Miscellaneous
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Miscellaneous.SnapshotIsolation
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Miscellaneous.AnsiNullDefault
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Miscellaneous.AnsiNullsEnabled
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Miscellaneous.AnsiPaddingEnabled
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Miscellaneous.AnsiWarningsEnabled
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Miscellaneous.ArithmeticAbortEnabled
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Miscellaneous.ConcatenateNullYieldsNull
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Miscellaneous.DatabaseOwnershipChaining
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Miscellaneous.DateCorrelationOptimization
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Miscellaneous.IsReadCommittedSnapshotOn
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Miscellaneous.NumericRoundAbortEnabled
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Miscellaneous.Parameterization
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Miscellaneous.QuotedIdentifiersEnabled
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Miscellaneous.RecursiveTriggersEnabled
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Miscellaneous.Trustworthy
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Miscellaneous.VarDecimalStorageFormatEnabled
			# 
			# 					# Options - Other Options - Recovery
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Recovery.PageVerify
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.Recovery.TargetRecoveryTimeSeconds
			# 
			# 					# Options - Other Options - ServiceBroker
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.ServiceBroker.BrokerEnabled
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.ServiceBroker.HonorBrokerPriority
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.ServiceBroker.ServiceBrokerIdentifier
			# 
			# 					# Options - Other Options - State
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.State.DatabaseReadOnly
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.State.EncryptionEnabled
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Options.OtherOptions.State.RestrictAccess
			# 
			# 					# AlwaysOn
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.AlwaysOn.AvailabilityDatabaseSynchronizationState
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.AlwaysOn.AvailabilityGroupName
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.AlwaysOn.IsAvailabilityGroupMember
			# 
			# 					# Change Tracking
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.ChangeTracking.IsEnabled
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.ChangeTracking.RetentionPeriod
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.ChangeTracking.RetentionPeriodUnits 
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.ChangeTracking.AutoCleanUp
			# 
			# 					# Mirroring
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Mirroring.IsEnabled
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Mirroring.MirroringFailoverLogSequenceNumber
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Mirroring.MirroringID
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Mirroring.MirroringPartner
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Mirroring.MirroringPartnerInstance
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Mirroring.MirroringRedoQueueMaxSizeKB
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Mirroring.MirroringRoleSequence
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Mirroring.MirroringSafetyLevel
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Mirroring.MirroringSafetySequence
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Mirroring.MirroringStatus
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Mirroring.MirroringTimeoutSeconds
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Mirroring.MirroringWitness
			# 					$WorksheetData[$Row++,$Col] = $_.Properties.Mirroring.MirroringWitnessStatus
			# 
			# 					$Col++
			# 
			# 				}
			# 			}
			# 			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			# 			$Range.Value2 = $WorksheetData
			# 			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			# 			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			# 			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null
			# 
			# 			$WorksheetFormat.Add($WorksheetNumber, @{
			# 					BoldFirstRow = $false
			# 					BoldFirstColumn = $true
			# 					AutoFilter = $false
			# 					FreezeAtCell = 'B2'
			# 					ColumnFormat = @()
			# 					RowFormat = @(
			# 						@{RowNumber = 5; NumberFormat = $XlNumFmtDate},
			# 						@{RowNumber = 6; NumberFormat = $XlNumFmtDate},
			# 						@{RowNumber = 7; NumberFormat = $XlNumFmtDate},
			# 						@{RowNumber = 10; NumberFormat = $XlNumFmtDate},
			# 						@{RowNumber = 11; NumberFormat = $XlNumFmtDate},
			# 						@{RowNumber = 20; NumberFormat = $XlNumFmtNumberS2},
			# 						@{RowNumber = 22; NumberFormat = $XlNumFmtNumberS2},
			# 						@{RowNumber = 23; NumberFormat = $XlNumFmtNumberS2},
			# 						@{RowNumber = 44; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 70; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 85; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 90; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 94; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 95; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 97; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 99; NumberFormat = $XlNumFmtNumberS0}
			# 					)
			# 				})
			# 
			# 			$WorksheetNumber++
			#endregion


			# Worksheet 33: Database Configuration - General
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Configuration - General"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'DB Config - General'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases }) | Measure-Object).Count + 1
			$ColumnCount = 24
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			# 2
			$WorksheetData[0,$Col++] = 'Backup >>'
			$WorksheetData[0,$Col++] = 'Last Full Backup Date'
			$WorksheetData[0,$Col++] = 'Last Differential Backup Date'
			$WorksheetData[0,$Col++] = 'Last Log Backup Date'
			# 6
			$WorksheetData[0,$Col++] = 'Database >>'
			$WorksheetData[0,$Col++] = 'Database Snapshot Base Name'
			$WorksheetData[0,$Col++] = 'Date Created'
			$WorksheetData[0,$Col++] = 'Is AlwaysOn AG Member'
			$WorksheetData[0,$Col++] = 'Is Database Snapshot'
			$WorksheetData[0,$Col++] = 'Is Database Snapshot Base'
			$WorksheetData[0,$Col++] = 'Full-Text Enabled'
			$WorksheetData[0,$Col++] = 'Is Mail Host'
			$WorksheetData[0,$Col++] = 'Is Management Data Warehouse'
			$WorksheetData[0,$Col++] = 'Mirroring Enabled'
			$WorksheetData[0,$Col++] = 'Is System Object'
			$WorksheetData[0,$Col++] = 'Number Of Users'
			$WorksheetData[0,$Col++] = 'Owner'
			$WorksheetData[0,$Col++] = 'Size (MB)'
			$WorksheetData[0,$Col++] = 'Space Available (KB)'
			$WorksheetData[0,$Col++] = 'Status'
			# 22
			$WorksheetData[0,$Col++] = 'Maintenance >>'
			$WorksheetData[0,$Col++] = 'Collation'
			# 24

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | ForEach-Object {

					$Col = 0

					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.Name

					# General - Backup
					$WorksheetData[$Row,$Col++] = $null
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Backup.LastFullBackupDate
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Backup.LastDifferentialBackupDate
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Backup.LastLogBackupDate

					# General - Database
					$WorksheetData[$Row,$Col++] = $null
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.DatabaseSnapshotBaseName
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.DateCreated
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.IsAvailabilityGroupMember
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.IsDatabaseSnapshot
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.IsDatabaseSnapshotBase
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.IsFullTextEnabled
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.IsMailHost
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.IsManagementDataWarehouse
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.IsMirroringEnabled
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.IsSystemObject
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.NumberOfUsers
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.Owner
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.SizeMB
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.SpaceAvailableKB
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.Status

					# General - Maintenance
					$WorksheetData[$Row,$Col++] = $null
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Maintenance.Collation

					$Row++

				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 18; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 20; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 21; NumberFormat = $XlNumFmtNumberS2}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 34: Database Configuration - Database Files
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Configuration - Files"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'DB Config - Files'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Properties.Files.DatabaseFiles } } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Properties.Files.DatabaseFiles | Where-Object { $_.ID } } } ) | Measure-Object).Count + 1

			$ColumnCount = 20
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Logical Name'
			$WorksheetData[0,$Col++] = 'Type'
			$WorksheetData[0,$Col++] = 'FileGroup'
			$WorksheetData[0,$Col++] = 'Current Size (MB)'
			$WorksheetData[0,$Col++] = 'Max Size (MB)'
			$WorksheetData[0,$Col++] = 'Used Space (MB)'
			$WorksheetData[0,$Col++] = 'Available Space (MB)'
			$WorksheetData[0,$Col++] = 'Is Primary File'
			$WorksheetData[0,$Col++] = 'Is Offline'
			$WorksheetData[0,$Col++] = 'Is Read Only'
			$WorksheetData[0,$Col++] = 'Is Media Read Only'
			$WorksheetData[0,$Col++] = 'Is Sparse'
			$WorksheetData[0,$Col++] = 'Growth'
			$WorksheetData[0,$Col++] = 'Growth Type'
			$WorksheetData[0,$Col++] = 'Volume Free Space (MB)'
			$WorksheetData[0,$Col++] = 'VLF Count'
			$WorksheetData[0,$Col++] = 'Path'
			$WorksheetData[0,$Col++] = 'Filename'


			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | ForEach-Object {

					$DatabaseName = $_.Name

					$_.Properties.Files.DatabaseFiles | Where-Object { $_.ID } | ForEach-Object {
						#$_.Properties.Files.DatabaseFiles | Where-Object { ($_ | Measure-Object).Count -gt 0 } | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.LogicalName
						$WorksheetData[$Row,$Col++] = $_.FileType
						$WorksheetData[$Row,$Col++] = $_.Filegroup
						$WorksheetData[$Row,$Col++] = "{0:N2}" -f ($_.SizeKB / 1KB)
						$WorksheetData[$Row,$Col++] = "{0:N2}" -f ($_.MaxSizeKB / 1KB)
						$WorksheetData[$Row,$Col++] = "{0:N2}" -f ($_.UsedSpaceKB / 1KB)
						$WorksheetData[$Row,$Col++] = if ($_.AvailableSpaceKB) { "{0:N2}" -f ($_.AvailableSpaceKB / 1KB) } else { $null }
						$WorksheetData[$Row,$Col++] = $_.IsPrimaryFile
						$WorksheetData[$Row,$Col++] = $_.IsOffline
						$WorksheetData[$Row,$Col++] = $_.IsReadOnly
						$WorksheetData[$Row,$Col++] = $_.IsReadOnlyMedia
						$WorksheetData[$Row,$Col++] = $_.IsSparse
						$WorksheetData[$Row,$Col++] = $_.Growth
						$WorksheetData[$Row,$Col++] = $_.GrowthType
						$WorksheetData[$Row,$Col++] = "{0:N2}" -f ($_.VolumeFreeSpaceBytes / 1MB)
						$WorksheetData[$Row,$Col++] = $_.VlfCount
						$WorksheetData[$Row,$Col++] = $_.Path
						$WorksheetData[$Row,$Col++] = $_.FileName
						$Row++ 
					}
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
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 17; NumberFormat = $XlNumFmtNumberS2}
						@{ColumnNumber = 18; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion 


			# Worksheet 35: Database Configuration - Filegroups
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Configuration - Filegroups"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'DB Config - Filegroups'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Properties.FileGroups.Rows } } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Properties.FileGroups.Rows | Where-Object { $_.ID } } } ) | Measure-Object).Count + 1

			$ColumnCount = 7
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Filegroup Name'
			$WorksheetData[0,$Col++] = 'Type'
			$WorksheetData[0,$Col++] = 'Files'
			$WorksheetData[0,$Col++] = 'Read-Only'
			$WorksheetData[0,$Col++] = 'Is Default'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | ForEach-Object {

					$DatabaseName = $_.Name

					$_.Properties.FileGroups.Rows | Where-Object { $_.ID } | ForEach-Object {
						#$_.Properties.FileGroups.Rows | Where-Object { ($_ | Measure-Object).Count -gt 0 } | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = 'Rows'
						$WorksheetData[$Row,$Col++] = $_.Files
						$WorksheetData[$Row,$Col++] = $_.ReadOnly
						$WorksheetData[$Row,$Col++] = $_.IsDefault 
						$Row++ 
					}

					# 					$_.Properties.FileGroups.Filestream | ForEach-Object {
					# 						$Col = 0
					# 						$WorksheetData[$Row,$Col++] = $ServerName
					# 						$WorksheetData[$Row,$Col++] = $DatabaseName
					# 						$WorksheetData[$Row,$Col++] = 'FILESTREAM'
					# 						$WorksheetData[$Row,$Col++] = $_.Name
					# 						$WorksheetData[$Row,$Col++] = $_.Files
					# 						$WorksheetData[$Row,$Col++] = $_.ReadOnly
					# 						$WorksheetData[$Row,$Col++] = $_.IsDefault 
					# 						$Row++ 
					# 					}

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
					ColumnFormat = @(
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 36: Database Configuration - Options
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Configuration - Options"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'DB Config - Options'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases }) | Measure-Object).Count + 1
			$ColumnCount = 53
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			# 2
			$WorksheetData[0,$Col++] = 'Collation'
			$WorksheetData[0,$Col++] = 'Recovery Model'
			$WorksheetData[0,$Col++] = 'Compatibility Level'
			$WorksheetData[0,$Col++] = 'Containment Type'
			# 6
			$WorksheetData[0,$Col++] = 'Other Options >>'
			$WorksheetData[0,$Col++] = 'Automatic >>>' 
			$WorksheetData[0,$Col++] = 'Auto Close'
			$WorksheetData[0,$Col++] = 'Auto Create Statistics'
			$WorksheetData[0,$Col++] = 'Auto Shrink'
			$WorksheetData[0,$Col++] = 'Auto Update Statistics'
			$WorksheetData[0,$Col++] = 'Auto UpdateStatistics Async'
			# 13
			$WorksheetData[0,$Col++] = 'Containment >>>'
			$WorksheetData[0,$Col++] = 'Default Full-Text Language'
			$WorksheetData[0,$Col++] = 'Default Language'
			$WorksheetData[0,$Col++] = 'Nested Triggers Enabled'
			$WorksheetData[0,$Col++] = 'Transform Noise Words'
			$WorksheetData[0,$Col++] = 'Two Digit Year Cutoff'
			# 19
			$WorksheetData[0,$Col++] = 'Cursor >>>'
			$WorksheetData[0,$Col++] = 'Close Cursor on Commit Enabled'
			$WorksheetData[0,$Col++] = 'Local Cursor Default'
			# 22
			$WorksheetData[0,$Col++] = 'FILESTREAM >>>'
			$WorksheetData[0,$Col++] = 'FILESTREAM DirectoryName'
			$WorksheetData[0,$Col++] = 'FILESTREAM Non-Transacted Access'
			# 25
			$WorksheetData[0,$Col++] = 'Miscellaneous >>>'
			$WorksheetData[0,$Col++] = 'Allow Snapshot Isolation'
			$WorksheetData[0,$Col++] = 'ANSI NULL Default'
			$WorksheetData[0,$Col++] = 'ANSI NULLS Enabled'
			$WorksheetData[0,$Col++] = 'ANSI Padding Enabled'
			$WorksheetData[0,$Col++] = 'ANSI Warnings Enabled'
			$WorksheetData[0,$Col++] = 'Arithmetic Abort Enabled'
			$WorksheetData[0,$Col++] = 'Concatenate Null Yields Null'
			$WorksheetData[0,$Col++] = 'Cross-Database Ownership Chaining Enabled'
			$WorksheetData[0,$Col++] = 'Date Correlation Optimization Enabled'
			$WorksheetData[0,$Col++] = 'Is Read Committed Snapshot On'
			$WorksheetData[0,$Col++] = 'Numeric Round-Abort Enabled'
			$WorksheetData[0,$Col++] = 'Parameterization'
			$WorksheetData[0,$Col++] = 'Quoted Identifiers Enabled'
			$WorksheetData[0,$Col++] = 'Recursive Triggers Enabled'
			$WorksheetData[0,$Col++] = 'Trustworthy'
			$WorksheetData[0,$Col++] = 'VarDecimal Storage Format Enabled'
			# 42
			$WorksheetData[0,$Col++] = 'Recovery >>>'
			$WorksheetData[0,$Col++] = 'Page Verify'
			$WorksheetData[0,$Col++] = 'Target Recovery Time (sec)'
			# 45
			$WorksheetData[0,$Col++] = 'Service Broker >>>'
			$WorksheetData[0,$Col++] = 'Broker Enabled'
			$WorksheetData[0,$Col++] = 'Honor Broker Priority'
			$WorksheetData[0,$Col++] = 'Service Broker Identifier'
			# 49
			$WorksheetData[0,$Col++] = 'State >>>'
			$WorksheetData[0,$Col++] = 'Database Read-Only'
			$WorksheetData[0,$Col++] = 'Encryption Enabled'
			$WorksheetData[0,$Col++] = 'Restrict Access'
			# 53

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | ForEach-Object {

					$Col = 0

					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.Name

					# Options
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.Collation
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.RecoveryModel
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.CompatibilityLevel
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.ContainmentType

					# Options - Other Options
					$WorksheetData[$Row,$Col++] = $null

					# Options - Other Options - Automatic
					$WorksheetData[$Row,$Col++] = $null
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Automatic.AutoClose
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Automatic.AutoCreateStatistics
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Automatic.AutoShrink
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Automatic.AutoUpdateStatistics
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Automatic.AutoUpdateStatisticsAsync

					# Options - Other Options - Containment
					$WorksheetData[$Row,$Col++] = $null
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Containment.DefaultFullTextLanguage
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Containment.DefaultLanguage
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Containment.NestedTriggersEnabled
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Containment.TransformNoiseWords
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Containment.TwoDigitYearCutoff

					# Options - Other Options - Cursor
					$WorksheetData[$Row,$Col++] = $null
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Cursor.CloseCursorsOnCommitEnabled
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Cursor.LocalCursorsDefault

					# Options - Other Options - FILESTREAM
					$WorksheetData[$Row,$Col++] = $null
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Filestream.FilestreamDirectoryName
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Filestream.FilestreamNonTransactedAccess

					# Options - Other Options - Miscellaneous
					$WorksheetData[$Row,$Col++] = $null
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Miscellaneous.SnapshotIsolation
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Miscellaneous.AnsiNullDefault
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Miscellaneous.AnsiNullsEnabled
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Miscellaneous.AnsiPaddingEnabled
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Miscellaneous.AnsiWarningsEnabled
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Miscellaneous.ArithmeticAbortEnabled
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Miscellaneous.ConcatenateNullYieldsNull
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Miscellaneous.DatabaseOwnershipChaining
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Miscellaneous.DateCorrelationOptimization
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Miscellaneous.IsReadCommittedSnapshotOn
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Miscellaneous.NumericRoundAbortEnabled
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Miscellaneous.Parameterization
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Miscellaneous.QuotedIdentifiersEnabled
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Miscellaneous.RecursiveTriggersEnabled
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Miscellaneous.Trustworthy
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Miscellaneous.VarDecimalStorageFormatEnabled

					# Options - Other Options - Recovery
					$WorksheetData[$Row,$Col++] = $null
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Recovery.PageVerify
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.Recovery.TargetRecoveryTimeSeconds

					# Options - Other Options - ServiceBroker
					$WorksheetData[$Row,$Col++] = $null
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.ServiceBroker.BrokerEnabled
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.ServiceBroker.HonorBrokerPriority
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.ServiceBroker.ServiceBrokerIdentifier

					# Options - Other Options - State
					$WorksheetData[$Row,$Col++] = $null
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.State.DatabaseReadOnly
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.State.EncryptionEnabled
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.OtherOptions.State.RestrictAccess

					$Row++

				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 19; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 45; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 37: Database Configuration - AlwaysOn
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Configuration - AlwaysOn"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'DB Config - AlwaysOn'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases }) | Measure-Object).Count + 1
			$ColumnCount = 6
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'AlwaysOn AG Enabled'
			# 3
			$WorksheetData[0,$Col++] = 'Availability Database Synchronization State'
			$WorksheetData[0,$Col++] = 'Availability Group Name'
			$WorksheetData[0,$Col++] = 'IsAvailability Group Member'
			# 5

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name
				$AlwaysOnAgEnabled = $_.Server.Configuration.General.IsHadrEnabled

				$_.Server.Databases | ForEach-Object {

					$Col = 0

					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.Name
					$WorksheetData[$Row,$Col++] = $AlwaysOnAgEnabled

					# AlwaysOn
					$WorksheetData[$Row,$Col++] = $_.Properties.AlwaysOn.AvailabilityDatabaseSynchronizationState
					$WorksheetData[$Row,$Col++] = $_.Properties.AlwaysOn.AvailabilityGroupName
					$WorksheetData[$Row,$Col++] = $_.Properties.AlwaysOn.IsAvailabilityGroupMember

					$Row++
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

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


			# Worksheet 38: Database Configuration - Change Tracking
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Configuration - Change Tracking"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'DB Config - Change Tracking'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases }) | Measure-Object).Count + 1
			$ColumnCount = 6
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Enabled'
			$WorksheetData[0,$Col++] = 'Retention Period'
			$WorksheetData[0,$Col++] = 'Retention Period Units'
			$WorksheetData[0,$Col++] = 'Auto Cleanup'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.Name
					$WorksheetData[$Row,$Col++] = $_.Properties.ChangeTracking.IsEnabled
					$WorksheetData[$Row,$Col++] = $_.Properties.ChangeTracking.RetentionPeriod
					$WorksheetData[$Row,$Col++] = $_.Properties.ChangeTracking.RetentionPeriodUnits 
					$WorksheetData[$Row,$Col++] = $_.Properties.ChangeTracking.AutoCleanUp
					$Row++
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 39: Database Configuration - Permissions
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Configuration - Permissions"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'DB Config - Permissions'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Properties.Permissions | Where-Object { $_.PermissionType } } } ) | Measure-Object).Count + 1
			$ColumnCount = 12
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Object Type'
			$WorksheetData[0,$Col++] = 'Object Schema'
			$WorksheetData[0,$Col++] = 'Object Name'
			$WorksheetData[0,$Col++] = 'Action'
			$WorksheetData[0,$Col++] = 'Permission'
			$WorksheetData[0,$Col++] = 'Column'
			$WorksheetData[0,$Col++] = 'Granted To'
			$WorksheetData[0,$Col++] = 'Grantee Type'
			$WorksheetData[0,$Col++] = 'Granted By'
			$WorksheetData[0,$Col++] = 'Grantor Type'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.ServerName

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {

					$DatabaseName = $_.Name

					$_.Properties.Permissions | Where-Object { $_.PermissionType } | 
					Sort-Object -Property ObjectClass, ObjectName, PermissionState, PermissionType, Grantee | ForEach-Object {

						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.ObjectClass
						$WorksheetData[$Row,$Col++] = $_.ObjectSchema
						$WorksheetData[$Row,$Col++] = $_.ObjectName
						$WorksheetData[$Row,$Col++] = $_.PermissionState
						$WorksheetData[$Row,$Col++] = $_.PermissionType
						$WorksheetData[$Row,$Col++] = $_.ColumnName
						$WorksheetData[$Row,$Col++] = $_.Grantee
						$WorksheetData[$Row,$Col++] = $_.GranteeType
						$WorksheetData[$Row,$Col++] = $_.Grantor
						$WorksheetData[$Row,$Col++] = $_.GrantorType
						$Row++
					} 
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 40: Database Configuration - Mirroring
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Configuration - Mirroring"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'DB Config - Mirroring'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases }) | Measure-Object).Count + 1
			$ColumnCount = 15
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Enabled'
			$WorksheetData[0,$Col++] = 'Failover Log Sequence Number'
			$WorksheetData[0,$Col++] = 'Mirroring ID'
			$WorksheetData[0,$Col++] = 'Mirroring Partner'
			$WorksheetData[0,$Col++] = 'Mirroring Partner Instance'
			$WorksheetData[0,$Col++] = 'Redo Queue Max Size (KB)'
			$WorksheetData[0,$Col++] = 'Role Sequence'
			$WorksheetData[0,$Col++] = 'Safety Level'
			$WorksheetData[0,$Col++] = 'Safety Sequence'
			$WorksheetData[0,$Col++] = 'Status'
			$WorksheetData[0,$Col++] = 'Timeout (sec)'
			$WorksheetData[0,$Col++] = 'Mirroring Witness'
			$WorksheetData[0,$Col++] = 'Mirroring Witness Status'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.Name
					$WorksheetData[$Row,$Col++] = $_.Properties.Mirroring.IsEnabled
					$WorksheetData[$Row,$Col++] = $_.Properties.Mirroring.MirroringFailoverLogSequenceNumber
					$WorksheetData[$Row,$Col++] = $_.Properties.Mirroring.MirroringID
					$WorksheetData[$Row,$Col++] = $_.Properties.Mirroring.MirroringPartner
					$WorksheetData[$Row,$Col++] = $_.Properties.Mirroring.MirroringPartnerInstance
					$WorksheetData[$Row,$Col++] = $_.Properties.Mirroring.MirroringRedoQueueMaxSizeKB
					$WorksheetData[$Row,$Col++] = $_.Properties.Mirroring.MirroringRoleSequence
					$WorksheetData[$Row,$Col++] = $_.Properties.Mirroring.MirroringSafetyLevel
					$WorksheetData[$Row,$Col++] = $_.Properties.Mirroring.MirroringSafetySequence
					$WorksheetData[$Row,$Col++] = $_.Properties.Mirroring.MirroringStatus
					$WorksheetData[$Row,$Col++] = $_.Properties.Mirroring.MirroringTimeoutSeconds
					$WorksheetData[$Row,$Col++] = $_.Properties.Mirroring.MirroringWitness
					$WorksheetData[$Row,$Col++] = $_.Properties.Mirroring.MirroringWitnessStatus
					$Row++
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 11; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion 


			# Worksheet 41: Database Security - Users
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Security - Users"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'DB Security - Users'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Security.User } } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Security.User | Where-Object { $_.ID } } } ) | Measure-Object).Count + 1
			$ColumnCount = 15
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'User Name'
			$WorksheetData[0,$Col++] = 'Login Type'
			$WorksheetData[0,$Col++] = 'User Type'
			$WorksheetData[0,$Col++] = 'Authentication Type'
			$WorksheetData[0,$Col++] = 'Login'
			$WorksheetData[0,$Col++] = 'Certificate'
			$WorksheetData[0,$Col++] = 'Asymmetric Key'
			$WorksheetData[0,$Col++] = 'Default Schema'
			$WorksheetData[0,$Col++] = 'Default Language'
			$WorksheetData[0,$Col++] = 'Has DB Access'
			$WorksheetData[0,$Col++] = 'Is System Object'
			$WorksheetData[0,$Col++] = 'Create Date'
			$WorksheetData[0,$Col++] = 'Last Modified Date'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.ServerName

				$_.Server.Databases | ForEach-Object {

					$DatabaseName = $_.Name

					$_.Security.User | Where-Object { $_.ID } | ForEach-Object { 
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.LoginType
						$WorksheetData[$Row,$Col++] = $_.UserType
						$WorksheetData[$Row,$Col++] = $_.AuthenticationType
						$WorksheetData[$Row,$Col++] = $_.Login
						$WorksheetData[$Row,$Col++] = $_.Certificate
						$WorksheetData[$Row,$Col++] = $_.AsymmetricKey
						$WorksheetData[$Row,$Col++] = $_.DefaultSchema
						$WorksheetData[$Row,$Col++] = $_.DefaultLanguage
						$WorksheetData[$Row,$Col++] = $_.HasDbAccess
						$WorksheetData[$Row,$Col++] = $_.IsSystemObject
						$WorksheetData[$Row,$Col++] = $_.CreateDate
						$WorksheetData[$Row,$Col++] = $_.DateLastModified
						$Row++
					} 
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
						@{ColumnNumber = 14; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion 


			# Worksheet 42: Database Security - Database Roles
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Security - Database Roles"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'DB Security - Database Roles'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Security.DatabaseRole } } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Security.DatabaseRole | Where-Object { $_.ID } } } ) | Measure-Object).Count + 1
			$ColumnCount = 10
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Role Name'
			$WorksheetData[0,$Col++] = 'Owner'
			$WorksheetData[0,$Col++] = 'IsFixedRole'
			$WorksheetData[0,$Col++] = 'Create Date'
			$WorksheetData[0,$Col++] = 'Last Modified Date'
			$WorksheetData[0,$Col++] = 'Members'
			$WorksheetData[0,$Col++] = 'Member Of'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.ServerName

				$_.Server.Databases | ForEach-Object {

					$DatabaseName = $_.Name

					$_.Security.DatabaseRole | Where-Object { $_.ID } | ForEach-Object { 
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.Owner
						$WorksheetData[$Row,$Col++] = $_.IsFixedRole
						$WorksheetData[$Row,$Col++] = $_.CreateDate
						$WorksheetData[$Row,$Col++] = $_.DateLastModified
						$WorksheetData[$Row,$Col++] = ($_.Member | Sort-Object) -join $Delimiter # PoSH is more forgiving here than [String]::Join if $_.Member is $NULL
						$WorksheetData[$Row,$Col++] = ($_.MemberOf | Sort-Object) -join $Delimiter # PoSH is more forgiving here than [String]::Join if $_.MemberOf is NUL
						$Row++
					} 
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
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion 


			# Worksheet 43: Application Roles
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Security - Application Roles"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'DB Security - Application Roles'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Security.ApplicationRole } } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Security.ApplicationRole | Where-Object { $_.ID } } } ) | Measure-Object).Count + 1
			$ColumnCount = 6
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Role Name'
			$WorksheetData[0,$Col++] = 'Default Schema'
			$WorksheetData[0,$Col++] = 'Create Date'
			$WorksheetData[0,$Col++] = 'Last Modified Date'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.ServerName

				$_.Server.Databases | ForEach-Object {

					$DatabaseName = $_.Name

					$_.Security.ApplicationRole | Where-Object { $_.ID } | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.DefaultSchema
						$WorksheetData[$Row,$Col++] = $_.CreateDate
						$WorksheetData[$Row,$Col++] = $_.DateLastModified
						$Row++
					}
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
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 44: Database Security - Schemas
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Security - Schemas"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'DB Security - Schemas'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Security.Schema } } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Security.Schema | Where-Object { $_.Name } } } ) | Measure-Object).Count + 1
			$ColumnCount = 5
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Owner'
			$WorksheetData[0,$Col++] = 'Is System Object'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.ServerName

				$_.Server.Databases | ForEach-Object {

					$DatabaseName = $_.Name

					$_.Security.Schema | Where-Object { $_.Name } | ForEach-Object { 
						#$_.Security.Schema | Where-Object { ($_ | Measure-Object).Count -gt 0 } | ForEach-Object { 
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.Owner
						$WorksheetData[$Row,$Col++] = $_.IsSystemObject
						$Row++
					} 
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


			# Worksheet 45: Database Security - Asymmetric Keys
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Security - Asymmetric Keys"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'DB Security - Asymmetric Keys'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Security.AsymmetricKeys } } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Security.AsymmetricKeys | Where-Object { $_.ID } } } ) | Measure-Object).Count + 1
			$ColumnCount = 8
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Key Name'
			$WorksheetData[0,$Col++] = 'Owner'
			$WorksheetData[0,$Col++] = 'Encryption Algorithm'
			$WorksheetData[0,$Col++] = 'Key Length'
			$WorksheetData[0,$Col++] = 'Private Key Encryption Type'
			$WorksheetData[0,$Col++] = 'Provider Name'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.ServerName

				$_.Server.Databases | ForEach-Object {

					$DatabaseName = $_.Name

					$_.Security.AsymmetricKeys | Where-Object { $_.ID } | ForEach-Object {
						#$_.Security.AsymmetricKeys | Where-Object { ($_ | Measure-Object).Count -gt 0 } | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.Owner
						$WorksheetData[$Row,$Col++] = $_.KeyEncryptionAlgorithm
						$WorksheetData[$Row,$Col++] = $_.KeyLength
						$WorksheetData[$Row,$Col++] = $_.PrivateKeyEncryptionType
						$WorksheetData[$Row,$Col++] = $_.ProviderName
						$Row++
					} 
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


			# Worksheet 46: Database Security - Certificates
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Security - Certificates"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'DB Security - Certificates'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Security.Certificates } } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Security.Certificates | Where-Object { $_.ID } } } ) | Measure-Object).Count + 1
			$ColumnCount = 12
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Certificate Name'
			$WorksheetData[0,$Col++] = 'Owner'
			$WorksheetData[0,$Col++] = 'Private Key Encryption Type'
			$WorksheetData[0,$Col++] = 'Active For Service Broker'
			$WorksheetData[0,$Col++] = 'Issuer'
			$WorksheetData[0,$Col++] = 'Serial Number'
			$WorksheetData[0,$Col++] = 'Subject'
			$WorksheetData[0,$Col++] = 'Start Date'
			$WorksheetData[0,$Col++] = 'Expiration Date'
			$WorksheetData[0,$Col++] = 'Last Backup Date'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.ServerName

				$_.Server.Databases | ForEach-Object {

					$DatabaseName = $_.Name

					$_.Security.Certificates | Where-Object { $_.ID } | ForEach-Object {
						#$_.Security.Certificates | Where-Object { ($_ | Measure-Object).Count -gt 0 } | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.Owner
						$WorksheetData[$Row,$Col++] = $_.PrivateKeyEncryptionType
						$WorksheetData[$Row,$Col++] = $_.ActiveForServiceBrokerDialog
						$WorksheetData[$Row,$Col++] = $_.Issuer
						$WorksheetData[$Row,$Col++] = $_.SerialNumber
						$WorksheetData[$Row,$Col++] = $_.Subject
						$WorksheetData[$Row,$Col++] = $_.StartDate
						$WorksheetData[$Row,$Col++] = $_.ExpirationDate
						$WorksheetData[$Row,$Col++] = $_.LastBackupDate
						$Row++
					} 
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
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 11; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 12; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 47: Database Security - Symmetric Keys
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Security - Symmetric Keys"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'DB Security - Symmetric Keys'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Security.SymmetricKeys } } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Security.SymmetricKeys | Where-Object { $_.ID } } } ) | Measure-Object).Count + 1
			$ColumnCount = 11
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Key Name'
			$WorksheetData[0,$Col++] = 'Owner'
			$WorksheetData[0,$Col++] = 'Create Date'
			$WorksheetData[0,$Col++] = 'Date Last Modified'
			$WorksheetData[0,$Col++] = 'Encryption Algorithm'
			$WorksheetData[0,$Col++] = 'Key Length'
			$WorksheetData[0,$Col++] = 'Key Guid'
			$WorksheetData[0,$Col++] = 'Provider Name'
			$WorksheetData[0,$Col++] = 'Is Open'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.ServerName

				$_.Server.Databases | ForEach-Object {

					$DatabaseName = $_.Name

					$_.Security.SymmetricKeys | Where-Object { $_.ID } | ForEach-Object {
						#$_.Security.SymmetricKeys | Where-Object { ($_ | Measure-Object).Count -gt 0 } | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.Owner
						$WorksheetData[$Row,$Col++] = $_.CreateDate
						$WorksheetData[$Row,$Col++] = $_.DateLastModified
						$WorksheetData[$Row,$Col++] = $_.EncryptionAlgorithm
						$WorksheetData[$Row,$Col++] = $_.KeyLength
						$WorksheetData[$Row,$Col++] = $_.KeyGuid
						$WorksheetData[$Row,$Col++] = $_.ProviderName
						$WorksheetData[$Row,$Col++] = $_.IsOpen
						$Row++
					} 
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
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet X: Database Security - Audit Specifications


			# 			# Worksheet 48: Agent Configuration (VERTICAL FORMAT)
			# 			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Agent Configuration"
			# 			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			# 			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			# 			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			# 			$Worksheet.Name = 'Agent Config'
			# 			#$Worksheet.Tab.Color = $AgentTabColor
			# 			$Worksheet.Tab.ThemeColor = $AgentTabColor
			# 
			# 			#$ColumnCount = ($SqlServerInventory.DatabaseServer | Where-Object { $_.Agent.Configuration } | Measure-Object).Count + 1
			# 
			# 			# Option 1: Report only instances that have the SQL Agent enabled
			# 			$ColumnCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Agent.Configuration } ) | Measure-Object).Count + 1
			# 
			# 			# Option 2: Report all instances, leave blanks for those that don't have the SQL Agent enabled
			# 			#$ColumnCount = ($SqlServerInventory.DatabaseServer | Measure-Object).Count + 1
			# 
			# 			$RowCount = 50
			# 			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount
			# 
			# 			$Row = 0
			# 			$WorksheetData[$Row++,0] = 'Server Name'
			# 			$WorksheetData[$Row++,0] = 'General'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Name'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Server Type'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Agent Service'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Auto Restart SQL Server'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Auto Restart SQL Server Agent'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Error Log'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'File Name'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Log Level'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Include Execution Trace Messages'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Write Oem File'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Net Send Recipient'
			# 			#13 
			# 			$WorksheetData[$Row++,0] = 'Advanced'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'SQL Server Event Forwarding'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Is Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Server'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Events To Forward'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Severity At Or Above'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Idle CPU Condition'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Is Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Avg CPU Usage Falls Below (%)'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Time CPU Remains Below (sec)'
			# 			# 23
			# 			$WorksheetData[$Row++,0] = 'Alert System'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Mail Session'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Is Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Mail System'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Mail Profile'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Save Sent Messages'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Pager Emails'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'To Template'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Cc Template'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Subject'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Include Body In Notification'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Failsafe Operator'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Operator'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Notification Method'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Email Address'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Pager Address'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Net Send Address'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Token Replacement'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Replace Alert Tokens'
			# 			# 42
			# 			$WorksheetData[$Row++,0] = 'Job System'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Shutdown Time-Out Interval (sec)'
			# 			# 44
			# 			$WorksheetData[$Row++,0] = 'Connection'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Alias Local Host Server'
			# 			#$WorksheetData[$Row++,0] = $IndentString__2 + 'SQL Server Connection'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Login Timeout (sec)'
			# 			# 47
			# 			$WorksheetData[$Row++,0] = 'History'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Max Job History Total Rows'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Max Job History Rows Per Job'
			# 			# 50
			# 
			# 
			# 			$Col = 1
			# 
			# 			# Option 1: Report only instances that have the SQL Agent enabled
			# 			$SqlServerInventory.DatabaseServer | Where-Object { $_.Agent.Configuration } | Sort-Object -Property $_.ServerName | ForEach-Object {
			# 
			# 				# Option 2: Report all instances, leave blanks for those that don't have the SQL Agent enabled
			# 				#$SqlServerInventory.DatabaseServer | Sort-Object -Property $_.ServerName | ForEach-Object {
			# 				$Row = 0
			# 
			# 				# General
			# 				$WorksheetData[$Row++,$Col] = $_.ServerName
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.General.Name
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.General.ServerType
			# 
			# 				# General - Agent Service
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.General.AgentService.AutoRestartSqlAgent
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.General.AgentService.AutoRestartSqlServer
			# 
			# 				# General - Error Log
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.General.ErrorLog.FileName
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.General.ErrorLog.LogLevel
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.General.ErrorLog.IncludeExecutionTraceMessages
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.General.ErrorLog.WriteOemFile
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.General.ErrorLog.NetSendRecipient
			# 
			# 				# Advanced
			# 				$WorksheetData[$Row++,$Col] = $null
			# 
			# 				# Advanced - SQL Server Event Forwarding
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.Advanced.EventForwarding.IsEnabled
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.Advanced.EventForwarding.Server
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.Advanced.EventForwarding.EventsToForward
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.Advanced.EventForwarding.SeverityAtOrAbove
			# 
			# 				# Advanced - Idle CPU Condition
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.Advanced.IdleCpuCondition.IsEnabled
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.Advanced.IdleCpuCondition.AvgCpuBelowPercent
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.Advanced.IdleCpuCondition.AvgCpuRemainsForSeconds
			# 
			# 				# Alert System
			# 				$WorksheetData[$Row++,$Col] = $null
			# 
			# 				# Alert System - Mail Session
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.AlertSystem.MailSession.IsEnabled
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.AlertSystem.MailSession.MailSystem
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.AlertSystem.MailSession.MailProfile
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.AlertSystem.MailSession.SaveSentMessages
			# 
			# 				# Alert System - Pager Emails
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.AlertSystem.PagerEmails.ToTemplate
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.AlertSystem.PagerEmails.CcTemplate
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.AlertSystem.PagerEmails.Subject
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.AlertSystem.PagerEmails.IncludeBody
			# 
			# 				# Alert System - Failsafe Operator
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.AlertSystem.FailSafeOperator.Operator
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.AlertSystem.FailSafeOperator.NotificationMethod
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.AlertSystem.FailSafeOperator.FailSafeEmailAddress
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.AlertSystem.FailSafeOperator.FailSafePagerAddress
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.AlertSystem.FailSafeOperator.FailSafeNetSendAddress
			# 
			# 				# Alert System - Token Replacement
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.AlertSystem.TokenReplacement.ReplaceAlertTokens
			# 
			# 				# Job System
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.JobSystem.ShutdownTimeoutIntervalSeconds
			# 
			# 				# Connection
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.Connection.AliasLocalHostServer
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.Connection.LoginTimeoutSeconds
			# 
			# 				# Connection
			# 				$WorksheetData[$Row++,$Col] = $null
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.History.MaxJobHistoryTotalRows
			# 				$WorksheetData[$Row++,$Col] = $_.Agent.Configuration.History.MaxJobHistoryRowsPerJob
			# 
			# 				$Col++
			# 			}
			# 			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			# 			$Range.Value2 = $WorksheetData
			# 			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			# 			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			# 			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null
			# 
			# 			$WorksheetFormat.Add($WorksheetNumber, @{
			# 					BoldFirstRow = $false
			# 					BoldFirstColumn = $true
			# 					AutoFilter = $false
			# 					FreezeAtCell = 'B2'
			# 					ColumnFormat = @()
			# 					RowFormat = @(
			# 						@{RowNumber = 22; NumberFormat = $XlNumFmtNumberGeneral},
			# 						@{RowNumber = 23; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 44; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 47; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 49; NumberFormat = $XlNumFmtNumberS0},
			# 						@{RowNumber = 50; NumberFormat = $XlNumFmtNumberS0}
			# 					)
			# 				})
			# 
			# 			$WorksheetNumber++
			#endregion


			# Worksheet 48: Agent Configuration
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Agent Configuration"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Agent Config'
			#$Worksheet.Tab.Color = $AgentTabColor
			$Worksheet.Tab.ThemeColor = $AgentTabColor


			# Option 1: Report only instances that have the SQL Agent enabled
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Agent.Configuration } ) | Measure-Object).Count + 1

			# Option 2: Report all instances, leave blanks for those that don't have the SQL Agent enabled
			#$RowCount = ($SqlServerInventory.DatabaseServer | Measure-Object).Count + 1

			$ColumnCount = 50
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'General >'
			$WorksheetData[0,$Col++] = 'Name'
			$WorksheetData[0,$Col++] = 'Server Type'
			$WorksheetData[0,$Col++] = 'Agent Service >>'
			$WorksheetData[0,$Col++] = 'Auto Restart SQL Server'
			$WorksheetData[0,$Col++] = 'Auto Restart SQL Server Agent'
			$WorksheetData[0,$Col++] = 'Error Log >>'
			$WorksheetData[0,$Col++] = 'File Name'
			$WorksheetData[0,$Col++] = 'Log Level'
			$WorksheetData[0,$Col++] = 'Include Execution Trace Messages'
			$WorksheetData[0,$Col++] = 'Write Oem File'
			$WorksheetData[0,$Col++] = 'Net Send Recipient'
			#13 
			$WorksheetData[0,$Col++] = 'Advanced >'
			$WorksheetData[0,$Col++] = 'SQL Server Event Forwarding >>'
			$WorksheetData[0,$Col++] = 'Is Enabled'
			$WorksheetData[0,$Col++] = 'Server'
			$WorksheetData[0,$Col++] = 'Events To Forward'
			$WorksheetData[0,$Col++] = 'Severity At Or Above'
			$WorksheetData[0,$Col++] = 'Idle CPU Condition >>'
			$WorksheetData[0,$Col++] = 'Is Enabled'
			$WorksheetData[0,$Col++] = 'Avg CPU Usage Falls Below (%)'
			$WorksheetData[0,$Col++] = 'Time CPU Remains Below (sec)'
			# 23
			$WorksheetData[0,$Col++] = 'Alert System >'
			$WorksheetData[0,$Col++] = 'Mail Session >>'
			$WorksheetData[0,$Col++] = 'Is Enabled'
			$WorksheetData[0,$Col++] = 'Mail System'
			$WorksheetData[0,$Col++] = 'Mail Profile'
			$WorksheetData[0,$Col++] = 'Save Sent Messages'
			$WorksheetData[0,$Col++] = 'Pager Emails >>'
			$WorksheetData[0,$Col++] = 'To Template'
			$WorksheetData[0,$Col++] = 'Cc Template'
			$WorksheetData[0,$Col++] = 'Subject'
			$WorksheetData[0,$Col++] = 'Include Body In Notification'
			$WorksheetData[0,$Col++] = 'Failsafe Operator >>'
			$WorksheetData[0,$Col++] = 'Operator'
			$WorksheetData[0,$Col++] = 'Notification Method'
			$WorksheetData[0,$Col++] = 'Email Address'
			$WorksheetData[0,$Col++] = 'Pager Address'
			$WorksheetData[0,$Col++] = 'Net Send Address'
			$WorksheetData[0,$Col++] = 'Token Replacement >>'
			$WorksheetData[0,$Col++] = 'Replace Alert Tokens'
			# 42
			$WorksheetData[0,$Col++] = 'Job System >'
			$WorksheetData[0,$Col++] = 'Shutdown Time-Out Interval (sec)'
			# 44
			$WorksheetData[0,$Col++] = 'Connection >'
			$WorksheetData[0,$Col++] = 'Alias Local Host Server'
			#$WorksheetData[0,$Col++] = 'SQL Server Connection'
			$WorksheetData[0,$Col++] = 'Login Timeout (sec)'
			# 47
			$WorksheetData[0,$Col++] = 'History >'
			$WorksheetData[0,$Col++] = 'Max Job History Total Rows'
			$WorksheetData[0,$Col++] = 'Max Job History Rows Per Job'
			# 50


			$Row = 1

			# Option 1: Report only instances that have the SQL Agent enabled
			$SqlServerInventory.DatabaseServer | Where-Object { $_.Agent.Configuration } | ForEach-Object {

				# Option 2: Report all instances, leave blanks for those that don't have the SQL Agent enabled
				#$SqlServerInventory.DatabaseServer | Sort-Object -Property $_.ServerName | ForEach-Object {
				$Col = 0

				# General
				$WorksheetData[$Row,$Col++] = $_.ServerName
				$WorksheetData[$Row,$Col++] = $null
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.General.Name
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.General.ServerType

				# General - Agent Service
				$WorksheetData[$Row,$Col++] = $null
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.General.AgentService.AutoRestartSqlAgent
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.General.AgentService.AutoRestartSqlServer

				# General - Error Log
				$WorksheetData[$Row,$Col++] = $null
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.General.ErrorLog.FileName
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.General.ErrorLog.LogLevel
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.General.ErrorLog.IncludeExecutionTraceMessages
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.General.ErrorLog.WriteOemFile
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.General.ErrorLog.NetSendRecipient

				# Advanced
				$WorksheetData[$Row,$Col++] = $null

				# Advanced - SQL Server Event Forwarding
				$WorksheetData[$Row,$Col++] = $null
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.Advanced.EventForwarding.IsEnabled
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.Advanced.EventForwarding.Server
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.Advanced.EventForwarding.EventsToForward
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.Advanced.EventForwarding.SeverityAtOrAbove

				# Advanced - Idle CPU Condition
				$WorksheetData[$Row,$Col++] = $null
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.Advanced.IdleCpuCondition.IsEnabled
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.Advanced.IdleCpuCondition.AvgCpuBelowPercent
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.Advanced.IdleCpuCondition.AvgCpuRemainsForSeconds

				# Alert System
				$WorksheetData[$Row,$Col++] = $null

				# Alert System - Mail Session
				$WorksheetData[$Row,$Col++] = $null
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.AlertSystem.MailSession.IsEnabled
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.AlertSystem.MailSession.MailSystem
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.AlertSystem.MailSession.MailProfile
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.AlertSystem.MailSession.SaveSentMessages

				# Alert System - Pager Emails
				$WorksheetData[$Row,$Col++] = $null
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.AlertSystem.PagerEmails.ToTemplate
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.AlertSystem.PagerEmails.CcTemplate
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.AlertSystem.PagerEmails.Subject
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.AlertSystem.PagerEmails.IncludeBody

				# Alert System - Failsafe Operator
				$WorksheetData[$Row,$Col++] = $null
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.AlertSystem.FailSafeOperator.Operator
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.AlertSystem.FailSafeOperator.NotificationMethod
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.AlertSystem.FailSafeOperator.FailSafeEmailAddress
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.AlertSystem.FailSafeOperator.FailSafePagerAddress
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.AlertSystem.FailSafeOperator.FailSafeNetSendAddress

				# Alert System - Token Replacement
				$WorksheetData[$Row,$Col++] = $null
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.AlertSystem.TokenReplacement.ReplaceAlertTokens

				# Job System
				$WorksheetData[$Row,$Col++] = $null
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.JobSystem.ShutdownTimeoutIntervalSeconds

				# Connection
				$WorksheetData[$Row,$Col++] = $null
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.Connection.AliasLocalHostServer
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.Connection.LoginTimeoutSeconds

				# Connection
				$WorksheetData[$Row,$Col++] = $null
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.History.MaxJobHistoryTotalRows
				$WorksheetData[$Row,$Col++] = $_.Agent.Configuration.History.MaxJobHistoryRowsPerJob

				$Row++
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'B2'
					ColumnFormat = @(
						@{ColumnNumber = 22; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 23; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 44; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 47; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 49; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 50; NumberFormat = $XlNumFmtNumberS0} 
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 49: Agent Jobs
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Agent Jobs"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Agent Jobs'
			#$Worksheet.Tab.Color = $AgentTabColor
			$Worksheet.Tab.ThemeColor = $AgentTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Agent.Jobs } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Agent.Jobs | Where-Object { $_.General.Name } } ) | Measure-Object).Count + 1
			$ColumnCount = 11
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Job Name'
			$WorksheetData[0,$Col++] = 'Owner'
			$WorksheetData[0,$Col++] = 'Category'
			$WorksheetData[0,$Col++] = 'Enabled'
			#$WorksheetData[0,$Col++] = 'Source'
			$WorksheetData[0,$Col++] = 'Create Date'
			$WorksheetData[0,$Col++] = 'Last Modify Date'
			$WorksheetData[0,$Col++] = 'Last Executed Date'
			$WorksheetData[0,$Col++] = 'Last Outcome'
			$WorksheetData[0,$Col++] = 'Next Run Date'
			$WorksheetData[0,$Col++] = 'Description'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.ServerName

				$_.Agent.Jobs | Where-Object { $_.General.Name } | ForEach-Object {
					#$_.Agent.Jobs | Where-Object { ($_ | Measure-Object).Count -gt 0 } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.General.Name
					$WorksheetData[$Row,$Col++] = $_.General.Owner
					$WorksheetData[$Row,$Col++] = $_.General.Category
					$WorksheetData[$Row,$Col++] = $_.General.Enabled
					#$WorksheetData[$Row,$Col++] = $null
					$WorksheetData[$Row,$Col++] = $_.General.Created
					$WorksheetData[$Row,$Col++] = $_.General.LastModified
					$WorksheetData[$Row,$Col++] = $_.General.LastExecuted
					$WorksheetData[$Row,$Col++] = $_.General.LastOutcome
					$WorksheetData[$Row,$Col++] = $_.General.NextRun
					$WorksheetData[$Row,$Col++] = $_.General.Description
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
					ColumnFormat = @(
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion 


			# Worksheet 50: Agent Jobs - Steps
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Agent Jobs - Steps"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Agent Jobs - Steps'
			#$Worksheet.Tab.Color = $AgentTabColor
			$Worksheet.Tab.ThemeColor = $AgentTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Agent.Jobs | ForEach-Object { $_.Steps } } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Agent.Jobs | ForEach-Object { $_.Steps | Where-Object { $_.Id } } } ) | Measure-Object).Count + 1
			$ColumnCount = 17
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Job Name'
			$WorksheetData[0,$Col++] = 'Job Enabled'
			$WorksheetData[0,$Col++] = 'Step Name'
			$WorksheetData[0,$Col++] = 'Step ID'
			$WorksheetData[0,$Col++] = 'Type'
			$WorksheetData[0,$Col++] = 'Run As'
			$WorksheetData[0,$Col++] = 'Database'
			$WorksheetData[0,$Col++] = 'Command'
			$WorksheetData[0,$Col++] = 'Success Exit Code'
			$WorksheetData[0,$Col++] = 'On Success Action'
			$WorksheetData[0,$Col++] = 'On Success Step'
			$WorksheetData[0,$Col++] = 'On Fail Action'
			$WorksheetData[0,$Col++] = 'On Fail Step'
			$WorksheetData[0,$Col++] = 'Retry Attempts'
			$WorksheetData[0,$Col++] = 'Retry Interval (mins)'
			$WorksheetData[0,$Col++] = 'Log File'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.ServerName

				$_.Agent.Jobs | ForEach-Object {

					$JobName = $_.General.Name
					$JobEnabled = $_.General.Enabled

					$_.Steps | Where-Object { $_.Id } | ForEach-Object {
						#$_.Steps | Where-Object { ($_ | Measure-Object).Count -gt 0 } | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $JobName
						$WorksheetData[$Row,$Col++] = $JobEnabled
						$WorksheetData[$Row,$Col++] = $_.General.StepName
						$WorksheetData[$Row,$Col++] = $_.Id
						$WorksheetData[$Row,$Col++] = $_.General.Type
						$WorksheetData[$Row,$Col++] = $_.General.RunAs
						$WorksheetData[$Row,$Col++] = $_.General.Database
						#$WorksheetData[$Row,$Col++] = if ($_.General.Command.Length -gt 32767) { $_.General.Command.Substring(0, 32764) + '...' } else { $_.General.Command }
						$WorksheetData[$Row,$Col++] = if ($_.General.Command.Length -gt 5000) { $_.General.Command.Substring(0, 4997) + '...' } else { $_.General.Command }
						$WorksheetData[$Row,$Col++] = $_.General.SuccessExitCode
						$WorksheetData[$Row,$Col++] = $_.Advanced.OnSuccessAction
						$WorksheetData[$Row,$Col++] = $_.Advanced.OnSuccessStep
						$WorksheetData[$Row,$Col++] = $_.Advanced.OnFailAction
						$WorksheetData[$Row,$Col++] = $_.Advanced.OnFailStep
						$WorksheetData[$Row,$Col++] = $_.Advanced.RetryAttempts
						$WorksheetData[$Row,$Col++] = $_.Advanced.RetryIntervalMinutes
						$WorksheetData[$Row,$Col++] = $_.Advanced.Logging.OutputFile
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(5), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 12; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 14; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 16; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion 


			# Worksheet 51: Agent Jobs - Schedules
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Agent Jobs - Schedules"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Agent Jobs - Schedules'
			#$Worksheet.Tab.Color = $AgentTabColor
			$Worksheet.Tab.ThemeColor = $AgentTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Agent.Jobs | ForEach-Object { $_.Schedules } } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Agent.Jobs | ForEach-Object { $_.Schedules | Where-Object { $_.Id } } } ) | Measure-Object).Count + 1
			$ColumnCount = 7
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Job Name'
			$WorksheetData[0,$Col++] = 'Job Enabled'
			$WorksheetData[0,$Col++] = 'Schedule Name'
			$WorksheetData[0,$Col++] = 'Schedule Enabled'
			$WorksheetData[0,$Col++] = 'Create Date'
			$WorksheetData[0,$Col++] = 'Description'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.ServerName

				$_.Agent.Jobs | ForEach-Object {

					$JobName = $_.General.Name
					$JobEnabled = $_.General.Enabled

					$_.Schedules | Where-Object { $_.Id } | ForEach-Object {
						#$_.Schedules | Where-Object { ($_ | Measure-Object).Count -gt 0 } | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $JobName
						$WorksheetData[$Row,$Col++] = $JobEnabled
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.IsEnabled
						$WorksheetData[$Row,$Col++] = $_.DateCreated
						$WorksheetData[$Row,$Col++] = $_.Description
						$Row++
					}
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
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 52: Agent Jobs - Alerts
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Agent Jobs - Alerts"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Agent Jobs - Alerts'
			#$Worksheet.Tab.Color = $AgentTabColor
			$Worksheet.Tab.ThemeColor = $AgentTabColor

			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Agent.Jobs | ForEach-Object { $_.Alerts | Where-Object { $_.General.Id } } } ) | Measure-Object).Count + 1
			$ColumnCount = 13
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Job Name'
			$WorksheetData[0,$Col++] = 'Job Enabled'
			$WorksheetData[0,$Col++] = 'Alert Name'
			$WorksheetData[0,$Col++] = 'Alert Enabled'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Error Number'
			$WorksheetData[0,$Col++] = 'Severity'
			$WorksheetData[0,$Col++] = 'Keyword'
			$WorksheetData[0,$Col++] = 'Performance Condition'
			$WorksheetData[0,$Col++] = 'WMI Namespace'
			$WorksheetData[0,$Col++] = 'WMI Query'
			$WorksheetData[0,$Col++] = 'Notify Operators'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.ServerName

				$_.Agent.Jobs | ForEach-Object {

					$JobName = $_.General.Name
					$JobEnabled = $_.General.Enabled

					$_.Alerts | Where-Object { $_.General.Id } | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $JobName
						$WorksheetData[$Row,$Col++] = $JobEnabled
						$WorksheetData[$Row,$Col++] = $_.General.Name
						$WorksheetData[$Row,$Col++] = $_.General.IsEnabled
						$WorksheetData[$Row,$Col++] = $_.Definition.DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Definition.ErrorNumber
						$WorksheetData[$Row,$Col++] = $_.Definition.Severity
						$WorksheetData[$Row,$Col++] = $_.Definition.EventDescriptionKeyword
						$WorksheetData[$Row,$Col++] = $_.Definition.PerformanceCondition
						$WorksheetData[$Row,$Col++] = $_.Definition.WmiNamespace
						$WorksheetData[$Row,$Col++] = $_.Definition.WmiQuery
						$WorksheetData[$Row,$Col++] = if (($_.Response.NotifyOperators | Measure-Object).Count -gt 0) {
							[String]::Join($Delimiter, @(
									$_.Response.NotifyOperators | Where-Object { $_.OperatorName } | ForEach-Object {
										$NotifyMethod = @()

										if ($_.UseEmail) { $NotifyMethod += 'E-mail' }
										if ($_.UsePager) { $NotifyMethod += 'Pager' }
										if ($_.UseNetSend) { $NotifyMethod += 'Net Send' }

										$_.OperatorName + ' (' + [String]::Join(',', $NotifyMethod) + ')'
									}
								))
						} else {
							[String]::Empty
						}
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(4), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion 


			# Worksheet 53: Agent Jobs - Notifications
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Agent Jobs - Notifications"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Agent Jobs - Notifications'
			#$Worksheet.Tab.Color = $AgentTabColor
			$Worksheet.Tab.ThemeColor = $AgentTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Agent.Jobs } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Agent.Jobs | Where-Object { $_.General.Name } } ) | Measure-Object).Count + 1
			$ColumnCount = 11
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Job Name'
			$WorksheetData[0,$Col++] = 'Email Operator'
			$WorksheetData[0,$Col++] = 'Email Condition'
			$WorksheetData[0,$Col++] = 'Page Operator'
			$WorksheetData[0,$Col++] = 'Page Condition'
			$WorksheetData[0,$Col++] = 'Net Send Operator'
			$WorksheetData[0,$Col++] = 'Net Send Condition'
			$WorksheetData[0,$Col++] = 'Write To Windows Application Event Log'
			$WorksheetData[0,$Col++] = 'Delete Job'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.ServerName

				$_.Agent.Jobs | Where-Object { $_.General.Name } | ForEach-Object {
					#$_.Agent.Jobs | Where-Object { ($_ | Measure-Object).Count -gt 0 } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.General.Name
					$WorksheetData[$Row,$Col++] = $_.Notifications.EmailOperator
					$WorksheetData[$Row,$Col++] = $_.Notifications.EmailCondition
					$WorksheetData[$Row,$Col++] = $_.Notifications.PageOperator
					$WorksheetData[$Row,$Col++] = $_.Notifications.PageCondition
					$WorksheetData[$Row,$Col++] = $_.Notifications.NetSendOperator
					$WorksheetData[$Row,$Col++] = $_.Notifications.NetSendCondition
					$WorksheetData[$Row,$Col++] = $_.Notifications.EventLogCondition
					$WorksheetData[$Row,$Col++] = $_.Notifications.DeleteJobCondition
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
					ColumnFormat = @(
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion 


			# Worksheet 54: Agent Alerts
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Agent Alerts"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Agent Alerts'
			#$Worksheet.Tab.Color = $AgentTabColor
			$Worksheet.Tab.ThemeColor = $AgentTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Agent.Alerts } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Agent.Alerts | Where-Object { $_.General.ID } } ) | Measure-Object).Count + 1
			$ColumnCount = 23
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Alert Name'
			$WorksheetData[0,$Col++] = 'Enabled'
			$WorksheetData[0,$Col++] = 'Category Name'
			$WorksheetData[0,$Col++] = 'Event Source'
			$WorksheetData[0,$Col++] = 'Type'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Error Number'
			$WorksheetData[0,$Col++] = 'Severity'
			$WorksheetData[0,$Col++] = 'Message Text Contains'
			$WorksheetData[0,$Col++] = 'Performance Condition'
			$WorksheetData[0,$Col++] = 'WMI Namespace'
			$WorksheetData[0,$Col++] = 'WMI Query'
			$WorksheetData[0,$Col++] = 'Execute Job ID'
			$WorksheetData[0,$Col++] = 'Execute Job Name'
			$WorksheetData[0,$Col++] = 'Notify Operators'
			$WorksheetData[0,$Col++] = 'Include Error Alert In'
			$WorksheetData[0,$Col++] = 'Additional Message To Send'
			$WorksheetData[0,$Col++] = 'Delay Between Responses (sec)'
			$WorksheetData[0,$Col++] = 'Last Alert'
			$WorksheetData[0,$Col++] = 'Last Response'
			$WorksheetData[0,$Col++] = 'Number Of Occurrences'
			$WorksheetData[0,$Col++] = 'Last Occurrence Count Reset'


			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.ServerName

				$_.Agent.Alerts | Where-Object { $_.General.ID } | ForEach-Object {
					#$_.Agent.Alerts | Where-Object { ($_ | Measure-Object).Count -gt 0 } | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.General.Name
					$WorksheetData[$Row,$Col++] = $_.General.IsEnabled
					$WorksheetData[$Row,$Col++] = $_.General.CategoryName
					$WorksheetData[$Row,$Col++] = $_.General.EventSource
					$WorksheetData[$Row,$Col++] = $_.General.Type
					$WorksheetData[$Row,$Col++] = $_.General.Definition.DatabaseName
					$WorksheetData[$Row,$Col++] = $_.General.Definition.ErrorNumber
					$WorksheetData[$Row,$Col++] = $_.General.Definition.Severity
					$WorksheetData[$Row,$Col++] = $_.General.Definition.EventDescriptionKeyword
					$WorksheetData[$Row,$Col++] = $_.General.Definition.PerformanceCondition
					$WorksheetData[$Row,$Col++] = $_.General.Definition.WmiNamespace
					$WorksheetData[$Row,$Col++] = $_.General.Definition.WmiQuery
					$WorksheetData[$Row,$Col++] = $_.Response.ExecuteJobID
					$WorksheetData[$Row,$Col++] = $_.Response.ExecuteJobName
					$WorksheetData[$Row,$Col++] = if (($_.Response.NotifyOperators | Measure-Object).Count -gt 0) {
						[String]::Join($Delimiter, @(
								$_.Response.NotifyOperators | Where-Object { $_.OperatorId } | ForEach-Object {

									$NotifyMethod = @()
									if ($_.UseEmail) { $NotifyMethod += 'E-mail' }
									if ($_.UsePager) { $NotifyMethod += 'Pager' }
									if ($_.UseNetSend) { $NotifyMethod += 'Net Send' } 

									$_.OperatorName + ' (' + [String]::Join(',', $NotifyMethod) + ')'

								} 
							)) 
					} else {
						[String]::Empty
					}
					$WorksheetData[$Row,$Col++] = $_.Options.IncludeErrorTextIn
					$WorksheetData[$Row,$Col++] = $_.Options.NotificationMessage
					$WorksheetData[$Row,$Col++] = $_.Options.DelaySecondsBetweenResponses
					$WorksheetData[$Row,$Col++] = $_.History.LastAlertDate
					$WorksheetData[$Row,$Col++] = $_.History.LastResponseDate
					$WorksheetData[$Row,$Col++] = $_.History.OccurrenceCount
					$WorksheetData[$Row,$Col++] = $_.History.CountResetDate
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
					ColumnFormat = @(
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 19; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 20; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 21; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 22; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 23; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion 


			# 			# Worksheet 55: Agent Operators  (VERTICAL FORMAT)
			# 			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Agent Operators"
			# 			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			# 			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			# 			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			# 			$Worksheet.Name = 'Agent Operators 2'
			# 			#$Worksheet.Tab.Color = $AgentTabColor
			# 			$Worksheet.Tab.ThemeColor = $AgentTabColor
			# 
			# 			#$RowCount = ($SqlServerInventory.DatabaseServer | Measure-Object).Count + 1
			# 			#$ColumnCount = 15
			# 			$ColumnCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Agent.Operators } ) | Measure-Object).Count + 1
			# 			$RowCount = 24
			# 			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount
			# 
			# 			$Row = 0
			# 			$WorksheetData[$Row++,0] = 'Server Name'
			# 			$WorksheetData[$Row++,0] = 'Operator Name'
			# 			# 2
			# 			$WorksheetData[$Row++,0] = 'General'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Category'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Enabled'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Notification Options'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'E-mail name'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Net send address'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Pager e-mail name'
			# 			# 9
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Pager On Duty Schedule'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'On Duty Days'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Mon-Fri Workday Begin'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Mon-Fri Workday End'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Sat Workday Begin'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Sat Workday End'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Sun Workday Begin'
			# 			$WorksheetData[$Row++,0] = $IndentString__2 + 'Sun Workday End'
			# 			# 17
			# 			$WorksheetData[$Row++,0] = 'Notifications'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Alerts'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Jobs'
			# 			# 20
			# 			$WorksheetData[$Row++,0] = 'History'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Last E-mail'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Last Pager'
			# 			$WorksheetData[$Row++,0] = $IndentString_1 + 'Last Net Send'
			# 			# 24
			# 
			# 
			# 			$Col = 1
			# 			$SqlServerInventory.DatabaseServer | Sort-Object -Property $_.ServerName | ForEach-Object {
			# 
			# 				$ServerName = $_.Server.Configuration.General.Name
			# 
			# 				$_.Agent.Operators | Where-Object { $_.General.ID } | Sort-Object -Property $_.General.Name | ForEach-Object {
			# 
			# 					$Row = 0
			# 
			# 					$WorksheetData[$Row++,$Col] = $ServerName
			# 					$WorksheetData[$Row++,$Col] = $_.General.Name
			# 					$WorksheetData[$Row++,$Col] = $_.General.CategoryName
			# 					$WorksheetData[$Row++,$Col] = $_.General.IsEnabled
			# 
			# 					# General - Notification Options
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.General.NotificationOptions.EmailAddress
			# 					$WorksheetData[$Row++,$Col] = $_.General.NotificationOptions.NetSendAddress
			# 					$WorksheetData[$Row++,$Col] = $_.General.NotificationOptions.PagerAddress
			# 
			# 					# General - Pager on duty schedule
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.General.OnDutySchedule.OnDutyDays
			# 					$WorksheetData[$Row++,$Col] = $_.General.OnDutySchedule.WeekdayStartTime
			# 					$WorksheetData[$Row++,$Col] = $_.General.OnDutySchedule.WeekdayEndTime
			# 					$WorksheetData[$Row++,$Col] = $_.General.OnDutySchedule.SaturdayStartTime
			# 					$WorksheetData[$Row++,$Col] = $_.General.OnDutySchedule.SaturdayEndTime
			# 					$WorksheetData[$Row++,$Col] = $_.General.OnDutySchedule.SundayStartTime
			# 					$WorksheetData[$Row++,$Col] = $_.General.OnDutySchedule.SundayEndTime
			# 
			# 					# Notifications
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = if (($_.Notifications.Alerts | Measure-Object).Count -gt 0) {
			# 						[String]::Join('`n', @(
			# 							$_.Notifications.Alerts | Where-Object { $_.AlertId } | ForEach-Object {
			# 		
			# 								$NotifyMethod = @()
			# 								if ($_.UseEmail) { $NotifyMethod += 'E-mail' }
			# 								if ($_.UsePager) { $NotifyMethod += 'Pager' }
			# 								if ($_.UseNetSend) { $NotifyMethod += 'Net Send' }						
			# 							
			# 								$_.AlertName + ' (' + [String]::Join(',', $NotifyMethod) + ')'
			# 								
			# 							}							
			# 						))					
			# 					} else {
			# 						[String]::Empty
			# 					}
			# 					
			# 					$WorksheetData[$Row++,$Col] = if (($_.Notifications.Jobs | Measure-Object).Count -gt 0) {
			# 						[String]::Join('`n', @(
			# 							$_.Notifications.Jobs | Where-Object { $_.JobId } | ForEach-Object {
			# 		
			# 								$NotifyMethod = @()
			# 								if ($_.UseEmail) { $NotifyMethod += 'E-mail' }
			# 								if ($_.UsePager) { $NotifyMethod += 'Pager' }
			# 								if ($_.UseNetSend) { $NotifyMethod += 'Net Send' }						
			# 							
			# 								$_.JobName + ' (' + [String]::Join(',', $NotifyMethod) + ')'
			# 								
			# 							}							
			# 						))					
			# 					} else {
			# 						[String]::Empty
			# 					}
			# 
			# 
			# 					# History
			# 					$WorksheetData[$Row++,$Col] = $null
			# 					$WorksheetData[$Row++,$Col] = $_.History.LastEmailDate
			# 					$WorksheetData[$Row++,$Col] = $_.History.LastPagerDate
			# 					$WorksheetData[$Row++,$Col] = $_.History.LastNetSendDate
			# 
			# 					$Col++
			# 
			# 				}
			# 			}
			# 			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			# 			$Range.Value2 = $WorksheetData
			# 			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			# 			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			# 			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null
			# 
			# 			$WorksheetFormat.Add($WorksheetNumber, @{
			# 					BoldFirstRow = $false
			# 					BoldFirstColumn = $true
			# 					AutoFilter = $false
			# 					FreezeAtCell = 'B2'
			# 					ColumnFormat = @()
			# 					RowFormat = @(
			# 						@{RowNumber = 12; NumberFormat = $XlNumFmtTime},
			# 						@{RowNumber = 13; NumberFormat = $XlNumFmtTime},
			# 						@{RowNumber = 14; NumberFormat = $XlNumFmtTime},
			# 						@{RowNumber = 15; NumberFormat = $XlNumFmtTime},
			# 						@{RowNumber = 16; NumberFormat = $XlNumFmtTime},
			# 						@{RowNumber = 17; NumberFormat = $XlNumFmtTime},
			# 						@{RowNumber = 22; NumberFormat = $XlNumFmtDate},
			# 						@{RowNumber = 23; NumberFormat = $XlNumFmtDate},
			# 						@{RowNumber = 24; NumberFormat = $XlNumFmtDate}
			# 					)
			# 				})
			# 
			# 			$WorksheetNumber++
			#endregion


			# Worksheet 55: Agent Operators
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Agent Operators"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Agent Operators'
			#$Worksheet.Tab.Color = $AgentTabColor
			$Worksheet.Tab.ThemeColor = $AgentTabColor

			#$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Agent.Operators } ) | Measure-Object).Count + 1
			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Agent.Operators | Where-Object { $_.General.ID } } ) | Measure-Object).Count + 1
			$ColumnCount = 19
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Operator Name'
			$WorksheetData[0,$Col++] = 'Category'
			$WorksheetData[0,$Col++] = 'Enabled'
			$WorksheetData[0,$Col++] = 'E-mail name'
			$WorksheetData[0,$Col++] = 'Net send address'
			$WorksheetData[0,$Col++] = 'Pager e-mail name'
			$WorksheetData[0,$Col++] = 'Pager On Duty Days'
			$WorksheetData[0,$Col++] = 'M-F Workday Begin'
			$WorksheetData[0,$Col++] = 'M-F Workday End'
			$WorksheetData[0,$Col++] = 'Sat Workday Begin'
			$WorksheetData[0,$Col++] = 'Sat Workday End'
			$WorksheetData[0,$Col++] = 'Sun Workday Begin'
			$WorksheetData[0,$Col++] = 'Sun Workday End'
			$WorksheetData[0,$Col++] = 'Last E-mail'
			$WorksheetData[0,$Col++] = 'Last Pager'
			$WorksheetData[0,$Col++] = 'Last Net Send'
			$WorksheetData[0,$Col++] = 'Alerts'
			$WorksheetData[0,$Col++] = 'Jobs'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.ServerName

				$_.Agent.Operators | Where-Object { $_.General.ID } | Sort-Object -Property $_.General.Name | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $_.General.Name
					$WorksheetData[$Row,$Col++] = $_.General.CategoryName
					$WorksheetData[$Row,$Col++] = $_.General.IsEnabled
					$WorksheetData[$Row,$Col++] = $_.General.NotificationOptions.EmailAddress
					$WorksheetData[$Row,$Col++] = $_.General.NotificationOptions.NetSendAddress
					$WorksheetData[$Row,$Col++] = $_.General.NotificationOptions.PagerAddress
					$WorksheetData[$Row,$Col++] = $_.General.OnDutySchedule.OnDutyDays
					$WorksheetData[$Row,$Col++] = $_.General.OnDutySchedule.WeekdayStartTime
					$WorksheetData[$Row,$Col++] = $_.General.OnDutySchedule.WeekdayEndTime
					$WorksheetData[$Row,$Col++] = $_.General.OnDutySchedule.SaturdayStartTime
					$WorksheetData[$Row,$Col++] = $_.General.OnDutySchedule.SaturdayEndTime
					$WorksheetData[$Row,$Col++] = $_.General.OnDutySchedule.SundayStartTime
					$WorksheetData[$Row,$Col++] = $_.General.OnDutySchedule.SundayEndTime
					$WorksheetData[$Row,$Col++] = $_.History.LastEmailDate
					$WorksheetData[$Row,$Col++] = $_.History.LastPagerDate
					$WorksheetData[$Row,$Col++] = $_.History.LastNetSendDate

					$WorksheetData[$Row,$Col++] = if (($_.Notifications.Alerts | Measure-Object).Count -gt 0) {
						[String]::Join($Delimiter, @(
								$_.Notifications.Alerts | Where-Object { $_.AlertId } | ForEach-Object {

									$NotifyMethod = @()
									if ($_.UseEmail) { $NotifyMethod += 'E-mail' }
									if ($_.UsePager) { $NotifyMethod += 'Pager' }
									if ($_.UseNetSend) { $NotifyMethod += 'Net Send' } 

									$_.AlertName + ' (' + [String]::Join(',', $NotifyMethod) + ')'

								} 
							)) 
					} else {
						[String]::Empty
					}

					$WorksheetData[$Row,$Col++] = if (($_.Notifications.Jobs | Measure-Object).Count -gt 0) {
						[String]::Join($Delimiter, @(
								$_.Notifications.Jobs | Where-Object { $_.JobId } | ForEach-Object {

									$NotifyMethod = @()
									if ($_.UseEmail) { $NotifyMethod += 'E-mail' }
									if ($_.UsePager) { $NotifyMethod += 'Pager' }
									if ($_.UseNetSend) { $NotifyMethod += 'Net Send' } 

									$_.JobName + ' (' + [String]::Join(',', $NotifyMethod) + ')'

								} 
							)) 
					} else {
						[String]::Empty
					}

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
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtTime},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtTime},
						@{ColumnNumber = 11; NumberFormat = $XlNumFmtTime},
						@{ColumnNumber = 12; NumberFormat = $XlNumFmtTime},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtTime},
						@{ColumnNumber = 14; NumberFormat = $XlNumFmtTime},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 16; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 17; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion 



			# Apply formatting to every worksheet
			# Work backwards so that the first sheet is active when the workbook is saved
			$ProgressStatus = 'Applying formatting to all worksheets'
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete 0 -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			for ($WorksheetNumber = $WorksheetCount; $WorksheetNumber -ge 1; $WorksheetNumber--) {

				$ProgressStatus = "Applying formatting to Worksheet #$($WorksheetNumber)"
				Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
				Write-Progress -Activity $ProgressActivity -PercentComplete (((($WorksheetCount * 2) - $WorksheetNumber + 1) / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId

				$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)

				# Switch to the worksheet
				$Worksheet.Activate() | Out-Null

				# Bold the header row
				#$Duration = (Measure-Command {
				$Worksheet.Rows.Item(1).Font.Bold = $WorksheetFormat[$WorksheetNumber].BoldFirstRow
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Bold Header Row Duration (ms): $Duration" -MessageLevel Verbose

				# Bold the 1st column
				#$Duration = (Measure-Command {
				$Worksheet.Columns.Item(1).Font.Bold = $WorksheetFormat[$WorksheetNumber].BoldFirstColumn
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Bold 1st Column Duration (ms): $Duration" -MessageLevel Verbose

				# Freeze View
				#$Duration = (Measure-Command {
				$Worksheet.Range($WorksheetFormat[$WorksheetNumber].FreezeAtCell).Select() | Out-Null
				$Worksheet.Application.ActiveWindow.FreezePanes = $true 
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Freeze View Duration (ms): $Duration" -MessageLevel Verbose


				# Apply Column formatting
				#$Duration = (Measure-Command {
				$WorksheetFormat[$WorksheetNumber].ColumnFormat | ForEach-Object {
					$Worksheet.Columns.Item($_.ColumnNumber).NumberFormat = $_.NumberFormat
				}
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Apply Column formatting Duration (ms): $Duration" -MessageLevel Verbose

				# Apply Row formatting
				#$Duration = (Measure-Command {
				$WorksheetFormat[$WorksheetNumber].RowFormat | ForEach-Object {
					$Worksheet.Rows.Item($_.RowNumber).NumberFormat = $_.NumberFormat
				}
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Apply Row formatting Duration (ms): $Duration" -MessageLevel Verbose

				# Update worksheet values so row and column formatting apply
				#$Duration = (Measure-Command {
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
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Apply Row and Column formatting - Update Values (ms): $Duration" -MessageLevel Verbose


				# Apply table formatting
				#$Duration = (Measure-Command {
				$ListObject = $Worksheet.ListObjects.Add($XlListObjectSourceType::xlSrcRange, $Worksheet.UsedRange, $null, $XlYesNoGuess::xlYes, $null) 
				$ListObject.Name = "Table $WorksheetNumber"
				$ListObject.TableStyle = $TableStyle
				$ListObject.ShowTableStyleFirstColumn = $WorksheetFormat[$WorksheetNumber].BoldFirstColumn # Put a background color behind the 1st column
				$ListObject.ShowAutoFilter = $WorksheetFormat[$WorksheetNumber].AutoFilter
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Apply table formatting Duration (ms): $Duration" -MessageLevel Verbose

				# Zoom back to 80%
				#$Duration = (Measure-Command {
				$Worksheet.Application.ActiveWindow.Zoom = 80
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Zoom to 80% Duration (ms): $Duration" -MessageLevel Verbose

				# Adjust the column widths to 250 before autofitting contents
				# This allows longer lines of text to remain on one line
				#$Duration = (Measure-Command {
				$Worksheet.UsedRange.EntireColumn.ColumnWidth = 250
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Change column width Duration (ms): $Duration" -MessageLevel Verbose

				# Autofit column and row contents
				#$Duration = (Measure-Command {
				$Worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null
				$Worksheet.UsedRange.EntireRow.AutoFit() | Out-Null
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Autofit contents Duration (ms): $Duration" -MessageLevel Verbose

				# Left align contents
				#$Duration = (Measure-Command {
				$Worksheet.UsedRange.EntireColumn.HorizontalAlignment = $XlHAlign::xlHAlignLeft
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Left align contents Duration (ms): $Duration" -MessageLevel Verbose

				# Vertical align contents
				#$Duration = (Measure-Command {
				$Worksheet.UsedRange.EntireColumn.VerticalAlignment = $XlVAlign::xlVAlignTop
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Vertical align contents Duration (ms): $Duration" -MessageLevel Verbose

				# Put the selection back to the upper left cell
				#$Duration = (Measure-Command {
				$Worksheet.Range('A1').Select() | Out-Null
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Reset selection Duration (ms): $Duration" -MessageLevel Verbose
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
		Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Information
		Write-SqlServerInventoryLog -Message "End Function: $($MyInvocation.InvocationName)" -MessageLevel Information


		# Cleanup
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

		Remove-Variable -Name TabCharLength
		Remove-Variable -Name IndentString_1
		Remove-Variable -Name IndentString__2
		Remove-Variable -Name IndentString___3

		Remove-Variable -Name ComputerName
		Remove-Variable -Name ServerName
		Remove-Variable -Name ProductName
		Remove-Variable -Name DatabaseName
		Remove-Variable -Name FileGroupName

		Remove-Variable -Name ColorThemePathPattern
		Remove-Variable -Name ColorThemePath

		Remove-Variable -Name WorksheetFormat

		Remove-Variable -Name XlSortOrder
		Remove-Variable -Name XlYesNoGuess
		Remove-Variable -Name XlHAlign
		Remove-Variable -Name XlVAlign
		Remove-Variable -Name XlListObjectSourceType
		Remove-Variable -Name XlThemeColor

		Remove-Variable -Name OverviewTabColor
		Remove-Variable -Name ServicesTabColor
		Remove-Variable -Name ServerTabColor
		Remove-Variable -Name DatabaseTabColor
		Remove-Variable -Name AgentTabColor

		Remove-Variable -Name TableStyle

		Remove-Variable -Name ProgressId
		Remove-Variable -Name ProgressActivity
		Remove-Variable -Name ProgressStatus

		# Release all lingering COM objects
		Remove-ComObject

	}
}

function Export-SqlServerInventoryDatabaseEngineAssessmentToExcel {
	<#
	.SYNOPSIS
		Writes an Excel file containing the Database Engine Assessment information from an assessment of a SQL Server Inventory.

	.DESCRIPTION
		The Export-SqlServerInventoryDatabaseEngineConfigToExcel function uses COM Interop to write an Excel file containing the Database Engine Assessment information from an assessment of a SQL Server Inventory.
				
		Microsoft Excel 2007 or higher must be installed in order to write the Excel file.
		
	.PARAMETER  DatabaseEngineAssessment
		A SQL Server Inventory Database Engine Assessment object returned by Get-SqlServerInventoryDatabaseEngineAssessment.
		
	.PARAMETER  Path
		Specifies the path where the Excel file will be written. This is a fully qualified path to a .XLSX file.
		
		If not specified then the file is named "SQL Server Inventory - [Year][Month][Day][Hour][Minute] - Database Engine Assessment.xlsx" and is written to your "My Documents" folder.

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
		Export-SqlServerInventoryDatabaseEngineAssessmentToExcel -DatabaseEngineAssessment $DbEngineAssessment 
		
		Description
		-----------
		Write a Database Engine inventory using the SQL Server Inventory Database Engine Assessment contained in $DbEngineAssessment.

		The Excel workbook will be written to your "My Documents" folder.
		
		The Office color theme and Medium color scheme will be used by default.
		
	.EXAMPLE
		Export-SqlServerInventoryDatabaseEngineAssessmentToExcel -DatabaseEngineAssessment $DbEngineAssessment -Path 'C:\DB Engine Inventory Assessment.xlsx'
		
		Description
		-----------
		Write a Database Engine inventory using the SQL Server Inventory Database Engine Assessment contained in $DbEngineAssessment.

		The Excel workbook will be written to your C:\DB Engine Inventory.xlsx.
		
		The Office color theme and Medium color scheme will be used by default.

	.EXAMPLE
		Export-SqlServerInventoryDatabaseEngineAssessmentToExcel -DatabaseEngineAssessment $Inventory -ColorTheme Blue -ColorScheme Dark
		
		Description
		-----------
		Write a Database Engine inventory using the SQL Server Inventory Database Engine Assessment contained in $DbEngineAssessment.

		The Excel workbook will be written to your "My Documents" folder.
		
		The Blue color theme and Dark color scheme will be used.
	
	.NOTES
		Blue and Green are nice looking Color Themes for Office 2013

		Waveform is a nice looking Color Theme for Office 2010

	.LINK
		Get-SqlServerInventory
		Get-SqlServerInventoryDatabaseEngineAssessment
#>
	[cmdletBinding()]
	param(
		[Parameter(Mandatory=$true, ValueFromPipeline=$true)]
		[PSCustomObject]
		$DatabaseEngineAssessment
		,
		[Parameter(Mandatory=$false)] 
		[ValidateNotNullOrEmpty()]
		[string]
		$Path = [System.IO.Path]::ChangeExtension((Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)) -ChildPath ('SQL Server Inventory - ' + (Get-Date -Format 'yyyy-MM-dd-HH-mm') + ' - Database Engine Assessment')), 'xlsx')
		,
		[Parameter(Mandatory=$false)] 
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


		<#
	.LINK
		"Export Windows PowerShell Data to Excel" (http://technet.microsoft.com/en-us/query/dd297620)
		
	.LINK
		"Microsoft.Office.Interop.Excel Namespace" (http://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.excel(v=office.14).aspx)
		
	.LINK
		"Excel Object Model Reference" (http://msdn.microsoft.com/en-us/library/ff846392.aspx)
		
	.LINK
		"Excel 2010 Enumerations" (http://msdn.microsoft.com/en-us/library/ff838815.aspx)
		
	.LINK
		"TableStyle Cheat Sheet in French/English" (http://msdn.microsoft.com/fr-fr/library/documentformat.openxml.spreadsheet.tablestyle.aspx)
		
	.LINK
		"Color Palette and the 56 Excel ColorIndex Colors" (http://dmcritchie.mvps.org/excel/colors.htm)
		
	.LINK
		"Adding Color to Excel 2007 Worksheets by Using the ColorIndex Property" (http://msdn.microsoft.com/en-us/library/cc296089(v=office.12).aspx)
		
	.LINK
		"XlRgbColor Enumeration" (http://msdn.microsoft.com/en-us/library/ff197459.aspx)
#>



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

		$TabCharLength = 4
		$IndentString_1 = [String]::Empty.PadLeft($TabCharLength * 1)
		$IndentString__2 = [String]::Empty.PadLeft($TabCharLength * 2)
		$IndentString___3 = [String]::Empty.PadLeft($TabCharLength * 3)

		$ComputerName = $null
		$ServerName = $null
		$ProductName = $null
		$DatabaseName = $null
		$FileGroupName = $null

		$ColorThemePathPattern = $null
		$ColorThemePath = $null

		# Used to hold all of the formatting to be applied at the end
		$WorksheetCount = 6
		$WorksheetFormat = @{}

		$XlSortOrder = 'Microsoft.Office.Interop.Excel.XlSortOrder' -as [Type]
		$XlYesNoGuess = 'Microsoft.Office.Interop.Excel.XlYesNoGuess' -as [Type]
		$XlHAlign = 'Microsoft.Office.Interop.Excel.XlHAlign' -as [Type]
		$XlVAlign = 'Microsoft.Office.Interop.Excel.XlVAlign' -as [Type]
		$XlListObjectSourceType = 'Microsoft.Office.Interop.Excel.XlListObjectSourceType' -as [Type]
		$XlThemeColor = 'Microsoft.Office.Interop.Excel.XlThemeColor' -as [Type]

		$OverviewTabColor = $XlThemeColor::xlThemeColorDark1
		$ServicesTabColor = $XlThemeColor::xlThemeColorLight1
		$ServerTabColor = $XlThemeColor::xlThemeColorAccent1
		$DatabaseTabColor = $XlThemeColor::xlThemeColorAccent2
		$SecurityTabColor = $XlThemeColor::xlThemeColorAccent3
		$ServerObjectsTabColor = $XlThemeColor::xlThemeColorAccent4
		$ManagementTabColor = $XlThemeColor::xlThemeColorAccent5
		$AgentTabColor = $XlThemeColor::xlThemeColorAccent6

		$TableStyle = switch ($ColorScheme) {
			'light' { 'TableStyleLight8' }
			'medium' { 'TableStyleMedium15' }
			'dark' { 'TableStyleDark1' }
		}

		$ProgressId = Get-Random
		$ProgressActivity = 'Export-SqlServerInventoryConfigurationCheckToExcel'
		$ProgressStatus = 'Beginning output to Excel'

		Write-SqlServerInventoryLog -Message "Start Function: $($MyInvocation.InvocationName)" -MessageLevel Debug
		Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Information
		Write-Progress -Activity $ProgressActivity -PercentComplete 0 -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId


		#region

		# Hide the Excel instance (is this necessary?)
		$Excel.visible = $false

		# Turn off screen updating
		$Excel.ScreenUpdating = $false

		# Turn off automatic calculations
		#$Excel.Calculation = [Microsoft.Office.Interop.Excel.XlCalculation]::xlCalculationManual

		# Add a workbook
		$Workbook = $Excel.Workbooks.Add()
		$Workbook.Title = 'SQL Server Inventory - Database Engine Configuration Report'

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
				Write-SqlServerInventoryLog -Message "Unable to find a theme named ""$ColorTheme"", using default Excel theme instead" -MessageLevel Warning
			}

		}


		# Add enough worksheets to get us to $WorksheetCount
		$Excel.Worksheets.Add($MissingType, $Excel.Worksheets.Item($Excel.Worksheets.Count), $WorksheetCount - $Excel.Worksheets.Count, $Excel.Worksheets.Item(1).Type) | Out-Null
		$WorksheetNumber = 1

		try {

			# Write each worksheet
			@($CatPerformance, $CatReliability, $CatSecurity, $CatAvailability, $CatRecovery, $CatInformation) | ForEach-Object {

				$Category = $_

				$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): $Category"
				Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
				Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
				#region
				$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
				$Worksheet.Name = $Category
				$Worksheet.Tab.ThemeColor = $ServerTabColor

				$RowCount = ($DatabaseEngineAssessment | Where-Object { $_.Category -ieq $Category } | Measure-Object).Count + 1
				$ColumnCount = 6
				$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

				$Col = 0
				$WorksheetData[0,$Col++] = 'Server Name'
				$WorksheetData[0,$Col++] = 'Database Name'
				$WorksheetData[0,$Col++] = 'Priority'
				$WorksheetData[0,$Col++] = 'Description'
				$WorksheetData[0,$Col++] = 'Details'
				$WorksheetData[0,$Col++] = 'URL'

				$Row = 1
				$DatabaseEngineAssessment | 
				Where-Object { $_.Category -ieq $Category } | 
				Sort-Object -Property Priority, Description, ServerName, DatabaseName |
				ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $_.ServerName
					$WorksheetData[$Row,$Col++] = $_.DatabaseName
					$WorksheetData[$Row,$Col++] = $_.PriorityValue
					$WorksheetData[$Row,$Col++] = $_.Description
					$WorksheetData[$Row,$Col++] = $_.Details
					$WorksheetData[$Row,$Col++] = $_.URL
					$Row++
				}

				$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
				$Range.Value2 = $WorksheetData
				#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
				#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
				#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

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


			}


			# Apply formatting to every worksheet
			# Work backwards so that the first sheet is active when the workbook is saved
			$ProgressStatus = 'Applying formatting to all worksheets'
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete 0 -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			for ($WorksheetNumber = $WorksheetCount; $WorksheetNumber -ge 1; $WorksheetNumber--) {

				$ProgressStatus = "Applying formatting to Worksheet #$($WorksheetNumber)"
				Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
				Write-Progress -Activity $ProgressActivity -PercentComplete (((($WorksheetCount * 2) - $WorksheetNumber + 1) / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId

				$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)

				# Switch to the worksheet
				$Worksheet.Activate() | Out-Null

				# Bold the header row
				#$Duration = (Measure-Command {
				$Worksheet.Rows.Item(1).Font.Bold = $WorksheetFormat[$WorksheetNumber].BoldFirstRow
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Bold Header Row Duration (ms): $Duration" -MessageLevel Verbose

				# Bold the 1st column
				#$Duration = (Measure-Command {
				$Worksheet.Columns.Item(1).Font.Bold = $WorksheetFormat[$WorksheetNumber].BoldFirstColumn
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Bold 1st Column Duration (ms): $Duration" -MessageLevel Verbose

				# Freeze View
				#$Duration = (Measure-Command {
				$Worksheet.Range($WorksheetFormat[$WorksheetNumber].FreezeAtCell).Select() | Out-Null
				$Worksheet.Application.ActiveWindow.FreezePanes = $true 
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Freeze View Duration (ms): $Duration" -MessageLevel Verbose


				# Apply Column formatting
				#$Duration = (Measure-Command {
				$WorksheetFormat[$WorksheetNumber].ColumnFormat | ForEach-Object {
					$Worksheet.Columns.Item($_.ColumnNumber).NumberFormat = $_.NumberFormat
				}
				#}).TotalMilliseconds
				Write-SqlServerInventoryLog -Message "Apply Column formatting Duration (ms): $Duration" -MessageLevel Verbose

				# Apply Row formatting
				#$Duration = (Measure-Command {
				$WorksheetFormat[$WorksheetNumber].RowFormat | ForEach-Object {
					$Worksheet.Rows.Item($_.RowNumber).NumberFormat = $_.NumberFormat
				}
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Apply Row formatting Duration (ms): $Duration" -MessageLevel Verbose

				# Update worksheet values so row and column formatting apply
				#$Duration = (Measure-Command {
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
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Apply Row and Column formatting - Update Values (ms): $Duration" -MessageLevel Verbose


				# Apply table formatting
				#$Duration = (Measure-Command {
				$ListObject = $Worksheet.ListObjects.Add($XlListObjectSourceType::xlSrcRange, $Worksheet.UsedRange, $null, $XlYesNoGuess::xlYes, $null) 
				$ListObject.Name = "Table $WorksheetNumber"
				$ListObject.TableStyle = $TableStyle
				$ListObject.ShowTableStyleFirstColumn = $WorksheetFormat[$WorksheetNumber].BoldFirstColumn # Put a background color behind the 1st column
				$ListObject.ShowAutoFilter = $WorksheetFormat[$WorksheetNumber].AutoFilter
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Apply table formatting Duration (ms): $Duration" -MessageLevel Verbose

				# Zoom back to 80%
				#$Duration = (Measure-Command {
				$Worksheet.Application.ActiveWindow.Zoom = 80
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Zoom to 80% Duration (ms): $Duration" -MessageLevel Verbose

				# Adjust the column widths to 150 before autofitting contents
				# This allows longer lines of text to remain on one line
				#$Duration = (Measure-Command {
				$Worksheet.UsedRange.EntireColumn.ColumnWidth = 150
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Change column width Duration (ms): $Duration" -MessageLevel Verbose

				# Wrap text
				#$Duration = (Measure-Command {
				$Worksheet.UsedRange.WrapText = $true
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Wrap text Duration (ms): $Duration" -MessageLevel Verbose

				# Autofit column and row contents
				#$Duration = (Measure-Command {
				$Worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null
				$Worksheet.UsedRange.EntireRow.AutoFit() | Out-Null
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Autofit contents Duration (ms): $Duration" -MessageLevel Verbose

				# Left align contents
				#$Duration = (Measure-Command {
				$Worksheet.UsedRange.EntireColumn.HorizontalAlignment = $XlHAlign::xlHAlignLeft
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Left align contents Duration (ms): $Duration" -MessageLevel Verbose

				# Vertical align contents
				#$Duration = (Measure-Command {
				$Worksheet.UsedRange.EntireColumn.VerticalAlignment = $XlVAlign::xlVAlignTop
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Vertical align contents Duration (ms): $Duration" -MessageLevel Verbose

				# Put the selection back to the upper left cell
				#$Duration = (Measure-Command {
				$Worksheet.Range('A1').Select() | Out-Null
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Reset selection Duration (ms): $Duration" -MessageLevel Verbose
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
		Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Information
		Write-SqlServerInventoryLog -Message "End Function: $($MyInvocation.InvocationName)" -MessageLevel Information


		# Cleanup
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

		Remove-Variable -Name TabCharLength
		Remove-Variable -Name IndentString_1
		Remove-Variable -Name IndentString__2
		Remove-Variable -Name IndentString___3

		Remove-Variable -Name ComputerName
		Remove-Variable -Name ServerName
		Remove-Variable -Name ProductName
		Remove-Variable -Name DatabaseName
		Remove-Variable -Name FileGroupName

		Remove-Variable -Name ColorThemePathPattern
		Remove-Variable -Name ColorThemePath

		Remove-Variable -Name WorksheetFormat

		Remove-Variable -Name XlSortOrder
		Remove-Variable -Name XlYesNoGuess
		Remove-Variable -Name XlHAlign
		Remove-Variable -Name XlVAlign
		Remove-Variable -Name XlListObjectSourceType
		Remove-Variable -Name XlThemeColor

		Remove-Variable -Name OverviewTabColor
		Remove-Variable -Name ServicesTabColor
		Remove-Variable -Name ServerTabColor
		Remove-Variable -Name DatabaseTabColor
		Remove-Variable -Name AgentTabColor

		Remove-Variable -Name TableStyle

		Remove-Variable -Name ProgressId
		Remove-Variable -Name ProgressActivity
		Remove-Variable -Name ProgressStatus

		# Release all lingering COM objects
		Remove-ComObject

	}
}

function Export-SqlServerInventoryDatabaseEngineDbObjectsToExcel {
	<#
	.SYNOPSIS
		Writes an Excel file containing the Database Engine Database Object information in a SQL Server Inventory.

	.DESCRIPTION
		The Export-SqlServerInventoryDatabaseEngineDbObjectsToExcel function uses COM Interop to write an Excel file containing the Database Engine Database Objects information in a SQL Server Inventory returned by Get-SqlServerInventory.
		
		Although the SQL Server Shared Management Objects (SMO) libraries are required to perform an inventory they are NOT required to write the Excel file.
		
		Microsoft Excel 2007 or higher must be installed in order to write the Excel file.
		
	.PARAMETER  SqlServerInventory
		A SQL Server Inventory object returned by Get-SqlServerInventory.
		
	.PARAMETER  Path
		Specifies the path where the Excel file will be written. This is a fully qualified path to a .XLSX file.
		
		If not specified then the file is named "SQL Server Inventory - [Year][Month][Day][Hour][Minute] - Database Engine Db Objects.xlsx" and is written to your "My Documents" folder.

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
		Export-SqlServerInventoryDatabaseEngineDbObjectsToExcel -SqlServerInventory $Inventory 
		
		Description
		-----------
		Write a Database Engine inventory using the SQL Server Inventory contained in $Inventory.

		The Excel workbook will be written to your "My Documents" folder.
		
		The Office color theme and Medium color scheme will be used by default.
		
	.EXAMPLE
		Export-SqlServerInventoryDatabaseEngineDbObjectsToExcel -SqlServerInventory $Inventory -Path 'C:\DB Engine Inventory.xlsx'
		
		Description
		-----------
		Write a Database Engine inventory using the SQL Server Inventory contained in $Inventory.

		The Excel workbook will be written to your C:\DB Engine Inventory.xlsx.
		
		The Office color theme and Medium color scheme will be used by default.

	.EXAMPLE
		Export-SqlServerInventoryDatabaseEngineDbObjectsToExcel -SqlServerInventory $Inventory -ColorTheme Blue -ColorScheme Dark
		
		Description
		-----------
		Write a Database Engine inventory using the SQL Server Inventory contained in $Inventory.

		The Excel workbook will be written to your "My Documents" folder.
		
		The Blue color theme and Dark color scheme will be used.
	
	.NOTES
		Blue and Green are nice looking Color Themes for Office 2013

		Waveform is a nice looking Color Theme for Office 2010

	.LINK
		Get-SqlServerInventory

#>
	[cmdletBinding()]
	param(
		[Parameter(Mandatory=$true, ValueFromPipeline=$true)]
		[PSCustomObject]
		$SqlServerInventory
		,
		[Parameter(Mandatory=$false)] 
		[ValidateNotNullOrEmpty()]
		[string]
		$Path = [System.IO.Path]::ChangeExtension((Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)) -ChildPath ('SQL Server Inventory - ' + (Get-Date -Format 'yyyy-MM-dd-HH-mm') + ' - Database Engine Db Objects')), 'xlsx')
		,
		[Parameter(Mandatory=$false)] 
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

		<#
	.LINK
		"Export Windows PowerShell Data to Excel" (http://technet.microsoft.com/en-us/query/dd297620)
		
	.LINK
		"Microsoft.Office.Interop.Excel Namespace" (http://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.excel(v=office.14).aspx)
		
	.LINK
		"Excel Object Model Reference" (http://msdn.microsoft.com/en-us/library/ff846392.aspx)
		
	.LINK
		"Excel 2010 Enumerations" (http://msdn.microsoft.com/en-us/library/ff838815.aspx)
		
	.LINK
		"TableStyle Cheat Sheet in French/English" (http://msdn.microsoft.com/fr-fr/library/documentformat.openxml.spreadsheet.tablestyle.aspx)
		
	.LINK
		"Color Palette and the 56 Excel ColorIndex Colors" (http://dmcritchie.mvps.org/excel/colors.htm)
		
	.LINK
		"Adding Color to Excel 2007 Worksheets by Using the ColorIndex Property" (http://msdn.microsoft.com/en-us/library/cc296089(v=office.12).aspx)
		
	.LINK
		"XlRgbColor Enumeration" (http://msdn.microsoft.com/en-us/library/ff197459.aspx)
#>

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

		$TabCharLength = 4
		$IndentString_1 = [String]::Empty.PadLeft($TabCharLength * 1)
		$IndentString__2 = [String]::Empty.PadLeft($TabCharLength * 2)
		$IndentString___3 = [String]::Empty.PadLeft($TabCharLength * 3)

		$ComputerName = $null
		$ServerName = $null
		$ProductName = $null
		$DatabaseName = $null
		$FileGroupName = $null

		$ColorThemePathPattern = $null
		$ColorThemePath = $null

		# Used to hold all of the formatting to be applied at the end
		$WorksheetCount = 74
		$WorksheetFormat = @{}

		$XlSortOrder = 'Microsoft.Office.Interop.Excel.XlSortOrder' -as [Type]
		$XlYesNoGuess = 'Microsoft.Office.Interop.Excel.XlYesNoGuess' -as [Type]
		$XlHAlign = 'Microsoft.Office.Interop.Excel.XlHAlign' -as [Type]
		$XlVAlign = 'Microsoft.Office.Interop.Excel.XlVAlign' -as [Type]
		$XlListObjectSourceType = 'Microsoft.Office.Interop.Excel.XlListObjectSourceType' -as [Type]
		$XlThemeColor = 'Microsoft.Office.Interop.Excel.XlThemeColor' -as [Type]

		$OverviewTabColor = $XlThemeColor::xlThemeColorDark1
		$ServicesTabColor = $XlThemeColor::xlThemeColorLight1
		$ServerTabColor = $XlThemeColor::xlThemeColorAccent1
		$DatabaseTabColor = $XlThemeColor::xlThemeColorAccent2
		$SecurityTabColor = $XlThemeColor::xlThemeColorAccent3
		$ServerObjectsTabColor = $XlThemeColor::xlThemeColorAccent4
		$ManagementTabColor = $XlThemeColor::xlThemeColorAccent5
		$AgentTabColor = $XlThemeColor::xlThemeColorAccent6

		$TableStyle = switch ($ColorScheme) {
			'light' { 'TableStyleLight8' }
			'medium' { 'TableStyleMedium15' }
			'dark' { 'TableStyleDark1' }
		}

		$ProgressId = Get-Random
		$ProgressActivity = 'Export-SqlServerInventoryDatabaseEngineDbObjectsToExcel'
		$ProgressStatus = 'Beginning output to Excel'

		Write-SqlServerInventoryLog -Message "Start Function: $($MyInvocation.InvocationName)" -MessageLevel Debug
		Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Information
		Write-Progress -Activity $ProgressActivity -PercentComplete 0 -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId


		#region

		# Hide the Excel instance (is this necessary?)
		$Excel.visible = $false

		# Turn off screen updating
		$Excel.ScreenUpdating = $false

		# Turn off automatic calculations
		#$Excel.Calculation = [Microsoft.Office.Interop.Excel.XlCalculation]::xlCalculationManual

		# Add a workbook
		$Workbook = $Excel.Workbooks.Add()
		$Workbook.Title = 'SQL Server Inventory - Database Engine Database Objects'

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
				Write-SqlServerInventoryLog -Message "Unable to find a theme named ""$ColorTheme"", using default Excel theme instead" -MessageLevel Warning
			}

		}


		# Add enough worksheets to get us to $WorksheetCount
		$Excel.Worksheets.Add($MissingType, $Excel.Worksheets.Item($Excel.Worksheets.Count), $WorksheetCount - $Excel.Worksheets.Count, $Excel.Worksheets.Item(1).Type) | Out-Null
		$WorksheetNumber = 1

		try {


			# Worksheet 1: Services
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Services"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Services'
			#$Worksheet.Tab.Color = $ServicesTabColor
			$Worksheet.Tab.ThemeColor = $ServicesTabColor

			$RowCount = ($SqlServerInventory.Service | Measure-Object).Count + 1
			$ColumnCount = 15
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Computer Name'
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Service Type'
			$WorksheetData[0,$Col++] = 'Service IP Address'
			$WorksheetData[0,$Col++] = 'Service Port'
			$WorksheetData[0,$Col++] = 'Status'
			$WorksheetData[0,$Col++] = 'Process ID'
			$WorksheetData[0,$Col++] = 'Start Date'
			$WorksheetData[0,$Col++] = 'Start Mode'
			$WorksheetData[0,$Col++] = 'Service Account'
			$WorksheetData[0,$Col++] = 'Clustered'
			$WorksheetData[0,$Col++] = 'AlwaysOn'
			$WorksheetData[0,$Col++] = 'Executable Path'
			$WorksheetData[0,$Col++] = 'Startup Parameters'

			$Row = 1
			$SqlServerInventory.Service | ForEach-Object {
				$Col = 0
				$WorksheetData[$Row,$Col++] = $_.ComputerName
				$WorksheetData[$Row,$Col++] = $_.ServerName
				$WorksheetData[$Row,$Col++] = $_.DisplayName #$_.ServiceTypeName
				$WorksheetData[$Row,$Col++] = $_.ServiceIpAddress
				$WorksheetData[$Row,$Col++] = $_.Port
				$WorksheetData[$Row,$Col++] = $_.ServiceState
				$WorksheetData[$Row,$Col++] = $_.ProcessId
				$WorksheetData[$Row,$Col++] = $_.ServiceStartDate
				$WorksheetData[$Row,$Col++] = $_.StartMode
				$WorksheetData[$Row,$Col++] = $_.ServiceAccount
				$WorksheetData[$Row,$Col++] = $_.IsClusteredInstance
				$WorksheetData[$Row,$Col++] = $_.IsHadrEnabled
				$WorksheetData[$Row,$Col++] = $_.PathName
				$WorksheetData[$Row,$Col++] = $_.StartupParameters
				$Row++
			}

			#$Worksheet.Columns.Item(5).NumberFormat = $XlNumFmtNumberGeneral
			#$Worksheet.Columns.Item(7).NumberFormat = $XlNumFmtNumberGeneral

			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'D2'
					ColumnFormat = @(
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 2: Server Overview - Servername, Scan Date, Version, Edition, OS, etc.
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Server Overview"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Server Overview'
			#$Worksheet.Tab.Color = $OverviewTabColor
			$Worksheet.Tab.ThemeColor = $OverviewTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | Measure-Object).Count + 1
			$ColumnCount = 19
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Computer Name'
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Scan Date (UTC)'
			$WorksheetData[0,$Col++] = 'Install Date'
			$WorksheetData[0,$Col++] = 'Startup Date'
			$WorksheetData[0,$Col++] = 'Product Name'
			$WorksheetData[0,$Col++] = 'Product Edition'
			$WorksheetData[0,$Col++] = 'Level'
			$WorksheetData[0,$Col++] = 'Platform'
			$WorksheetData[0,$Col++] = 'Version'
			$WorksheetData[0,$Col++] = 'Server Type'
			$WorksheetData[0,$Col++] = 'Clustered'
			$WorksheetData[0,$Col++] = 'Logical Processors'
			$WorksheetData[0,$Col++] = 'Total Memory (MB)'
			$WorksheetData[0,$Col++] = 'Instance Memory In Use (MB)'
			$WorksheetData[0,$Col++] = 'Operating System'
			$WorksheetData[0,$Col++] = 'System Manufacturer'
			$WorksheetData[0,$Col++] = 'System Type'
			$WorksheetData[0,$Col++] = 'Power Plan'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {
				$Col = 0
				$WorksheetData[$Row,$Col++] = $_.Machine.OperatingSystem.Settings.ComputerSystem.FullyQualifiedDomainName # $_.ComputerName
				$WorksheetData[$Row,$Col++] = $_.ServerName
				$WorksheetData[$Row,$Col++] = $_.ScanDateUTC
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.InstallDate
				$WorksheetData[$Row,$Col++] = $_.Server.Service.ServiceStartDate
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Product
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Edition
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.ProductLevel
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Platform
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.Version
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.ServerType
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.IsClustered
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.ProcessorCount
				$WorksheetData[$Row,$Col++] = $_.Server.Configuration.General.MemoryMB
				$WorksheetData[$Row,$Col++] = "{0:N2}" -f ($_.Server.Configuration.General.MemoryInUseKB / 1KB)
				$WorksheetData[$Row,$Col++] = $_.Machine.OperatingSystem.Settings.OperatingSystem.Name
				$WorksheetData[$Row,$Col++] = $_.Machine.OperatingSystem.Settings.ComputerSystemProduct.Manufacturer
				$WorksheetData[$Row,$Col++] = $_.Machine.OperatingSystem.Settings.ComputerSystemProduct.Name
				$WorksheetData[$Row,$Col++] = $_.Machine.OperatingSystem.Settings.PowerPlan | Where-Object { $_.IsActive -eq $true } | ForEach-Object { $_.PlanName }
				$Row++
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'C2'
					ColumnFormat = @(
						@{ColumnNumber = 3; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 14; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtNumberS2}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 3: Database Overview
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Overview"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Database Overview'
			#$Worksheet.Tab.Color = $OverviewTabColor
			$Worksheet.Tab.ThemeColor = $OverviewTabColor

			$RowCount = (($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases }) | Measure-Object).Count + 1
			$ColumnCount = 17
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Product Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Date Created'
			$WorksheetData[0,$Col++] = 'Status'
			$WorksheetData[0,$Col++] = 'Owner'
			$WorksheetData[0,$Col++] = 'Compatibility Level'
			$WorksheetData[0,$Col++] = 'Recovery Model'
			$WorksheetData[0,$Col++] = 'Data File Count'
			$WorksheetData[0,$Col++] = 'Log File Count'
			$WorksheetData[0,$Col++] = 'Data File Size (MB)'
			$WorksheetData[0,$Col++] = 'Log File Size (MB)'
			$WorksheetData[0,$Col++] = 'Available Data Space (MB)'
			$WorksheetData[0,$Col++] = 'Last Known DBCC Date'
			$WorksheetData[0,$Col++] = 'Last Full Backup'
			$WorksheetData[0,$Col++] = 'Last Diff Backup'
			$WorksheetData[0,$Col++] = 'Last Log Backup'

			$Row = 1
			$SqlServerInventory.DatabaseServer | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name
				$ProductName = $_.Server.Configuration.General.Product

				$_.Server.Databases | ForEach-Object {
					$Col = 0
					$WorksheetData[$Row,$Col++] = $ServerName
					$WorksheetData[$Row,$Col++] = $ProductName
					$WorksheetData[$Row,$Col++] = $_.Name
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.DateCreated
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.Status
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.Owner
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.CompatibilityLevel
					$WorksheetData[$Row,$Col++] = $_.Properties.Options.RecoveryModel
					$WorksheetData[$Row,$Col++] = ($_.Properties.Files.DatabaseFiles | Where-Object { $_.FileType -ieq 'rows data' } | Measure-Object).Count
					$WorksheetData[$Row,$Col++] = ($_.Properties.Files.DatabaseFiles | Where-Object { $_.FileType -ieq 'Log' } | Measure-Object).Count
					$WorksheetData[$Row,$Col++] = "{0:N2}" -f (($_.Properties.Files.DatabaseFiles | Where-Object { $_.FileType -ieq 'rows data' } | Measure-Object -Property SizeKB -Sum).Sum / 1KB)
					$WorksheetData[$Row,$Col++] = "{0:N2}" -f (($_.Properties.Files.DatabaseFiles | Where-Object { $_.FileType -ieq 'log' } | Measure-Object -Property SizeKB -Sum).Sum / 1KB)
					$WorksheetData[$Row,$Col++] = "{0:N2}" -f (($_.Properties.Files.DatabaseFiles | Where-Object { $_.FileType -ieq 'rows data' } | Measure-Object -Property AvailableSpaceKB -Sum).Sum / 1KB)
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Database.LastKnownGoodDbccDate
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Backup.LastFullBackupDate
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Backup.LastDifferentialBackupDate
					$WorksheetData[$Row,$Col++] = $_.Properties.General.Backup.LastLogBackupDate
					$Row++
				}

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType,S $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'D2'
					ColumnFormat = @(
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 11; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 12; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 14; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 16; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 17; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 4: Tables - General
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tables - General"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tables - General'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Tables | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 27
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Description >>'
			$WorksheetData[0,$Col++] = 'Created Date'
			$WorksheetData[0,$Col++] = 'Last Modified Date'
			$WorksheetData[0,$Col++] = 'System Object'
			$WorksheetData[0,$Col++] = 'Has After Trigger'
			$WorksheetData[0,$Col++] = 'Has Clustered Index'
			$WorksheetData[0,$Col++] = 'Has Compressed Partitions'
			$WorksheetData[0,$Col++] = 'Has Delete Trigger'
			$WorksheetData[0,$Col++] = 'Has Index'
			$WorksheetData[0,$Col++] = 'Has Insert Trigger'
			$WorksheetData[0,$Col++] = 'Has InsteadOf Trigger'
			$WorksheetData[0,$Col++] = 'Has Update Trigger'
			$WorksheetData[0,$Col++] = 'Is Indexable'
			$WorksheetData[0,$Col++] = 'Is Schema Owned'
			$WorksheetData[0,$Col++] = 'Options >>'
			$WorksheetData[0,$Col++] = 'Quoted Identifier'
			$WorksheetData[0,$Col++] = 'ANSI NULLs'
			$WorksheetData[0,$Col++] = 'Is File Table'
			$WorksheetData[0,$Col++] = 'Lock Escalation'
			$WorksheetData[0,$Col++] = 'Max Degree Of Parallelism'
			$WorksheetData[0,$Col++] = 'Online Heap Operation'
			$WorksheetData[0,$Col++] = 'Replication >>'
			$WorksheetData[0,$Col++] = 'Table Is Replicated'


			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Tables | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Schema
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Name
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.CreateDate
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.DateLastModified
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.IsSystemObject
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.HasAfterTrigger
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.HasClusteredIndex
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.HasCompressedPartitions
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.HasDeleteTrigger
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.HasIndex
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.HasInsertTrigger
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.HasInsteadOfTrigger
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.HasUpdateTrigger
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.IsIndexable
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.IsSchemaOwned
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.QuotedIdentifier
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.AnsiNulls
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.IsFileTable
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.LockEscalation
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.MaximumDegreeOfParallelism
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.OnlineHeapOperation
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.IsReplicated
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 5: Tables - Change Tracking
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tables - Change Tracking"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tables - Change Tracking'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Tables | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 6
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Change Tracking Enabled'
			$WorksheetData[0,$Col++] = 'Track Columns Updated'


			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Tables | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Schema
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Name
						$WorksheetData[$Row,$Col++] = $_.Properties.ChangeTracking.IsEnabled
						$WorksheetData[$Row,$Col++] = $_.Properties.ChangeTracking.TrackColumnsUpdated
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 6: Tables - Storage
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tables - Storage"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tables - Storage'
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Tables | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 23
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Compression >>'
			$WorksheetData[0,$Col++] = 'Partitions Not Compressed'
			$WorksheetData[0,$Col++] = 'Partitions Using Row Compression'
			$WorksheetData[0,$Col++] = 'Partitions Using Page Compression'
			$WorksheetData[0,$Col++] = 'Filegroups >>'
			$WorksheetData[0,$Col++] = 'Text Filegroup'
			$WorksheetData[0,$Col++] = 'Table Is Partitioned'
			$WorksheetData[0,$Col++] = 'Filegroup'
			$WorksheetData[0,$Col++] = 'FILESTREAM Filegroup'
			$WorksheetData[0,$Col++] = 'General >>'
			$WorksheetData[0,$Col++] = 'Vardecimal Storage Format Enabled'
			$WorksheetData[0,$Col++] = 'Index Space (MB)'
			$WorksheetData[0,$Col++] = 'Row Count'
			$WorksheetData[0,$Col++] = 'Data Space (MB)'
			$WorksheetData[0,$Col++] = 'Partitioning >>'
			$WorksheetData[0,$Col++] = 'Partition Scheme'
			$WorksheetData[0,$Col++] = 'Number of Partitions'
			$WorksheetData[0,$Col++] = 'Partition Column'
			$WorksheetData[0,$Col++] = 'FILESTREAM Partition Scheme'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Tables | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Schema
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Name
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.Storage.Compression.PartitionsNotCompressed
						$WorksheetData[$Row,$Col++] = $_.Properties.Storage.Compression.PartitionsUsingRowCompression
						$WorksheetData[$Row,$Col++] = $_.Properties.Storage.Compression.PartitionsUsingPageCompression
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.Storage.Filegroups.TextFileGroup
						$WorksheetData[$Row,$Col++] = $_.Properties.Storage.Filegroups.IsPartitioned
						$WorksheetData[$Row,$Col++] = $_.Properties.Storage.Filegroups.FileGroup
						$WorksheetData[$Row,$Col++] = $_.Properties.Storage.Filegroups.FileStreamFileGroup
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.Storage.General.IsVarDecimalStorageFormatEnabled
						$WorksheetData[$Row,$Col++] = $_.Properties.Storage.General.IndexSpaceUsedKB / 1KB
						$WorksheetData[$Row,$Col++] = $_.Properties.Storage.General.RowCount
						$WorksheetData[$Row,$Col++] = $_.Properties.Storage.General.DataSpaceUsedKB / 1KB
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.Storage.Partitioning.PartitionScheme
						$WorksheetData[$Row,$Col++] = $($_.Properties.Storage.Partitioning.PhysicalPartitions | Measure-Object).Count
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.Storage.Partitioning.FileStreamPartitionScheme

						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 16; NumberFormat = $XlNumFmtNumberS3},
						@{ColumnNumber = 17; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 18; NumberFormat = $XlNumFmtNumberS3},
						@{ColumnNumber = 21; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 7: Tables - Columns
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tables - Columns"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tables - Columns'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Tables | ForEach-Object { $_.Columns | Where-Object { $_.General.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 44
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Column Name'
			$WorksheetData[0,$Col++] = 'Binding >>'
			$WorksheetData[0,$Col++] = 'Default Binding'
			$WorksheetData[0,$Col++] = 'Default Schema'
			$WorksheetData[0,$Col++] = 'Rule'
			$WorksheetData[0,$Col++] = 'Rule Schema'
			$WorksheetData[0,$Col++] = 'Computed >>'
			$WorksheetData[0,$Col++] = 'Is Computed'
			$WorksheetData[0,$Col++] = 'Computed Text'
			$WorksheetData[0,$Col++] = 'General >>'
			$WorksheetData[0,$Col++] = 'Allow Nulls'
			$WorksheetData[0,$Col++] = 'ANSI Padding Status'
			$WorksheetData[0,$Col++] = 'Column ID'
			$WorksheetData[0,$Col++] = 'Data Type'
			$WorksheetData[0,$Col++] = 'Length'
			$WorksheetData[0,$Col++] = 'Numeric Precision'
			$WorksheetData[0,$Col++] = 'Numeric Scale'
			$WorksheetData[0,$Col++] = 'Primary Key'
			$WorksheetData[0,$Col++] = 'System Type'
			$WorksheetData[0,$Col++] = 'Identity >>'
			$WorksheetData[0,$Col++] = 'Identity'
			$WorksheetData[0,$Col++] = 'Identity Seed'
			$WorksheetData[0,$Col++] = 'Identity Increment'
			$WorksheetData[0,$Col++] = 'Miscellaneous >>'
			$WorksheetData[0,$Col++] = 'Collation'
			$WorksheetData[0,$Col++] = 'Full Text'
			$WorksheetData[0,$Col++] = 'Not For Replication'
			$WorksheetData[0,$Col++] = 'Statistical Semantics'
			$WorksheetData[0,$Col++] = 'Is Deterministic'
			$WorksheetData[0,$Col++] = 'Is FILESTREAM'
			$WorksheetData[0,$Col++] = 'Is Foreign Key'
			$WorksheetData[0,$Col++] = 'Is Persisted'
			$WorksheetData[0,$Col++] = 'Is Precise'
			$WorksheetData[0,$Col++] = 'Is ROWGUID'
			$WorksheetData[0,$Col++] = 'Sparse >>'
			$WorksheetData[0,$Col++] = 'Is Column Set'
			$WorksheetData[0,$Col++] = 'Is Sparse'
			$WorksheetData[0,$Col++] = 'XML >>'
			$WorksheetData[0,$Col++] = 'XML Schema Namespace'
			$WorksheetData[0,$Col++] = 'XML Schema Namespace Schema'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Tables | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Columns | Where-Object { $_.General.General.ID } | Sort-Object -Property @{Expression={$_.General.General.ID}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.General.Name
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Binding.DefaultBinding
							$WorksheetData[$Row,$Col++] = $_.General.Binding.DefaultSchema
							$WorksheetData[$Row,$Col++] = $_.General.Binding.Rule
							$WorksheetData[$Row,$Col++] = $_.General.Binding.RuleSchema
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Computed.IsComputed
							$WorksheetData[$Row,$Col++] = $_.General.Computed.ComputedText
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.General.AllowNulls
							$WorksheetData[$Row,$Col++] = $_.General.General.AnsiPaddingStatus
							$WorksheetData[$Row,$Col++] = $_.General.General.ID
							$WorksheetData[$Row,$Col++] = $_.General.General.DataType
							$WorksheetData[$Row,$Col++] = $_.General.General.Length
							$WorksheetData[$Row,$Col++] = $_.General.General.NumericPrecision
							$WorksheetData[$Row,$Col++] = $_.General.General.NumericScale
							$WorksheetData[$Row,$Col++] = $_.General.General.InPrimaryKey
							$WorksheetData[$Row,$Col++] = $_.General.General.SystemType
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Identity.IsIdentity
							$WorksheetData[$Row,$Col++] = $_.General.Identity.IdentityIncrement
							$WorksheetData[$Row,$Col++] = $_.General.Identity.IdentitySeed
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.Collation
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsFullTextIndexed
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsNotForReplication
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.StatisticalSemantics
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsDeterministic
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsFileStream
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsForeignKey
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsPersisted
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsPrecise
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsRowGuidCol
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Sparse.IsColumnSet
							$WorksheetData[$Row,$Col++] = $_.General.Sparse.IsSparse
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.XML.XmlSchemaNameSpace
							$WorksheetData[$Row,$Col++] = $_.General.XML.XmlSchemaNameSpaceSchema

							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 19; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 20; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 21; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 26; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 27; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 8: Tables - Default Constraints
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tables - Default Constraints"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tables - Default Constraints'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Tables | ForEach-Object { $_.Columns | Where-Object { $_.DefaultConstraint.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 13
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Column Name'
			$WorksheetData[0,$Col++] = 'Column ID'
			$WorksheetData[0,$Col++] = 'Constraint Name'
			$WorksheetData[0,$Col++] = 'Date Created'
			$WorksheetData[0,$Col++] = 'Date Modified'
			$WorksheetData[0,$Col++] = 'File Table'
			$WorksheetData[0,$Col++] = 'System Named'
			$WorksheetData[0,$Col++] = 'Text'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Tables | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Columns | Where-Object { $_.DefaultConstraint.ID } | Sort-Object -Property @{Expression={$_.General.General.ID}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.General.Name
							$WorksheetData[$Row,$Col++] = $_.General.General.ID
							$WorksheetData[$Row,$Col++] = $_.DefaultConstraint.Name
							$WorksheetData[$Row,$Col++] = $_.DefaultConstraint.CreateDate
							$WorksheetData[$Row,$Col++] = $_.DefaultConstraint.DateLastModified
							$WorksheetData[$Row,$Col++] = $_.DefaultConstraint.IsFileTableDefined
							$WorksheetData[$Row,$Col++] = $_.DefaultConstraint.IsSystemNamed
							$WorksheetData[$Row,$Col++] = if ($_.DefaultConstraint.Text.Length -gt 5000) { $_.DefaultConstraint.Text.Substring(0, 4997) + '...' } else { $_.DefaultConstraint.Text } 
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'G2'
					ColumnFormat = @(
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 9: Tables - Checks
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tables - Checks"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tables - Checks'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Tables | ForEach-Object { $_.Checks | Where-Object { $_.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 14
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Check Name'
			$WorksheetData[0,$Col++] = 'Date Created'
			$WorksheetData[0,$Col++] = 'Date Modified'
			$WorksheetData[0,$Col++] = 'Enabled'
			$WorksheetData[0,$Col++] = 'Checked'
			$WorksheetData[0,$Col++] = 'File Table'
			$WorksheetData[0,$Col++] = 'Not For Replication'
			$WorksheetData[0,$Col++] = 'Is System Named'
			$WorksheetData[0,$Col++] = 'Definition'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Tables | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Checks | Where-Object { $_.ID } | Sort-Object -Property Name | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.Name
							$WorksheetData[$Row,$Col++] = $_.CreateDate
							$WorksheetData[$Row,$Col++] = $_.DateLastModified
							$WorksheetData[$Row,$Col++] = $_.IsEnabled
							$WorksheetData[$Row,$Col++] = $_.IsChecked
							$WorksheetData[$Row,$Col++] = $_.IsFileTableDefined
							$WorksheetData[$Row,$Col++] = $_.IsNotForReplication
							$WorksheetData[$Row,$Col++] = $_.IsSystemNamed
							$WorksheetData[$Row,$Col++] = if ($_.Definition.Length -gt 5000) { $_.Definition.Substring(0, 4997) + '...' } else { $_.Definition }
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 10: Tables - Foreign Keys
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tables - Foreign Keys"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tables - Foreign Keys'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Tables | ForEach-Object { $_.ForeignKeys | Where-Object { $_.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 18
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Foreign Key Name'
			$WorksheetData[0,$Col++] = 'Column(s)'
			$WorksheetData[0,$Col++] = 'Referenced Schema'
			$WorksheetData[0,$Col++] = 'Referenced Table'
			$WorksheetData[0,$Col++] = 'Referenced Key'
			$WorksheetData[0,$Col++] = 'Date Created'
			$WorksheetData[0,$Col++] = 'Date Modified'
			$WorksheetData[0,$Col++] = 'Enabled'
			$WorksheetData[0,$Col++] = 'Checked'
			$WorksheetData[0,$Col++] = 'File Table'
			$WorksheetData[0,$Col++] = 'Not For Replication'
			$WorksheetData[0,$Col++] = 'Is System Named'
			$WorksheetData[0,$Col++] = 'Delete Action'
			$WorksheetData[0,$Col++] = 'Update Action'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Tables | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.ForeignKeys | Where-Object { $_.ID } | Sort-Object -Property Name | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.Name
							$WorksheetData[$Row,$Col++] = $_.Columns -join $Delimiter
							$WorksheetData[$Row,$Col++] = $_.ReferencedTableSchema
							$WorksheetData[$Row,$Col++] = $_.ReferencedTable
							$WorksheetData[$Row,$Col++] = $_.ReferencedKey
							$WorksheetData[$Row,$Col++] = $_.CreateDate
							$WorksheetData[$Row,$Col++] = $_.DateLastModified
							$WorksheetData[$Row,$Col++] = $_.IsEnabled
							$WorksheetData[$Row,$Col++] = $_.IsChecked
							$WorksheetData[$Row,$Col++] = $_.IsFileTableDefined
							$WorksheetData[$Row,$Col++] = $_.IsNotForReplication
							$WorksheetData[$Row,$Col++] = $_.IsSystemNamed
							$WorksheetData[$Row,$Col++] = $_.DeleteAction
							$WorksheetData[$Row,$Col++] = $_.UpdateAction
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 11; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 11: Tables - FullText Indexes
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tables - FullText Indexes"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tables - FullText Indexes'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Tables | ForEach-Object { $_.FullTextIndex | Where-Object { $_.General.CatalogName } } } } | Measure-Object).Count + 1
			$ColumnCount = 17
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Full-Text Catalog'
			$WorksheetData[0,$Col++] = 'Full-Text Index Filegroup'
			$WorksheetData[0,$Col++] = 'Full-Text Index Key'
			$WorksheetData[0,$Col++] = 'Full-Text Index Stoplist'
			$WorksheetData[0,$Col++] = 'Full-Text Index Stoplist Option'
			$WorksheetData[0,$Col++] = 'Full-Text Indexing Enabled'
			$WorksheetData[0,$Col++] = 'Search Property List'
			$WorksheetData[0,$Col++] = 'Change Tracking'
			$WorksheetData[0,$Col++] = 'Table Full-Text Docs Processed'
			$WorksheetData[0,$Col++] = 'Table Full-Text Fail Count'
			$WorksheetData[0,$Col++] = 'Table Full-Text Item Count'
			$WorksheetData[0,$Col++] = 'Table Full-Text Pending Changes'
			$WorksheetData[0,$Col++] = 'Table Full-Text Populate Status'


			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Tables | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.FullTextIndex | Where-Object { $_.General.CatalogName } | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.CatalogName
							$WorksheetData[$Row,$Col++] = $_.General.FilegroupName
							$WorksheetData[$Row,$Col++] = $_.General.UniqueIndexName
							$WorksheetData[$Row,$Col++] = $_.General.StopListName
							$WorksheetData[$Row,$Col++] = $_.General.StopListOption
							$WorksheetData[$Row,$Col++] = $_.General.IsEnabled
							$WorksheetData[$Row,$Col++] = $_.General.SearchPropertyListName
							$WorksheetData[$Row,$Col++] = $_.General.ChangeTracking
							$WorksheetData[$Row,$Col++] = $_.General.DocumentsProcessed
							$WorksheetData[$Row,$Col++] = $_.General.NumberOfFailures
							$WorksheetData[$Row,$Col++] = $_.General.ItemCount
							$WorksheetData[$Row,$Col++] = $_.General.PendingChanges
							$WorksheetData[$Row,$Col++] = $_.General.PopulationStatus
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'H2'
					ColumnFormat = @(
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 14; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 16; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 12: Tables - FullText Index Columns
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tables - FullText Index Columns"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tables - FullText Index Columns'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = (
				$SqlServerInventory.DatabaseServer | ForEach-Object { 
					$_.Server.Databases | ForEach-Object { 
						$_.Tables | ForEach-Object { 
							$_.FullTextIndex | Where-Object { $_.General.CatalogName } | ForEach-Object { 
								$_.Columns | Where-Object { $_.Name } 
							} 
						} 
					} 
				} | Measure-Object
			).Count + 1
			$ColumnCount = 11
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Full-Text Catalog'
			$WorksheetData[0,$Col++] = 'Full-Text Index Filegroup'
			$WorksheetData[0,$Col++] = 'Full-Text Index Key'
			$WorksheetData[0,$Col++] = 'Column Name'
			$WorksheetData[0,$Col++] = 'Language For Word Breaker'
			$WorksheetData[0,$Col++] = 'Type Column'
			$WorksheetData[0,$Col++] = 'Statistical Semantics'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Tables | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.FullTextIndex | Where-Object { $_.General.CatalogName } | ForEach-Object {

							$FullTextCatalogName = $_.General.CatalogName
							$FullTextFilegroupName = $_.General.CatalogName
							$FullTextUniqueIndexName = $_.General.UniqueIndexName

							$_.Columns | Where-Object { $_.Name } | Sort-Object -Property Name | ForEach-Object {
								$Col = 0
								$WorksheetData[$Row,$Col++] = $ServerName
								$WorksheetData[$Row,$Col++] = $DatabaseName
								$WorksheetData[$Row,$Col++] = $SchemaName
								$WorksheetData[$Row,$Col++] = $ObjectName
								$WorksheetData[$Row,$Col++] = $FullTextCatalogName
								$WorksheetData[$Row,$Col++] = $FullTextFilegroupName
								$WorksheetData[$Row,$Col++] = $FullTextUniqueIndexName
								$WorksheetData[$Row,$Col++] = $_.Name
								$WorksheetData[$Row,$Col++] = $_.LanguageForWordBreaker
								$WorksheetData[$Row,$Col++] = $_.TypeColumnName
								$WorksheetData[$Row,$Col++] = $_.StatisticalSemantics
								$Row++
							}
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'I2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 13: Tables - Indexes
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tables - Indexes"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tables - Indexes'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Tables | ForEach-Object { $_.Indexes | Where-Object { $_.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 26
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Index Name'
			$WorksheetData[0,$Col++] = 'Index Type'
			$WorksheetData[0,$Col++] = 'Key Type'
			$WorksheetData[0,$Col++] = 'Space Used (MB)'
			$WorksheetData[0,$Col++] = 'Compact Large Objects'
			$WorksheetData[0,$Col++] = 'Has Compressed Partitions'
			$WorksheetData[0,$Col++] = 'Filtered'
			$WorksheetData[0,$Col++] = 'Clustered'
			$WorksheetData[0,$Col++] = 'Disabled'
			$WorksheetData[0,$Col++] = 'Hypothetical'
			$WorksheetData[0,$Col++] = 'File Table'
			$WorksheetData[0,$Col++] = 'Full Text Key'
			$WorksheetData[0,$Col++] = 'On Computed Column'
			$WorksheetData[0,$Col++] = 'On Table'
			$WorksheetData[0,$Col++] = 'Partitioned'
			$WorksheetData[0,$Col++] = 'Spatial'
			$WorksheetData[0,$Col++] = 'System Named'
			$WorksheetData[0,$Col++] = 'System Object'
			$WorksheetData[0,$Col++] = 'Unique'
			$WorksheetData[0,$Col++] = 'XML Index'
			$WorksheetData[0,$Col++] = 'Parent XML Index'
			$WorksheetData[0,$Col++] = 'Secondary XML Index Type'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Tables | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Indexes | Where-Object { $_.General.ID } | Sort-Object -Property @{Expression={$_.General.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.Name
							$WorksheetData[$Row,$Col++] = $_.General.IndexType
							$WorksheetData[$Row,$Col++] = $_.General.IndexKeyType
							$WorksheetData[$Row,$Col++] = $_.General.SpaceUsedKB / 1KB
							$WorksheetData[$Row,$Col++] = $_.General.CompactLargeObjects
							$WorksheetData[$Row,$Col++] = $_.General.HasCompressedPartitions
							$WorksheetData[$Row,$Col++] = $_.General.HasFilter
							$WorksheetData[$Row,$Col++] = $_.General.IsClustered
							$WorksheetData[$Row,$Col++] = $_.General.IsDisabled
							$WorksheetData[$Row,$Col++] = $_.General.IsHypothetical
							$WorksheetData[$Row,$Col++] = $_.General.IsFileTableDefined
							$WorksheetData[$Row,$Col++] = $_.General.IsFullTextKey
							$WorksheetData[$Row,$Col++] = $_.General.IsIndexOnComputed
							$WorksheetData[$Row,$Col++] = $_.General.IsIndexOnTable
							$WorksheetData[$Row,$Col++] = $_.General.IsPartitioned
							$WorksheetData[$Row,$Col++] = $_.General.IsSpatialIndex
							$WorksheetData[$Row,$Col++] = $_.General.IsSystemNamed
							$WorksheetData[$Row,$Col++] = $_.General.IsSystemObject
							$WorksheetData[$Row,$Col++] = $_.General.IsUnique
							$WorksheetData[$Row,$Col++] = $_.General.IsXmlIndex
							$WorksheetData[$Row,$Col++] = $_.General.ParentXmlIndex
							$WorksheetData[$Row,$Col++] = $_.General.SecondaryXmlIndexType
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberS3}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 14: Tables - Index Options
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tables - Index Options"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tables - Index Options'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Tables | ForEach-Object { $_.Indexes | Where-Object { $_.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 18
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Index Name'
			$WorksheetData[0,$Col++] = 'General >>'
			$WorksheetData[0,$Col++] = 'Auto Recompute Statistics'
			$WorksheetData[0,$Col++] = 'Ignore Duplicate Values'
			$WorksheetData[0,$Col++] = 'Locks >>'
			$WorksheetData[0,$Col++] = 'Allow Row Locks'
			$WorksheetData[0,$Col++] = 'Allow Page Locks'
			$WorksheetData[0,$Col++] = 'Operation >>'
			$WorksheetData[0,$Col++] = 'Allow Online DML Processing'
			$WorksheetData[0,$Col++] = 'Max Degree Of Parallelism'
			$WorksheetData[0,$Col++] = 'Storage >>'
			$WorksheetData[0,$Col++] = 'Sort In TempDB'
			$WorksheetData[0,$Col++] = 'Fill Factor'
			$WorksheetData[0,$Col++] = 'Pad Index'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Tables | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Indexes | Where-Object { $_.General.ID } | Sort-Object -Property @{Expression={$_.General.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.Name
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = -not $_.Options.General.NoAutomaticRecomputation
							$WorksheetData[$Row,$Col++] = $_.Options.General.IgnoreDuplicateKeys
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = -not $_.Options.Locks.DisallowPageLocks
							$WorksheetData[$Row,$Col++] = -not $_.Options.Locks.DisallowRowLocks
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.Options.Operation.OnlineIndexOperation
							$WorksheetData[$Row,$Col++] = $_.Options.Operation.MaximumDegreeOfParallelism
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.Options.Storage.SortInTempdb
							$WorksheetData[$Row,$Col++] = $_.Options.Storage.FillFactor
							$WorksheetData[$Row,$Col++] = $_.Options.Storage.PadIndex
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 17; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 15: Tables - Index Storage
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tables - Index Storage"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tables - Index Storage'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Tables | ForEach-Object { $_.Indexes | Where-Object { $_.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 10
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Index Name'
			$WorksheetData[0,$Col++] = 'Filegroup'
			$WorksheetData[0,$Col++] = 'FILESTREAM Filegroup'
			$WorksheetData[0,$Col++] = 'Partition Scheme'
			$WorksheetData[0,$Col++] = 'FILESTREAM Partition Scheme'
			$WorksheetData[0,$Col++] = 'Partition Scheme Parameters'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Tables | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Indexes | Where-Object { $_.General.ID } | Sort-Object -Property @{Expression={$_.General.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.Name
							$WorksheetData[$Row,$Col++] = $_.Storage.FileGroup
							$WorksheetData[$Row,$Col++] = $_.Storage.FileStreamFileGroup
							$WorksheetData[$Row,$Col++] = $_.Storage.PartitionScheme
							$WorksheetData[$Row,$Col++] = $_.Storage.FileStreamPartitionScheme
							$WorksheetData[$Row,$Col++] = $($_.Storage.PartitionSchemeParameters | Where-Object { $_.ID } | ForEach-Object { $_.Name }) -join $Delimiter
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 16: Tables - Index Filters
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tables - Index Filters"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tables - Index Filters'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Tables | ForEach-Object { $_.Indexes | Where-Object { $_.General.HasFilter -eq $true } } } } | Measure-Object).Count + 1
			$ColumnCount = 6
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Index Name'
			$WorksheetData[0,$Col++] = 'Filter Expression'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Tables | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Indexes | Where-Object { $_.General.HasFilter -eq $true } | Sort-Object -Property @{Expression={$_.General.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.Name
							$WorksheetData[$Row,$Col++] = if ($_.FilterDefinition.Length -gt 5000) { $_.FilterDefinition.Substring(0, 4997) + '...' } else { $_.FilterDefinition }
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 17: Tables - Index Spatial
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tables - Index Spatial"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tables - Index Spatial'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Tables | ForEach-Object { $_.Indexes | Where-Object { $_.General.IsSpatialIndex -eq $true } } } } | Measure-Object).Count + 1
			$ColumnCount = 18
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Index Name'
			$WorksheetData[0,$Col++] = 'Bounding Box >>'
			$WorksheetData[0,$Col++] = 'X-min'
			$WorksheetData[0,$Col++] = 'Y-min'
			$WorksheetData[0,$Col++] = 'X-max'
			$WorksheetData[0,$Col++] = 'Y-max'
			$WorksheetData[0,$Col++] = 'General >>'
			$WorksheetData[0,$Col++] = 'Tessellation Scheme'
			$WorksheetData[0,$Col++] = 'Cells Per Object'
			$WorksheetData[0,$Col++] = 'Grids >>'
			$WorksheetData[0,$Col++] = 'Level 1'
			$WorksheetData[0,$Col++] = 'Level 2'
			$WorksheetData[0,$Col++] = 'Level 3'
			$WorksheetData[0,$Col++] = 'Level 4'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Tables | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Indexes | Where-Object { $_.General.IsSpatialIndex -eq $true } | Sort-Object -Property @{Expression={$_.General.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.Name
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.Spatial.BoundingBox.XMin
							$WorksheetData[$Row,$Col++] = $_.Spatial.BoundingBox.YMin
							$WorksheetData[$Row,$Col++] = $_.Spatial.BoundingBox.XMax
							$WorksheetData[$Row,$Col++] = $_.Spatial.BoundingBox.YMax
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.Spatial.General.SpatialIndexType
							$WorksheetData[$Row,$Col++] = $_.Spatial.General.CellsPerObject
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.Spatial.Grids.Level1Grid
							$WorksheetData[$Row,$Col++] = $_.Spatial.Grids.Level2Grid
							$WorksheetData[$Row,$Col++] = $_.Spatial.Grids.Level3Grid
							$WorksheetData[$Row,$Col++] = $_.Spatial.Grids.Level4Grid
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 18: Tables - Statistics
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tables - Statistics"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tables - Statistics'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Tables | ForEach-Object { $_.Statistics | Where-Object { $_.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 13
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Statistics Name'
			$WorksheetData[0,$Col++] = 'Date Updated'
			$WorksheetData[0,$Col++] = 'Auto Created'
			$WorksheetData[0,$Col++] = 'Auto Update Enabled'
			$WorksheetData[0,$Col++] = 'Filegroup'
			$WorksheetData[0,$Col++] = 'From Index Creation'
			$WorksheetData[0,$Col++] = 'Has Filter'
			$WorksheetData[0,$Col++] = 'Columns'
			$WorksheetData[0,$Col++] = 'Filter Definition'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Tables | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Statistics | Where-Object { $_.General.ID } | Sort-Object -Property @{Expression={$_.General.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.Name
							$WorksheetData[$Row,$Col++] = $_.General.LastUpdated
							$WorksheetData[$Row,$Col++] = $_.General.IsAutoCreated
							$WorksheetData[$Row,$Col++] = -not $_.General.NoAutomaticRecomputation
							$WorksheetData[$Row,$Col++] = $_.General.FileGroup
							$WorksheetData[$Row,$Col++] = $_.General.IsFromIndexCreation
							$WorksheetData[$Row,$Col++] = $_.General.HasFilter
							$WorksheetData[$Row,$Col++] = $($_.General.Columns | Where-Object { $_.ID } | ForEach-Object { $_.Name }) -join $Delimiter
							$WorksheetData[$Row,$Col++] = if ($_.FilterDefinition.Length -gt 5000) { $_.FilterDefinition.Substring(0, 4997) + '...' } else { $_.FilterDefinition }
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 19: Tables - Triggers
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tables - Triggers"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tables - Triggers'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Tables | ForEach-Object { $_.Triggers | Where-Object { $_.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 26
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Trigger Name'
			$WorksheetData[0,$Col++] = 'Implementation Type'
			$WorksheetData[0,$Col++] = 'Enabled'
			$WorksheetData[0,$Col++] = 'Encrypted'
			$WorksheetData[0,$Col++] = 'Is System Object'
			$WorksheetData[0,$Col++] = 'Date Created'
			$WorksheetData[0,$Col++] = 'Date Modified'
			$WorksheetData[0,$Col++] = 'Instead Of'
			$WorksheetData[0,$Col++] = 'For Insert'
			$WorksheetData[0,$Col++] = 'Insert Order'
			$WorksheetData[0,$Col++] = 'For Update'
			$WorksheetData[0,$Col++] = 'Update Order'
			$WorksheetData[0,$Col++] = 'For Delete'
			$WorksheetData[0,$Col++] = 'Delete Order'
			$WorksheetData[0,$Col++] = 'ANSI NULL'
			$WorksheetData[0,$Col++] = 'Quoted Identifier'
			$WorksheetData[0,$Col++] = 'Not For Replication'
			$WorksheetData[0,$Col++] = 'Execution Context'
			$WorksheetData[0,$Col++] = 'Execute As'
			$WorksheetData[0,$Col++] = 'CLR Assembly Name'
			$WorksheetData[0,$Col++] = 'CLR Class Name'
			$WorksheetData[0,$Col++] = 'CLR Method Name'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Tables | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Triggers | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.Name
							$WorksheetData[$Row,$Col++] = $_.ImplementationType
							$WorksheetData[$Row,$Col++] = $_.IsEnabled
							$WorksheetData[$Row,$Col++] = $_.IsEncrypted
							$WorksheetData[$Row,$Col++] = $_.IsSystemObject
							$WorksheetData[$Row,$Col++] = $_.CreateDate
							$WorksheetData[$Row,$Col++] = $_.DateLastModified
							$WorksheetData[$Row,$Col++] = $_.InsteadOf
							$WorksheetData[$Row,$Col++] = $_.Insert
							$WorksheetData[$Row,$Col++] = $_.InsertOrder
							$WorksheetData[$Row,$Col++] = $_.Update
							$WorksheetData[$Row,$Col++] = $_.UpdateOrder
							$WorksheetData[$Row,$Col++] = $_.Delete
							$WorksheetData[$Row,$Col++] = $_.DeleteOrder
							$WorksheetData[$Row,$Col++] = $_.AnsiNullsStatus
							$WorksheetData[$Row,$Col++] = $_.QuotedIdentifierStatus
							$WorksheetData[$Row,$Col++] = $_.NotForReplication
							$WorksheetData[$Row,$Col++] = $_.ExecutionContext
							$WorksheetData[$Row,$Col++] = $_.ExecutionContextPrincipal
							$WorksheetData[$Row,$Col++] = $_.AssemblyName
							$WorksheetData[$Row,$Col++] = $_.ClassName
							$WorksheetData[$Row,$Col++] = $_.MethodName
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberS3}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 20: Tables - Trigger Definitions
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Tables - Trigger Definitions"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Tables - Trigger Definitions'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Tables | ForEach-Object { $_.Triggers | Where-Object { $_.ImplementationType -ieq 'T-SQL' } } } } | Measure-Object).Count + 1
			$ColumnCount = 6
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Trigger Name'
			$WorksheetData[0,$Col++] = 'Definition'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Tables | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Triggers | Where-Object { $_.ImplementationType -ieq 'T-SQL' } | Sort-Object -Property @{Expression={$_.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.Name
							$WorksheetData[$Row,$Col++] = if ($_.Definition.Length -gt 5000) { $_.Definition.Substring(0, 4997) + '...' } else { $_.Definition }
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 21: Views - General
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Views - General"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Views - General'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Views | Where-Object { $_.Properties.General.Description.Name } } } | Measure-Object).Count + 1
			$ColumnCount = 23
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'View Name'
			$WorksheetData[0,$Col++] = 'Description >>'
			$WorksheetData[0,$Col++] = 'Created Date'
			$WorksheetData[0,$Col++] = 'Last Modified Date'
			$WorksheetData[0,$Col++] = 'System Object'
			$WorksheetData[0,$Col++] = 'Has After Trigger'
			$WorksheetData[0,$Col++] = 'Has Column Specification'
			$WorksheetData[0,$Col++] = 'Has Delete Trigger'
			$WorksheetData[0,$Col++] = 'Has Index'
			$WorksheetData[0,$Col++] = 'Has Insert Trigger'
			$WorksheetData[0,$Col++] = 'Has InsteadOf Trigger'
			$WorksheetData[0,$Col++] = 'Has Update Trigger'
			$WorksheetData[0,$Col++] = 'Is Indexable'
			$WorksheetData[0,$Col++] = 'Is Schema Owned'
			$WorksheetData[0,$Col++] = 'Options >>'
			$WorksheetData[0,$Col++] = 'ANSI NULLs'
			$WorksheetData[0,$Col++] = 'Encrypted'
			$WorksheetData[0,$Col++] = 'Quoted Identifier'
			$WorksheetData[0,$Col++] = 'Schema Bound'
			$WorksheetData[0,$Col++] = 'Returns View Metadata'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Views | Where-Object { $_.Properties.General.Description.Name } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Schema
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Name
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.CreateDate
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.DateLastModified
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.IsSystemObject
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.HasAfterTrigger
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.HasColumnSpecification
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.HasDeleteTrigger
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.HasIndex
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.HasInsertTrigger
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.HasInsteadOfTrigger
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.HasUpdateTrigger
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.IsIndexable
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.IsSchemaOwned
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.AnsiNulls
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.IsEncrypted
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.QuotedIdentifier
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.IsSchemaBound
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.ReturnsViewMetadata
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 22: Views - Columns
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Views - Columns"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Views - Columns'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Views | ForEach-Object { $_.Columns | Where-Object { $_.General.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 44
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Column Name'
			$WorksheetData[0,$Col++] = 'Binding >>'
			$WorksheetData[0,$Col++] = 'Default Binding'
			$WorksheetData[0,$Col++] = 'Default Schema'
			$WorksheetData[0,$Col++] = 'Rule'
			$WorksheetData[0,$Col++] = 'Rule Schema'
			$WorksheetData[0,$Col++] = 'Computed >>'
			$WorksheetData[0,$Col++] = 'Is Computed'
			$WorksheetData[0,$Col++] = 'Computed Text'
			$WorksheetData[0,$Col++] = 'General >>'
			$WorksheetData[0,$Col++] = 'Allow Nulls'
			$WorksheetData[0,$Col++] = 'ANSI Padding Status'
			$WorksheetData[0,$Col++] = 'Column ID'
			$WorksheetData[0,$Col++] = 'Data Type'
			$WorksheetData[0,$Col++] = 'Length'
			$WorksheetData[0,$Col++] = 'Numeric Precision'
			$WorksheetData[0,$Col++] = 'Numeric Scale'
			$WorksheetData[0,$Col++] = 'Primary Key'
			$WorksheetData[0,$Col++] = 'System Type'
			$WorksheetData[0,$Col++] = 'Identity >>'
			$WorksheetData[0,$Col++] = 'Identity'
			$WorksheetData[0,$Col++] = 'Identity Seed'
			$WorksheetData[0,$Col++] = 'Identity Increment'
			$WorksheetData[0,$Col++] = 'Miscellaneous >>'
			$WorksheetData[0,$Col++] = 'Collation'
			$WorksheetData[0,$Col++] = 'Full Text'
			$WorksheetData[0,$Col++] = 'Not For Replication'
			$WorksheetData[0,$Col++] = 'Statistical Semantics'
			$WorksheetData[0,$Col++] = 'Is Deterministic'
			$WorksheetData[0,$Col++] = 'Is FILESTREAM'
			$WorksheetData[0,$Col++] = 'Is Foreign Key'
			$WorksheetData[0,$Col++] = 'Is Persisted'
			$WorksheetData[0,$Col++] = 'Is Precise'
			$WorksheetData[0,$Col++] = 'Is ROWGUID'
			$WorksheetData[0,$Col++] = 'Sparse >>'
			$WorksheetData[0,$Col++] = 'Is Column Set'
			$WorksheetData[0,$Col++] = 'Is Sparse'
			$WorksheetData[0,$Col++] = 'XML >>'
			$WorksheetData[0,$Col++] = 'XML Schema Namespace'
			$WorksheetData[0,$Col++] = 'XML Schema Namespace Schema'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Views | Where-Object { $_.Properties.General.Description.Name } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Columns | Where-Object { $_.General.General.ID } | Sort-Object -Property @{Expression={$_.General.General.ID}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.General.Name
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Binding.DefaultBinding
							$WorksheetData[$Row,$Col++] = $_.General.Binding.DefaultSchema
							$WorksheetData[$Row,$Col++] = $_.General.Binding.Rule
							$WorksheetData[$Row,$Col++] = $_.General.Binding.RuleSchema
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Computed.IsComputed
							$WorksheetData[$Row,$Col++] = $_.General.Computed.ComputedText
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.General.AllowNulls
							$WorksheetData[$Row,$Col++] = $_.General.General.AnsiPaddingStatus
							$WorksheetData[$Row,$Col++] = $_.General.General.ID
							$WorksheetData[$Row,$Col++] = $_.General.General.DataType
							$WorksheetData[$Row,$Col++] = $_.General.General.Length
							$WorksheetData[$Row,$Col++] = $_.General.General.NumericPrecision
							$WorksheetData[$Row,$Col++] = $_.General.General.NumericScale
							$WorksheetData[$Row,$Col++] = $_.General.General.InPrimaryKey
							$WorksheetData[$Row,$Col++] = $_.General.General.SystemType
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Identity.IsIdentity
							$WorksheetData[$Row,$Col++] = $_.General.Identity.IdentityIncrement
							$WorksheetData[$Row,$Col++] = $_.General.Identity.IdentitySeed
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.Collation
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsFullTextIndexed
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsNotForReplication
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.StatisticalSemantics
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsDeterministic
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsFileStream
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsForeignKey
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsPersisted
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsPrecise
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsRowGuidCol
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Sparse.IsColumnSet
							$WorksheetData[$Row,$Col++] = $_.General.Sparse.IsSparse
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.XML.XmlSchemaNameSpace
							$WorksheetData[$Row,$Col++] = $_.General.XML.XmlSchemaNameSpaceSchema

							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 19; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 20; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 21; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 26; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 27; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 23: Views - FullText Indexes
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Views - FullText Indexes"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Views - FullText Indexes'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Views | ForEach-Object { $_.FullTextIndex | Where-Object { $_.General.CatalogName } } } } | Measure-Object).Count + 1
			$ColumnCount = 17
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'View Name'
			$WorksheetData[0,$Col++] = 'Full-Text Catalog'
			$WorksheetData[0,$Col++] = 'Full-Text Index Filegroup'
			$WorksheetData[0,$Col++] = 'Full-Text Index Key'
			$WorksheetData[0,$Col++] = 'Full-Text Index Stoplist'
			$WorksheetData[0,$Col++] = 'Full-Text Index Stoplist Option'
			$WorksheetData[0,$Col++] = 'Full-Text Indexing Enabled'
			$WorksheetData[0,$Col++] = 'Search Property List'
			$WorksheetData[0,$Col++] = 'Change Tracking'
			$WorksheetData[0,$Col++] = 'Table Full-Text Docs Processed'
			$WorksheetData[0,$Col++] = 'Table Full-Text Fail Count'
			$WorksheetData[0,$Col++] = 'Table Full-Text Item Count'
			$WorksheetData[0,$Col++] = 'Table Full-Text Pending Changes'
			$WorksheetData[0,$Col++] = 'Table Full-Text Populate Status'


			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Views | Where-Object { $_.Properties.General.Description.Name } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.FullTextIndex | Where-Object { $_.General.CatalogName } | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.CatalogName
							$WorksheetData[$Row,$Col++] = $_.General.FilegroupName
							$WorksheetData[$Row,$Col++] = $_.General.UniqueIndexName
							$WorksheetData[$Row,$Col++] = $_.General.StopListName
							$WorksheetData[$Row,$Col++] = $_.General.StopListOption
							$WorksheetData[$Row,$Col++] = $_.General.IsEnabled
							$WorksheetData[$Row,$Col++] = $_.General.SearchPropertyListName
							$WorksheetData[$Row,$Col++] = $_.General.ChangeTracking
							$WorksheetData[$Row,$Col++] = $_.General.DocumentsProcessed
							$WorksheetData[$Row,$Col++] = $_.General.NumberOfFailures
							$WorksheetData[$Row,$Col++] = $_.General.ItemCount
							$WorksheetData[$Row,$Col++] = $_.General.PendingChanges
							$WorksheetData[$Row,$Col++] = $_.General.PopulationStatus
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'H2'
					ColumnFormat = @(
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 14; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 16; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 24: Views - FullText Index Columns
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Views - FullText Index Columns"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Views - FullText Index Columns'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			$RowCount = (
				$SqlServerInventory.DatabaseServer | ForEach-Object { 
					$_.Server.Databases | ForEach-Object { 
						$_.Views | ForEach-Object { 
							$_.FullTextIndex | Where-Object { $_.General.CatalogName } | ForEach-Object { 
								$_.Columns | Where-Object { $_.Name } 
							} 
						} 
					} 
				} | Measure-Object
			).Count + 1
			$ColumnCount = 11
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'View Name'
			$WorksheetData[0,$Col++] = 'Full-Text Catalog'
			$WorksheetData[0,$Col++] = 'Full-Text Index Filegroup'
			$WorksheetData[0,$Col++] = 'Full-Text Index Key'
			$WorksheetData[0,$Col++] = 'Column Name'
			$WorksheetData[0,$Col++] = 'Language For Word Breaker'
			$WorksheetData[0,$Col++] = 'Type Column'
			$WorksheetData[0,$Col++] = 'Statistical Semantics'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Views | Where-Object { $_.Properties.General.Description.Name } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.FullTextIndex | Where-Object { $_.General.CatalogName } | ForEach-Object {

							$FullTextCatalogName = $_.General.CatalogName
							$FullTextFilegroupName = $_.General.CatalogName
							$FullTextUniqueIndexName = $_.General.UniqueIndexName

							$_.Columns | Where-Object { $_.Name } | Sort-Object -Property Name | ForEach-Object {
								$Col = 0
								$WorksheetData[$Row,$Col++] = $ServerName
								$WorksheetData[$Row,$Col++] = $DatabaseName
								$WorksheetData[$Row,$Col++] = $SchemaName
								$WorksheetData[$Row,$Col++] = $ObjectName
								$WorksheetData[$Row,$Col++] = $FullTextCatalogName
								$WorksheetData[$Row,$Col++] = $FullTextFilegroupName
								$WorksheetData[$Row,$Col++] = $FullTextUniqueIndexName
								$WorksheetData[$Row,$Col++] = $_.Name
								$WorksheetData[$Row,$Col++] = $_.LanguageForWordBreaker
								$WorksheetData[$Row,$Col++] = $_.TypeColumnName
								$WorksheetData[$Row,$Col++] = $_.StatisticalSemantics
								$Row++
							}
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'I2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 25: Views - Indexes
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Views - Indexes"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Views - Indexes'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Views | ForEach-Object { $_.Indexes | Where-Object { $_.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 26
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'View Name'
			$WorksheetData[0,$Col++] = 'Index Name'
			$WorksheetData[0,$Col++] = 'Index Type'
			$WorksheetData[0,$Col++] = 'Key Type'
			$WorksheetData[0,$Col++] = 'Space Used (MB)'
			$WorksheetData[0,$Col++] = 'Compact Large Objects'
			$WorksheetData[0,$Col++] = 'Has Compressed Partitions'
			$WorksheetData[0,$Col++] = 'Filtered'
			$WorksheetData[0,$Col++] = 'Clustered'
			$WorksheetData[0,$Col++] = 'Disabled'
			$WorksheetData[0,$Col++] = 'Hypothetical'
			$WorksheetData[0,$Col++] = 'File Table'
			$WorksheetData[0,$Col++] = 'Full Text Key'
			$WorksheetData[0,$Col++] = 'On Computed Column'
			$WorksheetData[0,$Col++] = 'On Table'
			$WorksheetData[0,$Col++] = 'Partitioned'
			$WorksheetData[0,$Col++] = 'Spatial'
			$WorksheetData[0,$Col++] = 'System Named'
			$WorksheetData[0,$Col++] = 'System Object'
			$WorksheetData[0,$Col++] = 'Unique'
			$WorksheetData[0,$Col++] = 'XML Index'
			$WorksheetData[0,$Col++] = 'Parent XML Index'
			$WorksheetData[0,$Col++] = 'Secondary XML Index Type'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Views | Where-Object { $_.Properties.General.Description.Name } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Indexes | Where-Object { $_.General.ID } | Sort-Object -Property @{Expression={$_.General.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.Name
							$WorksheetData[$Row,$Col++] = $_.General.IndexType
							$WorksheetData[$Row,$Col++] = $_.General.IndexKeyType
							$WorksheetData[$Row,$Col++] = $_.General.SpaceUsedKB / 1KB
							$WorksheetData[$Row,$Col++] = $_.General.CompactLargeObjects
							$WorksheetData[$Row,$Col++] = $_.General.HasCompressedPartitions
							$WorksheetData[$Row,$Col++] = $_.General.HasFilter
							$WorksheetData[$Row,$Col++] = $_.General.IsClustered
							$WorksheetData[$Row,$Col++] = $_.General.IsDisabled
							$WorksheetData[$Row,$Col++] = $_.General.IsHypothetical
							$WorksheetData[$Row,$Col++] = $_.General.IsFileTableDefined
							$WorksheetData[$Row,$Col++] = $_.General.IsFullTextKey
							$WorksheetData[$Row,$Col++] = $_.General.IsIndexOnComputed
							$WorksheetData[$Row,$Col++] = $_.General.IsIndexOnTable
							$WorksheetData[$Row,$Col++] = $_.General.IsPartitioned
							$WorksheetData[$Row,$Col++] = $_.General.IsSpatialIndex
							$WorksheetData[$Row,$Col++] = $_.General.IsSystemNamed
							$WorksheetData[$Row,$Col++] = $_.General.IsSystemObject
							$WorksheetData[$Row,$Col++] = $_.General.IsUnique
							$WorksheetData[$Row,$Col++] = $_.General.IsXmlIndex
							$WorksheetData[$Row,$Col++] = $_.General.ParentXmlIndex
							$WorksheetData[$Row,$Col++] = $_.General.SecondaryXmlIndexType
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberS3}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 26: Views - Index Options
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Views - Index Options"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Views - Index Options'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Views | ForEach-Object { $_.Indexes | Where-Object { $_.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 18
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'View Name'
			$WorksheetData[0,$Col++] = 'Index Name'
			$WorksheetData[0,$Col++] = 'General >>'
			$WorksheetData[0,$Col++] = 'Auto Recompute Statistics'
			$WorksheetData[0,$Col++] = 'Ignore Duplicate Values'
			$WorksheetData[0,$Col++] = 'Locks >>'
			$WorksheetData[0,$Col++] = 'Allow Row Locks'
			$WorksheetData[0,$Col++] = 'Allow Page Locks'
			$WorksheetData[0,$Col++] = 'Operation >>'
			$WorksheetData[0,$Col++] = 'Allow Online DML Processing'
			$WorksheetData[0,$Col++] = 'Max Degree Of Parallelism'
			$WorksheetData[0,$Col++] = 'Storage >>'
			$WorksheetData[0,$Col++] = 'Sort In TempDB'
			$WorksheetData[0,$Col++] = 'Fill Factor'
			$WorksheetData[0,$Col++] = 'Pad Index'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Views | Where-Object { $_.Properties.General.Description.Name } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Indexes | Where-Object { $_.General.ID } | Sort-Object -Property @{Expression={$_.General.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.Name
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = -not $_.Options.General.NoAutomaticRecomputation
							$WorksheetData[$Row,$Col++] = $_.Options.General.IgnoreDuplicateKeys
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = -not $_.Options.Locks.DisallowPageLocks
							$WorksheetData[$Row,$Col++] = -not $_.Options.Locks.DisallowRowLocks
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.Options.Operation.OnlineIndexOperation
							$WorksheetData[$Row,$Col++] = $_.Options.Operation.MaximumDegreeOfParallelism
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.Options.Storage.SortInTempdb
							$WorksheetData[$Row,$Col++] = $_.Options.Storage.FillFactor
							$WorksheetData[$Row,$Col++] = $_.Options.Storage.PadIndex
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 17; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 27: Views - Index Storage
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Views - Index Storage"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Views - Index Storage'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Views | ForEach-Object { $_.Indexes | Where-Object { $_.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 10
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'View Name'
			$WorksheetData[0,$Col++] = 'Index Name'
			$WorksheetData[0,$Col++] = 'Filegroup'
			$WorksheetData[0,$Col++] = 'FILESTREAM Filegroup'
			$WorksheetData[0,$Col++] = 'Partition Scheme'
			$WorksheetData[0,$Col++] = 'FILESTREAM Partition Scheme'
			$WorksheetData[0,$Col++] = 'Partition Scheme Parameters'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Views | Where-Object { $_.Properties.General.Description.Name } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Indexes | Where-Object { $_.General.ID } | Sort-Object -Property @{Expression={$_.General.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.Name
							$WorksheetData[$Row,$Col++] = $_.Storage.FileGroup
							$WorksheetData[$Row,$Col++] = $_.Storage.FileStreamFileGroup
							$WorksheetData[$Row,$Col++] = $_.Storage.PartitionScheme
							$WorksheetData[$Row,$Col++] = $_.Storage.FileStreamPartitionScheme
							$WorksheetData[$Row,$Col++] = $($_.Storage.PartitionSchemeParameters | Where-Object { $_.ID } | ForEach-Object { $_.Name }) -join $Delimiter
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 28: Views - Statistics
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Views - Statistics"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Views - Statistics'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Views | ForEach-Object { $_.Statistics | Where-Object { $_.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 13
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'View Name'
			$WorksheetData[0,$Col++] = 'Statistics Name'
			$WorksheetData[0,$Col++] = 'Date Updated'
			$WorksheetData[0,$Col++] = 'Auto Created'
			$WorksheetData[0,$Col++] = 'Auto Update Enabled'
			$WorksheetData[0,$Col++] = 'Filegroup'
			$WorksheetData[0,$Col++] = 'From Index Creation'
			$WorksheetData[0,$Col++] = 'Has Filter'
			$WorksheetData[0,$Col++] = 'Columns'
			$WorksheetData[0,$Col++] = 'Filter Definition'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Views | Where-Object { $_.Properties.General.Description.Name } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Statistics | Where-Object { $_.General.ID } | Sort-Object -Property @{Expression={$_.General.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.Name
							$WorksheetData[$Row,$Col++] = $_.General.LastUpdated
							$WorksheetData[$Row,$Col++] = $_.General.IsAutoCreated
							$WorksheetData[$Row,$Col++] = -not $_.General.NoAutomaticRecomputation
							$WorksheetData[$Row,$Col++] = $_.General.FileGroup
							$WorksheetData[$Row,$Col++] = $_.General.IsFromIndexCreation
							$WorksheetData[$Row,$Col++] = $_.General.HasFilter
							$WorksheetData[$Row,$Col++] = $($_.General.Columns | Where-Object { $_.ID } | ForEach-Object { $_.Name }) -join $Delimiter
							$WorksheetData[$Row,$Col++] = if ($_.FilterDefinition.Length -gt 5000) { $_.FilterDefinition.Substring(0, 4997) + '...' } else { $_.FilterDefinition }
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 29: Views - Triggers
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Views - Triggers"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Views - Triggers'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Views | ForEach-Object { $_.Triggers | Where-Object { $_.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 26
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'View Name'
			$WorksheetData[0,$Col++] = 'Trigger Name'
			$WorksheetData[0,$Col++] = 'Implementation Type'
			$WorksheetData[0,$Col++] = 'Enabled'
			$WorksheetData[0,$Col++] = 'Encrypted'
			$WorksheetData[0,$Col++] = 'Is System Object'
			$WorksheetData[0,$Col++] = 'Date Created'
			$WorksheetData[0,$Col++] = 'Date Modified'
			$WorksheetData[0,$Col++] = 'Instead Of'
			$WorksheetData[0,$Col++] = 'For Insert'
			$WorksheetData[0,$Col++] = 'Insert Order'
			$WorksheetData[0,$Col++] = 'For Update'
			$WorksheetData[0,$Col++] = 'Update Order'
			$WorksheetData[0,$Col++] = 'For Delete'
			$WorksheetData[0,$Col++] = 'Delete Order'
			$WorksheetData[0,$Col++] = 'ANSI NULL'
			$WorksheetData[0,$Col++] = 'Quoted Identifier'
			$WorksheetData[0,$Col++] = 'Not For Replication'
			$WorksheetData[0,$Col++] = 'Execution Context'
			$WorksheetData[0,$Col++] = 'Execute As'
			$WorksheetData[0,$Col++] = 'CLR Assembly Name'
			$WorksheetData[0,$Col++] = 'CLR Class Name'
			$WorksheetData[0,$Col++] = 'CLR Method Name'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Views | Where-Object { $_.Properties.General.Description.Name } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Triggers | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.Name
							$WorksheetData[$Row,$Col++] = $_.ImplementationType
							$WorksheetData[$Row,$Col++] = $_.IsEnabled
							$WorksheetData[$Row,$Col++] = $_.IsEncrypted
							$WorksheetData[$Row,$Col++] = $_.IsSystemObject
							$WorksheetData[$Row,$Col++] = $_.CreateDate
							$WorksheetData[$Row,$Col++] = $_.DateLastModified
							$WorksheetData[$Row,$Col++] = $_.InsteadOf
							$WorksheetData[$Row,$Col++] = $_.Insert
							$WorksheetData[$Row,$Col++] = $_.InsertOrder
							$WorksheetData[$Row,$Col++] = $_.Update
							$WorksheetData[$Row,$Col++] = $_.UpdateOrder
							$WorksheetData[$Row,$Col++] = $_.Delete
							$WorksheetData[$Row,$Col++] = $_.DeleteOrder
							$WorksheetData[$Row,$Col++] = $_.AnsiNullsStatus
							$WorksheetData[$Row,$Col++] = $_.QuotedIdentifierStatus
							$WorksheetData[$Row,$Col++] = $_.NotForReplication
							$WorksheetData[$Row,$Col++] = $_.ExecutionContext
							$WorksheetData[$Row,$Col++] = $_.ExecutionContextPrincipal
							$WorksheetData[$Row,$Col++] = $_.AssemblyName
							$WorksheetData[$Row,$Col++] = $_.ClassName
							$WorksheetData[$Row,$Col++] = $_.MethodName
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberS3}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 30: Views - Trigger Definitions
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Views - Trigger Definitions"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Views - Trigger Definitions'
			#$Worksheet.Tab.Color = $DatabaseTabColor
			$Worksheet.Tab.ThemeColor = $DatabaseTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Views | ForEach-Object { $_.Triggers | Where-Object { $_.ImplementationType -ieq 'T-SQL' } } } } | Measure-Object).Count + 1
			$ColumnCount = 6
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'View Name'
			$WorksheetData[0,$Col++] = 'Trigger Name'
			$WorksheetData[0,$Col++] = 'Definition'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Views | Where-Object { $_.Properties.General.Description.Name } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Triggers | Where-Object { $_.ImplementationType -ieq 'T-SQL' } | Sort-Object -Property @{Expression={$_.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.Name
							$WorksheetData[$Row,$Col++] = if ($_.Definition.Length -gt 5000) { $_.Definition.Substring(0, 4997) + '...' } else { $_.Definition }
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 31: Synonyms
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Synonyms"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Synonyms'
			#$Worksheet.Tab.Color = $SecurityTabColor
			$Worksheet.Tab.ThemeColor = $SecurityTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Synonyms | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 12
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Synonym Name'
			$WorksheetData[0,$Col++] = 'Base Server'
			$WorksheetData[0,$Col++] = 'Base Database'
			$WorksheetData[0,$Col++] = 'Base Schema'
			$WorksheetData[0,$Col++] = 'Base Object'
			$WorksheetData[0,$Col++] = 'Base Type'
			$WorksheetData[0,$Col++] = 'Is Schema Owned'
			$WorksheetData[0,$Col++] = 'Date Created'
			$WorksheetData[0,$Col++] = 'Date Modified'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Synonyms | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Schema}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Schema
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.BaseServer
						$WorksheetData[$Row,$Col++] = $_.BaseDatabase
						$WorksheetData[$Row,$Col++] = $_.BaseSchema
						$WorksheetData[$Row,$Col++] = $_.BaseObject
						$WorksheetData[$Row,$Col++] = $_.BaseType
						$WorksheetData[$Row,$Col++] = $_.IsSchemaOwned
						$WorksheetData[$Row,$Col++] = $_.CreateDate
						$WorksheetData[$Row,$Col++] = $_.DateLastModified
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 11; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 12; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 32: Stored Procedures - General
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Stored Procedures - General"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Stored Procedures - General'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.StoredProcedures | Where-Object { $_.Properties.General.Description.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 21
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Procedure Name'
			$WorksheetData[0,$Col++] = 'Description >>'
			$WorksheetData[0,$Col++] = 'Implementation Type'
			$WorksheetData[0,$Col++] = 'Created Date'
			$WorksheetData[0,$Col++] = 'Last Modified Date'
			$WorksheetData[0,$Col++] = 'System Object'
			$WorksheetData[0,$Col++] = 'Execution As'
			$WorksheetData[0,$Col++] = 'Execution Principal'
			$WorksheetData[0,$Col++] = 'CLR Assembly Name'
			$WorksheetData[0,$Col++] = 'CLR Class Name'
			$WorksheetData[0,$Col++] = 'CLR Method Name'
			$WorksheetData[0,$Col++] = 'Options >>'
			$WorksheetData[0,$Col++] = 'ANSI NULLs'
			$WorksheetData[0,$Col++] = 'Encrypted'
			$WorksheetData[0,$Col++] = 'Quoted Identifier'
			$WorksheetData[0,$Col++] = 'For Replication'
			$WorksheetData[0,$Col++] = 'Recompile'
			$WorksheetData[0,$Col++] = 'Startup Procedure'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.StoredProcedures | Where-Object { $_.Properties.General.Description.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Schema
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Name
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.ImplementationType
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.CreateDate
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.DateLastModified
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.IsSystemObject
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.ExecutionContext
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.ExecutionContextPrincipal
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.AssemblyName
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.ClassName
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.MethodName
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.AnsiNulls
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.IsEncrypted
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.QuotedIdentifierStatus
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.ForReplication
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.Recompile
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.Startup
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 33: Stored Procedures - Parameters
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Stored Procedures - Parameters"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Stored Procedures - Parameters'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = (
				$SqlServerInventory.DatabaseServer | ForEach-Object { 
					$_.Server.Databases | ForEach-Object { 
						$_.Programmability.StoredProcedures | Where-Object { $_.Properties.General.Description.ID } | ForEach-Object { 
							$_.Parameters | Where-Object { $_.ID } 
						} 
					} 
				} | Measure-Object
			).Count + 1
			$ColumnCount = 17
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Procedure Name'
			$WorksheetData[0,$Col++] = 'Parameter Name'
			$WorksheetData[0,$Col++] = 'Parameter ID'
			$WorksheetData[0,$Col++] = 'Cursor Parameter'
			$WorksheetData[0,$Col++] = 'Output Parameter'
			$WorksheetData[0,$Col++] = 'Read Only'
			$WorksheetData[0,$Col++] = 'Default Value'
			$WorksheetData[0,$Col++] = 'Data Type'
			$WorksheetData[0,$Col++] = 'Length'
			$WorksheetData[0,$Col++] = 'Numeric Precision'
			$WorksheetData[0,$Col++] = 'Numeric Scale'
			$WorksheetData[0,$Col++] = 'System Type'
			$WorksheetData[0,$Col++] = 'Schema'
			$WorksheetData[0,$Col++] = 'XML Document Constraint'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.StoredProcedures | Where-Object { $_.Properties.General.Description.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Parameters | Where-Object { $_.ID } | Sort-Object -Property ID | ForEach-Object {
							$Col = 0 
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.Name
							$WorksheetData[$Row,$Col++] = $_.ID
							$WorksheetData[$Row,$Col++] = $_.IsCursorParameter
							$WorksheetData[$Row,$Col++] = $_.IsOutputParameter
							$WorksheetData[$Row,$Col++] = $_.IsReadOnly
							$WorksheetData[$Row,$Col++] = $_.DefaultValue
							$WorksheetData[$Row,$Col++] = $_.DataType.Name
							$WorksheetData[$Row,$Col++] = $_.DataType.MaximumLength
							$WorksheetData[$Row,$Col++] = $_.DataType.NumericPrecision
							$WorksheetData[$Row,$Col++] = $_.DataType.NumericScale
							$WorksheetData[$Row,$Col++] = $_.DataType.SqlDataType
							$WorksheetData[$Row,$Col++] = $_.DataType.Schema
							$WorksheetData[$Row,$Col++] = $_.DataType.XmlDocumentConstraint
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtText},
						@{ColumnNumber = 12; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 14; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 34: Stored Procedures - Definition
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Stored Procedures - Definition"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Stored Procedures - Definition'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.StoredProcedures | Where-Object { $_.Properties.General.Description.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 5
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Procedure Name'
			$WorksheetData[0,$Col++] = 'Definition'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.StoredProcedures | Where-Object { $_.Properties.General.Description.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Schema
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Name
						$WorksheetData[$Row,$Col++] = if ($_.Definition.Length -gt 5000) { $_.Definition.Substring(0, 4997) + '...' } else { $_.Definition }
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 35: Extended Stored Procedures
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Extended Stored Procedures"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Extended Stored Procedures'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.ExtendedStoredProcedures | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 21
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Extended Procedure Name'
			$WorksheetData[0,$Col++] = 'Created Date'
			$WorksheetData[0,$Col++] = 'Last Modified Date'
			$WorksheetData[0,$Col++] = 'System Object'
			$WorksheetData[0,$Col++] = 'DLL Location'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.ExtendedStoredProcedures | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Schema}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Schema
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.CreateDate
						$WorksheetData[$Row,$Col++] = $_.DateLastModified
						$WorksheetData[$Row,$Col++] = $_.IsSystemObject
						$WorksheetData[$Row,$Col++] = $_.DllLocation
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 36: Functions - General
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Functions - General"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Functions - General'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Functions | Where-Object { $_.Properties.General.Description.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 31
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Procedure Name'
			$WorksheetData[0,$Col++] = 'Description >>'
			$WorksheetData[0,$Col++] = 'Implementation Type'
			$WorksheetData[0,$Col++] = 'Created Date'
			$WorksheetData[0,$Col++] = 'Last Modified Date'
			$WorksheetData[0,$Col++] = 'System Object'
			$WorksheetData[0,$Col++] = 'Execution As'
			$WorksheetData[0,$Col++] = 'Execution Principal'
			$WorksheetData[0,$Col++] = 'CLR Assembly Name'
			$WorksheetData[0,$Col++] = 'CLR Class Name'
			$WorksheetData[0,$Col++] = 'CLR Method Name'
			$WorksheetData[0,$Col++] = 'Options >>'
			$WorksheetData[0,$Col++] = 'ANSI NULLs'
			$WorksheetData[0,$Col++] = 'Encrypted'
			$WorksheetData[0,$Col++] = 'Is Deterministic'
			$WorksheetData[0,$Col++] = 'Function Type'
			$WorksheetData[0,$Col++] = 'Quoted Identifier'
			$WorksheetData[0,$Col++] = 'Returns NULL On NULL Input'
			$WorksheetData[0,$Col++] = 'Schema Bound'
			$WorksheetData[0,$Col++] = 'Schema Owned'
			$WorksheetData[0,$Col++] = 'Output >>'
			$WorksheetData[0,$Col++] = 'Data Type'
			$WorksheetData[0,$Col++] = 'Length'
			$WorksheetData[0,$Col++] = 'Numeric Precision'
			$WorksheetData[0,$Col++] = 'Numeric Scale'
			$WorksheetData[0,$Col++] = 'System Type'
			$WorksheetData[0,$Col++] = 'Schema'
			$WorksheetData[0,$Col++] = 'XML Document Constraint'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Functions | Where-Object { $_.Properties.General.Description.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Schema
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Name
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.ImplementationType
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.CreateDate
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.DateLastModified
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.IsSystemObject
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.ExecutionContext
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.ExecutionContextPrincipal
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.AssemblyName
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.ClassName
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.MethodName
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.AnsiNulls
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.IsEncrypted
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.IsDeterministic
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.FunctionType
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.QuotedIdentifierStatus
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.ReturnsNullOnNullInput
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.IsSchemaBound
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.IsSchemaOwned
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.DataType.Name
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.DataType.MaximumLength
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.DataType.NumericPrecision
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.DataType.NumericScale
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.DataType.SqlDataType
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.DataType.Schema
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.DataType.XmlDocumentConstraint

						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 26; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 27; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 28; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 37: Functions - Columns
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Functions - Columns"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Functions - Columns'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Functions | ForEach-Object { $_.Columns | Where-Object { $_.General.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 44
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Column Name'
			$WorksheetData[0,$Col++] = 'Binding >>'
			$WorksheetData[0,$Col++] = 'Default Binding'
			$WorksheetData[0,$Col++] = 'Default Schema'
			$WorksheetData[0,$Col++] = 'Rule'
			$WorksheetData[0,$Col++] = 'Rule Schema'
			$WorksheetData[0,$Col++] = 'Computed >>'
			$WorksheetData[0,$Col++] = 'Is Computed'
			$WorksheetData[0,$Col++] = 'Computed Text'
			$WorksheetData[0,$Col++] = 'General >>'
			$WorksheetData[0,$Col++] = 'Allow Nulls'
			$WorksheetData[0,$Col++] = 'ANSI Padding Status'
			$WorksheetData[0,$Col++] = 'Column ID'
			$WorksheetData[0,$Col++] = 'Data Type'
			$WorksheetData[0,$Col++] = 'Length'
			$WorksheetData[0,$Col++] = 'Numeric Precision'
			$WorksheetData[0,$Col++] = 'Numeric Scale'
			$WorksheetData[0,$Col++] = 'Primary Key'
			$WorksheetData[0,$Col++] = 'System Type'
			$WorksheetData[0,$Col++] = 'Identity >>'
			$WorksheetData[0,$Col++] = 'Identity'
			$WorksheetData[0,$Col++] = 'Identity Seed'
			$WorksheetData[0,$Col++] = 'Identity Increment'
			$WorksheetData[0,$Col++] = 'Miscellaneous >>'
			$WorksheetData[0,$Col++] = 'Collation'
			$WorksheetData[0,$Col++] = 'Full Text'
			$WorksheetData[0,$Col++] = 'Not For Replication'
			$WorksheetData[0,$Col++] = 'Statistical Semantics'
			$WorksheetData[0,$Col++] = 'Is Deterministic'
			$WorksheetData[0,$Col++] = 'Is FILESTREAM'
			$WorksheetData[0,$Col++] = 'Is Foreign Key'
			$WorksheetData[0,$Col++] = 'Is Persisted'
			$WorksheetData[0,$Col++] = 'Is Precise'
			$WorksheetData[0,$Col++] = 'Is ROWGUID'
			$WorksheetData[0,$Col++] = 'Sparse >>'
			$WorksheetData[0,$Col++] = 'Is Column Set'
			$WorksheetData[0,$Col++] = 'Is Sparse'
			$WorksheetData[0,$Col++] = 'XML >>'
			$WorksheetData[0,$Col++] = 'XML Schema Namespace'
			$WorksheetData[0,$Col++] = 'XML Schema Namespace Schema'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Functions | Where-Object { $_.Properties.General.Description.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Columns | Where-Object { $_.General.General.ID } | Sort-Object -Property @{Expression={$_.General.General.ID}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.General.Name
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Binding.DefaultBinding
							$WorksheetData[$Row,$Col++] = $_.General.Binding.DefaultSchema
							$WorksheetData[$Row,$Col++] = $_.General.Binding.Rule
							$WorksheetData[$Row,$Col++] = $_.General.Binding.RuleSchema
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Computed.IsComputed
							$WorksheetData[$Row,$Col++] = $_.General.Computed.ComputedText
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.General.AllowNulls
							$WorksheetData[$Row,$Col++] = $_.General.General.AnsiPaddingStatus
							$WorksheetData[$Row,$Col++] = $_.General.General.ID
							$WorksheetData[$Row,$Col++] = $_.General.General.DataType
							$WorksheetData[$Row,$Col++] = $_.General.General.Length
							$WorksheetData[$Row,$Col++] = $_.General.General.NumericPrecision
							$WorksheetData[$Row,$Col++] = $_.General.General.NumericScale
							$WorksheetData[$Row,$Col++] = $_.General.General.InPrimaryKey
							$WorksheetData[$Row,$Col++] = $_.General.General.SystemType
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Identity.IsIdentity
							$WorksheetData[$Row,$Col++] = $_.General.Identity.IdentityIncrement
							$WorksheetData[$Row,$Col++] = $_.General.Identity.IdentitySeed
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.Collation
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsFullTextIndexed
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsNotForReplication
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.StatisticalSemantics
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsDeterministic
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsFileStream
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsForeignKey
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsPersisted
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsPrecise
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsRowGuidCol
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Sparse.IsColumnSet
							$WorksheetData[$Row,$Col++] = $_.General.Sparse.IsSparse
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.XML.XmlSchemaNameSpace
							$WorksheetData[$Row,$Col++] = $_.General.XML.XmlSchemaNameSpaceSchema

							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 19; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 20; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 21; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 26; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 27; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 38: Functions - Default Constraints
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Functions - Default Constraints"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Functions - Default Constraints'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Functions | ForEach-Object { $_.Columns | Where-Object { $_.DefaultConstraint.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 13
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Column Name'
			$WorksheetData[0,$Col++] = 'Column ID'
			$WorksheetData[0,$Col++] = 'Constraint Name'
			$WorksheetData[0,$Col++] = 'Date Created'
			$WorksheetData[0,$Col++] = 'Date Modified'
			$WorksheetData[0,$Col++] = 'File Table'
			$WorksheetData[0,$Col++] = 'System Named'
			$WorksheetData[0,$Col++] = 'Text'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Functions | Where-Object { $_.Properties.General.Description.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Columns | Where-Object { $_.DefaultConstraint.ID } | Sort-Object -Property @{Expression={$_.General.General.ID}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.General.Name
							$WorksheetData[$Row,$Col++] = $_.General.General.ID
							$WorksheetData[$Row,$Col++] = $_.DefaultConstraint.Name
							$WorksheetData[$Row,$Col++] = $_.DefaultConstraint.CreateDate
							$WorksheetData[$Row,$Col++] = $_.DefaultConstraint.DateLastModified
							$WorksheetData[$Row,$Col++] = $_.DefaultConstraint.IsFileTableDefined
							$WorksheetData[$Row,$Col++] = $_.DefaultConstraint.IsSystemNamed
							$WorksheetData[$Row,$Col++] = if ($_.DefaultConstraint.Text.Length -gt 5000) { $_.DefaultConstraint.Text.Substring(0, 4997) + '...' } else { $_.DefaultConstraint.Text } 
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'G2'
					ColumnFormat = @(
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 39: Functions - Checks
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Functions - Checks"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Functions - Checks'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Functions | ForEach-Object { $_.Checks | Where-Object { $_.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 14
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Name'
			$WorksheetData[0,$Col++] = 'Check Name'
			$WorksheetData[0,$Col++] = 'Date Created'
			$WorksheetData[0,$Col++] = 'Date Modified'
			$WorksheetData[0,$Col++] = 'Enabled'
			$WorksheetData[0,$Col++] = 'Checked'
			$WorksheetData[0,$Col++] = 'File Table'
			$WorksheetData[0,$Col++] = 'Not For Replication'
			$WorksheetData[0,$Col++] = 'Is System Named'
			$WorksheetData[0,$Col++] = 'Text'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Functions | Where-Object { $_.Properties.General.Description.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Checks | Where-Object { $_.ID } | Sort-Object -Property Name | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.Name
							$WorksheetData[$Row,$Col++] = $_.CreateDate
							$WorksheetData[$Row,$Col++] = $_.DateLastModified
							$WorksheetData[$Row,$Col++] = $_.IsEnabled
							$WorksheetData[$Row,$Col++] = $_.IsChecked
							$WorksheetData[$Row,$Col++] = $_.IsFileTableDefined
							$WorksheetData[$Row,$Col++] = $_.IsNotForReplication
							$WorksheetData[$Row,$Col++] = $_.IsSystemNamed
							$WorksheetData[$Row,$Col++] = if ($_.Text.Length -gt 5000) { $_.Text.Substring(0, 4997) + '...' } else { $_.Text }
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 40: Functions - Indexes
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Functions - Indexes"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Functions - Indexes'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Functions | ForEach-Object { $_.Indexes | Where-Object { $_.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 26
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Function Name'
			$WorksheetData[0,$Col++] = 'Index Name'
			$WorksheetData[0,$Col++] = 'Index Type'
			$WorksheetData[0,$Col++] = 'Key Type'
			$WorksheetData[0,$Col++] = 'Space Used (MB)'
			$WorksheetData[0,$Col++] = 'Compact Large Objects'
			$WorksheetData[0,$Col++] = 'Has Compressed Partitions'
			$WorksheetData[0,$Col++] = 'Filtered'
			$WorksheetData[0,$Col++] = 'Clustered'
			$WorksheetData[0,$Col++] = 'Disabled'
			$WorksheetData[0,$Col++] = 'Hypothetical'
			$WorksheetData[0,$Col++] = 'File Table'
			$WorksheetData[0,$Col++] = 'Full Text Key'
			$WorksheetData[0,$Col++] = 'On Computed Column'
			$WorksheetData[0,$Col++] = 'On Table'
			$WorksheetData[0,$Col++] = 'Partitioned'
			$WorksheetData[0,$Col++] = 'Spatial'
			$WorksheetData[0,$Col++] = 'System Named'
			$WorksheetData[0,$Col++] = 'System Object'
			$WorksheetData[0,$Col++] = 'Unique'
			$WorksheetData[0,$Col++] = 'XML Index'
			$WorksheetData[0,$Col++] = 'Parent XML Index'
			$WorksheetData[0,$Col++] = 'Secondary XML Index Type'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Functions | Where-Object { $_.Properties.General.Description.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Indexes | Where-Object { $_.General.ID } | Sort-Object -Property @{Expression={$_.General.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.Name
							$WorksheetData[$Row,$Col++] = $_.General.IndexType
							$WorksheetData[$Row,$Col++] = $_.General.IndexKeyType
							$WorksheetData[$Row,$Col++] = $_.General.SpaceUsedKB / 1KB
							$WorksheetData[$Row,$Col++] = $_.General.CompactLargeObjects
							$WorksheetData[$Row,$Col++] = $_.General.HasCompressedPartitions
							$WorksheetData[$Row,$Col++] = $_.General.HasFilter
							$WorksheetData[$Row,$Col++] = $_.General.IsClustered
							$WorksheetData[$Row,$Col++] = $_.General.IsDisabled
							$WorksheetData[$Row,$Col++] = $_.General.IsHypothetical
							$WorksheetData[$Row,$Col++] = $_.General.IsFileTableDefined
							$WorksheetData[$Row,$Col++] = $_.General.IsFullTextKey
							$WorksheetData[$Row,$Col++] = $_.General.IsIndexOnComputed
							$WorksheetData[$Row,$Col++] = $_.General.IsIndexOnTable
							$WorksheetData[$Row,$Col++] = $_.General.IsPartitioned
							$WorksheetData[$Row,$Col++] = $_.General.IsSpatialIndex
							$WorksheetData[$Row,$Col++] = $_.General.IsSystemNamed
							$WorksheetData[$Row,$Col++] = $_.General.IsSystemObject
							$WorksheetData[$Row,$Col++] = $_.General.IsUnique
							$WorksheetData[$Row,$Col++] = $_.General.IsXmlIndex
							$WorksheetData[$Row,$Col++] = $_.General.ParentXmlIndex
							$WorksheetData[$Row,$Col++] = $_.General.SecondaryXmlIndexType
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberS3}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 41: Functions - Index Options
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Functions - Index Options"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Functions - Index Options'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Functions | ForEach-Object { $_.Indexes | Where-Object { $_.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 18
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Function Name'
			$WorksheetData[0,$Col++] = 'Index Name'
			$WorksheetData[0,$Col++] = 'General >>'
			$WorksheetData[0,$Col++] = 'Auto Recompute Statistics'
			$WorksheetData[0,$Col++] = 'Ignore Duplicate Values'
			$WorksheetData[0,$Col++] = 'Locks >>'
			$WorksheetData[0,$Col++] = 'Allow Row Locks'
			$WorksheetData[0,$Col++] = 'Allow Page Locks'
			$WorksheetData[0,$Col++] = 'Operation >>'
			$WorksheetData[0,$Col++] = 'Allow Online DML Processing'
			$WorksheetData[0,$Col++] = 'Max Degree Of Parallelism'
			$WorksheetData[0,$Col++] = 'Storage >>'
			$WorksheetData[0,$Col++] = 'Sort In TempDB'
			$WorksheetData[0,$Col++] = 'Fill Factor'
			$WorksheetData[0,$Col++] = 'Pad Index'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Functions | Where-Object { $_.Properties.General.Description.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Indexes | Where-Object { $_.General.ID } | Sort-Object -Property @{Expression={$_.General.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.Name
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = -not $_.Options.General.NoAutomaticRecomputation
							$WorksheetData[$Row,$Col++] = $_.Options.General.IgnoreDuplicateKeys
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = -not $_.Options.Locks.DisallowPageLocks
							$WorksheetData[$Row,$Col++] = -not $_.Options.Locks.DisallowRowLocks
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.Options.Operation.OnlineIndexOperation
							$WorksheetData[$Row,$Col++] = $_.Options.Operation.MaximumDegreeOfParallelism
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.Options.Storage.SortInTempdb
							$WorksheetData[$Row,$Col++] = $_.Options.Storage.FillFactor
							$WorksheetData[$Row,$Col++] = $_.Options.Storage.PadIndex
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 17; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 42: Functions - Index Storage
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Functions - Index Storage"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Functions - Index Storage'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Functions | ForEach-Object { $_.Indexes | Where-Object { $_.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 10
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Function Name'
			$WorksheetData[0,$Col++] = 'Index Name'
			$WorksheetData[0,$Col++] = 'Filegroup'
			$WorksheetData[0,$Col++] = 'FILESTREAM Filegroup'
			$WorksheetData[0,$Col++] = 'Partition Scheme'
			$WorksheetData[0,$Col++] = 'FILESTREAM Partition Scheme'
			$WorksheetData[0,$Col++] = 'Partition Scheme Parameters'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Functions | Where-Object { $_.Properties.General.Description.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Indexes | Where-Object { $_.General.ID } | Sort-Object -Property @{Expression={$_.General.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.Name
							$WorksheetData[$Row,$Col++] = $_.Storage.FileGroup
							$WorksheetData[$Row,$Col++] = $_.Storage.FileStreamFileGroup
							$WorksheetData[$Row,$Col++] = $_.Storage.PartitionScheme
							$WorksheetData[$Row,$Col++] = $_.Storage.FileStreamPartitionScheme
							$WorksheetData[$Row,$Col++] = $($_.Storage.PartitionSchemeParameters | Where-Object { $_.ID } | ForEach-Object { $_.Name }) -join $Delimiter
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 43: Functions - Parameters
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Functions - Parameters"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Functions - Parameters'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = (
				$SqlServerInventory.DatabaseServer | ForEach-Object { 
					$_.Server.Databases | ForEach-Object { 
						$_.Programmability.Functions | Where-Object { $_.Properties.General.Description.ID } | ForEach-Object { 
							$_.Parameters | Where-Object { $_.ID } 
						} 
					} 
				} | Measure-Object
			).Count + 1
			$ColumnCount = 17
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Procedure Name'
			$WorksheetData[0,$Col++] = 'Parameter Name'
			$WorksheetData[0,$Col++] = 'Parameter ID'
			$WorksheetData[0,$Col++] = 'Cursor Parameter'
			$WorksheetData[0,$Col++] = 'Output Parameter'
			$WorksheetData[0,$Col++] = 'Read Only'
			$WorksheetData[0,$Col++] = 'Default Value'
			$WorksheetData[0,$Col++] = 'Data Type'
			$WorksheetData[0,$Col++] = 'Length'
			$WorksheetData[0,$Col++] = 'Numeric Precision'
			$WorksheetData[0,$Col++] = 'Numeric Scale'
			$WorksheetData[0,$Col++] = 'System Type'
			$WorksheetData[0,$Col++] = 'Schema'
			$WorksheetData[0,$Col++] = 'XML Document Constraint'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Functions | Where-Object { $_.Properties.General.Description.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Parameters | Where-Object { $_.ID } | Sort-Object -Property ID | ForEach-Object {
							$Col = 0 
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.Name
							$WorksheetData[$Row,$Col++] = $_.ID
							$WorksheetData[$Row,$Col++] = $_.IsCursorParameter
							$WorksheetData[$Row,$Col++] = $_.IsOutputParameter
							$WorksheetData[$Row,$Col++] = $_.IsReadOnly
							$WorksheetData[$Row,$Col++] = $_.DefaultValue
							$WorksheetData[$Row,$Col++] = $_.DataType.Name
							$WorksheetData[$Row,$Col++] = $_.DataType.MaximumLength
							$WorksheetData[$Row,$Col++] = $_.DataType.NumericPrecision
							$WorksheetData[$Row,$Col++] = $_.DataType.NumericScale
							$WorksheetData[$Row,$Col++] = $_.DataType.SqlDataType
							$WorksheetData[$Row,$Col++] = $_.DataType.Schema
							$WorksheetData[$Row,$Col++] = $_.DataType.XmlDocumentConstraint
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtText},
						@{ColumnNumber = 12; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 14; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 44: Functions - Definition
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Functions - Definition"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Functions - Definition'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Functions | Where-Object { $_.Properties.General.Description.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 5
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Procedure Name'
			$WorksheetData[0,$Col++] = 'Definition'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Functions | Where-Object { $_.Properties.General.Description.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Schema
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Name
						$WorksheetData[$Row,$Col++] = if ($_.Definition.Length -gt 5000) { $_.Definition.Substring(0, 4997) + '...' } else { $_.Definition }
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 45: Database Triggers
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Triggers"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Database Triggers'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.DatabaseTriggers | Where-Object { $_.ID } } } | Measure-Object).Count + 1

			$ColumnCount = 18
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Trigger Name'
			$WorksheetData[0,$Col++] = 'Enabled'
			$WorksheetData[0,$Col++] = 'Encrypted'
			$WorksheetData[0,$Col++] = 'System Object'
			$WorksheetData[0,$Col++] = 'Events'
			$WorksheetData[0,$Col++] = 'Execution Context'
			$WorksheetData[0,$Col++] = 'Execution User'
			$WorksheetData[0,$Col++] = 'ANSI Nulls Status'
			$WorksheetData[0,$Col++] = 'Quoted Identifier Status'
			$WorksheetData[0,$Col++] = 'Implementation Type'
			$WorksheetData[0,$Col++] = 'CLR Assembly Name'
			$WorksheetData[0,$Col++] = 'CLR Class Name'
			$WorksheetData[0,$Col++] = 'CLR Method Name'
			$WorksheetData[0,$Col++] = 'Create Date'
			$WorksheetData[0,$Col++] = 'Modified Date'
			$WorksheetData[0,$Col++] = 'Not For Replication'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.DatabaseTriggers | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Schema}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.IsEnabled
						$WorksheetData[$Row,$Col++] = $_.IsEncrypted
						$WorksheetData[$Row,$Col++] = $_.IsSystemObject
						$WorksheetData[$Row,$Col++] = $_.DdlTriggerEvents
						$WorksheetData[$Row,$Col++] = $_.ExecutionContext
						$WorksheetData[$Row,$Col++] = $_.ExecutionContextUser
						$WorksheetData[$Row,$Col++] = $_.AnsiNullsStatus
						$WorksheetData[$Row,$Col++] = $_.QuotedIdentifierStatus
						$WorksheetData[$Row,$Col++] = $_.ImplementationType
						$WorksheetData[$Row,$Col++] = $_.AssemblyName
						$WorksheetData[$Row,$Col++] = $_.ClassName
						$WorksheetData[$Row,$Col++] = $_.MethodName
						$WorksheetData[$Row,$Col++] = $_.CreateDate
						$WorksheetData[$Row,$Col++] = $_.DateLastModified
						$WorksheetData[$Row,$Col++] = $_.NotForReplication
						$Row++
					}
				}

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'D2'
					ColumnFormat = @(
						@{ColumnNumber = 16; NumberFormat = $XlNumFmtDate}
						@{ColumnNumber = 17; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 46: Database Triggers - Definition
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Database Triggers - Definition"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Database Triggers - Definition'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor #$ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.DatabaseTriggers | Where-Object { $_.ID } } } | Measure-Object).Count + 1

			$ColumnCount = 18
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Trigger Name'
			$WorksheetData[0,$Col++] = 'Definition'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.DatabaseTriggers | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Schema}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = if ($_.Definition.Length -gt 5000) { $_.Definition.Substring(0, 4997) + '...' } else { $_.Definition }
						$Row++
					}
				}

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

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


			# Worksheet 47: Assemblies
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Assemblies"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Assemblies'
			#$Worksheet.Tab.Color = $ServerTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor #$ServerTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Assemblies | Where-Object { $_.Properties.General.ID } } } | Measure-Object).Count + 1

			$ColumnCount = 12
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Assembly Name'
			$WorksheetData[0,$Col++] = 'Assembly Owner'
			$WorksheetData[0,$Col++] = 'Permission Set'
			$WorksheetData[0,$Col++] = 'System Object'
			$WorksheetData[0,$Col++] = 'Visible'
			$WorksheetData[0,$Col++] = 'Date Created'
			$WorksheetData[0,$Col++] = 'Culture'
			$WorksheetData[0,$Col++] = 'Version'
			$WorksheetData[0,$Col++] = 'Path To Assembly'
			$WorksheetData[0,$Col++] = 'Public Key'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Assemblies | Where-Object { $_.Properties.General.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Owner}}, @{Expression={$_.Properties.General.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Name
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Owner
						$WorksheetData[$Row,$Col++] = $_.Properties.General.AssemblySecurityLevel
						$WorksheetData[$Row,$Col++] = $_.Properties.General.IsSystemObject
						$WorksheetData[$Row,$Col++] = $_.Properties.General.IsVisible
						$WorksheetData[$Row,$Col++] = $_.Properties.General.CreateDate
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Culture
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Version
						$WorksheetData[$Row,$Col++] = $($_.Properties.General.SqlAssemblyFiles | Sort-Object -Property ID | ForEach-Object { $_.Name }) -join $Delimiter
						$WorksheetData[$Row,$Col++] = $_.Properties.General.PublicKey
						$Row++
					}
				}

			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 48: Types - Aggregates
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Types - Aggregates"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Types - Aggregates'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Types.UserDefinedAggregates | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 18
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Owner Name'
			$WorksheetData[0,$Col++] = 'Aggregate Name'
			$WorksheetData[0,$Col++] = 'Created Date'
			$WorksheetData[0,$Col++] = 'Last Modified Date'
			$WorksheetData[0,$Col++] = 'CLR Assembly Name'
			$WorksheetData[0,$Col++] = 'CLR Class Name'
			$WorksheetData[0,$Col++] = 'Schema Owned'
			$WorksheetData[0,$Col++] = 'Output >>'
			$WorksheetData[0,$Col++] = 'Data Type'
			$WorksheetData[0,$Col++] = 'Length'
			$WorksheetData[0,$Col++] = 'Numeric Precision'
			$WorksheetData[0,$Col++] = 'Numeric Scale'
			$WorksheetData[0,$Col++] = 'System Type'
			$WorksheetData[0,$Col++] = 'Schema'
			$WorksheetData[0,$Col++] = 'XML Document Constraint'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Types.UserDefinedAggregates | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Schema}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Schema
						$WorksheetData[$Row,$Col++] = $_.Owner
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.CreateDate
						$WorksheetData[$Row,$Col++] = $_.DateLastModified
						$WorksheetData[$Row,$Col++] = $_.AssemblyName
						$WorksheetData[$Row,$Col++] = $_.ClassName
						$WorksheetData[$Row,$Col++] = $_.IsSchemaOwned
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.DataType.Name
						$WorksheetData[$Row,$Col++] = $_.DataType.MaximumLength
						$WorksheetData[$Row,$Col++] = $_.DataType.NumericPrecision
						$WorksheetData[$Row,$Col++] = $_.DataType.NumericScale
						$WorksheetData[$Row,$Col++] = $_.DataType.SqlDataType
						$WorksheetData[$Row,$Col++] = $_.DataType.Schema
						$WorksheetData[$Row,$Col++] = $_.DataType.XmlDocumentConstraint
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 14; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 49: Types - Aggregates Parameters
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Types - Aggregates Parameters"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Types - Aggregates Parameters'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Types.UserDefinedAggregates | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 13
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Aggregate Name'
			$WorksheetData[0,$Col++] = 'Parameter Name'
			$WorksheetData[0,$Col++] = 'Parameter ID'
			$WorksheetData[0,$Col++] = 'Data Type'
			$WorksheetData[0,$Col++] = 'Length'
			$WorksheetData[0,$Col++] = 'Numeric Precision'
			$WorksheetData[0,$Col++] = 'Numeric Scale'
			$WorksheetData[0,$Col++] = 'System Type'
			$WorksheetData[0,$Col++] = 'Schema'
			$WorksheetData[0,$Col++] = 'XML Document Constraint'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Types.UserDefinedAggregates | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Schema}}, @{Expression={$_.Name}} | ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Parameters | Where-Object { $_.ID } | Sort-Object -Property ID | ForEach-Object {
							$Col = 0 
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.Name
							$WorksheetData[$Row,$Col++] = $_.ID
							$WorksheetData[$Row,$Col++] = $_.DataType.Name
							$WorksheetData[$Row,$Col++] = $_.DataType.MaximumLength
							$WorksheetData[$Row,$Col++] = $_.DataType.NumericPrecision
							$WorksheetData[$Row,$Col++] = $_.DataType.NumericScale
							$WorksheetData[$Row,$Col++] = $_.DataType.SqlDataType
							$WorksheetData[$Row,$Col++] = $_.DataType.Schema
							$WorksheetData[$Row,$Col++] = $_.DataType.XmlDocumentConstraint
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 50: Types - Data Types
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Types - Data Types"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Types - Data Types'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Types.UserDefinedDataTypes | Where-Object { $_.Properties.General.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 21
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Owner Name'
			$WorksheetData[0,$Col++] = 'Data Type Name'
			$WorksheetData[0,$Col++] = 'General >>'
			$WorksheetData[0,$Col++] = 'Data Type'
			$WorksheetData[0,$Col++] = 'Length'
			$WorksheetData[0,$Col++] = 'Max Length'
			$WorksheetData[0,$Col++] = 'Variable Length'
			$WorksheetData[0,$Col++] = 'Numeric Precision'
			$WorksheetData[0,$Col++] = 'Numeric Scale'
			$WorksheetData[0,$Col++] = 'Allow NULLs'
			$WorksheetData[0,$Col++] = 'Allow Identity'
			$WorksheetData[0,$Col++] = 'Collation'
			$WorksheetData[0,$Col++] = 'Schema Owned'
			$WorksheetData[0,$Col++] = 'Binding >>'
			$WorksheetData[0,$Col++] = 'Default'
			$WorksheetData[0,$Col++] = 'Default Schema'
			$WorksheetData[0,$Col++] = 'Rule'
			$WorksheetData[0,$Col++] = 'Rule Schema'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Types.UserDefinedDataTypes | Where-Object { $_.Properties.General.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Schema}}, @{Expression={$_.Properties.General.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Schema
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Owner
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Name
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.General.SystemType
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Length
						$WorksheetData[$Row,$Col++] = $_.Properties.General.MaxLength
						$WorksheetData[$Row,$Col++] = $_.Properties.General.VariableLength
						$WorksheetData[$Row,$Col++] = $_.Properties.General.NumericPrecision
						$WorksheetData[$Row,$Col++] = $_.Properties.General.NumericScale
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Nullable
						$WorksheetData[$Row,$Col++] = $_.Properties.General.AllowIdentity
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Collation
						$WorksheetData[$Row,$Col++] = $_.Properties.General.IsSchemaOwned
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Default
						$WorksheetData[$Row,$Col++] = $_.Properties.General.DefaultSchema
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Rule
						$WorksheetData[$Row,$Col++] = $_.Properties.General.RuleSchema
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 11; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 12; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 51: Types - Table Types
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Types - Table Types"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Types - Table Types'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Types.UserDefinedTableTypes | Where-Object { $_.Properties.General.Description.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 13
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Owner Name'
			$WorksheetData[0,$Col++] = 'Table Type Name'
			$WorksheetData[0,$Col++] = 'Description >>'
			$WorksheetData[0,$Col++] = 'Created Date'
			$WorksheetData[0,$Col++] = 'Last Modified Date'
			$WorksheetData[0,$Col++] = 'System Object'
			$WorksheetData[0,$Col++] = 'Is Schema Owned'
			$WorksheetData[0,$Col++] = 'Options >>'
			$WorksheetData[0,$Col++] = 'Max Length'
			$WorksheetData[0,$Col++] = 'Nullable'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Types.UserDefinedTableTypes | 
					Where-Object { $_.Properties.General.Description.ID } | 
					Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | 
					ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Schema
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Owner
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.Name
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.CreateDate
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.DateLastModified
						$WorksheetData[$Row,$Col++] = -not $_.Properties.General.Description.IsUserDefined
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Description.IsSchemaOwned
						$WorksheetData[$Row,$Col++] = $null
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.MaxLength
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Options.Nullable
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 52: Types - Table Types Columns
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Types - Table Types Columns"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Types - Table Types Columns'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Types.UserDefinedTableTypes | ForEach-Object { $_.Columns | Where-Object { $_.General.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 44
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Type Name'
			$WorksheetData[0,$Col++] = 'Column Name'
			$WorksheetData[0,$Col++] = 'Binding >>'
			$WorksheetData[0,$Col++] = 'Default Binding'
			$WorksheetData[0,$Col++] = 'Default Schema'
			$WorksheetData[0,$Col++] = 'Rule'
			$WorksheetData[0,$Col++] = 'Rule Schema'
			$WorksheetData[0,$Col++] = 'Computed >>'
			$WorksheetData[0,$Col++] = 'Is Computed'
			$WorksheetData[0,$Col++] = 'Computed Text'
			$WorksheetData[0,$Col++] = 'General >>'
			$WorksheetData[0,$Col++] = 'Allow Nulls'
			$WorksheetData[0,$Col++] = 'ANSI Padding Status'
			$WorksheetData[0,$Col++] = 'Column ID'
			$WorksheetData[0,$Col++] = 'Data Type'
			$WorksheetData[0,$Col++] = 'Length'
			$WorksheetData[0,$Col++] = 'Numeric Precision'
			$WorksheetData[0,$Col++] = 'Numeric Scale'
			$WorksheetData[0,$Col++] = 'Primary Key'
			$WorksheetData[0,$Col++] = 'System Type'
			$WorksheetData[0,$Col++] = 'Identity >>'
			$WorksheetData[0,$Col++] = 'Identity'
			$WorksheetData[0,$Col++] = 'Identity Seed'
			$WorksheetData[0,$Col++] = 'Identity Increment'
			$WorksheetData[0,$Col++] = 'Miscellaneous >>'
			$WorksheetData[0,$Col++] = 'Collation'
			$WorksheetData[0,$Col++] = 'Full Text'
			$WorksheetData[0,$Col++] = 'Not For Replication'
			$WorksheetData[0,$Col++] = 'Statistical Semantics'
			$WorksheetData[0,$Col++] = 'Is Deterministic'
			$WorksheetData[0,$Col++] = 'Is FILESTREAM'
			$WorksheetData[0,$Col++] = 'Is Foreign Key'
			$WorksheetData[0,$Col++] = 'Is Persisted'
			$WorksheetData[0,$Col++] = 'Is Precise'
			$WorksheetData[0,$Col++] = 'Is ROWGUID'
			$WorksheetData[0,$Col++] = 'Sparse >>'
			$WorksheetData[0,$Col++] = 'Is Column Set'
			$WorksheetData[0,$Col++] = 'Is Sparse'
			$WorksheetData[0,$Col++] = 'XML >>'
			$WorksheetData[0,$Col++] = 'XML Schema Namespace'
			$WorksheetData[0,$Col++] = 'XML Schema Namespace Schema'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Types.UserDefinedTableTypes | 
					Where-Object { $_.Properties.General.Description.ID } | 
					Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | 
					ForEach-Object {

						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Columns | Where-Object { $_.General.General.ID } | Sort-Object -Property @{Expression={$_.General.General.ID}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.General.Name
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Binding.DefaultBinding
							$WorksheetData[$Row,$Col++] = $_.General.Binding.DefaultSchema
							$WorksheetData[$Row,$Col++] = $_.General.Binding.Rule
							$WorksheetData[$Row,$Col++] = $_.General.Binding.RuleSchema
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Computed.IsComputed
							$WorksheetData[$Row,$Col++] = $_.General.Computed.ComputedText
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.General.AllowNulls
							$WorksheetData[$Row,$Col++] = $_.General.General.AnsiPaddingStatus
							$WorksheetData[$Row,$Col++] = $_.General.General.ID
							$WorksheetData[$Row,$Col++] = $_.General.General.DataType
							$WorksheetData[$Row,$Col++] = $_.General.General.Length
							$WorksheetData[$Row,$Col++] = $_.General.General.NumericPrecision
							$WorksheetData[$Row,$Col++] = $_.General.General.NumericScale
							$WorksheetData[$Row,$Col++] = $_.General.General.InPrimaryKey
							$WorksheetData[$Row,$Col++] = $_.General.General.SystemType
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Identity.IsIdentity
							$WorksheetData[$Row,$Col++] = $_.General.Identity.IdentityIncrement
							$WorksheetData[$Row,$Col++] = $_.General.Identity.IdentitySeed
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.Collation
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsFullTextIndexed
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsNotForReplication
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.StatisticalSemantics
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsDeterministic
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsFileStream
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsForeignKey
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsPersisted
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsPrecise
							$WorksheetData[$Row,$Col++] = $_.General.Miscellaneous.IsRowGuidCol
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.Sparse.IsColumnSet
							$WorksheetData[$Row,$Col++] = $_.General.Sparse.IsSparse
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.General.XML.XmlSchemaNameSpace
							$WorksheetData[$Row,$Col++] = $_.General.XML.XmlSchemaNameSpaceSchema

							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 19; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 20; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 21; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 26; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 27; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 53: Types - Table Types Default Constraints
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Types - Table Types Default Constraints"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Types - Table Types DF CSTR'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Types.UserDefinedTableTypes | ForEach-Object { $_.Columns | Where-Object { $_.DefaultConstraint.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 13
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Type Name'
			$WorksheetData[0,$Col++] = 'Column Name'
			$WorksheetData[0,$Col++] = 'Column ID'
			$WorksheetData[0,$Col++] = 'Constraint Name'
			$WorksheetData[0,$Col++] = 'Date Created'
			$WorksheetData[0,$Col++] = 'Date Modified'
			$WorksheetData[0,$Col++] = 'File Table'
			$WorksheetData[0,$Col++] = 'System Named'
			$WorksheetData[0,$Col++] = 'Text'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Types.UserDefinedTableTypes | 
					Where-Object { $_.Properties.General.Description.ID } | 
					Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | 
					ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Columns | Where-Object { $_.DefaultConstraint.ID } | Sort-Object -Property @{Expression={$_.General.General.ID}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.General.Name
							$WorksheetData[$Row,$Col++] = $_.General.General.ID
							$WorksheetData[$Row,$Col++] = $_.DefaultConstraint.Name
							$WorksheetData[$Row,$Col++] = $_.DefaultConstraint.CreateDate
							$WorksheetData[$Row,$Col++] = $_.DefaultConstraint.DateLastModified
							$WorksheetData[$Row,$Col++] = $_.DefaultConstraint.IsFileTableDefined
							$WorksheetData[$Row,$Col++] = $_.DefaultConstraint.IsSystemNamed
							$WorksheetData[$Row,$Col++] = if ($_.DefaultConstraint.Text.Length -gt 5000) { $_.DefaultConstraint.Text.Substring(0, 4997) + '...' } else { $_.DefaultConstraint.Text } 
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'G2'
					ColumnFormat = @(
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 54: Types - Table Types Checks
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Types - Table Types Checks"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Types - Table Types Checks'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Types.UserDefinedTableTypes | ForEach-Object { $_.Checks | Where-Object { $_.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 14
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Type Name'
			$WorksheetData[0,$Col++] = 'Check Name'
			$WorksheetData[0,$Col++] = 'Date Created'
			$WorksheetData[0,$Col++] = 'Date Modified'
			$WorksheetData[0,$Col++] = 'Enabled'
			$WorksheetData[0,$Col++] = 'Checked'
			$WorksheetData[0,$Col++] = 'File Table'
			$WorksheetData[0,$Col++] = 'Not For Replication'
			$WorksheetData[0,$Col++] = 'Is System Named'
			$WorksheetData[0,$Col++] = 'Text'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Types.UserDefinedTableTypes | 
					Where-Object { $_.Properties.General.Description.ID } | 
					Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | 
					ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Checks | Where-Object { $_.ID } | Sort-Object -Property Name | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.Name
							$WorksheetData[$Row,$Col++] = $_.CreateDate
							$WorksheetData[$Row,$Col++] = $_.DateLastModified
							$WorksheetData[$Row,$Col++] = $_.IsEnabled
							$WorksheetData[$Row,$Col++] = $_.IsChecked
							$WorksheetData[$Row,$Col++] = $_.IsFileTableDefined
							$WorksheetData[$Row,$Col++] = $_.IsNotForReplication
							$WorksheetData[$Row,$Col++] = $_.IsSystemNamed
							$WorksheetData[$Row,$Col++] = if ($_.Text.Length -gt 5000) { $_.Text.Substring(0, 4997) + '...' } else { $_.Text }
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 55: Types - Table Types Indexes
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Types - Table Types Indexes"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Types - Table Types Indexes'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Types.UserDefinedTableTypes | ForEach-Object { $_.Indexes | Where-Object { $_.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 26
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Type Name'
			$WorksheetData[0,$Col++] = 'Index Name'
			$WorksheetData[0,$Col++] = 'Index Type'
			$WorksheetData[0,$Col++] = 'Key Type'
			$WorksheetData[0,$Col++] = 'Space Used (MB)'
			$WorksheetData[0,$Col++] = 'Compact Large Objects'
			$WorksheetData[0,$Col++] = 'Has Compressed Partitions'
			$WorksheetData[0,$Col++] = 'Filtered'
			$WorksheetData[0,$Col++] = 'Clustered'
			$WorksheetData[0,$Col++] = 'Disabled'
			$WorksheetData[0,$Col++] = 'Hypothetical'
			$WorksheetData[0,$Col++] = 'File Table'
			$WorksheetData[0,$Col++] = 'Full Text Key'
			$WorksheetData[0,$Col++] = 'On Computed Column'
			$WorksheetData[0,$Col++] = 'On Table'
			$WorksheetData[0,$Col++] = 'Partitioned'
			$WorksheetData[0,$Col++] = 'Spatial'
			$WorksheetData[0,$Col++] = 'System Named'
			$WorksheetData[0,$Col++] = 'System Object'
			$WorksheetData[0,$Col++] = 'Unique'
			$WorksheetData[0,$Col++] = 'XML Index'
			$WorksheetData[0,$Col++] = 'Parent XML Index'
			$WorksheetData[0,$Col++] = 'Secondary XML Index Type'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Types.UserDefinedTableTypes | 
					Where-Object { $_.Properties.General.Description.ID } | 
					Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | 
					ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Indexes | Where-Object { $_.General.ID } | Sort-Object -Property @{Expression={$_.General.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.Name
							$WorksheetData[$Row,$Col++] = $_.General.IndexType
							$WorksheetData[$Row,$Col++] = $_.General.IndexKeyType
							$WorksheetData[$Row,$Col++] = $_.General.SpaceUsedKB / 1KB
							$WorksheetData[$Row,$Col++] = $_.General.CompactLargeObjects
							$WorksheetData[$Row,$Col++] = $_.General.HasCompressedPartitions
							$WorksheetData[$Row,$Col++] = $_.General.HasFilter
							$WorksheetData[$Row,$Col++] = $_.General.IsClustered
							$WorksheetData[$Row,$Col++] = $_.General.IsDisabled
							$WorksheetData[$Row,$Col++] = $_.General.IsHypothetical
							$WorksheetData[$Row,$Col++] = $_.General.IsFileTableDefined
							$WorksheetData[$Row,$Col++] = $_.General.IsFullTextKey
							$WorksheetData[$Row,$Col++] = $_.General.IsIndexOnComputed
							$WorksheetData[$Row,$Col++] = $_.General.IsIndexOnTable
							$WorksheetData[$Row,$Col++] = $_.General.IsPartitioned
							$WorksheetData[$Row,$Col++] = $_.General.IsSpatialIndex
							$WorksheetData[$Row,$Col++] = $_.General.IsSystemNamed
							$WorksheetData[$Row,$Col++] = $_.General.IsSystemObject
							$WorksheetData[$Row,$Col++] = $_.General.IsUnique
							$WorksheetData[$Row,$Col++] = $_.General.IsXmlIndex
							$WorksheetData[$Row,$Col++] = $_.General.ParentXmlIndex
							$WorksheetData[$Row,$Col++] = $_.General.SecondaryXmlIndexType
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberS3}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 56: Types - Table Types Index Options
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Types - Table Types Index Options"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Types - Table Types Idx Options'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Types.UserDefinedTableTypes | ForEach-Object { $_.Indexes | Where-Object { $_.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 18
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Type Name'
			$WorksheetData[0,$Col++] = 'Index Name'
			$WorksheetData[0,$Col++] = 'General >>'
			$WorksheetData[0,$Col++] = 'Auto Recompute Statistics'
			$WorksheetData[0,$Col++] = 'Ignore Duplicate Values'
			$WorksheetData[0,$Col++] = 'Locks >>'
			$WorksheetData[0,$Col++] = 'Allow Row Locks'
			$WorksheetData[0,$Col++] = 'Allow Page Locks'
			$WorksheetData[0,$Col++] = 'Operation >>'
			$WorksheetData[0,$Col++] = 'Allow Online DML Processing'
			$WorksheetData[0,$Col++] = 'Max Degree Of Parallelism'
			$WorksheetData[0,$Col++] = 'Storage >>'
			$WorksheetData[0,$Col++] = 'Sort In TempDB'
			$WorksheetData[0,$Col++] = 'Fill Factor'
			$WorksheetData[0,$Col++] = 'Pad Index'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Types.UserDefinedTableTypes | 
					Where-Object { $_.Properties.General.Description.ID } | 
					Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | 
					ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Indexes | Where-Object { $_.General.ID } | Sort-Object -Property @{Expression={$_.General.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.Name
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = -not $_.Options.General.NoAutomaticRecomputation
							$WorksheetData[$Row,$Col++] = $_.Options.General.IgnoreDuplicateKeys
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = -not $_.Options.Locks.DisallowPageLocks
							$WorksheetData[$Row,$Col++] = -not $_.Options.Locks.DisallowRowLocks
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.Options.Operation.OnlineIndexOperation
							$WorksheetData[$Row,$Col++] = $_.Options.Operation.MaximumDegreeOfParallelism
							$WorksheetData[$Row,$Col++] = $null
							$WorksheetData[$Row,$Col++] = $_.Options.Storage.SortInTempdb
							$WorksheetData[$Row,$Col++] = $_.Options.Storage.FillFactor
							$WorksheetData[$Row,$Col++] = $_.Options.Storage.PadIndex
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 17; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 57: Types - Table Types Index Storage
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Types - Table Types Index Storage"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Types - Table Types Idx Storage'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Types.UserDefinedTableTypes | ForEach-Object { $_.Indexes | Where-Object { $_.General.ID } } } } | Measure-Object).Count + 1
			$ColumnCount = 10
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Table Type Name'
			$WorksheetData[0,$Col++] = 'Index Name'
			$WorksheetData[0,$Col++] = 'Filegroup'
			$WorksheetData[0,$Col++] = 'FILESTREAM Filegroup'
			$WorksheetData[0,$Col++] = 'Partition Scheme'
			$WorksheetData[0,$Col++] = 'FILESTREAM Partition Scheme'
			$WorksheetData[0,$Col++] = 'Partition Scheme Parameters'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Types.UserDefinedTableTypes | 
					Where-Object { $_.Properties.General.Description.ID } | 
					Sort-Object -Property @{Expression={$_.Properties.General.Description.Schema}}, @{Expression={$_.Properties.General.Description.Name}} | 
					ForEach-Object {
						$SchemaName = $_.Properties.General.Description.Schema
						$ObjectName = $_.Properties.General.Description.Name

						$_.Indexes | Where-Object { $_.General.ID } | Sort-Object -Property @{Expression={$_.General.Name}} | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $_.General.Name
							$WorksheetData[$Row,$Col++] = $_.Storage.FileGroup
							$WorksheetData[$Row,$Col++] = $_.Storage.FileStreamFileGroup
							$WorksheetData[$Row,$Col++] = $_.Storage.PartitionScheme
							$WorksheetData[$Row,$Col++] = $_.Storage.FileStreamPartitionScheme
							$WorksheetData[$Row,$Col++] = $($_.Storage.PartitionSchemeParameters | Where-Object { $_.ID } | ForEach-Object { $_.Name }) -join $Delimiter
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 58: Types - CLR Types
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Types - CLR Types"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Types - CLR Types'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Types.UserDefinedTypes | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 18
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Owner Name'
			$WorksheetData[0,$Col++] = 'CLR Type Name'
			$WorksheetData[0,$Col++] = 'CLR Assembly Name'
			$WorksheetData[0,$Col++] = 'CLR Class Name'
			$WorksheetData[0,$Col++] = 'Length'
			$WorksheetData[0,$Col++] = 'Numeric Precision'
			$WorksheetData[0,$Col++] = 'Numeric Scale'
			$WorksheetData[0,$Col++] = 'Collation'
			$WorksheetData[0,$Col++] = 'Binary Ordered'
			$WorksheetData[0,$Col++] = 'COM Visible'
			$WorksheetData[0,$Col++] = 'Fixed Length'
			$WorksheetData[0,$Col++] = 'Nullable'
			$WorksheetData[0,$Col++] = 'Schema Owned'
			$WorksheetData[0,$Col++] = 'User Defined Type Format'
			$WorksheetData[0,$Col++] = 'Binary Type Identifier'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Types.UserDefinedTypes | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Schema}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Schema
						$WorksheetData[$Row,$Col++] = $_.Owner
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.AssemblyName
						$WorksheetData[$Row,$Col++] = $_.ClassName
						$WorksheetData[$Row,$Col++] = $_.MaxLength
						$WorksheetData[$Row,$Col++] = $_.NumericPrecision
						$WorksheetData[$Row,$Col++] = $_.NumericScale
						$WorksheetData[$Row,$Col++] = $_.Collation
						$WorksheetData[$Row,$Col++] = $_.IsBinaryOrdered
						$WorksheetData[$Row,$Col++] = $_.IsComVisible
						$WorksheetData[$Row,$Col++] = $_.IsFixedLength
						$WorksheetData[$Row,$Col++] = $_.IsNullable
						$WorksheetData[$Row,$Col++] = $_.IsSchemaOwned
						$WorksheetData[$Row,$Col++] = $_.UserDefinedTypeFormat
						$WorksheetData[$Row,$Col++] = $_.BinaryTypeIdentifier
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 59: Types - XML Schema Collections
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Types - XML Schema Collections"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Types - XML Schema Collections'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Types.XmlSchemaCollections | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 7
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'XML Schema Collection Name'
			$WorksheetData[0,$Col++] = 'Created Date'
			$WorksheetData[0,$Col++] = 'Last Modified Date'
			$WorksheetData[0,$Col++] = 'Definition'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Types.XmlSchemaCollections | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Schema}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Schema
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.CreateDate
						$WorksheetData[$Row,$Col++] = $_.DateLastModified
						$WorksheetData[$Row,$Col++] = if ($_.Definition.Length -gt 5000) { $_.Definition.Substring(0, 4997) + '...' } else { $_.Definition }
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 60: Rules
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Rules"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Rules'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Rules | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 7
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Rule Name'
			$WorksheetData[0,$Col++] = 'Created Date'
			$WorksheetData[0,$Col++] = 'Last Modified Date'
			$WorksheetData[0,$Col++] = 'Definition'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Rules | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Schema}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Schema
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.CreateDate
						$WorksheetData[$Row,$Col++] = $_.DateLastModified
						$WorksheetData[$Row,$Col++] = if ($_.Definition.Length -gt 5000) { $_.Definition.Substring(0, 4997) + '...' } else { $_.Definition }
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 61: Defaults
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Defaults"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Defaults'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Defaults | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 5
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Default Name'
			$WorksheetData[0,$Col++] = 'Definition'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Defaults | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Schema}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Schema
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = if ($_.Definition.Length -gt 5000) { $_.Definition.Substring(0, 4997) + '...' } else { $_.Definition }
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 62: Plan Guides
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Plan Guides"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Plan Guides'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.PlanGuides | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 12
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Plan Guide Name'
			$WorksheetData[0,$Col++] = 'Enabled'
			$WorksheetData[0,$Col++] = 'Scope Type'
			$WorksheetData[0,$Col++] = 'Scope Batch'
			$WorksheetData[0,$Col++] = 'Scope Schema Name'
			$WorksheetData[0,$Col++] = 'Scope Object Name'
			$WorksheetData[0,$Col++] = 'Parameters'
			$WorksheetData[0,$Col++] = 'Hints'
			$WorksheetData[0,$Col++] = 'Statement'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.PlanGuides | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Schema}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Schema
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = -not $_.IsDisabled
						$WorksheetData[$Row,$Col++] = $_.ScopeType
						$WorksheetData[$Row,$Col++] = if ($_.ScopeBatch.Length -gt 5000) { $_.ScopeBatch.Substring(0, 4997) + '...' } else { $_.ScopeBatch }
						$WorksheetData[$Row,$Col++] = $_.ScopeSchemaName
						$WorksheetData[$Row,$Col++] = $_.ScopeObjectName
						$WorksheetData[$Row,$Col++] = $_.Parameters
						$WorksheetData[$Row,$Col++] = if ($_.Hints.Length -gt 5000) { $_.Hints.Substring(0, 4997) + '...' } else { $_.Hints }
						$WorksheetData[$Row,$Col++] = if ($_.Statement.Length -gt 5000) { $_.Statement.Substring(0, 4997) + '...' } else { $_.Statement }
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 63: Sequences
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Sequences"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Sequences'
			#$Worksheet.Tab.Color = $ServerObjectsTabColor
			$Worksheet.Tab.ThemeColor = $ServerObjectsTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Programmability.Sequences | Where-Object { $_.Properties.General.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 18
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Schema Name'
			$WorksheetData[0,$Col++] = 'Owner'
			$WorksheetData[0,$Col++] = 'Sequence Name'
			$WorksheetData[0,$Col++] = 'Created Date'
			$WorksheetData[0,$Col++] = 'Last Modified Date'
			$WorksheetData[0,$Col++] = 'Data Type'
			$WorksheetData[0,$Col++] = 'Numeric Precision'
			$WorksheetData[0,$Col++] = 'Start Value'
			$WorksheetData[0,$Col++] = 'Increment By'
			$WorksheetData[0,$Col++] = 'Minimum Value'
			$WorksheetData[0,$Col++] = 'Maximum Value'
			$WorksheetData[0,$Col++] = 'Cycle'
			$WorksheetData[0,$Col++] = 'Cache Option'
			$WorksheetData[0,$Col++] = 'Cache Size'
			$WorksheetData[0,$Col++] = 'Is Exhausted'
			$WorksheetData[0,$Col++] = 'Is Schema Owned'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Programmability.Sequences | Where-Object { $_.Properties.General.ID } | Sort-Object -Property @{Expression={$_.Properties.General.Schema}}, @{Expression={$_.Properties.General.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Schema
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Owner
						$WorksheetData[$Row,$Col++] = $_.Properties.General.Name
						$WorksheetData[$Row,$Col++] = $_.Properties.General.CreateDate
						$WorksheetData[$Row,$Col++] = $_.Properties.General.DateLastModified
						$WorksheetData[$Row,$Col++] = $_.Properties.General.DataType.Name
						$WorksheetData[$Row,$Col++] = $_.Properties.General.DataType.NumericPrecision
						$WorksheetData[$Row,$Col++] = $_.Properties.General.StartValue
						$WorksheetData[$Row,$Col++] = $_.Properties.General.IncrementValue
						$WorksheetData[$Row,$Col++] = $_.Properties.General.MaxValue
						$WorksheetData[$Row,$Col++] = $_.Properties.General.MinValue
						$WorksheetData[$Row,$Col++] = $_.Properties.General.IsCycleEnabled
						$WorksheetData[$Row,$Col++] = $_.Properties.General.SequenceCacheType
						$WorksheetData[$Row,$Col++] = $_.Properties.General.CacheSize
						$WorksheetData[$Row,$Col++] = $_.Properties.General.IsExhausted
						$WorksheetData[$Row,$Col++] = $_.Properties.General.IsSchemaOwned
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'F2'
					ColumnFormat = @(
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 9; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 10; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 11; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 12; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 16; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 64: Service Broker - Message Types
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Service Broker - Message Types"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Broker - Message Types'
			#$Worksheet.Tab.Color = $ManagementTabColor
			$Worksheet.Tab.ThemeColor = $ManagementTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.ServiceBroker.MessageTypes | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 8
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Owner Name'
			$WorksheetData[0,$Col++] = 'Message Type Name'
			$WorksheetData[0,$Col++] = 'System Object'
			$WorksheetData[0,$Col++] = 'Message Type Validation'
			$WorksheetData[0,$Col++] = 'XML Schema Collection'
			$WorksheetData[0,$Col++] = 'XML Schema Collection Schema'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.ServiceBroker.MessageTypes | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Owner}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Owner
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.IsSystemObject
						$WorksheetData[$Row,$Col++] = $_.MessageTypeValidation
						$WorksheetData[$Row,$Col++] = $_.ValidationXmlSchemaCollection
						$WorksheetData[$Row,$Col++] = $_.ValidationXmlSchemaCollectionSchema
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 65: Service Broker - Contracts
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Service Broker - Contracts"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Broker - Contracts'
			#$Worksheet.Tab.Color = $ManagementTabColor
			$Worksheet.Tab.ThemeColor = $ManagementTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.ServiceBroker.ServiceContracts | ForEach-Object { $_.MessageTypeMappings | Where-Object { $_.Name } } } } | Measure-Object).Count + 1
			$ColumnCount = 7
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Owner Name'
			$WorksheetData[0,$Col++] = 'Contract Name'
			$WorksheetData[0,$Col++] = 'System Object'
			$WorksheetData[0,$Col++] = 'Message Type'
			$WorksheetData[0,$Col++] = 'Sent By'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.ServiceBroker.ServiceContracts | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Owner}}, @{Expression={$_.Name}} | ForEach-Object {
						$SchemaName = $_.Owner
						$ObjectName = $_.Name
						$IsSystemObject = $_.IsSystemObject

						$_.MessageTypeMappings | Where-Object { $_.Name } | Sort-Object -Property MessageSource | ForEach-Object {
							$Col = 0
							$WorksheetData[$Row,$Col++] = $ServerName
							$WorksheetData[$Row,$Col++] = $DatabaseName
							$WorksheetData[$Row,$Col++] = $SchemaName
							$WorksheetData[$Row,$Col++] = $ObjectName
							$WorksheetData[$Row,$Col++] = $IsSystemObject
							$WorksheetData[$Row,$Col++] = $_.Name
							$WorksheetData[$Row,$Col++] = $_.MessageSource
							$Row++
						}
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 66: Service Broker - Queues
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Service Broker - Queues"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Broker - Queues'
			#$Worksheet.Tab.Color = $ManagementTabColor
			$Worksheet.Tab.ThemeColor = $ManagementTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.ServiceBroker.Queues | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 19
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Owner Name'
			$WorksheetData[0,$Col++] = 'Queue Name'
			$WorksheetData[0,$Col++] = 'System Object'
			$WorksheetData[0,$Col++] = 'Activation Execution Context'
			$WorksheetData[0,$Col++] = 'Create Date'
			$WorksheetData[0,$Col++] = 'Last Modified Date'
			$WorksheetData[0,$Col++] = 'Execution Context Principal'
			$WorksheetData[0,$Col++] = 'File Group'
			$WorksheetData[0,$Col++] = 'Is Activation Enabled'
			$WorksheetData[0,$Col++] = 'Is Enqueue Enabled'
			$WorksheetData[0,$Col++] = 'Is Poison Message Handing Enabled'
			$WorksheetData[0,$Col++] = 'Is Retention Enabled'
			$WorksheetData[0,$Col++] = 'Maximum Readers'
			$WorksheetData[0,$Col++] = 'Procedure Database'
			$WorksheetData[0,$Col++] = 'Procedure Name'
			$WorksheetData[0,$Col++] = 'Procedure Schema'
			$WorksheetData[0,$Col++] = 'Message Count'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.ServiceBroker.Queues | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Owner}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Owner
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.IsSystemObject
						$WorksheetData[$Row,$Col++] = $_.ActivationExecutionContext
						$WorksheetData[$Row,$Col++] = $_.CreateDate
						$WorksheetData[$Row,$Col++] = $_.DateLastModified
						$WorksheetData[$Row,$Col++] = $_.ExecutionContextPrincipal
						$WorksheetData[$Row,$Col++] = $_.FileGroup
						$WorksheetData[$Row,$Col++] = $_.IsActivationEnabled
						$WorksheetData[$Row,$Col++] = $_.IsEnqueueEnabled
						$WorksheetData[$Row,$Col++] = $_.IsPoisonMessageHandlingEnabled
						$WorksheetData[$Row,$Col++] = $_.IsRetentionEnabled
						$WorksheetData[$Row,$Col++] = $_.MaxReaders
						$WorksheetData[$Row,$Col++] = $_.ProcedureDatabase
						$WorksheetData[$Row,$Col++] = $_.ProcedureName
						$WorksheetData[$Row,$Col++] = $_.ProcedureSchema
						$WorksheetData[$Row,$Col++] = $_.RowCount
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 15; NumberFormat = $XlNumFmtNumberGeneral},
						@{ColumnNumber = 19; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 67: Service Broker - Services
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Service Broker - Services"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Broker - Services'
			#$Worksheet.Tab.Color = $ManagementTabColor
			$Worksheet.Tab.ThemeColor = $ManagementTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.ServiceBroker.Services | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 8
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Owner Name'
			$WorksheetData[0,$Col++] = 'Service Name'
			$WorksheetData[0,$Col++] = 'System Object'
			$WorksheetData[0,$Col++] = 'Queue Name'
			$WorksheetData[0,$Col++] = 'Queue Schema'
			$WorksheetData[0,$Col++] = 'Mapped Contracts'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.ServiceBroker.Services | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Owner}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Owner
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.IsSystemObject
						$WorksheetData[$Row,$Col++] = $_.QueueName
						$WorksheetData[$Row,$Col++] = $_.QueueSchema
						$WorksheetData[$Row,$Col++] = $($_.ServiceContractMappings | ForEach-Object { $_.Name }) -join $Delimiter
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 68: Service Broker - Routes
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Service Broker - Routes"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Broker - Routes'
			#$Worksheet.Tab.Color = $ManagementTabColor
			$Worksheet.Tab.ThemeColor = $ManagementTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.ServiceBroker.Routes | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 9
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Owner Name'
			$WorksheetData[0,$Col++] = 'Route Name'
			$WorksheetData[0,$Col++] = 'Address'
			$WorksheetData[0,$Col++] = 'Broker Instance'
			$WorksheetData[0,$Col++] = 'Expiration Date'
			$WorksheetData[0,$Col++] = 'Mirror Address'
			$WorksheetData[0,$Col++] = 'Remote Service'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.ServiceBroker.Routes | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Owner}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Owner
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.Address
						$WorksheetData[$Row,$Col++] = $_.BrokerInstance
						$WorksheetData[$Row,$Col++] = $_.ExpirationDate
						$WorksheetData[$Row,$Col++] = $_.MirrorAddress
						$WorksheetData[$Row,$Col++] = $_.RemoteService
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtDate}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 69: Service Broker - Remote Service Bindings
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Service Broker - Remote Service Bindings"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Broker - Remote Service Binding'
			#$Worksheet.Tab.Color = $ManagementTabColor
			$Worksheet.Tab.ThemeColor = $ManagementTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.ServiceBroker.RemoteServiceBindings | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 7
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Owner Name'
			$WorksheetData[0,$Col++] = 'Remove Service Binding Name'
			$WorksheetData[0,$Col++] = 'Is Anonymous'
			$WorksheetData[0,$Col++] = 'Certificate User'
			$WorksheetData[0,$Col++] = 'Remote Service'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.ServiceBroker.RemoteServiceBindings | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Owner}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Owner
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.IsAnonymous
						$WorksheetData[$Row,$Col++] = $_.CertificateUser
						$WorksheetData[$Row,$Col++] = $_.RemoteService
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 70: Service Broker - Priorities
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Service Broker - Priorities"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Broker - Priorities'
			#$Worksheet.Tab.Color = $ManagementTabColor
			$Worksheet.Tab.ThemeColor = $ManagementTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.ServiceBroker.Priorities | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 8
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Owner Name'
			$WorksheetData[0,$Col++] = 'Priority Name'
			$WorksheetData[0,$Col++] = 'Priority Level'
			$WorksheetData[0,$Col++] = 'Contract Name'
			$WorksheetData[0,$Col++] = 'Local Service Name'
			$WorksheetData[0,$Col++] = 'Remote Service Name'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.ServiceBroker.Priorities | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Owner}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Owner
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.PriorityLevel
						$WorksheetData[$Row,$Col++] = $_.ContractName
						$WorksheetData[$Row,$Col++] = $_.LocalServiceName
						$WorksheetData[$Row,$Col++] = $_.RemoteServiceName
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 3; NumberFormat = $XlNumFmtNumberGeneral}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 71: Storage - Full Text Catalogs
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Storage - Full Text Catalogs"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Storage - Full Text Catalogs'
			#$Worksheet.Tab.Color = $AgentTabColor
			$Worksheet.Tab.ThemeColor = $AgentTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Storage.FullTextCatalogs | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 16
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Owner Name'
			$WorksheetData[0,$Col++] = 'Catalog Name'
			$WorksheetData[0,$Col++] = 'Has Full Text Indexed Tables'
			$WorksheetData[0,$Col++] = 'Accent Sensitive'
			$WorksheetData[0,$Col++] = 'Full Text Index Size (MB)'
			$WorksheetData[0,$Col++] = 'Error Log Size (MB)'
			$WorksheetData[0,$Col++] = 'Default Catalog'
			$WorksheetData[0,$Col++] = 'Filegroup'
			$WorksheetData[0,$Col++] = 'Item Count'
			$WorksheetData[0,$Col++] = 'Last Population Date'
			$WorksheetData[0,$Col++] = 'Population Age (Sec)'
			$WorksheetData[0,$Col++] = 'Population Status'
			$WorksheetData[0,$Col++] = 'Root Path'
			$WorksheetData[0,$Col++] = 'Unique Key Count'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Storage.FullTextCatalogs | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Owner}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Owner
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.HasFullTextIndexedTables
						$WorksheetData[$Row,$Col++] = $_.IsAccentSensitive
						$WorksheetData[$Row,$Col++] = $_.FullTextIndexSizeMB
						$WorksheetData[$Row,$Col++] = $_.ErrorLogSizeBytes / 1MB
						$WorksheetData[$Row,$Col++] = $_.IsDefault
						$WorksheetData[$Row,$Col++] = $_.FileGroup
						$WorksheetData[$Row,$Col++] = $_.ItemCount
						$WorksheetData[$Row,$Col++] = $_.PopulationCompletionDate
						$WorksheetData[$Row,$Col++] = $_.PopulationCompletionAgeSeconds
						$WorksheetData[$Row,$Col++] = $_.PopulationStatus
						$WorksheetData[$Row,$Col++] = $_.RootPath
						$WorksheetData[$Row,$Col++] = $_.UniqueKeyCount
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 7; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberS2},
						@{ColumnNumber = 11; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 12; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 16; NumberFormat = $XlNumFmtNumberS0}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 72: Storage - Full Text Stoplists
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Storage - Full Text Stoplists"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Storage - Full Text Stoplists'
			#$Worksheet.Tab.Color = $AgentTabColor
			$Worksheet.Tab.ThemeColor = $AgentTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Storage.FullTextStopLists | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 6
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Owner Name'
			$WorksheetData[0,$Col++] = 'Stoplist Name'
			$WorksheetData[0,$Col++] = 'Word Count'
			$WorksheetData[0,$Col++] = 'Stop Words (First 100)'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Storage.FullTextStopLists | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Owner}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Owner
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $($_.StopWords | Measure-Object).Count
						$WorksheetData[$Row,$Col++] = $(
							$_.StopWords | 
							Sort-Object -Property StopWord, Language | 
							Select-Object -First 100 | 
							ForEach-Object { "'$($_.StopWord)' ($($_.Language))" }
						) -join $Delimiter
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @(
						@{ColumnNumber = 5; NumberFormat = $XlNumFmtNumberS0},
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtText}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 73: Storage - Partition Schemes
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Storage - Partition Schemes"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Storage - Partition Schemes'
			#$Worksheet.Tab.Color = $AgentTabColor
			$Worksheet.Tab.ThemeColor = $AgentTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Storage.PartitionSchemes | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 7
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Owner Name'
			$WorksheetData[0,$Col++] = 'Partition Scheme Name'
			$WorksheetData[0,$Col++] = 'Partition Function'
			$WorksheetData[0,$Col++] = 'Next Used File Group'
			$WorksheetData[0,$Col++] = 'File Groups'

			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Storage.PartitionSchemes | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Owner}}, @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Owner
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.PartitionFunction
						$WorksheetData[$Row,$Col++] = $_.NextUsedFileGroup
						$WorksheetData[$Row,$Col++] = $($_.FileGroups | Sort-Object) -join $Delimiter
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'E2'
					ColumnFormat = @()
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion


			# Worksheet 74: Storage - Partition Functions
			$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Storage - Partition Functions"
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			#region
			$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
			$Worksheet.Name = 'Storage - Partition Functions'
			#$Worksheet.Tab.Color = $AgentTabColor
			$Worksheet.Tab.ThemeColor = $AgentTabColor

			$RowCount = ($SqlServerInventory.DatabaseServer | ForEach-Object { $_.Server.Databases | ForEach-Object { $_.Storage.PartitionFunctions | Where-Object { $_.ID } } } | Measure-Object).Count + 1
			$ColumnCount = 8
			$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

			$Col = 0
			$WorksheetData[0,$Col++] = 'Server Name'
			$WorksheetData[0,$Col++] = 'Database Name'
			$WorksheetData[0,$Col++] = 'Partition Function Name'
			$WorksheetData[0,$Col++] = 'Create Date'
			$WorksheetData[0,$Col++] = 'Number Of Partitions'
			$WorksheetData[0,$Col++] = 'Parameters'
			$WorksheetData[0,$Col++] = 'Range Type'
			$WorksheetData[0,$Col++] = 'Range Values'


			$Row = 1
			$SqlServerInventory.DatabaseServer | Sort-Object -Property @{Expression={$_.Server.Configuration.General.Name}} | ForEach-Object {

				$ServerName = $_.Server.Configuration.General.Name

				$_.Server.Databases | Sort-Object -Property Name | ForEach-Object {
					$DatabaseName = $_.Name

					$_.Storage.PartitionFunctions | Where-Object { $_.ID } | Sort-Object -Property @{Expression={$_.Name}} | ForEach-Object {
						$Col = 0
						$WorksheetData[$Row,$Col++] = $ServerName
						$WorksheetData[$Row,$Col++] = $DatabaseName
						$WorksheetData[$Row,$Col++] = $_.Name
						$WorksheetData[$Row,$Col++] = $_.CreateDate
						$WorksheetData[$Row,$Col++] = $_.NumberOfPartitions
						$WorksheetData[$Row,$Col++] = $(
							$_.PartitionFunctionParameters | ForEach-Object {
								'{0} (Length: {1:D}; Precision: {2:D}; Scale: {3:D})' -f $_.Name, $_.Length, $_.NumericPrecision, $_.NumericScale
							}
						) -join $Delimiter
						$WorksheetData[$Row,$Col++] = $_.RangeType
						$WorksheetData[$Row,$Col++] = $($_.RangeValues | Sort-Object) -join $Delimiter
						$Row++
					}
				}
			}
			$Range = $Worksheet.Range($Worksheet.Cells.Item(1,1), $Worksheet.Cells.Item($RowCount,$ColumnCount))
			$Range.Value2 = $WorksheetData
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
			#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

			$WorksheetFormat.Add($WorksheetNumber, @{
					BoldFirstRow = $true
					BoldFirstColumn = $false
					AutoFilter = $true
					FreezeAtCell = 'D2'
					ColumnFormat = @(
						@{ColumnNumber = 4; NumberFormat = $XlNumFmtDate},
						@{ColumnNumber = 6; NumberFormat = $XlNumFmtText}
					)
					RowFormat = @()
				})

			$WorksheetNumber++
			#endregion




			# Apply formatting to every worksheet
			# Work backwards so that the first sheet is active when the workbook is saved
			$ProgressStatus = 'Applying formatting to all worksheets'
			Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
			Write-Progress -Activity $ProgressActivity -PercentComplete 0 -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId
			for ($WorksheetNumber = $WorksheetCount; $WorksheetNumber -ge 1; $WorksheetNumber--) {

				$ProgressStatus = "Applying formatting to Worksheet #$($WorksheetNumber)"
				Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Verbose
				Write-Progress -Activity $ProgressActivity -PercentComplete (((($WorksheetCount * 2) - $WorksheetNumber + 1) / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId

				$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)

				# Switch to the worksheet
				$Worksheet.Activate() | Out-Null

				# Bold the header row
				#$Duration = (Measure-Command {
				$Worksheet.Rows.Item(1).Font.Bold = $WorksheetFormat[$WorksheetNumber].BoldFirstRow
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Bold Header Row Duration (ms): $Duration" -MessageLevel Verbose

				# Bold the 1st column
				#$Duration = (Measure-Command {
				$Worksheet.Columns.Item(1).Font.Bold = $WorksheetFormat[$WorksheetNumber].BoldFirstColumn
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Bold 1st Column Duration (ms): $Duration" -MessageLevel Verbose

				# Freeze View
				#$Duration = (Measure-Command {
				$Worksheet.Range($WorksheetFormat[$WorksheetNumber].FreezeAtCell).Select() | Out-Null
				$Worksheet.Application.ActiveWindow.FreezePanes = $true 
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Freeze View Duration (ms): $Duration" -MessageLevel Verbose


				# Apply Column formatting
				#$Duration = (Measure-Command {
				$WorksheetFormat[$WorksheetNumber].ColumnFormat | ForEach-Object {
					$Worksheet.Columns.Item($_.ColumnNumber).NumberFormat = $_.NumberFormat
				}
				#}).TotalMilliseconds
				Write-SqlServerInventoryLog -Message "Apply Column formatting Duration (ms): $Duration" -MessageLevel Verbose

				# Apply Row formatting
				#$Duration = (Measure-Command {
				$WorksheetFormat[$WorksheetNumber].RowFormat | ForEach-Object {
					$Worksheet.Rows.Item($_.RowNumber).NumberFormat = $_.NumberFormat
				}
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Apply Row formatting Duration (ms): $Duration" -MessageLevel Verbose

				# Update worksheet values so row and column formatting apply
				#$Duration = (Measure-Command {
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
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Apply Row and Column formatting - Update Values (ms): $Duration" -MessageLevel Verbose


				# Apply table formatting
				#$Duration = (Measure-Command {
				$ListObject = $Worksheet.ListObjects.Add($XlListObjectSourceType::xlSrcRange, $Worksheet.UsedRange, $null, $XlYesNoGuess::xlYes, $null) 
				$ListObject.Name = "Table $WorksheetNumber"
				$ListObject.TableStyle = $TableStyle
				$ListObject.ShowTableStyleFirstColumn = $WorksheetFormat[$WorksheetNumber].BoldFirstColumn # Put a background color behind the 1st column
				$ListObject.ShowAutoFilter = $WorksheetFormat[$WorksheetNumber].AutoFilter
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Apply table formatting Duration (ms): $Duration" -MessageLevel Verbose

				# Zoom back to 80%
				#$Duration = (Measure-Command {
				$Worksheet.Application.ActiveWindow.Zoom = 80
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Zoom to 80% Duration (ms): $Duration" -MessageLevel Verbose

				# Adjust the column widths to 250 before autofitting contents
				# This allows longer lines of text to remain on one line
				#$Duration = (Measure-Command {
				$Worksheet.UsedRange.EntireColumn.ColumnWidth = 250
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Change column width Duration (ms): $Duration" -MessageLevel Verbose

				# Autofit column and row contents
				#$Duration = (Measure-Command {
				$Worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null
				$Worksheet.UsedRange.EntireRow.AutoFit() | Out-Null
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Autofit contents Duration (ms): $Duration" -MessageLevel Verbose

				# Left align contents
				#$Duration = (Measure-Command {
				$Worksheet.UsedRange.EntireColumn.HorizontalAlignment = $XlHAlign::xlHAlignLeft
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Left align contents Duration (ms): $Duration" -MessageLevel Verbose

				# Vertical align contents
				#$Duration = (Measure-Command {
				$Worksheet.UsedRange.EntireColumn.VerticalAlignment = $XlVAlign::xlVAlignTop
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Vertical align contents Duration (ms): $Duration" -MessageLevel Verbose

				# Put the selection back to the upper left cell
				#$Duration = (Measure-Command {
				$Worksheet.Range('A1').Select() | Out-Null
				#}).TotalMilliseconds
				#Write-SqlServerInventoryLog -Message "Reset selection Duration (ms): $Duration" -MessageLevel Verbose
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
		Write-SqlServerInventoryLog -Message $ProgressStatus -MessageLevel Information
		Write-SqlServerInventoryLog -Message "End Function: $($MyInvocation.InvocationName)" -MessageLevel Information


		# Cleanup
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

		Remove-Variable -Name TabCharLength
		Remove-Variable -Name IndentString_1
		Remove-Variable -Name IndentString__2
		Remove-Variable -Name IndentString___3

		Remove-Variable -Name ComputerName
		Remove-Variable -Name ServerName
		Remove-Variable -Name ProductName
		Remove-Variable -Name DatabaseName
		Remove-Variable -Name FileGroupName

		Remove-Variable -Name ColorThemePathPattern
		Remove-Variable -Name ColorThemePath

		Remove-Variable -Name WorksheetFormat

		Remove-Variable -Name XlSortOrder
		Remove-Variable -Name XlYesNoGuess
		Remove-Variable -Name XlHAlign
		Remove-Variable -Name XlVAlign
		Remove-Variable -Name XlListObjectSourceType
		Remove-Variable -Name XlThemeColor

		Remove-Variable -Name OverviewTabColor
		Remove-Variable -Name ServicesTabColor
		Remove-Variable -Name ServerTabColor
		Remove-Variable -Name DatabaseTabColor
		Remove-Variable -Name AgentTabColor

		Remove-Variable -Name TableStyle

		Remove-Variable -Name ProgressId
		Remove-Variable -Name ProgressActivity
		Remove-Variable -Name ProgressStatus

		# Release all lingering COM objects
		Remove-ComObject

	}
}

<#
Function Names Reserved For Future Consideration:

Export-SqlServerInventoryWindowsInventoryToExcel
Export-SqlServerInventoryWindowsMachineToExcel
Export-SqlServerInventoryWindowsConfigToExcel
Export-SqlServerInventoryWindowsOperatingSystemToExcel
Export-SqlServerInventoryWindowsOsToExcel
Export-SqlServerInventoryWindowsOStoExcel
Export-SqlServerInventoryWindowsToExcel

Export-SqlServerInventoryDbEngineConfigToExcel
Export-SqlServerInventoryDbEngineDbObjectsToExcel
Export-SqlServerInventoryDatabaseEngineConfigToExcel
Export-SqlServerInventoryDatabaseEngineDbObjectsToExcel

Export-SqlInventoryDatabaseEngineConfigToExcel
Export-SqlInventoryDatabaseEngineDbObjectsToExcel
Export-SqlInventoryDbEngineConfigToExcel
Export-SqlInventoryDbEngineDbObjectsToExcel
Export-SqlInvDatabaseEngineConfigToExcel
Export-SqlInvDbeEngineConfigToExcel
Export-SqlInvDbeEngineDbObjectsToExcel



Get-SqlServerServiceData
Get-DbEngineOverviewData
Get-DbEngineServerConfigGeneralData
Get-DbEngineServerConfigMemoryData
Get-DbEngineServerConfigProcessorsData
Get-DbEngineServerConfigSecurityData
Get-DbEngineServerConfigConnectionsData
Get-DbEngineServerConfigDatabaseSettingsData
Get-DbEngineServerConfigAdvancedData
Get-DbEngineServerConfigClusteringData
Get-DbEngineServerConfigAlwaysOnData
Get-DbEngineServerSecurityLoginsData
Get-DbEngineServerSecurityRolesData
Get-DbEngineServerSecurityCredentialsData
Get-DbEngineServerSecurityAuditsData
Get-DbEngineServerSecurityAuditSpecificationsData
Get-DbEngineServerObjectsEndpointsData
Get-DbEngineServerObjectsEndpointsData
Get-DbEngineServerObjectsLinkedServerConfigurationData
Get-DbEngineServerObjectsLinkedServerLoginsData
Get-DbEngineServerObjectsTraceFlagsData
Get-DbEngineServerObjectsServerTriggersData
Get-DbEngineServerManagementStartupProceduresData
Get-DbEngineServerManagementResourceGovernorData
Get-DbEngineServerManagementDatabaseMailAccountsData
Get-DbEngineServerManagementDatabaseMailProfilesData
Get-DbEngineServerManagementDatabaseMailSecurityData
Get-DbEngineServerManagementDatabaseMailConfigurationData
Get-DbEngineDatabaseOverviewData
Get-DbEngineDatabaseConfigGeneralData
Get-DbEngineDatabaseConfigFilesData
Get-DbEngineDatabaseConfigFilegroupsData
Get-DbEngineDatabaseConfigOptionsData
Get-DbEngineDatabaseConfigAlwaysOnData
Get-DbEngineDatabaseConfigChangeTrackingData
Get-DbEngineDatabaseConfigMirroringData
Get-DbEngineDatabaseSecuritySchemasData
Get-DbEngineDatabaseSecurityUsersData
Get-DbEngineDatabaseSecurityDatabaseRolesData
Get-DbEngineDatabaseSecurityApplicationRolesData
Get-DbEngineDatabaseSecurityCertificatesData
Get-DbEngineDatabaseSecurityAsymmetricKeysData
Get-DbEngineDatabaseSecuritySymmetricKeysData
Get-AgentConfigData
Get-AgentJobsData
Get-AgentJobSchedulesData
Get-AgentJobStepsData
Get-AgentJobAlertsData
Get-AgentJobNotificationsData
Get-AgentAlertsData
Get-AgentOperatorsData
#>