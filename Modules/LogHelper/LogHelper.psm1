##########################
# PRIVATE SCRIPT VARIABLES
##########################
New-Variable -Name LogFile -Value $null -Scope Script -Visibility Private
New-Variable -Name LoggingPreference -Value 'none' -Scope Script -Visibility Private
New-Variable -Name WriteFailureRetry -Value 10 -Scope Script -Visibility Private
New-Variable -Name LogQueue -Value $null -Scope Script -Visibility Private
New-Variable -Name LogQueueIsLocked -Value $false -Scope Script -Visibility Private

$script:LogQueue = [System.Collections.Queue]::Synchronized((New-Object -TypeName System.Collections.Queue))
$script:LogFile = [System.IO.Path]::ChangeExtension([IO.Path]::GetTempFileName(),'txt')



###################
# PUBLIC FUNCTIONS
###################
function Set-LogFile {
	[CmdletBinding()]
	param(
		[Parameter(Position=0, Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$Path
	)
	try {
		if ((Split-Path -Path $Path -Parent | Test-Path) -eq $true) {
			$script:LogFile = $Path
		} else {
			throw 'Invalid logfile path specified'
		}
	}
	catch {
		throw
	}
}

function Set-LoggingPreference {
	[CmdletBinding()]
	param(
		[Parameter(Position=0, Mandatory=$true)]
		#[ValidateSet('none','minimal','normal','verbose')]
		[ValidateSet('none','standard','verbose','debug')]
		[System.String]
		$Preference
	)
	try {
		$script:LoggingPreference = $Preference
	}
	catch {
		throw
	}
}

function Set-LogQueue {
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

function Get-LogFile {
	Write-Output $script:LogFile
}

function Get-LoggingPreference {
	Write-Output $script:LoggingPreference
}

function Write-Log {
	<#
		.SYNOPSIS
			A brief description of the function.

		.DESCRIPTION
			A detailed description of the function.

		.PARAMETER  ParameterA
			The description of the ParameterA parameter.

		.PARAMETER  ParameterB
			The description of the ParameterB parameter.

		.EXAMPLE
			PS C:\> Get-Something -ParameterA 'One value' -ParameterB 32

		.EXAMPLE
			PS C:\> Get-Something 'One value' 32

		.INPUTS
			System.String,System.Int32

		.OUTPUTS
			System.String

		.NOTES
			Additional information about the function go here.

		.LINK
			about_functions_advanced

		.LINK
			about_comment_based_help

	#>
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
		$WriteToLog = $false

		# Determine if we're going to write to the log based on the logging preferences and message level
		$WriteToLog = switch ($script:LoggingPreference) {
			'debug' {
				switch ($MessageLevel) {
					'error' { $true }
					'warning' { $true }
					'information' { $true }
					'verbose' { $true }
					'debug' { $true }
					default { $false }
				}
			}
			'verbose' {
				switch ($MessageLevel) {
					'error' { $true }
					'warning' { $true }
					'information' { $true }
					'verbose' { $true }
					default { $false }
				}
			}
			'standard' {
				switch ($MessageLevel) {
					'error' { $true }
					'warning' { $true }
					'information' { $true }
					default { $false }
				}
			}
			default {
				$false
			}
		}

		if ($WriteToLog -eq $true) {
			$Symbol = switch ($MessageLevel) {
				'information' { '?' }
				'verbose' { '$' }
				'debug' { '*' }
				'error' { '!' }
				'warning' { '+' }
			}

			$script:LogQueue.Enqueue(("{0} {1} {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.ffff'), $Symbol, $Message))

		} else {

			# If we're not writing to the log then we need some way to alert on warnings and errors
			# so write to the warning and error streams
			if ($MessageLevel -ieq 'warning') {
				Write-Warning $Message -WarningAction Continue
			}
			elseif ($MessageLevel -ieq 'error') {
				Write-Error $Message -ErrorAction Continue
			}
			elseif ($MessageLevel -ieq 'verbose') {
 				Write-Verbose $Message
 			} 
			elseif ($MessageLevel -ieq 'debug') {
 				Write-Debug $Message
 			}
			else {
				# For lack of a better "other" stream just write it to the debug stream
				Write-Debug $Message
			}

		}

		# Putting this outside of the $WriteToLog check b\c items may be in the queue
		# although $script:LoggingPreference may have changed (which would change $WriteToLog)
		# Another thought: Does this part belong in its own function? Or run as a background job?
		try {
			# Lock the queue...this prevents anything else from accessing it
			[System.Threading.Monitor]::Enter($script:LogQueue.SyncRoot)
			$script:LogQueueIsLocked = $true

			# Try and write each message in the queue
			# If we fail the item remains in the queue and will be handled on the next log entry
			while ($script:LogQueue.Count -gt 0) {
				$script:LogQueue.Peek() | Out-File -Encoding Default -FilePath $script:LogFile -Append
				$script:LogQueue.Dequeue() | Out-Null
			}
		}
		catch {
		}
		finally {
			if ($script:LogQueueIsLocked -eq $true) {
				[System.Threading.Monitor]::Exit($script:LogQueue.SyncRoot)
				$script:LogQueueIsLocked = $false
			}
		}

	}
	catch {
		Write-Host "$(Get-Date)`tError writing to log file: $Message"
		#throw
	}
}


#############################
# RUN WHEN MODULE IS UNLOADED
#############################
$MyInvocation.MyCommand.ScriptBlock.Module.OnRemove = {
	try {
		# Lock the queue...this prevents anything else from accessing it
		[System.Threading.Monitor]::Enter($script:LogQueue.SyncRoot)
		$script:LogQueueIsLocked = $true

		# Try and write each message in the queue
		# If we fail the item remains in the queue and will be handled on the next log entry
		while ($script:LogQueue.Count -gt 0) {
			$script:LogQueue.Peek() | Out-File -Encoding Default -FilePath $script:LogFile -Append
			$script:LogQueue.Dequeue() | Out-Null
		}
	}
	catch {
	}
	finally {
		if ($script:LogQueueIsLocked -eq $true) {
			[System.Threading.Monitor]::Exit($script:LogQueue.SyncRoot)
			$script:LogQueueIsLocked = $false
		}
	}
}
