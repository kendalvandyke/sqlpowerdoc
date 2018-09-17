<#
	.SYNOPSIS
		Writes Excel files containing the Database Engine and Windows Operating System information from a SQL Server Inventory file created by Get-SqlServerInventoryToClixml.ps1.

	.DESCRIPTION
		This script loads a SQL Server Inventory file created by Get-SqlServerInventoryToClixml.ps1 and calls the Export-SqlDatabaseEngineInventoryToExcel and Export-SqlWindowsInventoryToExcel functions in the SqlServerInventory module to write Excel files containing the Database Engine and Windows Operating System information from the inventory.

		Microsoft Excel 2007 or higher must be installed in order to write the Excel file.
		
	.PARAMETER  FromPath
		The literal path to the XML file created by Get-SqlServerInventoryToClixml.ps1.
		
	.PARAMETER  ToDirectoryPath
		Specifies the literal path to the directory where the Excel workbooks will be written. This path must exist prior to executing this script.
		
		If not specified then ToDirectoryPath defaults to the same directory specified by the FromPath paramter.
		
		Assuming the XML file specified in FromPath is named "SQL Server Inventory.xml" then:
			The Database Engine configuration details will be written to "SQL Server Inventory - Database Engine Config.xlsx"
			The Database Engine Database Objects details will be written to "SQL Server Inventory - Database Engine Db Objects.xlsx"
			The Database Engine configuration assessment will be written to "SQL Server Inventory - Database Engine Assessment.xlsx"
			The Windows Operating System details will be written to "SQL Server Inventory - Windows.xlsx"
		
	.PARAMETER  ColorTheme
		An Office Theme Color to apply to each worksheet. If not specified or if an unknown theme color is provided the default "Office" theme colors will be used.
		
		Office 2013 theme colors include: Aspect, Blue Green, Blue II, Blue Warm, Blue, Grayscale, Green Yellow, Green, Marquee, Median, Office, Office 2007 - 2010, Orange Red, Orange, Paper, Red Orange, Red Violet, Red, Slipstream, Violet II, Violet, Yellow Orange, Yellow
		
		Office 2010 theme colors include: Adjacency, Angles, Apex, Apothecary, Aspect, Austin, Black Tie, Civic, Clarity, Composite, Concourse, Couture, Elemental, Equity, Essential, Executive, Flow, Foundry, Grayscale, Grid, Hardcover, Horizon, Median, Metro, Module, Newsprint, Office, Opulent, Oriel, Origin, Paper, Perspective, Pushpin, Slipstream, Solstice, Technic, Thatch, Trek, Urban, Verve, Waveform

		Office 2007 theme colors include: Apex, Aspect, Civic, Concourse, Equity, Flow, Foundry, Grayscale, Median, Metro, Module, Office, Opulent, Oriel, Origin, Paper, Solstice, Technic, Trek, Urban, Verve
		
	.PARAMETER  ColorScheme
		The color theme to apply to each worksheet. Valid values are "Light", "Medium", and "Dark". 
		
		If not specified then "Medium" is used as the default value .

	.PARAMETER  LoggingPreference
		Specifies the logging verbosity to use when writing log entries.
		
		Valid values include: None, Standard, Verbose, and Debug.
		
		The default value is "None"
		
	.PARAMETER  LogPath
		A literal path to a log file to write details about what this script is doing. The filename does not need to exist prior to executing this script but the specified directory does.
		
		If a LoggingPreference other than None is specified and this parameter is not specified then the file is named "SQL Server Inventory - [Year][Month][Day][Hour][Minute].log" and is written to the same directory specified by the ToDirectoryPath paramter.
		
	.EXAMPLE
		.\Convert-SqlServerInventoryClixmlToExcel.ps1 -FromPath "C:\Inventory\SQL Server Inventory.xml" 
		
		Description
		-----------
		Writes Excel files for the Database Engine and Windows Operating System information contained in "C:\Inventory\SQL Server Inventory.xml" to "C:\Inventory\SQL Server - Database Engine.xlsx" and "C:\Inventory\SQL Server - Windows.xlsx", respectively.
		
		The Office color theme and Medium color scheme will be used by default.
		
	.EXAMPLE
		.\Convert-SqlServerInventoryClixmlToExcel.ps1 -FromPath "C:\Inventory\SQL Server Inventory.xml"  -ColorTheme Blue -ColorScheme Dark
		
		Description
		-----------
		Writes Excel files for the Database Engine and Windows Operating System information contained in "C:\Inventory\SQL Server Inventory.xml" to "C:\Inventory\SQL Server - Database Engine.xlsx" and "C:\Inventory\SQL Server - Windows.xlsx", respectively.
		
		The Blue color theme and Dark color scheme will be used.

	
	.NOTES
		Blue and Green are nice looking Color Themes for Office 2013

		Waveform is a nice looking Color Theme for Office 2010

	.LINK
		Get-SqlServerInventory

#>
[cmdletBinding(SupportsShouldProcess=$false)]
param(
	[Parameter(Mandatory=$true)] 
	[alias('From')]
	[ValidateNotNullOrEmpty()]
	[string]
	$FromPath
	,
	[Parameter(Mandatory=$false)] 
	[alias('To')]
	[ValidateNotNullOrEmpty()]
	[string]
	$ToDirectoryPath = ([System.IO.Path]::GetDirectoryName($FromPath))
	,
	[Parameter(Mandatory=$false)] 
	[alias('LogLevel')]
	[ValidateSet('None','Standard','Verbose','Debug')]
	[string]
	$LoggingPreference = 'none'
	,
	[Parameter(Mandatory=$false)] 
	[alias('Log')]
	[ValidateNotNullOrEmpty()]
	[string]
	$LogPath = (Join-Path -Path $ToDirectoryPath -ChildPath ("SQL Server Inventory - " + (Get-Date -Format "yyyy-MM-dd-HH-mm") + ".log"))
	,
	[Parameter(Mandatory=$false)] 
	[alias('Theme')]
	[string]
	$ColorTheme = 'Office'
	,
	[Parameter(Mandatory=$false)] 
	[ValidateSet('Dark','Light','Medium')]
	[string]
	$ColorScheme = 'Medium' 
)

######################
# FUNCTIONS
######################

function Write-LogMessage {
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
		} else {
			Write-Host $Message
		}
	}
	catch {
		throw
	}
}


######################
# VARIABLES
######################
$Inventory = $null
$ProgressId = Get-Random
$ProgressActivity = 'Convert-SqlServerInventoryClixmlToExcel'
$ProgressStatus = $null


######################
# BEGIN SCRIPT
######################

# Import Modules that we need
Import-Module -Name LogHelper, SqlServerInventory

# Set logging variables
Set-LogFile -Path $LogPath
Set-LoggingPreference -Preference $LoggingPreference

$ProgressStatus = "Starting Script: $($MyInvocation.MyCommand.Path)"
Write-LogMessage -Message $ProgressStatus -MessageLevel Information
Write-Progress -Activity $ProgressActivity -PercentComplete 0 -Status $ProgressStatus -Id $ProgressId

$ProgressStatus = "Loading inventory from '$FromPath'"
Write-LogMessage -Message $ProgressStatus -MessageLevel Information
Write-Progress -Activity $ProgressActivity -PercentComplete 25 -Status $ProgressStatus -Id $ProgressId

$Inventory = Import-SqlServerInventoryFromGzClixml -Path $FromPath

$ProgressStatus = "Writing inventory to Excel (Go get a coffee, this can take a few minutes...)"
Write-LogMessage -Message $ProgressStatus -MessageLevel Information
Write-Progress -Activity $ProgressActivity -PercentComplete 75 -Status $ProgressStatus -Id $ProgressId

Export-SqlServerInventoryToExcel `
-SqlServerInventory $Inventory `
-DirectoryPath $ToDirectoryPath `
-BaseFilename $([System.IO.Path]::GetFileNameWithoutExtension($FromPath)) `
-ColorTheme $ColorTheme `
-ColorScheme $ColorScheme `
-ParentProgressId $ProgressId


$ProgressStatus = "End Script: $($MyInvocation.MyCommand.Path)"
Write-LogMessage -Message $ProgressStatus -MessageLevel Information
Write-Progress -Activity $ProgressActivity -PercentComplete 100 -Status $ProgressStatus -Id $ProgressId -Completed

# Remove Variables
Remove-Variable -Name Inventory, ProgressId, ProgressActivity, ProgressStatus

# Remove Modules
Remove-Module -Name SqlServerInventory, LogHelper

# Call garbage collector
[System.GC]::Collect()

