@{

	# Script module or binary module file associated with this manifest
	ModuleToProcess = 'WindowsMachineInformation'

	# Version number of this module.
	ModuleVersion = '1.0.0.0'

	# ID used to uniquely identify this module
	GUID = '{f885012a-cde3-40d4-a7a2-428fb6047839}'

	# Author of this module
	Author = 'Kendal Van Dyke'

	# Company or vendor of this module
	CompanyName = 'Kendal Van Dyke'

	# Copyright statement for this module
	Copyright = '(c) 2013. All rights reserved.'

	# Description of the functionality provided by this module
	Description = 'Retrieves information about a Windows Machine using WMI and registry calls'

	# Minimum version of the Windows PowerShell engine required by this module
	PowerShellVersion = '2.0'

	# Minimum version of the .NET Framework required by this module
	DotNetFrameworkVersion = '2.0'

	# Minimum version of the common language runtime (CLR) required by this module
	CLRVersion = '2.0.50727'

	# Processor architecture (None, X86, Amd64, IA64) required by this module
	ProcessorArchitecture = 'None'

	# Modules that must be imported into the global environment prior to importing
	# this module
	RequiredModules = @()

	# Assemblies that must be loaded prior to importing this module
	RequiredAssemblies = @()

	# Script files (.ps1) that are run in the caller's environment prior to
	# importing this module
	ScriptsToProcess = @()

	# Type files (.ps1xml) to be loaded when importing this module
	TypesToProcess = @()

	# Format files (.ps1xml) to be loaded when importing this module
	FormatsToProcess = @()

	# Modules to import as nested modules of the module specified in
	# ModuleToProcess
	NestedModules = @()

	# Functions to export from this module
	FunctionsToExport = @('Get-WindowsMachineInformation')

	# Cmdlets to export from this module
	CmdletsToExport = @()

	# Variables to export from this module
	VariablesToExport = '*'

	# Aliases to export from this module
	AliasesToExport = @()

	# List of all modules packaged with this module
	ModuleList = @()

	# List of all files packaged with this module
	FileList = @('WindowsMachineInformation.psm1')

	# Private data to pass to the module specified in ModuleToProcess
	PrivateData = @{}

}