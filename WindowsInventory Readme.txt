Requirements:
- Windows Powershell 2.0 on all machines 
- Excel on the machine generating the report

Please send all feedback to kendal.vandyke@gmail.com


How to install:
1) Open a Powershell prompt in elevated mode
	Start -> All Programs -> Accessories -> Windows PowerShell -> Windows PowerShell (right click, choose "Run as Administrator")

2) Set the execution policy to allow for remotely signed scripts
	Set-ExecutionPolicy RemoteSigned -Force

3) Create Powershell and Powershell modules directory 
	New-Item -type directory -path "$([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments))\WindowsPowerShell\Modules"

4) Ensure ZIP file is not blocked by Windows
	Navigate to zip file in Windows Explorer
	Right click -> Properties
	On the general tab, if a security warning appears at the bottom of the window indicating that the file came from another computer click the "unblock" button
	
5) Extract contents of Zip file to the WindowsPowerShell directory in your "My Documents" folder



There are two ways to collect an inventory:

1) All in one script: Easiest when your machine is physically connected to the network you're taking an inventory of.
	- Start a PowerShell console session with credentials that have administrator access to the computers you're collecting information from. 
		- Either:
			- If your current login has the necessary credentials start a PowerShell console session
				Start -> All Programs -> Accessories -> Windows PowerShell -> Windows PowerShell
				
			- OR start a PowerShell console as another user that has administrator access (Replace DOMAIN\USERNAME with the appropriate credentials)
				Start -> Run -> runas /netonly /user:DOMAIN\USERNAME "%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe"
	
	- Change the current location to where your scripts live
		sl "$([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments))\WindowsPowerShell"
	
	- Execute .\Get-WindowsInventoryToExcel.ps1 with the appropriate parameters.
		- Parameters
			-DNSServer: "automatic" or a comma delimited list of IP addresses of AD DNS servers to query for DNS A records
			-DNSDomain: "automatic" or the AD domain name to query for DNS A Records (default is automatic if no value provided)
			-ExcludeSubnet: a comma delimited list of CIDR notation subnets to exclude from scanning
			-LimitSubnet: a comma delimited list of CIDR notation subnets to limit scanning to
			-ExcludeComputerName: a comma delimited list of computer names to exclude from scanning
			OR
			-Subnet: "automatic" or a comma delimited list of CIDR notation subnets to scan for hosts
			-LimitSubnet: a comma delimited list of CIDR notation subnets to limit scanning to
			-ExcludeComputerName: a comma delimited list of computer names to exclude from scanning
			OR
			-ComputerName: a comma delimited list of computer names that you want to inventory
			
			AND
			
			-PrivateOnly: If present, only probe private IP addresses (default is all IP addresses)
			-MaxConcurrencyThrottle: Number between 1-100 to indicate how many machine scans to run concurrently. (default is # of logical CPUs presented to your OS)
			-AdditionalData: None, or a comma delimited list of additional data points to collect from the following list: (Default is all)
						AdditionalHardware,BIOS,EventLog,FullyQualifiedDomainName,InstalledApplications,InstalledPatches,IPRoutes, `
						LastLoggedOnUser,LocalGroups,LocalUserAccounts,PowerPlans,Printers,PrintSpoolerLocation,Processes, `
						ProductKeys,RegistrySize,Services,Shares,StartupCommands,WindowsComponents
			-Filename: Path to write the Excel file. (default "My Documents\Windows Inventory - YYYY-MM-DD-HH-MM.xlsx")
			-LoggingPreference: Indicates the level to log information. Options are: none, standard, verbose, or debug (default is none)
			-LogPath: Path to write the log file. (default "My Documents\Windows Inventory - YYYY-MM-DD-HH-MM.log")
			-Zip: If present, add the Excel and Log File (if one was written) to a ZIP file in the same path as the Excel file and as the same name but with a ZIP extension
			
		- Example calls:
			.\Get-WindowsInventoryToExcel.ps1 -DNSServer automatic -DNSDomain automatic -LoggingPreference standard -PrivateOnly
			.\Get-WindowsInventoryToExcel.ps1 -Subnet 172.20.40.0/28 -LoggingPreference standard -PrivateOnly -Zip
			.\Get-WindowsInventoryToExcel.ps1 -ComputerName $env:COMPUTERNAME -LoggingPreference standard -Zip


			
2) Two scripts: Easiest when you have access to a machine on the network you're collecting an inventory from but want to use your machine to create the Excel file
	- Follow installation steps on both the remote machine and your local machine
	- On the remote machine:
		- Either:
			- Log into the machine with credentials that have administrator access to the computers you're collecting information from and start a PowerShell console session
				Start -> All Programs -> Accessories -> Windows PowerShell -> Windows PowerShell
				
			- OR log into the machine with a non-domain admin account and start a PowerShell session with credentials that have administrator access to the computers you're collecting information from. Replace DOMAIN\USERNAME with the appropriate credentials
				Start -> Run -> runas /netonly /user:DOMAIN\USERNAME "%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe"
		
		- Change the current location to where your scripts live
			sl "$([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments))\WindowsPowerShell"
		
		- Execute .\Get-WindowsInventoryToClixml.ps1 with the with the appropriate parameters.
			- Parameters
				-DNSServer: "automatic" or a comma delimited list of IP addresses of AD DNS servers to query for DNS A records
				-DNSDomain: "automatic" or the AD domain name to query for DNS A Records (default is automatic if no value provided)
				-ExcludeSubnet: a comma delimited list of CIDR notation subnets to exclude from scanning
				-LimitSubnet: a comma delimited list of CIDR notation subnets to limit scanning to
				-ExcludeComputerName: a comma delimited list of computer names to exclude from scanning
				OR
				-Subnet: "automatic" or a comma delimited list of CIDR notation subnets to scan for hosts
				-LimitSubnet: a comma delimited list of CIDR notation subnets to limit scanning to
				-ExcludeComputerName: a comma delimited list of computer names to exclude from scanning
				OR
				-ComputerName: a comma delimited list of computer names that you want to inventory

				AND
				-PrivateOnly: If present, only probe private IP addresses (default is all IP addresses)
				-MaxConcurrencyThrottle: Number between 1-100 to indicate how many machine scans to run concurrently. (default is # of logical CPUs presented to your OS)
				-AdditionalData: None, or a comma delimited list of additional data points to collect from the following list: (Default is all)
							AdditionalHardware,BIOS,EventLog,FullyQualifiedDomainName,InstalledApplications,InstalledPatches,IPRoutes, `
							LastLoggedOnUser,LocalGroups,LocalUserAccounts,PowerPlans,Printers,PrintSpoolerLocation,Processes, `
							ProductKeys,RegistrySize,Services,Shares,StartupCommands,WindowsComponents
				-DirectoryPath: Path to write the XML and log file. (default "My Documents\")
				-LoggingPreference: Indicates the level to log information. Options are: none, standard, verbose, or debug (default is none)
				-Zip: If present, add the XML and Log File (if one was written) to a ZIP file in the same path as the XML file and as the same name but with a ZIP extension

						
			- Example calls:
				.\Get-WindowsInventoryToClixml.ps1 -DNSServer automatic -DNSDomain automatic -LoggingPreference standard -PrivateOnly
				.\Get-WindowsInventoryToClixml.ps1 -Subnet 172.20.40.0/28 -LoggingPreference standard -PrivateOnly
				.\Get-WindowsInventoryToClixml.ps1 -DNSServer automatic -domain yesprep.local -LoggingPreference verbose -MaxConcurrencyThrottle 10 -PrivateOnly -Zip
				.\Get-WindowsInventoryToClixml.ps1 -ComputerName $env:COMPUTERNAME -LoggingPreference standard -Zip

		- Copy the XML\ZIP file created by Get-WindowsInventoryToClixml.ps1 to your local machine (and uncompress if it's a ZIP)
			- The XML file can be several Megabytes but will compress significantly; consider using the -Zip option to create a smaller file for transferring to another machine

			
	- On your local machine:
		- Start a PowerShell console session
			Start -> All Programs -> Accessories -> Windows PowerShell -> Windows PowerShell
			
		- Change the current location to where your scripts live
			sl "$([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments))\WindowsPowerShell"
			
		- Execute .\Convert-WindowsInventoryClixmlToExcel.ps1 with the appropriate parameters
			- Parameters
				-FromPath: The path to the XML file created by .\Get-WindowsInventoryToClixml.ps1
				-ToPath: The path to write the Excel file. (default is same as FromPath with XLSX extension)
				-LoggingPreference: Indicates the level to log information. Options are: none, standard, verbose, or debug (default is none)
				-LogPath: Path to write the log file. (default "My Documents\WindowsInventory-YYYY-MM-DD-HH-MM.log")
				-ColorTheme: Corresponds to the name of a Microsoft Office Color Theme (e.g. "Waveform" and "Slipstream" in Office 2010 or "Blue" and "Green" in Office 2013). Defualt is "Office"
				-ColorScheme: "light", "medium", or "dark" (Default is "medium")
				
			- Example calls:
				.\Convert-WindowsInventoryClixmlToExcel.ps1 -FromPath "C:\Users\kvandyke\Documents\WindowsInventory-2012-05-08-17-01.xml" -ColorTheme Blue -ColorScheme dark