Requirements:
- Windows Powershell 2.0 on all machines 
- SQL Server Shared Management Objects (SMO) on the machine collecting the inventory (http://www.microsoft.com/en-us/download/details.aspx?id=29065)
- Excel 2007 or higher on the machine generating the report

NOTE: Running a SQL Server Inventory will also perform a Windows Inventory!

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



There are two steps to collect an inventory. Step 1 runs on a machine on the network you're collecting an inventory from and step 2 uses your machine to create the Excel file

0) Follow installation steps on both the remote machine and your local machine

1) On the remote machine:
	- Either:
		- Log into the machine with credentials that have administrator access to the computers you're collecting information from and start a PowerShell console session
			Start -> All Programs -> Accessories -> Windows PowerShell -> Windows PowerShell
			
		- OR log into the machine with a non-domain admin account and start a PowerShell session with credentials that have administrator access to the computers you're collecting information from. Replace DOMAIN\USERNAME with the appropriate credentials
			Start -> Run -> runas /netonly /user:DOMAIN\USERNAME "%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe"
	
	- Change the current location to where your scripts live
		sl "$([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments))\WindowsPowerShell"
	
	- Execute .\Get-SqlServerInventoryToClixml.ps1 with the with the appropriate parameters.
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

			-Username: SQL Server username to use when connecting to an instance. If this is omitted Windows authentication will be used instead.
			-Password: SQL Server password to use when connecting to an instance

			AND

			-PrivateOnly: If present, only probe private IP addresses (default is all IP addresses)
			-MaxConcurrencyThrottle: Number between 1-100 to indicate how many machine scans to run concurrently. (default is # of logical CPUs presented to your OS)
			-DirectoryPath: Path to write the XML and log file. (default "My Documents\")
			-LoggingPreference: Indicates the level to log information. Options are: none, standard, verbose, or debug (default is none)
			-Zip: If present, add the XML and Log File (if one was written) to a ZIP file in the directory specified by the DirectoryPath parameter
			-IncludeDatabaseObjectInformation: If present, include database object information (i.e. tables, views, procedures, udfs, etc.). System objects are NOT included by default
			-IncludeDatabaseObjectPermissions: If present, include database object information (i.e. tables, views, procedures, udfs, etc.). System objects are NOT included by default
			-IncludeDatabaseSystemObjects: If present, include system objects when retrieving database object information. This has no effect if -IncludeDatabaseObjectInformation is not specified

			
		- Example calls:
			.\Get-SqlServerInventoryToClixml.ps1 -DNSServer automatic -DNSDomain automatic -LoggingPreference standard -PrivateOnly
			.\Get-SqlServerInventoryToClixml.ps1 -Subnet 172.20.40.0/28 -LoggingPreference standard -PrivateOnly
			.\Get-SqlServerInventoryToClixml.ps1 -DNSServer automatic -domain automatic -LoggingPreference verbose -MaxConcurrencyThrottle 10 -PrivateOnly -Zip
			.\Get-SqlServerInventoryToClixml.ps1 -ComputerName $env:COMPUTERNAME -LoggingPreference standard -Zip
			.\Get-SqlServerInventoryToClixml.ps1 -ComputerName SQL1, SQL2, SQL3 -LoggingPreference verbose -Zip
			.\Get-SqlServerInventoryToClixml.ps1 -ComputerName SQL1, SQL2, SQL3 -LoggingPreference verbose -Zip -IncludeDatabaseObjectInformation
			.\Get-SqlServerInventoryToClixml.ps1 -ComputerName SQL1, SQL2, SQL3 -LoggingPreference verbose -Zip -IncludeDatabaseObjectInformation -IncludeDatabaseSystemObjects

			
	- Copy the XML\ZIP file created by Get-SqlServerInventoryToClixml.ps1 to your local machine (and uncompress if it's a ZIP)
		- The XML file can be several Megabytes but will compress significantly; consider using the -Zip option to create a smaller file for transferring to another machine

		
- On your local machine:
	- Start a PowerShell console session
		Start -> All Programs -> Accessories -> Windows PowerShell -> Windows PowerShell
		
	- Change the current location to where your scripts live
		sl "$([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments))\WindowsPowerShell"
		
	- Execute .\Convert-SqlServerInventoryClixmlToExcel.ps1 with the appropriate parameters
		- Parameters
			-FromPath: The path to the XML file created by .\Get-SqlServerInventoryToClixml.ps1
			-ToDirectoryPath: The directory path to write the Excel files. (default is same as FromPath if this parameter is not provided)
			-LoggingPreference: Indicates the level to log information. Options are: none, standard, verbose, or debug (default is none)
			-LogPath: Path to write the log file. (default "SQL Server Inventory - YYYY-MM-DD-HH-MM.log" in the same directory as -ToDirectoryPath)
			-ColorTheme: Corresponds to the name of a Microsoft Office Color Theme (e.g. "Waveform" and "Slipstream" in Office 2010 or "Blue" and "Green" in Office 2013). Default is "Office"
			-ColorScheme: "light", "medium", or "dark" (Default is "medium")
			
		- Example calls:
			.\Convert-SqlServerInventoryClixmlToExcel.ps1 -FromPath "C:\Users\kvandyke\Documents\SQL Server Inventory - 2012-05-08-17-01.xml" -LoggingPreference standard
			.\Convert-SqlServerInventoryClixmlToExcel.ps1 -FromPath "C:\Users\kvandyke\Documents\SQL Server Inventory - 2012-05-08-17-01.xml" -ColorTheme Blue
			.\Convert-SqlServerInventoryClixmlToExcel.ps1 -FromPath "C:\Users\kvandyke\Documents\SQL Server Inventory - 2012-05-08-17-01.xml" -ColorTheme Blue -ColorScheme dark