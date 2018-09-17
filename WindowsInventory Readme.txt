NOTE: Running a SQL Server Inventory will also perform a Windows Inventory!

Requirements
-------------
Before you can document your Windows environment with SQL Power Doc you'll to meet the following requirements:

	Permissions
		- The account used to run SQL Power Doc will need Administrator rights to the OS in order to collect information about the Operating System.

	Windows Machine To Perform Inventory
		- This can be virtual or physical - either way it's recommended that its located on the same physical network as the servers you are collecting information from. SQL Power Doc collects a lot of information and you don't want network latency to become a bottleneck
			
		- PowerShell likes memory; For documenting a 10-20 server environment you'll want at least 2 GB of RAM available. Logical CPU count isn't as important, but SQL Power Doc can split its workload across multiple CPUs when doing it's thing so the more CPUs you've got the faster the job will get done. 

		- The following software needs to be installed: 
			-- Windows PowerShell 2.0 or higher:
				Windows PowerShell 2.0 is available on on all Windows Operating Systems going back to Windows Server 2003 and Windows XP. Chances are you've already got it installed and enabled, but in case you're not sure head over to  http://support.microsoft.com/kb/968929 for a list of requirements and instructions on how to get PowerShell working on your system.

		- Sometimes firewall and group policy restrictions will prevent SQL Power Doc from gathering information from servers. If that's your environment then you'll want to make sure to open up communications for both SQL Server and WMI (Windows Management Instrumentation). Start at  http://msdn.microsoft.com/en-us/library/windows/desktop/aa822854(v=vs.85).aspx for instructions on how to do so.

	Windows Machine To Create The Documentation
		- Usually this will be your laptop or desktop 
		
		- You'll need the following software installed: 
			-- Windows PowerShell 2.0 or higher 
			-- Microsoft Excel 2007 or higher
		
You can use the same machine to collect an inventory and create the documentation as long as it meets all the requirements outlined above.



Configure Windows PowerShell
----------------------------
You'll want to make sure that PowerShell is configured properly on both the machine that you're using to perform the inventory and the machine that's building the documentation. (Repeat: Do this on both machines!)

Set Execution Policy
	By default PowerShell tries to keep you from shooting yourself in the foot by not letting you run scripts that you download from the internet. In PowerShell lingo this is referred to as the Execution Policy (see  http://technet.microsoft.com/en-us/library/hh847748.aspx) and in order to use SQL Power Doc you'll need to change it by following these steps:
		1.Open a PowerShell console in elevated mode:
			Start -> All Programs -> Accessories -> Windows PowerShell -> Windows PowerShell (right click, choose "Run as Administrator")
		
		2.Set the execution policy to allow for remotely signed scripts
			Set-ExecutionPolicy RemoteSigned -Force
		
		3.Exit the PowerShell console

Configure Windows PowerShell Directory
	Now you'll need to create a directory to hold PowerShell code.
		1.Open a new PowerShell console (but not in elevated mode as when you set the execution policy):
			Start -> All Programs -> Accessories -> Windows PowerShell -> Windows PowerShell 

		2.Create PowerShell and PowerShell modules directory in your "My Documents" folder
			New-Item -type directory -path "$([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments))\WindowsPowerShell\Modules"

		3.Exit the PowerShell console

		
Download And Install
--------------------
Grab the latest version of the code from the  downloads page but don't extract the ZIP file yet! Because the file came from the internet it needs to be unblocked or PowerShell gets cranky because it's considered untrusted.

To unblock a file, navigate to it in Windows Explorer, right click, and choose the Properties menu option. On the General tab, click the Unblock button, then click the OK button to close the Properties dialog.

Once the file is unblocked you can extract the contents to the WindowsPowerShell folder (in your "My Documents" directory) that you created in the last step.

Note: Make sure to keep the folder names in the zip file intact so that everything in the Modules folder is extracted into WindowsPowerShell\Modules and the .ps1 files are extracted into the WindowsPowerShell folder.

The Windows Inventory portion of SQL Power Doc will attempt to use the RDS-Manager PowerShell module provided by the Microsoft Remote Desktop Services team. You can download this optional module from  http://gallery.technet.microsoft.com/ScriptCenter/e8c3af96-db10-45b0-88e3-328f087a8700/ . Make sure to save it in the WindowsPowerShell\Modules\RDS-Manager folder (in your "My Documents" directory) that you recently created. It's not the end of the world if it's missing but you'll get more details about users' desktop sessions when it's installed.


Collect Windows Inventory
-------------------------------------------
So far, so good...now it's time to discover your Windows Servers and perform an inventory! In this step you're going to run a PowerShell script which will discover Windows machines on your network (or verify they're running), collect information about them and their underlying OS, and write the results as an XML file that you'll use in the next step.

Start by opening a PowerShell console on the machine that will be collecting the information from your SQL Servers and set your current location to the WindowsPowerShell folder:
	Set-Location "$([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments))\WindowsPowerShell"

Running The Script, Choosing The Right Parameters

	You're going to execute the script .\Get-WindowsInventoryToClixml.ps1 to do all the work but it requires a few parameters to know what to do.

	Discover & Verify Windows Machines

		The first set of parameters define how to find & verify Windows machines.

			Find Windows machines by querying Active Directory DNS for hosts

				-DnsServer
				Valid values are "automatic" or a comma delimited list of AD DNS server IP addresses to query for DNS A records that may be Windows machines.

				-DnsDomain
				Optional. Valid values are "automatic" (the default if this parameter is not specified) or the AD domain name to query for DNS records from.

				-ExcludeSubnet
				Optional. This is a comma delimited list of  CIDR notation subnets to exclude when looking for Windows machines.

				-LimitSubnet
				Optional. This is an inclusive comma delimited list of CIDR notation subnets to limit the scope when looking for Windows machines.

				-ExcludeComputerName
				Optional. This is a comma delimited list of computer names to exclude when looking for Windows machines.

				-PrivateOnly
				Optional. This switch limits the scope to private class A, B, and C IP addresses when looking for Windows machines.


			Find Windows machines by scanning a subnet of IP Addresses

				-Subnet
				Valid values are "automatic" or a comma delimited list of CIDR notation subnets to scan for IPv4 hosts that may be Windows machines.

				-LimitSubnet
				Optional. This is an inclusive comma delimited list of CIDR notation subnets to limit the scope when looking for Windows machines.

				-ExcludeComputerName
				Optional. This is a comma delimited list of computer names to exclude when looking for Windows machines.

				-PrivateOnly
				Optional. This switch limits the scope to private class A, B, and C IP addresses when looking for Windows machines.


			Find Windows machines by computer name

				-ComputerName
				This is a comma delimited list of computer names that may be Windows machines.

				-PrivateOnly
				Optional. This switch limits the scope to private class A, B, and C IP addresses when looking for Windows machines.


	Additional Information To Collect

		By default, SQL Power Doc collects a minimal set of information about each Windows machine it finds. You can collect more information with the following parameter:

			-AdditionalData
			This is a comma delimited list of one or more of the following additional data points to collect:
				AdditionalHardware 
				BIOS 
				DesktopSessions 
				EventLog 
				FullyQualifiedDomainName 
				InstalledApplications 
				InstalledPatches 
				IPRoutes 
				LastLoggedOnUser 
				LocalGroups 
				LocalUserAccounts 
				PowerPlans 
				Printers 
				PrintSpoolerLocation 
				Processes 
				ProductKeys 
				RegistrySize 
				Services 
				Shares 
				StartupCommands 
				WindowsComponents

			Alternatively, you can specify the value "All" to include all of the data points listed above. 

			If this parameter is not provided the default is that none of the listed data points will be included.


	Logging, Output, & Resource Utilization

		Finally, the following parameters control logging, output, and system resources SQL Power Doc will use when finding and collecting information from Windows machines:

			-DirectoryPath
			Optional. A fully qualified directory path where all output will be written. The default value is your "My Documents" folder.

			-LoggingPreference
			Optional. Specifies how much information will be written to a log file (in the same directory as the output file). Valid values are None, Standard, Verbose, and Debug. The default value is None (i.e. no logging).

			-Zip
			Optional. If provided, create a Zip file containing all output in the directory specified by the DirectoryPath parameter.

			-MaxConcurrencyThrottle
			Optional. A number between 1-100 which indicates how many tasks to perform concurrently. The default is the number of logical CPUs present on the OS.

		
	Examples

		The following examples demonstrate how to combine all the parameters together when running the script.

		Example 1:

		.\Get-WindowsInventoryToClixml.psm1 -DNSServer automatic -DNSDomain automatic -PrivateOnly

			Collect an inventory by querying Active Directory for a list of hosts to scan for Windows machines. The list of hosts will be restricted to private IP addresses only.

			The Inventory file will be written to your "My Documents" folder.

			No Log file will be written.


		Example 2:

		.\Get-WindowsInventoryToClixml.psm1 -Subnet 172.20.40.0/28 -LoggingPreference Standard

			Collect an inventory by scanning all hosts in the subnet 172.20.40.0/28 for Windows machines.

			The Inventory and Log files will be written to your "My Documents" folder.

			Standard logging will be used.

		 
		Example 3:

		.\Get-WindowsInventoryToClixml.psm1 -Computername Server1,Server2,Server3

			Collect an inventory by scanning Server1, Server2, and Server3 for Windows machines. 

			The Inventory file will be written to your "My Documents" folder.

			No Log file will be written.
		 

		Example 4:

		.\Get-WindowsInventoryToClixml.psm1 -Computername $env:COMPUTERNAME -AdditionalData None -LoggingPreference Verbose

			Collect an inventory by scanning the local machine for Windows machines.

			Do not collect any data beyond the core set of information.

			The Inventory and Log files will be written to your "My Documents" folder.

			Verbose logging will be used.

 

How Long Will It Take?
----------------------
When run with the defaults on a machine with 2 CPUs you can expect the script to take about 5 minutes to complete an inventory of 20 machines. Your mileage will vary depending on how many machines you are collecting information from and what additional information you're including.

Progress is written to the PowerShell console to give you a better idea what the script's up to. If you've got logging enabled (highly recommended) you can also check the logs for progress updates.



Generate A Windows Inventory Report
-------------------------------
Once the inventory collection phase is complete you'll want to copy the output file to the machine where you'll create the inventory reports (Excel workbooks).

To create an inventory report, start by opening a PowerShell console and set your current location to the WindowsPowerShell folder:
	Set-Location "$([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments))\WindowsPowerShell"

This time you're going to execute the script .\Convert-WindowsInventoryClixmlToExcel.ps1 and supply the following parameters:

	-FromPath

	The literal path to the output file created by Get-WindowsInventoryToClixml.ps1.

	-ToDirectoryPath
	Optional. Specifies the literal path to the directory where the Excel workbooks will be written. This path must exist prior to executing the script. If this parameter is not provided the workbooks will be written to the same directory specified in the FromPath parameter. Assuming the XML file specified in FromPath is named "Windows Inventory.xml" then the Excel file will be written to "Windows Inventory.xlsx".


	-ColorTheme
	Optional. An Office Theme Color to apply to each worksheet. If not specified or if an unknown theme color is provided the default "Office" theme colors will be used.

	Office 2013 theme colors include: 
		Aspect, Blue Green, Blue II, Blue Warm, Blue, Grayscale, Green Yellow, Green, Marquee, Median, Office, Office 2007 - 2010, Orange Red, Orange, Paper, Red Orange, Red Violet, Red, Slipstream, Violet II, Violet, Yellow Orange, Yellow

	Office 2010 theme colors include: 
		Adjacency, Angles, Apex, Apothecary, Aspect, Austin, Black Tie, Civic, Clarity, Composite, Concourse, Couture, Elemental, Equity, Essential, Executive, Flow, Foundry, Grayscale, Grid, Hardcover, Horizon, Median, Metro, Module, Newsprint, Office, Opulent, Oriel, Origin, Paper, Perspective, Pushpin, Slipstream, Solstice, Technic, Thatch, Trek, Urban, Verve, Waveform

	Office 2007 theme colors include: 
		Apex, Aspect, Civic, Concourse, Equity, Flow, Foundry, Grayscale, Median, Metro, Module, Office, Opulent, Oriel, Origin, Paper, Solstice, Technic, Trek, Urban, Verve


	-ColorScheme
	Optional. The color theme to apply to each worksheet. Valid values are "Light", "Medium", and "Dark". If not specified then "Medium" is used as the default value .


	-LoggingPreference
	Optional. Specifies how much information will be written to a log file (location specified in the LogPath parameter). Valid values are None, Standard, Verbose, and Debug. The default value is None (i.e. no logging).

	
	-LogPath
	Optional. A literal path to a log file to write details about what this script is doing. The filename does not need to exist prior to executing this script but the specified directory does. 

	If a LoggingPreference other than "None" is specified and this parameter is not provided then the file is named "SQL Server Inventory - [Year][Month][Day][Hour][Minute].log" and is written to the same directory specified by the ToDirectoryPath paramter.

 

Examples

The following examples demonstrate how to combine all the parameters together when running the script.

	Example 1:

	.\Convert-WindowsInventoryClixmlToExcel.ps1 -FromPath "C:\Inventory\Windows Inventory.xml"

		Writes an Excel file for the Windows Operating System information contained in "C:\Inventory\Windows Inventory.xml" to "C:\Inventory\Windows Inventory.xlsx". 

		The Office color theme and Medium color scheme will be used by default.

	
	Example 2:

	.\Convert-WindowsInventoryClixmlToExcel.ps1 -FromPath "C:\Inventory\Windows Inventory.xml"  -ColorTheme Blue -ColorScheme Dark

		Writes an Excel file for the Windows Operating System information contained in "C:\Inventory\Windows Inventory.xml" to "C:\Inventory\Windows Inventory.xlsx".  

		The Blue color theme and Dark color scheme will be used.

 
 
Additional Help
---------------
If you're still having problems using SQL Power Doc after reading through this guide please post in the Discussions (https://sqlpowerdoc.codeplex.com/discussions) or reach out to @SQLDBA on Twitter.

	
