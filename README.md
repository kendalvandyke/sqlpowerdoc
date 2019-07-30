# sqlpowerdoc
SQL Power Doc is a collection of Windows PowerShell scripts and modules that discover, document, and diagnose SQL Server instances and their underlying Windows OS &amp; machine configurations.

**NOTE: Running a SQL Server Inventory will also perform a Windows Inventory!**


## Requirements
Before you can document your SQL Server environment with SQL Power Doc you'll to meet the following requirements:

	Permissions
		- SQL Power Doc makes connections to standalone SQL Server instances using either Windows Authentication or with a SQL Server username and password. Whichever way you connect, the login will need to be a member of the sysadmin server role on all standalone SQL Server instances you're documenting. 
		
		- For Windows Azure SQL Database (WASD) a SQL username and password is the only way you can connect. This login should be the WASD Administrator login. 
		
		- SQL Power Doc also tries to collect information about the Operating System that SQL Server in installed on. The account used to run SQL Power Doc will need Administrator rights to the OS in order to do this part.

	Windows Machine To Perform Inventory
		- This can be virtual or physical - either way it's recommended that its located on the same physical network as the servers you are collecting information from. SQL Power Doc collects a lot of information and you don't want network latency to become a bottleneck
			
			-- You can ignore this requirement if you're documenting a Windows Azure SQL Database since it's not likely that you'll have physical access to the hardware these databases run on!

		- Just like SQL Server, PowerShell likes memory; For documenting a 10-20 server environment you'll want at least 2 GB of RAM available. Logical CPU count isn't as important, but SQL Power Doc can split its workload across multiple CPUs when doing it's thing so the more CPUs you've got the faster the job will get done. 

		- The following software needs to be installed: 
			-- Windows PowerShell 2.0 or higher:
				Windows PowerShell 2.0 is available on on all Windows Operating Systems going back to Windows Server 2003 and Windows XP. Chances are you've already got it installed and enabled, but in case you're not sure head over to  http://support.microsoft.com/kb/968929 for a list of requirements and instructions on how to get PowerShell working on your system.

			-- SQL Server Management Objects (SMO):
				If SQL Server Management Studio is installed on this machine then it's already got SMO. 
				
				You don't need the absolute latest version installed, but it's a good idea to make sure that the version you do have installed at least matches the highest version of SQL Server that will be included in your inventory 
				
				SMO is part of the SQL 2012 Feature Pack and can be downloaded for free from  http://www.microsoft.com/en-us/download/details.aspx?id=29065 (Note - SMO requires the System CLR Types which are on the same download page)

		- Sometimes firewall and group policy restrictions will prevent SQL Power Doc from gathering information from servers. If that's your environment then you'll want to make sure to open up communications for both SQL Server and WMI (Windows Management Instrumentation). Start at  http://msdn.microsoft.com/en-us/library/windows/desktop/aa822854(v=vs.85).aspx for instructions on how to do so.

	Windows Machine To Create The Documentation
		- Usually this will be your laptop or desktop 
		
		- You'll need the following software installed: 
			-- Windows PowerShell 2.0 or higher 
			-- Microsoft Excel 2007 or higher
		
		- You do not need SMO or access to SQL Server for this step

You can use the same machine to collect an inventory and create the documentation as long as it meets all the requirements outlined above. This will be the case if you're documenting one or more Windows Azure SQL Databases.


## Configure Windows PowerShell
You'll want to make sure that PowerShell is configured properly on both the machine that you're using to perform the inventory and the machine that's building the documentation. (Repeat: Do this on both machines!)

### Set Execution Policy
By default PowerShell tries to keep you from shooting yourself in the foot by not letting you run scripts that you download from the internet. In PowerShell lingo this is referred to as the Execution Policy (see  http://technet.microsoft.com/en-us/library/hh847748.aspx) and in order to use SQL Power Doc you'll need to change it by following these steps:
		
1. Open a PowerShell console in elevated mode:
`Start -> All Programs -> Accessories -> Windows PowerShell -> Windows PowerShell` (right click, choose "Run as Administrator")
2. Set the execution policy to allow for remotely signed scripts
`Set-ExecutionPolicy RemoteSigned -Force`
3. Exit the PowerShell console

### Configure Windows PowerShell Directory
Now you'll need to create a directory to hold PowerShell code.
		
1. Open a new PowerShell console (but not in elevated mode as when you set the execution policy):
`Start -> All Programs -> Accessories -> Windows PowerShell -> Windows PowerShell`
2. Create PowerShell and PowerShell modules directory in your "My Documents" folder
`New-Item -type directory -path "$([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments))\WindowsPowerShell\Modules"`
3. Exit the PowerShell console

		
## Download And Install
Grab the latest version of the code from the  downloads page but don't extract the ZIP file yet! Because the file came from the internet it needs to be unblocked or PowerShell gets cranky because it's considered untrusted.

To unblock a file, navigate to it in Windows Explorer, right click, and choose the Properties menu option. On the General tab, click the Unblock button, then click the OK button to close the Properties dialog.

Once the file is unblocked you can extract the contents to the WindowsPowerShell folder (in your "My Documents" directory) that you created in the last step.

Note: Make sure to keep the folder names in the zip file intact so that everything in the Modules folder is extracted into WindowsPowerShell\Modules and the .ps1 files are extracted into the WindowsPowerShell folder.

The Windows Inventory portion of SQL Power Doc will attempt to use the RDS-Manager PowerShell module provided by the Microsoft Remote Desktop Services team. You can download this optional module from  http://gallery.technet.microsoft.com/ScriptCenter/e8c3af96-db10-45b0-88e3-328f087a8700/ . Make sure to save it in the WindowsPowerShell\Modules\RDS-Manager folder (in your "My Documents" directory) that you recently created. It's not the end of the world if it's missing but you'll get more details about users' desktop sessions when it's installed.


## Collect A SQL Server And Windows Inventory
So far, so good...now it's time to discover your SQL Servers and perform an inventory! In this step you're going to run a PowerShell script which will discover SQL Servers on your network (or verify they're running), collect information about them and their underlying OS, and write the results as a Gzip compressed XML file that you'll use in the next step.

Start by opening a PowerShell console on the machine that will be collecting the information from your SQL Servers and set your current location to the WindowsPowerShell folder:
	`Set-Location "$([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments))\WindowsPowerShell"`

Running The Script, Choosing The Right Parameters

	You're going to execute the script .\Get-SqlServerInventoryToClixml.ps1 to do all the work but it requires a few parameters to know what to do.

	Discover & Verify SQL Server Services

		The first set of parameters define how to find & verify machines with SQL Server services installed on them.

			Find SQL Server services by querying Active Directory DNS for hosts

				-DnsServer
				Valid values are "automatic" or a comma delimited list of AD DNS server IP addresses to query for DNS A records that may be running SQL Server.

				-DnsDomain
				Optional. Valid values are "automatic" (the default if this parameter is not specified) or the AD domain name to query for DNS records from.

				-ExcludeSubnet
				Optional. This is a comma delimited list of  CIDR notation subnets to exclude when looking for SQL Servers.

				-LimitSubnet
				Optional. This is an inclusive comma delimited list of CIDR notation subnets to limit the scope when looking for SQL Servers.

				-ExcludeComputerName
				Optional. This is a comma delimited list of computer names to exclude when looking for SQL Servers.

				-PrivateOnly
				Optional. This switch limits the scope to private class A, B, and C IP addresses when looking for SQL Servers.


			Find SQL Server services by scanning a subnet of IP Addresses

				-Subnet
				Valid values are "automatic" or a comma delimited list of CIDR notation subnets to scan for IPv4 hosts that may be running SQL Server.

				-LimitSubnet
				Optional. This is an inclusive comma delimited list of CIDR notation subnets to limit the scope when looking for SQL Servers.

				-ExcludeComputerName
				Optional. This is a comma delimited list of computer names to exclude when looking for SQL Servers.

				-PrivateOnly
				Optional. This switch limits the scope to private class A, B, and C IP addresses when looking for SQL Servers.


			Find SQL Server services by computer name

				-ComputerName
				This is a comma delimited list of computer names that may be running SQL Server.

				-PrivateOnly
				Optional. This switch limits the scope to private class A, B, and C IP addresses when looking for SQL Servers.


	Authentication

		SQL Power Doc will default to using Windows Authentication when attempting to connect to SQL Server instances that it finds. If you want to connect using SQL Server authentication instead, or if you are connecting to Windows Azure SQL Database instances, provide the following parameters:

			-Username
			SQL Server username to use when connecting to an instance.

			-Password
			SQL Server password to use when connecting to an instance.

	Additional Information To Collect

		By default, SQL Power Doc does not  collect Database object information (e.g. tables, views, procedures, etc.), Database object permissions, or information about Database system objects. The following switch parameters alter the default behavior:

			-IncludeDatabaseObjectInformation
			Include database object information (e.g. tables, views, procedures, functions, etc.).

			-IncludeDatabaseObjectPermissions
			Include database object permissions (e.g. GRANT SELECT on tables, columns, views, etc.).

			-IncludeDatabaseSystemObjects
			Include system objects when collecting information about database objects and\or database object permissions.

	Logging, Output, & Resource Utilization

		Finally, the following parameters control logging, output, and system resources SQL Power Doc will use when finding and collecting information from SQL Servers:

			-DirectoryPath
			Optional. A fully qualified directory path where all output will be written. The default value is your "My Documents" folder.

			-LoggingPreference
			Optional. Specifies how much information will be written to a log file (in the same directory as the output file). Valid values are None, Standard, Verbose, and Debug. The default value is None (i.e. no logging).

			-Zip
			Optional. If provided, create a Zip file containing all output in the directory specified by the DirectoryPath parameter.

			-MaxConcurrencyThrottle
			Optional. A number between 1-100 which indicates how many tasks to perform concurrently. The default is the number of logical CPUs present on the OS.

## Examples
The following examples demonstrate how to combine all the parameters together when running the script.

### Example 1:
```powershell
.\Get-SqlServerInventoryToClixml.ps1 -DNSServer automatic -DNSDomain automatic -PrivateOnly
```
Collect an inventory by querying Active Directory for a list of hosts to scan for SQL Server instances.

The list of hosts will be restricted to private IP addresses only.

Windows Authentication will be used to connect to each instance.

Database objects will NOT be included in the results.

The Inventory file will be written to your "My Documents" folder.

No log file will be written.

### Example 2:
```powershell
.\Get-SqlServerInventoryToClixml.ps1 -Subnet 172.20.40.0/28 -Username sa -Password BetterNotBeBlank
```
Collect an inventory by scanning all hosts in the subnet 172.20.40.0/28 for SQL Server instances.

SQL authentication (username = "sa", password = "BetterNotBeBlank") will be used to connect to the instance.

Database objects will NOT be included in the results.

The Inventory file will be written to your "My Documents" folder.

No log file will be written.
		 
### Example 3:
```powershell
.\Get-SqlServerInventoryToClixml.ps1 -Computername Server1,Server2,Server3 -LoggingPreference Standard
```
Collect an inventory by scanning Server1, Server2, and Server3 for SQL Server instances.

Windows Authentication will be used to connect to the instance.

Database objects will NOT be included in the results.

The Inventory file will be written to your "My Documents" folder. 

Standard logging will be used.
		 

### Example 4:
```powershell
.\Get-SqlServerInventoryToClixml.ps1 -Computername $env:COMPUTERNAME -IncludeDatabaseObjectInformation -LoggingPreference Verbose
```
Collect an inventory by scanning the local machine for SQL Server instances. 

Windows Authentication will be used to connect to the instance.

Database objects (EXCLUDING system objects) will be included in the results. 

The Inventory file will be written to your "My Documents" folder. 

Verbose logging will be used.

### Example 5:
```powershell
.\Get-SqlServerInventoryToClixml.ps1 -Computername $env:COMPUTERNAME -IncludeDatabaseObjectInformation -IncludeDatabaseSystemObjects
```
Collect an inventory by scanning the local machine for SQL Server instances. 

Windows Authentication will be used to connect to the instance. 

Database objects (INCLUDING system objects) will be included in the results.


## How Long Will It Take?
When run with the defaults (i.e. do not collect database object information or permissions) on a machine with 2 CPUs you can expect the script to take about 15-20 minutes to complete an inventory of 20 instances. Your mileage will vary depending on how many instances you are collecting information from and what additional information you're including.

A few suggestions\observations:

- Features have been added to each version of SQL Server so it stands to reason that the more recent your versions of SQL Server are, the longer it will take to collect information from them - especially if you include database object information and system objects as part of the inventory. 
- Progress is written to the PowerShell console to give you a better idea what the script's up to. If you've got logging enabled (highly recommended) you can also check the logs for progress updates. 
- Just like any other software it's possible to run into issues if you try and do too much. In other words, you may have more success getting the script to complete in a reasonable amount of time if you limit what it's doing. If you're collecting database object information and permissions, consider limiting the number of instances you include in the inventory. Do you really need to include ALL of development, QA, test, and production in a single inventory? 
- SQL Server has a LOT of system objects under the covers so unless you have a good reason to, don't bother collecting system object details...and if you need to, you may want to limit the inventory to a few machines at a time. 


## Generate A SQL Inventory Report
Once the inventory collection phase is complete you'll want to copy the output file to the machine where you'll create the inventory reports (Excel workbooks).

To create an inventory report, start by opening a PowerShell console and set your current location to the WindowsPowerShell folder:
	`Set-Location "$([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments))\WindowsPowerShell"`

This time you're going to execute the script `.\Convert-SqlServerInventoryClixmlToExcel.ps1` and supply the following parameters:

	-FromPath

	The literal path to the output file created by Get-SqlServerInventoryToClixml.ps1.

	-ToDirectoryPath
	Optional. Specifies the literal path to the directory where the Excel workbooks will be written. This path must exist prior to executing the script. If this parameter is not provided the workbooks will be written to the same directory specified in the FromPath parameter.


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

 

## Examples
The following examples demonstrate how to combine all the parameters together when running the script.

Example 1:
```powershell
.\Convert-SqlServerInventoryClixmlToExcel.ps1 -FromPath "C:\Inventory\SQL Server Inventory.xml.gz"
```

Writes Excel files for the Database Engine and Windows Operating System information contained in "C:\Inventory\SQL Server Inventory.xml.gz" to "C:\Inventory\SQL Server - Database Engine.xlsx" and "C:\Inventory\SQL Server - Windows.xlsx", respectively.
The Office color theme and Medium color scheme will be used by default.
	
Example 2:
```powershell
.\Convert-SqlServerInventoryClixmlToExcel.ps1 -FromPath "C:\Inventory\SQL Server Inventory.xml.gz"  -ColorTheme Blue -ColorScheme Dark
```

Writes Excel files for the Database Engine and Windows Operating System information contained in `"C:\Inventory\SQL Server Inventory.xml.gz"` to `"C:\Inventory\SQL Server - Database Engine.xlsx"` and `"C:\Inventory\SQL Server - Windows.xlsx"`, respectively.  
The Blue color theme and Dark color scheme will be used.

 
 
## Additional Help
If you're still having problems using SQL Power Doc after reading through this guide please post in the [Discussions](https://sqlpowerdoc.codeplex.com/discussions) or reach out to [@SQLDBA](https://twitter.com/sqldba) on Twitter.


## License
[MIT](/LICENSE)
