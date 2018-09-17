######################
# CONSTANTS
######################

# Registry constants
New-Variable -Name HKEY_CLASSES_ROOT -Value 2147483648 -Scope Script -Option Constant
New-Variable -Name HKEY_CURRENT_USER -Value 2147483649 -Scope Script -Option Constant
New-Variable -Name HKEY_LOCAL_MACHINE -Value 2147483650 -Scope Script -Option Constant
New-Variable -Name HKEY_USERS -Value 2147483651 -Scope Script -Option Constant
New-Variable -Name HKEY_CURRENT_CONFIG -Value 2147483653 -Scope Script -Option Constant
New-Variable -Name HKEY_DYN_DATA -Value 2147483654 -Scope Script -Option Constant

# Script constants
New-Variable -Name Version -Value '1.0.0' -Scope Script -Option Constant


# OS Versions
# See http://msdn.microsoft.com/en-us/library/ms724832(v=VS.85).aspx for version numbers
# and http://en.wikipedia.org/wiki/Timeline_of_Microsoft_Windows for version timeline

New-Object -TypeName System.Version -ArgumentList '5.0' | New-Variable -Name Windows2000 -Scope Script -Option Constant
New-Object -TypeName System.Version -ArgumentList '5.1' | New-Variable -Name WindowsXP -Scope Script -Option Constant
New-Object -TypeName System.Version -ArgumentList '5.2' | New-Variable -Name WindowsXP64bit -Scope Script -Option Constant
New-Object -TypeName System.Version -ArgumentList '5.2' | New-Variable -Name WindowsServer2003 -Scope Script -Option Constant
New-Object -TypeName System.Version -ArgumentList '5.2' | New-Variable -Name WindowsServer2003R2 -Scope Script -Option Constant
New-Object -TypeName System.Version -ArgumentList '6.0' | New-Variable -Name WindowsVista -Scope Script -Option Constant
New-Object -TypeName System.Version -ArgumentList '6.0' | New-Variable -Name WindowsServer2008 -Scope Script -Option Constant
New-Object -TypeName System.Version -ArgumentList '6.1' | New-Variable -Name WindowsServer2008R2 -Scope Script -Option Constant
New-Object -TypeName System.Version -ArgumentList '6.1' | New-Variable -Name Windows7 -Scope Script -Option Constant
New-Object -TypeName System.Version -ArgumentList '6.2' | New-Variable -Name Windows8 -Scope Script -Option Constant
New-Object -TypeName System.Version -ArgumentList '6.2' | New-Variable -Name WindowsServer2012 -Scope Script -Option Constant



######################
# LOOKUP FUNCTIONS
######################

function Get-BIOSCharacteristic($BIOSCharacteristic) {
	Write-Output $(
		switch ($BIOSCharacteristic) {
			0 {'Reserved'}
			1 {'Reserved'}
			2 {'Unknown'}
			3 {'BIOS Characteristics Not Supported'}
			4 {'ISA is supported'}
			5 {'MCA is supported'}
			6 {'EISA is supported'}
			7 {'PCI is supported'}
			8 {'PC Card (PCMCIA) is supported'}
			9 {'Plug and Play is supported'}
			10 {'APM is supported'}
			11 {'BIOS is Upgradable (Flash)'}
			12 {'BIOS shadowing is allowed'}
			13 {'VL-VESA is supported'}
			14 {'ESCD support is available'}
			15 {'Boot from CD is supported'}
			16 {'Selectable Boot is supported'}
			17 {'BIOS ROM is socketed'}
			18 {'Boot From PC Card (PCMCIA) is supported'}
			19 {'EDD (Enhanced Disk Drive) Specification is supported'}
			20 {'Int 13h - Japanese Floppy for NEC 9800 1.2mb (3.5, 1k Bytes/Sector, 360 RPM) is supported'}
			21 {'Int 13h - Japanese Floppy for Toshiba 1.2mb (3.5, 360 RPM) is supported'}
			22 {'Int 13h - 5.25 / 360 KB Floppy Services are supported'}
			23 {'Int 13h - 5.25 /1.2MB Floppy Services are supported'}
			24 {'13h - 3.5 / 720 KB Floppy Services are supported'}
			25 {'Int 13h - 3.5 / 2.88 MB Floppy Services are supported'}
			26 {'Int 5h, Print Screen Service is supported'}
			27 {'Int 9h, 8042 Keyboard services are supported'}
			28 {'Int 14h, Serial Services are supported'}
			29 {'Int 17h, printer services are supported'}
			30 {'Int 10h, CGA/Mono Video Services are supported'}
			31 {'NEC PC-98'}
			32 {'ACPI supported'}
			33 {'USB Legacy is supported'}
			34 {'AGP is supported'}
			35 {'I2O boot is supported'}
			36 {'LS-120 boot is supported'}
			37 {'ATAPI ZIP Drive boot is supported'}
			38 {'1394 boot is supported'}
			39 {'Smart Battery supported'}
			default {'Unknown (Undocumented)'} 
		}
	)
}

function Get-PhysicalMemoryFormFactor($PhysicalMemoryFormFactor) {
	Write-Output $(
		switch ($PhysicalMemoryFormFactor) {
			0 {'Unknown'}
			1 {'Other'}
			2 {'SIP'}
			3 {'DIP'}
			4 {'ZIP'}
			5 {'SOJ'}
			6 {'Proprietary'}
			7 {'SIMM'}
			8 {'DIMM'}
			9 {'TSOP'}
			10 {'PGA'}
			11 {'RIMM'}
			12 {'SODIMM'}
			13 {'SRIMM'}
			14 {'SMD'}
			15 {'SSMP'}
			16 {'QFP'}
			17 {'TQFP'}
			18 {'SOIC'}
			19 {'LCC'}
			20 {'PLCC'}
			21 {'BGA'}
			22 {'FPBGA'}
			23 {'LGA'}
			default {'Unknown (Undocumented)'} 
		}
	)
}

function Get-PhysicalMemoryType($PhysicalMemoryType) {
	Write-Output $(
		switch ($PhysicalMemoryType) {
			0 {'Unknown'}
			1 {'Other'}
			2 {'DRAM'}
			3 {'Synchronous DRAM'}
			4 {'Cache DRAM'}
			5 {'EDO'}
			6 {'EDRAM'}
			7 {'VRAM'}
			8 {'SRAM'}
			9 {'RAM'}
			10 {'ROM'}
			11 {'Flash'}
			12 {'EEPROM'}
			13 {'FEPROM'}
			14 {'EPROM'}
			15 {'CDRAM'}
			16 {'3DRAM'}
			17 {'SDRAM'}
			18 {'SGRAM'}
			19 {'RDRAM'}
			20 {'DDR'}
			21 {'DDR-2'}
			default {'Unknown (Undocumented)'} 
		}
	)
}

function Get-PhysicalMemoryTypeDetail($PhysicalMemoryTypeDetail) {
	Write-Output $(
		switch ($PhysicalMemoryTypeDetail) {
			1 {'Reserved'}
			2 {'Other'}
			4 {'Unknown'}
			8 {'Fast-paged'}
			16 {'Static column'}
			32 {'Pseudo-static'}
			64 {'RAMBUS'}
			128 {'Synchronous'}
			256 {'CMOS'}
			512 {'EDO'}
			1024 {'Window DRAM'}
			2048 {'Cache DRAM'}
			4096 {'Nonvolatile'}
			default {'Unknown (Undocumented)'} 
		}
	)
}

function Get-ProcessorFamily($ProcessorFamily) {
	Write-Output $(
		switch ($ProcessorFamily) {
			1 {'Other'}
			2 {'Unknown'}
			3 {'8086'}
			4 {'80286'}
			5 {'Intel386(TM) Processor'}
			6 {'Intel486(TM) Processor'}
			7 {'8087'}
			8 {'80287'}
			9 {'80387'}
			10 {'80487'}
			11 {'Pentium Brand'}
			12 {'Pentium Pro'}
			13 {'Pentium II'}
			14 {'Pentium Processor with MMX(TM) Technology'}
			15 {'Celeron(TM)'}
			16 {'Pentium II Xeon(TM)'}
			17 {'Pentium III'}
			18 {'M1 Family'}
			19 {'M2 Family'}
			24 {'AMD Duron(TM) Processor Family'}
			25 {'K5 Family'}
			26 {'K6 Family'}
			27 {'K6-2'}
			28 {'K6-3'}
			29 {'AMD Athlon(TM) Processor Family'}
			30 {'AMD2900 Family'}
			31 {'K6-2+'}
			32 {'Power PC Family'}
			33 {'Power PC 601'}
			34 {'Power PC 603'}
			35 {'Power PC 603+'}
			36 {'Power PC 604'}
			37 {'Power PC 620'}
			38 {'Power PC X704'}
			39 {'Power PC 750'}
			48 {'Alpha Family'}
			49 {'Alpha 21064'}
			50 {'Alpha 21066'}
			51 {'Alpha 21164'}
			52 {'Alpha 21164PC'}
			53 {'Alpha 21164a'}
			54 {'Alpha 21264'}
			55 {'Alpha 21364'}
			64 {'MIPS Family'}
			65 {'MIPS R4000'}
			66 {'MIPS R4200'}
			67 {'MIPS R4400'}
			68 {'MIPS R4600'}
			69 {'MIPS R10000'}
			80 {'SPARC Family'}
			81 {'SuperSPARC'}
			82 {'microSPARC II'}
			83 {'microSPARC IIep'}
			84 {'UltraSPARC'}
			85 {'UltraSPARC II'}
			86 {'UltraSPARC IIi'}
			87 {'UltraSPARC III'}
			88 {'UltraSPARC IIIi'}
			96 {'68040'}
			97 {'68xxx Family'}
			98 {'68000'}
			99 {'68010'}
			100 {'68020'}
			101 {'68030'}
			112 {'Hobbit Family'}
			120 {'Crusoe(TM) TM5000 Family'}
			121 {'Crusoe(TM) TM3000 Family'}
			122 {'Efficeon(TM) TM8000 Family'}
			128 {'Weitek'}
			130 {'Itanium(TM) Processor'}
			131 {'AMD Athlon(TM) 64 Processor Famiily'}
			132 {'AMD Opteron(TM) Processor Family'}
			144 {'PA-RISC Family'}
			145 {'PA-RISC 8500'}
			146 {'PA-RISC 8000'}
			147 {'PA-RISC 7300LC'}
			148 {'PA-RISC 7200'}
			149 {'PA-RISC 7100LC'}
			150 {'PA-RISC 7100'}
			160 {'V30 Family'}
			176 {'Pentium III Xeon(TM) Processor'}
			177 {'Pentium III Processor with Intel SpeedStep(TM) Technology'}
			178 {'Pentium 4'}
			179 {'Intel Xeon(TM)'}
			180 {'AS400 Family'}
			181 {'Intel Xeon(TM) Processor MP'}
			182 {'AMD Athlon(TM) XP Family'}
			183 {'AMD Athlon(TM) MP Family'}
			184 {'Intel Itanium 2'}
			185 {'Intel Pentium M Processor'}
			190 {'K7'}
			200 {'IBM390 Family'}
			201 {'G4'}
			202 {'G5'}
			203 {'G6'}
			204 {'z/Architecture Base'}
			250 {'i860'}
			251 {'i960'}
			260 {'SH-3'}
			261 {'SH-4'}
			280 {'ARM'}
			281 {'StrongARM'}
			300 {'6x86'}
			301 {'MediaGX'}
			302 {'MII'}
			320 {'WinChip'}
			350 {'DSP'}
			500 {'Video Processor'}
			default {'Unknown (Undocumented)'} 
		}
	)
}

function Get-ShareType($ShareType) {
	Write-Output $(
		switch ($ShareType) {
			0 {'Disk Drive'}
			1 {'Print Queue'}
			2 {'Device'}
			3 {'IPC'}
			2147483648 {'Disk Drive Admin'}
			2147483649 {'Print Queue Admin'}
			2147483650 {'Device Admin'}
			2147483651 {'IPC Admin'}
			default {'Unknown (Undocumented)'} 
		}
	)
}

function Get-ChassisType($ChassisType) {
	Write-Output $(
		switch ($ChassisType) {
			1 {'Other'}
			2 {'Unknown'}
			3 {'Desktop'}
			4 {'Low Profile Desktop'}
			5 {'Pizza Box'}
			6 {'Mini Tower'}
			7 {'Tower'}
			8 {'Portable'}
			9 {'Laptop'}
			10 {'Notebook'}
			11 {'Hand Held'}
			12 {'Docking Station'}
			13 {'All in One'}
			14 {'Sub Notebook'}
			15 {'Space-Saving'}
			16 {'Lunch Box'}
			17 {'Main System Chassis'}
			18 {'Expansion Chassis'}
			19 {'SubChassis'}
			20 {'Bus Expansion Chassis'}
			21 {'Peripheral Chassis'}
			22 {'Storage Chassis'}
			23 {'Rack Mount Chassis'}
			24 {'Sealed-Case PC'}
			default {'Unknown (Undocumented)'} 
		}
	)
}


#function Get-Windows2003ComponentInformation($ComponentName) {
#	$ComponentInformation = switch ($ComponentName.ToLower()) {
#		'accessopt'	{ 'Accessibility Wizard' }
#	}
#	
#}


#function Get-Windows2000ComponentInformation($ComponentName) {
#	$ComponentInformation = switch ($ComponentName.ToLower()) {	
#	}
#}


function Get-ComponentInformation($ComponentName) {
	Write-Output $(
		switch ($ComponentName.ToLower()) {
			'accessopt' { @{ 
					'Class' = '0100' ; 
					'ClassName' = 'Accessories' ; 
					'DisplayName' = 'Accessibility Wizard' ; 
					'Level' = 2 ; 
				}
			}
			'cdplayer' { @{ 
					'Class' = '0100' ; 
					'ClassName' = 'Accessories' ; 
					'DisplayName' = 'CD Player' ; 
					'Level' = $null ; 
				}
			}
			'calc' { @{ 
					'Class' = '0100' ; 
					'ClassName' = 'Accessories' ; 
					'DisplayName' = 'Calculator' ; 
					'Level' = $null ; 
				}
			}
			'charmap' { @{ 
					'Class' = '0100' ; 
					'ClassName' = 'Accessories' ; 
					'DisplayName' = 'Character Map' ; 
					'Level' = $null ; 
				}
			}
			'clipbook' { @{ 
					'Class' = '0100' ; 
					'ClassName' = 'Accessories' ; 
					'DisplayName' = 'Clipboard Viewer' ; 
					'Level' = $null ; 
				}
			}
			'deskpaper' { @{ 
					'Class' = '0100' ; 
					'ClassName' = 'Accessories' ; 
					'DisplayName' = 'Desktop Wallpaper' ; 
					'Level' = $null ; 
				}
			}
			'imagevue' { @{ 
					'Class' = '0100' ; 
					'ClassName' = 'Accessories' ; 
					'DisplayName' = 'Imaging' ; 
					'Level' = 2 ; 
				}
			}
			'mousepoint' { @{ 
					'Class' = '0100' ; 
					'ClassName' = 'Accessories' ; 
					'DisplayName' = 'Mouse Pointers' ; 
					'Level' = 2 ; 
				}
			}
			'mswordpad' { @{ 
					'Class' = '0100' ; 
					'ClassName' = 'Accessories' ; 
					'DisplayName' = 'Wordpad' ; 
					'Level' = 2 ; 
				}
			}
			'objectpkg' { @{ 
					'Class' = '0100' ; 
					'ClassName' = 'Accessories' ; 
					'DisplayName' = 'Object Packager' ; 
					'Level' = 2 ; 
				}
			}
			'paint' { @{ 
					'Class' = '0100' ; 
					'ClassName' = 'Accessories' ; 
					'DisplayName' = 'Paint' ; 
					'Level' = 2 ; 
				}
			}
			'templates' { @{ 
					'Class' = '0100' ; 
					'ClassName' = 'Accessories' ; 
					'DisplayName' = 'Document Templates' ; 
					'Level' = 2 ; 
				}
			}
			'chat' { @{ 
					'Class' = '0140' ; 
					'ClassName' = 'Communication' ; 
					'DisplayName' = 'Chat' ; 
					'Level' = 2 ; 
				}
			}
			'dialer' { @{ 
					'Class' = '0140' ; 
					'ClassName' = 'Communication' ; 
					'DisplayName' = 'Phone Dialer' ; 
					'Level' = 2 ; 
				}
			}
			'hypertrm' { @{ 
					'Class' = '0140' ; 
					'ClassName' = 'Communication' ; 
					'DisplayName' = 'HyperTerminal' ; 
					'Level' = 2 ; 
				}
			}
			'freecell' { @{ 
					'Class' = '0150' ; 
					'ClassName' = 'Games' ; 
					'DisplayName' = '' ; 
					'Level' = 2 ; 
				}
			}
			'hearts' { @{ 
					'Class' = '0150' ; 
					'ClassName' = 'Games' ; 
					'DisplayName' = 'Hearts' ; 
					'Level' = 2 ; 
				}
			}
			'freecell' { @{ 
					'Class' = '0150' ; 
					'ClassName' = 'Games' ; 
					'DisplayName' = '' ; 
					'Level' = 2 ; 
				}
			}
			'minesweeper' { @{ 
					'Class' = '0150' ; 
					'ClassName' = 'Games' ; 
					'DisplayName' = 'Minesweeper' ; 
					'Level' = 2 ; 
				}
			}
			'solitaire' { @{ 
					'Class' = '0150' ; 
					'ClassName' = 'Games' ; 
					'DisplayName' = 'Solitaire' ; 
					'Level' = 2 ; 
				}
			}
			'spider' { @{ 
					'Class' = '0150' ; 
					'ClassName' = 'Games' ; 
					'DisplayName' = 'Spider Solitaire' ; 
					'Level' = 2 ; 
				}
			}
			'pinball' { @{ 
					'Class' = '0150' ; 
					'ClassName' = 'Games' ; 
					'DisplayName' = 'Pinball' ; 
					'Level' = 2 ; 
				}
			}
			'zonegames' { @{ 
					'Class' = '0150' ; 
					'ClassName' = 'Games' ; 
					'DisplayName' = 'Internet Games' ; 
					'Level' = 2 ; 
				}
			}
			'media_clips' { @{ 
					'Class' = '0160' ; 
					'ClassName' = 'Multimedia' ; 
					'DisplayName' = 'Sample Sounds' ; 
					'Level' = 2 ; 
				}
			}
			'com' { @{ 
					'Class' = '0170' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'COM+' ; 
					'Level' = 1 ; 
				}
			}
			'dtc' { @{ 
					'Class' = '0180' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Distributed Transaction Coordinator' ; 
					'Level' = 1 ; 
				}
			}
			'media_utopia' { @{ 
					'Class' = '0160' ; 
					'ClassName' = 'Multimedia' ; 
					'DisplayName' = 'Utopia Sound Scheme' ; 
					'Level' = 2 ; 
				}
			}
			'mplay' { @{ 
					'Class' = '0160' ; 
					'ClassName' = 'Multimedia' ; 
					'DisplayName' = 'Media Player' ; 
					'Level' = 2 ; 
				}
			}
			'rec' { @{ 
					'Class' = '0160' ; 
					'ClassName' = 'Multimedia' ; 
					'DisplayName' = 'Sound Recorder' ; 
					'Level' = 2 ; 
				}
			}
			'vol' { @{ 
					'Class' = '0160' ; 
					'ClassName' = 'Multimedia' ; 
					'DisplayName' = 'Volume Control' ; 
					'Level' = 2 ; 
				}
			}
			'adam' { @{ 
					'Class' = '0190' ; 
					'ClassName' = 'Active Directory Services' ; 
					'DisplayName' = 'Active Directory Application Mode (ADAM)' ; 
					'Level' = 2 ; 
				}
			}
			'adfs' { @{ 
					'Class' = '0190' ; 
					'ClassName' = 'Active Directory Services' ; 
					'DisplayName' = 'Active Directory Federation Services (ADFS)' ; 
					'Level' = 2 ; 
				}
			}
			'appsrv_console' { @{ 
					'Class' = '0200' ; 
					'ClassName' = 'Application Server' ; 
					'DisplayName' = 'Application Server Console' ; 
					'Level' = 2 ; 
				}
			}
			'aspnet' { @{ 
					'Class' = '0200' ; 
					'ClassName' = 'Application Server' ; 
					'DisplayName' = 'ASP.NET' ; 
					'Level' = 2 ; 
				}
			}
			'dtcnetwork' { @{ 
					'Class' = '0200' ; 
					'ClassName' = 'Enable Network DTC Access' ; 
					'DisplayName' = '' ; 
					'Level' = 2 ; 
				}
			}
			'complusnetwork' { @{ 
					'Class' = '0200' ; 
					'ClassName' = 'Enable Network COM+ Access' ; 
					'DisplayName' = '' ; 
					'Level' = 2 ; 
				}
			}
			'bitsserverextensionsisapi' { @{ 
					'Class' = '0210' ; 
					'ClassName' = 'Internet Information Server (IIS)' ; 
					'DisplayName' = 'BITS Server Extention ISAPI' ; 
					'Level' = 2 ; 
				}
			}
			'bitsserverextensionsmanager' { @{ 
					'Class' = '0210' ; 
					'ClassName' = 'Internet Information Server (IIS)' ; 
					'DisplayName' = 'BITS Management Console Snap-In' ; 
					'Level' = 2 ; 
				}
			}
			'fp_extensions' { @{ 
					'Class' = '0210' ; 
					'ClassName' = 'Internet Information Server (IIS)' ; 
					'DisplayName' = 'Frontpage Server Extensions' ; 
					'Level' = 2 ; 
				}
			}
			'iis_common' { @{ 
					'Class' = '0210' ; 
					'ClassName' = 'Internet Information Server (IIS)' ; 
					'DisplayName' = 'Common Files' ; 
					'Level' = 2 ; 
				}
			}
			'iis_doc' { @{ 
					'Class' = '0210' ; 
					'ClassName' = 'Internet Information Server (IIS)' ; 
					'DisplayName' = 'Documentation' ; 
					'Level' = 2 ; 
				}
			}
			'iis_ftp' { @{ 
					'Class' = '0210' ; 
					'ClassName' = 'Internet Information Server (IIS)' ; 
					'DisplayName' = 'FTP Server' ; 
					'Level' = 2 ; 
				}
			}
			'iis_htmla' { @{ 
					'Class' = '0210' ; 
					'ClassName' = 'Internet Information Server (IIS)' ; 
					'DisplayName' = 'IIS Manager (HTML)' ; 
					'Level' = 2 ; 
				}
			}
			'iis_inetmgr' { @{ 
					'Class' = '0210' ; 
					'ClassName' = 'Internet Information Server (IIS)' ; 
					'DisplayName' = 'Internet Information Services Manager' ; 
					'Level' = 2 ; 
				}
			}
			'iis_nntp' { @{ 
					'Class' = '0210' ; 
					'ClassName' = 'Internet Information Server (IIS)' ; 
					'DisplayName' = 'NNTP Service' ; 
					'Level' = 2 ; 
				}
			}
			'iis_pwmgr' { @{ 
					'Class' = '0210' ; 
					'ClassName' = 'Internet Information Server (IIS)' ; 
					'DisplayName' = 'Personal Web Manager' ; 
					'Level' = 2 ; 
				}
			}
			'iis_smtp' { @{ 
					'Class' = '0210' ; 
					'ClassName' = 'Internet Information Server (IIS)' ; 
					'DisplayName' = 'SMTP Service' ; 
					'Level' = 2 ; 
				}
			}
			'iis_www' { @{ 
					'Class' = '0210' ; 
					'ClassName' = 'Internet Information Server (IIS)' ; 
					'DisplayName' = 'World Wide Web Server' ; 
					'Level' = 2 ; 
				}
			}
			'inetprint' { @{ 
					'Class' = '0210' ; 
					'ClassName' = 'Internet Information Server (IIS)' ; 
					'DisplayName' = 'Internet Printing' ; 
					'Level' = 2 ; 
				}
			}
			'iis_asp' { @{ 
					'Class' = '0215' ; 
					'ClassName' = 'World Wide Web Server' ; 
					'DisplayName' = 'Active Server Pages' ; 
					'Level' = 3 ; 
				}
			}
			'iis_internetdataconnector' { @{ 
					'Class' = '0215' ; 
					'ClassName' = 'World Wide Web Server' ; 
					'DisplayName' = 'Internet Data Connector' ; 
					'Level' = 3 ; 
				}
			}
			'iis_serversideincludes' { @{ 
					'Class' = '0215' ; 
					'ClassName' = 'World Wide Web Server' ; 
					'DisplayName' = 'Server Side Includes' ; 
					'Level' = 3 ; 
				}
			}
			'iis_webdav' { @{ 
					'Class' = '0215' ; 
					'ClassName' = 'World Wide Web Server' ; 
					'DisplayName' = 'WebDAV Publishing' ; 
					'Level' = 3 ; 
				}
			}
			'iis_www_vdir_printers' { @{ 
					'Class' = '0215' ; 
					'ClassName' = 'World Wide Web Server' ; 
					'DisplayName' = 'Printers Virtual Directory' ; 
					'Level' = 3 ; 
				}
			}
			'iis_www_vdir_scripts' { @{ 
					'Class' = '0215' ; 
					'ClassName' = 'World Wide Web Server' ; 
					'DisplayName' = 'Scripts Virtual Directory' ; 
					'Level' = 3 ; 
				}
			}
			'sakit_web' { @{ 
					'Class' = '0215' ; 
					'ClassName' = 'World Wide Web Server' ; 
					'DisplayName' = 'Remote Administration (HTML)' ; 
					'Level' = 3 ; 
				}
			}
			'tswebclient' { @{ 
					'Class' = '0215' ; 
					'ClassName' = 'World Wide Web Server' ; 
					'DisplayName' = 'Remote Desktop Web Connection' ; 
					'Level' = 3 ; 
				}
			}
			'certsrv_client' { @{ 
					'Class' = '0300' ; 
					'ClassName' = 'Certificate Services' ; 
					'DisplayName' = 'Certificate Services Web Enrollment Support' ; 
					'Level' = 2 ; 
				}
			}
			'certsrv_server' { @{ 
					'Class' = '0300' ; 
					'ClassName' = 'Certificate Services' ; 
					'DisplayName' = 'Certificate Services CA' ; 
					'Level' = 2 ; 
				}
			}
			'dfsext' { @{ 
					'Class' = '0305' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'DFS Extentions Library' ; 
					'Level' = 1 ; 
				}
			}
			'cluster' { @{ 
					'Class' = '0350' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Cluster Service' ; 
					'Level' = 1 ; 
				}
			}
			'pop3admin' { @{ 
					'Class' = '0400' ; 
					'ClassName' = 'E-Mail Services' ; 
					'DisplayName' = 'POP3 Service Web Administration' ; 
					'Level' = 2 ; 
				}
			}
			'pop3service' { @{ 
					'Class' = '0400' ; 
					'ClassName' = 'E-Mail Services' ; 
					'DisplayName' = 'POP3 Service' ; 
					'Level' = 2 ; 
				}
			}
			'fax' { @{ 
					'Class' = '0500' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Fax Services' ; 
					'Level' = 1 ; 
				}
			}
			'indexsrv_system' { @{ 
					'Class' = '0600' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Indexing Service' ; 
					'Level' = 1 ; 
				}
			}
			'iehardenadmin' { @{ 
					'Class' = '0700' ; 
					'ClassName' = 'Internet Explorer Enhanced Security Configuration' ; 
					'DisplayName' = 'For administrator groups' ; 
					'Level' = 2 ; 
				}
			}
			'iehardenuser' { @{ 
					'Class' = '0700' ; 
					'ClassName' = 'Internet Explorer Enhanced Security Configuration' ; 
					'DisplayName' = 'For all other user groups' ; 
					'Level' = 2 ; 
				}
			}
			'ieaccess' { @{ 
					'Class' = '0800' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Internet Explorer (from Start Menu and Desktop)' ; 
					'Level' = 1 ; 
				}
			}
			'netcm' { @{ 
					'Class' = '0900' ; 
					'ClassName' = 'Management and Monitoring Tools' ; 
					'DisplayName' = 'Connection Manager Components' ; 
					'Level' = 2 ; 
				}
			}
			'netcmak' { @{ 
					'Class' = '0900' ; 
					'ClassName' = 'Management and Monitoring Tools' ; 
					'DisplayName' = 'Connection Manager Administration Kit' ; 
					'Level' = 2 ; 
				}
			}
			'netcps' { @{ 
					'Class' = '0900' ; 
					'ClassName' = 'Management and Monitoring Tools' ; 
					'DisplayName' = 'Connection Point Services' ; 
					'Level' = 2 ; 
				}
			}
			'fsrstandard' { @{ 
					'Class' = '0900' ; 
					'ClassName' = 'Management and Monitoring Tools' ; 
					'DisplayName' = 'File Server Management' ; 
					'Level' = 2 ; 
				}
			}
			'srm' { @{ 
					'Class' = '0900' ; 
					'ClassName' = 'Management and Monitoring Tools' ; 
					'DisplayName' = 'File Server Resource Management' ; 
					'Level' = 2 ; 
				}
			}
			'hwmgmt' { @{ 
					'Class' = '0900' ; 
					'ClassName' = 'Management and Monitoring Tools' ; 
					'DisplayName' = 'Hardware Management' ; 
					'Level' = 2 ; 
				}
			}
			'netmontools' { @{ 
					'Class' = '0900' ; 
					'ClassName' = 'Management and Monitoring Tools' ; 
					'DisplayName' = 'Network Monitor Tools' ; 
					'Level' = 2 ; 
				}
			}
			'pmcsnap' { @{ 
					'Class' = '0900' ; 
					'ClassName' = 'Management and Monitoring Tools' ; 
					'DisplayName' = 'Print Management Component' ; 
					'Level' = 2 ; 
				}
			}
			'sanmgmt' { @{ 
					'Class' = '0900' ; 
					'ClassName' = 'Management and Monitoring Tools' ; 
					'DisplayName' = 'Storage Manager for SANs' ; 
					'Level' = 2 ; 
				}
			}
			'snmp' { @{ 
					'Class' = '0900' ; 
					'ClassName' = 'Management and Monitoring Tools' ; 
					'DisplayName' = 'Simple Network Management Protocol' ; 
					'Level' = 2 ; 
				}
			}
			'wbemmsi' { @{ 
					'Class' = '0900' ; 
					'ClassName' = 'Management and Monitoring Tools' ; 
					'DisplayName' = 'WMI Windows Installer Provider' ; 
					'Level' = 2 ; 
				}
			}
			'wbemsnmp' { @{ 
					'Class' = '0900' ; 
					'ClassName' = 'Management and Monitoring Tools' ; 
					'DisplayName' = 'WMI SNMP Provider' ; 
					'Level' = 2 ; 
				}
			}
			'freestyle' { @{ 
					'Class' = '0910' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Media Center' ; 
					'Level' = 1 ; 
				}
			}
			'netfx20' { @{ 
					'Class' = '0920' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Microsoft .NET Framework 2.0' ; 
					'Level' = 1 ; 
				}
			}
			'msmq' { @{ 
					'Class' = '1000' ; 
					'ClassName' = 'Message Queuing' ; 
					'DisplayName' = 'Message Queuing Services' ; 
					'Level' = 2 ; 
				}
			}
			'msmq_adintegrated' { @{ 
					'Class' = '1000' ; 
					'ClassName' = 'Message Queuing' ; 
					'DisplayName' = 'Active Directory Integration' ; 
					'Level' = 2 ; 
				}
			}
			'msmq_core' { @{ 
					'Class' = '1000' ; 
					'ClassName' = 'Message Queuing' ; 
					'DisplayName' = 'Common (Core Functionallity)' ; 
					'Level' = 2 ; 
				}
			}
			'msmq_httpsupport' { @{ 
					'Class' = '1000' ; 
					'ClassName' = 'Message Queuing' ; 
					'DisplayName' = 'MSMQ HTTP Support' ; 
					'Level' = 2 ; 
				}
			}
			'msmq_localstorage' { @{ 
					'Class' = '1000' ; 
					'ClassName' = 'Message Queuing' ; 
					'DisplayName' = 'Common (Local Storage)' ; 
					'Level' = 2 ; 
				}
			}
			'msmq_mqdsservice' { @{ 
					'Class' = '1000' ; 
					'ClassName' = 'Message Queuing' ; 
					'DisplayName' = 'Downlevel Client Support' ; 
					'Level' = 2 ; 
				}
			}
			'msmq_routingsupport' { @{ 
					'Class' = '1000' ; 
					'ClassName' = 'Message Queuing' ; 
					'DisplayName' = 'Routing Support' ; 
					'Level' = 2 ; 
				}
			}
			'msmq_triggersservice' { @{ 
					'Class' = '1000' ; 
					'ClassName' = 'Message Queuing' ; 
					'DisplayName' = 'Triggers' ; 
					'Level' = 2 ; 
				}
			}
			'msnexplr' { @{ 
					'Class' = '1100' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'MSN Explorer' ; 
					'Level' = 2 ; 
				}
			}
			'computeserver' { @{ 
					'Class' = '1150' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Microsoft Windows Compute Server' ; 
					'Level' = 1 ; 
				}
			}
			'storageserver' { @{ 
					'Class' = '1160' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Microsoft Windows Storage Server' ; 
					'Level' = 1 ; 
				}
			}
			'acs' { @{ 
					'Class' = '1200' ; 
					'ClassName' = 'Networking Services' ; 
					'DisplayName' = 'QoS Admission Control Service' ; 
					'Level' = 2 ; 
				}
			}
			'beacon' { @{ 
					'Class' = '1200' ; 
					'ClassName' = 'Networking Services' ; 
					'DisplayName' = 'Internet Gateway Device Discovery and Control Client' ; 
					'Level' = 2 ; 
				}
			}
			'dhcpserver' { @{ 
					'Class' = '1200' ; 
					'ClassName' = 'Networking Services' ; 
					'DisplayName' = 'Dynamic Host Configuration Protocol (DHCP)' ; 
					'Level' = 2 ; 
				}
			}
			'dns' { @{ 
					'Class' = '1200' ; 
					'ClassName' = 'Networking Services' ; 
					'DisplayName' = 'Domain Name System (DNS)' ; 
					'Level' = 2 ; 
				}
			}
			'ias' { @{ 
					'Class' = '1200' ; 
					'ClassName' = 'Networking Services' ; 
					'DisplayName' = 'Internet Authentication Service' ; 
					'Level' = 2 ; 
				}
			}
			'netrqs' { @{ 
					'Class' = '1200' ; 
					'ClassName' = 'Networking Services' ; 
					'DisplayName' = 'Remote Access Quarantine Service' ; 
					'Level' = 2 ; 
				}
			}
			'iprip' { @{ 
					'Class' = '1200' ; 
					'ClassName' = 'Networking Services' ; 
					'DisplayName' = 'RIP Listener' ; 
					'Level' = 2 ; 
				}
			}
			'netcis' { @{ 
					'Class' = '1200' ; 
					'ClassName' = 'Networking Services' ; 
					'DisplayName' = 'RPC Over HTTP Proxy' ; 
					'Level' = 2 ; 
				}
			}
			'p2p' { @{ 
					'Class' = '1200' ; 
					'ClassName' = 'Networking Services' ; 
					'DisplayName' = 'Peer-to-Peer' ; 
					'Level' = 2 ; 
				}
			}
			'upnp' { @{ 
					'Class' = '1200' ; 
					'ClassName' = 'Networking Services' ; 
					'DisplayName' = 'Universal Plug and Play' ; 
					'Level' = 2 ; 
				}
			}
			'wins' { @{ 
					'Class' = '1200' ; 
					'ClassName' = 'Networking Services' ; 
					'DisplayName' = 'Windows Internet Name Service (WINS)' ; 
					'Level' = 2 ; 
				}
			}
			'simptcp' { @{ 
					'Class' = '1200' ; 
					'ClassName' = 'Networking Services' ; 
					'DisplayName' = 'Simple TCP/IP Services' ; 
					'Level' = 2 ; 
				}
			}
			'ils' { @{ 
					'Class' = '1200' ; 
					'ClassName' = 'Networking Services' ; 
					'DisplayName' = 'Site Server ILS Services' ; 
					'Level' = 2 ; 
				}
			}
			'lpdsvc' { @{ 
					'Class' = '1300' ; 
					'ClassName' = 'Other Network File and Print Services' ; 
					'DisplayName' = 'Print Services for Unix' ; 
					'Level' = 2 ; 
				}
			}
			'macprint' { @{ 
					'Class' = '1300' ; 
					'ClassName' = 'Other Network File and Print Services' ; 
					'DisplayName' = 'Print Services for Macintosh' ; 
					'Level' = 2 ; 
				}
			}
			'macsrv' { @{ 
					'Class' = '1300' ; 
					'ClassName' = 'Other Network File and Print Services' ; 
					'DisplayName' = 'File Services for Macintosh' ; 
					'Level' = 2 ; 
				}
			}
			'reminst' { @{ 
					'Class' = '1400' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Remote Installation Services' ; 
					'Level' = 1 ; 
				}
			}
			'scw' { @{ 
					'Class' = '1450' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Security Configuration Wizard' ; 
					'Level' = 1 ; 
				}
			}
			'sua' { @{ 
					'Class' = '1460' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Subsystem for UNIX-based Applications' ; 
					'Level' = 1 ; 
				}
			}
			'rstorage' { @{ 
					'Class' = '1500' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Remote Storage' ; 
					'Level' = 1 ; 
				}
			}
			'iisdbg' { @{ 
					'Class' = '1600' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Script Debugger' ; 
					'Level' = 1 ; 
				}
			}
			'tsclients' { @{ 
					'Class' = '1700' ; 
					'ClassName' = 'Terminal Services' ; 
					'DisplayName' = 'Client Creator Files' ; 
					'Level' = 2 ; 
				}
			}
			'tsenable' { @{ 
					'Class' = '1700' ; 
					'ClassName' = 'Terminal Services' ; 
					'DisplayName' = 'Enable Terminal Services' ; 
					'Level' = 2 ; 
				}
			}
			'licenseserver' { @{ 
					'Class' = '1800' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Terminal Services Licensing' ; 
					'Level' = 1 ; 
				}
			}
			'uddiadmin' { @{ 
					'Class' = '1900' ; 
					'ClassName' = 'UDDI Services' ; 
					'DisplayName' = 'UDDI Services Administration Console' ; 
					'Level' = 2 ; 
				}
			}
			'uddidatabase' { @{ 
					'Class' = '1900' ; 
					'ClassName' = 'UDDI Services' ; 
					'DisplayName' = 'UDDI Services Database Components' ; 
					'Level' = 2 ; 
				}
			}
			'uddiweb' { @{ 
					'Class' = '1900' ; 
					'ClassName' = 'UDDI Services' ; 
					'DisplayName' = 'UDDI Services Web Server Components' ; 
					'Level' = 2 ; 
				}
			}
			'oeaccess' { @{ 
					'Class' = '2000' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Outlook Express (on start menu)' ; 
					'Level' = 1 ; 
				}
			}
			'terminalserver' { @{ 
					'Class' = '2050' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Terminal Server' ; 
					'Level' = 1 ; 
				}
			}
			'rootautoupdate' { @{ 
					'Class' = '2100' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Update Root Certificates' ; 
					'Level' = 1 ; 
				}
			}
			'autoupdate' { @{ 
					'Class' = '2150' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Automatic Updates' ; 
					'Level' = 1 ; 
				}
			}
			'wmpocm' { @{ 
					'Class' = '2200' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Windows Media Player (from Start Menu and Desktop)' ; 
					'Level' = 1 ; 
				}
			}
			'msmsgs' { @{ 
					'Class' = '2300' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Windows Messenger' ; 
					'Level' = 1 ; 
				}
			}
			'wmaccess' { @{ 
					'Class' = '2305' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Windows Messenger (from Start Menu)' ; 
					'Level' = 1 ; 
				}
			}
			'wms_admin' { @{ 
					'Class' = '2400' ; 
					'ClassName' = 'Windows Media Services' ; 
					'DisplayName' = 'Windows Media Services Admin' ; 
					'Level' = 2 ; 
				}
			}
			'wms_admin_asp' { @{ 
					'Class' = '2400' ; 
					'ClassName' = 'Windows Media Services' ; 
					'DisplayName' = 'Windows Media Services Administrator for the Web' ; 
					'Level' = 2 ; 
				}
			}
			'wms_admin_mmc' { @{ 
					'Class' = '2400' ; 
					'ClassName' = 'Windows Media Services' ; 
					'DisplayName' = 'Windows Media Services snap-in' ; 
					'Level' = 2 ; 
				}
			}
			'wms_isapi' { @{ 
					'Class' = '2400' ; 
					'ClassName' = 'Windows Media Services' ; 
					'DisplayName' = 'Multicast and Advertisement Logging Agent' ; 
					'Level' = 2 ; 
				}
			}
			'wms_server' { @{ 
					'Class' = '2400' ; 
					'ClassName' = 'Windows Media Services' ; 
					'DisplayName' = 'Windows Media Services' ; 
					'Level' = 2 ; 
				}
			}
			'wbem' { @{ 
					'Class' = '2500' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'WMI' ; 
					'Level' = 1 ; 
				}
			}
			'sharepoint' { @{ 
					'Class' = '2600' ; 
					'ClassName' = $null ; 
					'DisplayName' = 'Windows Sharepoint Services' ; 
					'Level' = 1 ; 
				}
			}
			'authman' { @{ 
					'Class' = '9999' ; 
					'ClassName' = 'Hidden' ; 
					'DisplayName' = 'authman' ; 
					'Level' = 1 ; 
				}
			}
			'cfscommonuifx' { @{ 
					'Class' = '9999' ; 
					'ClassName' = 'Hidden' ; 
					'DisplayName' = 'cfscommonuifx' ; 
					'Level' = 1 ; 
				}
			}
			'display' { @{ 
					'Class' = '9999' ; 
					'ClassName' = 'Hidden' ; 
					'DisplayName' = 'display' ; 
					'Level' = 1 ; 
				}
			}
			'fp_vid_deploy' { @{ 
					'Class' = '9999' ; 
					'ClassName' = 'Hidden' ; 
					'DisplayName' = 'fp_vid_deploy' ; 
					'Level' = 1 ; 
				}
			}
			'fsrcommon' { @{ 
					'Class' = '9999' ; 
					'ClassName' = 'Hidden' ; 
					'DisplayName' = 'fsrcommon' ; 
					'Level' = 1 ; 
				}
			}
			'netfx' { @{ 
					'Class' = '9999' ; 
					'ClassName' = 'Hidden' ; 
					'DisplayName' = 'netfx' ; 
					'Level' = 1 ; 
				}
			}
			'notebook' { @{ 
					'Class' = '9999' ; 
					'ClassName' = 'Hidden' ; 
					'DisplayName' = 'notebook' ; 
					'Level' = 1 ; 
				}
			}
			'ntcomponents' { @{ 
					'Class' = '9999' ; 
					'ClassName' = 'Hidden' ; 
					'DisplayName' = 'ntcomponents' ; 
					'Level' = 1 ; 
				}
			}
			'oobe' { @{ 
					'Class' = '9999' ; 
					'ClassName' = 'Hidden' ; 
					'DisplayName' = 'oobe' ; 
					'Level' = 1 ; 
				}
			}
			'starter' { @{ 
					'Class' = '9999' ; 
					'ClassName' = 'Hidden' ; 
					'DisplayName' = 'starter' ; 
					'Level' = 1 ; 
				}
			}
			'stickynotes' { @{ 
					'Class' = '9999' ; 
					'ClassName' = 'Hidden' ; 
					'DisplayName' = 'stickynotes' ; 
					'Level' = 1 ; 
				}
			}
			'system' { @{ 
					'Class' = '9999' ; 
					'ClassName' = 'Hidden' ; 
					'DisplayName' = 'system' ; 
					'Level' = 1 ; 
				}
			}
			'tpg' { @{ 
					'Class' = '9999' ; 
					'ClassName' = 'Hidden' ; 
					'DisplayName' = 'tpg' ; 
					'Level' = 1 ; 
				}
			}
			'wms_svrtyplib' { @{ 
					'Class' = '9999' ; 
					'ClassName' = 'Hidden' ; 
					'DisplayName' = 'wms_svrtyplib' ; 
					'Level' = 1 ; 
				}
			}
			default { @{ 
					'Class' = '3000' ; 
					'ClassName' = 'Hidden' ; 
					'DisplayName' = '* Unknown ($($ComponentName.ToLower()))' ; 
					'Level' = 1 ; 
				}
			}
			#		'Name' { @{ 
			#				'Class' = '' ; 
			#				'ClassName' = '' ; 
			#				'DisplayName' = '' ; 
			#				'Level' = 1 ; 
			#			}
			#		}
		}
	)

	#	# Update $script:SystemRole
	#	switch ($ComponentName.ToLower()) {
	#		'iis_ftp' { $script:SystemRole['FTP'] = $true }
	#		'iis_nntp' { $script:SystemRole['News'] = $true }
	#		'iis_smtp' { $script:SystemRole['SMTP'] = $true }
	#		'iis_www' { $script:SystemRole['WWW'] = $true }
	#		'certsrv_server' { $script:SystemRole['PKI'] = $true }
	#		'dhcpserver' { $script:SystemRole['DHCP'] = $true }
	#		'dns' { $script:SystemRole['DNS'] = $true }
	#		'ias' { $script:SystemRole['IAS'] = $true }
	#		'wins' { $script:SystemRole['WINS'] = $true }
	#		'reminst' { $script:SystemRole['RIS'] = $true }
	#		'tsenable' { if ($script:Reg_TerminalServerMode) { $script:SystemRole['TS'] = $true } }
	#		'terminalserver' { if ($script:Reg_TerminalServerMode) { $script:SystemRole['TS'] = $true } }
	#		'wms_server' { $script:SystemRole['Media'] = $true }
	#	}

}

function Get-OSLanguage($LanguageCode) {
	Write-Output $(
		switch ($LanguageCode) {
			1 {'Arabic'}
			4 {'Chinese (Simplified)– China'}
			9 {'English'}
			1025 {'Arabic – Saudi Arabia'}
			1026 {'Bulgarian'}
			1027 {'Catalan'}
			1028 {'Chinese (Traditional) – Taiwan'}
			1029 {'Czech'}
			1030 {'Danish'}
			1031 {'German – Germany'}
			1032 {'Greek'}
			1033 {'English – United States'}
			1034 {'Spanish – Traditional Sort'}
			1035 {'Finnish'}
			1036 {'French – France'}
			1037 {'Hebrew'}
			1038 {'Hungarian'}
			1039 {'Icelandic'}
			1040 {'Italian – Italy'}
			1041 {'Japanese'}
			1042 {'Korean'}
			1043 {'Dutch – Netherlands'}
			1044 {'Norwegian – Bokmal'}
			1045 {'Polish'}
			1046 {'Portuguese – Brazil'}
			1047 {'Rhaeto-Romanic'}
			1048 {'Romanian'}
			1049 {'Russian'}
			1050 {'Croatian'}
			1051 {'Slovak'}
			1052 {'Albanian'}
			1053 {'Swedish'}
			1054 {'Thai'}
			1055 {'Turkish'}
			1056 {'Urdu'}
			1057 {'Indonesian'}
			1058 {'Ukrainian'}
			1059 {'Belarusian'}
			1060 {'Slovenian'}
			1061 {'Estonian'}
			1062 {'Latvian'}
			1063 {'Lithuanian'}
			1065 {'Persian'}
			1066 {'Vietnamese'}
			1069 {'Basque (Basque) – Basque'}
			1070 {'Serbian'}
			1071 {'Macedonian (FYROM)'}
			1072 {'Sutu'}
			1073 {'Tsonga'}
			1074 {'Tswana'}
			1076 {'Xhosa'}
			1077 {'Zulu'}
			1078 {'Afrikaans'}
			1080 {'Faeroese'}
			1081 {'Hindi'}
			1082 {'Maltese'}
			1084 {'Scottish Gaelic (United Kingdom)'}
			1085 {'Yiddish'}
			1086 {'Malay – Malaysia'}
			2049 {'Arabic – Iraq'}
			2052 {'Chinese (Simplified) – PRC'}
			2055 {'German – Switzerland'}
			2057 {'English – United Kingdom'}
			2058 {'Spanish – Mexico'}
			2060 {'French – Belgium'}
			2064 {'Italian – Switzerland'}
			2067 {'Dutch – Belgium'}
			2068 {'Norwegian – Nynorsk'}
			2070 {'Portuguese – Portugal'}
			2072 {'Romanian – Moldova'}
			2073 {'Russian – Moldova'}
			2074 {'Serbian – Latin'}
			2077 {'Swedish – Finland'}
			3073 {'Arabic – Egypt'}
			3076 {'Chinese (Traditional) – Hong Kong SAR'}
			3079 {'German – Austria'}
			3081 {'English – Australia'}
			3082 {'Spanish – International Sort'}
			3084 {'French – Canada'}
			3098 {'Serbian – Cyrillic'}
			4097 {'Arabic – Libya'}
			4100 {'Chinese (Simplified) – Singapore'}
			4103 {'German – Luxembourg'}
			4105 {'English – Canada'}
			4106 {'Spanish – Guatemala'}
			4108 {'French – Switzerland'}
			5121 {'Arabic – Algeria'}
			5127 {'German – Liechtenstein'}
			5129 {'English – New Zealand'}
			5130 {'Spanish – Costa Rica'}
			5132 {'French – Luxembourg'}
			6145 {'Arabic – Morocco'}
			6153 {'English – Ireland'}
			6154 {'Spanish – Panama'}
			7169 {'Arabic – Tunisia'}
			7177 {'English – South Africa'}
			7178 {'Spanish – Dominican Republic'}
			8193 {'Arabic – Oman'}
			8201 {'English – Jamaica'}
			8202 {'Spanish – Venezuela'}
			9217 {'Arabic – Yemen'}
			9226 {'Spanish – Colombia'}
			10241 {'Arabic – Syria'}
			10249 {'English – Belize'}
			10250 {'Spanish – Peru'}
			11265 {'Arabic – Jordan'}
			11273 {'English – Trinidad'}
			11274 {'Spanish – Argentina'}
			12289 {'Arabic – Lebanon'}
			12298 {'Spanish – Ecuador'}
			13313 {'Arabic – Kuwait'}
			13322 {'Spanish – Chile'}
			14337 {'Arabic – U.A.E.'}
			14346 {'Spanish – Uruguay'}
			15361 {'Arabic – Bahrain'}
			15370 {'Spanish – Paraguay'}
			16385 {'Arabic – Qatar'}
			16394 {'Spanish – Bolivia'}
			17418 {'Spanish – El Salvador'}
			18442 {'Spanish – Honduras'}
			19466 {'Spanish – Nicaragua'}
			20490 {'Spanish – Puerto Rico'}
			default {'Unknown (Undocumented)'} 
		}
	)
}


######################
# WMI FUNCTIONS
######################

function Get-OSInfo([string]$Computer) {
	# Win32_OperatingSystem: http://msdn.microsoft.com/en-us/library/aa394239(VS.85).aspx

	$OSInfo = $null

	#$Win32_OperatingSystem = Get-WmiObjectWithTimeout -Namespace root\CIMV2 -Class Win32_OperatingSystem -Property Name, CSDVersion, InstallDate, OSLanguage, Version, WindowsDirectory -ComputerName $Computer
	$Win32_OperatingSystem = Get-WmiObject -Namespace root\CIMV2 -Class Win32_OperatingSystem -Property Name, CSDVersion, InstallDate, OSLanguage, Version, WindowsDirectory -ComputerName $Computer

	# This call should only bring back 1 row but just in case there are more we'll iterate through and make the last one the winner
	$Win32_OperatingSystem | ForEach-Object {
		$OSInfo = New-Object -TypeName psobject -Property @{
			Name = $_.Name.split('|')[0].Trim()
			ServicePack = $_.CSDVersion
			InstallDateUTC = $Win32_OperatingSystem.ConvertToDateTime($_.InstallDate).ToUniversalTime()
			Language = Get-OSLanguage -LanguageCode $_.OSLanguage
			Version = $_.Version
			WindowsDirectory = $_.WindowsDirectory
			VersionNumber = [Single]($_.Version).Substring(0, ($_.Version).LastIndexOf('.'))
		} 

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tOperating System:"
		$OSInfo.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	}
	Write-Output $OSInfo

	Remove-Variable -Name OSInfo, Win32_OperatingSystem

}

function Get-BIOSInfo([string]$Computer) {
	# Win32_BIOS: http://msdn.microsoft.com/en-us/library/aa394077(VS.85).aspx

	$BIOSInfo = $null

	#$Win32_BIOS = Get-WmiObjectWithTimeout -Namespace root\CIMV2 -Class Win32_BIOS -Property BiosCharacteristics, SMBIOSBIOSVersion, SMBIOSMajorVersion, SMBIOSMinorVersion, Version -ComputerName $Computer
	$Win32_BIOS = Get-WmiObject -Namespace root\CIMV2 -Class Win32_BIOS -Property BiosCharacteristics, SMBIOSBIOSVersion, SMBIOSMajorVersion, SMBIOSMinorVersion, Version -ComputerName $Computer

	# This call should only bring back 1 row but just in case there are more we'll iterate through and make the last one the winner
	$Win32_BIOS | ForEach-Object {
		$BIOSInfo = New-Object -TypeName psobject -Property @{
			SMBIOSBIOSVersion = $_.SMBIOSBIOSVersion
			SMBIOSMajorVersion = $_.SMBIOSMajorVersion
			SMBIOSMinorVersion = $_.SMBIOSMinorVersion
			Version = $_.Version
			Characteristics = $_.BiosCharacteristics
		}

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tBIOS:"
		$BIOSInfo.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	}
	Write-Output $BIOSInfo

	Remove-Variable -Name BIOSInfo, Win32_BIOS
}

function Get-SystemInformation([string]$Computer) {
	# Win32_ComputerSystem: http://msdn.microsoft.com/en-us/library/aa394102(VS.85).aspx

	$SystemInformation = $null
	$ComputerRole = $null
	$DomainType = $null

	$ComputerSystem = $null
	#$Win32_ComputerSystem = Get-WmiObjectWithTimeout -Namespace root\CIMV2 -Class Win32_ComputerSystem -Property Domain, DomainRole, Name, NumberOfProcessors, TotalPhysicalMemory -ComputerName $Computer
	$Win32_ComputerSystem = Get-WmiObject -Namespace root\CIMV2 -Class Win32_ComputerSystem -Property Domain, DomainRole, Name, NumberOfProcessors, TotalPhysicalMemory -ComputerName $Computer

	$Win32_ComputerSystem | ForEach-Object {

		switch ($_.DomainRole) {
			0 {
				$ComputerRole = 'Standalone Workstation'
				$DomainType = 'workgroup'

			}
			1 {
				$ComputerRole = 'Member Workstation'
				$DomainType = 'domain'

			}
			2 {
				$ComputerRole = 'Standalone Server'
				$DomainType = 'workgroup'

			}
			3 {
				$ComputerRole = 'Member Server'
				$DomainType = 'domain'

			}
			4 {
				$ComputerRole = 'Domain Controller'
				$DomainType = 'domain'

				#	            $script:bWMILocalAccounts = $false
				#	            $script:bWMILocalGroups = $false
				#				$script:SystemRole['DC'] = $true

			}
			5 {
				$ComputerRole = 'Domain Controller (PDC Emulator)'
				$DomainType = 'domain'

				#	            $script:bWMILocalAccounts = $false
				#	            $script:bWMILocalGroups = $false
				#				$script:SystemRole['DC'] = $true   
			}
			default {
				$ComputerRole = 'Unknown'
				$DomainType = 'unknown'

			} 
		} 

		## I was using Add-Member to figure out the FQDN but ran into a problem
		## when this module was nested within other modules on Windows XP
		## Yeah, Windows XP, I know...but I had to run this code on an XP machine
		## so I removed the Add-Member call.

		# 		$SystemInformation = New-Object -TypeName psobject -Property @{
		# 			Domain = $_.Domain
		# 			DomainRole = $_.DomainRole
		# 			Name = $_.Name
		# 			NumberOfProcessors = $_.NumberOfProcessors
		# 			TotalPhysicalMemoryBytes = $_.TotalPhysicalMemory
		# 			ComputerRole = $ComputerRole
		# 			DomainType = $DomainType
		# 		} | Add-Member -MemberType ScriptProperty -Name FullyQualifiedDomainName -Value {
		# 			if ($this.DomainType -ieq 'domain') {
		# 				[String]::Join('.', @($this.Name, $this.Domain))
		# 			} else {
		# 				$this.Name
		# 			}
		# 		} -PassThru

		$SystemInformation = New-Object -TypeName psobject -Property @{
			Domain = $_.Domain
			DomainRole = $_.DomainRole
			Name = $_.Name
			FullyQualifiedDomainName = if ($DomainType -ieq 'domain') {
				[String]::Join('.', @($_.Name, $_.Domain))
			} else {
				$_.Name
			}
			NumberOfProcessors = $_.NumberOfProcessors
			TotalPhysicalMemoryBytes = $_.TotalPhysicalMemory
			ComputerRole = $ComputerRole
			DomainType = $DomainType
		}
	}

	Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tSystem:"
	$SystemInformation.psobject.Properties | ForEach-Object {
		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
	}

	Write-Output $SystemInformation

	Remove-Variable -Name SystemInformation, ComputerRole, DomainType, ComputerSystem, Win32_ComputerSystem
}

function Get-CDROMInformation([string]$Computer) {
	# Win32_CDROMDrive: http://msdn.microsoft.com/en-us/library/aa394081(VS.85).aspx

	$CDROMInformation = @()
	$CDROMDrive = $null

	#$Win32_CDROMDrive = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_CDROMDrive -Property Drive, Manufacturer, Name -ComputerName $Computer
	$Win32_CDROMDrive = Get-WMIObject -Namespace root\CIMV2 -Class Win32_CDROMDrive -Property Drive, Manufacturer, Name -ComputerName $Computer

	$Win32_CDROMDrive | Sort-Object -Property Drive | ForEach-Object {

		$CDROMDrive = New-Object -TypeName psobject -Property @{ 
			Drive = $_.Drive
			Manufacturer = $_.Manufacturer
			Name = $_.Name
		}

		$CDROMInformation += $CDROMDrive

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tCDROM:"
		$CDROMDrive.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	}
	Write-Output $CDROMInformation

	Remove-Variable -Name CDROMInformation, CDROMDrive, Win32_CDROMDrive
}

function Get-ComputerSystemProductInformation([string]$Computer) {
	# Win32_ComputerSystemProduct: http://msdn.microsoft.com/en-us/library/aa394105(VS.85).aspx

	$ComputerSystemProductInformation = $null

	#$Win32_ComputerSystemProduct = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_ComputerSystemProduct -Property Vendor, Name, IdentifyingNumber, Version -ComputerName $Computer
	$Win32_ComputerSystemProduct = Get-WMIObject -Namespace root\CIMV2 -Class Win32_ComputerSystemProduct -Property Vendor, Name, IdentifyingNumber, Version -ComputerName $Computer

	# This call should only bring back 1 row but just in case there are more we'll iterate through and make the last one the winner
	$Win32_ComputerSystemProduct | ForEach-Object {
		$ComputerSystemProductInformation = New-Object -TypeName psobject -Property @{
			Manufacturer = $_.Vendor
			Name = $_.Name
			IdentifyingNumber = $_.IdentifyingNumber
			Version = $_.Version
		}

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tComputer System Product:"
		$ComputerSystemProductInformation.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	}
	Write-Output $ComputerSystemProductInformation

	Remove-Variable -Name ComputerSystemProductInformation, Win32_ComputerSystemProduct
}

function Get-DiskInformation([string]$Computer, [System.Version]$OSVersion) {
	# Win32_DiskDrive: http://msdn.microsoft.com/en-us/library/aa394132(VS.85).aspx
	# Win32_DiskDriveToDiskPartition: http://msdn.microsoft.com/en-us/library/aa394134(VS.85).aspx
	# Win32_DiskPartition: http://msdn.microsoft.com/en-us/library/aa394135(VS.85).aspx
	# Win32_LogicalDiskToPartition: http://msdn.microsoft.com/en-us/library/aa394175(VS.85).aspx
	# Win32_LogicalDisk: http://msdn.microsoft.com/en-us/library/aa394173(VS.85).aspx
	# Win32_Volume: http://msdn.microsoft.com/en-us/library/aa394515(VS.85).aspx

	$DiskInformation = @()
	$Drive = $null
	$Partition = $null
	$Win32_DiskPartition = $null
	$Win32_LogicalDisk = $null
	$Win32_Volume = $null
	$Win32_DiskDrive = $null
	$DeviceID = $null
	$PartitionDeviceID = $null
	$AllocationUnitSizeBytes = $null

	#$Win32_DiskDrive = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_DiskDrive -Property Caption, DeviceID, Interfacetype, Size -ComputerName $Computer
	$Win32_DiskDrive = Get-WMIObject -Namespace root\CIMV2 -Class Win32_DiskDrive -Property Caption, DeviceID, Interfacetype, Size -ComputerName $Computer

	#$WMIObject | Get-Member -MemberType Property | Format-Table -Property "Name"

	## Win32_DiskDrive -> Win32_DiskDriveToDiskPartition -> Win32_DiskPartition -> Win32_LogicalDiskToPartition -> Win32_LogicalDisk
	#Get-WmiObject -Class Win32_DiskDrive
	#Get-WmiObject -Class Win32_DiskDriveToDiskPartition
	#Get-WmiObject -Class Win32_DiskPartition
	#Get-WmiObject -Class Win32_LogicalDiskToPartition
	#Get-WmiObject -Class Win32_LogicalDisk
	#
	#Get-WmiObject -Class Win32_Volume
	#Get-WmiObject -Class Win32_DiskDrivePhysicalMedia	# Kind of useless
	#Get-WmiObject -Class Win32_MountPoint

	$Win32_DiskDrive | ForEach-Object {

		$Drive = New-Object -TypeName psobject -Property @{
			Caption = $_.Caption
			DeviceID = $_.DeviceID ; 
			Interfacetype = $_.Interfacetype ; 
			SizeBytes = $_.Size ;
			Partitions = @()
		}

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tDrive:"
		$Drive.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}

		$DeviceID = $_.DeviceID.Replace('\','\\')
		#$Win32_DiskPartition = Get-WmiObject -Namespace root\CIMV2 -Query "Associators of {Win32_DiskDrive.DeviceID=""$DeviceID""} WHERE AssocClass = Win32_DiskDriveToDiskPartition" -ComputerName $Computer
		$Win32_DiskPartition = Get-WmiObject -Namespace root\CIMV2 -Query "Associators of {Win32_DiskDrive.DeviceID=""$DeviceID""} WHERE ResultClass = Win32_DiskPartition" -ComputerName $Computer
		#$Win32_DiskPartition = $_.GetRelated("Win32_DiskPartition")

		$Win32_DiskPartition | Where-Object { $_.DeviceID } | ForEach-Object {

			$Partition = New-Object -TypeName psobject -Property @{
				Caption = $_.Caption ;
				StartingOffsetBytes = $_.StartingOffset ;
				BlockSizeBytes = $_.BlockSize ;
				LogicalDisks = @() ;
			}

			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`tPartition:"
			$Partition.psobject.Properties | ForEach-Object {
				Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`t`t$($_.Name): $($_.Value)"
			}

			$PartitionDeviceID = $_.DeviceID
			#$Win32_LogicalDisk = Get-WmiObject -Namespace root\CIMV2 -Query "Associators of {Win32_DiskPartition.DeviceID=""$PartitionDeviceID""} WHERE AssocClass = Win32_LogicalDiskToPartition" -ComputerName $Computer
			$Win32_LogicalDisk = Get-WmiObject -Namespace root\CIMV2 -Query "Associators of {Win32_DiskPartition.DeviceID=""$PartitionDeviceID""} WHERE ResultClass = Win32_LogicalDisk" -ComputerName $Computer
			#$Win32_LogicalDisk = $_.GetRelated("Win32_LogicalDisk")

			$Win32_LogicalDisk | Where-Object { $_.DeviceID } | ForEach-Object {

				# Win32_Volume only supported in Windows Server 2003 and higher
				if ($OSVersion.CompareTo($WindowsServer2003) -ge 0) {

					try {

						#		 				# Method 1 - via Win32_MountPoint
						#						$Win32_MountPoint = Get-WmiObject -Query "select Volume FROM Win32_MountPoint WHERE (Directory = `'Win32_Directory.Name=""$Caption\\\\""`')" -ComputerName $Computer
						#						$VolumePath = $Win32_MountPoint.Volume.Replace("\","\\")
						#						
						#						$Win32_Volume = Get-WmiObject -Query "select BlockSize FROM Win32_Volume WHERE __RELPATH = `'$VolumePath`'" -ComputerName $Computer
						#						$AllocationUnitSize = $Win32_Volume.BlockSize

						# Method 2 - using drive letter
						#$Win32_Volume = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_Volume -Property BlockSize -Filter "DriveLetter = ""$($_.DeviceID)""" -ComputerName $Computer
						$Win32_Volume = Get-WMIObject -Namespace root\CIMV2 -Class Win32_Volume -Property BlockSize -Filter "DriveLetter = ""$($_.DeviceID)""" -ComputerName $Computer

						$AllocationUnitSizeBytes = $Win32_Volume.BlockSize

					}
					catch {
						$AllocationUnitSizeBytes = $null
					}

				}
				else {
					$AllocationUnitSizeBytes = $null
				}

				$LogicalDisk = New-Object -TypeName psobject -Property @{
					AllocationUnitSizeBytes = $AllocationUnitSizeBytes
					Caption = $_.Caption
					FreeSpaceBytes = $_.FreeSpace
					FileSystem = $_.FileSystem
					SizeBytes = $_.Size
					VolumeName = $_.VolumeName
				}

				$Partition.LogicalDisks += $LogicalDisk

				Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`t`t`tLogical Disk:"
				$LogicalDisk.psobject.Properties | ForEach-Object {
					Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`t`t`t`t$($_.Name): $($_.Value)"
				}

			}
			$Drive.Partitions += $Partition
		}
		$DiskInformation += $Drive
	}
	Write-Output $DiskInformation

	Remove-Variable -Name DiskInformation, Drive, Partition, Win32_DiskPartition, Win32_LogicalDisk, Win32_Volume, Win32_DiskDrive, DeviceID, PartitionDeviceID, AllocationUnitSizeBytes

}

function Get-LocalGroupsInformation([string]$Computer, [System.Version]$OSVersion) {
	# Win32_Group: http://msdn.microsoft.com/en-us/library/aa394151(VS.85).aspx
	# Win32_GroupUser: http://msdn.microsoft.com/en-us/library/aa394153(VS.85).aspx

	$LocalGroupsInformation = @()
	$Win32_GroupUser = $null
	$Group = $null
	$User = $null
	$PartComponent = $null


	# Query is different for Windows 2000 vs. Windows XP and higher
	if ($OSVersion.CompareTo($WindowsXP) -ge 0) {
		#$Win32_Group = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_Group -Property Name, Domain -Filter 'LocalAccount = True' -ComputerName $Computer
		$Win32_Group = Get-WMIObject -Namespace root\CIMV2 -Class Win32_Group -Property Name, Domain -Filter 'LocalAccount = True' -ComputerName $Computer
	} else {
		#$Win32_Group = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_Group -Property Name, Domain -Filter "(__SERVER = Domain) or (Domain = 'BUILTIN')" -ComputerName $Computer
		$Win32_Group = Get-WMIObject -Namespace root\CIMV2 -Class Win32_Group -Property Name, Domain -Filter "(__SERVER = Domain) or (Domain = 'BUILTIN')" -ComputerName $Computer
	}

	$Win32_Group | Sort-Object -Property Name | ForEach-Object {

		$Group = New-Object -TypeName psobject -Property @{
			Name = $_.Name
			Members = @()
		}

		## Example calls
		#Get-WmiObject -Query "Select * from Win32_GroupUser where GroupComponent='Win32_Group.Domain=""T420WIN7-005"",Name=""Guests""'"
		#Get-WmiObject -Query "Select PartComponent from Win32_GroupUser where GroupComponent='Win32_Group.Domain=""T420WIN7-005"",Name=""Administrators""'"

		#$Win32_GroupUser = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_GroupUser -Property PartComponent -Filter "GroupComponent='Win32_Group.Domain=""$($_.Domain)"",Name=""$($_.Name)""'" -ComputerName $Computer
		$Win32_GroupUser = Get-WMIObject -Namespace root\CIMV2 -Class Win32_GroupUser -Property PartComponent -Filter "GroupComponent='Win32_Group.Domain=""$($_.Domain)"",Name=""$($_.Name)""'" -ComputerName $Computer

		$Win32_GroupUser | Sort-Object -Property PartComponent | ForEach-Object {

			If (!$_.PartComponent) {
				$User = 'UNKNOWN'
			} else {
				# Could this be done better with a regular expression and matches?
				$PartComponent = $_.PartComponent.Split('"')
				$User = "$($PartComponent[1])\$($PartComponent[3])"
			}

			$Group.Members += $User
		}

		$LocalGroupsInformation += $Group

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tGroup:"
		$Group.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		} 
	}
	Write-Output $LocalGroupsInformation

	Remove-Variable -Name LocalGroupsInformation, Win32_GroupUser, Group, User, PartComponent

}

function Get-IpRouteInformation([string]$Computer) {
	# Win32_IP4RouteTable: http://msdn.microsoft.com/en-us/library/aa394162(VS.85).aspx

	$IpRouteInformation = @()
	$Route = $null

	#$Win32_IP4RouteTable = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_IP4RouteTable -Property Destination, Mask, NextHop, Metric1, Metric2, Metric3, Metric4, Metric5 -ComputerName $Computer
	$Win32_IP4RouteTable = Get-WMIObject -Namespace root\CIMV2 -Class Win32_IP4RouteTable -Property Destination, Mask, NextHop, Metric1, Metric2, Metric3, Metric4, Metric5 -ComputerName $Computer

	$Win32_IP4RouteTable | ForEach-Object {
		$Route = New-Object -TypeName psobject -Property @{
			Destination = $_.Destination
			Mask = $_.Mask
			NextHop = $_.NextHop
			Metric1 = $_.Metric1
			Metric2 = $_.Metric2
			Metric3 = $_.Metric3
			Metric4 = $_.Metric4
			Metric5 = $_.Metric5
		}

		$IpRouteInformation += $Route

		$Route.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	}
	Write-Output $IpRouteInformation

	Remove-Variable -Name IpRouteInformation, Route, Win32_IP4RouteTable
}


function Get-DesktopSessionInformation([string]$Computer) {

	# Uses RDS-Manager module from http://gallery.technet.microsoft.com/ScriptCenter/e8c3af96-db10-45b0-88e3-328f087a8700/
	# Or, if RDS-Manager is not loaded, uses Win32_Process: http://msdn.microsoft.com/en-us/library/aa394372(VS.85).aspx

	$DesktopSession = @()
	$WTSConnectStateEnum = $null
	$Session = $null
	$RDSessionSuccess = $false

	if ((Test-Path -Path 'function:Get-RDSession') -eq $true) {

		try {

			$WTSConnectStateEnum = 'RDSManager.PowerShell.WTSConnectState' -as [Type]

			Get-RDSession -RDSHost $Computer | ForEach-Object {
				$Session = New-Object -TypeName psobject -Property @{
					Client = $_.Client
					Host = $_.Host
					IdleTime = $_.IdleTime
					LogonTimeUTC = if ($_.LogonTime) { (Get-Date -Date $_.LogonTime).ToUniversalTime() } else { $null }
					ProtocolType = $_.ProtocolType
					Session = $_.Session
					SessionID = $_.SessionID

					# See http://msdn.microsoft.com/en-us/library/windows/desktop/aa383860(v=vs.85).aspx for Connection State Info
					State = switch ($_.State) {
						$WTSConnectStateEnum::WTSActive { 'Active' }
						$WTSConnectStateEnum::WTSConnected { 'Connected' }
						$WTSConnectStateEnum::WTSConnectQuery { 'Connecting' }
						$WTSConnectStateEnum::WTSShadow { 'Shadow' }
						$WTSConnectStateEnum::WTSDisconnected { 'Disconnected' }
						$WTSConnectStateEnum::WTSIdle { 'Idle' }
						$WTSConnectStateEnum::WTSListen { 'Listening' }
						$WTSConnectStateEnum::WTSReset { 'Resetting' }
						$WTSConnectStateEnum::WTSDown { 'Down Due To Error' }
						$WTSConnectStateEnum::WTSInit { 'Initializing' }
					} 
					User = $_.User
				}

				$DesktopSession += $Session

				Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tDesktop Session:"
				$Session.psobject.Properties | ForEach-Object {
					Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
				}

			}
		}
		catch {
			$ErrorRecord = $_.Exception.ErrorRecord
			Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering information about logged on users: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
			Write-WindowsMachineInformationLog -Message "[$ComputerName] Reverting to Win32_Process to determine logged on users" -MessageLevel Warning
			$RDSessionSuccess = $false
			$DesktopSession = @()
		}

	}

	if ($RDSessionSuccess -ne $true) {

		#Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_Process -Filter 'name="explorer.exe"' -ComputerName $Computer | ForEach-Object {
		Get-WMIObject -Namespace root\CIMV2 -Class Win32_Process -Filter 'name="explorer.exe"' -ComputerName $Computer | ForEach-Object {
			$Session = New-Object -TypeName psobject -Property @{
				Client = $null
				Host = $null
				IdleTime = $null
				LogonTimeUTC = $_.ConvertToDateTime($_.CreationDate).ToUniversalTime()
				ProtocolType = $null
				Session = $null
				SessionID = $null
				State = $null
				User = $_.GetOwner() | ForEach-Object { [String]::Join('\', @($_.Domain, $_.User)) }
			}

			$DesktopSession += $Session

			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tDesktop Session:"
			$Session.psobject.Properties | ForEach-Object {
				Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
			} 

		}

	}

	Write-Output $DesktopSession

	Remove-Variable -Name DesktopSession, WTSConnectStateEnum, Session, RDSessionSuccess

}


function Get-NetworkAdapterConfig([string]$Computer) {
	# Win32_NetworkAdapterConfiguration: http://msdn.microsoft.com/en-us/library/aa394217(VS.85).aspx

	$NetworkAdapterConfig = @()
	$Adapter = $null 

	#$Win32_NetworkAdapterConfiguration = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_NetworkAdapterConfiguration -Property Description, MACAddress, DNSHostName, DHCPEnabled, DHCPServer, DNSDomain, WINSPrimaryServer, WINSSecondaryServer, IPAddress, IPSubnet, DefaultIPGateway, DNSServerSearchOrder -Filter 'IPEnabled = True' -ComputerName $Computer
	$Win32_NetworkAdapterConfiguration = Get-WMIObject -Namespace root\CIMV2 -Class Win32_NetworkAdapterConfiguration -Property Description, MACAddress, DNSHostName, DHCPEnabled, DHCPServer, DNSDomain, WINSPrimaryServer, WINSSecondaryServer, IPAddress, IPSubnet, DefaultIPGateway, DNSServerSearchOrder -Filter 'IPEnabled = True' -ComputerName $Computer

	$Win32_NetworkAdapterConfiguration | ForEach-Object {

		$Adapter = New-Object -TypeName psobject -Property @{
			Description = $_.Description
			MACAddress = $_.MACAddress
			DNSHostName = $_.DNSHostName
			DHCPEnabled = $_.DHCPEnabled
			DHCPServer = $_.DHCPServer
			DNS = $_.DNSDomain
			WINSPrimaryServer = $_.WINSPrimaryServer
			WINSSecondaryServer = $_.WINSSecondaryServer
			IPAddress = $_.IPAddress
			IPSubnet = $_.IPSubnet
			DefaultIPGateway = $_.DefaultIPGateway
			DNSServerSearchOrder = $_.DNSServerSearchOrder
		}

		$NetworkAdapterConfig += $Adapter

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tAdapter:"
		$Adapter.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	}
	Write-Output $NetworkAdapterConfig

	Remove-Variable -Name NetworkAdapterConfig, Adapter, Win32_NetworkAdapterConfiguration
}

function Get-EventLogSettings([string]$Computer) {
	# Win32_NTEventLogFile: http://msdn.microsoft.com/en-us/library/aa394225(VS.85).aspx

	$EventLogSettings = @()
	$EventLog = $null

	#$Win32_NTEventLogFile = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_NTEventlogFile -Property LogFileName, MaxFileSize, Name, OverwritePolicy -ComputerName $Computer
	$Win32_NTEventLogFile = Get-WMIObject -Namespace root\CIMV2 -Class Win32_NTEventlogFile -Property LogFileName, MaxFileSize, Name, OverwritePolicy -ComputerName $Computer

	$Win32_NTEventLogFile | ForEach-Object {

		$EventLog = New-Object -TypeName psobject -Property @{
			FileName = $_.Name
			MaxFileSizeBytes = $_.MaxFileSize
			LogName = $_.LogFileName
			OverwritePolicy = $_.OverwritePolicy
		}

		$EventLogSettings += $EventLog

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tEvent Log:"
		$EventLog.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	}
	Write-Output $EventLogSettings

	Remove-Variable -Name EventLogSettings, EventLog, Win32_NTEventLogFile
}

function Get-PagefileInformation([string]$Computer, [System.Version]$OSVersion) {
	# Win32_ComputerSystem: http://msdn.microsoft.com/en-us/library/aa394102(VS.85).aspx
	# Win32_PageFileUsage: http://msdn.microsoft.com/en-us/library/aa394246(VS.85).aspx
	# Win32_PageFileSetting: http://msdn.microsoft.com/en-us/library/aa394245(VS.85).aspx

	$PagefileInformation = @()
	$PageFile = $null 
	$MaximumSizeMB = $null
	$AutomaticManagedPagefile = $false
	$Win32_ComputerSystem = $null
	$Win32_PageFileSetting = $null
	$Win32_PageFileUsage = $null 

	# In Windows Server 2008 (OS Version 6) and up you query Win32_ComputerSystem.AutomaticManagedPagefile to find out if enabled
	#	If yes, Win32_PageFileSetting will be empty
	# 	If no, Win32_PageFileSetting will contain information
	#	Either way, use Win32_PageFileUsage to get actual usage information
	# For < Windows 2008 always use both Win32_PageFileSetting and Win32_PageFileUsage to get settings and runtime information

	# Windows Vista & 2008 introduced Automatic Page File Settings.
	if ($OSVersion.CompareTo($WindowsVista) -ge 0) {
		#$Win32_ComputerSystem = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_ComputerSystem -Property AutomaticManagedPagefile -ComputerName $Computer
		$Win32_ComputerSystem = Get-WMIObject -Namespace root\CIMV2 -Class Win32_ComputerSystem -Property AutomaticManagedPagefile -ComputerName $Computer
		if ( $Win32_ComputerSystem.AutomaticManagedPagefile -ieq 'true') { 
			$AutomaticManagedPagefile = $true
		}
	}

	# TempPageFile not available via WMI in Windows 2000 and below
	if ($OSVersion.CompareTo($WindowsXP) -ge 0) {
		#$Win32_PageFileUsage = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_PageFileUsage -Property AllocatedBaseSize, CurrentUsage, Name, PeakUsage, TempPageFile -ComputerName $Computer
		$Win32_PageFileUsage = Get-WMIObject -Namespace root\CIMV2 -Class Win32_PageFileUsage -Property AllocatedBaseSize, CurrentUsage, Name, PeakUsage, TempPageFile -ComputerName $Computer
	} else {
		#$Win32_PageFileUsage = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_PageFileUsage -Property AllocatedBaseSize, CurrentUsage, Name, PeakUsage -ComputerName $Computer
		$Win32_PageFileUsage = Get-WMIObject -Namespace root\CIMV2 -Class Win32_PageFileUsage -Property AllocatedBaseSize, CurrentUsage, Name, PeakUsage -ComputerName $Computer
	}

	$Win32_PageFileUsage | Sort-Object -Property Name | ForEach-Object {

		if ($AutomaticManagedPagefile -ne $true) {
			#$Win32_PageFileSetting = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_PageFileSetting -Property MaximumSize -Filter "Name = '$($_.Name.Replace(""\"",""\\""))'" -ComputerName $Computer
			$Win32_PageFileSetting = Get-WMIObject -Namespace root\CIMV2 -Class Win32_PageFileSetting -Property MaximumSize -Filter "Name = '$($_.Name.Replace(""\"",""\\""))'" -ComputerName $Computer
			$MaximumSizeMB = $Win32_PageFileSetting.MaximumSize
		} else {
			$MaximumSizeMB = $null
		}

		$PageFile = New-Object -TypeName psobject -Property @{
			Drive = $($_.Name).Split('\')[0]
			InitialSizeMB = $_.AllocatedBaseSize
			MaximumSizeMB = $MaximumSizeMB
			CurrentSizeMB = $_.CurrentUsage
			PeakSizeMB = $_.PeakUsage
			IsTemporary = $_.TempPageFile
			IsAutoManaged = $AutomaticManagedPagefile
		}

		$PagefileInformation += $PageFile

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tPage File:"
		$PageFile.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		} 
	}
	Write-Output $PagefileInformation

	Remove-Variable -Name PagefileInformation, PageFile, MaximumSizeMB, AutomaticManagedPagefile, Win32_ComputerSystem, Win32_PageFileSetting, Win32_PageFileUsage

}

function Get-PhysicalMemoryInformation([string]$Computer) {
	# Win32_PhysicalMemory: http://msdn.microsoft.com/en-us/library/aa394347(VS.85).aspx

	$PhysicalMemoryInformation = @()
	$PhysicalMemory = $null
	$FormFactor = $null
	$HotSwappable = $null

	#$Win32_PhysicalMemory = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_PhysicalMemory -Property BankLabel, Capacity, DeviceLocator, FormFactor, HotSwappable, Manufacturer, MemoryType, PartNumber, SerialNumber, Speed, TypeDetail -ComputerName $Computer
	$Win32_PhysicalMemory = Get-WMIObject -Namespace root\CIMV2 -Class Win32_PhysicalMemory -Property BankLabel, Capacity, DeviceLocator, FormFactor, HotSwappable, Manufacturer, MemoryType, PartNumber, SerialNumber, Speed, TypeDetail -ComputerName $Computer

	$Win32_PhysicalMemory | Sort-Object -Property @{Expression={$_['BankLabel']};Descending=$true}, @{Expression={$_['DeviceLocator']};Ascending=$true} | ForEach-Object {

		if (!$_.FormFactor) {
			$FormFactor = 'Unknown'
		} else {
			$FormFactor = Get-PhysicalMemoryFormFactor($_.FormFactor)
		}

		if ($_.HotSwappable -ine 'true') {
			$HotSwappable = $false
		} else {
			$HotSwappable = $true
		}

		$PhysicalMemory = New-Object -TypeName psobject -Property @{
			BankLabel = $_.BankLabel
			CapacityBytes = $_.Capacity
			DeviceLocator = $_.DeviceLocator
			FormFactor = $FormFactor
			HotSwappable = $HotSwappable
			Manufacturer = $_.Manufacturer
			MemoryType = Get-PhysicalMemoryType($_.MemoryType)
			PartNumber = $_.PartNumber
			SerialNumber = $_.SerialNumber
			Speed = $_.Speed
			TypeDetail = Get-PhysicalMemoryTypeDetail($_.TypeDetail)
		}

		$PhysicalMemoryInformation += $PhysicalMemory

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tPhysical Memory:"
		$PhysicalMemory.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	}
	Write-Output $PhysicalMemoryInformation
}

function Get-PowerPlanInformation([string]$Computer) {
	# Win32_PowerPlan: http://msdn.microsoft.com/en-us/library/dd904531(VS.85).aspx
	# Win32_PowerSettingDataIndex: http://msdn.microsoft.com/en-us/library/dd904534(VS.85).aspx
	# Win32_PowerSetting: http://msdn.microsoft.com/en-us/library/dd904532(VS.85).aspx

	$PowerPlanInformation = @()
	$PowerPlan = $null
	$Win32_PowerPlan = $null

	try {
		#$Win32_PowerPlan = Get-WMIObjectWithTimeout -Namespace root\CIMV2\power -Class Win32_PowerPlan -Property ElementName, Description, IsActive -ComputerName $Computer -ErrorAction Stop
		$Win32_PowerPlan = Get-WMIObject -Namespace root\CIMV2\power -Class Win32_PowerPlan -Property ElementName, Description, IsActive -ComputerName $Computer -ErrorAction Stop
	}
	catch {
		Write-WindowsMachineInformationLog -Message "`t[$Computer] Unable to gather information about power plans: Check that Win32_PowerPlan class is installed" -MessageLevel Warning
		return 
	}

	$Win32_PowerPlan | Sort-Object -Property ElementName | ForEach-Object {

		$PowerPlan = New-Object -TypeName psobject -Property @{
			PlanName = $_.ElementName
			Description = $_.Description
			IsActive = $_.IsActive
		}

		$PowerPlanInformation += $PowerPlan

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tPower Plan:"
		$PowerPlan.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	} 
	Write-Output $PowerPlanInformation

	Remove-Variable -Name PowerPlanInformation, PowerPlan, Win32_PowerPlan
}

function Get-PrinterInformation([string]$Computer) {
	# Win32_Printer: http://msdn.microsoft.com/en-us/library/aa394363(VS.85).aspx

	$PrinterInformation = @()
	$Printer = $null

	#$Win32_Printer = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_Printer -Property DriverName, Name, PortName -Filter 'ServerName = Null' -ComputerName $Computer
	$Win32_Printer = Get-WMIObject -Namespace root\CIMV2 -Class Win32_Printer -Property DriverName, Name, PortName -Filter 'ServerName = Null' -ComputerName $Computer

	$Win32_Printer | Sort-Object -Property Name | ForEach-Object {

		$Printer = New-Object -TypeName psobject -Property @{
			DriverName = $_.DriverName
			Name = $_.Name
			PortName = $_.PortName
		}

		$PrinterInformation += $Printer

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tPrinter:"
		$Printer.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	}
	Write-Output $PrinterInformation

	Remove-Variable -Name PrinterInformation, Printer, Win32_Printer
}

function Get-ProcessInformation([string]$Computer) {
	# Win32_Process: http://msdn.microsoft.com/en-us/library/aa394372(VS.85).aspx

	$ProcessInformation = @()
	$Process = $null
	$ProcessOwner = $null
	$OwnerUser = $null
	$OwnerDomain = $null

	#$Win32_Process = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_Process -ComputerName $Computer
	$Win32_Process = Get-WMIObject -Namespace root\CIMV2 -Class Win32_Process -ComputerName $Computer

	$Win32_Process | ForEach-Object {

		try {
			$ProcessOwner = $_.GetOwner()
			if ($ProcessOwner.ReturnValue -eq 0) {
				$OwnerUser = $ProcessOwner.User 
				$OwnerDomain = $ProcessOwner.Domain
			} else {
				$OwnerUser = $null
				$OwnerDomain = $null
			}
		}
		catch {
			$OwnerUser = $null
			$OwnerDomain = $null
		}

		$Process = New-Object -TypeName psobject -Property @{
			Caption = $_.Caption
			ExecutablePath = $_.ExecutablePath
			OwnerUser = $OwnerUser
			OwnerDomain = $OwnerDomain
		}

		$ProcessInformation += $Process

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tProcess:"
		$Process.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	} 
	Write-Output $ProcessInformation

	Remove-Variable -Name ProcessInformation, Process, ProcessOwner, OwnerUser, OwnerDomain, Win32_Process
}

function Get-ProcessorInformation([string]$Computer, [System.Version]$OSVersion) {
	# Win32_QuickFixEngineering: http://msdn.microsoft.com/en-us/library/aa394391(VS.85).aspx
	# Win32_Processor: http://msdn.microsoft.com/en-us/library/aa394373(VS.85).aspx

	$ProcessorInformation = $null
	$Win32_Processor = $null
	$Proc = $null
	$Hyperthreading = $false
	$NumberOfCores = $null
	$NumberOfLogicalProcessors = $null
	$NumberofPhysicalProcessors = $null
	$L3CacheSize = $null
	$L3CacheSpeed = $null


	# Windows 2008 and higher or Windows 2003 with KB932370 (http://support.microsoft.com/kb/932370/) applied
	# See http://www.symantec.com/connect/downloads/identifying-physical-hyperthreaded-and-multicore-processors-windows for more information
	if (
		$OSVersion.CompareTo($WindowsVista) -ge 0 -or
		(
			$OSVersion.CompareTo($WindowsServer2003) -ge 0 -and
			$(@(Get-WMIObject -Namespace root\CIMV2 -Class Win32_QuickFixEngineering -ComputerName $Computer | Where-Object { $_.HotFixID -like 'KB932370*' }).Count -gt 0)
			#$(@(Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_QuickFixEngineering -ComputerName $Computer | Where-Object { $_.HotFixID -like 'KB932370*' }).Count -gt 0)
		)
	) {

		# Windows Vista, Windows 2008, and higher provides L3 cache size and speed
		if ($OSVersion.CompareTo($WindowsVista) -ge 0) {
			#$Win32_Processor = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_Processor -Property DataWidth, Description, ExtClock, Family, L2CacheSize, L2CacheSpeed, L3CacheSize, L3CacheSpeed, Manufacturer, MaxClockSpeed, Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, SocketDesignation -ComputerName $Computer
			$Win32_Processor = Get-WMIObject -Namespace root\CIMV2 -Class Win32_Processor -Property DataWidth, Description, ExtClock, Family, L2CacheSize, L2CacheSpeed, L3CacheSize, L3CacheSpeed, Manufacturer, MaxClockSpeed, Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, SocketDesignation -ComputerName $Computer
		} else {
			#$Win32_Processor = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_Processor -Property DataWidth, Description, ExtClock, Family, L2CacheSize, L2CacheSpeed, Manufacturer, MaxClockSpeed, Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, SocketDesignation -ComputerName $Computer
			$Win32_Processor = Get-WMIObject -Namespace root\CIMV2 -Class Win32_Processor -Property DataWidth, Description, ExtClock, Family, L2CacheSize, L2CacheSpeed, Manufacturer, MaxClockSpeed, Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, SocketDesignation -ComputerName $Computer
		}

		# In a multi-socket system all processors have to be the same so we can just inspect the first one for details
		$Proc = @($Win32_Processor)[0]

		$NumberOfCores = $Proc.NumberOfCores
		$NumberOfLogicalProcessors = $Proc.NumberOfLogicalProcessors
		$NumberofPhysicalProcessors = @(@($Win32_Processor) | ForEach-Object {$_.SocketDesignation} | Select-Object -unique).count

		if ($NumberOfLogicalProcessors -gt $NumberOfCores) {
			$Hyperthreading = $true
		} else {
			$Hyperthreading = $false
		}

		if ($OSVersion.CompareTo($WindowsVista) -ge 0) {
			$L3CacheSize = $Proc.L3CacheSize
			$L3CacheSpeed = $Proc.L3CacheSpeed
		}

	} else {

		#$Win32_Processor = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_Processor -Property DataWidth, Description, ExtClock, Family, L2CacheSize, L2CacheSpeed, Manufacturer, MaxClockSpeed, Name, MaxClockSpeed, SocketDesignation -ComputerName $Computer
		$Win32_Processor = Get-WMIObject -Namespace root\CIMV2 -Class Win32_Processor -Property DataWidth, Description, ExtClock, Family, L2CacheSize, L2CacheSpeed, Manufacturer, MaxClockSpeed, Name, MaxClockSpeed, SocketDesignation -ComputerName $Computer

		# In a multi-socket system all processors have to be the same so we can just inspect the first one for details
		$Proc = @($Win32_Processor)[0]

		$NumberOfCores = $null
		$NumberOfLogicalProcessors = @($Win32_Processor).count
		$NumberofPhysicalProcessors = @(@($Win32_Processor) | ForEach-Object {$_.SocketDesignation} | select-object -unique).count

		if ($NumberOfLogicalProcessors -eq $NumberofPhysicalProcessors) {
			$Hyperthreading = $false
		} else {
			$Hyperthreading = 'Unknown'
		}

	}

	$ProcessorInformation = New-Object -TypeName psobject -Property @{
		DataWidth = $Proc.DataWidth
		Description = $Proc.Description
		ExtClockMHz = $Proc.ExtClock
		Family = Get-ProcessorFamily($Proc.Family)
		Hyperthreading = $Hyperthreading
		L2CacheSizeKB = $Proc.L2CacheSize
		L2CacheSpeedMHz = $Proc.L2CacheSpeed
		L3CacheSizeKB = $L3CacheSize
		L3CacheSpeedMHz = $L3CacheSpeed
		Manufacturer = $Proc.Manufacturer
		MaxClockSpeedMHz = $Proc.MaxClockSpeed
		Name = $Proc.Name -replace '\s{2,}',' '
		NumberOfCores = $NumberOfCores
		NumberOfLogicalProcessors = $NumberOfLogicalProcessors
		NumberofPhysicalProcessors = $NumberofPhysicalProcessors
	}

	$ProcessorInformation.psobject.Properties | ForEach-Object {
		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
	} 

	Write-Output $ProcessorInformation

	Remove-Variable -Name ProcessorInformation, Win32_Processor, Proc, Hyperthreading, NumberOfCores, NumberOfLogicalProcessors, NumberofPhysicalProcessors, L3CacheSize, L3CacheSpeed

}

function Get-ApplicationInformationFromWMI([string]$Computer, [System.Version]$OSVersion) {
	# Win32_Product: http://msdn.microsoft.com/en-us/library/aa394378(VS.85).aspx

	$ApplicationInformation = @()
	$Application = $null
	$InstallDateUTC = $null
	$Win32_Product = $null

	try {
		# Starting with Windows XP use InstallDate2 to get the correct install date
		if ($OSVersion.CompareTo($WindowsVista) -ge 0) {
			#$Win32_Product = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_Product -Property Name, Vendor, Version, InstallLocation, InstallDate, InstallDate2, HelpLink, URLInfoAbout, URLUpdateInfo -Filter 'Name <> Null' -ComputerName $Computer -ErrorAction Stop
			$Win32_Product = Get-WMIObject -Namespace root\CIMV2 -Class Win32_Product -Property Name, Vendor, Version, InstallLocation, InstallDate, InstallDate2, HelpLink, URLInfoAbout, URLUpdateInfo -Filter 'Name <> Null' -ComputerName $Computer -ErrorAction Stop
		} 
		elseif ($OSVersion.CompareTo($WindowsXP) -ge 0) {
			#$Win32_Product = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_Product -Property Name, Vendor, Version, InstallLocation, InstallDate, InstallDate2 -Filter 'Name <> Null' -ComputerName $Computer -ErrorAction Stop
			$Win32_Product = Get-WMIObject -Namespace root\CIMV2 -Class Win32_Product -Property Name, Vendor, Version, InstallLocation, InstallDate, InstallDate2 -Filter 'Name <> Null' -ComputerName $Computer -ErrorAction Stop
		} 
		else {
			#$Win32_Product = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_Product -Property Name, Vendor, Version, InstallLocation, InstallDate -Filter 'Name <> Null' -ComputerName $Computer -ErrorAction Stop
			$Win32_Product = Get-WMIObject -Namespace root\CIMV2 -Class Win32_Product -Property Name, Vendor, Version, InstallLocation, InstallDate -Filter 'Name <> Null' -ComputerName $Computer -ErrorAction Stop
		}
	}
	catch {
		Write-WindowsMachineInformationLog -Message "[$Computer] Unable to gather application information from WMI: Check that Win32_Product class is installed or Windows Installer Applications will not appear" -MessageLevel Warning
		Write-WindowsMachineInformationLog -Message "[$Computer] You can add it with Add/Remove Windows Components -> Management and Monitoring -> WMI Windows Installer Provider" -MessageLevel Warning
		#return
	}

	$Win32_Product | ForEach-Object {

		# According to http://msdn.microsoft.com/en-us/library/aa394378(VS.85).aspx starting with Windows XP use InstallDate2 to get the correct install date
		# But...on my systems tested it's not populated while InstallDate still is
		# So...as a fallback we can do some string parsing to produce InstallDate as yyyy-mm-dd

		if ($_.InstallDate2) {
			$InstallDateUTC = $_.ConvertToDateTime($_.InstallDate2).ToUniversalTime()
		} elseif ($_.InstallDate) {
			# If populated the format should be YYYYMMDD
			if ($($_.InstallDate).Length -eq 8) {
				$InstallDateUTC = $(Get-Date -Year $($_.InstallDate).Substring(0,4) -Month $($_.InstallDate).Substring(4,2) -Day $($_.InstallDate).Substring(6,2) -Hour 0 -Minute 0 -Second 0 )
			} else {
				# If we can't make a date out of it then don't use the date at all
				$InstallDateUTC = $null
			}
		}
		else {
			# Can't resolve a date so don't use any date at all
			$InstallDateUTC = $null
		}

		$Application = New-Object -TypeName psobject -Property @{
			ProductName = $_.Name
			Vendor = $_.Vendor
			Version = $_.Version
			InstallDateUTC = $InstallDateUTC
			InstallLocation = $_.InstallLocation
			HelpURL = $_.HelpLink
			SupportURL = $_.URLInfoAbout
			UpdateInfoURL = $_.URLUpdateInfo
			Source = 'WMI'
		}

		$ApplicationInformation += $Application

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tInstalled Application:"
		$Application.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	} 

	Write-Output $ApplicationInformation

	Remove-Variable -Name ApplicationInformation, Application, InstallDateUTC, Win32_Product
}

function Get-RegistrySizeInformation([string]$Computer) {
	# Win32_Registry: http://msdn.microsoft.com/en-us/library/aa394394(VS.85).aspx

	$RegistrySizeInformation = $null

	#$Win32_Registry = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_Registry -Property CurrentSize, MaximumSize -ComputerName $Computer
	$Win32_Registry = Get-WMIObject -Namespace root\CIMV2 -Class Win32_Registry -Property CurrentSize, MaximumSize -ComputerName $Computer

	# This call should only bring back 1 row but just in case there are more we'll iterate through and make the last one the winner
	$Win32_Registry | ForEach-Object {
		$RegistrySizeInformation = New-Object -TypeName psobject -Property @{
			CurrentSizeMB = $_.CurrentSize
			MaximumSizeMB = $_.MaximumSize
		}
	}

	Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`Registry:"
	$RegistrySizeInformation.psobject.Properties | ForEach-Object {
		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
	}

	Write-Output $RegistrySizeInformation

	Remove-Variable -Name RegistrySizeInformation, Win32_Registry
}

function Get-ServicesInformation([string]$Computer) {
	# Win32_Service: http://msdn.microsoft.com/en-us/library/aa394418(VS.85).aspx

	$ServicesInformation = @()
	$Service = $null

	#$Win32_Service = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_Service -Property Caption, PathName, Started, StartMode, StartName -Filter "ServiceType ='Share Process' Or ServiceType ='Own Process'" -ComputerName $Computer
	$Win32_Service = Get-WMIObject -Namespace root\CIMV2 -Class Win32_Service -Property Caption, PathName, Started, StartMode, StartName -Filter "ServiceType ='Share Process' Or ServiceType ='Own Process'" -ComputerName $Computer

	$Win32_Service | Sort-Object -Property Caption | ForEach-Object {

		$Service = New-Object -TypeName psobject -Property @{
			Caption = $_.Caption
			Started = $_.Started
			PathName = $_.PathName
			StartMode = $_.StartMode
			StartAs = $_.StartName
		}

		$ServicesInformation += $Service

		#		if ($Service.Caption -ieq "mssqlserver") {
		#			$script:SystemRole["SQL"] = $true
		#		}

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tService:"
		$Service.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	} 
	Write-Output $ServicesInformation

	Remove-Variable -Name ServicesInformation, Service, Win32_Service
}

function Get-ShareInformation([string]$Computer) {
	# Win32_Share: http://msdn.microsoft.com/en-us/library/aa394435(VS.85).aspx

	$ShareInformation = @()
	$Share = $null

	#$Win32_Share = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_Share -Property Name, Description, Path, Type -ComputerName $Computer
	$Win32_Share = Get-WMIObject -Namespace root\CIMV2 -Class Win32_Share -Property Name, Description, Path, Type -ComputerName $Computer

	$Win32_Share | Sort-Object -Property Name | ForEach-Object {

		$Share = New-Object -TypeName psobject -Property @{
			Name = $_.Name
			Description = $_.Description
			Path = $_.Path
			ShareType = Get-ShareType($_.Type)
		}

		$ShareInformation += $Share

		#		if ($_.Type = 0) {
		#			$script:SystemRole["File"] = $true
		#		} elseif ($_.Type = 1) {
		#			$script:SystemRole["Print"] = $true
		#		}

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tShare:"
		$Share.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}

	} 
	Write-Output $ShareInformation

	Remove-Variable -Name ShareInformation, Share, Win32_Share
}

function Get-SoundDeviceInformation([string]$Computer) {
	# Win32_SoundDevice: http://msdn.microsoft.com/en-us/library/aa394463(VS.85).aspx

	$SoundDeviceInformation = @()
	$SoundDevice = $null

	#$Win32_SoundDevice = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_SoundDevice -Property Name, Manufacturer -ComputerName $Computer
	$Win32_SoundDevice = Get-WMIObject -Namespace root\CIMV2 -Class Win32_SoundDevice -Property Name, Manufacturer -ComputerName $Computer

	$Win32_SoundDevice | ForEach-Object {

		$SoundDevice = New-Object -TypeName psobject -Property @{
			Name = $_.Name
			Manufacturer = $_.Manufacturer
		}

		$SoundDeviceInformation += $SoundDevice

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tSound Device:"
		$SoundDevice.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	} 
	Write-Output $SoundDeviceInformation

	Remove-Variable -Name SoundDeviceInformation, SoundDevice, Win32_SoundDevice
}

function Get-StartupCommandInformation([string]$Computer) {
	# Win32_StartupCommand: http://msdn.microsoft.com/en-us/library/aa394464(VS.85).aspx

	$StartupCommandInformation = @()
	$StartupCommand = $null

	#$Win32_StartupCommand = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_StartupCommand -Property Command, Name, User -ComputerName $Computer
	$Win32_StartupCommand = Get-WMIObject -Namespace root\CIMV2 -Class Win32_StartupCommand -Property Command, Name, User -ComputerName $Computer

	$Win32_StartupCommand | Sort-Object -Property User, Name | ForEach-Object {

		$StartupCommand = New-Object -TypeName psobject -Property @{
			Name = $_.Name
			Command = $_.Command
			User = $_.User
		}

		$StartupCommandInformation += $StartupCommand

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tStartup Command:"
		$StartupCommand.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	} 
	Write-Output $StartupCommandInformation

	Remove-Variable -Name StartupCommandInformation, StartupCommand, Win32_StartupCommand
}

function Test-ServiceInstallState([string]$Computer, [string]$ServiceName) {
	#if (@(Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_Service -Filter "name='$ServiceName'" -ComputerName $Computer -ErrorAction SilentlyContinue).Count -gt 0) {
	if (@(Get-WMIObject -Namespace root\CIMV2 -Class Win32_Service -Filter "name='$ServiceName'" -ComputerName $Computer -ErrorAction SilentlyContinue).Count -gt 0) {
		Write-Output $true
	} else {
		Write-Output $false
	}
}

# This is one that we can get more info from - but how useful?
function Get-SystemEnclosureInformation([string]$Computer) {
	# Win32_SystemEnclosure: http://msdn.microsoft.com/en-us/library/aa394474(VS.85).aspx

	$SystemEnclosureInformation = $null

	#$Win32_SystemEnclosure = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_SystemEnclosure -Property ChassisTypes -ComputerName $Computer
	$Win32_SystemEnclosure = Get-WMIObject -Namespace root\CIMV2 -Class Win32_SystemEnclosure -Property ChassisTypes -ComputerName $Computer

	$Win32_SystemEnclosure | ForEach-Object {
		$SystemEnclosureInformation = New-Object -TypeName psobject -Property @{
			ChassisType = Get-ChassisType($_.ChassisTypes)
		}
	} 

	$SystemEnclosureInformation.psobject.Properties | ForEach-Object {
		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t$($_.Name): $($_.Value)"
	}

	Write-Output $SystemEnclosureInformation

	Remove-Variable -Name SystemEnclosureInformation, Win32_SystemEnclosure

}

function Get-TapeDriveInformation([string]$Computer) {
	# Win32_TapeDrive: http://msdn.microsoft.com/en-us/library/aa394491(VS.85).aspx

	$TapeDriveInformation = @()
	$TapeDrive = $null

	#$Win32_TapeDrive = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_TapeDrive -Property Name, Description, Manufacturer -ComputerName $Computer
	$Win32_TapeDrive = Get-WMIObject -Namespace root\CIMV2 -Class Win32_TapeDrive -Property Name, Description, Manufacturer -ComputerName $Computer

	# Not sure why but this is bringing back one empty row on my laptop so I added where-object to filter out the empty row
	$Win32_TapeDrive | Where-Object {$_.Name} | Sort-Object -Property Name | ForEach-Object {

		$TapeDrive = New-Object -TypeName psobject -Property @{
			Name = $_.Name
			Description = $_.Description
			Manufacturer = $_.Manufacturer
		}

		$TapeDriveInformation += $TapeDrive

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tTape Drive:"
		$TapeDrive.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	} 
	Write-Output $TapeDriveInformation

	Remove-Variable -Name TapeDriveInformation, TapeDrive, Win32_TapeDrive
}

function Get-TimeZoneInformation([string]$Computer) {
	# Win32_TimeZone: http://msdn.microsoft.com/en-us/library/aa394498(VS.85).aspx

	$TimeZoneInformation = $null

	#$Win32_TimeZone = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_TimeZone -Property Description -ComputerName $Computer
	$Win32_TimeZone = Get-WMIObject -Namespace root\CIMV2 -Class Win32_TimeZone -Property Description -ComputerName $Computer

	$Win32_TimeZone | ForEach-Object {
		$TimeZoneInformation = New-Object -TypeName psobject -Property @{
			Description = $_.Description
		}
	}

	$TimeZoneInformation.psobject.Properties | ForEach-Object {
		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t$($_.Name): $($_.Value)"
	}

	Write-Output $TimeZoneInformation

	Remove-Variable -Name TimeZoneInformation, Win32_TimeZone
}

function Get-PatchInformationFromWMI([string]$Computer) {
	# Using Get-Hotfix instead of Win32_QuickFixEngineering because it takes care of date formatting
	# Win32_QuickFixEngineering: http://msdn.microsoft.com/en-us/library/aa394391(VS.85).aspx

	$PatchInformation = @()
	$Patch = $null

	$HotFix = Get-HotFix -ComputerName $Computer

	# Don't include patches that look like {ABC0D0F6-019E-4AC3-AD46-9C044E7B19F3}
	$HotFix | Where-Object { (-not (($_.HotFixID.Length -eq 38) -and ($_.HotFixID.StartsWith('{')) -and ($_.HotFixID.EndsWith('}')))) } | ForEach-Object {

		$Patch = New-Object -TypeName psobject -Property @{
			Caption = $_.Caption
			Description = $_.Description
			HotFixID = $_.HotFixID
			InstalledBy = $_.InstalledBy
			InstallDateUTC = if (($_.psbase.Properties['InstalledOn'].Value -ne [String]::Empty) -and ($_.InstalledOn -ne $null)) { $_.InstalledOn.ToUniversalTime() } else { $null }
		}

		$PatchInformation += $Patch

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tPatch:"
		$Patch.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}

	} 
	Write-Output $PatchInformation

	Remove-Variable -Name PatchInformation, Patch, HotFix
}

function Get-UserAccountInformation([string]$Computer, [System.Version]$OSVersion) {
	# Win32_UserAccount: http://msdn.microsoft.com/en-us/library/aa394507(VS.85).aspx

	$UserAccountInformation = @()
	$UserAccount = $null

	# Windows 2000 doesn't have the LocalAccount property
	if ($OSVersion.CompareTo($WindowsXP) -ge 0) {
		#$Win32_UserAccount = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_UserAccount -Property Description, Name, Disabled, FullName, Lockout, PasswordChangeable, PasswordExpires, PasswordRequired, PasswordRequired -Filter 'LocalAccount = True' -ComputerName $Computer
		$Win32_UserAccount = Get-WMIObject -Namespace root\CIMV2 -Class Win32_UserAccount -Property Description, Name, Disabled, FullName, Lockout, PasswordChangeable, PasswordExpires, PasswordRequired, PasswordRequired -Filter 'LocalAccount = True' -ComputerName $Computer
	} else {
		#$Win32_UserAccount = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_UserAccount -Property Description, Name, Disabled, FullName, Lockout, PasswordChangeable, PasswordExpires, PasswordRequired, PasswordRequired -Filter '__SERVER = Domain' -ComputerName $Computer
		$Win32_UserAccount = Get-WMIObject -Namespace root\CIMV2 -Class Win32_UserAccount -Property Description, Name, Disabled, FullName, Lockout, PasswordChangeable, PasswordExpires, PasswordRequired, PasswordRequired -Filter '__SERVER = Domain' -ComputerName $Computer
	}

	$Win32_UserAccount | ForEach-Object {

		$UserAccount = New-Object -TypeName psobject -Property @{
			Description = $_.Description
			Disabled = $_.Disabled
			FullName = $_.FullName
			Lockout = $_.Lockout
			PasswordChangeable = $_.PasswordChangeable
			PasswordExpires = $_.PasswordExpires
			PasswordRequired = $_.PasswordRequired
			UserName = $_.Name
		}

		$UserAccountInformation += $UserAccount

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tUser Account:"
		$UserAccount.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	} 
	Write-Output $UserAccountInformation

	Remove-Variable -Name UserAccountInformation, UserAccount, Win32_UserAccount
}

function Get-VideoControllerInformation([string]$Computer) {
	# Win32_VideoController: http://msdn.microsoft.com/en-us/library/aa394512(VS.85).aspx

	$VideoControllerInformation = @()
	$VideoController = $null

	#$Win32_VideoController = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_VideoController -Property AdapterCompatibility, AdapterRAM, Name -ComputerName $Computer
	$Win32_VideoController = Get-WMIObject -Namespace root\CIMV2 -Class Win32_VideoController -Property AdapterCompatibility, AdapterRAM, Name -ComputerName $Computer

	$Win32_VideoController | ForEach-Object {

		$VideoController = New-Object -TypeName psobject -Property @{
			AdapterCompatibility = $_.AdapterCompatibility
			AdapterRAMBytes = $_.AdapterRAM
			Name = $_.Name
		}

		$VideoControllerInformation += $VideoController

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tVideo Controller:"
		$VideoController.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}
	}
	Write-Output $VideoControllerInformation

	Remove-Variable -Name VideoControllerInformation, VideoController, Win32_VideoController
}

# Not fully implemented yet
function Get-ServerFeatureInformation([string]$Computer) {
	# Win32_ServerFeature: http://msdn.microsoft.com/en-us/library/cc280268(VS.85).aspx

	Write-WindowsMachineInformationLog -Message "[$Computer] Gathering information about Server Features" -MessageLevel Verbose

	#	$ServerFeature = $null
	#	$ServerFeatureGuid = $null

	#$Win32_ServerFeature = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_ServerFeature -Property ID, Name, ParentID -ComputerName $Computer
	$Win32_ServerFeature = Get-WMIObject -Namespace root\CIMV2 -Class Win32_ServerFeature -Property ID, Name, ParentID -ComputerName $Computer


	# Roles
	# Features
	# Role Services

	$ServerRoleIDs = (1,2,5,6,7,8,9,10,11,12,13,14,16,17,18,19,20,21)

	$Win32_ServerFeature | ForEach-Object {

		#		$ServerFeatureGuid = [guid]::NewGuid()
		#
		#		# Determine if the ID is a feature or a role
		#		if ($ServerRoleIDs -contains $_.ID) {
		#
		#			$script:ServerRole["$ServerFeatureGuid"] = @{
		#				"ID" = $_.ID ;
		#				"Name" = $_.Name ;
		#				"ParentID" = $_.ParentID ;
		#			}
		#
		#			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "	Role:"
		#			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "		ID: $($script:ServerRole[""$ServerFeatureGuid""].ID)"
		#			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "		Name: $($script:ServerRole[""$ServerFeatureGuid""].Name)"
		#			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "		ParentID: $($script:ServerRole[""$ServerFeatureGuid""].ParentID)"
		#
		#		
		#		} else {
		#			$script:ServerFeature["$ServerFeatureGuid"] = @{
		#				"ID" = $_.ID ;
		#				"Name" = $_.Name ;
		#				"ParentID" = $_.ParentID ;
		#			}
		#
		#			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "	Feature:"
		#			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "		ID: $($script:ServerFeature[""$ServerFeatureGuid""].ID)"
		#			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "		Name: $($script:ServerFeature[""$ServerFeatureGuid""].Name)"
		#			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "		ParentID: $($script:ServerFeature[""$ServerFeatureGuid""].ParentID)"
		#
		#		
		#		}

		#		# This needs to be updated for Windows 2008 since there are new roles and services
		#		switch ($_.ID) {
		#			10 { $script:SystemRole["DC"] = $true }
		#			12 { $script:SystemRole["DHCP"] = $true }
		#			13 { $script:SystemRole["DNS"] = $true }
		#			6 { $script:SystemRole["File"] = $true }
		#			184 { $script:SystemRole["FTP"] = $true }
		#			14 { $script:SystemRole["IAS"] = $true }
		#			3 { $script:SystemRole["Media"] = $true }
		#			7 { $script:SystemRole["Print"] = $true }
		#			207 { $script:SystemRole["RAS"] = $true }
		#			48 { $script:SystemRole["SMTP"] = $true }
		#			18 { $script:SystemRole["TS"] = $true }
		#			40 { $script:SystemRole["WINS"] = $true }
		#			2 { $script:SystemRole["WWW"] = $true }
		#		}

	} 

	Remove-Variable -Name Win32_ServerFeature, ServerRoleIDs

}


function Get-OptionalFeatureInformation([string]$Computer) {
	# Win32_OptionalFeature: http://msdn.microsoft.com/en-us/library/ee309383(v=VS.85).aspx
	# Minimum supported server: Windows Server 2008 R2
	# Minimum supported client: Windows 7

	$OptionalFeatureInformation = @()
	$OptionalFeature = $null
	$FeatureState = $null

	Write-WindowsMachineInformationLog -Message "[$Computer] Gathering information about Optional Features" -MessageLevel Verbose

	#$Win32_OptionalFeature = Get-WMIObjectWithTimeout -Namespace root\CIMV2 -Class Win32_OptionalFeature -Property Caption, Name, InstallState -ComputerName $Computer
	$Win32_OptionalFeature = Get-WMIObject -Namespace root\CIMV2 -Class Win32_OptionalFeature -Property Caption, Name, InstallState -ComputerName $Computer

	$Win32_OptionalFeature | ForEach-Object {

		$FeatureState = switch ($_.InstallState) {
			1 { 'Enabled' }
			2 { 'Disabled' }
			3 { 'Absent' }
			4 { 'Unknown' }
		}

		$OptionalFeature = New-Object -TypeName psobject -Property @{
			Name = $_.Name
			Caption = $_.Caption
			InstallState = $FeatureState
		}

		$OptionalFeatureInformation += $OptionalFeature

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tOptional Feature:"
		$OptionalFeature.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		}

	} 
	Write-Output $OptionalFeatureInformation

	Remove-Variable -Name OptionalFeatureInformation, OptionalFeature, FeatureState, Win32_OptionalFeature

}


## IIS Functions

function Get-IIS6WebInformation([string]$Computer) {
	# IISWebServerSetting: http://msdn.microsoft.com/en-us/library/ms524332.aspx
	# IISWebVirtualDirSetting: http://msdn.microsoft.com/en-us/library/ms525005(VS.90).aspx

	$IIS6WebInformation = @()

	Write-WindowsMachineInformationLog -Message "[$Computer] Start subroutine: GetIIS6Information" -MessageLevel Verbose

	$WebServer = $null
	$Setting = $null
	$Binding = $null
	$VirtualDir = $null
	$VirtualDirSetting = $null
	$IISWebServerSetting = $null
	$IIsWebVirtualDir = $null
	$IIsWebVirtualDirSetting = $null
	$IISWebServer = $null

	try {
		#$IISWebServer = Get-WMIObjectWithTimeout -Namespace root\MicrosoftIISv2 -Class IISWebServer -Authentication PacketPrivacy -ComputerName $Computer
		$IISWebServer = Get-WMIObject -Namespace root\MicrosoftIISv2 -Class IISWebServer -Authentication PacketPrivacy -ComputerName $Computer
	}
	catch {
		Write-WindowsMachineInformationLog -Message "`t[$Computer] Unable to gather information about IIS: Check that IISWebServer class is installed" -MessageLevel Warning
		return 
	}

	$IISWebServer | ForEach-Object {

		$WebServer = New-Object -TypeName psobject -Property @{
			Name = $_.Name
			Settings = @()
			VirtualDirectories = @()
		}

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tWeb Server:"
		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`tName: $($WebServer.Name)"
		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`tSettings:"

		$IISWebServerSetting = $_.GetRelated('IISWebServerSetting')

		$IISWebServerSetting | ForEach-Object {

			$Setting = New-Object -TypeName psobject -Property @{
				Name = $_.Name
				ServerComment = $_.ServerComment
				Bindings = @()
			}

			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`tName: $($Setting.Name)"
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`tServerComment: $($Setting.ServerComment)"
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`tBindings:"

			$_.ServerBindings | ForEach-Object {
				$Binding = New-Object -TypeName psobject -Property @{
					Hostname = $_.Hostname
					IP = $_.Ip
					Port = $_.Port
				}

				$Setting.Bindings += $Binding

				Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`t`tHostname: $($Binding.Hostname)"
				Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`t`tIP: $($Binding.IP)"
				Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`t`tPort: $($Binding.Port)"
			}

			$WebServer.Settings += $Setting
		}

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`tVirtual Directories:"

		$IIsWebVirtualDir = $_.GetRelated('IIsWebVirtualDir')

		$IIsWebVirtualDir | ForEach-Object {

			$VirtualDir = New-Object -TypeName psobject -Property @{
				Name = $_.Name
				Settings = @()
			}

			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`tName: $($VirtualDir.Name)"
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`tSettings:"

			$IIsWebVirtualDirSetting = $_.GetRelated('IISWebVirtualDirSetting')

			$IIsWebVirtualDirSetting | ForEach-Object {

				$VirtualDirSetting = New-Object -TypeName psobject -Property @{
					Name = $_.Name
					Path = $null
					PhysicalPath = $_.Path
				}

				$VirtualDir.Settings += $VirtualDirSetting

				Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`t`tName: $($VirtualDirSetting.Name)"
				Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`t`tPath: $($VirtualDirSetting.Path)"
				Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`t`tPhysicalPath: $($VirtualDirSetting.PhysicalPath)"

			}

			$WebServer.VirtualDirectories += $VirtualDir
		}

		$IIS6WebInformation += $WebServer
	}
	Write-Output $IIS6WebInformation

	Remove-Variable -Name IIS6WebInformation, WebServer, Setting, Binding, VirtualDir, VirtualDirSetting, IISWebServerSetting, IIsWebVirtualDir, IIsWebVirtualDirSetting, IISWebServer

}

function Get-IIS7WebInformation([string]$Computer) {
	# Significant architectural changes in IIS 7
	# IIS 7 WMI provider: http://msdn.microsoft.com/en-us/library/Aa347459
	# See:
	#	"Converting Metabase Properties to Configuration Settings [IIS 7]": http://msdn.microsoft.com/en-us/library/aa347565(v=VS.90).aspx
	#	"Mapping IIS 6.0 WMI Methods to IIS 7 WMI Methods": http://msdn.microsoft.com/en-us/library/bb398762(v=VS.90).aspx


	# Site: http://msdn.microsoft.com/en-us/library/ms689454(v=VS.90).aspx
	# Binding: http://msdn.microsoft.com/en-us/library/ms689497(v=VS.90).aspx
	#		BindingInformation - A nonempty read/write string value with three colon-delimited fields that specify binding information 
	#							for a Web site. The first field is a specific IP address or an asterisk (an asterisk specifies all unassigned 
	#							IP addresses). The second field is the port number; the default is 80. The third field is an optional host header. 

	$IIS7WebInformation = @()
	$Site = $null
	$Binding = $null
	$BindingInfo = $null
	$Application = $null
	$VirtualDir = $null
	$IIS7Application = $null
	$IIS7VirtualDir = $null
	$IIS7Site = $null

	Write-WindowsMachineInformationLog -Message "[$Computer] Start subroutine: GetIIS7Information" -MessageLevel Verbose

	try {
		#$IIS7Site = Get-WMIObjectWithTimeout -Namespace root\WebAdministration -Class Site -Property Name, Bindings -Authentication PacketPrivacy -ComputerName $Computer
		$IIS7Site = Get-WMIObject -Namespace root\WebAdministration -Class Site -Property Name, Bindings -Authentication PacketPrivacy -ComputerName $Computer
	}
	catch {
		Write-WindowsMachineInformationLog -Message "`t[$Computer] Unable to gather IIS information: Check that Site class is installed" -MessageLevel Warning
		return 
	}

	#	$uname = "DOMAIN\USERNAME"
	#	$pword = "password123"
	#	$spword = ConvertTo-SecureString -String $pword -AsPlainText -Force
	#	$Creds = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $uname, $spword

	$IIS7Site | ForEach-Object {

		$Site = New-Object -TypeName psobject -Property @{
			Name = $_.Name
			Settings = @()
			VirtualDirectories = @()
		}

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tWeb Server:"
		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`tName: $($Site.Name)"
		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`tSettings:"

		$Setting = New-Object -TypeName psobject -Property @{
			Name = $_.Name
			ServerComment = $_.Name
			Bindings = @()
		}

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`tName: $($Setting.Name)"
		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`tServerComment: $($Setting.ServerComment)"
		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`tBindings:"

		$_.Bindings | where {$_.Protocol -ieq 'http'} | ForEach-Object {

			$BindingInfo = $($_.BindingInformation).Split(':')

			$Binding = New-Object -TypeName psobject -Property @{
				Hostname = $BindingInfo[2]
				IP = $($BindingInfo[0]).Replace('*','')
				Port = $BindingInfo[1]
			}

			$Setting.Bindings += $Binding

			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`t`tHostname: $($Binding.Hostname)"
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`t`tIP: $($Binding.IP)"
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`t`tPort: $($Binding.Port)"

		}

		$Site.Settings += $Setting


		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`tVirtual Directories:" 

		$IIS7Application = $_.GetRelated('Application')

		$IIS7Application | ForEach-Object {

			$IIS7VirtualDir = $_.GetRelated('VirtualDirectory')

			$IIS7VirtualDir | ForEach-Object {

				$VirtualDir = New-Object -TypeName psobject -Property @{
					Name = $_.Name
					Settings = @()
				}

				Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`tName: $($VirtualDir.Name)"
				Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`tSettings:"

				$VirtualDirSetting = New-Object -TypeName psobject -Property @{
					Name = $_.SiteName
					Path = $_.Path
					PhysicalPath = $_.PhysicalPath
				}

				$VirtualDir.Settings += $VirtualDirSetting

				Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`t`tName: $($VirtualDirSetting.Name)"
				Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`t`tPath: $($VirtualDirSetting.Path)"
				Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t`t`tPhysicalPath: $($VirtualDirSetting.PhysicalPath)"

				$Site.VirtualDirectories += $VirtualDir 
			}

		}

		$IIS7WebInformation += $Site
	} 
	Write-Output $IIS7WebInformation

	Remove-Variable -Name IIS7WebInformation, Site, Binding, BindingInfo, Application, VirtualDir, IIS7Application, IIS7VirtualDir, IIS7Site

}



######################
# REGISTRY PROVIDER FUNCTIONS
######################


function Get-DomainInformation($RegistryProvider) {
	Write-WindowsMachineInformationLog -Message "[$Computer] Reading domain information" -MessageLevel Verbose
	$DomainInformation = $($RegistryProvider.GetStringValue($HKEY_LOCAL_MACHINE,'SYSTEM\CurrentControlSet\Services\Tcpip\Parameters','Domain')).sValue
	Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tDomain: $DomainInformation"
	Write-Output $DomainInformation
	Remove-Variable -Name DomainInformation
}

function Get-PrintSpoolLocation($RegistryProvider) {
	Write-WindowsMachineInformationLog -Message "[$Computer] Reading print spool location" -MessageLevel Verbose
	$PrintSpoolLocation = $($RegistryProvider.GetStringValue($HKEY_LOCAL_MACHINE,'SYSTEM\CurrentControlSet\Control\Print\Printers','DefaultSpoolDirectory')).sValue
	Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`rDefaultSpoolDirectory: $PrintSpoolLocation"
	Write-Output $PrintSpoolLocation
	Remove-Variable -Name PrintSpoolLocation
}

function Get-TerminalServerMode($RegistryProvider) {
	Write-WindowsMachineInformationLog -Message "[$Computer] Reading terminal server compatibility" -MessageLevel Verbose
	$TerminalServerMode = $($RegistryProvider.GetDWORDValue($HKEY_LOCAL_MACHINE,'SYSTEM\CurrentControlSet\Control\Terminal Server','TSAppCompat')).sValue
	Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tTSAppCompat: $TerminalServerMode"
	Write-Output $TerminalServerMode
	Remove-Variable -Name TerminalServerMode
}

function Get-ApplicationInformationFromRegistry($RegistryProvider) {

	$ApplicationInformation = @()
	$Application = $null
	$RegItemCollection = $null
	$IncludeProgram = $false
	$DisplayName = $null
	$InstallDate = $null

	$RegItemCollection = $RegistryProvider.EnumKey($HKEY_LOCAL_MACHINE,'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')

	$RegItemCollection.sNames | Where-Object { (($_.Length -ne 38) -or (($_.Length -eq 38) -and (-not $_.StartsWith('{')))) } | ForEach-Object {

		$IncludeProgram = $true

		$DisplayName = $($RegistryProvider.GetStringValue($HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($_)",'DisplayName')).sValue

		# Only include the program if it's got a display name
		if (-not $DisplayName) {

			$DisplayName = $($RegistryProvider.GetStringValue($HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($_)",'HiddenDisplayName')).sValue

			if (-not $DisplayName) {
				$IncludeProgram = $false
			}
		}


		# If it's not an MSI app and has a displayable name then include it
		if ($IncludeProgram -eq $true) {

			$Application = New-Object -TypeName psobject -Property @{
				ProductName = $DisplayName
				Vendor = $null
				Version = $null
				InstallDateUTC = $null
				InstallLocation = $null
				HelpURL = $null
				SupportURL = $null
				UpdateInfoURL = $null
				Source = 'Registry'
			}

			$Application.Vendor = $($RegistryProvider.GetStringValue($HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($_)",'Publisher')).sValue
			$Application.Version = $($RegistryProvider.GetStringValue($HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($_)",'DisplayVersion')).sValue
			$Application.InstallLocation = $($RegistryProvider.GetStringValue($HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($_)",'InstallLocation')).sValue
			$Application.HelpURL = $($RegistryProvider.GetStringValue($HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($_)",'HelpLink')).sValue
			$Application.SupportURL = $($RegistryProvider.GetStringValue($HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($_)",'URLInfoAbout')).sValue
			$Application.UpdateInfoURL = $($RegistryProvider.GetStringValue($HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($_)",'URLUpdateInfo')).sValue

			$InstallDate = $($RegistryProvider.GetStringValue($HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($_)",'InstallDate')).sValue
			if ($InstallDate -and 
				$InstallDate.Length -eq 8 -and
				$InstallDate -imatch '^\d{8}$'
			) {
				# If populated the format should be YYYYMMDD
				$Application.InstallDateUTC = $(Get-Date -Year $($InstallDate).Substring(0,4) -Month $($InstallDate).Substring(4,2) -Day $($InstallDate).Substring(6,2) -Hour 0 -Minute 0 -Second 0 )
			}

			$ApplicationInformation += $Application

			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tInstalled Application:"
			$Application.psobject.Properties | ForEach-Object {
				Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
			} 
		}

	}

	Write-Output $ApplicationInformation

	Remove-Variable -Name ApplicationInformation, Application, RegItemCollection, IncludeProgram, DisplayName, InstallDate

}

function Get-LastLoggedInUser($RegistryProvider) {

	$LastLoggedInUser = $null

	Write-WindowsMachineInformationLog -Message "[$Computer] Reading last user" -MessageLevel Verbose

	$LastUserDomain = $($RegistryProvider.GetStringValue($HKEY_LOCAL_MACHINE,'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon','DefaultDomainName')).sValue
	$LastUser = $($RegistryProvider.GetStringValue($HKEY_LOCAL_MACHINE,'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon','DefaultUserName')).sValue

	if ($LastUser.Length -gt 0) {
		$LastLoggedInUser = $LastUserDomain + '\' + $LastUser
	}

	Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tLastUser: $LastLoggedInUser"
	Write-Output $LastLoggedInUser

	Remove-Variable -Name LastLoggedInUser, LastUserDomain, LastUser
}

function Get-WindowsComponentInformation($RegistryProvider) {

	$WindowsComponentsInformation = @()
	$Component = $null
	$ComponentInformation = $null
	$RegValue = $null
	$RegItemCollection = $null

	Write-WindowsMachineInformationLog -Message "[$Computer] Reading Windows Components information" -MessageLevel Verbose

	$RegItemCollection = $RegistryProvider.EnumValues($HKEY_LOCAL_MACHINE,'SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\OC Manager\Subcomponents')

	foreach ($RegValue in $RegItemCollection.sNames) {

		# Double check that we have a value
		if ($RegValue) {

			$ComponentInformation = Get-ComponentInformation -ComponentName $RegValue

			$Component = New-Object -TypeName psobject -Property $ComponentInformation

			$WindowsComponentsInformation += $Component

			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tComponent:"
			$Component.psobject.Properties | ForEach-Object {
				Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
			} 
		}
	}
	Write-Output $WindowsComponentsInformation

	Remove-Variable -Name WindowsComponentsInformation, Component, ComponentInformation, RegValue, RegItemCollection
}

function Get-PatchInformationFromRegistry($RegistryProvider) {

	$PatchInformation = @()
	$Patch = $null
	$RegItemCollection = $null
	$InstallDate = $null

	$RegItemCollection = $RegistryProvider.EnumKey($HKEY_LOCAL_MACHINE,'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')

	$RegItemCollection.sNames | Where-Object { ($_ -imatch '^(KB)?\d+$') } | ForEach-Object {

		$Patch = New-Object -TypeName psobject -Property @{
			Caption = $null
			Description = $null
			HotFixID = $_
			InstalledBy = $null
			InstallDateUTC = $null
		}

		$Patch.Caption = $($RegistryProvider.GetStringValue($HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($_)",'HelpLink')).sValue
		$Patch.Description = $($RegistryProvider.GetStringValue($HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($_)",'ReleaseType')).sValue

		$InstallDate = $($RegistryProvider.GetStringValue($HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($_)",'InstallDate')).sValue

		if (($InstallDate) -and ($InstallDate.Length -eq 8)) {
			# If populated the format should be YYYYMMDD
			$Patch.InstallDateUTC = $(Get-Date -Year $($InstallDate).Substring(0,4) -Month $($InstallDate).Substring(4,2) -Day $($InstallDate).Substring(6,2) -Hour 0 -Minute 0 -Second 0 ).ToUniversalTime()
		}

		$PatchInformation += $Patch

		Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`tPatch:"
		$Patch.psobject.Properties | ForEach-Object {
			Write-WindowsMachineInformationLog -MessageLevel Debug -Message "`t`t$($_.Name): $($_.Value)"
		} 

	}

	Write-Output $PatchInformation

	Remove-Variable -Name PatchInformation, Patch, RegItemCollection, InstallDate

}


######################
# EXTERNAL EXE FUNCTIONS
######################

# This function is a total hackjob
# It should be easier to do this in PowerShell but it's not
# So we're stuck with having to Frankenstein together psExec and WMI to get it to work :-(
# For more information on secedit see http://technet.microsoft.com/en-us/library/cc742472(WS.10).aspx#BKMK_3

#function Get-WindowsUserRightsAssignment {
function Get-LocalSecurityPolicyInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[String]
		$Computer
		,
		[Parameter(Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[String[]]
		$Policy 
	)
	try {
		$RightsAssignment = $null
		$Account = $null
		$SecurityIdentifier = $null
		$NTAccount = $null

		Invoke-Command -ScriptBlock { &psexec \\$Computer cmd "/C `"secedit /export /cfg %TMP%\rights.inf /quiet && type %TMP%\rights.inf && del %TMP%\rights.inf`"" } | 
		Where-Object { $_ -ilike 'se* = *' } | 
		ForEach-Object {
			$RightsAssignment = $_.Split('=')

			if ($Policy -icontains $RightsAssignment[0].Trim()) {
				$Account = @()

				$RightsAssignment[1].Trim().Split(',') | ForEach-Object {
					if ($_ -ilike '`*S-*') {
						$SecurityIdentifier = New-Object -TypeName System.Security.Principal.SecurityIdentifier -ArgumentList $_.Replace('*','')
						$NTAccount = $SecurityIdentifier.Translate([System.Security.Principal.NTAccount])
						$Account += $NTAccount.Value
					} else {
						$Account += $_
					}
				}

				Write-Output (
					New-Object -TypeName psobject -Property @{
						Policy = $RightsAssignment[0].Trim()
						Account = $Account
					}
				) 
			}
		}
	}
	catch {
		Throw
	}
}



######################
# HELPER FUNCTIONS
######################

function Get-WMIObjectWithTimeout {
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
		$Query = "select " + [String]::Join(',',$Property) + " from " + $Class

		if ($Filter) {
			$Query = $Query + " where " + $Filter
		}

		$WmiSearcher.Options.Timeout = [TimeSpan]::FromSeconds($TimeoutSeconds)
		$WmiSearcher.Options.ReturnImmediately = $true
		$WmiSearcher.Scope.Path = "\\" + $ComputerName + "\" + $NameSpace
		$WmiSearcher.Query = $Query
		$WmiSearcher.Get() 
	}
	catch {
		throw
	}
	finally {
		Remove-Variable -Name WmiSearcher, Query
	}
}

# Wrapper function for logging
function Write-WindowsMachineInformationLog {
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
		throw
	}
}


function Get-IISInformation([string]$Computer, [System.Version]$OSVersion){
	if ($OSVersion.CompareTo($WindowsVista) -ge 0) { 
		Get-IIS7WebInformation -Computer $Computer 
	} else { 
		Get-IIS6WebInformation -Computer $Computer 
	}
}




######################
# MAIN FUNCTIONS
######################

function Get-WindowsMachineInformation {
	<#
	.SYNOPSIS
		Gets information about a Windows Operating system and the hardware its running on.

	.DESCRIPTION
		The Get-WindowsMachineInformation function uses Windows Management Instrumentation (WMI) to retrieve comprehensive information about a host running a Windows Operating System.
				
	.PARAMETER  ComputerName
		The computer name (or IP address) to collect information from.

	.PARAMETER  AdditionalData
		A comma delimited list of additional data to collect.
		
		Valid values include: AdditionalHardware, BIOS, DesktopSessions, EventLog, FullyQualifiedDomainName, InstalledApplications, InstalledPatches, IPRoutes, LastLoggedOnUser, LocalGroups, LocalUserAccounts, None, PowerPlans, Printers, PrintSpoolerLocation, Processes, ProductKeys, RegistrySize, Services, Shares, StartupCommands, WindowsComponents
		
		Use "None" to bypass collecting all additional information.
		
		The default value "None"		
		
	.PARAMETER  StopAtErrorCount
		The number of errors that can be encountered before this function calls it quits and returns.
		
		The default value is 3 errors.
		
	.EXAMPLE
		Get-WindowsMachineInformation -ComputerName $env:COMPUTERNAME
		
		Description
		-----------
		This will retrieve information about the local machine.
		
		All core and additional data will be collected.
		
		The function will return after encountering 3 errors.		

	.EXAMPLE
		Get-WindowsMachineInformation -ComputerName $env:COMPUTERNAME -AdditionalData None
		
		Description
		-----------
		This will retrieve information about the local machine.
		
		Do not collect any data beyond the core set of information.
		
		The function will return after encountering 3 errors.		

	.EXAMPLE
		Get-WindowsMachineInformation -ComputerName $env:COMPUTERNAME -StopAtErrorCount 5
		
		Description
		-----------
		This will retrieve information about the local machine.
		
		All core and additional data will be collected.
		
		The function will return after encountering 5 errors.	

	.OUTPUTS
		System.Management.Automation.PSObject

	.NOTES

#>
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$false)] 
		[alias('data')]
		[ValidateSet('AdditionalHardware','All','BIOS','DesktopSessions','EventLog','FullyQualifiedDomainName','InstalledApplications','InstalledPatches','IPRoutes', `
			'LastLoggedOnUser','LocalGroups','LocalUserAccounts','None','PowerPlans','Printers','PrintSpoolerLocation','Processes', `
			'ProductKeys','RegistrySize','Services','Shares','StartupCommands','WindowsComponents')]
		[string[]]
		$AdditionalData = @('None')
		,
		[Parameter(Mandatory=$false)] 
		[alias('computer')]
		[string]
		$ComputerName = $env:COMPUTERNAME
		,
		[Parameter(Mandatory=$false)] 
		[alias('errors')]
		[int]
		$StopAtErrorCount = 3
	)
	process {

		$ScanDateUTC = $(Get-Date).ToUniversalTime()
		$HasMicrosoftIIS = (Test-ServiceInstallState -Computer $ComputerName -ServiceName 'w3svc')
		$OSVersion = [System.Version]'0.0'
		$LastErrorActionPreference = $ErrorActionPreference
		$ErrorCount = 0 #$global:Error.Count
		$StdRegProv = $null

		$MachineInformation = New-Object -TypeName psobject -Property @{
			Hardware = New-Object -TypeName psobject -Property @{
				MotherboardControllerAndPort = New-Object -TypeName psobject -Property @{
					BIOS = @()
					PhysicalMemory = @()
					Processor = @()
					SoundDevice = @()
					SystemEnclosure = @()
				}
				NetworkAdapter = @()
				Printer = @()
				Storage = New-Object -TypeName psobject -Property @{
					CDROMDrive = @()
					DiskDrive = @()
					TapeDrive = @()
				}
				VideoAndMonitor = New-Object -TypeName psobject -Property @{
					VideoController = @()
				}
			}
			OperatingSystem = New-Object -TypeName psobject -Property @{
				Desktop = New-Object -TypeName psobject -Property @{
					TimeZone = @()
				}
				EventLog = @()
				Network = New-Object -TypeName psobject -Property @{
					IPV4RouteTable = @()
				}
				PageFile = @()
				Registry = @()
				Settings = New-Object -TypeName psobject -Property @{
					OperatingSystem = @()
					ComputerSystem = @()
					ComputerSystemProduct = @()
					PowerPlan = @()
					StartupCommands = @()
				}
				RunningProcesses = @()
				Services = @()
				Shares = @()
				Users = New-Object -TypeName psobject -Property @{
					LocalGroups = @()
					LocalUsers = @()
					DesktopSessions = @()
				}
			}
			Software = New-Object -TypeName psobject -Property @{
				InstalledApplications = @()
				Patches = @()
			}
			ScanDateUTC = $ScanDateUTC
			ScanErrorCount = 0
		}

		$ErrorActionPreference = 'Stop'

		# 
		If ($AdditionalData -icontains 'All') {
			$AdditionalData = @('AdditionalHardware','BIOS','DesktopSessions','EventLog','FullyQualifiedDomainName','InstalledApplications','InstalledPatches','IPRoutes', `
				'LastLoggedOnUser','LocalGroups','LocalUserAccounts','PowerPlans','Printers','PrintSpoolerLocation','Processes', `
				'ProductKeys','RegistrySize','Services','Shares','StartupCommands','WindowsComponents')
		}

		try {

			#########################
			# Gather core data first
			#########################

			# Get-SystemInformation
			#region
			while (-not $MachineInformation.OperatingSystem.Settings.ComputerSystem.Name) {
				try {
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering computer system information" -MessageLevel Verbose
					$MachineInformation.OperatingSystem.Settings.ComputerSystem = Get-SystemInformation -Computer $ComputerName
				} 
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering computer system information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw }
				}
			}
			#endregion

			# Get-ComputerSystemProductInformation
			#region
			try { 
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering computer system product information" -MessageLevel Verbose
				$MachineInformation.OperatingSystem.Settings.ComputerSystemProduct = Get-ComputerSystemProductInformation -Computer $ComputerName
			}
			catch {
				$ErrorRecord = $_.Exception.ErrorRecord
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering computer system product information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
				if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
			}
			#endregion

			# Get-OSInfo
			#region
			try { 
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering OS information" -MessageLevel Verbose
				$MachineInformation.OperatingSystem.Settings.OperatingSystem = Get-OSInfo -Computer $ComputerName
				$OSVersion = [System.Version]$MachineInformation.OperatingSystem.Settings.OperatingSystem.Version
			} 
			catch {
				$ErrorRecord = $_.Exception.ErrorRecord
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering OS information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
				if (++$ErrorCount -ge $StopAtErrorCount) { throw }
			}
			#endregion

			# Get-TimeZoneInformation
			#region
			try { 
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering time zone" -MessageLevel Verbose
				$MachineInformation.OperatingSystem.Desktop.TimeZone = Get-TimeZoneInformation -Computer $ComputerName 
			} 
			catch {
				$ErrorRecord = $_.Exception.ErrorRecord
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering time zone: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
				if (++$ErrorCount -ge $StopAtErrorCount) { throw }
			}
			#endregion

			# Get-PagefileInformation
			#region
			try { 
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering pagefile information" -MessageLevel Verbose
				$MachineInformation.OperatingSystem.PageFile = Get-PagefileInformation -Computer $ComputerName -OSVersion $OSVersion 
			} 
			catch {
				$ErrorRecord = $_.Exception.ErrorRecord
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering pagefile information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
				if (++$ErrorCount -ge $StopAtErrorCount) { throw }
			}
			#endregion

			# Get-NetworkAdapterConfig
			#region
			try { 
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering network adapter configuration" -MessageLevel Verbose
				$MachineInformation.Hardware.NetworkAdapter = Get-NetworkAdapterConfig -Computer $ComputerName 
			} 
			catch {
				$ErrorRecord = $_.Exception.ErrorRecord
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering network adapter configuration: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
				if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
			}
			#endregion

			# Get-PhysicalMemoryInformation
			#region
			try { 
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering information about physical memory" -MessageLevel Verbose
				$MachineInformation.Hardware.MotherboardControllerAndPort.PhysicalMemory = Get-PhysicalMemoryInformation -Computer $ComputerName 
			}
			catch {
				$ErrorRecord = $_.Exception.ErrorRecord
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering information about physical memory: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
				if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
			}
			#endregion

			# Get-ProcessorInformation
			#region
			try { 
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering processor information" -MessageLevel Verbose
				$MachineInformation.Hardware.MotherboardControllerAndPort.Processor = Get-ProcessorInformation -Computer $ComputerName -OSVersion $OSVersion
			}
			catch {
				$ErrorRecord = $_.Exception.ErrorRecord
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering processor information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
				if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
			}
			#endregion

			# Get-SystemEnclosureInformation
			#region
			try { 
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering system enclosure information" -MessageLevel Verbose
				$MachineInformation.Hardware.MotherboardControllerAndPort.SystemEnclosure = Get-SystemEnclosureInformation -Computer $ComputerName
			}
			catch {
				$ErrorRecord = $_.Exception.ErrorRecord
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering system enclosure information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
				if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
			}
			#endregion

			# Get-CDROMInformation
			#region
			try { 
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering CD-ROM Information" -MessageLevel Verbose
				$MachineInformation.Hardware.Storage.CDROMDrive = Get-CDROMInformation -Computer $ComputerName 
			} 
			catch {
				$ErrorRecord = $_.Exception.ErrorRecord
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering CD-ROM Information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
				if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
			}
			#endregion

			# Get-DiskInformation
			#region
			try { 
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering disk information" -MessageLevel Verbose
				$MachineInformation.Hardware.Storage.DiskDrive = Get-DiskInformation -Computer $ComputerName -OSVersion $OSVersion
			}
			catch {
				$ErrorRecord = $_.Exception.ErrorRecord
				Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering disk information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
				if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
			}
			#endregion

			#	# Roles & Features only supported in server class OSs for Windows 2008 and higher
			#	if (([int]$script:OS["Level"] -ge 60) -and ($script:OS["Name"] -ilike 'server')) { Get-ServerFeatureInformation -Computer $Computer }
			#
			#	# Optional Features only supported in Windows 7, Windows Server 2008 R2, and higher
			#	#if ([int]$script:OS["Level"] -ge 61) { Get-OptionalFeatureInformation -Computer $Computer }


			#########################
			# Gather optional data
			#########################

			# Get-BIOSInfo
			#region
			if ($AdditionalData -icontains 'BIOS') { 
				try { 
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering BIOS information" -MessageLevel Verbose
					$MachineInformation.Hardware.MotherboardControllerAndPort.BIOS = Get-BIOSInfo -Computer $ComputerName 
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering BIOS information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}
			}
			#endregion

			# Get-LocalGroupsInformation
			#region
			if ($AdditionalData -icontains 'LocalGroups') { 
				try { 
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering local groups and group members" -MessageLevel Verbose
					$MachineInformation.OperatingSystem.Users.LocalGroups = Get-LocalGroupsInformation -Computer $ComputerName -OSVersion $OSVersion
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering local groups and group members: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}
			}
			#endregion

			# Get-UserAccountInformation
			#region
			if ($AdditionalData -icontains 'LocalUserAccounts') { 
				try { 
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering local users" -MessageLevel Verbose
					$MachineInformation.OperatingSystem.Users.LocalUsers = Get-UserAccountInformation -Computer $ComputerName -OSVersion $OSVersion
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering local users: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}
			}
			#endregion

			# Get-DesktopSessionInformation
			#region
			if ($AdditionalData -icontains 'DesktopSessions') { 
				try {
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering information about logged on users" -MessageLevel Verbose
					$MachineInformation.OperatingSystem.Users.DesktopSessions = Get-DesktopSessionInformation -Computer $ComputerName 
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering information about logged on users: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}
			} 
			#endregion




			# Get-IpRouteInformation
			# IPV4 routes only supported in Windows XP, Windows Server 2003, and higher
			#region
			if (
				$AdditionalData -icontains 'IPRoutes' -and 
				$OSVersion.CompareTo($WindowsXP) -ge 0
			) { 
				try { 
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering IPv4 Route information" -MessageLevel Verbose
					$MachineInformation.OperatingSystem.Network.IPV4RouteTable = Get-IpRouteInformation -Computer $ComputerName 
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering IPv4 Route information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}
			}
			#endregion

			# Get-EventLogSettings
			#region
			if ($AdditionalData -icontains 'EventLog') { 
				try { 
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering event log settings" -MessageLevel Verbose
					$MachineInformation.OperatingSystem.EventLog = Get-EventLogSettings -Computer $ComputerName 
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering event log settings: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}
			}
			#endregion

			# Get-PowerPlanInformation
			# Power plans only supported in Windows 7, Windows Server 2008 R2, and higher
			#region
			if (
				$AdditionalData -icontains 'PowerPlans' -and 
				$OSVersion.CompareTo($WindowsServer2008R2) -ge 0
			) { 
				try { 
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering information about power plans" -MessageLevel Verbose
					$MachineInformation.OperatingSystem.Settings.PowerPlan = Get-PowerPlanInformation -Computer $ComputerName 
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering information about power plans: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}
			}
			#endregion

			# Get-PrinterInformation
			#region
			if ($AdditionalData -icontains 'Printers') { 
				try { 
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering printer information" -MessageLevel Verbose
					$MachineInformation.Hardware.Printer = Get-PrinterInformation -Computer $ComputerName 
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering printer information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}
			}
			#endregion

			# Get-ProcessInformation
			#region
			if ($AdditionalData -icontains 'Processes') { 
				try { 
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering process information" -MessageLevel Verbose
					$MachineInformation.OperatingSystem.RunningProcesses = Get-ProcessInformation -Computer $ComputerName 
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering process information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}
			}
			#endregion

			# Get-RegistrySizeInformation
			#region
			if ($AdditionalData -icontains 'RegistrySize') { 
				try { 
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering registry size information" -MessageLevel Verbose
					$MachineInformation.OperatingSystem.Registry = Get-RegistrySizeInformation -Computer $ComputerName 
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering registry size information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw }
				}
			}
			#endregion

			# Get-ServicesInformation
			#region
			if ($AdditionalData -icontains 'Services') { 
				try { 
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering information about services" -MessageLevel Verbose
					$MachineInformation.OperatingSystem.Services = Get-ServicesInformation -Computer $ComputerName 
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering information about services: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}
			}
			#endregion

			# Get-ShareInformation
			#region
			if ($AdditionalData -icontains 'Shares') { 
				try { 
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering information about shares" -MessageLevel Verbose
					$MachineInformation.OperatingSystem.Shares = Get-ShareInformation -Computer $ComputerName 
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering information about shares: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}
			}
			#endregion

			if ($AdditionalData -icontains 'AdditionalHardware') { 

				# Get-SoundDeviceInformation
				#region
				try { 
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering Sound Device information" -MessageLevel Verbose
					$MachineInformation.Hardware.MotherboardControllerAndPort.SoundDevice = Get-SoundDeviceInformation -Computer $ComputerName 
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering Sound Device information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}
				#endregion

				# Get-TapeDriveInformation
				#region
				try { 
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering TapeDrive information" -MessageLevel Verbose
					$MachineInformation.Hardware.Storage.TapeDrive = Get-TapeDriveInformation -Computer $ComputerName 
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering TapeDrive information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}
				#endregion

				# Get-VideoControllerInformation
				#region
				try { 
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering Video Controller information" -MessageLevel Verbose
					$MachineInformation.Hardware.VideoAndMonitor.VideoController = Get-VideoControllerInformation -Computer $ComputerName 
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering Video Controller information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}
				#endregion

			}

			# Get-StartupCommandInformation
			#region
			if ($AdditionalData -icontains 'StartupCommands') { 
				try { 
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering Startup Commands information" -MessageLevel Verbose
					$MachineInformation.OperatingSystem.Settings.StartupCommands = Get-StartupCommandInformation -Computer $ComputerName 
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering Startup Commands information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}
			}
			#endregion



			# Can't be instantiated in the "normal" way with a named class like this:
			#$StdRegProv = Get-WmiObject -Namespace root\DEFAULT -Class StdRegProv -ComputerName "127.0.0.1"

			# Instead, it has to be called via query
			# For info on using this WMI class see http://msdn.microsoft.com/en-us/library/windows/desktop/aa393664(v=vs.85).aspx
			$StdRegProv = Get-WmiObject -Namespace root\DEFAULT -Query "select * FROM meta_class WHERE __Class = 'StdRegProv'" -ComputerName $ComputerName


			# Get-ApplicationInformation
			#region
			if ($AdditionalData -icontains 'InstalledApplications') {
				try { 
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering application information from WMI" -MessageLevel Verbose
					$MachineInformation.Software.InstalledApplications += Get-ApplicationInformationFromWMI -Computer $ComputerName -OSVersion $OSVersion
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering application information from WMI: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}

				try {
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering application information from registry" -MessageLevel Verbose
					$MachineInformation.Software.InstalledApplications += Get-ApplicationInformationFromRegistry -RegistryProvider $StdRegProv
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering application information from registry: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}
			}
			#endregion


			# Get-PatchInformation
			#region
			if ($AdditionalData -icontains 'InstalledPatches') { 
				try {
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering information about patches from WMI" -MessageLevel Verbose
					$MachineInformation.Software.Patches = Get-PatchInformationFromWMI -Computer $ComputerName 
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering information about patches from WMI: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				}

				try { 
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Gathering information about patches from registry" -MessageLevel Verbose
					$MachineInformation.Software.Patches += Get-PatchInformationFromRegistry -RegistryProvider $StdRegProv 
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering information about patches from registry: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { throw } 
				} 
			}
			#endregion



			#########################
			# Gather registry information
			#########################

			#if ($RegistryOptions -icontains "PrintSpoolerLocation") { $null = Get-PrintSpoolLocation -RegistryProvider $StdRegProv }

			# Get-ApplicationInformationFromRegistry
			# region
			#if ($RegistryOptions -icontains 'WindowsComponents') { $null = Get-WindowsComponentInformation -RegistryProvider $StdRegProv }
			#if ($RegistryOptions -icontains 'FullyQualifiedDomainName') { $null = Get-DomainInformation -RegistryProvider $StdRegProv }
			#if ($RegistryOptions -icontains 'ProductKeys') { $null }
			#if ($RegistryOptions -icontains 'LastLoggedOnUser') { $null = Get-LastLoggedInUser -RegistryProvider $StdRegProv }

			#$null = Get-TerminalServerMode -RegistryProvider $StdRegProv

		}
		catch {
			# If we hit this point we've reached the max error threshold
			Write-WindowsMachineInformationLog -Message "[$ComputerName] Error gathering machine information - max error threshold reached ($StopAtErrorCount)" -MessageLevel Warning
		}
		finally {

			# Record the number of scan errors
			$MachineInformation.ScanErrorCount = $ErrorCount

			# Reset the $ErrorActionPreference and return the $MachineInformation object
			$ErrorActionPreference = $LastErrorActionPreference
			Write-Output $MachineInformation

			Remove-Variable -Name ScanDateUTC, HasMicrosoftIIS, OSVersion, LastErrorActionPreference, ErrorCount, StdRegProv, MachineInformation

			# Invoke the garbage collector
			[System.GC]::Collect()

		}

	}
}

<# 
TODO: 
- Add ability to pass alternate credentials
- Refactor IIS7 format? (Maybe make an IIS6 and IIS7 specific node?)
- Fix Server Role and Features
- Product Keys (I never used this anyways)
#>