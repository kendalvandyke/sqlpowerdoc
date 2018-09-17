<#
TODO:
- Store all DateTime values as UTC?
- Server level
	- AlwaysOn High Availability
#>

######################
# CONSTANTS
######################

# SQL Versions
# See http://social.technet.microsoft.com/wiki/contents/articles/783.sql-server-versions.aspx for version timeline
# Also see http://support.microsoft.com/kb/321185
# Also see http://sqlserverbuilds.blogspot.com/

New-Object -TypeName System.Version -ArgumentList '7.0.0.0' | New-Variable -Name SQLServer7 -Scope Script -Option Constant
New-Object -TypeName System.Version -ArgumentList '8.0.0.0' | New-Variable -Name SQLServer2000 -Scope Script -Option Constant
New-Object -TypeName System.Version -ArgumentList '9.0.0.0' | New-Variable -Name SQLServer2005 -Scope Script -Option Constant
New-Object -TypeName System.Version -ArgumentList '10.0.0.0' | New-Variable -Name SQLServer2008 -Scope Script -Option Constant
New-Object -TypeName System.Version -ArgumentList '10.50.0.0' | New-Variable -Name SQLServer2008R2 -Scope Script -Option Constant
New-Object -TypeName System.Version -ArgumentList '11.0.0.0' | New-Variable -Name SQLServer2012 -Scope Script -Option Constant

New-Variable -Name StandaloneDbEngine -Scope Script -Option Constant -Value 'Standalone'
New-Variable -Name AzureDbEngine -Scope Script -Option Constant -Value 'Windows Azure SQL Database'



# SMO Major Versions
# 9	:	SQL 2005
# 10:	SQL 2008 & 2008 R2 
# 11:	SQL 2012

# Used to compare against dates\times
[DateTime]'01/01/0001 12:00:00 AM' | New-Variable -Name SmoEpoch -Scope Script -Option Constant
[DateTime]'01/01/1900 12:00:00 AM' | New-Variable -Name DbccEpoch -Scope Script -Option Constant


# Privilege Constants
# See http://msdn.microsoft.com/en-us/library/windows/desktop/bb530716(v=vs.85).aspx
New-Variable -Name LockPagesInMemory -Scope Script -Value 'SeLockMemoryPrivilege' -Option Constant
New-Variable -Name PerformVolumeMaintenanceTasks -Scope Script -Value 'SeManageVolumePrivilege' -Option Constant

######################
# VARIABLES
######################
New-Variable -Name SmoMajorVersion -Scope Script -Value $null

$LookupAccountSidDefinition = @'
public const int NO_ERROR = 0;
public const int ERROR_INSUFFICIENT_BUFFER = 122;

public enum SID_NAME_USE : int {
    SidTypeUser = 1,
    SidTypeGroup,
    SidTypeDomain,
    SidTypeAlias,
    SidTypeWellKnownGroup,
    SidTypeDeletedAccount,
    SidTypeInvalid,
    SidTypeUnknown,
    SidTypeComputer
};

[DllImport("advapi32.dll", CharSet=CharSet.Auto, SetLastError = true)]
 public static extern bool LookupAccountSid(
     string lpSystemName,
     [MarshalAs(UnmanagedType.LPArray)] byte[] Sid,
     System.Text.StringBuilder lpName,
     ref uint cchName,
     System.Text.StringBuilder ReferencedDomainName,
     ref uint cchReferencedDomainName,
     out SID_NAME_USE peUse);
'@

if (-not ('WindowsAPI.Authorization' -as [Type])) {
    Add-Type -MemberDefinition $LookupAccountSidDefinition -Name Authorization -Namespace WindowsAPI -Using System.Text
}


###################
# PRIVATE FUNCTIONS
###################


# http://msdn.microsoft.com/en-us/library/windows/desktop/aa379166.aspx
function Resolve-AccountSid {
	[CmdletBinding()]
	[OutputType([int])]
	Param
	(
		[Parameter(Mandatory=$true,
			ValueFromPipelineByPropertyName=$true,
			Position=0)]
		[byte[]]
		$Sid,
		[Parameter(Mandatory=$false,
			ValueFromPipelineByPropertyName=$true,
			Position=1)]
		[string]
		$ComputerName = $null
	)
	Begin
	{
		$lpName = New-Object -TypeName System.Text.StringBuilder
		$cchName = $lpName.Capacity
		$ReferencedDomainName = New-Object -TypeName System.Text.StringBuilder
		$cchReferencedDomainName = $ReferencedDomainName.Capacity
		[WindowsAPI.Authorization+SID_NAME_USE]$peUse = 1
		$LastError = [WindowsAPI.Authorization]::NO_ERROR
	}
	Process
	{
		try {
			$Result = [WindowsAPI.Authorization]::LookupAccountSid($ComputerName, $Sid, $lpName, [ref]$cchName, $ReferencedDomainName, [ref]$cchReferencedDomainName, [ref]$peUse)

			if (-not $Result) {
				$LastError = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
				if ($LastError -eq [WindowsAPI.Authorization]::ERROR_INSUFFICIENT_BUFFER) {
					$lpName.EnsureCapacity($cchName) | Out-Null
					$ReferencedDomainName.EnsureCapacity($cchReferencedDomainName) | Out-Null
					$LastError = [WindowsAPI.Authorization]::NO_ERROR

					$Result = [WindowsAPI.Authorization]::LookupAccountSid($ComputerName, $Sid, $lpName, [ref]$cchName, $ReferencedDomainName, [ref]$cchReferencedDomainName, [ref]$peUse)

					if (-not $Result) {
						$LastError = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
					}
				}
			}

			Write-Output (
				New-Object -TypeName PSObject -Property @{
					IsResolved = $Result
					ReferencedDomainName = $ReferencedDomainName.ToString()
					AccountName = $lpName
					AccountType = [string]$(switch ($peUse) {
							$([WindowsAPI.Authorization+SID_NAME_USE]::SidTypeUser) { 'User' }
							$([WindowsAPI.Authorization+SID_NAME_USE]::SidTypeGroup) { 'Group' }
							$([WindowsAPI.Authorization+SID_NAME_USE]::SidTypeDomain) { 'Domain' }
							$([WindowsAPI.Authorization+SID_NAME_USE]::SidTypeAlias) { 'Alias' }
							$([WindowsAPI.Authorization+SID_NAME_USE]::SidTypeWellKnownGroup) { 'Well Known Group' }
							$([WindowsAPI.Authorization+SID_NAME_USE]::SidTypeDeletedAccount) { 'Deleted Account' }
							$([WindowsAPI.Authorization+SID_NAME_USE]::SidTypeInvalid) { 'Invalid' }
							$([WindowsAPI.Authorization+SID_NAME_USE]::SidTypeUnknown) { 'Unknown' }
							$([WindowsAPI.Authorization+SID_NAME_USE]::SidTypeComputer) { 'Computer' }
							default { $null }
						})
					ErrorCode = $LastError

				}
			)

		}
		catch {
			# Do nothing for now...maybe something later?
		} 
	}
	End
	{
		Remove-Variable -Name lpName, cchName, ReferencedDomainName, cchReferencedDomainName, peUse, LastError
	}
}



function Get-TablePropertyQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT tbl.name AS [Name], tbl.object_id AS [ID], tbl.schema_id AS [SchemaID], tbl.create_date AS [CreateDate], tbl.modify_date AS [DateLastModified], ISNULL(stbl.name, N'') AS [Owner], CAST(CASE WHEN tbl.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(tbl.schema_id) AS [Schema], CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(OBJECTPROPERTY(tbl.object_id, N'HasAfterTrigger') AS BIT) AS [HasAfterTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasInsertTrigger') AS BIT) AS [HasInsertTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasDeleteTrigger') AS BIT) AS [HasDeleteTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasInsteadOfTrigger') AS BIT) AS [HasInsteadOfTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasUpdateTrigger') AS BIT) AS [HasUpdateTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'IsIndexed') AS BIT) AS [HasIndex], CAST(OBJECTPROPERTY(tbl.object_id, N'IsIndexable') AS BIT) AS [IsIndexable], CAST(CASE idx.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [HasClusteredIndex], tbl.uses_ansi_nulls AS [AnsiNullsStatus], CAST(OBJECTPROPERTY(tbl.object_id, N'IsQuotedIdentOn') AS BIT) AS [QuotedIdentifierStatus], CAST(0 AS BIT) AS [FakeSystemTable], ISNULL(ftc.distribution_name, N'') AS [DistributionName], ISNULL(c.name, N'') AS [FederationColumnName], ISNULL(ftc.column_id, 0) AS [FederationColumnID], tbl.is_replicated AS [Replicated], tbl.lock_escalation AS [LockEscalation] FROM sys.tables AS tbl LEFT OUTER JOIN sys.database_principals AS stbl ON stbl.principal_id = ISNULL(tbl.principal_id, ( OBJECTPROPERTY(tbl.object_id, 'OwnerId') )) INNER JOIN sys.indexes AS idx ON idx.object_id = tbl.object_id AND idx.index_id < 2 LEFT OUTER JOIN sys.federated_table_columns AS ftc ON ( ftc.object_id = tbl.object_id ) LEFT OUTER JOIN sys.columns AS c ON ( c.object_id = tbl.object_id ) AND ( c.column_id = ftc.column_id )
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT tbl.name AS [Name], tbl.object_id AS [ID], tbl.schema_id AS [SchemaID], tbl.create_date AS [CreateDate], tbl.modify_date AS [DateLastModified], ISNULL(stbl.name, N'') AS [Owner], CAST(CASE WHEN tbl.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(tbl.schema_id) AS [Schema], CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(OBJECTPROPERTY(tbl.object_id, N'HasAfterTrigger') AS BIT) AS [HasAfterTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasInsertTrigger') AS BIT) AS [HasInsertTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasDeleteTrigger') AS BIT) AS [HasDeleteTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasInsteadOfTrigger') AS BIT) AS [HasInsteadOfTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasUpdateTrigger') AS BIT) AS [HasUpdateTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'IsIndexed') AS BIT) AS [HasIndex], CAST(OBJECTPROPERTY(tbl.object_id, N'IsIndexable') AS BIT) AS [IsIndexable], CAST(CASE idx.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [HasClusteredIndex], tbl.uses_ansi_nulls AS [AnsiNullsStatus], CAST(OBJECTPROPERTY(tbl.object_id, N'IsQuotedIdentOn') AS BIT) AS [QuotedIdentifierStatus], CAST(0 AS BIT) AS [FakeSystemTable], ISNULL(dstext.name, N'') AS [TextFileGroup], ISNULL((SELECT SUM(spart.rows) FROM sys.partitions spart WHERE spart.object_id = tbl.object_id AND spart.index_id < 2), 0) AS [RowCount], tbl.is_replicated AS [Replicated], tbl.lock_escalation AS [LockEscalation], CAST(CASE WHEN ctt.object_id IS NULL THEN 0 ELSE 1 END AS BIT) AS [ChangeTrackingEnabled], CAST(ISNULL(ctt.is_track_columns_updated_on, 0) AS BIT) AS [TrackColumnsUpdatedEnabled], tbl.is_filetable AS [IsFileTable], ISNULL(ft.directory_name, N'') AS [FileTableDirectoryName], ISNULL(ft.filename_collation_name, N'') AS [FileTableNameColumnCollation], CAST(ISNULL(ft.is_enabled, 0) AS BIT) AS [FileTableNamespaceEnabled], CASE WHEN 'FG' = dsidx.type THEN dsidx.name ELSE N'' END AS [FileGroup], CASE WHEN 'PS' = dsidx.type THEN dsidx.name ELSE N'' END AS [PartitionScheme], CAST(CASE WHEN 'PS' = dsidx.type THEN 1 ELSE 0 END AS BIT) AS [IsPartitioned], CASE WHEN 'FD' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamFileGroup], CASE WHEN 'PS' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamPartitionScheme] FROM sys.tables AS tbl LEFT OUTER JOIN sys.database_principals AS stbl ON stbl.principal_id = ISNULL(tbl.principal_id, ( OBJECTPROPERTY(tbl.object_id, 'OwnerId') )) INNER JOIN sys.indexes AS idx ON idx.object_id = tbl.object_id AND idx.index_id < 2 LEFT OUTER JOIN sys.data_spaces AS dstext ON tbl.lob_data_space_id = dstext.data_space_id LEFT OUTER JOIN sys.change_tracking_tables AS ctt ON ctt.object_id = tbl.object_id LEFT OUTER JOIN sys.filetables AS ft ON ft.object_id = tbl.object_id LEFT OUTER JOIN sys.data_spaces AS dsidx ON dsidx.data_space_id = idx.data_space_id LEFT OUTER JOIN sys.tables AS t ON t.object_id = idx.object_id LEFT OUTER JOIN sys.data_spaces AS dstbl ON dstbl.data_space_id = t.Filestream_data_space_id AND idx.index_id < 2
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT tbl.name AS [Name], tbl.object_id AS [ID], tbl.schema_id AS [SchemaID], tbl.create_date AS [CreateDate], tbl.modify_date AS [DateLastModified], ISNULL(stbl.name, N'') AS [Owner], CAST(CASE WHEN tbl.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(tbl.schema_id) AS [Schema], CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(OBJECTPROPERTY(tbl.object_id, N'HasAfterTrigger') AS BIT) AS [HasAfterTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasInsertTrigger') AS BIT) AS [HasInsertTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasDeleteTrigger') AS BIT) AS [HasDeleteTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasInsteadOfTrigger') AS BIT) AS [HasInsteadOfTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasUpdateTrigger') AS BIT) AS [HasUpdateTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'IsIndexed') AS BIT) AS [HasIndex], CAST(OBJECTPROPERTY(tbl.object_id, N'IsIndexable') AS BIT) AS [IsIndexable], CAST(CASE idx.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [HasClusteredIndex], tbl.uses_ansi_nulls AS [AnsiNullsStatus], CAST(OBJECTPROPERTY(tbl.object_id, N'IsQuotedIdentOn') AS BIT) AS [QuotedIdentifierStatus], CAST(0 AS BIT) AS [FakeSystemTable], ISNULL(dstext.name, N'') AS [TextFileGroup], ISNULL((SELECT SUM(spart.rows) FROM sys.partitions spart WHERE spart.object_id = tbl.object_id AND spart.index_id < 2), 0) AS [RowCount], tbl.is_replicated AS [Replicated], tbl.lock_escalation AS [LockEscalation], CAST(CASE WHEN ctt.object_id IS NULL THEN 0 ELSE 1 END AS BIT) AS [ChangeTrackingEnabled], CAST(ISNULL(ctt.is_track_columns_updated_on, 0) AS BIT) AS [TrackColumnsUpdatedEnabled], CASE WHEN 'FG' = dsidx.type THEN dsidx.name ELSE N'' END AS [FileGroup], CASE WHEN 'PS' = dsidx.type THEN dsidx.name ELSE N'' END AS [PartitionScheme], CAST(CASE WHEN 'PS' = dsidx.type THEN 1 ELSE 0 END AS BIT) AS [IsPartitioned], CASE WHEN 'FD' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamFileGroup], CASE WHEN 'PS' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamPartitionScheme] FROM sys.tables AS tbl LEFT OUTER JOIN sys.database_principals AS stbl ON stbl.principal_id = ISNULL(tbl.principal_id, ( OBJECTPROPERTY(tbl.object_id, 'OwnerId') )) INNER JOIN sys.indexes AS idx ON idx.object_id = tbl.object_id AND idx.index_id < 2 LEFT OUTER JOIN sys.data_spaces AS dstext ON tbl.lob_data_space_id = dstext.data_space_id LEFT OUTER JOIN sys.change_tracking_tables AS ctt ON ctt.object_id = tbl.object_id LEFT OUTER JOIN sys.data_spaces AS dsidx ON dsidx.data_space_id = idx.data_space_id LEFT OUTER JOIN sys.tables AS t ON t.object_id = idx.object_id LEFT OUTER JOIN sys.data_spaces AS dstbl ON dstbl.data_space_id = t.Filestream_data_space_id AND idx.index_id < 2 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT tbl.name AS [Name], tbl.object_id AS [ID], tbl.schema_id AS [SchemaID], tbl.create_date AS [CreateDate], tbl.modify_date AS [DateLastModified], ISNULL(stbl.name, N'') AS [Owner], CAST(CASE WHEN tbl.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(tbl.schema_id) AS [Schema], CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(OBJECTPROPERTY(tbl.object_id, N'HasAfterTrigger') AS BIT) AS [HasAfterTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasInsertTrigger') AS BIT) AS [HasInsertTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasDeleteTrigger') AS BIT) AS [HasDeleteTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasInsteadOfTrigger') AS BIT) AS [HasInsteadOfTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasUpdateTrigger') AS BIT) AS [HasUpdateTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'IsIndexed') AS BIT) AS [HasIndex], CAST(OBJECTPROPERTY(tbl.object_id, N'IsIndexable') AS BIT) AS [IsIndexable], CAST(CASE idx.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [HasClusteredIndex], tbl.uses_ansi_nulls AS [AnsiNullsStatus], CAST(OBJECTPROPERTY(tbl.object_id, N'IsQuotedIdentOn') AS BIT) AS [QuotedIdentifierStatus], CAST(0 AS BIT) AS [FakeSystemTable], ISNULL(dstext.name, N'') AS [TextFileGroup], ISNULL((SELECT SUM(spart.rows) FROM sys.partitions spart WHERE spart.object_id = tbl.object_id AND spart.index_id < 2), 0) AS [RowCount], tbl.is_replicated AS [Replicated], tbl.lock_escalation AS [LockEscalation], CAST(CASE WHEN ctt.object_id IS NULL THEN 0 ELSE 1 END AS BIT) AS [ChangeTrackingEnabled], CAST(ISNULL(ctt.is_track_columns_updated_on, 0) AS BIT) AS [TrackColumnsUpdatedEnabled], CASE WHEN 'FG' = dsidx.type THEN dsidx.name ELSE N'' END AS [FileGroup], CASE WHEN 'PS' = dsidx.type THEN dsidx.name ELSE N'' END AS [PartitionScheme], CAST(CASE WHEN 'PS' = dsidx.type THEN 1 ELSE 0 END AS BIT) AS [IsPartitioned], CASE WHEN 'FD' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamFileGroup], CASE WHEN 'PS' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamPartitionScheme] FROM sys.tables AS tbl LEFT OUTER JOIN sys.database_principals AS stbl ON stbl.principal_id = ISNULL(tbl.principal_id, ( OBJECTPROPERTY(tbl.object_id, 'OwnerId') )) INNER JOIN sys.indexes AS idx ON idx.object_id = tbl.object_id AND idx.index_id < 2 LEFT OUTER JOIN sys.data_spaces AS dstext ON tbl.lob_data_space_id = dstext.data_space_id LEFT OUTER JOIN sys.change_tracking_tables AS ctt ON ctt.object_id = tbl.object_id LEFT OUTER JOIN sys.data_spaces AS dsidx ON dsidx.data_space_id = idx.data_space_id LEFT OUTER JOIN sys.tables AS t ON t.object_id = idx.object_id LEFT OUTER JOIN sys.data_spaces AS dstbl ON dstbl.data_space_id = t.Filestream_data_space_id AND idx.index_id < 2 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT tbl.name AS [Name], tbl.object_id AS [ID], tbl.schema_id AS [SchemaID], tbl.create_date AS [CreateDate], tbl.modify_date AS [DateLastModified], ISNULL(stbl.name, N'') AS [Owner], CAST(CASE WHEN tbl.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(tbl.schema_id) AS [Schema], CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(OBJECTPROPERTY(tbl.object_id, N'HasAfterTrigger') AS BIT) AS [HasAfterTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasInsertTrigger') AS BIT) AS [HasInsertTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasDeleteTrigger') AS BIT) AS [HasDeleteTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasInsteadOfTrigger') AS BIT) AS [HasInsteadOfTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'HasUpdateTrigger') AS BIT) AS [HasUpdateTrigger], CAST(OBJECTPROPERTY(tbl.object_id, N'IsIndexed') AS BIT) AS [HasIndex], CAST(OBJECTPROPERTY(tbl.object_id, N'IsIndexable') AS BIT) AS [IsIndexable], CAST(CASE idx.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [HasClusteredIndex], tbl.uses_ansi_nulls AS [AnsiNullsStatus], CAST(OBJECTPROPERTY(tbl.object_id, N'IsQuotedIdentOn') AS BIT) AS [QuotedIdentifierStatus], CAST(0 AS BIT) AS [FakeSystemTable], ISNULL(dstext.name, N'') AS [TextFileGroup], ISNULL((SELECT SUM(spart.rows) FROM sys.partitions spart WHERE spart.object_id = tbl.object_id AND spart.index_id < 2), 0) AS [RowCount], tbl.is_replicated AS [Replicated], CASE WHEN 'FG' = dsidx.type THEN dsidx.name ELSE N'' END AS [FileGroup], CASE WHEN 'PS' = dsidx.type THEN dsidx.name ELSE N'' END AS [PartitionScheme], CAST(CASE WHEN 'PS' = dsidx.type THEN 1 ELSE 0 END AS BIT) AS [IsPartitioned] FROM sys.tables AS tbl LEFT OUTER JOIN sys.database_principals AS stbl ON stbl.principal_id = ISNULL(tbl.principal_id, ( OBJECTPROPERTY(tbl.object_id, 'OwnerId') )) INNER JOIN sys.indexes AS idx ON idx.object_id = tbl.object_id AND idx.index_id < 2 LEFT OUTER JOIN sys.data_spaces AS dstext ON tbl.lob_data_space_id = dstext.data_space_id LEFT OUTER JOIN sys.data_spaces AS dsidx ON dsidx.data_space_id = idx.data_space_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(tbl.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(tbl.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT tbl.name AS [Name], tbl.id AS [ID], tbl.crdate AS [CreateDate], stbl.name AS [Schema], stbl.uid AS [SchemaID], stbl.name AS [Owner], CAST(CASE WHEN ( OBJECTPROPERTY(tbl.id, N'IsMSShipped') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(tbl.id, N'IsSystemTable') THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(OBJECTPROPERTY(tbl.id, N'HasAfterTrigger') AS BIT) AS [HasAfterTrigger], CAST(OBJECTPROPERTY(tbl.id, N'HasInsertTrigger') AS BIT) AS [HasInsertTrigger], CAST(OBJECTPROPERTY(tbl.id, N'HasDeleteTrigger') AS BIT) AS [HasDeleteTrigger], CAST(OBJECTPROPERTY(tbl.id, N'HasInsteadOfTrigger') AS BIT) AS [HasInsteadOfTrigger], CAST(OBJECTPROPERTY(tbl.id, N'HasUpdateTrigger') AS BIT) AS [HasUpdateTrigger], CAST(OBJECTPROPERTY(tbl.id, N'IsIndexed') AS BIT) AS [HasIndex], CAST(OBJECTPROPERTY(tbl.id, N'IsIndexable') AS BIT) AS [IsIndexable], CAST(CASE WHEN ( OBJECTPROPERTY(tbl.id, N'tableisfake') = 1 ) THEN 1 ELSE 0 END AS BIT) AS [FakeSystemTable], CAST(CASE idx.indid WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [HasClusteredIndex], ISNULL((SELECT TOP 1 s.groupname FROM dbo.sysfilegroups s, dbo.sysindexes i WHERE i.id = tbl.id AND i.indid IN ( 0, 1 ) AND i.groupid = s.groupid), N'') AS [TextFileGroup], CAST(tbl.replinfo AS BIT) AS [Replicated], CAST(OBJECTPROPERTY(tbl.id, N'IsAnsiNullsOn') AS BIT) AS [AnsiNullsStatus], CAST(OBJECTPROPERTY(tbl.id, N'IsQuotedIdentOn') AS BIT) AS [QuotedIdentifierStatus], CAST(idx.rowcnt AS BIGINT) AS [RowCount], fgidx.groupname AS [FileGroup] FROM dbo.sysobjects AS tbl INNER JOIN sysusers AS stbl ON stbl.uid = tbl.uid INNER JOIN dbo.sysindexes AS idx ON idx.id = tbl.id AND idx.indid < 2 LEFT OUTER JOIN dbo.sysfilegroups AS fgidx ON fgidx.groupid = idx.groupid WHERE ( tbl.type = 'U' OR tbl.type = 'S' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-TablePhysicalPartitionQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$null
	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
	SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], p.partition_number AS [PartitionNumber], prv.value AS [RightBoundaryValue], fg.name AS [FileGroupName], CAST(pf.boundary_value_on_right AS INT) AS [RangeType], CAST(p.rows AS FLOAT) AS [RowCount], p.data_compression AS [DataCompression] FROM sys.tables AS tbl INNER JOIN sys.indexes AS idx ON idx.object_id = tbl.object_id AND idx.index_id < 2 INNER JOIN sys.partitions AS p ON p.object_id = CAST(tbl.object_id AS INT) AND p.index_id = idx.index_id LEFT OUTER JOIN sys.destination_data_spaces AS dds ON dds.partition_scheme_id = idx.data_space_id AND dds.destination_id = p.partition_number LEFT OUTER JOIN sys.partition_schemes AS ps ON ps.data_space_id = idx.data_space_id LEFT OUTER JOIN sys.partition_range_values AS prv ON prv.boundary_id = p.partition_number AND prv.function_id = ps.function_id LEFT OUTER JOIN sys.filegroups AS fg ON fg.data_space_id = dds.data_space_id OR fg.data_space_id = idx.data_space_id LEFT OUTER JOIN sys.partition_functions AS pf ON pf.function_id = prv.function_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
	SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], p.partition_number AS [PartitionNumber], prv.value AS [RightBoundaryValue], fg.name AS [FileGroupName], CAST(pf.boundary_value_on_right AS int) AS [RangeType], CAST(p.rows AS float) AS [RowCount], p.data_compression AS [DataCompression] FROM sys.tables AS tbl INNER JOIN sys.indexes AS idx ON idx.object_id = tbl.object_id and idx.index_id < 2 INNER JOIN sys.partitions AS p ON p.object_id = CAST(tbl.object_id AS int) AND p.index_id = idx.index_id LEFT OUTER JOIN sys.destination_data_spaces AS dds ON dds.partition_scheme_id = idx.data_space_id and dds.destination_id = p.partition_number LEFT OUTER JOIN sys.partition_schemes AS ps ON ps.data_space_id = idx.data_space_id LEFT OUTER JOIN sys.partition_range_values AS prv ON prv.boundary_id = p.partition_number and prv.function_id = ps.function_id LEFT OUTER JOIN sys.filegroups AS fg ON fg.data_space_id = dds.data_space_id or fg.data_space_id = idx.data_space_id LEFT OUTER JOIN sys.partition_functions AS pf ON pf.function_id = prv.function_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
	SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], p.partition_number AS [PartitionNumber], prv.value AS [RightBoundaryValue], fg.name AS [FileGroupName], CAST(pf.boundary_value_on_right AS INT) AS [RangeType], CAST(p.rows AS FLOAT) AS [RowCount], p.data_compression AS [DataCompression] FROM sys.tables AS tbl INNER JOIN sys.indexes AS idx ON idx.object_id = tbl.object_id AND idx.index_id < 2 INNER JOIN sys.partitions AS p ON p.object_id = CAST(tbl.object_id AS INT) AND p.index_id = idx.index_id LEFT OUTER JOIN sys.destination_data_spaces AS dds ON dds.partition_scheme_id = idx.data_space_id AND dds.destination_id = p.partition_number LEFT OUTER JOIN sys.partition_schemes AS ps ON ps.data_space_id = idx.data_space_id LEFT OUTER JOIN sys.partition_range_values AS prv ON prv.boundary_id = p.partition_number AND prv.function_id = ps.function_id LEFT OUTER JOIN sys.filegroups AS fg ON fg.data_space_id = dds.data_space_id OR fg.data_space_id = idx.data_space_id LEFT OUTER JOIN sys.partition_functions AS pf ON pf.function_id = prv.function_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
	SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], p.partition_number AS [PartitionNumber], prv.value AS [RightBoundaryValue], fg.name AS [FileGroupName], CAST(pf.boundary_value_on_right AS INT) AS [RangeType], CAST(p.rows AS FLOAT) AS [RowCount] FROM sys.tables AS tbl INNER JOIN sys.indexes AS idx ON idx.object_id = tbl.object_id AND idx.index_id < 2 INNER JOIN sys.partitions AS p ON p.object_id = CAST(tbl.object_id AS INT) AND p.index_id = idx.index_id LEFT OUTER JOIN sys.destination_data_spaces AS dds ON dds.partition_scheme_id = idx.data_space_id AND dds.destination_id = p.partition_number LEFT OUTER JOIN sys.partition_schemes AS ps ON ps.data_space_id = idx.data_space_id LEFT OUTER JOIN sys.partition_range_values AS prv ON prv.boundary_id = p.partition_number AND prv.function_id = ps.function_id LEFT OUTER JOIN sys.filegroups AS fg ON fg.data_space_id = dds.data_space_id OR fg.data_space_id = idx.data_space_id LEFT OUTER JOIN sys.partition_functions AS pf ON pf.function_id = prv.function_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$null
		}
	}
}

function Get-TablePartitionSchemeParameterQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$null
	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], CAST(ic.partition_ordinal AS INT) AS [ID], c.name AS [Name] FROM sys.tables AS tbl INNER JOIN sys.indexes AS idx ON idx.object_id = tbl.object_id AND idx.index_id < 2 INNER JOIN sys.index_columns ic ON ( ic.partition_ordinal > 0 ) AND ( ic.index_id = idx.index_id AND ic.object_id = CAST(tbl.object_id AS INT) ) INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$null
		}
	}
}

function Get-TableCheckQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], ~cstr.is_not_trusted AS [IsChecked], ~cstr.is_disabled AS [IsEnabled], cstr.is_not_for_replication AS [NotForReplication], cstr.definition AS [Definition] FROM sys.tables AS tbl INNER JOIN sys.check_constraints AS cstr ON cstr.parent_object_id = tbl.object_id
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], ~cstr.is_not_trusted AS [IsChecked], ~cstr.is_disabled AS [IsEnabled], cstr.is_not_for_replication AS [NotForReplication], CAST(CASE WHEN filetableobj.object_id IS NULL THEN 0 ELSE 1 END AS BIT) AS [IsFileTableDefined], cstr.definition AS [Definition] FROM sys.tables AS tbl INNER JOIN sys.check_constraints AS cstr ON cstr.parent_object_id = tbl.object_id LEFT OUTER JOIN sys.filetable_system_defined_objects AS filetableobj ON filetableobj.object_id = cstr.object_id
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], ~cstr.is_not_trusted AS [IsChecked], ~cstr.is_disabled AS [IsEnabled], cstr.is_not_for_replication AS [NotForReplication], cstr.definition AS [Definition] FROM sys.tables AS tbl INNER JOIN sys.check_constraints AS cstr ON cstr.parent_object_id = tbl.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], ~cstr.is_not_trusted AS [IsChecked], ~cstr.is_disabled AS [IsEnabled], cstr.is_not_for_replication AS [NotForReplication], cstr.definition AS [Definition] FROM sys.tables AS tbl INNER JOIN sys.check_constraints AS cstr ON cstr.parent_object_id = tbl.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], ~cstr.is_not_trusted AS [IsChecked], ~cstr.is_disabled AS [IsEnabled], cstr.is_not_for_replication AS [NotForReplication], cstr.definition AS [Definition] FROM sys.tables AS tbl INNER JOIN sys.check_constraints AS cstr ON cstr.parent_object_id = tbl.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(tbl.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(tbl.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT tbl.id AS [TableID], stbl.uid AS [SchemaID], cstr.name AS [Name], cstr.id AS [ID], cstr.crdate AS [CreateDate], CAST(cstr.status & 4 AS BIT) AS [IsSystemNamed], CAST(1 - ISNULL(OBJECTPROPERTY(cstr.id, N'CnstIsNotTrusted'), 0) AS BIT) AS [IsChecked], CAST(1 - ISNULL(OBJECTPROPERTY(cstr.id, N'CnstIsDisabled'), 0) AS BIT) AS [IsEnabled], CAST(ISNULL(OBJECTPROPERTY(cstr.id, N'CnstIsNotRepl'), 0) AS BIT) AS [NotForReplication], c.text AS [Definition] FROM dbo.sysobjects AS tbl INNER JOIN sysusers AS stbl ON stbl.uid = tbl.uid INNER JOIN dbo.sysobjects AS cstr ON ( cstr.type = 'C' ) AND ( cstr.parent_obj = tbl.id ) LEFT OUTER JOIN dbo.syscomments c ON c.id = cstr.id AND CASE WHEN c.number > 1 THEN c.number ELSE 0 END = 0 WHERE ( tbl.type = 'U' OR tbl.type = 'S' ) 
"@ + $SystemObjectWhereClause
		}
	}

}

function Get-TableColumnQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], clmns.name AS [Name], clmns.column_id AS [ID], clmns.is_nullable AS [Nullable], clmns.is_computed AS [Computed], CAST(ISNULL(cik.index_column_id, 0) AS BIT) AS [InPrimaryKey], clmns.is_ansi_padded AS [AnsiPaddingStatus], CAST(clmns.is_rowguidcol AS BIT) AS [RowGuidCol], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsDeterministic'), 0) AS BIT) AS [IsDeterministic], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsPrecise'), 0) AS BIT) AS [IsPrecise], CAST(ISNULL(cc.is_persisted, 0) AS BIT) AS [IsPersisted], ISNULL(clmns.collation_name, N'') AS [Collation], CAST(ISNULL((SELECT TOP 1 1 FROM sys.foreign_key_columns AS colfk WHERE colfk.parent_column_id = clmns.column_id AND colfk.parent_object_id = clmns.object_id), 0) AS BIT) AS [IsForeignKey], clmns.is_identity AS [Identity], CAST(ISNULL(ic.seed_value, 0) AS BIGINT) AS [IdentitySeed], CAST(ISNULL(ic.increment_value, 0) AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE SCHEMA_NAME(d.schema_id) END ) AS [DefaultSchema], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE SCHEMA_NAME(r.schema_id) END ) AS [RuleSchema], ISNULL(ic.is_not_for_replication, 0) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], CAST(clmns.is_filestream AS BIT) AS [IsFileStream], CAST(clmns.is_sparse AS BIT) AS [IsSparse], CAST(clmns.is_column_set AS BIT) AS [IsColumnSet], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND clmns.max_length <> -1 THEN clmns.max_length / 2 ELSE clmns.max_length END AS INT) AS [Length], CAST(clmns.precision AS INT) AS [NumericPrecision], CAST(clmns.scale AS INT) AS [NumericScale], ISNULL(xscclmns.name, N'') AS [XmlSchemaNamespace], ISNULL(s2clmns.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE clmns.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], ISNULL(cc.definition, N'') AS [ComputedText] FROM sys.tables AS tbl INNER JOIN sys.all_columns AS clmns ON clmns.object_id = tbl.object_id LEFT OUTER JOIN sys.indexes AS ik ON ik.object_id = clmns.object_id AND 1 = ik.is_primary_key LEFT OUTER JOIN sys.index_columns AS cik ON cik.index_id = ik.index_id AND cik.column_id = clmns.column_id AND cik.object_id = clmns.object_id AND 0 = cik.is_included_column LEFT OUTER JOIN sys.computed_columns AS cc ON cc.object_id = clmns.object_id AND cc.column_id = clmns.column_id LEFT OUTER JOIN sys.identity_columns AS ic ON ic.object_id = clmns.object_id AND ic.column_id = clmns.column_id LEFT OUTER JOIN sys.objects AS d ON d.object_id = clmns.default_object_id LEFT OUTER JOIN sys.objects AS r ON r.object_id = clmns.rule_object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = clmns.user_type_id LEFT OUTER JOIN sys.schemas AS s1clmns ON s1clmns.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = clmns.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = clmns.system_type_id ) AND ( baset.user_type_id = clmns.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscclmns ON xscclmns.xml_collection_id = clmns.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2clmns ON s2clmns.schema_id = xscclmns.schema_id
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		} 

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], clmns.name AS [Name], clmns.column_id AS [ID], clmns.is_nullable AS [Nullable], clmns.is_computed AS [Computed], CAST(ISNULL(cik.index_column_id, 0) AS BIT) AS [InPrimaryKey], clmns.is_ansi_padded AS [AnsiPaddingStatus], CAST(clmns.is_rowguidcol AS BIT) AS [RowGuidCol], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsDeterministic'), 0) AS BIT) AS [IsDeterministic], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsPrecise'), 0) AS BIT) AS [IsPrecise], CAST(ISNULL(cc.is_persisted, 0) AS BIT) AS [IsPersisted], ISNULL(clmns.collation_name, N'') AS [Collation], CAST(ISNULL((SELECT TOP 1 1 FROM sys.foreign_key_columns AS colfk WHERE colfk.parent_column_id = clmns.column_id AND colfk.parent_object_id = clmns.object_id), 0) AS BIT) AS [IsForeignKey], clmns.is_identity AS [Identity], CAST(ISNULL(ic.seed_value, 0) AS BIGINT) AS [IdentitySeed], CAST(ISNULL(ic.increment_value, 0) AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE SCHEMA_NAME(d.schema_id) END ) AS [DefaultSchema], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE SCHEMA_NAME(r.schema_id) END ) AS [RuleSchema], ISNULL(ic.is_not_for_replication, 0) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'StatisticalSemantics') AS INT) AS [StatisticalSemantics], CAST(clmns.is_filestream AS BIT) AS [IsFileStream], CAST(clmns.is_sparse AS BIT) AS [IsSparse], CAST(clmns.is_column_set AS BIT) AS [IsColumnSet], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND clmns.max_length <> -1 THEN clmns.max_length / 2 ELSE clmns.max_length END AS INT) AS [Length], CAST(clmns.precision AS INT) AS [NumericPrecision], CAST(clmns.scale AS INT) AS [NumericScale], ISNULL(xscclmns.name, N'') AS [XmlSchemaNamespace], ISNULL(s2clmns.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE clmns.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], ISNULL(cc.definition, N'') AS [ComputedText] FROM sys.tables AS tbl INNER JOIN sys.all_columns AS clmns ON clmns.object_id = tbl.object_id LEFT OUTER JOIN sys.indexes AS ik ON ik.object_id = clmns.object_id AND 1 = ik.is_primary_key LEFT OUTER JOIN sys.index_columns AS cik ON cik.index_id = ik.index_id AND cik.column_id = clmns.column_id AND cik.object_id = clmns.object_id AND 0 = cik.is_included_column LEFT OUTER JOIN sys.computed_columns AS cc ON cc.object_id = clmns.object_id AND cc.column_id = clmns.column_id LEFT OUTER JOIN sys.identity_columns AS ic ON ic.object_id = clmns.object_id AND ic.column_id = clmns.column_id LEFT OUTER JOIN sys.objects AS d ON d.object_id = clmns.default_object_id LEFT OUTER JOIN sys.objects AS r ON r.object_id = clmns.rule_object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = clmns.user_type_id LEFT OUTER JOIN sys.schemas AS s1clmns ON s1clmns.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = clmns.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = clmns.system_type_id ) AND ( baset.user_type_id = clmns.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscclmns ON xscclmns.xml_collection_id = clmns.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2clmns ON s2clmns.schema_id = xscclmns.schema_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], clmns.name AS [Name], clmns.column_id AS [ID], clmns.is_nullable AS [Nullable], clmns.is_computed AS [Computed], CAST(ISNULL(cik.index_column_id, 0) AS BIT) AS [InPrimaryKey], clmns.is_ansi_padded AS [AnsiPaddingStatus], CAST(clmns.is_rowguidcol AS BIT) AS [RowGuidCol], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsDeterministic'), 0) AS BIT) AS [IsDeterministic], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsPrecise'), 0) AS BIT) AS [IsPrecise], CAST(ISNULL(cc.is_persisted, 0) AS BIT) AS [IsPersisted], ISNULL(clmns.collation_name, N'') AS [Collation], CAST(ISNULL((SELECT TOP 1 1 FROM sys.foreign_key_columns AS colfk WHERE colfk.parent_column_id = clmns.column_id AND colfk.parent_object_id = clmns.object_id), 0) AS BIT) AS [IsForeignKey], clmns.is_identity AS [Identity], CAST(ISNULL(ic.seed_value, 0) AS BIGINT) AS [IdentitySeed], CAST(ISNULL(ic.increment_value, 0) AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE SCHEMA_NAME(d.schema_id) END ) AS [DefaultSchema], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE SCHEMA_NAME(r.schema_id) END ) AS [RuleSchema], ISNULL(ic.is_not_for_replication, 0) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], CAST(clmns.is_filestream AS BIT) AS [IsFileStream], CAST(clmns.is_sparse AS BIT) AS [IsSparse], CAST(clmns.is_column_set AS BIT) AS [IsColumnSet], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND clmns.max_length <> -1 THEN clmns.max_length / 2 ELSE clmns.max_length END AS INT) AS [Length], CAST(clmns.precision AS INT) AS [NumericPrecision], CAST(clmns.scale AS INT) AS [NumericScale], ISNULL(xscclmns.name, N'') AS [XmlSchemaNamespace], ISNULL(s2clmns.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE clmns.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], ISNULL(cc.definition, N'') AS [ComputedText] FROM sys.tables AS tbl INNER JOIN sys.all_columns AS clmns ON clmns.object_id = tbl.object_id LEFT OUTER JOIN sys.indexes AS ik ON ik.object_id = clmns.object_id AND 1 = ik.is_primary_key LEFT OUTER JOIN sys.index_columns AS cik ON cik.index_id = ik.index_id AND cik.column_id = clmns.column_id AND cik.object_id = clmns.object_id AND 0 = cik.is_included_column LEFT OUTER JOIN sys.computed_columns AS cc ON cc.object_id = clmns.object_id AND cc.column_id = clmns.column_id LEFT OUTER JOIN sys.identity_columns AS ic ON ic.object_id = clmns.object_id AND ic.column_id = clmns.column_id LEFT OUTER JOIN sys.objects AS d ON d.object_id = clmns.default_object_id LEFT OUTER JOIN sys.objects AS r ON r.object_id = clmns.rule_object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = clmns.user_type_id LEFT OUTER JOIN sys.schemas AS s1clmns ON s1clmns.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = clmns.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = clmns.system_type_id ) AND ( baset.user_type_id = clmns.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscclmns ON xscclmns.xml_collection_id = clmns.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2clmns ON s2clmns.schema_id = xscclmns.schema_id
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], clmns.name AS [Name], clmns.column_id AS [ID], clmns.is_nullable AS [Nullable], clmns.is_computed AS [Computed], CAST(ISNULL(cik.index_column_id, 0) AS BIT) AS [InPrimaryKey], clmns.is_ansi_padded AS [AnsiPaddingStatus], CAST(clmns.is_rowguidcol AS BIT) AS [RowGuidCol], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsDeterministic'), 0) AS BIT) AS [IsDeterministic], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsPrecise'), 0) AS BIT) AS [IsPrecise], CAST(ISNULL(cc.is_persisted, 0) AS BIT) AS [IsPersisted], ISNULL(clmns.collation_name, N'') AS [Collation], CAST(ISNULL((SELECT TOP 1 1 FROM sys.foreign_key_columns AS colfk WHERE colfk.parent_column_id = clmns.column_id AND colfk.parent_object_id = clmns.object_id), 0) AS BIT) AS [IsForeignKey], clmns.is_identity AS [Identity], CAST(ISNULL(ic.seed_value, 0) AS BIGINT) AS [IdentitySeed], CAST(ISNULL(ic.increment_value, 0) AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE SCHEMA_NAME(d.schema_id) END ) AS [DefaultSchema], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE SCHEMA_NAME(r.schema_id) END ) AS [RuleSchema], ISNULL(ic.is_not_for_replication, 0) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], CAST(clmns.is_filestream AS BIT) AS [IsFileStream], CAST(clmns.is_sparse AS BIT) AS [IsSparse], CAST(clmns.is_column_set AS BIT) AS [IsColumnSet], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND clmns.max_length <> -1 THEN clmns.max_length / 2 ELSE clmns.max_length END AS INT) AS [Length], CAST(clmns.precision AS INT) AS [NumericPrecision], CAST(clmns.scale AS INT) AS [NumericScale], ISNULL(xscclmns.name, N'') AS [XmlSchemaNamespace], ISNULL(s2clmns.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE clmns.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], ISNULL(cc.definition, N'') AS [ComputedText] FROM sys.tables AS tbl INNER JOIN sys.all_columns AS clmns ON clmns.object_id = tbl.object_id LEFT OUTER JOIN sys.indexes AS ik ON ik.object_id = clmns.object_id AND 1 = ik.is_primary_key LEFT OUTER JOIN sys.index_columns AS cik ON cik.index_id = ik.index_id AND cik.column_id = clmns.column_id AND cik.object_id = clmns.object_id AND 0 = cik.is_included_column LEFT OUTER JOIN sys.computed_columns AS cc ON cc.object_id = clmns.object_id AND cc.column_id = clmns.column_id LEFT OUTER JOIN sys.identity_columns AS ic ON ic.object_id = clmns.object_id AND ic.column_id = clmns.column_id LEFT OUTER JOIN sys.objects AS d ON d.object_id = clmns.default_object_id LEFT OUTER JOIN sys.objects AS r ON r.object_id = clmns.rule_object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = clmns.user_type_id LEFT OUTER JOIN sys.schemas AS s1clmns ON s1clmns.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = clmns.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = clmns.system_type_id ) AND ( baset.user_type_id = clmns.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscclmns ON xscclmns.xml_collection_id = clmns.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2clmns ON s2clmns.schema_id = xscclmns.schema_id
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], clmns.name AS [Name], clmns.column_id AS [ID], clmns.is_nullable AS [Nullable], clmns.is_computed AS [Computed], CAST(ISNULL(cik.index_column_id, 0) AS BIT) AS [InPrimaryKey], clmns.is_ansi_padded AS [AnsiPaddingStatus], CAST(clmns.is_rowguidcol AS BIT) AS [RowGuidCol], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsDeterministic'), 0) AS BIT) AS [IsDeterministic], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsPrecise'), 0) AS BIT) AS [IsPrecise], CAST(ISNULL(cc.is_persisted, 0) AS BIT) AS [IsPersisted], ISNULL(clmns.collation_name, N'') AS [Collation], CAST(ISNULL((SELECT TOP 1 1 FROM sys.foreign_key_columns AS colfk WHERE colfk.parent_column_id = clmns.column_id AND colfk.parent_object_id = clmns.object_id), 0) AS BIT) AS [IsForeignKey], clmns.is_identity AS [Identity], CAST(ISNULL(ic.seed_value, 0) AS BIGINT) AS [IdentitySeed], CAST(ISNULL(ic.increment_value, 0) AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE SCHEMA_NAME(d.schema_id) END ) AS [DefaultSchema], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE SCHEMA_NAME(r.schema_id) END ) AS [RuleSchema], ISNULL(ic.is_not_for_replication, 0) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND clmns.max_length <> -1 THEN clmns.max_length / 2 ELSE clmns.max_length END AS INT) AS [Length], CAST(clmns.precision AS INT) AS [NumericPrecision], CAST(clmns.scale AS INT) AS [NumericScale], ISNULL(xscclmns.name, N'') AS [XmlSchemaNamespace], ISNULL(s2clmns.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE clmns.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], ISNULL(cc.definition, N'') AS [ComputedText] FROM sys.tables AS tbl INNER JOIN sys.all_columns AS clmns ON clmns.object_id = tbl.object_id LEFT OUTER JOIN sys.indexes AS ik ON ik.object_id = clmns.object_id AND 1 = ik.is_primary_key LEFT OUTER JOIN sys.index_columns AS cik ON cik.index_id = ik.index_id AND cik.column_id = clmns.column_id AND cik.object_id = clmns.object_id AND 0 = cik.is_included_column LEFT OUTER JOIN sys.computed_columns AS cc ON cc.object_id = clmns.object_id AND cc.column_id = clmns.column_id LEFT OUTER JOIN sys.identity_columns AS ic ON ic.object_id = clmns.object_id AND ic.column_id = clmns.column_id LEFT OUTER JOIN sys.objects AS d ON d.object_id = clmns.default_object_id LEFT OUTER JOIN sys.objects AS r ON r.object_id = clmns.rule_object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = clmns.user_type_id LEFT OUTER JOIN sys.schemas AS s1clmns ON s1clmns.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = clmns.system_type_id AND baset.user_type_id = baset.system_type_id ) LEFT OUTER JOIN sys.xml_schema_collections AS xscclmns ON xscclmns.xml_collection_id = clmns.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2clmns ON s2clmns.schema_id = xscclmns.schema_id
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(tbl.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(tbl.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT tbl.id AS [TableID], stbl.uid AS [SchemaID], clmns.name AS [Name], CAST(clmns.colid AS INT) AS [ID], CAST(clmns.isnullable AS BIT) AS [Nullable], CAST(clmns.iscomputed AS BIT) AS [Computed], CAST(ISNULL(cik.colid, 0) AS BIT) AS [InPrimaryKey], CAST(ISNULL(COLUMNPROPERTY(clmns.id, clmns.name, N'UsesAnsiTrim'), 0) AS BIT) AS [AnsiPaddingStatus], CAST(clmns.colstat & 2 AS BIT) AS [RowGuidCol], CAST(clmns.colstat & 8 AS BIT) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], CAST(COLUMNPROPERTY(clmns.id, clmns.name, N'IsIdentity') AS BIT) AS [Identity], CAST(ISNULL((SELECT TOP 1 1 FROM dbo.sysforeignkeys AS colfk WHERE colfk.fkey = clmns.colid AND colfk.fkeyid = clmns.id), 0) AS BIT) AS [IsForeignKey], ISNULL(clmns.collation, N'') AS [Collation], CAST(CASE COLUMNPROPERTY(clmns.id, clmns.name, N'IsIdentity') WHEN 1 THEN IDENT_SEED(QUOTENAME(stbl.name) + '.' + QUOTENAME(tbl.name)) ELSE 0 END AS BIGINT) AS [IdentitySeed], CAST(CASE COLUMNPROPERTY(clmns.id, clmns.name, N'IsIdentity') WHEN 1 THEN IDENT_INCR(QUOTENAME(stbl.name) + '.' + QUOTENAME(tbl.name)) ELSE 0 END AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.cdefault = 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.cdefault = 0 THEN N'' ELSE USER_NAME(d.uid) END ) AS [DefaultSchema], ( CASE WHEN clmns.domain = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.domain = 0 THEN N'' ELSE USER_NAME(r.uid) END ) AS [RuleSchema], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'char', N'varchar', N'binary', N'varbinary', N'nchar', N'nvarchar' ) THEN clmns.prec ELSE clmns.length END AS INT) AS [Length], CAST(clmns.xprec AS INT) AS [NumericPrecision], CAST(clmns.xscale AS INT) AS [NumericScale], comt.text AS [ComputedText] FROM dbo.sysobjects AS tbl INNER JOIN sysusers AS stbl ON stbl.uid = tbl.uid INNER JOIN dbo.syscolumns AS clmns ON clmns.id = tbl.id LEFT OUTER JOIN dbo.syscomments comt ON comt.number = CAST(clmns.colid AS INT) AND comt.id = clmns.id LEFT OUTER JOIN dbo.sysindexes AS ik ON ik.id = clmns.id AND 0 != ik.status & 0x0800 LEFT OUTER JOIN dbo.sysindexkeys AS cik ON cik.indid = ik.indid AND cik.colid = clmns.colid AND cik.id = clmns.id LEFT OUTER JOIN dbo.sysobjects AS d ON d.id = clmns.cdefault AND 0 = d.category & 0x0800 LEFT OUTER JOIN dbo.sysobjects AS r ON r.id = clmns.domain LEFT OUTER JOIN systypes AS usrt ON usrt.xusertype = clmns.xusertype LEFT OUTER JOIN sysusers AS s1clmns ON s1clmns.uid = usrt.uid LEFT OUTER JOIN systypes AS baset ON baset.xusertype = clmns.xtype AND baset.xusertype = baset.xtype WHERE ( tbl.type = 'U' OR tbl.type = 'S' )
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-TableColumnDefaultConstraintQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], clmns.column_id AS [ColumnID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], cstr.definition AS [Text] FROM sys.tables AS tbl INNER JOIN sys.all_columns AS clmns ON clmns.object_id = tbl.object_id INNER JOIN sys.default_constraints AS cstr ON cstr.object_id = clmns.default_object_id 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], clmns.column_id AS [ColumnID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], CAST(CASE WHEN filetableobj.object_id IS NULL THEN 0 ELSE 1 END AS BIT) AS [IsFileTableDefined], cstr.definition AS [Text] FROM sys.tables AS tbl INNER JOIN sys.all_columns AS clmns ON clmns.object_id = tbl.object_id INNER JOIN sys.default_constraints AS cstr ON cstr.object_id = clmns.default_object_id LEFT OUTER JOIN sys.filetable_system_defined_objects AS filetableobj ON filetableobj.object_id = cstr.object_id
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], clmns.column_id AS [ColumnID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], cstr.definition AS [Text] FROM sys.tables AS tbl INNER JOIN sys.all_columns AS clmns ON clmns.object_id = tbl.object_id INNER JOIN sys.default_constraints AS cstr ON cstr.object_id = clmns.default_object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], clmns.column_id AS [ColumnID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], cstr.definition AS [Text] FROM sys.tables AS tbl INNER JOIN sys.all_columns AS clmns ON clmns.object_id = tbl.object_id INNER JOIN sys.default_constraints AS cstr ON cstr.object_id = clmns.default_object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], clmns.object_id AS [ColumnID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], cstr.definition AS [Text] FROM sys.tables AS tbl INNER JOIN sys.all_columns AS clmns ON clmns.object_id = tbl.object_id INNER JOIN sys.default_constraints AS cstr ON cstr.object_id = clmns.default_object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(tbl.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(tbl.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT tbl.id AS [TableID], stbl.uid AS [SchemaID], c.colid AS [ColumnID], cstr.name AS [Name], cstr.id AS [ID], cstr.crdate AS [CreateDate], CAST(cstr.status & 4 AS BIT) AS [IsSystemNamed], c.text AS [Text] FROM dbo.sysobjects AS tbl INNER JOIN sysusers AS stbl ON stbl.uid = tbl.uid INNER JOIN dbo.syscolumns AS clmns ON clmns.id = tbl.id INNER JOIN dbo.sysobjects AS cstr ON ( cstr.xtype = 'D' AND cstr.name NOT LIKE N'#%%' AND 0 != CONVERT(BIT, cstr.category & 0x0800) ) AND ( cstr.id = clmns.cdefault ) LEFT OUTER JOIN dbo.syscomments c ON c.id = cstr.id AND CASE WHEN c.number > 1 THEN c.number ELSE 0 END = 0 WHERE ( tbl.type = 'U' OR tbl.type = 'S' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-TableForeignKeyQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], ~cstr.is_not_trusted AS [IsChecked], ~cstr.is_disabled AS [IsEnabled], cstr.is_not_for_replication AS [NotForReplication], ki.name AS [ReferencedKey], rtbl.name AS [ReferencedTable], SCHEMA_NAME(rtbl.schema_id) AS [ReferencedTableSchema], cstr.delete_referential_action AS [DeleteAction], cstr.update_referential_action AS [UpdateAction] FROM sys.tables AS tbl INNER JOIN sys.foreign_keys AS cstr ON cstr.parent_object_id = tbl.object_id LEFT OUTER JOIN sys.indexes AS ki ON ki.index_id = cstr.key_index_id AND ki.object_id = cstr.referenced_object_id INNER JOIN sys.tables rtbl ON rtbl.object_id = cstr.referenced_object_id 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], ~cstr.is_not_trusted AS [IsChecked], ~cstr.is_disabled AS [IsEnabled], cstr.is_not_for_replication AS [NotForReplication], ki.name AS [ReferencedKey], rtbl.name AS [ReferencedTable], SCHEMA_NAME(rtbl.schema_id) AS [ReferencedTableSchema], cstr.delete_referential_action AS [DeleteAction], cstr.update_referential_action AS [UpdateAction], CAST(CASE WHEN filetableobj.object_id IS NULL THEN 0 ELSE 1 END AS BIT) AS [IsFileTableDefined] FROM sys.tables AS tbl INNER JOIN sys.foreign_keys AS cstr ON cstr.parent_object_id = tbl.object_id LEFT OUTER JOIN sys.indexes AS ki ON ki.index_id = cstr.key_index_id AND ki.object_id = cstr.referenced_object_id INNER JOIN sys.tables rtbl ON rtbl.object_id = cstr.referenced_object_id LEFT OUTER JOIN sys.filetable_system_defined_objects AS filetableobj ON filetableobj.object_id = cstr.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], ~cstr.is_not_trusted AS [IsChecked], ~cstr.is_disabled AS [IsEnabled], cstr.is_not_for_replication AS [NotForReplication], ki.name AS [ReferencedKey], rtbl.name AS [ReferencedTable], SCHEMA_NAME(rtbl.schema_id) AS [ReferencedTableSchema], cstr.delete_referential_action AS [DeleteAction], cstr.update_referential_action AS [UpdateAction] FROM sys.tables AS tbl INNER JOIN sys.foreign_keys AS cstr ON cstr.parent_object_id = tbl.object_id LEFT OUTER JOIN sys.indexes AS ki ON ki.index_id = cstr.key_index_id AND ki.object_id = cstr.referenced_object_id INNER JOIN sys.tables rtbl ON rtbl.object_id = cstr.referenced_object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], ~cstr.is_not_trusted AS [IsChecked], ~cstr.is_disabled AS [IsEnabled], cstr.is_not_for_replication AS [NotForReplication], ki.name AS [ReferencedKey], rtbl.name AS [ReferencedTable], SCHEMA_NAME(rtbl.schema_id) AS [ReferencedTableSchema], cstr.delete_referential_action AS [DeleteAction], cstr.update_referential_action AS [UpdateAction] FROM sys.tables AS tbl INNER JOIN sys.foreign_keys AS cstr ON cstr.parent_object_id = tbl.object_id LEFT OUTER JOIN sys.indexes AS ki ON ki.index_id = cstr.key_index_id AND ki.object_id = cstr.referenced_object_id INNER JOIN sys.tables rtbl ON rtbl.object_id = cstr.referenced_object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], ~cstr.is_not_trusted AS [IsChecked], ~cstr.is_disabled AS [IsEnabled], cstr.is_not_for_replication AS [NotForReplication], ki.name AS [ReferencedKey], rtbl.name AS [ReferencedTable], SCHEMA_NAME(rtbl.schema_id) AS [ReferencedTableSchema], cstr.delete_referential_action AS [DeleteAction], cstr.update_referential_action AS [UpdateAction] FROM sys.tables AS tbl INNER JOIN sys.foreign_keys AS cstr ON cstr.parent_object_id = tbl.object_id LEFT OUTER JOIN sys.indexes AS ki ON ki.index_id = cstr.key_index_id AND ki.object_id = cstr.referenced_object_id INNER JOIN sys.tables rtbl ON rtbl.object_id = cstr.referenced_object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(tbl.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(tbl.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT tbl.id AS [TableID], stbl.uid AS [SchemaID], cstr.name AS [Name], cstr.id AS [ID], cstr.crdate AS [CreateDate], CAST(cstr.status & 4 AS BIT) AS [IsSystemNamed], CAST(1 - ISNULL(OBJECTPROPERTY(cstr.id, N'CnstIsNotTrusted'), 0) AS BIT) AS [IsChecked], CAST(1 - ISNULL(OBJECTPROPERTY(cstr.id, N'CnstIsDisabled'), 0) AS BIT) AS [IsEnabled], CAST(ISNULL(OBJECTPROPERTY(cstr.id, N'CnstIsNotRepl'), 0) AS BIT) AS [NotForReplication], ki.name AS [ReferencedKey], rtbl.name AS [ReferencedTable], USER_NAME(rtbl.uid) AS [ReferencedTableSchema], OBJECTPROPERTY(cstr.id, N'CnstIsDeleteCascade') AS [DeleteAction], OBJECTPROPERTY(cstr.id, N'CnstIsUpdateCascade') AS [UpdateAction] FROM dbo.sysobjects AS tbl INNER JOIN sysusers AS stbl ON stbl.uid = tbl.uid INNER JOIN dbo.sysobjects AS cstr ON ( cstr.type = N'F' ) AND ( cstr.parent_obj = tbl.id ) INNER JOIN dbo.sysreferences AS rfr ON rfr.constid = cstr.id LEFT OUTER JOIN dbo.sysindexes AS ki ON ki.indid = rfr.rkeyindid AND ki.id = rfr.rkeyid INNER JOIN dbo.sysobjects AS rtbl ON rfr.rkeyid = rtbl.id WHERE ( tbl.type = 'U' OR tbl.type = 'S' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-TableForeignKeyColumnQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cstr.object_id AS [ForeignKeyID], fk.constraint_column_id AS [ID], cfk.name AS [Name] FROM sys.tables AS tbl INNER JOIN sys.foreign_keys AS cstr ON cstr.parent_object_id = tbl.object_id INNER JOIN sys.foreign_key_columns AS fk ON fk.constraint_object_id = cstr.object_id INNER JOIN sys.columns AS cfk ON fk.parent_column_id = cfk.column_id AND fk.parent_object_id = cfk.object_id 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cstr.object_id AS [ForeignKeyID], fk.constraint_column_id AS [ID], cfk.name AS [Name] FROM sys.tables AS tbl INNER JOIN sys.foreign_keys AS cstr ON cstr.parent_object_id = tbl.object_id INNER JOIN sys.foreign_key_columns AS fk ON fk.constraint_object_id = cstr.object_id INNER JOIN sys.columns AS cfk ON fk.parent_column_id = cfk.column_id AND fk.parent_object_id = cfk.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cstr.object_id AS [ForeignKeyID], fk.constraint_column_id AS [ID], cfk.name AS [Name] FROM sys.tables AS tbl INNER JOIN sys.foreign_keys AS cstr ON cstr.parent_object_id = tbl.object_id INNER JOIN sys.foreign_key_columns AS fk ON fk.constraint_object_id = cstr.object_id INNER JOIN sys.columns AS cfk ON fk.parent_column_id = cfk.column_id AND fk.parent_object_id = cfk.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cstr.object_id AS [ForeignKeyID], fk.constraint_column_id AS [ID], cfk.name AS [Name] FROM sys.tables AS tbl INNER JOIN sys.foreign_keys AS cstr ON cstr.parent_object_id = tbl.object_id INNER JOIN sys.foreign_key_columns AS fk ON fk.constraint_object_id = cstr.object_id INNER JOIN sys.columns AS cfk ON fk.parent_column_id = cfk.column_id AND fk.parent_object_id = cfk.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cstr.object_id AS [ForeignKeyID], fk.constraint_column_id AS [ID], cfk.name AS [Name] FROM sys.tables AS tbl INNER JOIN sys.foreign_keys AS cstr ON cstr.parent_object_id = tbl.object_id INNER JOIN sys.foreign_key_columns AS fk ON fk.constraint_object_id = cstr.object_id INNER JOIN sys.columns AS cfk ON fk.parent_column_id = cfk.column_id AND fk.parent_object_id = cfk.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(tbl.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(tbl.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT tbl.id AS [TableID], stbl.uid AS [SchemaID], fk.constid AS [ForeignKeyID], CAST(fk.keyno AS INT) AS [ID], cfk.name AS [Name] FROM dbo.sysobjects AS tbl INNER JOIN sysusers AS stbl ON stbl.uid = tbl.uid INNER JOIN dbo.sysobjects AS cstr ON ( cstr.type = N'F' ) AND ( cstr.parent_obj = tbl.id ) INNER JOIN dbo.sysforeignkeys AS fk ON fk.constid = cstr.id INNER JOIN dbo.syscolumns AS cfk ON cfk.colid = fk.fkey AND cfk.id = fk.fkeyid WHERE ( tbl.type = 'U' OR tbl.type = 'S' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-TableFullTextIndexQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$null
	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cat.name AS [CatalogName], CAST(fti.is_enabled AS BIT) AS [IsEnabled], OBJECTPROPERTY(fti.object_id, 'TableFullTextPopulateStatus') AS [PopulationStatus], ( CASE change_tracking_state WHEN 'M' THEN 1 WHEN 'A' THEN 2 ELSE 0 END ) AS [ChangeTracking], OBJECTPROPERTY(fti.object_id, 'TableFullTextItemCount') AS [ItemCount], OBJECTPROPERTY(fti.object_id, 'TableFullTextDocsProcessed') AS [DocumentsProcessed], OBJECTPROPERTY(fti.object_id, 'TableFullTextPendingChanges') AS [PendingChanges], OBJECTPROPERTY(fti.object_id, 'TableFullTextFailCount') AS [NumberOfFailures], ( CASE WHEN fti.stoplist_id IS NULL THEN 0 WHEN fti.stoplist_id = 0 THEN 1 ELSE 2 END ) AS [StopListOption], ISNULL(sl.name, N'') AS [StopListName], fg.name AS [FilegroupName], si.name AS [UniqueIndexName], ISNULL(spl.name, N'') AS [SearchPropertyListName] FROM sys.tables AS tbl INNER JOIN sys.fulltext_indexes AS fti ON fti.object_id = tbl.object_id INNER JOIN sys.fulltext_catalogs AS cat ON cat.fulltext_catalog_id = fti.fulltext_catalog_id LEFT OUTER JOIN sys.fulltext_stoplists AS sl ON sl.stoplist_id = fti.stoplist_id INNER JOIN sys.filegroups AS fg ON fg.data_space_id = fti.data_space_id INNER JOIN sys.indexes AS si ON si.index_id = fti.unique_index_id AND si.object_id = fti.object_id LEFT OUTER JOIN sys.registered_search_property_lists AS spl ON spl.property_list_id = fti.property_list_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cat.name AS [CatalogName], CAST(fti.is_enabled AS BIT) AS [IsEnabled], OBJECTPROPERTY(fti.object_id, 'TableFullTextPopulateStatus') AS [PopulationStatus], ( CASE change_tracking_state WHEN 'M' THEN 1 WHEN 'A' THEN 2 ELSE 0 END ) AS [ChangeTracking], OBJECTPROPERTY(fti.object_id, 'TableFullTextItemCount') AS [ItemCount], OBJECTPROPERTY(fti.object_id, 'TableFullTextDocsProcessed') AS [DocumentsProcessed], OBJECTPROPERTY(fti.object_id, 'TableFullTextPendingChanges') AS [PendingChanges], OBJECTPROPERTY(fti.object_id, 'TableFullTextFailCount') AS [NumberOfFailures], ( CASE WHEN fti.stoplist_id IS NULL THEN 0 WHEN fti.stoplist_id = 0 THEN 1 ELSE 2 END ) AS [StopListOption], ISNULL(sl.name, N'') AS [StopListName], fg.name AS [FilegroupName], si.name AS [UniqueIndexName] FROM sys.tables AS tbl INNER JOIN sys.fulltext_indexes AS fti ON fti.object_id = tbl.object_id INNER JOIN sys.fulltext_catalogs AS cat ON cat.fulltext_catalog_id = fti.fulltext_catalog_id LEFT OUTER JOIN sys.fulltext_stoplists AS sl ON sl.stoplist_id = fti.stoplist_id INNER JOIN sys.filegroups AS fg ON fg.data_space_id = fti.data_space_id INNER JOIN sys.indexes AS si ON si.index_id = fti.unique_index_id AND si.object_id = fti.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cat.name AS [CatalogName], CAST(fti.is_enabled AS BIT) AS [IsEnabled], OBJECTPROPERTY(fti.object_id, 'TableFullTextPopulateStatus') AS [PopulationStatus], ( CASE change_tracking_state WHEN 'M' THEN 1 WHEN 'A' THEN 2 ELSE 0 END ) AS [ChangeTracking], OBJECTPROPERTY(fti.object_id, 'TableFullTextItemCount') AS [ItemCount], OBJECTPROPERTY(fti.object_id, 'TableFullTextDocsProcessed') AS [DocumentsProcessed], OBJECTPROPERTY(fti.object_id, 'TableFullTextPendingChanges') AS [PendingChanges], OBJECTPROPERTY(fti.object_id, 'TableFullTextFailCount') AS [NumberOfFailures], si.name AS [UniqueIndexName] FROM sys.tables AS tbl INNER JOIN sys.fulltext_indexes AS fti ON fti.object_id = tbl.object_id INNER JOIN sys.fulltext_catalogs AS cat ON cat.fulltext_catalog_id = fti.fulltext_catalog_id INNER JOIN sys.indexes AS si ON si.index_id = fti.unique_index_id AND si.object_id = fti.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], cat.name AS [CatalogName], CAST(fti.is_enabled AS BIT) AS [IsEnabled], OBJECTPROPERTY(fti.object_id, 'TableFullTextPopulateStatus') AS [PopulationStatus], ( CASE change_tracking_state WHEN 'M' THEN 1 WHEN 'A' THEN 2 ELSE 0 END ) AS [ChangeTracking], OBJECTPROPERTY(fti.object_id, 'TableFullTextItemCount') AS [ItemCount], OBJECTPROPERTY(fti.object_id, 'TableFullTextDocsProcessed') AS [DocumentsProcessed], OBJECTPROPERTY(fti.object_id, 'TableFullTextPendingChanges') AS [PendingChanges], OBJECTPROPERTY(fti.object_id, 'TableFullTextFailCount') AS [NumberOfFailures], si.name AS [UniqueIndexName] FROM sys.tables AS tbl INNER JOIN sys.fulltext_indexes AS fti ON fti.object_id = tbl.object_id INNER JOIN sys.fulltext_catalogs AS cat ON cat.fulltext_catalog_id = fti.fulltext_catalog_id INNER JOIN sys.indexes AS si ON si.index_id = fti.unique_index_id AND si.object_id = fti.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(tbl.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(tbl.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT tbl.id AS [TableID], stbl.uid AS [SchemaID], cat.name AS [CatalogName], CAST(OBJECTPROPERTY(tbl.id, 'TableHasActiveFulltextIndex') AS BIT) AS [IsEnabled], ISNULL(OBJECTPROPERTY(tbl.id, 'TableFullTextPopulateStatus'), 0) AS [PopulationStatus], ISNULL(OBJECTPROPERTY(tbl.id, 'TableFullTextBackgroundUpdateIndexOn'), 0) + ISNULL(OBJECTPROPERTY(tbl.id, 'TableFullTextChangeTrackingOn'), 0) AS [ChangeTracking], si.name AS [UniqueIndexName] FROM dbo.sysobjects AS tbl INNER JOIN sysusers AS stbl ON stbl.uid = tbl.uid INNER JOIN sysfulltextcatalogs AS cat ON ( cat.ftcatid = OBJECTPROPERTY(tbl.id, 'TableFullTextCatalogId') ) AND ( 1 = CAST(OBJECTPROPERTY(tbl.id, 'TableFullTextCatalogId') AS BIT) ) INNER JOIN sysindexes AS si ON si.id = tbl.id AND INDEXPROPERTY(tbl.id, si.name, 'IsFulltextKey') <> 0 WHERE ( tbl.type = 'U' OR tbl.type = 'S' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-TableFullTextIndexColumnQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$null
	} else {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], col.name AS [Name], sl.name AS [Language], ISNULL(col2.name, N'') AS [TypeColumnName], icol.statistical_semantics AS [StatisticalSemantics] FROM sys.tables AS tbl INNER JOIN sys.fulltext_indexes AS fti ON fti.object_id = tbl.object_id INNER JOIN sys.fulltext_index_columns AS icol ON icol.object_id = fti.object_id INNER JOIN sys.columns AS col ON col.object_id = icol.object_id AND col.column_id = icol.column_id INNER JOIN sys.fulltext_languages AS sl ON sl.lcid = icol.language_id LEFT OUTER JOIN sys.columns AS col2 ON col2.column_id = icol.type_column_id AND col2.object_id = icol.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], col.name AS [Name], sl.name AS [Language], ISNULL(col2.name, N'') AS [TypeColumnName] FROM sys.tables AS tbl INNER JOIN sys.fulltext_indexes AS fti ON fti.object_id = tbl.object_id INNER JOIN sys.fulltext_index_columns AS icol ON icol.object_id = fti.object_id INNER JOIN sys.columns AS col ON col.object_id = icol.object_id AND col.column_id = icol.column_id INNER JOIN sys.fulltext_languages AS sl ON sl.lcid = icol.language_id LEFT OUTER JOIN sys.columns AS col2 ON col2.column_id = icol.type_column_id AND col2.object_id = icol.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], col.name AS [Name], sl.name AS [Language], ISNULL(col2.name, N'') AS [TypeColumnName] FROM sys.tables AS tbl INNER JOIN sys.fulltext_indexes AS fti ON fti.object_id = tbl.object_id INNER JOIN sys.fulltext_index_columns AS icol ON icol.object_id = fti.object_id INNER JOIN sys.columns AS col ON col.object_id = icol.object_id AND col.column_id = icol.column_id INNER JOIN sys.fulltext_languages AS sl ON sl.lcid = icol.language_id LEFT OUTER JOIN sys.columns AS col2 ON col2.column_id = icol.type_column_id AND col2.object_id = icol.object_id 		
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], col.name AS [Name], sl.name AS [Language], ISNULL(col2.name, N'') AS [TypeColumnName] FROM sys.tables AS tbl INNER JOIN sys.fulltext_indexes AS fti ON fti.object_id = tbl.object_id INNER JOIN sys.fulltext_index_columns AS icol ON icol.object_id = fti.object_id INNER JOIN sys.columns AS col ON col.object_id = icol.object_id AND col.column_id = icol.column_id INNER JOIN sys.fulltext_languages AS sl ON sl.lcid = icol.language_id LEFT OUTER JOIN sys.columns AS col2 ON col2.column_id = icol.type_column_id AND col2.object_id = icol.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(tbl.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(tbl.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT tbl.id AS [TableID], stbl.uid AS [SchemaID], cols.name AS [Name], ISNULL((SELECT scol2.name FROM sysdepends AS sdep, syscolumns AS scol2 WHERE cols.colid = sdep.number AND cols.id = sdep.id AND sdep.deptype = 1 AND cols.id = scol2.id AND sdep.depnumber = scol2.colid), N'') AS [TypeColumnName], sl.alias AS [Language] FROM dbo.sysobjects AS tbl INNER JOIN sysusers AS stbl ON stbl.uid = tbl.uid INNER JOIN syscolumns cols ON ( COLUMNPROPERTY(cols.id, cols.name, 'IsFulltextIndexed') <> 0 ) AND ( cols.id = tbl.id ) LEFT OUTER JOIN master.dbo.syslanguages AS sl ON sl.lcid = cols.language WHERE ( tbl.type = 'U' OR tbl.type = 'S' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-TableIndexQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], i.name AS [Name], CAST(i.index_id AS INT) AS [ID], CAST(OBJECTPROPERTY(i.object_id, N'IsMSShipped') AS BIT) AS [IsSystemObject], ISNULL(s.no_recompute, 0) AS [NoAutomaticRecomputation], i.fill_factor AS [FillFactor], CAST(CASE i.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], i.is_primary_key + 2 * i.is_unique_constraint AS [IndexKeyType], i.is_unique AS [IsUnique], i.ignore_dup_key AS [IgnoreDuplicateKeys], ~i.allow_row_locks AS [DisallowRowLocks], ~i.allow_page_locks AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], i.is_disabled AS [IsDisabled], CAST(ISNULL(k.is_system_named, 0) AS BIT) AS [IsSystemNamed], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN i.type = 3 THEN 1 ELSE 0 END AS BIT) AS [IsXmlIndex], CAST(ISNULL(spi.spatial_index_type, 0) AS TINYINT) AS [SpatialIndexType], CAST(ISNULL(si.bounding_box_xmin, 0) AS FLOAT(53)) AS [BoundingBoxXMin], CAST(ISNULL(si.bounding_box_ymin, 0) AS FLOAT(53)) AS [BoundingBoxYMin], CAST(ISNULL(si.bounding_box_xmax, 0) AS FLOAT(53)) AS [BoundingBoxXMax], CAST(ISNULL(si.bounding_box_ymax, 0) AS FLOAT(53)) AS [BoundingBoxYMax], CAST(ISNULL(si.level_1_grid, 0) AS SMALLINT) AS [Level1Grid], CAST(ISNULL(si.level_2_grid, 0) AS SMALLINT) AS [Level2Grid], CAST(ISNULL(si.level_3_grid, 0) AS SMALLINT) AS [Level3Grid], CAST(ISNULL(si.level_4_grid, 0) AS SMALLINT) AS [Level4Grid], CAST(ISNULL(si.cells_per_object, 0) AS INT) AS [CellsPerObject], CAST(CASE WHEN i.type = 4 THEN 1 ELSE 0 END AS BIT) AS [IsSpatialIndex], i.has_filter AS [HasFilter], ISNULL(i.filter_definition, N'') AS [FilterDefinition], CAST(CASE i.type WHEN 1 THEN 0 WHEN 4 THEN 4 ELSE 1 END AS TINYINT) AS [IndexType], i.is_hypothetical AS [IsHypothetical] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) LEFT OUTER JOIN sys.stats AS s ON s.stats_id = i.index_id AND s.object_id = i.object_id LEFT OUTER JOIN sys.key_constraints AS k ON k.parent_object_id = i.object_id AND k.unique_index_id = i.index_id LEFT OUTER JOIN sys.spatial_indexes AS spi ON i.object_id = spi.object_id AND i.index_id = spi.index_id LEFT OUTER JOIN sys.spatial_index_tessellations AS si ON i.object_id = si.object_id AND i.index_id = si.index_id 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], i.name AS [Name], CAST(i.index_id AS INT) AS [ID], CAST(OBJECTPROPERTY(i.object_id, N'IsMSShipped') AS BIT) AS [IsSystemObject], ISNULL(s.no_recompute, 0) AS [NoAutomaticRecomputation], i.fill_factor AS [FillFactor], CAST(CASE i.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], i.is_primary_key + 2 * i.is_unique_constraint AS [IndexKeyType], i.is_unique AS [IsUnique], i.ignore_dup_key AS [IgnoreDuplicateKeys], ~i.allow_row_locks AS [DisallowRowLocks], ~i.allow_page_locks AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], i.is_disabled AS [IsDisabled], CAST(ISNULL(k.is_system_named, 0) AS BIT) AS [IsSystemNamed], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN i.type = 3 THEN 1 ELSE 0 END AS BIT) AS [IsXmlIndex], CASE UPPER(ISNULL(xi.secondary_type, '')) WHEN 'P' THEN 1 WHEN 'V' THEN 2 WHEN 'R' THEN 3 ELSE 0 END AS [SecondaryXmlIndexType], ISNULL(xi2.name, N'') AS [ParentXmlIndex], CAST(CASE i.type WHEN 1 THEN 0 WHEN 3 THEN CASE WHEN xi.using_xml_index_id IS NULL THEN 2 ELSE 3 END WHEN 4 THEN 4 WHEN 6 THEN 5 ELSE 1 END AS TINYINT) AS [IndexType], CAST(ISNULL(spi.spatial_index_type, 0) AS TINYINT) AS [SpatialIndexType], CAST(ISNULL(si.bounding_box_xmin, 0) AS FLOAT(53)) AS [BoundingBoxXMin], CAST(ISNULL(si.bounding_box_ymin, 0) AS FLOAT(53)) AS [BoundingBoxYMin], CAST(ISNULL(si.bounding_box_xmax, 0) AS FLOAT(53)) AS [BoundingBoxXMax], CAST(ISNULL(si.bounding_box_ymax, 0) AS FLOAT(53)) AS [BoundingBoxYMax], CAST(ISNULL(si.level_1_grid, 0) AS SMALLINT) AS [Level1Grid], CAST(ISNULL(si.level_2_grid, 0) AS SMALLINT) AS [Level2Grid], CAST(ISNULL(si.level_3_grid, 0) AS SMALLINT) AS [Level3Grid], CAST(ISNULL(si.level_4_grid, 0) AS SMALLINT) AS [Level4Grid], CAST(ISNULL(si.cells_per_object, 0) AS INT) AS [CellsPerObject], CAST(CASE WHEN i.type = 4 THEN 1 ELSE 0 END AS BIT) AS [IsSpatialIndex], i.has_filter AS [HasFilter], ISNULL(i.filter_definition, N'') AS [FilterDefinition], CASE WHEN 'FG' = dsi.type THEN dsi.name ELSE N'' END AS [FileGroup], CASE WHEN 'PS' = dsi.type THEN dsi.name ELSE N'' END AS [PartitionScheme], CAST(CASE WHEN 'PS' = dsi.type THEN 1 ELSE 0 END AS BIT) AS [IsPartitioned], CASE WHEN 'FD' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamFileGroup], CASE WHEN 'PS' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamPartitionScheme], CAST(CASE WHEN filetableobj.object_id IS NULL THEN 0 ELSE 1 END AS BIT) AS [IsFileTableDefined], i.is_hypothetical AS [IsHypothetical] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) LEFT OUTER JOIN sys.stats AS s ON s.stats_id = i.index_id AND s.object_id = i.object_id LEFT OUTER JOIN sys.key_constraints AS k ON k.parent_object_id = i.object_id AND k.unique_index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi ON xi.object_id = i.object_id AND xi.index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi2 ON xi2.object_id = xi.object_id AND xi2.index_id = xi.using_xml_index_id LEFT OUTER JOIN sys.spatial_indexes AS spi ON i.object_id = spi.object_id AND i.index_id = spi.index_id LEFT OUTER JOIN sys.spatial_index_tessellations AS si ON i.object_id = si.object_id AND i.index_id = si.index_id LEFT OUTER JOIN sys.data_spaces AS dsi ON dsi.data_space_id = i.data_space_id LEFT OUTER JOIN sys.tables AS t ON t.object_id = i.object_id LEFT OUTER JOIN sys.data_spaces AS dstbl ON dstbl.data_space_id = t.Filestream_data_space_id AND i.index_id < 2 LEFT OUTER JOIN sys.filetable_system_defined_objects AS filetableobj ON i.object_id = filetableobj.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], i.name AS [Name], CAST(i.index_id AS INT) AS [ID], CAST(OBJECTPROPERTY(i.object_id, N'IsMSShipped') AS BIT) AS [IsSystemObject], ISNULL(s.no_recompute, 0) AS [NoAutomaticRecomputation], i.fill_factor AS [FillFactor], CAST(CASE i.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], i.is_primary_key + 2 * i.is_unique_constraint AS [IndexKeyType], i.is_unique AS [IsUnique], i.ignore_dup_key AS [IgnoreDuplicateKeys], ~i.allow_row_locks AS [DisallowRowLocks], ~i.allow_page_locks AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], i.is_disabled AS [IsDisabled], CAST(ISNULL(k.is_system_named, 0) AS BIT) AS [IsSystemNamed], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN i.type = 3 THEN 1 ELSE 0 END AS BIT) AS [IsXmlIndex], CASE UPPER(ISNULL(xi.secondary_type, '')) WHEN 'P' THEN 1 WHEN 'V' THEN 2 WHEN 'R' THEN 3 ELSE 0 END AS [SecondaryXmlIndexType], ISNULL(xi2.name, N'') AS [ParentXmlIndex], CAST(CASE i.type WHEN 1 THEN 0 WHEN 3 THEN CASE WHEN xi.using_xml_index_id IS NULL THEN 2 ELSE 3 END WHEN 4 THEN 4 WHEN 6 THEN 5 ELSE 1 END AS TINYINT) AS [IndexType], CAST(ISNULL(spi.spatial_index_type, 0) AS TINYINT) AS [SpatialIndexType], CAST(ISNULL(si.bounding_box_xmin, 0) AS FLOAT(53)) AS [BoundingBoxXMin], CAST(ISNULL(si.bounding_box_ymin, 0) AS FLOAT(53)) AS [BoundingBoxYMin], CAST(ISNULL(si.bounding_box_xmax, 0) AS FLOAT(53)) AS [BoundingBoxXMax], CAST(ISNULL(si.bounding_box_ymax, 0) AS FLOAT(53)) AS [BoundingBoxYMax], CAST(ISNULL(si.level_1_grid, 0) AS SMALLINT) AS [Level1Grid], CAST(ISNULL(si.level_2_grid, 0) AS SMALLINT) AS [Level2Grid], CAST(ISNULL(si.level_3_grid, 0) AS SMALLINT) AS [Level3Grid], CAST(ISNULL(si.level_4_grid, 0) AS SMALLINT) AS [Level4Grid], CAST(ISNULL(si.cells_per_object, 0) AS INT) AS [CellsPerObject], CAST(CASE WHEN i.type = 4 THEN 1 ELSE 0 END AS BIT) AS [IsSpatialIndex], i.has_filter AS [HasFilter], ISNULL(i.filter_definition, N'') AS [FilterDefinition], CASE WHEN 'FG' = dsi.type THEN dsi.name ELSE N'' END AS [FileGroup], CASE WHEN 'PS' = dsi.type THEN dsi.name ELSE N'' END AS [PartitionScheme], CAST(CASE WHEN 'PS' = dsi.type THEN 1 ELSE 0 END AS BIT) AS [IsPartitioned], CASE WHEN 'FD' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamFileGroup], CASE WHEN 'PS' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamPartitionScheme], i.is_hypothetical AS [IsHypothetical] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) LEFT OUTER JOIN sys.stats AS s ON s.stats_id = i.index_id AND s.object_id = i.object_id LEFT OUTER JOIN sys.key_constraints AS k ON k.parent_object_id = i.object_id AND k.unique_index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi ON xi.object_id = i.object_id AND xi.index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi2 ON xi2.object_id = xi.object_id AND xi2.index_id = xi.using_xml_index_id LEFT OUTER JOIN sys.spatial_indexes AS spi ON i.object_id = spi.object_id AND i.index_id = spi.index_id LEFT OUTER JOIN sys.spatial_index_tessellations AS si ON i.object_id = si.object_id AND i.index_id = si.index_id LEFT OUTER JOIN sys.data_spaces AS dsi ON dsi.data_space_id = i.data_space_id LEFT OUTER JOIN sys.tables AS t ON t.object_id = i.object_id LEFT OUTER JOIN sys.data_spaces AS dstbl ON dstbl.data_space_id = t.Filestream_data_space_id AND i.index_id < 2 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], i.name AS [Name], CAST(i.index_id AS INT) AS [ID], CAST(OBJECTPROPERTY(i.object_id, N'IsMSShipped') AS BIT) AS [IsSystemObject], ISNULL(s.no_recompute, 0) AS [NoAutomaticRecomputation], i.fill_factor AS [FillFactor], CAST(CASE i.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], i.is_primary_key + 2 * i.is_unique_constraint AS [IndexKeyType], i.is_unique AS [IsUnique], i.ignore_dup_key AS [IgnoreDuplicateKeys], ~i.allow_row_locks AS [DisallowRowLocks], ~i.allow_page_locks AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], i.is_disabled AS [IsDisabled], CAST(ISNULL(k.is_system_named, 0) AS BIT) AS [IsSystemNamed], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN i.type = 3 THEN 1 ELSE 0 END AS BIT) AS [IsXmlIndex], CASE UPPER(ISNULL(xi.secondary_type, '')) WHEN 'P' THEN 1 WHEN 'V' THEN 2 WHEN 'R' THEN 3 ELSE 0 END AS [SecondaryXmlIndexType], ISNULL(xi2.name, N'') AS [ParentXmlIndex], CAST(CASE i.type WHEN 1 THEN 0 WHEN 3 THEN CASE WHEN xi.using_xml_index_id IS NULL THEN 2 ELSE 3 END WHEN 4 THEN 4 WHEN 6 THEN 5 ELSE 1 END AS TINYINT) AS [IndexType], CAST(ISNULL(spi.spatial_index_type, 0) AS TINYINT) AS [SpatialIndexType], CAST(ISNULL(si.bounding_box_xmin, 0) AS FLOAT(53)) AS [BoundingBoxXMin], CAST(ISNULL(si.bounding_box_ymin, 0) AS FLOAT(53)) AS [BoundingBoxYMin], CAST(ISNULL(si.bounding_box_xmax, 0) AS FLOAT(53)) AS [BoundingBoxXMax], CAST(ISNULL(si.bounding_box_ymax, 0) AS FLOAT(53)) AS [BoundingBoxYMax], CAST(ISNULL(si.level_1_grid, 0) AS SMALLINT) AS [Level1Grid], CAST(ISNULL(si.level_2_grid, 0) AS SMALLINT) AS [Level2Grid], CAST(ISNULL(si.level_3_grid, 0) AS SMALLINT) AS [Level3Grid], CAST(ISNULL(si.level_4_grid, 0) AS SMALLINT) AS [Level4Grid], CAST(ISNULL(si.cells_per_object, 0) AS INT) AS [CellsPerObject], CAST(CASE WHEN i.type = 4 THEN 1 ELSE 0 END AS BIT) AS [IsSpatialIndex], i.has_filter AS [HasFilter], ISNULL(i.filter_definition, N'') AS [FilterDefinition], CASE WHEN 'FG' = dsi.type THEN dsi.name ELSE N'' END AS [FileGroup], CASE WHEN 'PS' = dsi.type THEN dsi.name ELSE N'' END AS [PartitionScheme], CAST(CASE WHEN 'PS' = dsi.type THEN 1 ELSE 0 END AS BIT) AS [IsPartitioned], CASE WHEN 'FD' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamFileGroup], CASE WHEN 'PS' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamPartitionScheme], i.is_hypothetical AS [IsHypothetical] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) LEFT OUTER JOIN sys.stats AS s ON s.stats_id = i.index_id AND s.object_id = i.object_id LEFT OUTER JOIN sys.key_constraints AS k ON k.parent_object_id = i.object_id AND k.unique_index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi ON xi.object_id = i.object_id AND xi.index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi2 ON xi2.object_id = xi.object_id AND xi2.index_id = xi.using_xml_index_id LEFT OUTER JOIN sys.spatial_indexes AS spi ON i.object_id = spi.object_id AND i.index_id = spi.index_id LEFT OUTER JOIN sys.spatial_index_tessellations AS si ON i.object_id = si.object_id AND i.index_id = si.index_id LEFT OUTER JOIN sys.data_spaces AS dsi ON dsi.data_space_id = i.data_space_id LEFT OUTER JOIN sys.tables AS t ON t.object_id = i.object_id LEFT OUTER JOIN sys.data_spaces AS dstbl ON dstbl.data_space_id = t.Filestream_data_space_id AND i.index_id < 2 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], i.name AS [Name], CAST(i.index_id AS INT) AS [ID], CAST(OBJECTPROPERTY(i.object_id, N'IsMSShipped') AS BIT) AS [IsSystemObject], ISNULL(s.no_recompute, 0) AS [NoAutomaticRecomputation], i.fill_factor AS [FillFactor], CAST(CASE i.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], i.is_primary_key + 2 * i.is_unique_constraint AS [IndexKeyType], i.is_unique AS [IsUnique], i.ignore_dup_key AS [IgnoreDuplicateKeys], ~i.allow_row_locks AS [DisallowRowLocks], ~i.allow_page_locks AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], i.is_disabled AS [IsDisabled], CAST(ISNULL(k.is_system_named, 0) AS BIT) AS [IsSystemNamed], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN i.type = 3 THEN 1 ELSE 0 END AS BIT) AS [IsXmlIndex], CASE UPPER(ISNULL(xi.secondary_type, '')) WHEN 'P' THEN 1 WHEN 'V' THEN 2 WHEN 'R' THEN 3 ELSE 0 END AS [SecondaryXmlIndexType], ISNULL(xi2.name, N'') AS [ParentXmlIndex], CAST(CASE i.type WHEN 1 THEN 0 WHEN 3 THEN CASE WHEN xi.using_xml_index_id IS NULL THEN 2 ELSE 3 END WHEN 4 THEN 4 WHEN 6 THEN 5 ELSE 1 END AS TINYINT) AS [IndexType], CASE WHEN 'FG' = dsi.type THEN dsi.name ELSE N'' END AS [FileGroup], CASE WHEN 'PS' = dsi.type THEN dsi.name ELSE N'' END AS [PartitionScheme], CAST(CASE WHEN 'PS' = dsi.type THEN 1 ELSE 0 END AS BIT) AS [IsPartitioned], i.is_hypothetical AS [IsHypothetical] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) LEFT OUTER JOIN sys.stats AS s ON s.stats_id = i.index_id AND s.object_id = i.object_id LEFT OUTER JOIN sys.key_constraints AS k ON k.parent_object_id = i.object_id AND k.unique_index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi ON xi.object_id = i.object_id AND xi.index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi2 ON xi2.object_id = xi.object_id AND xi2.index_id = xi.using_xml_index_id LEFT OUTER JOIN sys.data_spaces AS dsi ON dsi.data_space_id = i.data_space_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(tbl.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(tbl.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT tbl.id AS [TableID], stbl.uid AS [SchemaID], i.name AS [Name], CAST(i.indid AS INT) AS [ID], CAST(OBJECTPROPERTY(i.id, N'IsMSShipped') AS BIT) AS [IsSystemObject], CAST(INDEXPROPERTY(i.id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN ( i.status & 0x1000000 ) <> 0 THEN 1 ELSE 0 END AS BIT) AS [NoAutomaticRecomputation], CAST(INDEXPROPERTY(i.id, i.name, N'IndexFillFactor') AS TINYINT) AS [FillFactor], CAST(CASE i.indid WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], CASE WHEN 0 != i.status & 0x800 THEN 1 WHEN 0 != i.status & 0x1000 THEN 2 ELSE 0 END AS [IndexKeyType], CAST(i.status & 2 AS BIT) AS [IsUnique], CAST(CASE WHEN 0 != ( i.status & 0x01 ) THEN 1 ELSE 0 END AS BIT) AS [IgnoreDuplicateKeys], CAST(INDEXPROPERTY(i.id, i.name, N'IsRowLockDisallowed') AS BIT) AS [DisallowRowLocks], CAST(INDEXPROPERTY(i.id, i.name, N'IsPageLockDisallowed') AS BIT) AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], CAST(ISNULL(k.status & 4, 0) AS BIT) AS [IsSystemNamed], CAST(CASE i.indid WHEN 1 THEN 0 ELSE 1 END AS TINYINT) AS [IndexType], fgi.groupname AS [FileGroup], CAST(INDEXPROPERTY(i.id, i.name, N'IsHypothetical') AS BIT) AS [IsHypothetical] FROM dbo.sysobjects AS tbl INNER JOIN sysusers AS stbl ON stbl.uid = tbl.uid INNER JOIN dbo.sysindexes AS i ON ( i.indid > 0 AND i.indid < 255 AND 1 != INDEXPROPERTY(i.id, i.name, N'IsStatistics') ) AND ( i.id = tbl.id ) LEFT OUTER JOIN dbo.sysobjects AS k ON k.parent_obj = i.id AND k.name = i.name AND k.xtype IN ( N'PK', N'UQ' ) LEFT OUTER JOIN dbo.sysfilegroups AS fgi ON fgi.groupid = i.groupid WHERE ( tbl.type = 'U' OR tbl.type = 'S' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-TableIndexColumnQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], clmns.name AS [Name], ( CASE ic.key_ordinal WHEN 0 THEN ic.index_column_id ELSE ic.key_ordinal END ) AS [ID], CAST(COLUMNPROPERTY(ic.object_id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], ic.is_descending_key AS [Descending], ic.is_included_column AS [IsIncluded] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) INNER JOIN sys.index_columns AS ic ON ( ic.column_id > 0 AND ( ic.key_ordinal > 0 OR ic.partition_ordinal = 0 OR ic.is_included_column != 0 ) ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = i.object_id ) INNER JOIN sys.columns AS clmns ON clmns.object_id = ic.object_id AND clmns.column_id = ic.column_id 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], clmns.name AS [Name], ( CASE ic.key_ordinal WHEN 0 THEN ic.index_column_id ELSE ic.key_ordinal END ) AS [ID], CAST(COLUMNPROPERTY(ic.object_id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], ic.is_descending_key AS [Descending], ic.is_included_column AS [IsIncluded] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) INNER JOIN sys.index_columns AS ic ON ( ic.column_id > 0 AND ( ic.key_ordinal > 0 OR ic.partition_ordinal = 0 OR ic.is_included_column != 0 ) ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = i.object_id ) INNER JOIN sys.columns AS clmns ON clmns.object_id = ic.object_id AND clmns.column_id = ic.column_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], clmns.name AS [Name], ( CASE ic.key_ordinal WHEN 0 THEN ic.index_column_id ELSE ic.key_ordinal END ) AS [ID], CAST(COLUMNPROPERTY(ic.object_id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], ic.is_descending_key AS [Descending], ic.is_included_column AS [IsIncluded] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) INNER JOIN sys.index_columns AS ic ON ( ic.column_id > 0 AND ( ic.key_ordinal > 0 OR ic.partition_ordinal = 0 OR ic.is_included_column != 0 ) ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = i.object_id ) INNER JOIN sys.columns AS clmns ON clmns.object_id = ic.object_id AND clmns.column_id = ic.column_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], clmns.name AS [Name], ( CASE ic.key_ordinal WHEN 0 THEN ic.index_column_id ELSE ic.key_ordinal END ) AS [ID], CAST(COLUMNPROPERTY(ic.object_id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], ic.is_descending_key AS [Descending], ic.is_included_column AS [IsIncluded] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) INNER JOIN sys.index_columns AS ic ON ( ic.column_id > 0 AND ( ic.key_ordinal > 0 OR ic.partition_ordinal = 0 OR ic.is_included_column != 0 ) ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = i.object_id ) INNER JOIN sys.columns AS clmns ON clmns.object_id = ic.object_id AND clmns.column_id = ic.column_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], clmns.name AS [Name], ( CASE ic.key_ordinal WHEN 0 THEN ic.index_column_id ELSE ic.key_ordinal END ) AS [ID], CAST(COLUMNPROPERTY(ic.object_id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], ic.is_descending_key AS [Descending], ic.is_included_column AS [IsIncluded] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) INNER JOIN sys.index_columns AS ic ON ( ic.column_id > 0 AND ( ic.key_ordinal > 0 OR ic.partition_ordinal = 0 OR ic.is_included_column != 0 ) ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = i.object_id ) INNER JOIN sys.columns AS clmns ON clmns.object_id = ic.object_id AND clmns.column_id = ic.column_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(tbl.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(tbl.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT tbl.id AS [TableID], stbl.uid AS [SchemaID], CAST(i.indid AS INT) AS [IndexID], clmns.name AS [Name], CAST(ic.keyno AS INT) AS [ID], CAST(COLUMNPROPERTY(ic.id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], CAST(INDEXKEY_PROPERTY(ic.id, ic.indid, ic.keyno, N'IsDescending') AS BIT) AS [Descending] FROM dbo.sysobjects AS tbl INNER JOIN sysusers AS stbl ON stbl.uid = tbl.uid INNER JOIN dbo.sysindexes AS i ON ( i.indid > 0 AND i.indid < 255 AND 1 != INDEXPROPERTY(i.id, i.name, N'IsStatistics') ) AND ( i.id = tbl.id ) INNER JOIN dbo.sysindexkeys AS ic ON CAST(ic.indid AS INT) = CAST(i.indid AS INT) AND ic.id = i.id INNER JOIN dbo.syscolumns AS clmns ON clmns.id = ic.id AND clmns.colid = ic.colid AND clmns.number = 0 WHERE ( tbl.type = 'U' OR tbl.type = 'S' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-TableIndexPartitionSchemeParameterQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$null
	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], CAST(ic.partition_ordinal AS INT) AS [ID], c.name AS [Name] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) INNER JOIN sys.index_columns ic ON ( ic.partition_ordinal > 0 ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = CAST(i.object_id AS INT) ) INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], CAST(ic.partition_ordinal AS INT) AS [ID], c.name AS [Name] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) INNER JOIN sys.index_columns ic ON ( ic.partition_ordinal > 0 ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = CAST(i.object_id AS INT) ) INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], CAST(ic.partition_ordinal AS INT) AS [ID], c.name AS [Name] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) INNER JOIN sys.index_columns ic ON ( ic.partition_ordinal > 0 ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = CAST(i.object_id AS INT) ) INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], CAST(ic.partition_ordinal AS INT) AS [ID], c.name AS [Name] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) INNER JOIN sys.index_columns ic ON ( ic.partition_ordinal > 0 ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = CAST(i.object_id AS INT) ) INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$null
		}
	}
}

function Get-TableIndexPhysicalPartitionQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$null
	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], p.partition_number AS [PartitionNumber], prv.value AS [RightBoundaryValue], fg.name AS [FileGroupName], CAST(pf.boundary_value_on_right AS INT) AS [RangeType], CAST(p.rows AS FLOAT) AS [RowCount], p.data_compression AS [DataCompression] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) LEFT OUTER JOIN sys.all_objects AS allobj ON allobj.name = 'extended_index_' + CAST(i.object_id AS VARCHAR) + '_' + CAST(i.index_id AS VARCHAR) AND allobj.type = 'IT' INNER JOIN sys.partitions AS p ON p.object_id = CAST(( CASE WHEN i.type = 4 THEN allobj.object_id ELSE i.object_id END ) AS INT) AND p.index_id = CAST(( CASE WHEN i.type = 4 THEN 1 ELSE i.index_id END ) AS INT) LEFT OUTER JOIN sys.destination_data_spaces AS dds ON dds.partition_scheme_id = i.data_space_id AND dds.destination_id = p.partition_number LEFT OUTER JOIN sys.partition_schemes AS ps ON ps.data_space_id = i.data_space_id LEFT OUTER JOIN sys.partition_range_values AS prv ON prv.boundary_id = p.partition_number AND prv.function_id = ps.function_id LEFT OUTER JOIN sys.filegroups AS fg ON fg.data_space_id = dds.data_space_id OR fg.data_space_id = i.data_space_id LEFT OUTER JOIN sys.partition_functions AS pf ON pf.function_id = prv.function_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], p.partition_number AS [PartitionNumber], prv.value AS [RightBoundaryValue], fg.name AS [FileGroupName], CAST(pf.boundary_value_on_right AS INT) AS [RangeType], CAST(p.rows AS FLOAT) AS [RowCount], p.data_compression AS [DataCompression] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) LEFT OUTER JOIN sys.all_objects AS allobj ON allobj.name = 'extended_index_' + CAST(i.object_id AS VARCHAR) + '_' + CAST(i.index_id AS VARCHAR) AND allobj.type = 'IT' INNER JOIN sys.partitions AS p ON p.object_id = CAST(( CASE WHEN i.type = 4 THEN allobj.object_id ELSE i.object_id END ) AS INT) AND p.index_id = CAST(( CASE WHEN i.type = 4 THEN 1 ELSE i.index_id END ) AS INT) LEFT OUTER JOIN sys.destination_data_spaces AS dds ON dds.partition_scheme_id = i.data_space_id AND dds.destination_id = p.partition_number LEFT OUTER JOIN sys.partition_schemes AS ps ON ps.data_space_id = i.data_space_id LEFT OUTER JOIN sys.partition_range_values AS prv ON prv.boundary_id = p.partition_number AND prv.function_id = ps.function_id LEFT OUTER JOIN sys.filegroups AS fg ON fg.data_space_id = dds.data_space_id OR fg.data_space_id = i.data_space_id LEFT OUTER JOIN sys.partition_functions AS pf ON pf.function_id = prv.function_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], p.partition_number AS [PartitionNumber], prv.value AS [RightBoundaryValue], fg.name AS [FileGroupName], CAST(pf.boundary_value_on_right AS INT) AS [RangeType], CAST(p.rows AS FLOAT) AS [RowCount], p.data_compression AS [DataCompression] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) LEFT OUTER JOIN sys.all_objects AS allobj ON allobj.name = 'extended_index_' + CAST(i.object_id AS VARCHAR) + '_' + CAST(i.index_id AS VARCHAR) AND allobj.type = 'IT' INNER JOIN sys.partitions AS p ON p.object_id = CAST(( CASE WHEN i.type = 4 THEN allobj.object_id ELSE i.object_id END ) AS INT) AND p.index_id = CAST(( CASE WHEN i.type = 4 THEN 1 ELSE i.index_id END ) AS INT) LEFT OUTER JOIN sys.destination_data_spaces AS dds ON dds.partition_scheme_id = i.data_space_id AND dds.destination_id = p.partition_number LEFT OUTER JOIN sys.partition_schemes AS ps ON ps.data_space_id = i.data_space_id LEFT OUTER JOIN sys.partition_range_values AS prv ON prv.boundary_id = p.partition_number AND prv.function_id = ps.function_id LEFT OUTER JOIN sys.filegroups AS fg ON fg.data_space_id = dds.data_space_id OR fg.data_space_id = i.data_space_id LEFT OUTER JOIN sys.partition_functions AS pf ON pf.function_id = prv.function_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], p.partition_number AS [PartitionNumber], prv.value AS [RightBoundaryValue], fg.name AS [FileGroupName], CAST(pf.boundary_value_on_right AS INT) AS [RangeType], CAST(p.rows AS FLOAT) AS [RowCount] FROM sys.tables AS tbl INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = tbl.object_id ) INNER JOIN sys.partitions AS p ON p.object_id = CAST(i.object_id AS INT) AND p.index_id = CAST(i.index_id AS INT) LEFT OUTER JOIN sys.destination_data_spaces AS dds ON dds.partition_scheme_id = i.data_space_id AND dds.destination_id = p.partition_number LEFT OUTER JOIN sys.partition_schemes AS ps ON ps.data_space_id = i.data_space_id LEFT OUTER JOIN sys.partition_range_values AS prv ON prv.boundary_id = p.partition_number AND prv.function_id = ps.function_id LEFT OUTER JOIN sys.filegroups AS fg ON fg.data_space_id = dds.data_space_id OR fg.data_space_id = i.data_space_id LEFT OUTER JOIN sys.partition_functions AS pf ON pf.function_id = prv.function_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$null
		}
	}
}

function Get-TableStatisticsQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], st.name AS [Name], st.stats_id AS [ID], st.no_recompute AS [NoAutomaticRecomputation], STATS_DATE(st.object_id, st.stats_id) AS [LastUpdated], CAST(1 - INDEXPROPERTY(st.object_id, st.name, N'IsStatistics') AS BIT) AS [IsFromIndexCreation], st.auto_created AS [IsAutoCreated], '' AS [FileGroup], st.has_filter AS [HasFilter], ISNULL(st.filter_definition, N'') AS [FilterDefinition] FROM sys.tables AS tbl INNER JOIN sys.stats st ON st.object_id = tbl.object_id 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], st.name AS [Name], st.stats_id AS [ID], st.no_recompute AS [NoAutomaticRecomputation], STATS_DATE(st.object_id, st.stats_id) AS [LastUpdated], CAST(1 - INDEXPROPERTY(st.object_id, st.name, N'IsStatistics') AS BIT) AS [IsFromIndexCreation], st.auto_created AS [IsAutoCreated], '' AS [FileGroup], st.has_filter AS [HasFilter], ISNULL(st.filter_definition, N'') AS [FilterDefinition], st.is_temporary AS [IsTemporary] FROM sys.tables AS tbl INNER JOIN sys.stats st ON st.object_id = tbl.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], st.name AS [Name], st.stats_id AS [ID], st.no_recompute AS [NoAutomaticRecomputation], STATS_DATE(st.object_id, st.stats_id) AS [LastUpdated], CAST(1 - INDEXPROPERTY(st.object_id, st.name, N'IsStatistics') AS BIT) AS [IsFromIndexCreation], st.auto_created AS [IsAutoCreated], '' AS [FileGroup], st.has_filter AS [HasFilter], ISNULL(st.filter_definition, N'') AS [FilterDefinition] FROM sys.tables AS tbl INNER JOIN sys.stats st ON st.object_id = tbl.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], st.name AS [Name], st.stats_id AS [ID], st.no_recompute AS [NoAutomaticRecomputation], STATS_DATE(st.object_id, st.stats_id) AS [LastUpdated], CAST(1 - INDEXPROPERTY(st.object_id, st.name, N'IsStatistics') AS BIT) AS [IsFromIndexCreation], st.auto_created AS [IsAutoCreated], '' AS [FileGroup], st.has_filter AS [HasFilter], ISNULL(st.filter_definition, N'') AS [FilterDefinition] FROM sys.tables AS tbl INNER JOIN sys.stats st ON st.object_id = tbl.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], st.name AS [Name], st.stats_id AS [ID], st.no_recompute AS [NoAutomaticRecomputation], STATS_DATE(st.object_id, st.stats_id) AS [LastUpdated], CAST(1 - INDEXPROPERTY(st.object_id, st.name, N'IsStatistics') AS BIT) AS [IsFromIndexCreation], st.auto_created AS [IsAutoCreated], '' AS [FileGroup] FROM sys.tables AS tbl INNER JOIN sys.stats st ON st.object_id = tbl.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(tbl.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(tbl.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT tbl.id AS [TableID], stbl.uid AS [SchemaID], st.name AS [Name], CAST(st.indid AS INT) AS [ID], CAST(CASE WHEN ( st.status & 16777216 ) <> 0 THEN 1 ELSE 0 END AS BIT) AS [NoAutomaticRecomputation], STATS_DATE(st.id, st.indid) AS [LastUpdated], CAST(1 - INDEXPROPERTY(st.id, st.name, N'IsStatistics') AS BIT) AS [IsFromIndexCreation], CAST(INDEXPROPERTY(st.id, st.name, N'IsAutoStatistics') AS BIT) AS [IsAutoCreated], '' AS [FileGroup] FROM dbo.sysobjects AS tbl INNER JOIN sysusers AS stbl ON stbl.uid = tbl.uid INNER JOIN dbo.sysindexes st ON ( ( st.indid <> 0 AND st.indid <> 255 ) AND 0 = OBJECTPROPERTY(st.id, N'IsMSShipped') ) AND ( st.id = tbl.id ) WHERE ( tbl.type = 'U' OR tbl.type = 'S' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-TableStatisticsColumnQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], st.stats_id AS [StatisticID], sic.stats_column_id AS [ID], COL_NAME(sic.object_id, sic.column_id) AS [Name] FROM sys.tables AS tbl INNER JOIN sys.stats st ON st.object_id = tbl.object_id INNER JOIN sys.stats_columns sic ON sic.stats_id = st.stats_id AND sic.object_id = st.object_id 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], st.stats_id AS [StatisticID], sic.stats_column_id AS [ID], COL_NAME(sic.object_id, sic.column_id) AS [Name] FROM sys.tables AS tbl INNER JOIN sys.stats st ON st.object_id = tbl.object_id INNER JOIN sys.stats_columns sic ON sic.stats_id = st.stats_id AND sic.object_id = st.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], st.stats_id AS [StatisticID], sic.stats_column_id AS [ID], COL_NAME(sic.object_id, sic.column_id) AS [Name] FROM sys.tables AS tbl INNER JOIN sys.stats st ON st.object_id = tbl.object_id INNER JOIN sys.stats_columns sic ON sic.stats_id = st.stats_id AND sic.object_id = st.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], st.stats_id AS [StatisticID], sic.stats_column_id AS [ID], COL_NAME(sic.object_id, sic.column_id) AS [Name] FROM sys.tables AS tbl INNER JOIN sys.stats st ON st.object_id = tbl.object_id INNER JOIN sys.stats_columns sic ON sic.stats_id = st.stats_id AND sic.object_id = st.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], st.stats_id AS [StatisticID], sic.stats_column_id AS [ID], COL_NAME(sic.object_id, sic.column_id) AS [Name] FROM sys.tables AS tbl INNER JOIN sys.stats st ON st.object_id = tbl.object_id INNER JOIN sys.stats_columns sic ON sic.stats_id = st.stats_id AND sic.object_id = st.object_id 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(tbl.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(tbl.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT tbl.id AS [TableID], stbl.uid AS [SchemaID], CAST(st.indid AS INT) AS [StatisticID], CAST(c.keyno AS INT) AS [ID], clmns.name AS [Name] FROM dbo.sysobjects AS tbl INNER JOIN sysusers AS stbl ON stbl.uid = tbl.uid INNER JOIN dbo.sysindexes st ON ( ( st.indid <> 0 AND st.indid <> 255 ) AND 0 = OBJECTPROPERTY(st.id, N'IsMSShipped') ) AND ( st.id = tbl.id ) INNER JOIN dbo.sysindexkeys c ON c.indid = CAST(st.indid AS INT) AND c.id = st.id INNER JOIN dbo.syscolumns clmns ON clmns.id = c.id AND clmns.colid = c.colid WHERE ( tbl.type = 'U' OR tbl.type = 'S' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-TableTriggerQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], tr.name AS [Name], tr.object_id AS [ID], tr.create_date AS [CreateDate], tr.modify_date AS [DateLastModified], CAST(tr.is_ms_shipped AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(CASE WHEN ISNULL(smtr.definition, ssmtr.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CASE ISNULL(smtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END AS [ExecutionContext], ISNULL(USER_NAME(smtr.execute_as_principal_id), N'') AS [ExecutionContextPrincipal], ~trr.is_disabled AS [IsEnabled], trr.is_instead_of_trigger AS [InsteadOf], CAST(ISNULL(tei.object_id, 0) AS BIT) AS [Insert], CASE WHEN tei.is_first = 1 THEN 0 WHEN tei.is_last = 1 THEN 2 ELSE 1 END AS [InsertOrder], CAST(ISNULL(teu.object_id, 0) AS BIT) AS [Update], CASE WHEN teu.is_first = 1 THEN 0 WHEN teu.is_last = 1 THEN 2 ELSE 1 END AS [UpdateOrder], CAST(ISNULL(ted.object_id, 0) AS BIT) AS [Delete], CASE WHEN ted.is_first = 1 THEN 0 WHEN ted.is_last = 1 THEN 2 ELSE 1 END AS [DeleteOrder], CASE WHEN tr.type = N'TR' THEN 1 WHEN tr.type = N'TA' THEN 2 ELSE 1 END AS [ImplementationType], trr.is_not_for_replication AS [NotForReplication], NULL AS [Text], ISNULL(smtr.definition, ssmtr.definition) AS [Definition] FROM sys.tables AS tbl INNER JOIN sys.objects AS tr ON ( tr.type IN ( 'TR', 'TA' ) ) AND ( tr.parent_object_id = tbl.object_id ) LEFT OUTER JOIN sys.sql_modules AS smtr ON smtr.object_id = tr.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmtr ON ssmtr.object_id = tr.object_id INNER JOIN sys.triggers AS trr ON trr.object_id = tr.object_id LEFT OUTER JOIN sys.trigger_events AS tei ON tei.object_id = tr.object_id AND tei.type = 1 LEFT OUTER JOIN sys.trigger_events AS teu ON teu.object_id = tr.object_id AND teu.type = 2 LEFT OUTER JOIN sys.trigger_events AS ted ON ted.object_id = tr.object_id AND ted.type = 3 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' WHERE CAST(CASE WHEN tbl.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = tbl.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], tr.name AS [Name], tr.object_id AS [ID], tr.create_date AS [CreateDate], tr.modify_date AS [DateLastModified], CAST(tr.is_ms_shipped AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(CASE WHEN ISNULL(smtr.definition, ssmtr.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE asmbltr.name END AS [AssemblyName], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE amtr.assembly_class END AS [ClassName], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE amtr.assembly_method END AS [MethodName], CASE WHEN amtr.object_id IS NULL THEN CASE ISNULL(smtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END ELSE CASE ISNULL(amtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END END AS [ExecutionContext], CASE WHEN amtr.object_id IS NULL THEN ISNULL(USER_NAME(smtr.execute_as_principal_id), N'') ELSE USER_NAME(amtr.execute_as_principal_id) END AS [ExecutionContextPrincipal], ~trr.is_disabled AS [IsEnabled], trr.is_instead_of_trigger AS [InsteadOf], CAST(ISNULL(tei.object_id, 0) AS BIT) AS [Insert], CASE WHEN tei.is_first = 1 THEN 0 WHEN tei.is_last = 1 THEN 2 ELSE 1 END AS [InsertOrder], CAST(ISNULL(teu.object_id, 0) AS BIT) AS [Update], CASE WHEN teu.is_first = 1 THEN 0 WHEN teu.is_last = 1 THEN 2 ELSE 1 END AS [UpdateOrder], CAST(ISNULL(ted.object_id, 0) AS BIT) AS [Delete], CASE WHEN ted.is_first = 1 THEN 0 WHEN ted.is_last = 1 THEN 2 ELSE 1 END AS [DeleteOrder], CASE WHEN tr.type = N'TR' THEN 1 WHEN tr.type = N'TA' THEN 2 ELSE 1 END AS [ImplementationType], trr.is_not_for_replication AS [NotForReplication], NULL AS [Text], ISNULL(smtr.definition, ssmtr.definition) AS [Definition] FROM sys.tables AS tbl INNER JOIN sys.objects AS tr ON ( tr.type IN ( 'TR', 'TA' ) ) AND ( tr.parent_object_id = tbl.object_id ) LEFT OUTER JOIN sys.assembly_modules AS mod ON mod.object_id = tr.object_id LEFT OUTER JOIN sys.sql_modules AS smtr ON smtr.object_id = tr.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmtr ON ssmtr.object_id = tr.object_id LEFT OUTER JOIN sys.assemblies AS asmbl ON asmbl.assembly_id = mod.assembly_id LEFT OUTER JOIN sys.assembly_modules AS amtr ON amtr.object_id = tr.object_id LEFT OUTER JOIN sys.assemblies AS asmbltr ON asmbltr.assembly_id = amtr.assembly_id INNER JOIN sys.triggers AS trr ON trr.object_id = tr.object_id LEFT OUTER JOIN sys.trigger_events AS tei ON tei.object_id = tr.object_id AND tei.type = 1 LEFT OUTER JOIN sys.trigger_events AS teu ON teu.object_id = tr.object_id AND teu.type = 2 LEFT OUTER JOIN sys.trigger_events AS ted ON ted.object_id = tr.object_id AND ted.type = 3 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], tr.name AS [Name], tr.object_id AS [ID], tr.create_date AS [CreateDate], tr.modify_date AS [DateLastModified], CAST(tr.is_ms_shipped AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(CASE WHEN ISNULL(smtr.definition, ssmtr.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE asmbltr.name END AS [AssemblyName], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE amtr.assembly_class END AS [ClassName], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE amtr.assembly_method END AS [MethodName], CASE WHEN amtr.object_id IS NULL THEN CASE ISNULL(smtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END ELSE CASE ISNULL(amtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END END AS [ExecutionContext], CASE WHEN amtr.object_id IS NULL THEN ISNULL(USER_NAME(smtr.execute_as_principal_id), N'') ELSE USER_NAME(amtr.execute_as_principal_id) END AS [ExecutionContextPrincipal], ~trr.is_disabled AS [IsEnabled], trr.is_instead_of_trigger AS [InsteadOf], CAST(ISNULL(tei.object_id, 0) AS BIT) AS [Insert], CASE WHEN tei.is_first = 1 THEN 0 WHEN tei.is_last = 1 THEN 2 ELSE 1 END AS [InsertOrder], CAST(ISNULL(teu.object_id, 0) AS BIT) AS [Update], CASE WHEN teu.is_first = 1 THEN 0 WHEN teu.is_last = 1 THEN 2 ELSE 1 END AS [UpdateOrder], CAST(ISNULL(ted.object_id, 0) AS BIT) AS [Delete], CASE WHEN ted.is_first = 1 THEN 0 WHEN ted.is_last = 1 THEN 2 ELSE 1 END AS [DeleteOrder], CASE WHEN tr.type = N'TR' THEN 1 WHEN tr.type = N'TA' THEN 2 ELSE 1 END AS [ImplementationType], trr.is_not_for_replication AS [NotForReplication], NULL AS [Text], ISNULL(smtr.definition, ssmtr.definition) AS [Definition] FROM sys.tables AS tbl INNER JOIN sys.objects AS tr ON ( tr.type IN ( 'TR', 'TA' ) ) AND ( tr.parent_object_id = tbl.object_id ) LEFT OUTER JOIN sys.assembly_modules AS mod ON mod.object_id = tr.object_id LEFT OUTER JOIN sys.sql_modules AS smtr ON smtr.object_id = tr.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmtr ON ssmtr.object_id = tr.object_id LEFT OUTER JOIN sys.assemblies AS asmbl ON asmbl.assembly_id = mod.assembly_id LEFT OUTER JOIN sys.assembly_modules AS amtr ON amtr.object_id = tr.object_id LEFT OUTER JOIN sys.assemblies AS asmbltr ON asmbltr.assembly_id = amtr.assembly_id INNER JOIN sys.triggers AS trr ON trr.object_id = tr.object_id LEFT OUTER JOIN sys.trigger_events AS tei ON tei.object_id = tr.object_id AND tei.type = 1 LEFT OUTER JOIN sys.trigger_events AS teu ON teu.object_id = tr.object_id AND teu.type = 2 LEFT OUTER JOIN sys.trigger_events AS ted ON ted.object_id = tr.object_id AND ted.type = 3 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], tr.name AS [Name], tr.object_id AS [ID], tr.create_date AS [CreateDate], tr.modify_date AS [DateLastModified], CAST(tr.is_ms_shipped AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(CASE WHEN ISNULL(smtr.definition, ssmtr.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE asmbltr.name END AS [AssemblyName], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE amtr.assembly_class END AS [ClassName], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE amtr.assembly_method END AS [MethodName], CASE WHEN amtr.object_id IS NULL THEN CASE ISNULL(smtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END ELSE CASE ISNULL(amtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END END AS [ExecutionContext], CASE WHEN amtr.object_id IS NULL THEN ISNULL(USER_NAME(smtr.execute_as_principal_id), N'') ELSE USER_NAME(amtr.execute_as_principal_id) END AS [ExecutionContextPrincipal], ~trr.is_disabled AS [IsEnabled], trr.is_instead_of_trigger AS [InsteadOf], CAST(ISNULL(tei.object_id, 0) AS BIT) AS [Insert], CASE WHEN tei.is_first = 1 THEN 0 WHEN tei.is_last = 1 THEN 2 ELSE 1 END AS [InsertOrder], CAST(ISNULL(teu.object_id, 0) AS BIT) AS [Update], CASE WHEN teu.is_first = 1 THEN 0 WHEN teu.is_last = 1 THEN 2 ELSE 1 END AS [UpdateOrder], CAST(ISNULL(ted.object_id, 0) AS BIT) AS [Delete], CASE WHEN ted.is_first = 1 THEN 0 WHEN ted.is_last = 1 THEN 2 ELSE 1 END AS [DeleteOrder], CASE WHEN tr.type = N'TR' THEN 1 WHEN tr.type = N'TA' THEN 2 ELSE 1 END AS [ImplementationType], trr.is_not_for_replication AS [NotForReplication], NULL AS [Text], ISNULL(smtr.definition, ssmtr.definition) AS [Definition] FROM sys.tables AS tbl INNER JOIN sys.objects AS tr ON ( tr.type IN ( 'TR', 'TA' ) ) AND ( tr.parent_object_id = tbl.object_id ) LEFT OUTER JOIN sys.assembly_modules AS mod ON mod.object_id = tr.object_id LEFT OUTER JOIN sys.sql_modules AS smtr ON smtr.object_id = tr.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmtr ON ssmtr.object_id = tr.object_id LEFT OUTER JOIN sys.assemblies AS asmbl ON asmbl.assembly_id = mod.assembly_id LEFT OUTER JOIN sys.assembly_modules AS amtr ON amtr.object_id = tr.object_id LEFT OUTER JOIN sys.assemblies AS asmbltr ON asmbltr.assembly_id = amtr.assembly_id INNER JOIN sys.triggers AS trr ON trr.object_id = tr.object_id LEFT OUTER JOIN sys.trigger_events AS tei ON tei.object_id = tr.object_id AND tei.type = 1 LEFT OUTER JOIN sys.trigger_events AS teu ON teu.object_id = tr.object_id AND teu.type = 2 LEFT OUTER JOIN sys.trigger_events AS ted ON ted.object_id = tr.object_id AND ted.type = 3 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT tbl.object_id AS [TableID], tbl.schema_id AS [SchemaID], tr.name AS [Name], tr.object_id AS [ID], tr.create_date AS [CreateDate], tr.modify_date AS [DateLastModified], CAST(tr.is_ms_shipped AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(CASE WHEN ISNULL(smtr.definition, ssmtr.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE asmbltr.name END AS [AssemblyName], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE amtr.assembly_class END AS [ClassName], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE amtr.assembly_method END AS [MethodName], CASE WHEN amtr.object_id IS NULL THEN CASE ISNULL(smtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END ELSE CASE ISNULL(amtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END END AS [ExecutionContext], CASE WHEN amtr.object_id IS NULL THEN ISNULL(USER_NAME(smtr.execute_as_principal_id), N'') ELSE USER_NAME(amtr.execute_as_principal_id) END AS [ExecutionContextPrincipal], ~trr.is_disabled AS [IsEnabled], trr.is_instead_of_trigger AS [InsteadOf], CAST(ISNULL(tei.object_id, 0) AS BIT) AS [Insert], CASE WHEN tei.is_first = 1 THEN 0 WHEN tei.is_last = 1 THEN 2 ELSE 1 END AS [InsertOrder], CAST(ISNULL(teu.object_id, 0) AS BIT) AS [Update], CASE WHEN teu.is_first = 1 THEN 0 WHEN teu.is_last = 1 THEN 2 ELSE 1 END AS [UpdateOrder], CAST(ISNULL(ted.object_id, 0) AS BIT) AS [Delete], CASE WHEN ted.is_first = 1 THEN 0 WHEN ted.is_last = 1 THEN 2 ELSE 1 END AS [DeleteOrder], CASE WHEN tr.type = N'TR' THEN 1 WHEN tr.type = N'TA' THEN 2 ELSE 1 END AS [ImplementationType], trr.is_not_for_replication AS [NotForReplication], NULL AS [Text], ISNULL(smtr.definition, ssmtr.definition) AS [Definition] FROM sys.tables AS tbl INNER JOIN sys.objects AS tr ON ( tr.type IN ( 'TR', 'TA' ) ) AND ( tr.parent_object_id = tbl.object_id ) LEFT OUTER JOIN sys.assembly_modules AS mod ON mod.object_id = tr.object_id LEFT OUTER JOIN sys.sql_modules AS smtr ON smtr.object_id = tr.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmtr ON ssmtr.object_id = tr.object_id LEFT OUTER JOIN sys.assemblies AS asmbl ON asmbl.assembly_id = mod.assembly_id LEFT OUTER JOIN sys.assembly_modules AS amtr ON amtr.object_id = tr.object_id LEFT OUTER JOIN sys.assemblies AS asmbltr ON asmbltr.assembly_id = amtr.assembly_id INNER JOIN sys.triggers AS trr ON trr.object_id = tr.object_id LEFT OUTER JOIN sys.trigger_events AS tei ON tei.object_id = tr.object_id AND tei.type = 1 LEFT OUTER JOIN sys.trigger_events AS teu ON teu.object_id = tr.object_id AND teu.type = 2 LEFT OUTER JOIN sys.trigger_events AS ted ON ted.object_id = tr.object_id AND ted.type = 3 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(tbl.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(tbl.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT tbl.id AS [TableID], stbl.uid AS [SchemaID], tr.name AS [Name], tr.id AS [ID], tr.crdate AS [CreateDate], CAST(CASE WHEN ( OBJECTPROPERTY(tr.id, N'IsMSShipped') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(tr.id, N'IsSystemTable') THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(OBJECTPROPERTY(tr.id, N'ExecIsAnsiNullsOn') AS BIT) AS [AnsiNullsStatus], CAST(OBJECTPROPERTY(tr.id, N'ExecIsQuotedIdentOn') AS BIT) AS [QuotedIdentifierStatus], CAST((SELECT TOP 1 encrypted FROM dbo.syscomments p WHERE tr.id = p.id AND p.colid = 1 AND p.number < 2) AS BIT) AS [IsEncrypted], CAST(1 - OBJECTPROPERTY(tr.id, N'ExecIsTriggerDisabled') AS BIT) AS [IsEnabled], CAST(OBJECTPROPERTY(tr.id, N'ExecIsInsteadOfTrigger') AS BIT) AS [InsteadOf], CAST(OBJECTPROPERTY(tr.id, N'ExecIsInsertTrigger') AS BIT) AS [Insert], CASE WHEN OBJECTPROPERTY(tr.id, N'ExecIsFirstInsertTrigger') = 1 THEN 0 WHEN OBJECTPROPERTY(tr.id, N'ExecIsLastInsertTrigger') = 1 THEN 2 ELSE 1 END AS [InsertOrder], CAST(OBJECTPROPERTY(tr.id, N'ExecIsUpdateTrigger') AS BIT) AS [Update], CASE WHEN OBJECTPROPERTY(tr.id, N'ExecIsFirstUpdateTrigger') = 1 THEN 0 WHEN OBJECTPROPERTY(tr.id, N'ExecIsLastUpdateTrigger') = 1 THEN 2 ELSE 1 END AS [UpdateOrder], CAST(OBJECTPROPERTY(tr.id, N'ExecIsDeleteTrigger') AS BIT) AS [Delete], CASE WHEN OBJECTPROPERTY(tr.id, N'ExecIsFirstDeleteTrigger') = 1 THEN 0 WHEN OBJECTPROPERTY(tr.id, N'ExecIsLastDeleteTrigger') = 1 THEN 2 ELSE 1 END AS [DeleteOrder], CAST(OBJECTPROPERTY(tr.id, N'ExecIsTriggerNotForRepl') AS BIT) AS [NotForReplication], 1 AS [ImplementationType], c.text AS [Definition] FROM dbo.sysobjects AS tbl INNER JOIN sysusers AS stbl ON stbl.uid = tbl.uid INNER JOIN dbo.sysobjects AS tr ON ( tr.type = 'TR' ) AND ( tr.parent_obj = tbl.id ) LEFT OUTER JOIN dbo.syscomments c ON c.id = tr.id AND CASE WHEN c.number > 1 THEN c.number ELSE 0 END = 0 WHERE ( tbl.type = 'U' OR tbl.type = 'S' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-ViewPropertyQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT v.name AS [Name], v.object_id AS [ID], v.create_date AS [CreateDate], v.modify_date AS [DateLastModified], ISNULL(sv.name, N'') AS [Owner], CAST(CASE WHEN v.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(v.schema_id) AS [Schema], v.schema_id AS [SchemaID], CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(v.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(v.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(ISNULL(OBJECTPROPERTYEX(v.object_id, N'IsSchemaBound'), 0) AS BIT) AS [IsSchemaBound], CAST(CASE WHEN ISNULL(smv.definition, ssmv.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CAST(OBJECTPROPERTY(v.object_id, N'HasAfterTrigger') AS BIT) AS [HasAfterTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasInsertTrigger') AS BIT) AS [HasInsertTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasDeleteTrigger') AS BIT) AS [HasDeleteTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasInsteadOfTrigger') AS BIT) AS [HasInsteadOfTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasUpdateTrigger') AS BIT) AS [HasUpdateTrigger], CAST(OBJECTPROPERTY(v.object_id, N'IsIndexed') AS BIT) AS [HasIndex], CAST(OBJECTPROPERTY(v.object_id, N'IsIndexable') AS BIT) AS [IsIndexable], v.has_opaque_metadata AS [ReturnsViewMetadata], NULL AS [Text], ISNULL(smv.definition, ssmv.definition) AS [Definition] FROM sys.all_views AS v LEFT OUTER JOIN sys.database_principals AS sv ON sv.principal_id = ISNULL(v.principal_id, ( OBJECTPROPERTY(v.object_id, 'OwnerId') )) LEFT OUTER JOIN sys.sql_modules AS smv ON smv.object_id = v.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmv ON ssmv.object_id = v.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = v.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT v.name AS [Name], v.object_id AS [ID], v.create_date AS [CreateDate], v.modify_date AS [DateLastModified], ISNULL(sv.name, N'') AS [Owner], CAST(CASE WHEN v.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(v.schema_id) AS [Schema], v.schema_id AS [SchemaID], CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = v.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(v.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(v.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(ISNULL(OBJECTPROPERTYEX(v.object_id, N'IsSchemaBound'), 0) AS BIT) AS [IsSchemaBound], CAST(CASE WHEN ISNULL(smv.definition, ssmv.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CAST(OBJECTPROPERTY(v.object_id, N'HasAfterTrigger') AS BIT) AS [HasAfterTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasInsertTrigger') AS BIT) AS [HasInsertTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasDeleteTrigger') AS BIT) AS [HasDeleteTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasInsteadOfTrigger') AS BIT) AS [HasInsteadOfTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasUpdateTrigger') AS BIT) AS [HasUpdateTrigger], CAST(OBJECTPROPERTY(v.object_id, N'IsIndexed') AS BIT) AS [HasIndex], CAST(OBJECTPROPERTY(v.object_id, N'IsIndexable') AS BIT) AS [IsIndexable], v.has_opaque_metadata AS [ReturnsViewMetadata], NULL AS [Text], ISNULL(smv.definition, ssmv.definition) AS [Definition] FROM sys.all_views AS v LEFT OUTER JOIN sys.database_principals AS sv ON sv.principal_id = ISNULL(v.principal_id, ( OBJECTPROPERTY(v.object_id, 'OwnerId') )) LEFT OUTER JOIN sys.sql_modules AS smv ON smv.object_id = v.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmv ON ssmv.object_id = v.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT v.name AS [Name], v.object_id AS [ID], v.create_date AS [CreateDate], v.modify_date AS [DateLastModified], ISNULL(sv.name, N'') AS [Owner], CAST(CASE WHEN v.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(v.schema_id) AS [Schema], v.schema_id AS [SchemaID], CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = v.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(v.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(v.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(ISNULL(OBJECTPROPERTYEX(v.object_id, N'IsSchemaBound'), 0) AS BIT) AS [IsSchemaBound], CAST(CASE WHEN ISNULL(smv.definition, ssmv.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CAST(OBJECTPROPERTY(v.object_id, N'HasAfterTrigger') AS BIT) AS [HasAfterTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasInsertTrigger') AS BIT) AS [HasInsertTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasDeleteTrigger') AS BIT) AS [HasDeleteTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasInsteadOfTrigger') AS BIT) AS [HasInsteadOfTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasUpdateTrigger') AS BIT) AS [HasUpdateTrigger], CAST(OBJECTPROPERTY(v.object_id, N'IsIndexed') AS BIT) AS [HasIndex], CAST(OBJECTPROPERTY(v.object_id, N'IsIndexable') AS BIT) AS [IsIndexable], v.has_opaque_metadata AS [ReturnsViewMetadata], NULL AS [Text], ISNULL(smv.definition, ssmv.definition) AS [Definition] FROM sys.all_views AS v LEFT OUTER JOIN sys.database_principals AS sv ON sv.principal_id = ISNULL(v.principal_id, ( OBJECTPROPERTY(v.object_id, 'OwnerId') )) LEFT OUTER JOIN sys.sql_modules AS smv ON smv.object_id = v.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmv ON ssmv.object_id = v.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT v.name AS [Name], v.object_id AS [ID], v.create_date AS [CreateDate], v.modify_date AS [DateLastModified], ISNULL(sv.name, N'') AS [Owner], CAST(CASE WHEN v.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(v.schema_id) AS [Schema], v.schema_id AS [SchemaID], CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = v.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(v.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(v.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(ISNULL(OBJECTPROPERTYEX(v.object_id, N'IsSchemaBound'), 0) AS BIT) AS [IsSchemaBound], CAST(CASE WHEN ISNULL(smv.definition, ssmv.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CAST(OBJECTPROPERTY(v.object_id, N'HasAfterTrigger') AS BIT) AS [HasAfterTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasInsertTrigger') AS BIT) AS [HasInsertTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasDeleteTrigger') AS BIT) AS [HasDeleteTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasInsteadOfTrigger') AS BIT) AS [HasInsteadOfTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasUpdateTrigger') AS BIT) AS [HasUpdateTrigger], CAST(OBJECTPROPERTY(v.object_id, N'IsIndexed') AS BIT) AS [HasIndex], CAST(OBJECTPROPERTY(v.object_id, N'IsIndexable') AS BIT) AS [IsIndexable], v.has_opaque_metadata AS [ReturnsViewMetadata], NULL AS [Text], ISNULL(smv.definition, ssmv.definition) AS [Definition] FROM sys.all_views AS v LEFT OUTER JOIN sys.database_principals AS sv ON sv.principal_id = ISNULL(v.principal_id, ( OBJECTPROPERTY(v.object_id, 'OwnerId') )) LEFT OUTER JOIN sys.sql_modules AS smv ON smv.object_id = v.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmv ON ssmv.object_id = v.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT v.name AS [Name], v.object_id AS [ID], v.create_date AS [CreateDate], v.modify_date AS [DateLastModified], ISNULL(sv.name, N'') AS [Owner], CAST(CASE WHEN v.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(v.schema_id) AS [Schema], v.schema_id AS [SchemaID], CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = v.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(v.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(v.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(ISNULL(OBJECTPROPERTYEX(v.object_id, N'IsSchemaBound'), 0) AS BIT) AS [IsSchemaBound], CAST(CASE WHEN ISNULL(smv.definition, ssmv.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CAST(OBJECTPROPERTY(v.object_id, N'HasAfterTrigger') AS BIT) AS [HasAfterTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasInsertTrigger') AS BIT) AS [HasInsertTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasDeleteTrigger') AS BIT) AS [HasDeleteTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasInsteadOfTrigger') AS BIT) AS [HasInsteadOfTrigger], CAST(OBJECTPROPERTY(v.object_id, N'HasUpdateTrigger') AS BIT) AS [HasUpdateTrigger], CAST(OBJECTPROPERTY(v.object_id, N'IsIndexed') AS BIT) AS [HasIndex], CAST(OBJECTPROPERTY(v.object_id, N'IsIndexable') AS BIT) AS [IsIndexable], v.has_opaque_metadata AS [ReturnsViewMetadata], NULL AS [Text], ISNULL(smv.definition, ssmv.definition) AS [Definition] FROM sys.all_views AS v LEFT OUTER JOIN sys.database_principals AS sv ON sv.principal_id = ISNULL(v.principal_id, ( OBJECTPROPERTY(v.object_id, 'OwnerId') )) LEFT OUTER JOIN sys.sql_modules AS smv ON smv.object_id = v.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmv ON ssmv.object_id = v.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(v.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(v.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT v.name AS [Name], v.id AS [ID], v.crdate AS [CreateDate], sv.name AS [Schema], sv.uid AS [SchemaID], sv.name AS [Owner], CAST(CASE WHEN ( OBJECTPROPERTY(v.id, N'IsMSShipped') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(v.id, N'IsSystemTable') THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(OBJECTPROPERTY(v.id, N'ExecIsAnsiNullsOn') AS BIT) AS [AnsiNullsStatus], CAST(OBJECTPROPERTY(v.id, N'ExecIsQuotedIdentOn') AS BIT) AS [QuotedIdentifierStatus], CAST(OBJECTPROPERTY(v.id, N'IsSchemaBound') AS BIT) AS [IsSchemaBound], CAST((SELECT TOP 1 encrypted FROM dbo.syscomments p WHERE v.id = p.id AND p.colid = 1 AND p.number < 2) AS BIT) AS [IsEncrypted], CAST(OBJECTPROPERTY(v.id, N'HasAfterTrigger') AS BIT) AS [HasAfterTrigger], CAST(OBJECTPROPERTY(v.id, N'HasInsertTrigger') AS BIT) AS [HasInsertTrigger], CAST(OBJECTPROPERTY(v.id, N'HasDeleteTrigger') AS BIT) AS [HasDeleteTrigger], CAST(OBJECTPROPERTY(v.id, N'HasInsteadOfTrigger') AS BIT) AS [HasInsteadOfTrigger], CAST(OBJECTPROPERTY(v.id, N'HasUpdateTrigger') AS BIT) AS [HasUpdateTrigger], CAST(OBJECTPROPERTY(v.id, N'IsIndexed') AS BIT) AS [HasIndex], CAST(OBJECTPROPERTY(v.id, N'IsIndexable') AS BIT) AS [IsIndexable], c.text AS [Definition] FROM dbo.sysobjects AS v INNER JOIN sysusers AS sv ON sv.uid = v.uid LEFT OUTER JOIN dbo.syscomments c ON c.id = v.id AND CASE WHEN c.number > 1 THEN c.number ELSE 0 END = 0 WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-ViewColumnQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], clmns.name AS [Name], clmns.column_id AS [ID], clmns.is_nullable AS [Nullable], clmns.is_computed AS [Computed], CAST(ISNULL(cik.index_column_id, 0) AS BIT) AS [InPrimaryKey], clmns.is_ansi_padded AS [AnsiPaddingStatus], CAST(clmns.is_rowguidcol AS BIT) AS [RowGuidCol], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsDeterministic'), 0) AS BIT) AS [IsDeterministic], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsPrecise'), 0) AS BIT) AS [IsPrecise], CAST(ISNULL(cc.is_persisted, 0) AS BIT) AS [IsPersisted], ISNULL(clmns.collation_name, N'') AS [Collation], CAST(ISNULL((SELECT TOP 1 1 FROM sys.foreign_key_columns AS colfk WHERE colfk.parent_column_id = clmns.column_id AND colfk.parent_object_id = clmns.object_id), 0) AS BIT) AS [IsForeignKey], clmns.is_identity AS [Identity], CAST(ISNULL(ic.seed_value, 0) AS BIGINT) AS [IdentitySeed], CAST(ISNULL(ic.increment_value, 0) AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE SCHEMA_NAME(d.schema_id) END ) AS [DefaultSchema], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE SCHEMA_NAME(r.schema_id) END ) AS [RuleSchema], ISNULL(ic.is_not_for_replication, 0) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], CAST(clmns.is_filestream AS BIT) AS [IsFileStream], CAST(clmns.is_sparse AS BIT) AS [IsSparse], CAST(clmns.is_column_set AS BIT) AS [IsColumnSet], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND clmns.max_length <> -1 THEN clmns.max_length / 2 ELSE clmns.max_length END AS INT) AS [Length], CAST(clmns.precision AS INT) AS [NumericPrecision], CAST(clmns.scale AS INT) AS [NumericScale], ISNULL(xscclmns.name, N'') AS [XmlSchemaNamespace], ISNULL(s2clmns.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE clmns.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], ISNULL(cc.definition, N'') AS [ComputedText] FROM sys.all_views AS v INNER JOIN sys.all_columns AS clmns ON clmns.object_id = v.object_id LEFT OUTER JOIN sys.indexes AS ik ON ik.object_id = clmns.object_id AND 1 = ik.is_primary_key LEFT OUTER JOIN sys.index_columns AS cik ON cik.index_id = ik.index_id AND cik.column_id = clmns.column_id AND cik.object_id = clmns.object_id AND 0 = cik.is_included_column LEFT OUTER JOIN sys.computed_columns AS cc ON cc.object_id = clmns.object_id AND cc.column_id = clmns.column_id LEFT OUTER JOIN sys.identity_columns AS ic ON ic.object_id = clmns.object_id AND ic.column_id = clmns.column_id LEFT OUTER JOIN sys.objects AS d ON d.object_id = clmns.default_object_id LEFT OUTER JOIN sys.objects AS r ON r.object_id = clmns.rule_object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = clmns.user_type_id LEFT OUTER JOIN sys.schemas AS s1clmns ON s1clmns.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = clmns.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = clmns.system_type_id ) AND ( baset.user_type_id = clmns.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscclmns ON xscclmns.xml_collection_id = clmns.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2clmns ON s2clmns.schema_id = xscclmns.schema_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = v.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], clmns.name AS [Name], clmns.column_id AS [ID], clmns.is_nullable AS [Nullable], clmns.is_computed AS [Computed], CAST(ISNULL(cik.index_column_id, 0) AS BIT) AS [InPrimaryKey], clmns.is_ansi_padded AS [AnsiPaddingStatus], CAST(clmns.is_rowguidcol AS BIT) AS [RowGuidCol], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsDeterministic'), 0) AS BIT) AS [IsDeterministic], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsPrecise'), 0) AS BIT) AS [IsPrecise], CAST(ISNULL(cc.is_persisted, 0) AS BIT) AS [IsPersisted], ISNULL(clmns.collation_name, N'') AS [Collation], CAST(ISNULL((SELECT TOP 1 1 FROM sys.foreign_key_columns AS colfk WHERE colfk.parent_column_id = clmns.column_id AND colfk.parent_object_id = clmns.object_id), 0) AS BIT) AS [IsForeignKey], clmns.is_identity AS [Identity], CAST(ISNULL(ic.seed_value, 0) AS BIGINT) AS [IdentitySeed], CAST(ISNULL(ic.increment_value, 0) AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE SCHEMA_NAME(d.schema_id) END ) AS [DefaultSchema], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE SCHEMA_NAME(r.schema_id) END ) AS [RuleSchema], ISNULL(ic.is_not_for_replication, 0) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'StatisticalSemantics') AS INT) AS [StatisticalSemantics], CAST(clmns.is_filestream AS BIT) AS [IsFileStream], CAST(clmns.is_sparse AS BIT) AS [IsSparse], CAST(clmns.is_column_set AS BIT) AS [IsColumnSet], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND clmns.max_length <> -1 THEN clmns.max_length / 2 ELSE clmns.max_length END AS INT) AS [Length], CAST(clmns.precision AS INT) AS [NumericPrecision], CAST(clmns.scale AS INT) AS [NumericScale], ISNULL(xscclmns.name, N'') AS [XmlSchemaNamespace], ISNULL(s2clmns.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE clmns.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], ISNULL(cc.definition, N'') AS [ComputedText] FROM sys.all_views AS v INNER JOIN sys.all_columns AS clmns ON clmns.object_id = v.object_id LEFT OUTER JOIN sys.indexes AS ik ON ik.object_id = clmns.object_id AND 1 = ik.is_primary_key LEFT OUTER JOIN sys.index_columns AS cik ON cik.index_id = ik.index_id AND cik.column_id = clmns.column_id AND cik.object_id = clmns.object_id AND 0 = cik.is_included_column LEFT OUTER JOIN sys.computed_columns AS cc ON cc.object_id = clmns.object_id AND cc.column_id = clmns.column_id LEFT OUTER JOIN sys.identity_columns AS ic ON ic.object_id = clmns.object_id AND ic.column_id = clmns.column_id LEFT OUTER JOIN sys.objects AS d ON d.object_id = clmns.default_object_id LEFT OUTER JOIN sys.objects AS r ON r.object_id = clmns.rule_object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = clmns.user_type_id LEFT OUTER JOIN sys.schemas AS s1clmns ON s1clmns.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = clmns.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = clmns.system_type_id ) AND ( baset.user_type_id = clmns.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscclmns ON xscclmns.xml_collection_id = clmns.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2clmns ON s2clmns.schema_id = xscclmns.schema_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], clmns.name AS [Name], clmns.column_id AS [ID], clmns.is_nullable AS [Nullable], clmns.is_computed AS [Computed], CAST(ISNULL(cik.index_column_id, 0) AS BIT) AS [InPrimaryKey], clmns.is_ansi_padded AS [AnsiPaddingStatus], CAST(clmns.is_rowguidcol AS BIT) AS [RowGuidCol], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsDeterministic'), 0) AS BIT) AS [IsDeterministic], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsPrecise'), 0) AS BIT) AS [IsPrecise], CAST(ISNULL(cc.is_persisted, 0) AS BIT) AS [IsPersisted], ISNULL(clmns.collation_name, N'') AS [Collation], CAST(ISNULL((SELECT TOP 1 1 FROM sys.foreign_key_columns AS colfk WHERE colfk.parent_column_id = clmns.column_id AND colfk.parent_object_id = clmns.object_id), 0) AS BIT) AS [IsForeignKey], clmns.is_identity AS [Identity], CAST(ISNULL(ic.seed_value, 0) AS BIGINT) AS [IdentitySeed], CAST(ISNULL(ic.increment_value, 0) AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE SCHEMA_NAME(d.schema_id) END ) AS [DefaultSchema], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE SCHEMA_NAME(r.schema_id) END ) AS [RuleSchema], ISNULL(ic.is_not_for_replication, 0) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], CAST(clmns.is_filestream AS BIT) AS [IsFileStream], CAST(clmns.is_sparse AS BIT) AS [IsSparse], CAST(clmns.is_column_set AS BIT) AS [IsColumnSet], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND clmns.max_length <> -1 THEN clmns.max_length / 2 ELSE clmns.max_length END AS INT) AS [Length], CAST(clmns.precision AS INT) AS [NumericPrecision], CAST(clmns.scale AS INT) AS [NumericScale], ISNULL(xscclmns.name, N'') AS [XmlSchemaNamespace], ISNULL(s2clmns.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE clmns.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], ISNULL(cc.definition, N'') AS [ComputedText] FROM sys.all_views AS v INNER JOIN sys.all_columns AS clmns ON clmns.object_id = v.object_id LEFT OUTER JOIN sys.indexes AS ik ON ik.object_id = clmns.object_id AND 1 = ik.is_primary_key LEFT OUTER JOIN sys.index_columns AS cik ON cik.index_id = ik.index_id AND cik.column_id = clmns.column_id AND cik.object_id = clmns.object_id AND 0 = cik.is_included_column LEFT OUTER JOIN sys.computed_columns AS cc ON cc.object_id = clmns.object_id AND cc.column_id = clmns.column_id LEFT OUTER JOIN sys.identity_columns AS ic ON ic.object_id = clmns.object_id AND ic.column_id = clmns.column_id LEFT OUTER JOIN sys.objects AS d ON d.object_id = clmns.default_object_id LEFT OUTER JOIN sys.objects AS r ON r.object_id = clmns.rule_object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = clmns.user_type_id LEFT OUTER JOIN sys.schemas AS s1clmns ON s1clmns.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = clmns.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = clmns.system_type_id ) AND ( baset.user_type_id = clmns.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscclmns ON xscclmns.xml_collection_id = clmns.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2clmns ON s2clmns.schema_id = xscclmns.schema_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], clmns.name AS [Name], clmns.column_id AS [ID], clmns.is_nullable AS [Nullable], clmns.is_computed AS [Computed], CAST(ISNULL(cik.index_column_id, 0) AS BIT) AS [InPrimaryKey], clmns.is_ansi_padded AS [AnsiPaddingStatus], CAST(clmns.is_rowguidcol AS BIT) AS [RowGuidCol], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsDeterministic'), 0) AS BIT) AS [IsDeterministic], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsPrecise'), 0) AS BIT) AS [IsPrecise], CAST(ISNULL(cc.is_persisted, 0) AS BIT) AS [IsPersisted], ISNULL(clmns.collation_name, N'') AS [Collation], CAST(ISNULL((SELECT TOP 1 1 FROM sys.foreign_key_columns AS colfk WHERE colfk.parent_column_id = clmns.column_id AND colfk.parent_object_id = clmns.object_id), 0) AS BIT) AS [IsForeignKey], clmns.is_identity AS [Identity], CAST(ISNULL(ic.seed_value, 0) AS BIGINT) AS [IdentitySeed], CAST(ISNULL(ic.increment_value, 0) AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE SCHEMA_NAME(d.schema_id) END ) AS [DefaultSchema], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE SCHEMA_NAME(r.schema_id) END ) AS [RuleSchema], ISNULL(ic.is_not_for_replication, 0) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], CAST(clmns.is_filestream AS BIT) AS [IsFileStream], CAST(clmns.is_sparse AS BIT) AS [IsSparse], CAST(clmns.is_column_set AS BIT) AS [IsColumnSet], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND clmns.max_length <> -1 THEN clmns.max_length / 2 ELSE clmns.max_length END AS INT) AS [Length], CAST(clmns.precision AS INT) AS [NumericPrecision], CAST(clmns.scale AS INT) AS [NumericScale], ISNULL(xscclmns.name, N'') AS [XmlSchemaNamespace], ISNULL(s2clmns.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE clmns.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], ISNULL(cc.definition, N'') AS [ComputedText] FROM sys.all_views AS v INNER JOIN sys.all_columns AS clmns ON clmns.object_id = v.object_id LEFT OUTER JOIN sys.indexes AS ik ON ik.object_id = clmns.object_id AND 1 = ik.is_primary_key LEFT OUTER JOIN sys.index_columns AS cik ON cik.index_id = ik.index_id AND cik.column_id = clmns.column_id AND cik.object_id = clmns.object_id AND 0 = cik.is_included_column LEFT OUTER JOIN sys.computed_columns AS cc ON cc.object_id = clmns.object_id AND cc.column_id = clmns.column_id LEFT OUTER JOIN sys.identity_columns AS ic ON ic.object_id = clmns.object_id AND ic.column_id = clmns.column_id LEFT OUTER JOIN sys.objects AS d ON d.object_id = clmns.default_object_id LEFT OUTER JOIN sys.objects AS r ON r.object_id = clmns.rule_object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = clmns.user_type_id LEFT OUTER JOIN sys.schemas AS s1clmns ON s1clmns.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = clmns.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = clmns.system_type_id ) AND ( baset.user_type_id = clmns.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscclmns ON xscclmns.xml_collection_id = clmns.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2clmns ON s2clmns.schema_id = xscclmns.schema_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], clmns.name AS [Name], clmns.column_id AS [ID], clmns.is_nullable AS [Nullable], clmns.is_computed AS [Computed], CAST(ISNULL(cik.index_column_id, 0) AS BIT) AS [InPrimaryKey], clmns.is_ansi_padded AS [AnsiPaddingStatus], CAST(clmns.is_rowguidcol AS BIT) AS [RowGuidCol], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsDeterministic'), 0) AS BIT) AS [IsDeterministic], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsPrecise'), 0) AS BIT) AS [IsPrecise], CAST(ISNULL(cc.is_persisted, 0) AS BIT) AS [IsPersisted], ISNULL(clmns.collation_name, N'') AS [Collation], CAST(ISNULL((SELECT TOP 1 1 FROM sys.foreign_key_columns AS colfk WHERE colfk.parent_column_id = clmns.column_id AND colfk.parent_object_id = clmns.object_id), 0) AS BIT) AS [IsForeignKey], clmns.is_identity AS [Identity], CAST(ISNULL(ic.seed_value, 0) AS BIGINT) AS [IdentitySeed], CAST(ISNULL(ic.increment_value, 0) AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE SCHEMA_NAME(d.schema_id) END ) AS [DefaultSchema], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE SCHEMA_NAME(r.schema_id) END ) AS [RuleSchema], ISNULL(ic.is_not_for_replication, 0) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND clmns.max_length <> -1 THEN clmns.max_length / 2 ELSE clmns.max_length END AS INT) AS [Length], CAST(clmns.precision AS INT) AS [NumericPrecision], CAST(clmns.scale AS INT) AS [NumericScale], ISNULL(xscclmns.name, N'') AS [XmlSchemaNamespace], ISNULL(s2clmns.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE clmns.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], ISNULL(cc.definition, N'') AS [ComputedText] FROM sys.all_views AS v INNER JOIN sys.all_columns AS clmns ON clmns.object_id = v.object_id LEFT OUTER JOIN sys.indexes AS ik ON ik.object_id = clmns.object_id AND 1 = ik.is_primary_key LEFT OUTER JOIN sys.index_columns AS cik ON cik.index_id = ik.index_id AND cik.column_id = clmns.column_id AND cik.object_id = clmns.object_id AND 0 = cik.is_included_column LEFT OUTER JOIN sys.computed_columns AS cc ON cc.object_id = clmns.object_id AND cc.column_id = clmns.column_id LEFT OUTER JOIN sys.identity_columns AS ic ON ic.object_id = clmns.object_id AND ic.column_id = clmns.column_id LEFT OUTER JOIN sys.objects AS d ON d.object_id = clmns.default_object_id LEFT OUTER JOIN sys.objects AS r ON r.object_id = clmns.rule_object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = clmns.user_type_id LEFT OUTER JOIN sys.schemas AS s1clmns ON s1clmns.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = clmns.system_type_id AND baset.user_type_id = baset.system_type_id ) LEFT OUTER JOIN sys.xml_schema_collections AS xscclmns ON xscclmns.xml_collection_id = clmns.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2clmns ON s2clmns.schema_id = xscclmns.schema_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(v.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(v.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT v.id AS [ViewID], sv.uid AS [SchemaID], clmns.name AS [Name], CAST(clmns.colid AS INT) AS [ID], CAST(clmns.isnullable AS BIT) AS [Nullable], CAST(clmns.iscomputed AS BIT) AS [Computed], CAST(ISNULL(cik.colid, 0) AS BIT) AS [InPrimaryKey], CAST(ISNULL(COLUMNPROPERTY(clmns.id, clmns.name, N'UsesAnsiTrim'), 0) AS BIT) AS [AnsiPaddingStatus], CAST(clmns.colstat & 2 AS BIT) AS [RowGuidCol], CAST(clmns.colstat & 8 AS BIT) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], CAST(COLUMNPROPERTY(clmns.id, clmns.name, N'IsIdentity') AS BIT) AS [Identity], CAST(ISNULL((SELECT TOP 1 1 FROM dbo.sysforeignkeys AS colfk WHERE colfk.fkey = clmns.colid AND colfk.fkeyid = clmns.id), 0) AS BIT) AS [IsForeignKey], ISNULL(clmns.collation, N'') AS [Collation], CAST(CASE COLUMNPROPERTY(clmns.id, clmns.name, N'IsIdentity') WHEN 1 THEN IDENT_SEED(QUOTENAME(sv.name) + '.' + QUOTENAME(v.name)) ELSE 0 END AS BIGINT) AS [IdentitySeed], CAST(CASE COLUMNPROPERTY(clmns.id, clmns.name, N'IsIdentity') WHEN 1 THEN IDENT_INCR(QUOTENAME(sv.name) + '.' + QUOTENAME(v.name)) ELSE 0 END AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.cdefault = 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.cdefault = 0 THEN N'' ELSE USER_NAME(d.uid) END ) AS [DefaultSchema], ( CASE WHEN clmns.domain = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.domain = 0 THEN N'' ELSE USER_NAME(r.uid) END ) AS [RuleSchema], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'char', N'varchar', N'binary', N'varbinary', N'nchar', N'nvarchar' ) THEN clmns.prec ELSE clmns.length END AS INT) AS [Length], CAST(clmns.xprec AS INT) AS [NumericPrecision], CAST(clmns.xscale AS INT) AS [NumericScale], comt.text AS [ComputedText] FROM dbo.sysobjects AS v INNER JOIN sysusers AS sv ON sv.uid = v.uid INNER JOIN dbo.syscolumns AS clmns ON clmns.id = v.id LEFT OUTER JOIN dbo.sysindexes AS ik ON ik.id = clmns.id AND 0 != ik.status & 0x0800 LEFT OUTER JOIN dbo.sysindexkeys AS cik ON cik.indid = ik.indid AND cik.colid = clmns.colid AND cik.id = clmns.id LEFT OUTER JOIN dbo.sysobjects AS d ON d.id = clmns.cdefault AND 0 = d.category & 0x0800 LEFT OUTER JOIN dbo.sysobjects AS r ON r.id = clmns.domain LEFT OUTER JOIN systypes AS usrt ON usrt.xusertype = clmns.xusertype LEFT OUTER JOIN sysusers AS s1clmns ON s1clmns.uid = usrt.uid LEFT OUTER JOIN systypes AS baset ON baset.xusertype = clmns.xtype AND baset.xusertype = baset.xtype LEFT OUTER JOIN dbo.syscomments comt ON comt.number = CAST(clmns.colid AS INT) AND comt.id = clmns.id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-ViewFullTextIndexQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$null
	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = v.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], cat.name AS [CatalogName], CAST(fti.is_enabled AS BIT) AS [IsEnabled], OBJECTPROPERTY(fti.object_id, 'TableFullTextPopulateStatus') AS [PopulationStatus], ( CASE change_tracking_state WHEN 'M' THEN 1 WHEN 'A' THEN 2 ELSE 0 END ) AS [ChangeTracking], OBJECTPROPERTY(fti.object_id, 'TableFullTextItemCount') AS [ItemCount], OBJECTPROPERTY(fti.object_id, 'TableFullTextDocsProcessed') AS [DocumentsProcessed], OBJECTPROPERTY(fti.object_id, 'TableFullTextPendingChanges') AS [PendingChanges], OBJECTPROPERTY(fti.object_id, 'TableFullTextFailCount') AS [NumberOfFailures], ( CASE WHEN fti.stoplist_id IS NULL THEN 0 WHEN fti.stoplist_id = 0 THEN 1 ELSE 2 END ) AS [StopListOption], ISNULL(sl.name, N'') AS [StopListName], fg.name AS [FilegroupName], si.name AS [UniqueIndexName], ISNULL(spl.name, N'') AS [SearchPropertyListName] FROM sys.all_views AS v INNER JOIN sys.fulltext_indexes AS fti ON fti.object_id = v.object_id INNER JOIN sys.fulltext_catalogs AS cat ON cat.fulltext_catalog_id = fti.fulltext_catalog_id LEFT OUTER JOIN sys.fulltext_stoplists AS sl ON sl.stoplist_id = fti.stoplist_id INNER JOIN sys.filegroups AS fg ON fg.data_space_id = fti.data_space_id INNER JOIN sys.indexes AS si ON si.index_id = fti.unique_index_id AND si.object_id = fti.object_id LEFT OUTER JOIN sys.registered_search_property_lists AS spl ON spl.property_list_id = fti.property_list_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], cat.name AS [CatalogName], CAST(fti.is_enabled AS BIT) AS [IsEnabled], OBJECTPROPERTY(fti.object_id, 'TableFullTextPopulateStatus') AS [PopulationStatus], ( CASE change_tracking_state WHEN 'M' THEN 1 WHEN 'A' THEN 2 ELSE 0 END ) AS [ChangeTracking], OBJECTPROPERTY(fti.object_id, 'TableFullTextItemCount') AS [ItemCount], OBJECTPROPERTY(fti.object_id, 'TableFullTextDocsProcessed') AS [DocumentsProcessed], OBJECTPROPERTY(fti.object_id, 'TableFullTextPendingChanges') AS [PendingChanges], OBJECTPROPERTY(fti.object_id, 'TableFullTextFailCount') AS [NumberOfFailures], ( CASE WHEN fti.stoplist_id IS NULL THEN 0 WHEN fti.stoplist_id = 0 THEN 1 ELSE 2 END ) AS [StopListOption], ISNULL(sl.name, N'') AS [StopListName], fg.name AS [FilegroupName], si.name AS [UniqueIndexName] FROM sys.all_views AS v INNER JOIN sys.fulltext_indexes AS fti ON fti.object_id = v.object_id INNER JOIN sys.fulltext_catalogs AS cat ON cat.fulltext_catalog_id = fti.fulltext_catalog_id LEFT OUTER JOIN sys.fulltext_stoplists AS sl ON sl.stoplist_id = fti.stoplist_id INNER JOIN sys.filegroups AS fg ON fg.data_space_id = fti.data_space_id INNER JOIN sys.indexes AS si ON si.index_id = fti.unique_index_id AND si.object_id = fti.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], cat.name AS [CatalogName], CAST(fti.is_enabled AS BIT) AS [IsEnabled], OBJECTPROPERTY(fti.object_id, 'TableFullTextPopulateStatus') AS [PopulationStatus], ( CASE change_tracking_state WHEN 'M' THEN 1 WHEN 'A' THEN 2 ELSE 0 END ) AS [ChangeTracking], OBJECTPROPERTY(fti.object_id, 'TableFullTextItemCount') AS [ItemCount], OBJECTPROPERTY(fti.object_id, 'TableFullTextDocsProcessed') AS [DocumentsProcessed], OBJECTPROPERTY(fti.object_id, 'TableFullTextPendingChanges') AS [PendingChanges], OBJECTPROPERTY(fti.object_id, 'TableFullTextFailCount') AS [NumberOfFailures], ( CASE WHEN fti.stoplist_id IS NULL THEN 0 WHEN fti.stoplist_id = 0 THEN 1 ELSE 2 END ) AS [StopListOption], ISNULL(sl.name, N'') AS [StopListName], fg.name AS [FilegroupName], si.name AS [UniqueIndexName] FROM sys.all_views AS v INNER JOIN sys.fulltext_indexes AS fti ON fti.object_id = v.object_id INNER JOIN sys.fulltext_catalogs AS cat ON cat.fulltext_catalog_id = fti.fulltext_catalog_id LEFT OUTER JOIN sys.fulltext_stoplists AS sl ON sl.stoplist_id = fti.stoplist_id INNER JOIN sys.filegroups AS fg ON fg.data_space_id = fti.data_space_id INNER JOIN sys.indexes AS si ON si.index_id = fti.unique_index_id AND si.object_id = fti.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], cat.name AS [CatalogName], CAST(fti.is_enabled AS BIT) AS [IsEnabled], OBJECTPROPERTY(fti.object_id, 'TableFullTextPopulateStatus') AS [PopulationStatus], ( CASE change_tracking_state WHEN 'M' THEN 1 WHEN 'A' THEN 2 ELSE 0 END ) AS [ChangeTracking], OBJECTPROPERTY(fti.object_id, 'TableFullTextItemCount') AS [ItemCount], OBJECTPROPERTY(fti.object_id, 'TableFullTextDocsProcessed') AS [DocumentsProcessed], OBJECTPROPERTY(fti.object_id, 'TableFullTextPendingChanges') AS [PendingChanges], OBJECTPROPERTY(fti.object_id, 'TableFullTextFailCount') AS [NumberOfFailures], si.name AS [UniqueIndexName] FROM sys.all_views AS v INNER JOIN sys.fulltext_indexes AS fti ON fti.object_id = v.object_id INNER JOIN sys.fulltext_catalogs AS cat ON cat.fulltext_catalog_id = fti.fulltext_catalog_id INNER JOIN sys.indexes AS si ON si.index_id = fti.unique_index_id AND si.object_id = fti.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$null
		}
	}
}

function Get-ViewFullTextIndexColumnQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$null
	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = v.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], col.name AS [Name], sl.name AS [Language], ISNULL(col2.name, N'') AS [TypeColumnName], icol.statistical_semantics AS [StatisticalSemantics] FROM sys.all_views AS v INNER JOIN sys.fulltext_indexes AS fti ON fti.object_id = v.object_id INNER JOIN sys.fulltext_index_columns AS icol ON icol.object_id = fti.object_id INNER JOIN sys.columns AS col ON col.object_id = icol.object_id AND col.column_id = icol.column_id INNER JOIN sys.fulltext_languages AS sl ON sl.lcid = icol.language_id LEFT OUTER JOIN sys.columns AS col2 ON col2.column_id = icol.type_column_id AND col2.object_id = icol.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], col.name AS [Name], sl.name AS [Language], ISNULL(col2.name, N'') AS [TypeColumnName] FROM sys.all_views AS v INNER JOIN sys.fulltext_indexes AS fti ON fti.object_id = v.object_id INNER JOIN sys.fulltext_index_columns AS icol ON icol.object_id = fti.object_id INNER JOIN sys.columns AS col ON col.object_id = icol.object_id AND col.column_id = icol.column_id INNER JOIN sys.fulltext_languages AS sl ON sl.lcid = icol.language_id LEFT OUTER JOIN sys.columns AS col2 ON col2.column_id = icol.type_column_id AND col2.object_id = icol.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], col.name AS [Name], sl.name AS [Language], ISNULL(col2.name, N'') AS [TypeColumnName] FROM sys.all_views AS v INNER JOIN sys.fulltext_indexes AS fti ON fti.object_id = v.object_id INNER JOIN sys.fulltext_index_columns AS icol ON icol.object_id = fti.object_id INNER JOIN sys.columns AS col ON col.object_id = icol.object_id AND col.column_id = icol.column_id INNER JOIN sys.fulltext_languages AS sl ON sl.lcid = icol.language_id LEFT OUTER JOIN sys.columns AS col2 ON col2.column_id = icol.type_column_id AND col2.object_id = icol.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], col.name AS [Name], sl.name AS [Language], ISNULL(col2.name, N'') AS [TypeColumnName] FROM sys.all_views AS v INNER JOIN sys.fulltext_indexes AS fti ON fti.object_id = v.object_id INNER JOIN sys.fulltext_index_columns AS icol ON icol.object_id = fti.object_id INNER JOIN sys.columns AS col ON col.object_id = icol.object_id AND col.column_id = icol.column_id INNER JOIN sys.fulltext_languages AS sl ON sl.lcid = icol.language_id LEFT OUTER JOIN sys.columns AS col2 ON col2.column_id = icol.type_column_id AND col2.object_id = icol.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$null
		}
	}
}

function Get-ViewIndexQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], i.name AS [Name], CAST(i.index_id AS INT) AS [ID], CAST(OBJECTPROPERTY(i.object_id, N'IsMSShipped') AS BIT) AS [IsSystemObject], ISNULL(s.no_recompute, 0) AS [NoAutomaticRecomputation], i.fill_factor AS [FillFactor], CAST(CASE i.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], i.is_primary_key + 2 * i.is_unique_constraint AS [IndexKeyType], i.is_unique AS [IsUnique], i.ignore_dup_key AS [IgnoreDuplicateKeys], ~i.allow_row_locks AS [DisallowRowLocks], ~i.allow_page_locks AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], i.is_disabled AS [IsDisabled], CAST(ISNULL(k.is_system_named, 0) AS BIT) AS [IsSystemNamed], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN i.type = 3 THEN 1 ELSE 0 END AS BIT) AS [IsXmlIndex], CAST(ISNULL(spi.spatial_index_type, 0) AS TINYINT) AS [SpatialIndexType], CAST(ISNULL(si.bounding_box_xmin, 0) AS FLOAT(53)) AS [BoundingBoxXMin], CAST(ISNULL(si.bounding_box_ymin, 0) AS FLOAT(53)) AS [BoundingBoxYMin], CAST(ISNULL(si.bounding_box_xmax, 0) AS FLOAT(53)) AS [BoundingBoxXMax], CAST(ISNULL(si.bounding_box_ymax, 0) AS FLOAT(53)) AS [BoundingBoxYMax], CAST(ISNULL(si.level_1_grid, 0) AS SMALLINT) AS [Level1Grid], CAST(ISNULL(si.level_2_grid, 0) AS SMALLINT) AS [Level2Grid], CAST(ISNULL(si.level_3_grid, 0) AS SMALLINT) AS [Level3Grid], CAST(ISNULL(si.level_4_grid, 0) AS SMALLINT) AS [Level4Grid], CAST(ISNULL(si.cells_per_object, 0) AS INT) AS [CellsPerObject], CAST(CASE WHEN i.type = 4 THEN 1 ELSE 0 END AS BIT) AS [IsSpatialIndex], i.has_filter AS [HasFilter], ISNULL(i.filter_definition, N'') AS [FilterDefinition], CAST(CASE i.type WHEN 1 THEN 0 WHEN 4 THEN 4 ELSE 1 END AS TINYINT) AS [IndexType], i.is_hypothetical AS [IsHypothetical] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) LEFT OUTER JOIN sys.stats AS s ON s.stats_id = i.index_id AND s.object_id = i.object_id LEFT OUTER JOIN sys.key_constraints AS k ON k.parent_object_id = i.object_id AND k.unique_index_id = i.index_id LEFT OUTER JOIN sys.spatial_indexes AS spi ON i.object_id = spi.object_id AND i.index_id = spi.index_id LEFT OUTER JOIN sys.spatial_index_tessellations AS si ON i.object_id = si.object_id AND i.index_id = si.index_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = v.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], i.name AS [Name], CAST(i.index_id AS INT) AS [ID], CAST(OBJECTPROPERTY(i.object_id, N'IsMSShipped') AS BIT) AS [IsSystemObject], ISNULL(s.no_recompute, 0) AS [NoAutomaticRecomputation], i.fill_factor AS [FillFactor], CAST(CASE i.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], i.is_primary_key + 2 * i.is_unique_constraint AS [IndexKeyType], i.is_unique AS [IsUnique], i.ignore_dup_key AS [IgnoreDuplicateKeys], ~i.allow_row_locks AS [DisallowRowLocks], ~i.allow_page_locks AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], i.is_disabled AS [IsDisabled], CAST(ISNULL(k.is_system_named, 0) AS BIT) AS [IsSystemNamed], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN i.type = 3 THEN 1 ELSE 0 END AS BIT) AS [IsXmlIndex], CASE UPPER(ISNULL(xi.secondary_type, '')) WHEN 'P' THEN 1 WHEN 'V' THEN 2 WHEN 'R' THEN 3 ELSE 0 END AS [SecondaryXmlIndexType], ISNULL(xi2.name, N'') AS [ParentXmlIndex], CAST(CASE i.type WHEN 1 THEN 0 WHEN 3 THEN CASE WHEN xi.using_xml_index_id IS NULL THEN 2 ELSE 3 END WHEN 4 THEN 4 WHEN 6 THEN 5 ELSE 1 END AS TINYINT) AS [IndexType], CAST(ISNULL(spi.spatial_index_type, 0) AS TINYINT) AS [SpatialIndexType], CAST(ISNULL(si.bounding_box_xmin, 0) AS FLOAT(53)) AS [BoundingBoxXMin], CAST(ISNULL(si.bounding_box_ymin, 0) AS FLOAT(53)) AS [BoundingBoxYMin], CAST(ISNULL(si.bounding_box_xmax, 0) AS FLOAT(53)) AS [BoundingBoxXMax], CAST(ISNULL(si.bounding_box_ymax, 0) AS FLOAT(53)) AS [BoundingBoxYMax], CAST(ISNULL(si.level_1_grid, 0) AS SMALLINT) AS [Level1Grid], CAST(ISNULL(si.level_2_grid, 0) AS SMALLINT) AS [Level2Grid], CAST(ISNULL(si.level_3_grid, 0) AS SMALLINT) AS [Level3Grid], CAST(ISNULL(si.level_4_grid, 0) AS SMALLINT) AS [Level4Grid], CAST(ISNULL(si.cells_per_object, 0) AS INT) AS [CellsPerObject], CAST(CASE WHEN i.type = 4 THEN 1 ELSE 0 END AS BIT) AS [IsSpatialIndex], i.has_filter AS [HasFilter], ISNULL(i.filter_definition, N'') AS [FilterDefinition], CASE WHEN 'FG' = dsi.type THEN dsi.name ELSE N'' END AS [FileGroup], CASE WHEN 'PS' = dsi.type THEN dsi.name ELSE N'' END AS [PartitionScheme], CAST(CASE WHEN 'PS' = dsi.type THEN 1 ELSE 0 END AS BIT) AS [IsPartitioned], CASE WHEN 'FD' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamFileGroup], CASE WHEN 'PS' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamPartitionScheme], CAST(CASE WHEN filetableobj.object_id IS NULL THEN 0 ELSE 1 END AS BIT) AS [IsFileTableDefined], i.is_hypothetical AS [IsHypothetical] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) LEFT OUTER JOIN sys.stats AS s ON s.stats_id = i.index_id AND s.object_id = i.object_id LEFT OUTER JOIN sys.key_constraints AS k ON k.parent_object_id = i.object_id AND k.unique_index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi ON xi.object_id = i.object_id AND xi.index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi2 ON xi2.object_id = xi.object_id AND xi2.index_id = xi.using_xml_index_id LEFT OUTER JOIN sys.spatial_indexes AS spi ON i.object_id = spi.object_id AND i.index_id = spi.index_id LEFT OUTER JOIN sys.spatial_index_tessellations AS si ON i.object_id = si.object_id AND i.index_id = si.index_id LEFT OUTER JOIN sys.data_spaces AS dsi ON dsi.data_space_id = i.data_space_id LEFT OUTER JOIN sys.tables AS t ON t.object_id = i.object_id LEFT OUTER JOIN sys.data_spaces AS dstbl ON dstbl.data_space_id = t.Filestream_data_space_id AND i.index_id < 2 LEFT OUTER JOIN sys.filetable_system_defined_objects AS filetableobj ON i.object_id = filetableobj.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], i.name AS [Name], CAST(i.index_id AS INT) AS [ID], CAST(OBJECTPROPERTY(i.object_id, N'IsMSShipped') AS BIT) AS [IsSystemObject], ISNULL(s.no_recompute, 0) AS [NoAutomaticRecomputation], i.fill_factor AS [FillFactor], CAST(CASE i.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], i.is_primary_key + 2 * i.is_unique_constraint AS [IndexKeyType], i.is_unique AS [IsUnique], i.ignore_dup_key AS [IgnoreDuplicateKeys], ~i.allow_row_locks AS [DisallowRowLocks], ~i.allow_page_locks AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], i.is_disabled AS [IsDisabled], CAST(ISNULL(k.is_system_named, 0) AS BIT) AS [IsSystemNamed], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN i.type = 3 THEN 1 ELSE 0 END AS BIT) AS [IsXmlIndex], CASE UPPER(ISNULL(xi.secondary_type, '')) WHEN 'P' THEN 1 WHEN 'V' THEN 2 WHEN 'R' THEN 3 ELSE 0 END AS [SecondaryXmlIndexType], ISNULL(xi2.name, N'') AS [ParentXmlIndex], CAST(CASE i.type WHEN 1 THEN 0 WHEN 3 THEN CASE WHEN xi.using_xml_index_id IS NULL THEN 2 ELSE 3 END WHEN 4 THEN 4 WHEN 6 THEN 5 ELSE 1 END AS TINYINT) AS [IndexType], CAST(ISNULL(spi.spatial_index_type, 0) AS TINYINT) AS [SpatialIndexType], CAST(ISNULL(si.bounding_box_xmin, 0) AS FLOAT(53)) AS [BoundingBoxXMin], CAST(ISNULL(si.bounding_box_ymin, 0) AS FLOAT(53)) AS [BoundingBoxYMin], CAST(ISNULL(si.bounding_box_xmax, 0) AS FLOAT(53)) AS [BoundingBoxXMax], CAST(ISNULL(si.bounding_box_ymax, 0) AS FLOAT(53)) AS [BoundingBoxYMax], CAST(ISNULL(si.level_1_grid, 0) AS SMALLINT) AS [Level1Grid], CAST(ISNULL(si.level_2_grid, 0) AS SMALLINT) AS [Level2Grid], CAST(ISNULL(si.level_3_grid, 0) AS SMALLINT) AS [Level3Grid], CAST(ISNULL(si.level_4_grid, 0) AS SMALLINT) AS [Level4Grid], CAST(ISNULL(si.cells_per_object, 0) AS INT) AS [CellsPerObject], CAST(CASE WHEN i.type = 4 THEN 1 ELSE 0 END AS BIT) AS [IsSpatialIndex], i.has_filter AS [HasFilter], ISNULL(i.filter_definition, N'') AS [FilterDefinition], CASE WHEN 'FG' = dsi.type THEN dsi.name ELSE N'' END AS [FileGroup], CASE WHEN 'PS' = dsi.type THEN dsi.name ELSE N'' END AS [PartitionScheme], CAST(CASE WHEN 'PS' = dsi.type THEN 1 ELSE 0 END AS BIT) AS [IsPartitioned], CASE WHEN 'FD' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamFileGroup], CASE WHEN 'PS' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamPartitionScheme], i.is_hypothetical AS [IsHypothetical] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) LEFT OUTER JOIN sys.stats AS s ON s.stats_id = i.index_id AND s.object_id = i.object_id LEFT OUTER JOIN sys.key_constraints AS k ON k.parent_object_id = i.object_id AND k.unique_index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi ON xi.object_id = i.object_id AND xi.index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi2 ON xi2.object_id = xi.object_id AND xi2.index_id = xi.using_xml_index_id LEFT OUTER JOIN sys.spatial_indexes AS spi ON i.object_id = spi.object_id AND i.index_id = spi.index_id LEFT OUTER JOIN sys.spatial_index_tessellations AS si ON i.object_id = si.object_id AND i.index_id = si.index_id LEFT OUTER JOIN sys.data_spaces AS dsi ON dsi.data_space_id = i.data_space_id LEFT OUTER JOIN sys.tables AS t ON t.object_id = i.object_id LEFT OUTER JOIN sys.data_spaces AS dstbl ON dstbl.data_space_id = t.Filestream_data_space_id AND i.index_id < 2 WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], i.name AS [Name], CAST(i.index_id AS INT) AS [ID], CAST(OBJECTPROPERTY(i.object_id, N'IsMSShipped') AS BIT) AS [IsSystemObject], ISNULL(s.no_recompute, 0) AS [NoAutomaticRecomputation], i.fill_factor AS [FillFactor], CAST(CASE i.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], i.is_primary_key + 2 * i.is_unique_constraint AS [IndexKeyType], i.is_unique AS [IsUnique], i.ignore_dup_key AS [IgnoreDuplicateKeys], ~i.allow_row_locks AS [DisallowRowLocks], ~i.allow_page_locks AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], i.is_disabled AS [IsDisabled], CAST(ISNULL(k.is_system_named, 0) AS BIT) AS [IsSystemNamed], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN i.type = 3 THEN 1 ELSE 0 END AS BIT) AS [IsXmlIndex], CASE UPPER(ISNULL(xi.secondary_type, '')) WHEN 'P' THEN 1 WHEN 'V' THEN 2 WHEN 'R' THEN 3 ELSE 0 END AS [SecondaryXmlIndexType], ISNULL(xi2.name, N'') AS [ParentXmlIndex], CAST(CASE i.type WHEN 1 THEN 0 WHEN 3 THEN CASE WHEN xi.using_xml_index_id IS NULL THEN 2 ELSE 3 END WHEN 4 THEN 4 WHEN 6 THEN 5 ELSE 1 END AS TINYINT) AS [IndexType], CAST(ISNULL(spi.spatial_index_type, 0) AS TINYINT) AS [SpatialIndexType], CAST(ISNULL(si.bounding_box_xmin, 0) AS FLOAT(53)) AS [BoundingBoxXMin], CAST(ISNULL(si.bounding_box_ymin, 0) AS FLOAT(53)) AS [BoundingBoxYMin], CAST(ISNULL(si.bounding_box_xmax, 0) AS FLOAT(53)) AS [BoundingBoxXMax], CAST(ISNULL(si.bounding_box_ymax, 0) AS FLOAT(53)) AS [BoundingBoxYMax], CAST(ISNULL(si.level_1_grid, 0) AS SMALLINT) AS [Level1Grid], CAST(ISNULL(si.level_2_grid, 0) AS SMALLINT) AS [Level2Grid], CAST(ISNULL(si.level_3_grid, 0) AS SMALLINT) AS [Level3Grid], CAST(ISNULL(si.level_4_grid, 0) AS SMALLINT) AS [Level4Grid], CAST(ISNULL(si.cells_per_object, 0) AS INT) AS [CellsPerObject], CAST(CASE WHEN i.type = 4 THEN 1 ELSE 0 END AS BIT) AS [IsSpatialIndex], i.has_filter AS [HasFilter], ISNULL(i.filter_definition, N'') AS [FilterDefinition], CASE WHEN 'FG' = dsi.type THEN dsi.name ELSE N'' END AS [FileGroup], CASE WHEN 'PS' = dsi.type THEN dsi.name ELSE N'' END AS [PartitionScheme], CAST(CASE WHEN 'PS' = dsi.type THEN 1 ELSE 0 END AS BIT) AS [IsPartitioned], CASE WHEN 'FD' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamFileGroup], CASE WHEN 'PS' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamPartitionScheme], i.is_hypothetical AS [IsHypothetical] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) LEFT OUTER JOIN sys.stats AS s ON s.stats_id = i.index_id AND s.object_id = i.object_id LEFT OUTER JOIN sys.key_constraints AS k ON k.parent_object_id = i.object_id AND k.unique_index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi ON xi.object_id = i.object_id AND xi.index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi2 ON xi2.object_id = xi.object_id AND xi2.index_id = xi.using_xml_index_id LEFT OUTER JOIN sys.spatial_indexes AS spi ON i.object_id = spi.object_id AND i.index_id = spi.index_id LEFT OUTER JOIN sys.spatial_index_tessellations AS si ON i.object_id = si.object_id AND i.index_id = si.index_id LEFT OUTER JOIN sys.data_spaces AS dsi ON dsi.data_space_id = i.data_space_id LEFT OUTER JOIN sys.tables AS t ON t.object_id = i.object_id LEFT OUTER JOIN sys.data_spaces AS dstbl ON dstbl.data_space_id = t.Filestream_data_space_id AND i.index_id < 2 WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], i.name AS [Name], CAST(i.index_id AS INT) AS [ID], CAST(OBJECTPROPERTY(i.object_id, N'IsMSShipped') AS BIT) AS [IsSystemObject], ISNULL(s.no_recompute, 0) AS [NoAutomaticRecomputation], i.fill_factor AS [FillFactor], CAST(CASE i.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], i.is_primary_key + 2 * i.is_unique_constraint AS [IndexKeyType], i.is_unique AS [IsUnique], i.ignore_dup_key AS [IgnoreDuplicateKeys], ~i.allow_row_locks AS [DisallowRowLocks], ~i.allow_page_locks AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], i.is_disabled AS [IsDisabled], CAST(ISNULL(k.is_system_named, 0) AS BIT) AS [IsSystemNamed], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN i.type = 3 THEN 1 ELSE 0 END AS BIT) AS [IsXmlIndex], CASE UPPER(ISNULL(xi.secondary_type, '')) WHEN 'P' THEN 1 WHEN 'V' THEN 2 WHEN 'R' THEN 3 ELSE 0 END AS [SecondaryXmlIndexType], ISNULL(xi2.name, N'') AS [ParentXmlIndex], CAST(CASE i.type WHEN 1 THEN 0 WHEN 3 THEN CASE WHEN xi.using_xml_index_id IS NULL THEN 2 ELSE 3 END WHEN 4 THEN 4 WHEN 6 THEN 5 ELSE 1 END AS TINYINT) AS [IndexType], CASE WHEN 'FG' = dsi.type THEN dsi.name ELSE N'' END AS [FileGroup], CASE WHEN 'PS' = dsi.type THEN dsi.name ELSE N'' END AS [PartitionScheme], CAST(CASE WHEN 'PS' = dsi.type THEN 1 ELSE 0 END AS BIT) AS [IsPartitioned], i.is_hypothetical AS [IsHypothetical] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) LEFT OUTER JOIN sys.stats AS s ON s.stats_id = i.index_id AND s.object_id = i.object_id LEFT OUTER JOIN sys.key_constraints AS k ON k.parent_object_id = i.object_id AND k.unique_index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi ON xi.object_id = i.object_id AND xi.index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi2 ON xi2.object_id = xi.object_id AND xi2.index_id = xi.using_xml_index_id LEFT OUTER JOIN sys.data_spaces AS dsi ON dsi.data_space_id = i.data_space_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(v.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(v.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT v.id AS [ViewID], sv.uid AS [SchemaID], i.name AS [Name], CAST(i.indid AS INT) AS [ID], CAST(OBJECTPROPERTY(i.id, N'IsMSShipped') AS BIT) AS [IsSystemObject], CAST(INDEXPROPERTY(i.id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN ( i.status & 0x1000000 ) <> 0 THEN 1 ELSE 0 END AS BIT) AS [NoAutomaticRecomputation], CAST(INDEXPROPERTY(i.id, i.name, N'IndexFillFactor') AS TINYINT) AS [FillFactor], CAST(CASE i.indid WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], CASE WHEN 0 != i.status & 0x800 THEN 1 WHEN 0 != i.status & 0x1000 THEN 2 ELSE 0 END AS [IndexKeyType], CAST(i.status & 2 AS BIT) AS [IsUnique], CAST(CASE WHEN 0 != ( i.status & 0x01 ) THEN 1 ELSE 0 END AS BIT) AS [IgnoreDuplicateKeys], CAST(INDEXPROPERTY(i.id, i.name, N'IsRowLockDisallowed') AS BIT) AS [DisallowRowLocks], CAST(INDEXPROPERTY(i.id, i.name, N'IsPageLockDisallowed') AS BIT) AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], CAST(ISNULL(k.status & 4, 0) AS BIT) AS [IsSystemNamed], CAST(CASE i.indid WHEN 1 THEN 0 ELSE 1 END AS TINYINT) AS [IndexType], fgi.groupname AS [FileGroup], CAST(INDEXPROPERTY(i.id, i.name, N'IsHypothetical') AS BIT) AS [IsHypothetical] FROM dbo.sysobjects AS v INNER JOIN sysusers AS sv ON sv.uid = v.uid INNER JOIN dbo.sysindexes AS i ON ( i.indid > 0 AND i.indid < 255 AND 1 != INDEXPROPERTY(i.id, i.name, N'IsStatistics') ) AND ( i.id = v.id ) LEFT OUTER JOIN dbo.sysobjects AS k ON k.parent_obj = i.id AND k.name = i.name AND k.xtype IN ( N'PK', N'UQ' ) LEFT OUTER JOIN dbo.sysfilegroups AS fgi ON fgi.groupid = i.groupid WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-ViewIndexColumnQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], clmns.name AS [Name], ( CASE ic.key_ordinal WHEN 0 THEN ic.index_column_id ELSE ic.key_ordinal END ) AS [ID], CAST(COLUMNPROPERTY(ic.object_id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], ic.is_descending_key AS [Descending], ic.is_included_column AS [IsIncluded] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) INNER JOIN sys.index_columns AS ic ON ( ic.column_id > 0 AND ( ic.key_ordinal > 0 OR ic.partition_ordinal = 0 OR ic.is_included_column != 0 ) ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = i.object_id ) INNER JOIN sys.columns AS clmns ON clmns.object_id = ic.object_id AND clmns.column_id = ic.column_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = v.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], clmns.name AS [Name], ( CASE ic.key_ordinal WHEN 0 THEN ic.index_column_id ELSE ic.key_ordinal END ) AS [ID], CAST(COLUMNPROPERTY(ic.object_id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], ic.is_descending_key AS [Descending], ic.is_included_column AS [IsIncluded] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) INNER JOIN sys.index_columns AS ic ON ( ic.column_id > 0 AND ( ic.key_ordinal > 0 OR ic.partition_ordinal = 0 OR ic.is_included_column != 0 ) ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = i.object_id ) INNER JOIN sys.columns AS clmns ON clmns.object_id = ic.object_id AND clmns.column_id = ic.column_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], clmns.name AS [Name], ( CASE ic.key_ordinal WHEN 0 THEN ic.index_column_id ELSE ic.key_ordinal END ) AS [ID], CAST(COLUMNPROPERTY(ic.object_id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], ic.is_descending_key AS [Descending], ic.is_included_column AS [IsIncluded] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) INNER JOIN sys.index_columns AS ic ON ( ic.column_id > 0 AND ( ic.key_ordinal > 0 OR ic.partition_ordinal = 0 OR ic.is_included_column != 0 ) ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = i.object_id ) INNER JOIN sys.columns AS clmns ON clmns.object_id = ic.object_id AND clmns.column_id = ic.column_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], clmns.name AS [Name], ( CASE ic.key_ordinal WHEN 0 THEN ic.index_column_id ELSE ic.key_ordinal END ) AS [ID], CAST(COLUMNPROPERTY(ic.object_id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], ic.is_descending_key AS [Descending], ic.is_included_column AS [IsIncluded] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) INNER JOIN sys.index_columns AS ic ON ( ic.column_id > 0 AND ( ic.key_ordinal > 0 OR ic.partition_ordinal = 0 OR ic.is_included_column != 0 ) ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = i.object_id ) INNER JOIN sys.columns AS clmns ON clmns.object_id = ic.object_id AND clmns.column_id = ic.column_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], clmns.name AS [Name], ( CASE ic.key_ordinal WHEN 0 THEN ic.index_column_id ELSE ic.key_ordinal END ) AS [ID], CAST(COLUMNPROPERTY(ic.object_id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], ic.is_descending_key AS [Descending], ic.is_included_column AS [IsIncluded] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) INNER JOIN sys.index_columns AS ic ON ( ic.column_id > 0 AND ( ic.key_ordinal > 0 OR ic.partition_ordinal = 0 OR ic.is_included_column != 0 ) ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = i.object_id ) INNER JOIN sys.columns AS clmns ON clmns.object_id = ic.object_id AND clmns.column_id = ic.column_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(v.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(v.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT v.id AS [ViewID], sv.uid AS [SchemaID], CAST(i.indid AS INT) AS [IndexID], clmns.name AS [Name], CAST(ic.keyno AS INT) AS [ID], CAST(COLUMNPROPERTY(ic.id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], CAST(INDEXKEY_PROPERTY(ic.id, ic.indid, ic.keyno, N'IsDescending') AS BIT) AS [Descending] FROM dbo.sysobjects AS v INNER JOIN sysusers AS sv ON sv.uid = v.uid INNER JOIN dbo.sysindexes AS i ON ( i.indid > 0 AND i.indid < 255 AND 1 != INDEXPROPERTY(i.id, i.name, N'IsStatistics') ) AND ( i.id = v.id ) INNER JOIN dbo.sysindexkeys AS ic ON CAST(ic.indid AS INT) = CAST(i.indid AS INT) AND ic.id = i.id INNER JOIN dbo.syscolumns AS clmns ON clmns.id = ic.id AND clmns.colid = ic.colid AND clmns.number = 0 WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-ViewIndexPartitionSchemeParameterQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$null
	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = v.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], CAST(ic.partition_ordinal AS INT) AS [ID], c.name AS [Name] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) INNER JOIN sys.index_columns ic ON ( ic.partition_ordinal > 0 ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = CAST(i.object_id AS INT) ) INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], CAST(ic.partition_ordinal AS INT) AS [ID], c.name AS [Name] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) INNER JOIN sys.index_columns ic ON ( ic.partition_ordinal > 0 ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = CAST(i.object_id AS INT) ) INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], CAST(ic.partition_ordinal AS INT) AS [ID], c.name AS [Name] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) INNER JOIN sys.index_columns ic ON ( ic.partition_ordinal > 0 ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = CAST(i.object_id AS INT) ) INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], CAST(ic.partition_ordinal AS INT) AS [ID], c.name AS [Name] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) INNER JOIN sys.index_columns ic ON ( ic.partition_ordinal > 0 ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = CAST(i.object_id AS INT) ) INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$null
		}
	}
}

function Get-ViewIndexPhysicalPartitionQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$null
	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = v.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], p.partition_number AS [PartitionNumber], prv.value AS [RightBoundaryValue], fg.name AS [FileGroupName], CAST(pf.boundary_value_on_right AS INT) AS [RangeType], CAST(p.rows AS FLOAT) AS [RowCount], p.data_compression AS [DataCompression] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) LEFT OUTER JOIN sys.all_objects AS allobj ON allobj.name = 'extended_index_' + CAST(i.object_id AS VARCHAR) + '_' + CAST(i.index_id AS VARCHAR) AND allobj.type = 'IT' INNER JOIN sys.partitions AS p ON p.object_id = CAST(( CASE WHEN i.type = 4 THEN allobj.object_id ELSE i.object_id END ) AS INT) AND p.index_id = CAST(( CASE WHEN i.type = 4 THEN 1 ELSE i.index_id END ) AS INT) LEFT OUTER JOIN sys.destination_data_spaces AS dds ON dds.partition_scheme_id = i.data_space_id AND dds.destination_id = p.partition_number LEFT OUTER JOIN sys.partition_schemes AS ps ON ps.data_space_id = i.data_space_id LEFT OUTER JOIN sys.partition_range_values AS prv ON prv.boundary_id = p.partition_number AND prv.function_id = ps.function_id LEFT OUTER JOIN sys.filegroups AS fg ON fg.data_space_id = dds.data_space_id OR fg.data_space_id = i.data_space_id LEFT OUTER JOIN sys.partition_functions AS pf ON pf.function_id = prv.function_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], p.partition_number AS [PartitionNumber], prv.value AS [RightBoundaryValue], fg.name AS [FileGroupName], CAST(pf.boundary_value_on_right AS INT) AS [RangeType], CAST(p.rows AS FLOAT) AS [RowCount], p.data_compression AS [DataCompression] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) LEFT OUTER JOIN sys.all_objects AS allobj ON allobj.name = 'extended_index_' + CAST(i.object_id AS VARCHAR) + '_' + CAST(i.index_id AS VARCHAR) AND allobj.type = 'IT' INNER JOIN sys.partitions AS p ON p.object_id = CAST(( CASE WHEN i.type = 4 THEN allobj.object_id ELSE i.object_id END ) AS INT) AND p.index_id = CAST(( CASE WHEN i.type = 4 THEN 1 ELSE i.index_id END ) AS INT) LEFT OUTER JOIN sys.destination_data_spaces AS dds ON dds.partition_scheme_id = i.data_space_id AND dds.destination_id = p.partition_number LEFT OUTER JOIN sys.partition_schemes AS ps ON ps.data_space_id = i.data_space_id LEFT OUTER JOIN sys.partition_range_values AS prv ON prv.boundary_id = p.partition_number AND prv.function_id = ps.function_id LEFT OUTER JOIN sys.filegroups AS fg ON fg.data_space_id = dds.data_space_id OR fg.data_space_id = i.data_space_id LEFT OUTER JOIN sys.partition_functions AS pf ON pf.function_id = prv.function_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], p.partition_number AS [PartitionNumber], prv.value AS [RightBoundaryValue], fg.name AS [FileGroupName], CAST(pf.boundary_value_on_right AS INT) AS [RangeType], CAST(p.rows AS FLOAT) AS [RowCount], p.data_compression AS [DataCompression] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) LEFT OUTER JOIN sys.all_objects AS allobj ON allobj.name = 'extended_index_' + CAST(i.object_id AS VARCHAR) + '_' + CAST(i.index_id AS VARCHAR) AND allobj.type = 'IT' INNER JOIN sys.partitions AS p ON p.object_id = CAST(( CASE WHEN i.type = 4 THEN allobj.object_id ELSE i.object_id END ) AS INT) AND p.index_id = CAST(( CASE WHEN i.type = 4 THEN 1 ELSE i.index_id END ) AS INT) LEFT OUTER JOIN sys.destination_data_spaces AS dds ON dds.partition_scheme_id = i.data_space_id AND dds.destination_id = p.partition_number LEFT OUTER JOIN sys.partition_schemes AS ps ON ps.data_space_id = i.data_space_id LEFT OUTER JOIN sys.partition_range_values AS prv ON prv.boundary_id = p.partition_number AND prv.function_id = ps.function_id LEFT OUTER JOIN sys.filegroups AS fg ON fg.data_space_id = dds.data_space_id OR fg.data_space_id = i.data_space_id LEFT OUTER JOIN sys.partition_functions AS pf ON pf.function_id = prv.function_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], p.partition_number AS [PartitionNumber], prv.value AS [RightBoundaryValue], fg.name AS [FileGroupName], CAST(pf.boundary_value_on_right AS INT) AS [RangeType], CAST(p.rows AS FLOAT) AS [RowCount] FROM sys.all_views AS v INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = v.object_id ) INNER JOIN sys.partitions AS p ON p.object_id = CAST(i.object_id AS INT) AND p.index_id = CAST(i.index_id AS INT) LEFT OUTER JOIN sys.destination_data_spaces AS dds ON dds.partition_scheme_id = i.data_space_id AND dds.destination_id = p.partition_number LEFT OUTER JOIN sys.partition_schemes AS ps ON ps.data_space_id = i.data_space_id LEFT OUTER JOIN sys.partition_range_values AS prv ON prv.boundary_id = p.partition_number AND prv.function_id = ps.function_id LEFT OUTER JOIN sys.filegroups AS fg ON fg.data_space_id = dds.data_space_id OR fg.data_space_id = i.data_space_id LEFT OUTER JOIN sys.partition_functions AS pf ON pf.function_id = prv.function_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$null
		}
	}
}

function Get-ViewStatisticsQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], st.name AS [Name], st.stats_id AS [ID], st.no_recompute AS [NoAutomaticRecomputation], STATS_DATE(st.object_id, st.stats_id) AS [LastUpdated], CAST(1 - INDEXPROPERTY(st.object_id, st.name, N'IsStatistics') AS BIT) AS [IsFromIndexCreation], st.auto_created AS [IsAutoCreated], '' AS [FileGroup], st.has_filter AS [HasFilter], ISNULL(st.filter_definition, N'') AS [FilterDefinition], st.is_temporary AS [IsTemporary] FROM sys.all_views AS v INNER JOIN sys.stats st ON st.object_id = v.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = v.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], st.name AS [Name], st.stats_id AS [ID], st.no_recompute AS [NoAutomaticRecomputation], STATS_DATE(st.object_id, st.stats_id) AS [LastUpdated], CAST(1 - INDEXPROPERTY(st.object_id, st.name, N'IsStatistics') AS BIT) AS [IsFromIndexCreation], st.auto_created AS [IsAutoCreated], '' AS [FileGroup], st.has_filter AS [HasFilter], ISNULL(st.filter_definition, N'') AS [FilterDefinition], st.is_temporary AS [IsTemporary] FROM sys.all_views AS v INNER JOIN sys.stats st ON st.object_id = v.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], st.name AS [Name], st.stats_id AS [ID], st.no_recompute AS [NoAutomaticRecomputation], STATS_DATE(st.object_id, st.stats_id) AS [LastUpdated], CAST(1 - INDEXPROPERTY(st.object_id, st.name, N'IsStatistics') AS BIT) AS [IsFromIndexCreation], st.auto_created AS [IsAutoCreated], '' AS [FileGroup], st.has_filter AS [HasFilter], ISNULL(st.filter_definition, N'') AS [FilterDefinition] FROM sys.all_views AS v INNER JOIN sys.stats st ON st.object_id = v.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], st.name AS [Name], st.stats_id AS [ID], st.no_recompute AS [NoAutomaticRecomputation], STATS_DATE(st.object_id, st.stats_id) AS [LastUpdated], CAST(1 - INDEXPROPERTY(st.object_id, st.name, N'IsStatistics') AS BIT) AS [IsFromIndexCreation], st.auto_created AS [IsAutoCreated], '' AS [FileGroup], st.has_filter AS [HasFilter], ISNULL(st.filter_definition, N'') AS [FilterDefinition] FROM sys.all_views AS v INNER JOIN sys.stats st ON st.object_id = v.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], st.name AS [Name], st.stats_id AS [ID], st.no_recompute AS [NoAutomaticRecomputation], STATS_DATE(st.object_id, st.stats_id) AS [LastUpdated], CAST(1 - INDEXPROPERTY(st.object_id, st.name, N'IsStatistics') AS BIT) AS [IsFromIndexCreation], st.auto_created AS [IsAutoCreated], '' AS [FileGroup] FROM sys.all_views AS v INNER JOIN sys.stats st ON st.object_id = v.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(v.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(v.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT v.id AS [ViewID], sv.uid AS [SchemaID], st.name AS [Name], CAST(st.indid AS INT) AS [ID], CAST(CASE WHEN ( st.status & 16777216 ) <> 0 THEN 1 ELSE 0 END AS BIT) AS [NoAutomaticRecomputation], STATS_DATE(st.id, st.indid) AS [LastUpdated], CAST(1 - INDEXPROPERTY(st.id, st.name, N'IsStatistics') AS BIT) AS [IsFromIndexCreation], CAST(INDEXPROPERTY(st.id, st.name, N'IsAutoStatistics') AS BIT) AS [IsAutoCreated], '' AS [FileGroup] FROM dbo.sysobjects AS v INNER JOIN sysusers AS sv ON sv.uid = v.uid INNER JOIN dbo.sysindexes st ON ( ( st.indid <> 0 AND st.indid <> 255 ) AND 0 = OBJECTPROPERTY(st.id, N'IsMSShipped') ) AND ( st.id = v.id ) WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-ViewStatisticsColumnQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], st.stats_id AS [StatisticID], sic.stats_column_id AS [ID], COL_NAME(sic.object_id, sic.column_id) AS [Name] FROM sys.all_views AS v INNER JOIN sys.stats st ON st.object_id = v.object_id INNER JOIN sys.stats_columns sic ON sic.stats_id = st.stats_id AND sic.object_id = st.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = v.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], st.stats_id AS [StatisticID], sic.stats_column_id AS [ID], COL_NAME(sic.object_id, sic.column_id) AS [Name] FROM sys.all_views AS v INNER JOIN sys.stats st ON st.object_id = v.object_id INNER JOIN sys.stats_columns sic ON sic.stats_id = st.stats_id AND sic.object_id = st.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], st.stats_id AS [StatisticID], sic.stats_column_id AS [ID], COL_NAME(sic.object_id, sic.column_id) AS [Name] FROM sys.all_views AS v INNER JOIN sys.stats st ON st.object_id = v.object_id INNER JOIN sys.stats_columns sic ON sic.stats_id = st.stats_id AND sic.object_id = st.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], st.stats_id AS [StatisticID], sic.stats_column_id AS [ID], COL_NAME(sic.object_id, sic.column_id) AS [Name] FROM sys.all_views AS v INNER JOIN sys.stats st ON st.object_id = v.object_id INNER JOIN sys.stats_columns sic ON sic.stats_id = st.stats_id AND sic.object_id = st.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], st.stats_id AS [StatisticID], sic.stats_column_id AS [ID], COL_NAME(sic.object_id, sic.column_id) AS [Name] FROM sys.all_views AS v INNER JOIN sys.stats st ON st.object_id = v.object_id INNER JOIN sys.stats_columns sic ON sic.stats_id = st.stats_id AND sic.object_id = st.object_id WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(v.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(v.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT v.id AS [ViewID], sv.uid AS [SchemaID], CAST(st.indid AS INT) AS [StatisticID], CAST(c.keyno AS INT) AS [ID], clmns.name AS [Name] FROM dbo.sysobjects AS v INNER JOIN sysusers AS sv ON sv.uid = v.uid INNER JOIN dbo.sysindexes st ON ( ( st.indid <> 0 AND st.indid <> 255 ) AND 0 = OBJECTPROPERTY(st.id, N'IsMSShipped') ) AND ( st.id = v.id ) INNER JOIN dbo.sysindexkeys c ON c.indid = CAST(st.indid AS INT) AND c.id = st.id INNER JOIN dbo.syscolumns clmns ON clmns.id = c.id AND clmns.colid = c.colid WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-ViewTriggerQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], tr.name AS [Name], tr.object_id AS [ID], tr.create_date AS [CreateDate], tr.modify_date AS [DateLastModified], CAST(tr.is_ms_shipped AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(CASE WHEN ISNULL(smtr.definition, ssmtr.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CASE ISNULL(smtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END AS [ExecutionContext], ISNULL(USER_NAME(smtr.execute_as_principal_id), N'') AS [ExecutionContextPrincipal], ~trr.is_disabled AS [IsEnabled], trr.is_instead_of_trigger AS [InsteadOf], CAST(ISNULL(tei.object_id, 0) AS BIT) AS [Insert], CASE WHEN tei.is_first = 1 THEN 0 WHEN tei.is_last = 1 THEN 2 ELSE 1 END AS [InsertOrder], CAST(ISNULL(teu.object_id, 0) AS BIT) AS [Update], CASE WHEN teu.is_first = 1 THEN 0 WHEN teu.is_last = 1 THEN 2 ELSE 1 END AS [UpdateOrder], CAST(ISNULL(ted.object_id, 0) AS BIT) AS [Delete], CASE WHEN ted.is_first = 1 THEN 0 WHEN ted.is_last = 1 THEN 2 ELSE 1 END AS [DeleteOrder], CASE WHEN tr.type = N'TR' THEN 1 WHEN tr.type = N'TA' THEN 2 ELSE 1 END AS [ImplementationType], trr.is_not_for_replication AS [NotForReplication], NULL AS [Text], ISNULL(smtr.definition, ssmtr.definition) AS [Definition] FROM sys.all_views AS v INNER JOIN sys.objects AS tr ON ( tr.type IN ( 'TR', 'TA' ) ) AND ( tr.parent_object_id = v.object_id ) LEFT OUTER JOIN sys.sql_modules AS smtr ON smtr.object_id = tr.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmtr ON ssmtr.object_id = tr.object_id INNER JOIN sys.triggers AS trr ON trr.object_id = tr.object_id LEFT OUTER JOIN sys.trigger_events AS tei ON tei.object_id = tr.object_id AND tei.type = 1 LEFT OUTER JOIN sys.trigger_events AS teu ON teu.object_id = tr.object_id AND teu.type = 2 LEFT OUTER JOIN sys.trigger_events AS ted ON ted.object_id = tr.object_id AND ted.type = 3 WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN v.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = v.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], tr.name AS [Name], tr.object_id AS [ID], tr.create_date AS [CreateDate], tr.modify_date AS [DateLastModified], CAST(tr.is_ms_shipped AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(CASE WHEN ISNULL(smtr.definition, ssmtr.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE asmbltr.name END AS [AssemblyName], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE amtr.assembly_class END AS [ClassName], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE amtr.assembly_method END AS [MethodName], CASE WHEN amtr.object_id IS NULL THEN CASE ISNULL(smtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END ELSE CASE ISNULL(amtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END END AS [ExecutionContext], CASE WHEN amtr.object_id IS NULL THEN ISNULL(USER_NAME(smtr.execute_as_principal_id), N'') ELSE USER_NAME(amtr.execute_as_principal_id) END AS [ExecutionContextPrincipal], ~trr.is_disabled AS [IsEnabled], trr.is_instead_of_trigger AS [InsteadOf], CAST(ISNULL(tei.object_id, 0) AS BIT) AS [Insert], CASE WHEN tei.is_first = 1 THEN 0 WHEN tei.is_last = 1 THEN 2 ELSE 1 END AS [InsertOrder], CAST(ISNULL(teu.object_id, 0) AS BIT) AS [Update], CASE WHEN teu.is_first = 1 THEN 0 WHEN teu.is_last = 1 THEN 2 ELSE 1 END AS [UpdateOrder], CAST(ISNULL(ted.object_id, 0) AS BIT) AS [Delete], CASE WHEN ted.is_first = 1 THEN 0 WHEN ted.is_last = 1 THEN 2 ELSE 1 END AS [DeleteOrder], CASE WHEN tr.type = N'TR' THEN 1 WHEN tr.type = N'TA' THEN 2 ELSE 1 END AS [ImplementationType], trr.is_not_for_replication AS [NotForReplication], NULL AS [Text], ISNULL(smtr.definition, ssmtr.definition) AS [Definition] FROM sys.all_views AS v INNER JOIN sys.objects AS tr ON ( tr.type IN ( 'TR', 'TA' ) ) AND ( tr.parent_object_id = v.object_id ) LEFT OUTER JOIN sys.assembly_modules AS mod ON mod.object_id = tr.object_id LEFT OUTER JOIN sys.sql_modules AS smtr ON smtr.object_id = tr.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmtr ON ssmtr.object_id = tr.object_id LEFT OUTER JOIN sys.assemblies AS asmbl ON asmbl.assembly_id = mod.assembly_id LEFT OUTER JOIN sys.assembly_modules AS amtr ON amtr.object_id = tr.object_id LEFT OUTER JOIN sys.assemblies AS asmbltr ON asmbltr.assembly_id = amtr.assembly_id INNER JOIN sys.triggers AS trr ON trr.object_id = tr.object_id LEFT OUTER JOIN sys.trigger_events AS tei ON tei.object_id = tr.object_id AND tei.type = 1 LEFT OUTER JOIN sys.trigger_events AS teu ON teu.object_id = tr.object_id AND teu.type = 2 LEFT OUTER JOIN sys.trigger_events AS ted ON ted.object_id = tr.object_id AND ted.type = 3 WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], tr.name AS [Name], tr.object_id AS [ID], tr.create_date AS [CreateDate], tr.modify_date AS [DateLastModified], CAST(tr.is_ms_shipped AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(CASE WHEN ISNULL(smtr.definition, ssmtr.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE asmbltr.name END AS [AssemblyName], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE amtr.assembly_class END AS [ClassName], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE amtr.assembly_method END AS [MethodName], CASE WHEN amtr.object_id IS NULL THEN CASE ISNULL(smtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END ELSE CASE ISNULL(amtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END END AS [ExecutionContext], CASE WHEN amtr.object_id IS NULL THEN ISNULL(USER_NAME(smtr.execute_as_principal_id), N'') ELSE USER_NAME(amtr.execute_as_principal_id) END AS [ExecutionContextPrincipal], ~trr.is_disabled AS [IsEnabled], trr.is_instead_of_trigger AS [InsteadOf], CAST(ISNULL(tei.object_id, 0) AS BIT) AS [Insert], CASE WHEN tei.is_first = 1 THEN 0 WHEN tei.is_last = 1 THEN 2 ELSE 1 END AS [InsertOrder], CAST(ISNULL(teu.object_id, 0) AS BIT) AS [Update], CASE WHEN teu.is_first = 1 THEN 0 WHEN teu.is_last = 1 THEN 2 ELSE 1 END AS [UpdateOrder], CAST(ISNULL(ted.object_id, 0) AS BIT) AS [Delete], CASE WHEN ted.is_first = 1 THEN 0 WHEN ted.is_last = 1 THEN 2 ELSE 1 END AS [DeleteOrder], CASE WHEN tr.type = N'TR' THEN 1 WHEN tr.type = N'TA' THEN 2 ELSE 1 END AS [ImplementationType], trr.is_not_for_replication AS [NotForReplication], NULL AS [Text], ISNULL(smtr.definition, ssmtr.definition) AS [Definition] FROM sys.all_views AS v INNER JOIN sys.objects AS tr ON ( tr.type IN ( 'TR', 'TA' ) ) AND ( tr.parent_object_id = v.object_id ) LEFT OUTER JOIN sys.assembly_modules AS mod ON mod.object_id = tr.object_id LEFT OUTER JOIN sys.sql_modules AS smtr ON smtr.object_id = tr.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmtr ON ssmtr.object_id = tr.object_id LEFT OUTER JOIN sys.assemblies AS asmbl ON asmbl.assembly_id = mod.assembly_id LEFT OUTER JOIN sys.assembly_modules AS amtr ON amtr.object_id = tr.object_id LEFT OUTER JOIN sys.assemblies AS asmbltr ON asmbltr.assembly_id = amtr.assembly_id INNER JOIN sys.triggers AS trr ON trr.object_id = tr.object_id LEFT OUTER JOIN sys.trigger_events AS tei ON tei.object_id = tr.object_id AND tei.type = 1 LEFT OUTER JOIN sys.trigger_events AS teu ON teu.object_id = tr.object_id AND teu.type = 2 LEFT OUTER JOIN sys.trigger_events AS ted ON ted.object_id = tr.object_id AND ted.type = 3 WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], tr.name AS [Name], tr.object_id AS [ID], tr.create_date AS [CreateDate], tr.modify_date AS [DateLastModified], CAST(tr.is_ms_shipped AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(CASE WHEN ISNULL(smtr.definition, ssmtr.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE asmbltr.name END AS [AssemblyName], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE amtr.assembly_class END AS [ClassName], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE amtr.assembly_method END AS [MethodName], CASE WHEN amtr.object_id IS NULL THEN CASE ISNULL(smtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END ELSE CASE ISNULL(amtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END END AS [ExecutionContext], CASE WHEN amtr.object_id IS NULL THEN ISNULL(USER_NAME(smtr.execute_as_principal_id), N'') ELSE USER_NAME(amtr.execute_as_principal_id) END AS [ExecutionContextPrincipal], ~trr.is_disabled AS [IsEnabled], trr.is_instead_of_trigger AS [InsteadOf], CAST(ISNULL(tei.object_id, 0) AS BIT) AS [Insert], CASE WHEN tei.is_first = 1 THEN 0 WHEN tei.is_last = 1 THEN 2 ELSE 1 END AS [InsertOrder], CAST(ISNULL(teu.object_id, 0) AS BIT) AS [Update], CASE WHEN teu.is_first = 1 THEN 0 WHEN teu.is_last = 1 THEN 2 ELSE 1 END AS [UpdateOrder], CAST(ISNULL(ted.object_id, 0) AS BIT) AS [Delete], CASE WHEN ted.is_first = 1 THEN 0 WHEN ted.is_last = 1 THEN 2 ELSE 1 END AS [DeleteOrder], CASE WHEN tr.type = N'TR' THEN 1 WHEN tr.type = N'TA' THEN 2 ELSE 1 END AS [ImplementationType], trr.is_not_for_replication AS [NotForReplication], NULL AS [Text], ISNULL(smtr.definition, ssmtr.definition) AS [Definition] FROM sys.all_views AS v INNER JOIN sys.objects AS tr ON ( tr.type IN ( 'TR', 'TA' ) ) AND ( tr.parent_object_id = v.object_id ) LEFT OUTER JOIN sys.assembly_modules AS mod ON mod.object_id = tr.object_id LEFT OUTER JOIN sys.sql_modules AS smtr ON smtr.object_id = tr.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmtr ON ssmtr.object_id = tr.object_id LEFT OUTER JOIN sys.assemblies AS asmbl ON asmbl.assembly_id = mod.assembly_id LEFT OUTER JOIN sys.assembly_modules AS amtr ON amtr.object_id = tr.object_id LEFT OUTER JOIN sys.assemblies AS asmbltr ON asmbltr.assembly_id = amtr.assembly_id INNER JOIN sys.triggers AS trr ON trr.object_id = tr.object_id LEFT OUTER JOIN sys.trigger_events AS tei ON tei.object_id = tr.object_id AND tei.type = 1 LEFT OUTER JOIN sys.trigger_events AS teu ON teu.object_id = tr.object_id AND teu.type = 2 LEFT OUTER JOIN sys.trigger_events AS ted ON ted.object_id = tr.object_id AND ted.type = 3 WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT v.object_id AS [ViewID], v.schema_id AS [SchemaID], tr.name AS [Name], tr.object_id AS [ID], tr.create_date AS [CreateDate], tr.modify_date AS [DateLastModified], CAST(tr.is_ms_shipped AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(tr.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(CASE WHEN ISNULL(smtr.definition, ssmtr.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE asmbltr.name END AS [AssemblyName], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE amtr.assembly_class END AS [ClassName], CASE WHEN amtr.object_id IS NULL THEN N'' ELSE amtr.assembly_method END AS [MethodName], CASE WHEN amtr.object_id IS NULL THEN CASE ISNULL(smtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END ELSE CASE ISNULL(amtr.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END END AS [ExecutionContext], CASE WHEN amtr.object_id IS NULL THEN ISNULL(USER_NAME(smtr.execute_as_principal_id), N'') ELSE USER_NAME(amtr.execute_as_principal_id) END AS [ExecutionContextPrincipal], ~trr.is_disabled AS [IsEnabled], trr.is_instead_of_trigger AS [InsteadOf], CAST(ISNULL(tei.object_id, 0) AS BIT) AS [Insert], CASE WHEN tei.is_first = 1 THEN 0 WHEN tei.is_last = 1 THEN 2 ELSE 1 END AS [InsertOrder], CAST(ISNULL(teu.object_id, 0) AS BIT) AS [Update], CASE WHEN teu.is_first = 1 THEN 0 WHEN teu.is_last = 1 THEN 2 ELSE 1 END AS [UpdateOrder], CAST(ISNULL(ted.object_id, 0) AS BIT) AS [Delete], CASE WHEN ted.is_first = 1 THEN 0 WHEN ted.is_last = 1 THEN 2 ELSE 1 END AS [DeleteOrder], CASE WHEN tr.type = N'TR' THEN 1 WHEN tr.type = N'TA' THEN 2 ELSE 1 END AS [ImplementationType], trr.is_not_for_replication AS [NotForReplication], NULL AS [Text], ISNULL(smtr.definition, ssmtr.definition) AS [Definition] FROM sys.all_views AS v INNER JOIN sys.objects AS tr ON ( tr.type IN ( 'TR', 'TA' ) ) AND ( tr.parent_object_id = v.object_id ) LEFT OUTER JOIN sys.assembly_modules AS mod ON mod.object_id = tr.object_id LEFT OUTER JOIN sys.sql_modules AS smtr ON smtr.object_id = tr.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmtr ON ssmtr.object_id = tr.object_id LEFT OUTER JOIN sys.assemblies AS asmbl ON asmbl.assembly_id = mod.assembly_id LEFT OUTER JOIN sys.assembly_modules AS amtr ON amtr.object_id = tr.object_id LEFT OUTER JOIN sys.assemblies AS asmbltr ON asmbltr.assembly_id = amtr.assembly_id INNER JOIN sys.triggers AS trr ON trr.object_id = tr.object_id LEFT OUTER JOIN sys.trigger_events AS tei ON tei.object_id = tr.object_id AND tei.type = 1 LEFT OUTER JOIN sys.trigger_events AS teu ON teu.object_id = tr.object_id AND teu.type = 2 LEFT OUTER JOIN sys.trigger_events AS ted ON ted.object_id = tr.object_id AND ted.type = 3 WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(v.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(v.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT v.id AS [ViewID], sv.uid AS [SchemaID], tr.name AS [Name], tr.id AS [ID], tr.crdate AS [CreateDate], CAST(CASE WHEN ( OBJECTPROPERTY(tr.id, N'IsMSShipped') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(tr.id, N'IsSystemTable') THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(OBJECTPROPERTY(tr.id, N'ExecIsAnsiNullsOn') AS BIT) AS [AnsiNullsStatus], CAST(OBJECTPROPERTY(tr.id, N'ExecIsQuotedIdentOn') AS BIT) AS [QuotedIdentifierStatus], CAST((SELECT TOP 1 encrypted FROM dbo.syscomments p WHERE tr.id = p.id AND p.colid = 1 AND p.number < 2) AS BIT) AS [IsEncrypted], CAST(1 - OBJECTPROPERTY(tr.id, N'ExecIsTriggerDisabled') AS BIT) AS [IsEnabled], CAST(OBJECTPROPERTY(tr.id, N'ExecIsInsteadOfTrigger') AS BIT) AS [InsteadOf], CAST(OBJECTPROPERTY(tr.id, N'ExecIsInsertTrigger') AS BIT) AS [Insert], CASE WHEN OBJECTPROPERTY(tr.id, N'ExecIsFirstInsertTrigger') = 1 THEN 0 WHEN OBJECTPROPERTY(tr.id, N'ExecIsLastInsertTrigger') = 1 THEN 2 ELSE 1 END AS [InsertOrder], CAST(OBJECTPROPERTY(tr.id, N'ExecIsUpdateTrigger') AS BIT) AS [Update], CASE WHEN OBJECTPROPERTY(tr.id, N'ExecIsFirstUpdateTrigger') = 1 THEN 0 WHEN OBJECTPROPERTY(tr.id, N'ExecIsLastUpdateTrigger') = 1 THEN 2 ELSE 1 END AS [UpdateOrder], CAST(OBJECTPROPERTY(tr.id, N'ExecIsDeleteTrigger') AS BIT) AS [Delete], CASE WHEN OBJECTPROPERTY(tr.id, N'ExecIsFirstDeleteTrigger') = 1 THEN 0 WHEN OBJECTPROPERTY(tr.id, N'ExecIsLastDeleteTrigger') = 1 THEN 2 ELSE 1 END AS [DeleteOrder], CAST(OBJECTPROPERTY(tr.id, N'ExecIsTriggerNotForRepl') AS BIT) AS [NotForReplication], 1 AS [ImplementationType], c.text AS [Definition] FROM dbo.sysobjects AS v INNER JOIN sysusers AS sv ON sv.uid = v.uid INNER JOIN dbo.sysobjects AS tr ON ( tr.type = 'TR' ) AND ( tr.parent_obj = v.id ) LEFT OUTER JOIN dbo.syscomments c ON c.id = tr.id AND CASE WHEN c.number > 1 THEN c.number ELSE 0 END = 0 WHERE ( v.type = 'V' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-StoredProcedureQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN sp.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT sp.name AS [Name], sp.object_id AS [ID], sp.create_date AS [CreateDate], sp.modify_date AS [DateLastModified], ISNULL(ssp.name, N'') AS [Owner], CAST(CASE WHEN sp.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(sp.schema_id) AS [Schema], sp.schema_id AS [SchemaID], CAST(CASE WHEN sp.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(sp.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(sp.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(CASE WHEN ISNULL(smsp.definition, ssmsp.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CAST(ISNULL(smsp.is_recompiled, ssmsp.is_recompiled) AS BIT) AS [Recompile], CASE ISNULL(smsp.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END AS [ExecutionContext], ISNULL(USER_NAME(smsp.execute_as_principal_id), N'') AS [ExecutionContextPrincipal], CAST(ISNULL(spp.is_auto_executed, 0) AS BIT) AS [Startup], CASE WHEN sp.type = N'P' THEN 1 WHEN sp.type = N'PC' THEN 2 ELSE 1 END AS [ImplementationType], CAST(CASE sp.type WHEN N'RF' THEN 1 ELSE 0 END AS BIT) AS [ForReplication], NULL AS [Text], ISNULL(smsp.definition, ssmsp.definition) AS [Definition] FROM sys.all_objects AS sp LEFT OUTER JOIN sys.database_principals AS ssp ON ssp.principal_id = ISNULL(sp.principal_id, ( OBJECTPROPERTY(sp.object_id, 'OwnerId') )) LEFT OUTER JOIN sys.sql_modules AS smsp ON smsp.object_id = sp.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmsp ON ssmsp.object_id = sp.object_id LEFT OUTER JOIN sys.procedures AS spp ON spp.object_id = sp.object_id WHERE ( sp.type = 'P' OR sp.type = 'RF' OR sp.type = 'PC' ) 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN sp.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = sp.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT sp.name AS [Name], sp.object_id AS [ID], sp.create_date AS [CreateDate], sp.modify_date AS [DateLastModified], ISNULL(ssp.name, N'') AS [Owner], CAST(CASE WHEN sp.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(sp.schema_id) AS [Schema], sp.schema_id AS [SchemaID], CAST(CASE WHEN sp.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = sp.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(sp.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(sp.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(CASE WHEN ISNULL(smsp.definition, ssmsp.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CAST(ISNULL(smsp.is_recompiled, ssmsp.is_recompiled) AS BIT) AS [Recompile], CASE WHEN amsp.object_id IS NULL THEN N'' ELSE asmblsp.name END AS [AssemblyName], CASE WHEN amsp.object_id IS NULL THEN N'' ELSE amsp.assembly_class END AS [ClassName], CASE WHEN amsp.object_id IS NULL THEN N'' ELSE amsp.assembly_method END AS [MethodName], CASE WHEN amsp.object_id IS NULL THEN CASE ISNULL(smsp.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END ELSE CASE ISNULL(amsp.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END END AS [ExecutionContext], CASE WHEN amsp.object_id IS NULL THEN ISNULL(USER_NAME(smsp.execute_as_principal_id), N'') ELSE USER_NAME(amsp.execute_as_principal_id) END AS [ExecutionContextPrincipal], CAST(ISNULL(spp.is_auto_executed, 0) AS BIT) AS [Startup], CASE WHEN sp.type = N'P' THEN 1 WHEN sp.type = N'PC' THEN 2 ELSE 1 END AS [ImplementationType], CAST(CASE sp.type WHEN N'RF' THEN 1 ELSE 0 END AS BIT) AS [ForReplication], NULL AS [Text], ISNULL(smsp.definition, ssmsp.definition) AS [Definition] FROM sys.all_objects AS sp LEFT OUTER JOIN sys.database_principals AS ssp ON ssp.principal_id = ISNULL(sp.principal_id, ( OBJECTPROPERTY(sp.object_id, 'OwnerId') )) LEFT OUTER JOIN sys.sql_modules AS smsp ON smsp.object_id = sp.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmsp ON ssmsp.object_id = sp.object_id LEFT OUTER JOIN sys.assembly_modules AS amsp ON amsp.object_id = sp.object_id LEFT OUTER JOIN sys.assemblies AS asmblsp ON asmblsp.assembly_id = amsp.assembly_id LEFT OUTER JOIN sys.procedures AS spp ON spp.object_id = sp.object_id WHERE ( sp.type = 'P' OR sp.type = 'RF' OR sp.type = 'PC' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT sp.name AS [Name], sp.object_id AS [ID], sp.create_date AS [CreateDate], sp.modify_date AS [DateLastModified], ISNULL(ssp.name, N'') AS [Owner], CAST(CASE WHEN sp.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(sp.schema_id) AS [Schema], sp.schema_id AS [SchemaID], CAST(CASE WHEN sp.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = sp.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(sp.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(sp.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(CASE WHEN ISNULL(smsp.definition, ssmsp.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CAST(ISNULL(smsp.is_recompiled, ssmsp.is_recompiled) AS BIT) AS [Recompile], CASE WHEN amsp.object_id IS NULL THEN N'' ELSE asmblsp.name END AS [AssemblyName], CASE WHEN amsp.object_id IS NULL THEN N'' ELSE amsp.assembly_class END AS [ClassName], CASE WHEN amsp.object_id IS NULL THEN N'' ELSE amsp.assembly_method END AS [MethodName], CASE WHEN amsp.object_id IS NULL THEN CASE ISNULL(smsp.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END ELSE CASE ISNULL(amsp.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END END AS [ExecutionContext], CASE WHEN amsp.object_id IS NULL THEN ISNULL(USER_NAME(smsp.execute_as_principal_id), N'') ELSE USER_NAME(amsp.execute_as_principal_id) END AS [ExecutionContextPrincipal], CAST(ISNULL(spp.is_auto_executed, 0) AS BIT) AS [Startup], CASE WHEN sp.type = N'P' THEN 1 WHEN sp.type = N'PC' THEN 2 ELSE 1 END AS [ImplementationType], CAST(CASE sp.type WHEN N'RF' THEN 1 ELSE 0 END AS BIT) AS [ForReplication], NULL AS [Text], ISNULL(smsp.definition, ssmsp.definition) AS [Definition] FROM sys.all_objects AS sp LEFT OUTER JOIN sys.database_principals AS ssp ON ssp.principal_id = ISNULL(sp.principal_id, ( OBJECTPROPERTY(sp.object_id, 'OwnerId') )) LEFT OUTER JOIN sys.sql_modules AS smsp ON smsp.object_id = sp.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmsp ON ssmsp.object_id = sp.object_id LEFT OUTER JOIN sys.assembly_modules AS amsp ON amsp.object_id = sp.object_id LEFT OUTER JOIN sys.assemblies AS asmblsp ON asmblsp.assembly_id = amsp.assembly_id LEFT OUTER JOIN sys.procedures AS spp ON spp.object_id = sp.object_id WHERE ( sp.type = 'P' OR sp.type = 'RF' OR sp.type = 'PC' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT sp.name AS [Name], sp.object_id AS [ID], sp.create_date AS [CreateDate], sp.modify_date AS [DateLastModified], ISNULL(ssp.name, N'') AS [Owner], CAST(CASE WHEN sp.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(sp.schema_id) AS [Schema], sp.schema_id AS [SchemaID], CAST(CASE WHEN sp.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = sp.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(sp.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(sp.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(CASE WHEN ISNULL(smsp.definition, ssmsp.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CAST(ISNULL(smsp.is_recompiled, ssmsp.is_recompiled) AS BIT) AS [Recompile], CASE WHEN amsp.object_id IS NULL THEN N'' ELSE asmblsp.name END AS [AssemblyName], CASE WHEN amsp.object_id IS NULL THEN N'' ELSE amsp.assembly_class END AS [ClassName], CASE WHEN amsp.object_id IS NULL THEN N'' ELSE amsp.assembly_method END AS [MethodName], CASE WHEN amsp.object_id IS NULL THEN CASE ISNULL(smsp.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END ELSE CASE ISNULL(amsp.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END END AS [ExecutionContext], CASE WHEN amsp.object_id IS NULL THEN ISNULL(USER_NAME(smsp.execute_as_principal_id), N'') ELSE USER_NAME(amsp.execute_as_principal_id) END AS [ExecutionContextPrincipal], CAST(ISNULL(spp.is_auto_executed, 0) AS BIT) AS [Startup], CASE WHEN sp.type = N'P' THEN 1 WHEN sp.type = N'PC' THEN 2 ELSE 1 END AS [ImplementationType], CAST(CASE sp.type WHEN N'RF' THEN 1 ELSE 0 END AS BIT) AS [ForReplication], NULL AS [Text], ISNULL(smsp.definition, ssmsp.definition) AS [Definition] FROM sys.all_objects AS sp LEFT OUTER JOIN sys.database_principals AS ssp ON ssp.principal_id = ISNULL(sp.principal_id, ( OBJECTPROPERTY(sp.object_id, 'OwnerId') )) LEFT OUTER JOIN sys.sql_modules AS smsp ON smsp.object_id = sp.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmsp ON ssmsp.object_id = sp.object_id LEFT OUTER JOIN sys.assembly_modules AS amsp ON amsp.object_id = sp.object_id LEFT OUTER JOIN sys.assemblies AS asmblsp ON asmblsp.assembly_id = amsp.assembly_id LEFT OUTER JOIN sys.procedures AS spp ON spp.object_id = sp.object_id WHERE ( sp.type = 'P' OR sp.type = 'RF' OR sp.type = 'PC' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT sp.name AS [Name], sp.object_id AS [ID], sp.create_date AS [CreateDate], sp.modify_date AS [DateLastModified], ISNULL(ssp.name, N'') AS [Owner], CAST(CASE WHEN sp.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(sp.schema_id) AS [Schema], sp.schema_id AS [SchemaID], CAST(CASE WHEN sp.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = sp.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(ISNULL(OBJECTPROPERTYEX(sp.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(sp.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], CAST(CASE WHEN ISNULL(smsp.definition, ssmsp.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CAST(ISNULL(smsp.is_recompiled, ssmsp.is_recompiled) AS BIT) AS [Recompile], CASE WHEN amsp.object_id IS NULL THEN N'' ELSE asmblsp.name END AS [AssemblyName], CASE WHEN amsp.object_id IS NULL THEN N'' ELSE amsp.assembly_class END AS [ClassName], CASE WHEN amsp.object_id IS NULL THEN N'' ELSE amsp.assembly_method END AS [MethodName], CASE WHEN amsp.object_id IS NULL THEN CASE ISNULL(smsp.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END ELSE CASE ISNULL(amsp.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END END AS [ExecutionContext], CASE WHEN amsp.object_id IS NULL THEN ISNULL(USER_NAME(smsp.execute_as_principal_id), N'') ELSE USER_NAME(amsp.execute_as_principal_id) END AS [ExecutionContextPrincipal], CAST(ISNULL(spp.is_auto_executed, 0) AS BIT) AS [Startup], CASE WHEN sp.type = N'P' THEN 1 WHEN sp.type = N'PC' THEN 2 ELSE 1 END AS [ImplementationType], CAST(CASE sp.type WHEN N'RF' THEN 1 ELSE 0 END AS BIT) AS [ForReplication], NULL AS [Text], ISNULL(smsp.definition, ssmsp.definition) AS [Definition] FROM sys.all_objects AS sp LEFT OUTER JOIN sys.database_principals AS ssp ON ssp.principal_id = ISNULL(sp.principal_id, ( OBJECTPROPERTY(sp.object_id, 'OwnerId') )) LEFT OUTER JOIN sys.sql_modules AS smsp ON smsp.object_id = sp.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmsp ON ssmsp.object_id = sp.object_id LEFT OUTER JOIN sys.assembly_modules AS amsp ON amsp.object_id = sp.object_id LEFT OUTER JOIN sys.assemblies AS asmblsp ON asmblsp.assembly_id = amsp.assembly_id LEFT OUTER JOIN sys.procedures AS spp ON spp.object_id = sp.object_id WHERE ( sp.type = 'P' OR sp.type = 'RF' OR sp.type = 'PC' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(sp.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(sp.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT sp.name AS [Name], sp.id AS [ID], sp.crdate AS [CreateDate], ssp.name AS [Schema], ssp.uid AS [SchemaID], ssp.name AS [Owner], CAST(CASE WHEN ( OBJECTPROPERTY(sp.id, N'IsMSShipped') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(sp.id, N'IsSystemTable') THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], CAST(OBJECTPROPERTY(sp.id, N'ExecIsAnsiNullsOn') AS BIT) AS [AnsiNullsStatus], CAST(OBJECTPROPERTY(sp.id, N'ExecIsQuotedIdentOn') AS BIT) AS [QuotedIdentifierStatus], CAST((SELECT TOP 1 encrypted FROM dbo.syscomments p WHERE sp.id = p.id AND p.colid = 1 AND p.number < 2) AS BIT) AS [IsEncrypted], CAST(sp.status & 4 AS BIT) AS [Recompile], CAST(OBJECTPROPERTY(sp.id, N'ExecIsStartup') AS BIT) AS [Startup], CAST(CASE sp.xtype WHEN N'RF' THEN 1 ELSE 0 END AS BIT) AS [ForReplication], 1 AS [ImplementationType], c.text AS [Definition] FROM dbo.sysobjects AS sp INNER JOIN sysusers AS ssp ON ssp.uid = sp.uid LEFT OUTER JOIN dbo.syscomments c ON c.id = sp.id AND CASE WHEN c.number > 1 THEN c.number ELSE 0 END = 0 WHERE ( sp.xtype = 'P' OR sp.xtype = 'RF' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-StoredProcedureParameterQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN sp.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT sp.object_id AS [StoredProcedureID], sp.schema_id AS [SchemaID], param.name AS [Name], param.parameter_id AS [ID], param.default_value AS [DefaultValue], param.has_default_value AS [HasDefaultValue], usrt.name AS [DataType], s1param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND param.max_length <> -1 THEN param.max_length / 2 ELSE param.max_length END AS INT) AS [Length], CAST(param.precision AS INT) AS [NumericPrecision], CAST(param.scale AS INT) AS [NumericScale], ISNULL(xscparam.name, N'') AS [XmlSchemaNamespace], ISNULL(s2param.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE param.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], param.is_output AS [IsOutputParameter], param.is_cursor_ref AS [IsCursorParameter], param.is_readonly AS [IsReadOnly], sp.object_id AS [IDText], DB_NAME() AS [DatabaseName], param.name AS [ParamName], CAST(CASE WHEN sp.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) AS [ParentSysObj], 1 AS [Number] FROM sys.all_objects AS sp INNER JOIN sys.all_parameters AS param ON param.object_id = sp.object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = param.user_type_id LEFT OUTER JOIN sys.schemas AS s1param ON s1param.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = param.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = param.system_type_id ) AND ( baset.user_type_id = param.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscparam ON xscparam.xml_collection_id = param.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2param ON s2param.schema_id = xscparam.schema_id WHERE ( sp.type = 'P' OR sp.type = 'RF' OR sp.type = 'PC' ) 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN sp.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = sp.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT sp.object_id AS [StoredProcedureID], sp.schema_id AS [SchemaID], param.name AS [Name], param.parameter_id AS [ID], param.default_value AS [DefaultValue], param.has_default_value AS [HasDefaultValue], usrt.name AS [DataType], s1param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND param.max_length <> -1 THEN param.max_length / 2 ELSE param.max_length END AS INT) AS [Length], CAST(param.precision AS INT) AS [NumericPrecision], CAST(param.scale AS INT) AS [NumericScale], ISNULL(xscparam.name, N'') AS [XmlSchemaNamespace], ISNULL(s2param.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE param.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], param.is_output AS [IsOutputParameter], param.is_cursor_ref AS [IsCursorParameter], param.is_readonly AS [IsReadOnly], sp.object_id AS [IDText], DB_NAME() AS [DatabaseName], param.name AS [ParamName], CAST(CASE WHEN sp.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = sp.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [ParentSysObj], 1 AS [Number] FROM sys.all_objects AS sp INNER JOIN sys.all_parameters AS param ON param.object_id = sp.object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = param.user_type_id LEFT OUTER JOIN sys.schemas AS s1param ON s1param.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = param.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = param.system_type_id ) AND ( baset.user_type_id = param.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscparam ON xscparam.xml_collection_id = param.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2param ON s2param.schema_id = xscparam.schema_id WHERE ( sp.type = 'P' OR sp.type = 'RF' OR sp.type = 'PC' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT sp.object_id AS [StoredProcedureID], sp.schema_id AS [SchemaID], param.name AS [Name], param.parameter_id AS [ID], param.default_value AS [DefaultValue], param.has_default_value AS [HasDefaultValue], usrt.name AS [DataType], s1param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND param.max_length <> -1 THEN param.max_length / 2 ELSE param.max_length END AS INT) AS [Length], CAST(param.precision AS INT) AS [NumericPrecision], CAST(param.scale AS INT) AS [NumericScale], ISNULL(xscparam.name, N'') AS [XmlSchemaNamespace], ISNULL(s2param.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE param.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], param.is_output AS [IsOutputParameter], param.is_cursor_ref AS [IsCursorParameter], param.is_readonly AS [IsReadOnly], sp.object_id AS [IDText], DB_NAME() AS [DatabaseName], param.name AS [ParamName], CAST(CASE WHEN sp.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = sp.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [ParentSysObj], 1 AS [Number] FROM sys.all_objects AS sp INNER JOIN sys.all_parameters AS param ON param.object_id = sp.object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = param.user_type_id LEFT OUTER JOIN sys.schemas AS s1param ON s1param.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = param.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = param.system_type_id ) AND ( baset.user_type_id = param.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscparam ON xscparam.xml_collection_id = param.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2param ON s2param.schema_id = xscparam.schema_id WHERE ( sp.type = 'P' OR sp.type = 'RF' OR sp.type = 'PC' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT sp.object_id AS [StoredProcedureID], sp.schema_id AS [SchemaID], param.name AS [Name], param.parameter_id AS [ID], param.default_value AS [DefaultValue], param.has_default_value AS [HasDefaultValue], usrt.name AS [DataType], s1param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND param.max_length <> -1 THEN param.max_length / 2 ELSE param.max_length END AS INT) AS [Length], CAST(param.precision AS INT) AS [NumericPrecision], CAST(param.scale AS INT) AS [NumericScale], ISNULL(xscparam.name, N'') AS [XmlSchemaNamespace], ISNULL(s2param.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE param.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], param.is_output AS [IsOutputParameter], param.is_cursor_ref AS [IsCursorParameter], param.is_readonly AS [IsReadOnly], sp.object_id AS [IDText], DB_NAME() AS [DatabaseName], param.name AS [ParamName], CAST(CASE WHEN sp.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = sp.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [ParentSysObj], 1 AS [Number] FROM sys.all_objects AS sp INNER JOIN sys.all_parameters AS param ON param.object_id = sp.object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = param.user_type_id LEFT OUTER JOIN sys.schemas AS s1param ON s1param.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = param.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = param.system_type_id ) AND ( baset.user_type_id = param.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscparam ON xscparam.xml_collection_id = param.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2param ON s2param.schema_id = xscparam.schema_id WHERE ( sp.type = 'P' OR sp.type = 'RF' OR sp.type = 'PC' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT sp.object_id AS [StoredProcedureID], sp.schema_id AS [SchemaID], param.name AS [Name], param.parameter_id AS [ID], param.default_value AS [DefaultValue], param.has_default_value AS [HasDefaultValue], usrt.name AS [DataType], s1param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND param.max_length <> -1 THEN param.max_length / 2 ELSE param.max_length END AS INT) AS [Length], CAST(param.precision AS INT) AS [NumericPrecision], CAST(param.scale AS INT) AS [NumericScale], ISNULL(xscparam.name, N'') AS [XmlSchemaNamespace], ISNULL(s2param.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE param.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], param.is_output AS [IsOutputParameter], param.is_cursor_ref AS [IsCursorParameter], sp.object_id AS [IDText], DB_NAME() AS [DatabaseName], param.name AS [ParamName], CAST(CASE WHEN sp.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = sp.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [ParentSysObj], 1 AS [Number] FROM sys.all_objects AS sp INNER JOIN sys.all_parameters AS param ON param.object_id = sp.object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = param.user_type_id LEFT OUTER JOIN sys.schemas AS s1param ON s1param.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = param.system_type_id AND baset.user_type_id = baset.system_type_id ) LEFT OUTER JOIN sys.xml_schema_collections AS xscparam ON xscparam.xml_collection_id = param.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2param ON s2param.schema_id = xscparam.schema_id WHERE ( sp.type = 'P' OR sp.type = 'RF' OR sp.type = 'PC' ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(sp.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(sp.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT sp.id AS [StoredProcedureID], ssp.uid AS [SchemaID], param.name AS [Name], CAST(param.colid AS INT) AS [ID], NULL AS [DefaultValue], usrt.name AS [DataType], s1param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'char', N'varchar', N'binary', N'varbinary', N'nchar', N'nvarchar' ) THEN param.prec ELSE param.length END AS INT) AS [Length], CAST(param.xprec AS INT) AS [NumericPrecision], CAST(param.xscale AS INT) AS [NumericScale], CAST(CASE param.isoutparam WHEN 1 THEN param.isoutparam WHEN 0 THEN CASE param.name WHEN '' THEN 1 ELSE 0 END END AS BIT) AS [IsOutputParameter], sp.id AS [IDText], DB_NAME() AS [DatabaseName], param.name AS [ParamName], CAST(CASE WHEN ( OBJECTPROPERTY(sp.id, N'IsMSShipped') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(sp.id, N'IsSystemTable') THEN 1 ELSE 0 END AS BIT) AS [ParentSysObj], 1 AS [Number] FROM dbo.sysobjects AS sp INNER JOIN sysusers AS ssp ON ssp.uid = sp.uid INNER JOIN syscolumns AS param ON ( param.number = 1 ) AND ( param.id = sp.id ) LEFT OUTER JOIN systypes AS usrt ON usrt.xusertype = param.xusertype LEFT OUTER JOIN sysusers AS s1param ON s1param.uid = usrt.uid LEFT OUTER JOIN systypes AS baset ON baset.xusertype = param.xtype AND baset.xusertype = baset.xtype WHERE ( sp.xtype = 'P' OR sp.xtype = 'RF' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-UserDefinedFunctionQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT udf.name AS [Name], udf.object_id AS [ID], udf.create_date AS [CreateDate], udf.modify_date AS [DateLastModified], ISNULL(sudf.name, N'') AS [Owner], CAST(CASE WHEN udf.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(udf.schema_id) AS [Schema], udf.schema_id AS [SchemaID], CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], usrt.name AS [DataType], s1ret_param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND ret_param.max_length <> -1 THEN ret_param.max_length / 2 ELSE ret_param.max_length END AS INT) AS [Length], CAST(ret_param.precision AS INT) AS [NumericPrecision], CAST(ret_param.scale AS INT) AS [NumericScale], ISNULL(xscret_param.name, N'') AS [XmlSchemaNamespace], ISNULL(s2ret_param.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE ret_param.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], CAST(ISNULL(OBJECTPROPERTYEX(udf.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(udf.object_id, N'IsSchemaBound'), 0) AS BIT) AS [IsSchemaBound], CAST(CASE WHEN ISNULL(smudf.definition, ssmudf.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CAST(CAST(smudf.null_on_null_input AS BIT) AS BIT) AS [ReturnsNullOnNullInput], CASE ISNULL(smudf.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END AS [ExecutionContext], ISNULL(USER_NAME(smudf.execute_as_principal_id), N'') AS [ExecutionContextPrincipal], CAST(OBJECTPROPERTYEX(udf.object_id, N'IsDeterministic') AS BIT) AS [IsDeterministic], ( CASE WHEN 'FN' = udf.type THEN 1 WHEN 'FS' = udf.type THEN 1 WHEN 'IF' = udf.type THEN 3 WHEN 'TF' = udf.type THEN 2 WHEN 'FT' = udf.type THEN 2 ELSE 0 END ) AS [FunctionType], CASE WHEN udf.type IN ( 'FN', 'IF', 'TF' ) THEN 1 WHEN udf.type IN ( 'FS', 'FT' ) THEN 2 ELSE 1 END AS [ImplementationType], CAST(ISNULL(OBJECTPROPERTYEX(udf.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], ret_param.name AS [TableVariableName], ISNULL(smudf.definition, ssmudf.definition) AS [Definition] FROM sys.all_objects AS udf LEFT OUTER JOIN sys.database_principals AS sudf ON sudf.principal_id = ISNULL(udf.principal_id, ( OBJECTPROPERTY(udf.object_id, 'OwnerId') )) LEFT OUTER JOIN sys.all_parameters AS ret_param ON ret_param.object_id = udf.object_id AND ret_param.is_output = 1 LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = ret_param.user_type_id LEFT OUTER JOIN sys.schemas AS s1ret_param ON s1ret_param.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = ret_param.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = ret_param.system_type_id ) AND ( baset.user_type_id = ret_param.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscret_param ON xscret_param.xml_collection_id = ret_param.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2ret_param ON s2ret_param.schema_id = xscret_param.schema_id LEFT OUTER JOIN sys.sql_modules AS smudf ON smudf.object_id = udf.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmudf ON ssmudf.object_id = udf.object_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = udf.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT udf.name AS [Name], udf.object_id AS [ID], udf.create_date AS [CreateDate], udf.modify_date AS [DateLastModified], ISNULL(sudf.name, N'') AS [Owner], CAST(CASE WHEN udf.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(udf.schema_id) AS [Schema], udf.schema_id AS [SchemaID], CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = udf.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], usrt.name AS [DataType], s1ret_param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND ret_param.max_length <> -1 THEN ret_param.max_length / 2 ELSE ret_param.max_length END AS INT) AS [Length], CAST(ret_param.precision AS INT) AS [NumericPrecision], CAST(ret_param.scale AS INT) AS [NumericScale], ISNULL(xscret_param.name, N'') AS [XmlSchemaNamespace], ISNULL(s2ret_param.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE ret_param.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], CAST(ISNULL(OBJECTPROPERTYEX(udf.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(udf.object_id, N'IsSchemaBound'), 0) AS BIT) AS [IsSchemaBound], CAST(CASE WHEN ISNULL(smudf.definition, ssmudf.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CASE WHEN amudf.object_id IS NULL THEN N'' ELSE asmbludf.name END AS [AssemblyName], CASE WHEN amudf.object_id IS NULL THEN N'' ELSE amudf.assembly_class END AS [ClassName], CASE WHEN amudf.object_id IS NULL THEN N'' ELSE amudf.assembly_method END AS [MethodName], CAST(CASE WHEN amudf.object_id IS NULL THEN CAST(smudf.null_on_null_input AS BIT) ELSE amudf.null_on_null_input END AS BIT) AS [ReturnsNullOnNullInput], CASE WHEN amudf.object_id IS NULL THEN CASE ISNULL(smudf.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END ELSE CASE ISNULL(amudf.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END END AS [ExecutionContext], CASE WHEN amudf.object_id IS NULL THEN ISNULL(USER_NAME(smudf.execute_as_principal_id), N'') ELSE USER_NAME(amudf.execute_as_principal_id) END AS [ExecutionContextPrincipal], CAST(OBJECTPROPERTYEX(udf.object_id, N'IsDeterministic') AS BIT) AS [IsDeterministic], ( CASE WHEN 'FN' = udf.type THEN 1 WHEN 'FS' = udf.type THEN 1 WHEN 'IF' = udf.type THEN 3 WHEN 'TF' = udf.type THEN 2 WHEN 'FT' = udf.type THEN 2 ELSE 0 END ) AS [FunctionType], CASE WHEN udf.type IN ( 'FN', 'IF', 'TF' ) THEN 1 WHEN udf.type IN ( 'FS', 'FT' ) THEN 2 ELSE 1 END AS [ImplementationType], CAST(ISNULL(OBJECTPROPERTYEX(udf.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], ret_param.name AS [TableVariableName], ISNULL(smudf.definition, ssmudf.definition) AS [Definition] FROM sys.all_objects AS udf LEFT OUTER JOIN sys.database_principals AS sudf ON sudf.principal_id = ISNULL(udf.principal_id, ( OBJECTPROPERTY(udf.object_id, 'OwnerId') )) LEFT OUTER JOIN sys.all_parameters AS ret_param ON ret_param.object_id = udf.object_id AND ret_param.is_output = 1 LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = ret_param.user_type_id LEFT OUTER JOIN sys.schemas AS s1ret_param ON s1ret_param.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = ret_param.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = ret_param.system_type_id ) AND ( baset.user_type_id = ret_param.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscret_param ON xscret_param.xml_collection_id = ret_param.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2ret_param ON s2ret_param.schema_id = xscret_param.schema_id LEFT OUTER JOIN sys.sql_modules AS smudf ON smudf.object_id = udf.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmudf ON ssmudf.object_id = udf.object_id LEFT OUTER JOIN sys.assembly_modules AS amudf ON amudf.object_id = udf.object_id LEFT OUTER JOIN sys.assemblies AS asmbludf ON asmbludf.assembly_id = amudf.assembly_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT udf.name AS [Name], udf.object_id AS [ID], udf.create_date AS [CreateDate], udf.modify_date AS [DateLastModified], ISNULL(sudf.name, N'') AS [Owner], CAST(CASE WHEN udf.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(udf.schema_id) AS [Schema], udf.schema_id AS [SchemaID], CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = udf.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], usrt.name AS [DataType], s1ret_param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND ret_param.max_length <> -1 THEN ret_param.max_length / 2 ELSE ret_param.max_length END AS INT) AS [Length], CAST(ret_param.precision AS INT) AS [NumericPrecision], CAST(ret_param.scale AS INT) AS [NumericScale], ISNULL(xscret_param.name, N'') AS [XmlSchemaNamespace], ISNULL(s2ret_param.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE ret_param.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], CAST(ISNULL(OBJECTPROPERTYEX(udf.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(udf.object_id, N'IsSchemaBound'), 0) AS BIT) AS [IsSchemaBound], CAST(CASE WHEN ISNULL(smudf.definition, ssmudf.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CASE WHEN amudf.object_id IS NULL THEN N'' ELSE asmbludf.name END AS [AssemblyName], CASE WHEN amudf.object_id IS NULL THEN N'' ELSE amudf.assembly_class END AS [ClassName], CASE WHEN amudf.object_id IS NULL THEN N'' ELSE amudf.assembly_method END AS [MethodName], CAST(CASE WHEN amudf.object_id IS NULL THEN CAST(smudf.null_on_null_input AS BIT) ELSE amudf.null_on_null_input END AS BIT) AS [ReturnsNullOnNullInput], CASE WHEN amudf.object_id IS NULL THEN CASE ISNULL(smudf.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END ELSE CASE ISNULL(amudf.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END END AS [ExecutionContext], CASE WHEN amudf.object_id IS NULL THEN ISNULL(USER_NAME(smudf.execute_as_principal_id), N'') ELSE USER_NAME(amudf.execute_as_principal_id) END AS [ExecutionContextPrincipal], CAST(OBJECTPROPERTYEX(udf.object_id, N'IsDeterministic') AS BIT) AS [IsDeterministic], ( CASE WHEN 'FN' = udf.type THEN 1 WHEN 'FS' = udf.type THEN 1 WHEN 'IF' = udf.type THEN 3 WHEN 'TF' = udf.type THEN 2 WHEN 'FT' = udf.type THEN 2 ELSE 0 END ) AS [FunctionType], CASE WHEN udf.type IN ( 'FN', 'IF', 'TF' ) THEN 1 WHEN udf.type IN ( 'FS', 'FT' ) THEN 2 ELSE 1 END AS [ImplementationType], CAST(ISNULL(OBJECTPROPERTYEX(udf.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], ret_param.name AS [TableVariableName], ISNULL(smudf.definition, ssmudf.definition) AS [Definition] FROM sys.all_objects AS udf LEFT OUTER JOIN sys.database_principals AS sudf ON sudf.principal_id = ISNULL(udf.principal_id, ( OBJECTPROPERTY(udf.object_id, 'OwnerId') )) LEFT OUTER JOIN sys.all_parameters AS ret_param ON ret_param.object_id = udf.object_id AND ret_param.is_output = 1 LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = ret_param.user_type_id LEFT OUTER JOIN sys.schemas AS s1ret_param ON s1ret_param.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = ret_param.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = ret_param.system_type_id ) AND ( baset.user_type_id = ret_param.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscret_param ON xscret_param.xml_collection_id = ret_param.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2ret_param ON s2ret_param.schema_id = xscret_param.schema_id LEFT OUTER JOIN sys.sql_modules AS smudf ON smudf.object_id = udf.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmudf ON ssmudf.object_id = udf.object_id LEFT OUTER JOIN sys.assembly_modules AS amudf ON amudf.object_id = udf.object_id LEFT OUTER JOIN sys.assemblies AS asmbludf ON asmbludf.assembly_id = amudf.assembly_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT udf.name AS [Name], udf.object_id AS [ID], udf.create_date AS [CreateDate], udf.modify_date AS [DateLastModified], ISNULL(sudf.name, N'') AS [Owner], CAST(CASE WHEN udf.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(udf.schema_id) AS [Schema], udf.schema_id AS [SchemaID], CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = udf.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], usrt.name AS [DataType], s1ret_param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND ret_param.max_length <> -1 THEN ret_param.max_length / 2 ELSE ret_param.max_length END AS INT) AS [Length], CAST(ret_param.precision AS INT) AS [NumericPrecision], CAST(ret_param.scale AS INT) AS [NumericScale], ISNULL(xscret_param.name, N'') AS [XmlSchemaNamespace], ISNULL(s2ret_param.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE ret_param.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], CAST(ISNULL(OBJECTPROPERTYEX(udf.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(udf.object_id, N'IsSchemaBound'), 0) AS BIT) AS [IsSchemaBound], CAST(CASE WHEN ISNULL(smudf.definition, ssmudf.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CASE WHEN amudf.object_id IS NULL THEN N'' ELSE asmbludf.name END AS [AssemblyName], CASE WHEN amudf.object_id IS NULL THEN N'' ELSE amudf.assembly_class END AS [ClassName], CASE WHEN amudf.object_id IS NULL THEN N'' ELSE amudf.assembly_method END AS [MethodName], CAST(CASE WHEN amudf.object_id IS NULL THEN CAST(smudf.null_on_null_input AS BIT) ELSE amudf.null_on_null_input END AS BIT) AS [ReturnsNullOnNullInput], CASE WHEN amudf.object_id IS NULL THEN CASE ISNULL(smudf.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END ELSE CASE ISNULL(amudf.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END END AS [ExecutionContext], CASE WHEN amudf.object_id IS NULL THEN ISNULL(USER_NAME(smudf.execute_as_principal_id), N'') ELSE USER_NAME(amudf.execute_as_principal_id) END AS [ExecutionContextPrincipal], CAST(OBJECTPROPERTYEX(udf.object_id, N'IsDeterministic') AS BIT) AS [IsDeterministic], ( CASE WHEN 'FN' = udf.type THEN 1 WHEN 'FS' = udf.type THEN 1 WHEN 'IF' = udf.type THEN 3 WHEN 'TF' = udf.type THEN 2 WHEN 'FT' = udf.type THEN 2 ELSE 0 END ) AS [FunctionType], CASE WHEN udf.type IN ( 'FN', 'IF', 'TF' ) THEN 1 WHEN udf.type IN ( 'FS', 'FT' ) THEN 2 ELSE 1 END AS [ImplementationType], CAST(ISNULL(OBJECTPROPERTYEX(udf.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], ret_param.name AS [TableVariableName], ISNULL(smudf.definition, ssmudf.definition) AS [Definition] FROM sys.all_objects AS udf LEFT OUTER JOIN sys.database_principals AS sudf ON sudf.principal_id = ISNULL(udf.principal_id, ( OBJECTPROPERTY(udf.object_id, 'OwnerId') )) LEFT OUTER JOIN sys.all_parameters AS ret_param ON ret_param.object_id = udf.object_id AND ret_param.is_output = 1 LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = ret_param.user_type_id LEFT OUTER JOIN sys.schemas AS s1ret_param ON s1ret_param.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = ret_param.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = ret_param.system_type_id ) AND ( baset.user_type_id = ret_param.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscret_param ON xscret_param.xml_collection_id = ret_param.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2ret_param ON s2ret_param.schema_id = xscret_param.schema_id LEFT OUTER JOIN sys.sql_modules AS smudf ON smudf.object_id = udf.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmudf ON ssmudf.object_id = udf.object_id LEFT OUTER JOIN sys.assembly_modules AS amudf ON amudf.object_id = udf.object_id LEFT OUTER JOIN sys.assemblies AS asmbludf ON asmbludf.assembly_id = amudf.assembly_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT udf.name AS [Name], udf.object_id AS [ID], udf.create_date AS [CreateDate], udf.modify_date AS [DateLastModified], ISNULL(sudf.name, N'') AS [Owner], CAST(CASE WHEN udf.principal_id IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsSchemaOwned], SCHEMA_NAME(udf.schema_id) AS [Schema], udf.schema_id AS [SchemaID], CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = udf.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], usrt.name AS [DataType], s1ret_param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND ret_param.max_length <> -1 THEN ret_param.max_length / 2 ELSE ret_param.max_length END AS INT) AS [Length], CAST(ret_param.precision AS INT) AS [NumericPrecision], CAST(ret_param.scale AS INT) AS [NumericScale], ISNULL(xscret_param.name, N'') AS [XmlSchemaNamespace], ISNULL(s2ret_param.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE ret_param.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CAST(ISNULL(OBJECTPROPERTYEX(udf.object_id, N'ExecIsAnsiNullsOn'), 0) AS BIT) AS [AnsiNullsStatus], CAST(ISNULL(OBJECTPROPERTYEX(udf.object_id, N'IsSchemaBound'), 0) AS BIT) AS [IsSchemaBound], CAST(CASE WHEN ISNULL(smudf.definition, ssmudf.definition) IS NULL THEN 1 ELSE 0 END AS BIT) AS [IsEncrypted], CASE WHEN amudf.object_id IS NULL THEN N'' ELSE asmbludf.name END AS [AssemblyName], CASE WHEN amudf.object_id IS NULL THEN N'' ELSE amudf.assembly_class END AS [ClassName], CASE WHEN amudf.object_id IS NULL THEN N'' ELSE amudf.assembly_method END AS [MethodName], CAST(CASE WHEN amudf.object_id IS NULL THEN CAST(smudf.null_on_null_input AS BIT) ELSE amudf.null_on_null_input END AS BIT) AS [ReturnsNullOnNullInput], CASE WHEN amudf.object_id IS NULL THEN CASE ISNULL(smudf.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END ELSE CASE ISNULL(amudf.execute_as_principal_id, -1) WHEN -1 THEN 1 WHEN -2 THEN 2 ELSE 3 END END AS [ExecutionContext], CASE WHEN amudf.object_id IS NULL THEN ISNULL(USER_NAME(smudf.execute_as_principal_id), N'') ELSE USER_NAME(amudf.execute_as_principal_id) END AS [ExecutionContextPrincipal], CAST(OBJECTPROPERTYEX(udf.object_id, N'IsDeterministic') AS BIT) AS [IsDeterministic], ( CASE WHEN 'FN' = udf.type THEN 1 WHEN 'FS' = udf.type THEN 1 WHEN 'IF' = udf.type THEN 3 WHEN 'TF' = udf.type THEN 2 WHEN 'FT' = udf.type THEN 2 ELSE 0 END ) AS [FunctionType], CASE WHEN udf.type IN ( 'FN', 'IF', 'TF' ) THEN 1 WHEN udf.type IN ( 'FS', 'FT' ) THEN 2 ELSE 1 END AS [ImplementationType], CAST(ISNULL(OBJECTPROPERTYEX(udf.object_id, N'ExecIsQuotedIdentOn'), 0) AS BIT) AS [QuotedIdentifierStatus], ret_param.name AS [TableVariableName], ISNULL(smudf.definition, ssmudf.definition) AS [Definition] FROM sys.all_objects AS udf LEFT OUTER JOIN sys.database_principals AS sudf ON sudf.principal_id = ISNULL(udf.principal_id, ( OBJECTPROPERTY(udf.object_id, 'OwnerId') )) LEFT OUTER JOIN sys.all_parameters AS ret_param ON ret_param.object_id = udf.object_id AND ret_param.is_output = 1 LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = ret_param.user_type_id LEFT OUTER JOIN sys.schemas AS s1ret_param ON s1ret_param.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = ret_param.system_type_id AND baset.user_type_id = baset.system_type_id ) LEFT OUTER JOIN sys.xml_schema_collections AS xscret_param ON xscret_param.xml_collection_id = ret_param.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2ret_param ON s2ret_param.schema_id = xscret_param.schema_id LEFT OUTER JOIN sys.sql_modules AS smudf ON smudf.object_id = udf.object_id LEFT OUTER JOIN sys.system_sql_modules AS ssmudf ON ssmudf.object_id = udf.object_id LEFT OUTER JOIN sys.assembly_modules AS amudf ON amudf.object_id = udf.object_id LEFT OUTER JOIN sys.assemblies AS asmbludf ON asmbludf.assembly_id = amudf.assembly_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(udf.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(udf.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT udf.name AS [Name], udf.id AS [ID], udf.crdate AS [CreateDate], sudf.name AS [Schema], sudf.uid AS [SchemaID], sudf.name AS [Owner], CAST(CASE WHEN ( OBJECTPROPERTY(udf.id, N'IsMSShipped') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(udf.id, N'IsSystemTable') THEN 1 ELSE 0 END AS BIT) AS [IsSystemObject], usrt.name AS [DataType], s1ret_param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'char', N'varchar', N'binary', N'varbinary', N'nchar', N'nvarchar' ) THEN ret_param.prec ELSE ret_param.length END AS INT) AS [Length], CAST(ret_param.xprec AS INT) AS [NumericPrecision], CAST(ret_param.xscale AS INT) AS [NumericScale], CAST(OBJECTPROPERTY(udf.id, N'ExecIsAnsiNullsOn') AS BIT) AS [AnsiNullsStatus], CAST(OBJECTPROPERTY(udf.id, N'IsSchemaBound') AS BIT) AS [IsSchemaBound], CAST((SELECT TOP 1 encrypted FROM dbo.syscomments p WHERE udf.id = p.id AND p.colid = 1 AND p.number < 2) AS BIT) AS [IsEncrypted], CAST(OBJECTPROPERTY(udf.id, N'IsDeterministic') AS BIT) AS [IsDeterministic], ( CASE WHEN 1 = OBJECTPROPERTY(udf.id, N'IsScalarFunction') THEN 1 WHEN 1 = OBJECTPROPERTY(udf.id, N'IsInlineFunction') THEN 3 WHEN 1 = OBJECTPROPERTY(udf.id, N'IsTableFunction') THEN 2 ELSE 0 END ) AS [FunctionType], 1 AS [ImplementationType], CAST(ISNULL(OBJECTPROPERTYEX(udf.id, N'IsQuotedIdentOn'), 1) AS BIT) AS [QuotedIdentifierStatus], ret_param.name AS [TableVariableName], c.text AS [Definition] FROM dbo.sysobjects AS udf INNER JOIN sysusers AS sudf ON sudf.uid = udf.uid LEFT OUTER JOIN syscolumns AS ret_param ON ret_param.id = udf.id AND ret_param.number = 0 AND ret_param.name = '' LEFT OUTER JOIN systypes AS usrt ON usrt.xusertype = ret_param.xusertype LEFT OUTER JOIN sysusers AS s1ret_param ON s1ret_param.uid = usrt.uid LEFT OUTER JOIN systypes AS baset ON baset.xusertype = ret_param.xtype AND baset.xusertype = baset.xtype LEFT OUTER JOIN dbo.syscomments c ON c.id = udf.id AND CASE WHEN c.number > 1 THEN c.number ELSE 0 END = 0 WHERE ( udf.xtype IN ( 'TF', 'FN', 'IF' ) AND udf.name NOT LIKE N'#%%' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-UserDefinedFunctionCheckQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], ~cstr.is_not_trusted AS [IsChecked], ~cstr.is_disabled AS [IsEnabled], cstr.is_not_for_replication AS [NotForReplication], cstr.definition AS [Definition] FROM sys.all_objects AS udf INNER JOIN sys.check_constraints AS cstr ON cstr.parent_object_id = udf.object_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = udf.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], ~cstr.is_not_trusted AS [IsChecked], ~cstr.is_disabled AS [IsEnabled], cstr.is_not_for_replication AS [NotForReplication], CAST(CASE WHEN filetableobj.object_id IS NULL THEN 0 ELSE 1 END AS BIT) AS [IsFileTableDefined], cstr.definition AS [Definition] FROM sys.all_objects AS udf INNER JOIN sys.check_constraints AS cstr ON cstr.parent_object_id = udf.object_id LEFT OUTER JOIN sys.filetable_system_defined_objects AS filetableobj ON filetableobj.object_id = cstr.object_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], ~cstr.is_not_trusted AS [IsChecked], ~cstr.is_disabled AS [IsEnabled], cstr.is_not_for_replication AS [NotForReplication], cstr.definition AS [Definition] FROM sys.all_objects AS udf INNER JOIN sys.check_constraints AS cstr ON cstr.parent_object_id = udf.object_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], ~cstr.is_not_trusted AS [IsChecked], ~cstr.is_disabled AS [IsEnabled], cstr.is_not_for_replication AS [NotForReplication], cstr.definition AS [Definition] FROM sys.all_objects AS udf INNER JOIN sys.check_constraints AS cstr ON cstr.parent_object_id = udf.object_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], ~cstr.is_not_trusted AS [IsChecked], ~cstr.is_disabled AS [IsEnabled], cstr.is_not_for_replication AS [NotForReplication], cstr.definition AS [Definition] FROM sys.all_objects AS udf INNER JOIN sys.check_constraints AS cstr ON cstr.parent_object_id = udf.object_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(udf.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(udf.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT udf.id AS [FunctionID], sudf.uid AS [SchemaID], cstr.name AS [Name], cstr.id AS [ID], cstr.crdate AS [CreateDate], CAST(cstr.status & 4 AS BIT) AS [IsSystemNamed], CAST(1 - ISNULL(OBJECTPROPERTY(cstr.id, N'CnstIsNotTrusted'), 0) AS BIT) AS [IsChecked], CAST(1 - ISNULL(OBJECTPROPERTY(cstr.id, N'CnstIsDisabled'), 0) AS BIT) AS [IsEnabled], CAST(ISNULL(OBJECTPROPERTY(cstr.id, N'CnstIsNotRepl'), 0) AS BIT) AS [NotForReplication], c.text AS [Definition] FROM dbo.sysobjects AS udf INNER JOIN sysusers AS sudf ON sudf.uid = udf.uid INNER JOIN dbo.sysobjects AS cstr ON ( cstr.type = 'C' ) AND ( cstr.parent_obj = udf.id ) LEFT OUTER JOIN dbo.syscomments c ON c.id = cstr.id AND CASE WHEN c.number > 1 THEN c.number ELSE 0 END = 0 WHERE ( udf.xtype IN ( 'TF', 'FN', 'IF' ) AND udf.name NOT LIKE N'#%%' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-UserDefinedFunctionColumnQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], clmns.name AS [Name], clmns.column_id AS [ID], clmns.is_nullable AS [Nullable], clmns.is_computed AS [Computed], CAST(ISNULL(cik.index_column_id, 0) AS BIT) AS [InPrimaryKey], clmns.is_ansi_padded AS [AnsiPaddingStatus], CAST(clmns.is_rowguidcol AS BIT) AS [RowGuidCol], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsDeterministic'), 0) AS BIT) AS [IsDeterministic], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsPrecise'), 0) AS BIT) AS [IsPrecise], CAST(ISNULL(cc.is_persisted, 0) AS BIT) AS [IsPersisted], ISNULL(clmns.collation_name, N'') AS [Collation], CAST(ISNULL((SELECT TOP 1 1 FROM sys.foreign_key_columns AS colfk WHERE colfk.parent_column_id = clmns.column_id AND colfk.parent_object_id = clmns.object_id), 0) AS BIT) AS [IsForeignKey], clmns.is_identity AS [Identity], CAST(ISNULL(ic.seed_value, 0) AS BIGINT) AS [IdentitySeed], CAST(ISNULL(ic.increment_value, 0) AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE SCHEMA_NAME(d.schema_id) END ) AS [DefaultSchema], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE SCHEMA_NAME(r.schema_id) END ) AS [RuleSchema], ISNULL(ic.is_not_for_replication, 0) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], CAST(clmns.is_filestream AS BIT) AS [IsFileStream], CAST(clmns.is_sparse AS BIT) AS [IsSparse], CAST(clmns.is_column_set AS BIT) AS [IsColumnSet], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND clmns.max_length <> -1 THEN clmns.max_length / 2 ELSE clmns.max_length END AS INT) AS [Length], CAST(clmns.precision AS INT) AS [NumericPrecision], CAST(clmns.scale AS INT) AS [NumericScale], ISNULL(xscclmns.name, N'') AS [XmlSchemaNamespace], ISNULL(s2clmns.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE clmns.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], ISNULL(cc.definition, N'') AS [ComputedText] FROM sys.all_objects AS udf INNER JOIN sys.all_columns AS clmns ON clmns.object_id = udf.object_id LEFT OUTER JOIN sys.indexes AS ik ON ik.object_id = clmns.object_id AND 1 = ik.is_primary_key LEFT OUTER JOIN sys.index_columns AS cik ON cik.index_id = ik.index_id AND cik.column_id = clmns.column_id AND cik.object_id = clmns.object_id AND 0 = cik.is_included_column LEFT OUTER JOIN sys.computed_columns AS cc ON cc.object_id = clmns.object_id AND cc.column_id = clmns.column_id LEFT OUTER JOIN sys.identity_columns AS ic ON ic.object_id = clmns.object_id AND ic.column_id = clmns.column_id LEFT OUTER JOIN sys.objects AS d ON d.object_id = clmns.default_object_id LEFT OUTER JOIN sys.objects AS r ON r.object_id = clmns.rule_object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = clmns.user_type_id LEFT OUTER JOIN sys.schemas AS s1clmns ON s1clmns.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = clmns.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = clmns.system_type_id ) AND ( baset.user_type_id = clmns.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscclmns ON xscclmns.xml_collection_id = clmns.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2clmns ON s2clmns.schema_id = xscclmns.schema_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = udf.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], clmns.name AS [Name], clmns.column_id AS [ID], clmns.is_nullable AS [Nullable], clmns.is_computed AS [Computed], CAST(ISNULL(cik.index_column_id, 0) AS BIT) AS [InPrimaryKey], clmns.is_ansi_padded AS [AnsiPaddingStatus], CAST(clmns.is_rowguidcol AS BIT) AS [RowGuidCol], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsDeterministic'), 0) AS BIT) AS [IsDeterministic], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsPrecise'), 0) AS BIT) AS [IsPrecise], CAST(ISNULL(cc.is_persisted, 0) AS BIT) AS [IsPersisted], ISNULL(clmns.collation_name, N'') AS [Collation], CAST(ISNULL((SELECT TOP 1 1 FROM sys.foreign_key_columns AS colfk WHERE colfk.parent_column_id = clmns.column_id AND colfk.parent_object_id = clmns.object_id), 0) AS BIT) AS [IsForeignKey], clmns.is_identity AS [Identity], CAST(ISNULL(ic.seed_value, 0) AS BIGINT) AS [IdentitySeed], CAST(ISNULL(ic.increment_value, 0) AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE SCHEMA_NAME(d.schema_id) END ) AS [DefaultSchema], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE SCHEMA_NAME(r.schema_id) END ) AS [RuleSchema], ISNULL(ic.is_not_for_replication, 0) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'StatisticalSemantics') AS INT) AS [StatisticalSemantics], CAST(clmns.is_filestream AS BIT) AS [IsFileStream], CAST(clmns.is_sparse AS BIT) AS [IsSparse], CAST(clmns.is_column_set AS BIT) AS [IsColumnSet], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND clmns.max_length <> -1 THEN clmns.max_length / 2 ELSE clmns.max_length END AS INT) AS [Length], CAST(clmns.precision AS INT) AS [NumericPrecision], CAST(clmns.scale AS INT) AS [NumericScale], ISNULL(xscclmns.name, N'') AS [XmlSchemaNamespace], ISNULL(s2clmns.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE clmns.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], ISNULL(cc.definition, N'') AS [ComputedText] FROM sys.all_objects AS udf INNER JOIN sys.all_columns AS clmns ON clmns.object_id = udf.object_id LEFT OUTER JOIN sys.indexes AS ik ON ik.object_id = clmns.object_id AND 1 = ik.is_primary_key LEFT OUTER JOIN sys.index_columns AS cik ON cik.index_id = ik.index_id AND cik.column_id = clmns.column_id AND cik.object_id = clmns.object_id AND 0 = cik.is_included_column LEFT OUTER JOIN sys.computed_columns AS cc ON cc.object_id = clmns.object_id AND cc.column_id = clmns.column_id LEFT OUTER JOIN sys.identity_columns AS ic ON ic.object_id = clmns.object_id AND ic.column_id = clmns.column_id LEFT OUTER JOIN sys.objects AS d ON d.object_id = clmns.default_object_id LEFT OUTER JOIN sys.objects AS r ON r.object_id = clmns.rule_object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = clmns.user_type_id LEFT OUTER JOIN sys.schemas AS s1clmns ON s1clmns.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = clmns.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = clmns.system_type_id ) AND ( baset.user_type_id = clmns.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscclmns ON xscclmns.xml_collection_id = clmns.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2clmns ON s2clmns.schema_id = xscclmns.schema_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], clmns.name AS [Name], clmns.column_id AS [ID], clmns.is_nullable AS [Nullable], clmns.is_computed AS [Computed], CAST(ISNULL(cik.index_column_id, 0) AS BIT) AS [InPrimaryKey], clmns.is_ansi_padded AS [AnsiPaddingStatus], CAST(clmns.is_rowguidcol AS BIT) AS [RowGuidCol], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsDeterministic'), 0) AS BIT) AS [IsDeterministic], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsPrecise'), 0) AS BIT) AS [IsPrecise], CAST(ISNULL(cc.is_persisted, 0) AS BIT) AS [IsPersisted], ISNULL(clmns.collation_name, N'') AS [Collation], CAST(ISNULL((SELECT TOP 1 1 FROM sys.foreign_key_columns AS colfk WHERE colfk.parent_column_id = clmns.column_id AND colfk.parent_object_id = clmns.object_id), 0) AS BIT) AS [IsForeignKey], clmns.is_identity AS [Identity], CAST(ISNULL(ic.seed_value, 0) AS BIGINT) AS [IdentitySeed], CAST(ISNULL(ic.increment_value, 0) AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE SCHEMA_NAME(d.schema_id) END ) AS [DefaultSchema], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE SCHEMA_NAME(r.schema_id) END ) AS [RuleSchema], ISNULL(ic.is_not_for_replication, 0) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], CAST(clmns.is_filestream AS BIT) AS [IsFileStream], CAST(clmns.is_sparse AS BIT) AS [IsSparse], CAST(clmns.is_column_set AS BIT) AS [IsColumnSet], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND clmns.max_length <> -1 THEN clmns.max_length / 2 ELSE clmns.max_length END AS INT) AS [Length], CAST(clmns.precision AS INT) AS [NumericPrecision], CAST(clmns.scale AS INT) AS [NumericScale], ISNULL(xscclmns.name, N'') AS [XmlSchemaNamespace], ISNULL(s2clmns.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE clmns.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], ISNULL(cc.definition, N'') AS [ComputedText] FROM sys.all_objects AS udf INNER JOIN sys.all_columns AS clmns ON clmns.object_id = udf.object_id LEFT OUTER JOIN sys.indexes AS ik ON ik.object_id = clmns.object_id AND 1 = ik.is_primary_key LEFT OUTER JOIN sys.index_columns AS cik ON cik.index_id = ik.index_id AND cik.column_id = clmns.column_id AND cik.object_id = clmns.object_id AND 0 = cik.is_included_column LEFT OUTER JOIN sys.computed_columns AS cc ON cc.object_id = clmns.object_id AND cc.column_id = clmns.column_id LEFT OUTER JOIN sys.identity_columns AS ic ON ic.object_id = clmns.object_id AND ic.column_id = clmns.column_id LEFT OUTER JOIN sys.objects AS d ON d.object_id = clmns.default_object_id LEFT OUTER JOIN sys.objects AS r ON r.object_id = clmns.rule_object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = clmns.user_type_id LEFT OUTER JOIN sys.schemas AS s1clmns ON s1clmns.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = clmns.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = clmns.system_type_id ) AND ( baset.user_type_id = clmns.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscclmns ON xscclmns.xml_collection_id = clmns.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2clmns ON s2clmns.schema_id = xscclmns.schema_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], clmns.name AS [Name], clmns.column_id AS [ID], clmns.is_nullable AS [Nullable], clmns.is_computed AS [Computed], CAST(ISNULL(cik.index_column_id, 0) AS BIT) AS [InPrimaryKey], clmns.is_ansi_padded AS [AnsiPaddingStatus], CAST(clmns.is_rowguidcol AS BIT) AS [RowGuidCol], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsDeterministic'), 0) AS BIT) AS [IsDeterministic], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsPrecise'), 0) AS BIT) AS [IsPrecise], CAST(ISNULL(cc.is_persisted, 0) AS BIT) AS [IsPersisted], ISNULL(clmns.collation_name, N'') AS [Collation], CAST(ISNULL((SELECT TOP 1 1 FROM sys.foreign_key_columns AS colfk WHERE colfk.parent_column_id = clmns.column_id AND colfk.parent_object_id = clmns.object_id), 0) AS BIT) AS [IsForeignKey], clmns.is_identity AS [Identity], CAST(ISNULL(ic.seed_value, 0) AS BIGINT) AS [IdentitySeed], CAST(ISNULL(ic.increment_value, 0) AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE SCHEMA_NAME(d.schema_id) END ) AS [DefaultSchema], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE SCHEMA_NAME(r.schema_id) END ) AS [RuleSchema], ISNULL(ic.is_not_for_replication, 0) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], CAST(clmns.is_filestream AS BIT) AS [IsFileStream], CAST(clmns.is_sparse AS BIT) AS [IsSparse], CAST(clmns.is_column_set AS BIT) AS [IsColumnSet], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND clmns.max_length <> -1 THEN clmns.max_length / 2 ELSE clmns.max_length END AS INT) AS [Length], CAST(clmns.precision AS INT) AS [NumericPrecision], CAST(clmns.scale AS INT) AS [NumericScale], ISNULL(xscclmns.name, N'') AS [XmlSchemaNamespace], ISNULL(s2clmns.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE clmns.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], ISNULL(cc.definition, N'') AS [ComputedText] FROM sys.all_objects AS udf INNER JOIN sys.all_columns AS clmns ON clmns.object_id = udf.object_id LEFT OUTER JOIN sys.indexes AS ik ON ik.object_id = clmns.object_id AND 1 = ik.is_primary_key LEFT OUTER JOIN sys.index_columns AS cik ON cik.index_id = ik.index_id AND cik.column_id = clmns.column_id AND cik.object_id = clmns.object_id AND 0 = cik.is_included_column LEFT OUTER JOIN sys.computed_columns AS cc ON cc.object_id = clmns.object_id AND cc.column_id = clmns.column_id LEFT OUTER JOIN sys.identity_columns AS ic ON ic.object_id = clmns.object_id AND ic.column_id = clmns.column_id LEFT OUTER JOIN sys.objects AS d ON d.object_id = clmns.default_object_id LEFT OUTER JOIN sys.objects AS r ON r.object_id = clmns.rule_object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = clmns.user_type_id LEFT OUTER JOIN sys.schemas AS s1clmns ON s1clmns.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = clmns.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = clmns.system_type_id ) AND ( baset.user_type_id = clmns.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscclmns ON xscclmns.xml_collection_id = clmns.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2clmns ON s2clmns.schema_id = xscclmns.schema_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], clmns.name AS [Name], clmns.column_id AS [ID], clmns.is_nullable AS [Nullable], clmns.is_computed AS [Computed], CAST(ISNULL(cik.index_column_id, 0) AS BIT) AS [InPrimaryKey], clmns.is_ansi_padded AS [AnsiPaddingStatus], CAST(clmns.is_rowguidcol AS BIT) AS [RowGuidCol], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsDeterministic'), 0) AS BIT) AS [IsDeterministic], CAST(ISNULL(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsPrecise'), 0) AS BIT) AS [IsPrecise], CAST(ISNULL(cc.is_persisted, 0) AS BIT) AS [IsPersisted], ISNULL(clmns.collation_name, N'') AS [Collation], CAST(ISNULL((SELECT TOP 1 1 FROM sys.foreign_key_columns AS colfk WHERE colfk.parent_column_id = clmns.column_id AND colfk.parent_object_id = clmns.object_id), 0) AS BIT) AS [IsForeignKey], clmns.is_identity AS [Identity], CAST(ISNULL(ic.seed_value, 0) AS BIGINT) AS [IdentitySeed], CAST(ISNULL(ic.increment_value, 0) AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.default_object_id = 0 THEN N'' WHEN d.parent_object_id > 0 THEN N'' ELSE SCHEMA_NAME(d.schema_id) END ) AS [DefaultSchema], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.rule_object_id = 0 THEN N'' ELSE SCHEMA_NAME(r.schema_id) END ) AS [RuleSchema], ISNULL(ic.is_not_for_replication, 0) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.object_id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND clmns.max_length <> -1 THEN clmns.max_length / 2 ELSE clmns.max_length END AS INT) AS [Length], CAST(clmns.precision AS INT) AS [NumericPrecision], CAST(clmns.scale AS INT) AS [NumericScale], ISNULL(xscclmns.name, N'') AS [XmlSchemaNamespace], ISNULL(s2clmns.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE clmns.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], ISNULL(cc.definition, N'') AS [ComputedText] FROM sys.all_objects AS udf INNER JOIN sys.all_columns AS clmns ON clmns.object_id = udf.object_id LEFT OUTER JOIN sys.indexes AS ik ON ik.object_id = clmns.object_id AND 1 = ik.is_primary_key LEFT OUTER JOIN sys.index_columns AS cik ON cik.index_id = ik.index_id AND cik.column_id = clmns.column_id AND cik.object_id = clmns.object_id AND 0 = cik.is_included_column LEFT OUTER JOIN sys.computed_columns AS cc ON cc.object_id = clmns.object_id AND cc.column_id = clmns.column_id LEFT OUTER JOIN sys.identity_columns AS ic ON ic.object_id = clmns.object_id AND ic.column_id = clmns.column_id LEFT OUTER JOIN sys.objects AS d ON d.object_id = clmns.default_object_id LEFT OUTER JOIN sys.objects AS r ON r.object_id = clmns.rule_object_id LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = clmns.user_type_id LEFT OUTER JOIN sys.schemas AS s1clmns ON s1clmns.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = clmns.system_type_id AND baset.user_type_id = baset.system_type_id ) LEFT OUTER JOIN sys.xml_schema_collections AS xscclmns ON xscclmns.xml_collection_id = clmns.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2clmns ON s2clmns.schema_id = xscclmns.schema_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(udf.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(udf.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT udf.id AS [FunctionID], sudf.uid AS [SchemaID], clmns.name AS [Name], CAST(clmns.colid AS INT) AS [ID], CAST(clmns.isnullable AS BIT) AS [Nullable], CAST(clmns.iscomputed AS BIT) AS [Computed], CAST(ISNULL(cik.colid, 0) AS BIT) AS [InPrimaryKey], CAST(ISNULL(COLUMNPROPERTY(clmns.id, clmns.name, N'UsesAnsiTrim'), 0) AS BIT) AS [AnsiPaddingStatus], CAST(clmns.colstat & 2 AS BIT) AS [RowGuidCol], CAST(clmns.colstat & 8 AS BIT) AS [NotForReplication], CAST(COLUMNPROPERTY(clmns.id, clmns.name, N'IsFulltextIndexed') AS BIT) AS [IsFullTextIndexed], CAST(COLUMNPROPERTY(clmns.id, clmns.name, N'IsIdentity') AS BIT) AS [Identity], CAST(ISNULL((SELECT TOP 1 1 FROM dbo.sysforeignkeys AS colfk WHERE colfk.fkey = clmns.colid AND colfk.fkeyid = clmns.id), 0) AS BIT) AS [IsForeignKey], ISNULL(clmns.collation, N'') AS [Collation], CAST(CASE COLUMNPROPERTY(clmns.id, clmns.name, N'IsIdentity') WHEN 1 THEN IDENT_SEED(QUOTENAME(sudf.name) + '.' + QUOTENAME(udf.name)) ELSE 0 END AS BIGINT) AS [IdentitySeed], CAST(CASE COLUMNPROPERTY(clmns.id, clmns.name, N'IsIdentity') WHEN 1 THEN IDENT_INCR(QUOTENAME(sudf.name) + '.' + QUOTENAME(udf.name)) ELSE 0 END AS BIGINT) AS [IdentityIncrement], ( CASE WHEN clmns.cdefault = 0 THEN N'' ELSE d.name END ) AS [Default], ( CASE WHEN clmns.cdefault = 0 THEN N'' ELSE USER_NAME(d.uid) END ) AS [DefaultSchema], ( CASE WHEN clmns.domain = 0 THEN N'' ELSE r.name END ) AS [Rule], ( CASE WHEN clmns.domain = 0 THEN N'' ELSE USER_NAME(r.uid) END ) AS [RuleSchema], usrt.name AS [DataType], s1clmns.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'char', N'varchar', N'binary', N'varbinary', N'nchar', N'nvarchar' ) THEN clmns.prec ELSE clmns.length END AS INT) AS [Length], CAST(clmns.xprec AS INT) AS [NumericPrecision], CAST(clmns.xscale AS INT) AS [NumericScale], comt.text AS [ComputedText] FROM dbo.sysobjects AS udf INNER JOIN sysusers AS sudf ON sudf.uid = udf.uid INNER JOIN dbo.syscolumns AS clmns ON clmns.id = udf.id LEFT OUTER JOIN dbo.sysindexes AS ik ON ik.id = clmns.id AND 0 != ik.status & 0x0800 LEFT OUTER JOIN dbo.sysindexkeys AS cik ON cik.indid = ik.indid AND cik.colid = clmns.colid AND cik.id = clmns.id LEFT OUTER JOIN dbo.sysobjects AS d ON d.id = clmns.cdefault AND 0 = d.category & 0x0800 LEFT OUTER JOIN dbo.sysobjects AS r ON r.id = clmns.domain LEFT OUTER JOIN systypes AS usrt ON usrt.xusertype = clmns.xusertype LEFT OUTER JOIN sysusers AS s1clmns ON s1clmns.uid = usrt.uid LEFT OUTER JOIN systypes AS baset ON baset.xusertype = clmns.xtype AND baset.xusertype = baset.xtype LEFT OUTER JOIN dbo.syscomments comt ON comt.number = CAST(clmns.colid AS INT) AND comt.id = clmns.id WHERE ( udf.xtype IN ( 'TF', 'FN', 'IF' ) AND udf.name NOT LIKE N'#%%' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-UserDefinedFunctionDefaultConstraint([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], clmns.column_id AS [ColumnID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], cstr.definition AS [Text] FROM sys.all_objects AS udf INNER JOIN sys.all_columns AS clmns ON clmns.object_id = udf.object_id INNER JOIN sys.default_constraints AS cstr ON cstr.object_id = clmns.default_object_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = udf.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], clmns.column_id AS [ColumnID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], CAST(CASE WHEN filetableobj.object_id IS NULL THEN 0 ELSE 1 END AS BIT) AS [IsFileTableDefined], cstr.definition AS [Text] FROM sys.all_objects AS udf INNER JOIN sys.all_columns AS clmns ON clmns.object_id = udf.object_id INNER JOIN sys.default_constraints AS cstr ON cstr.object_id = clmns.default_object_id LEFT OUTER JOIN sys.filetable_system_defined_objects AS filetableobj ON filetableobj.object_id = cstr.object_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], clmns.column_id AS [ColumnID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], cstr.definition AS [Text] FROM sys.all_objects AS udf INNER JOIN sys.all_columns AS clmns ON clmns.object_id = udf.object_id INNER JOIN sys.default_constraints AS cstr ON cstr.object_id = clmns.default_object_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], clmns.column_id AS [ColumnID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], cstr.definition AS [Text] FROM sys.all_objects AS udf INNER JOIN sys.all_columns AS clmns ON clmns.object_id = udf.object_id INNER JOIN sys.default_constraints AS cstr ON cstr.object_id = clmns.default_object_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], clmns.column_id AS [ColumnID], cstr.name AS [Name], cstr.object_id AS [ID], cstr.create_date AS [CreateDate], cstr.modify_date AS [DateLastModified], CAST(cstr.is_system_named AS BIT) AS [IsSystemNamed], cstr.definition AS [Text] FROM sys.all_objects AS udf INNER JOIN sys.all_columns AS clmns ON clmns.object_id = udf.object_id INNER JOIN sys.default_constraints AS cstr ON cstr.object_id = clmns.default_object_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(udf.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(udf.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT udf.id AS [FunctionID], sudf.uid AS [SchemaID], clmns.colid AS [ColumnID], cstr.name AS [Name], cstr.id AS [ID], cstr.crdate AS [CreateDate], CAST(cstr.status & 4 AS BIT) AS [IsSystemNamed], c.text AS [Text] FROM dbo.sysobjects AS udf INNER JOIN sysusers AS sudf ON sudf.uid = udf.uid INNER JOIN dbo.syscolumns AS clmns ON clmns.id = udf.id INNER JOIN dbo.sysobjects AS cstr ON ( cstr.xtype = 'D' AND cstr.name NOT LIKE N'#%%' AND 0 != CONVERT(BIT, cstr.category & 0x0800) ) AND ( cstr.id = clmns.cdefault ) LEFT OUTER JOIN dbo.syscomments c ON c.id = cstr.id AND CASE WHEN c.number > 1 THEN c.number ELSE 0 END = 0 WHERE ( udf.xtype IN ( 'TF', 'FN', 'IF' ) AND udf.name NOT LIKE N'#%%' AND clmns.number = 0 AND 0 = OBJECTPROPERTY(clmns.id, N'IsScalarFunction') ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-UserDefinedFunctionIndexQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], i.name AS [Name], CAST(i.index_id AS INT) AS [ID], CAST(OBJECTPROPERTY(i.object_id, N'IsMSShipped') AS BIT) AS [IsSystemObject], ISNULL(s.no_recompute, 0) AS [NoAutomaticRecomputation], i.fill_factor AS [FillFactor], CAST(CASE i.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], i.is_primary_key + 2 * i.is_unique_constraint AS [IndexKeyType], i.is_unique AS [IsUnique], i.ignore_dup_key AS [IgnoreDuplicateKeys], ~i.allow_row_locks AS [DisallowRowLocks], ~i.allow_page_locks AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], i.is_disabled AS [IsDisabled], CAST(ISNULL(k.is_system_named, 0) AS BIT) AS [IsSystemNamed], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN i.type = 3 THEN 1 ELSE 0 END AS BIT) AS [IsXmlIndex], CAST(ISNULL(spi.spatial_index_type, 0) AS TINYINT) AS [SpatialIndexType], CAST(ISNULL(si.bounding_box_xmin, 0) AS FLOAT(53)) AS [BoundingBoxXMin], CAST(ISNULL(si.bounding_box_ymin, 0) AS FLOAT(53)) AS [BoundingBoxYMin], CAST(ISNULL(si.bounding_box_xmax, 0) AS FLOAT(53)) AS [BoundingBoxXMax], CAST(ISNULL(si.bounding_box_ymax, 0) AS FLOAT(53)) AS [BoundingBoxYMax], CAST(ISNULL(si.level_1_grid, 0) AS SMALLINT) AS [Level1Grid], CAST(ISNULL(si.level_2_grid, 0) AS SMALLINT) AS [Level2Grid], CAST(ISNULL(si.level_3_grid, 0) AS SMALLINT) AS [Level3Grid], CAST(ISNULL(si.level_4_grid, 0) AS SMALLINT) AS [Level4Grid], CAST(ISNULL(si.cells_per_object, 0) AS INT) AS [CellsPerObject], CAST(CASE WHEN i.type = 4 THEN 1 ELSE 0 END AS BIT) AS [IsSpatialIndex], i.has_filter AS [HasFilter], ISNULL(i.filter_definition, N'') AS [FilterDefinition], CAST(CASE i.type WHEN 1 THEN 0 WHEN 4 THEN 4 ELSE 1 END AS TINYINT) AS [IndexType], i.is_hypothetical AS [IsHypothetical] FROM sys.all_objects AS udf INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = udf.object_id ) LEFT OUTER JOIN sys.stats AS s ON s.stats_id = i.index_id AND s.object_id = i.object_id LEFT OUTER JOIN sys.key_constraints AS k ON k.parent_object_id = i.object_id AND k.unique_index_id = i.index_id LEFT OUTER JOIN sys.spatial_indexes AS spi ON i.object_id = spi.object_id AND i.index_id = spi.index_id LEFT OUTER JOIN sys.spatial_index_tessellations AS si ON i.object_id = si.object_id AND i.index_id = si.index_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = udf.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], i.name AS [Name], CAST(i.index_id AS INT) AS [ID], CAST(OBJECTPROPERTY(i.object_id, N'IsMSShipped') AS BIT) AS [IsSystemObject], ISNULL(s.no_recompute, 0) AS [NoAutomaticRecomputation], i.fill_factor AS [FillFactor], CAST(CASE i.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], i.is_primary_key + 2 * i.is_unique_constraint AS [IndexKeyType], i.is_unique AS [IsUnique], i.ignore_dup_key AS [IgnoreDuplicateKeys], ~i.allow_row_locks AS [DisallowRowLocks], ~i.allow_page_locks AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], i.is_disabled AS [IsDisabled], CAST(ISNULL(k.is_system_named, 0) AS BIT) AS [IsSystemNamed], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN i.type = 3 THEN 1 ELSE 0 END AS BIT) AS [IsXmlIndex], CASE UPPER(ISNULL(xi.secondary_type, '')) WHEN 'P' THEN 1 WHEN 'V' THEN 2 WHEN 'R' THEN 3 ELSE 0 END AS [SecondaryXmlIndexType], ISNULL(xi2.name, N'') AS [ParentXmlIndex], CAST(CASE i.type WHEN 1 THEN 0 WHEN 3 THEN CASE WHEN xi.using_xml_index_id IS NULL THEN 2 ELSE 3 END WHEN 4 THEN 4 WHEN 6 THEN 5 ELSE 1 END AS TINYINT) AS [IndexType], CAST(ISNULL(spi.spatial_index_type, 0) AS TINYINT) AS [SpatialIndexType], CAST(ISNULL(si.bounding_box_xmin, 0) AS FLOAT(53)) AS [BoundingBoxXMin], CAST(ISNULL(si.bounding_box_ymin, 0) AS FLOAT(53)) AS [BoundingBoxYMin], CAST(ISNULL(si.bounding_box_xmax, 0) AS FLOAT(53)) AS [BoundingBoxXMax], CAST(ISNULL(si.bounding_box_ymax, 0) AS FLOAT(53)) AS [BoundingBoxYMax], CAST(ISNULL(si.level_1_grid, 0) AS SMALLINT) AS [Level1Grid], CAST(ISNULL(si.level_2_grid, 0) AS SMALLINT) AS [Level2Grid], CAST(ISNULL(si.level_3_grid, 0) AS SMALLINT) AS [Level3Grid], CAST(ISNULL(si.level_4_grid, 0) AS SMALLINT) AS [Level4Grid], CAST(ISNULL(si.cells_per_object, 0) AS INT) AS [CellsPerObject], CAST(CASE WHEN i.type = 4 THEN 1 ELSE 0 END AS BIT) AS [IsSpatialIndex], i.has_filter AS [HasFilter], ISNULL(i.filter_definition, N'') AS [FilterDefinition], CASE WHEN 'FG' = dsi.type THEN dsi.name ELSE N'' END AS [FileGroup], CASE WHEN 'PS' = dsi.type THEN dsi.name ELSE N'' END AS [PartitionScheme], CAST(CASE WHEN 'PS' = dsi.type THEN 1 ELSE 0 END AS BIT) AS [IsPartitioned], CASE WHEN 'FD' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamFileGroup], CASE WHEN 'PS' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamPartitionScheme], CAST(CASE WHEN filetableobj.object_id IS NULL THEN 0 ELSE 1 END AS BIT) AS [IsFileTableDefined], i.is_hypothetical AS [IsHypothetical] FROM sys.all_objects AS udf INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = udf.object_id ) LEFT OUTER JOIN sys.stats AS s ON s.stats_id = i.index_id AND s.object_id = i.object_id LEFT OUTER JOIN sys.key_constraints AS k ON k.parent_object_id = i.object_id AND k.unique_index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi ON xi.object_id = i.object_id AND xi.index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi2 ON xi2.object_id = xi.object_id AND xi2.index_id = xi.using_xml_index_id LEFT OUTER JOIN sys.spatial_indexes AS spi ON i.object_id = spi.object_id AND i.index_id = spi.index_id LEFT OUTER JOIN sys.spatial_index_tessellations AS si ON i.object_id = si.object_id AND i.index_id = si.index_id LEFT OUTER JOIN sys.data_spaces AS dsi ON dsi.data_space_id = i.data_space_id LEFT OUTER JOIN sys.tables AS t ON t.object_id = i.object_id LEFT OUTER JOIN sys.data_spaces AS dstbl ON dstbl.data_space_id = t.Filestream_data_space_id AND i.index_id < 2 LEFT OUTER JOIN sys.filetable_system_defined_objects AS filetableobj ON i.object_id = filetableobj.object_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], i.name AS [Name], CAST(i.index_id AS INT) AS [ID], CAST(OBJECTPROPERTY(i.object_id, N'IsMSShipped') AS BIT) AS [IsSystemObject], ISNULL(s.no_recompute, 0) AS [NoAutomaticRecomputation], i.fill_factor AS [FillFactor], CAST(CASE i.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], i.is_primary_key + 2 * i.is_unique_constraint AS [IndexKeyType], i.is_unique AS [IsUnique], i.ignore_dup_key AS [IgnoreDuplicateKeys], ~i.allow_row_locks AS [DisallowRowLocks], ~i.allow_page_locks AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], i.is_disabled AS [IsDisabled], CAST(ISNULL(k.is_system_named, 0) AS BIT) AS [IsSystemNamed], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN i.type = 3 THEN 1 ELSE 0 END AS BIT) AS [IsXmlIndex], CASE UPPER(ISNULL(xi.secondary_type, '')) WHEN 'P' THEN 1 WHEN 'V' THEN 2 WHEN 'R' THEN 3 ELSE 0 END AS [SecondaryXmlIndexType], ISNULL(xi2.name, N'') AS [ParentXmlIndex], CAST(CASE i.type WHEN 1 THEN 0 WHEN 3 THEN CASE WHEN xi.using_xml_index_id IS NULL THEN 2 ELSE 3 END WHEN 4 THEN 4 WHEN 6 THEN 5 ELSE 1 END AS TINYINT) AS [IndexType], CAST(ISNULL(spi.spatial_index_type, 0) AS TINYINT) AS [SpatialIndexType], CAST(ISNULL(si.bounding_box_xmin, 0) AS FLOAT(53)) AS [BoundingBoxXMin], CAST(ISNULL(si.bounding_box_ymin, 0) AS FLOAT(53)) AS [BoundingBoxYMin], CAST(ISNULL(si.bounding_box_xmax, 0) AS FLOAT(53)) AS [BoundingBoxXMax], CAST(ISNULL(si.bounding_box_ymax, 0) AS FLOAT(53)) AS [BoundingBoxYMax], CAST(ISNULL(si.level_1_grid, 0) AS SMALLINT) AS [Level1Grid], CAST(ISNULL(si.level_2_grid, 0) AS SMALLINT) AS [Level2Grid], CAST(ISNULL(si.level_3_grid, 0) AS SMALLINT) AS [Level3Grid], CAST(ISNULL(si.level_4_grid, 0) AS SMALLINT) AS [Level4Grid], CAST(ISNULL(si.cells_per_object, 0) AS INT) AS [CellsPerObject], CAST(CASE WHEN i.type = 4 THEN 1 ELSE 0 END AS BIT) AS [IsSpatialIndex], i.has_filter AS [HasFilter], ISNULL(i.filter_definition, N'') AS [FilterDefinition], CASE WHEN 'FG' = dsi.type THEN dsi.name ELSE N'' END AS [FileGroup], CASE WHEN 'PS' = dsi.type THEN dsi.name ELSE N'' END AS [PartitionScheme], CAST(CASE WHEN 'PS' = dsi.type THEN 1 ELSE 0 END AS BIT) AS [IsPartitioned], CASE WHEN 'FD' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamFileGroup], CASE WHEN 'PS' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamPartitionScheme], i.is_hypothetical AS [IsHypothetical] FROM sys.all_objects AS udf INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = udf.object_id ) LEFT OUTER JOIN sys.stats AS s ON s.stats_id = i.index_id AND s.object_id = i.object_id LEFT OUTER JOIN sys.key_constraints AS k ON k.parent_object_id = i.object_id AND k.unique_index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi ON xi.object_id = i.object_id AND xi.index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi2 ON xi2.object_id = xi.object_id AND xi2.index_id = xi.using_xml_index_id LEFT OUTER JOIN sys.spatial_indexes AS spi ON i.object_id = spi.object_id AND i.index_id = spi.index_id LEFT OUTER JOIN sys.spatial_index_tessellations AS si ON i.object_id = si.object_id AND i.index_id = si.index_id LEFT OUTER JOIN sys.data_spaces AS dsi ON dsi.data_space_id = i.data_space_id LEFT OUTER JOIN sys.tables AS t ON t.object_id = i.object_id LEFT OUTER JOIN sys.data_spaces AS dstbl ON dstbl.data_space_id = t.Filestream_data_space_id AND i.index_id < 2 WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], i.name AS [Name], CAST(i.index_id AS INT) AS [ID], CAST(OBJECTPROPERTY(i.object_id, N'IsMSShipped') AS BIT) AS [IsSystemObject], ISNULL(s.no_recompute, 0) AS [NoAutomaticRecomputation], i.fill_factor AS [FillFactor], CAST(CASE i.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], i.is_primary_key + 2 * i.is_unique_constraint AS [IndexKeyType], i.is_unique AS [IsUnique], i.ignore_dup_key AS [IgnoreDuplicateKeys], ~i.allow_row_locks AS [DisallowRowLocks], ~i.allow_page_locks AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], i.is_disabled AS [IsDisabled], CAST(ISNULL(k.is_system_named, 0) AS BIT) AS [IsSystemNamed], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN i.type = 3 THEN 1 ELSE 0 END AS BIT) AS [IsXmlIndex], CASE UPPER(ISNULL(xi.secondary_type, '')) WHEN 'P' THEN 1 WHEN 'V' THEN 2 WHEN 'R' THEN 3 ELSE 0 END AS [SecondaryXmlIndexType], ISNULL(xi2.name, N'') AS [ParentXmlIndex], CAST(CASE i.type WHEN 1 THEN 0 WHEN 3 THEN CASE WHEN xi.using_xml_index_id IS NULL THEN 2 ELSE 3 END WHEN 4 THEN 4 WHEN 6 THEN 5 ELSE 1 END AS TINYINT) AS [IndexType], CAST(ISNULL(spi.spatial_index_type, 0) AS TINYINT) AS [SpatialIndexType], CAST(ISNULL(si.bounding_box_xmin, 0) AS FLOAT(53)) AS [BoundingBoxXMin], CAST(ISNULL(si.bounding_box_ymin, 0) AS FLOAT(53)) AS [BoundingBoxYMin], CAST(ISNULL(si.bounding_box_xmax, 0) AS FLOAT(53)) AS [BoundingBoxXMax], CAST(ISNULL(si.bounding_box_ymax, 0) AS FLOAT(53)) AS [BoundingBoxYMax], CAST(ISNULL(si.level_1_grid, 0) AS SMALLINT) AS [Level1Grid], CAST(ISNULL(si.level_2_grid, 0) AS SMALLINT) AS [Level2Grid], CAST(ISNULL(si.level_3_grid, 0) AS SMALLINT) AS [Level3Grid], CAST(ISNULL(si.level_4_grid, 0) AS SMALLINT) AS [Level4Grid], CAST(ISNULL(si.cells_per_object, 0) AS INT) AS [CellsPerObject], CAST(CASE WHEN i.type = 4 THEN 1 ELSE 0 END AS BIT) AS [IsSpatialIndex], i.has_filter AS [HasFilter], ISNULL(i.filter_definition, N'') AS [FilterDefinition], CASE WHEN 'FG' = dsi.type THEN dsi.name ELSE N'' END AS [FileGroup], CASE WHEN 'PS' = dsi.type THEN dsi.name ELSE N'' END AS [PartitionScheme], CAST(CASE WHEN 'PS' = dsi.type THEN 1 ELSE 0 END AS BIT) AS [IsPartitioned], CASE WHEN 'FD' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamFileGroup], CASE WHEN 'PS' = dstbl.type THEN dstbl.name ELSE N'' END AS [FileStreamPartitionScheme], i.is_hypothetical AS [IsHypothetical] FROM sys.all_objects AS udf INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = udf.object_id ) LEFT OUTER JOIN sys.stats AS s ON s.stats_id = i.index_id AND s.object_id = i.object_id LEFT OUTER JOIN sys.key_constraints AS k ON k.parent_object_id = i.object_id AND k.unique_index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi ON xi.object_id = i.object_id AND xi.index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi2 ON xi2.object_id = xi.object_id AND xi2.index_id = xi.using_xml_index_id LEFT OUTER JOIN sys.spatial_indexes AS spi ON i.object_id = spi.object_id AND i.index_id = spi.index_id LEFT OUTER JOIN sys.spatial_index_tessellations AS si ON i.object_id = si.object_id AND i.index_id = si.index_id LEFT OUTER JOIN sys.data_spaces AS dsi ON dsi.data_space_id = i.data_space_id LEFT OUTER JOIN sys.tables AS t ON t.object_id = i.object_id LEFT OUTER JOIN sys.data_spaces AS dstbl ON dstbl.data_space_id = t.Filestream_data_space_id AND i.index_id < 2 WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], i.name AS [Name], CAST(i.index_id AS INT) AS [ID], CAST(OBJECTPROPERTY(i.object_id, N'IsMSShipped') AS BIT) AS [IsSystemObject], ISNULL(s.no_recompute, 0) AS [NoAutomaticRecomputation], i.fill_factor AS [FillFactor], CAST(CASE i.index_id WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], i.is_primary_key + 2 * i.is_unique_constraint AS [IndexKeyType], i.is_unique AS [IsUnique], i.ignore_dup_key AS [IgnoreDuplicateKeys], ~i.allow_row_locks AS [DisallowRowLocks], ~i.allow_page_locks AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], i.is_disabled AS [IsDisabled], CAST(ISNULL(k.is_system_named, 0) AS BIT) AS [IsSystemNamed], CAST(INDEXPROPERTY(i.object_id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN i.type = 3 THEN 1 ELSE 0 END AS BIT) AS [IsXmlIndex], CASE UPPER(ISNULL(xi.secondary_type, '')) WHEN 'P' THEN 1 WHEN 'V' THEN 2 WHEN 'R' THEN 3 ELSE 0 END AS [SecondaryXmlIndexType], ISNULL(xi2.name, N'') AS [ParentXmlIndex], CAST(CASE i.type WHEN 1 THEN 0 WHEN 3 THEN CASE WHEN xi.using_xml_index_id IS NULL THEN 2 ELSE 3 END WHEN 4 THEN 4 WHEN 6 THEN 5 ELSE 1 END AS TINYINT) AS [IndexType], CASE WHEN 'FG' = dsi.type THEN dsi.name ELSE N'' END AS [FileGroup], CASE WHEN 'PS' = dsi.type THEN dsi.name ELSE N'' END AS [PartitionScheme], CAST(CASE WHEN 'PS' = dsi.type THEN 1 ELSE 0 END AS BIT) AS [IsPartitioned], i.is_hypothetical AS [IsHypothetical] FROM sys.all_objects AS udf INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = udf.object_id ) LEFT OUTER JOIN sys.stats AS s ON s.stats_id = i.index_id AND s.object_id = i.object_id LEFT OUTER JOIN sys.key_constraints AS k ON k.parent_object_id = i.object_id AND k.unique_index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi ON xi.object_id = i.object_id AND xi.index_id = i.index_id LEFT OUTER JOIN sys.xml_indexes AS xi2 ON xi2.object_id = xi.object_id AND xi2.index_id = xi.using_xml_index_id LEFT OUTER JOIN sys.data_spaces AS dsi ON dsi.data_space_id = i.data_space_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(udf.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(udf.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT udf.id AS [FunctionID], sudf.uid AS [SchemaID], i.name AS [Name], CAST(i.indid AS INT) AS [ID], CAST(OBJECTPROPERTY(i.id, N'IsMSShipped') AS BIT) AS [IsSystemObject], CAST(INDEXPROPERTY(i.id, i.name, N'IsFulltextKey') AS BIT) AS [IsFullTextKey], CAST(CASE WHEN ( i.status & 0x1000000 ) <> 0 THEN 1 ELSE 0 END AS BIT) AS [NoAutomaticRecomputation], CAST(INDEXPROPERTY(i.id, i.name, N'IndexFillFactor') AS TINYINT) AS [FillFactor], CAST(CASE i.indid WHEN 1 THEN 1 ELSE 0 END AS BIT) AS [IsClustered], CASE WHEN 0 != i.status & 0x800 THEN 1 WHEN 0 != i.status & 0x1000 THEN 2 ELSE 0 END AS [IndexKeyType], CAST(i.status & 2 AS BIT) AS [IsUnique], CAST(CASE WHEN 0 != ( i.status & 0x01 ) THEN 1 ELSE 0 END AS BIT) AS [IgnoreDuplicateKeys], CAST(INDEXPROPERTY(i.id, i.name, N'IsRowLockDisallowed') AS BIT) AS [DisallowRowLocks], CAST(INDEXPROPERTY(i.id, i.name, N'IsPageLockDisallowed') AS BIT) AS [DisallowPageLocks], CAST(INDEXPROPERTY(i.id, i.name, N'IsPadIndex') AS BIT) AS [PadIndex], CAST(ISNULL(k.status & 4, 0) AS BIT) AS [IsSystemNamed], CAST(CASE i.indid WHEN 1 THEN 0 ELSE 1 END AS TINYINT) AS [IndexType], fgi.groupname AS [FileGroup], CAST(INDEXPROPERTY(i.id, i.name, N'IsHypothetical') AS BIT) AS [IsHypothetical] FROM dbo.sysobjects AS udf INNER JOIN sysusers AS sudf ON sudf.uid = udf.uid INNER JOIN dbo.sysindexes AS i ON ( i.indid > 0 AND i.indid < 255 AND 1 != INDEXPROPERTY(i.id, i.name, N'IsStatistics') ) AND ( i.id = udf.id ) LEFT OUTER JOIN dbo.sysobjects AS k ON k.parent_obj = i.id AND k.name = i.name AND k.xtype IN ( N'PK', N'UQ' ) LEFT OUTER JOIN dbo.sysfilegroups AS fgi ON fgi.groupid = i.groupid WHERE ( udf.xtype IN ( 'TF', 'FN', 'IF' ) AND udf.name NOT LIKE N'#%%' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-UserDefinedFunctionIndexColumnQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], clmns.name AS [Name], ( CASE ic.key_ordinal WHEN 0 THEN ic.index_column_id ELSE ic.key_ordinal END ) AS [ID], CAST(COLUMNPROPERTY(ic.object_id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], ic.is_descending_key AS [Descending], ic.is_included_column AS [IsIncluded] FROM sys.all_objects AS udf INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = udf.object_id ) INNER JOIN sys.index_columns AS ic ON ( ic.column_id > 0 AND ( ic.key_ordinal > 0 OR ic.partition_ordinal = 0 OR ic.is_included_column != 0 ) ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = i.object_id ) INNER JOIN sys.columns AS clmns ON clmns.object_id = ic.object_id AND clmns.column_id = ic.column_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = udf.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], clmns.name AS [Name], ( CASE ic.key_ordinal WHEN 0 THEN ic.index_column_id ELSE ic.key_ordinal END ) AS [ID], CAST(COLUMNPROPERTY(ic.object_id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], ic.is_descending_key AS [Descending], ic.is_included_column AS [IsIncluded] FROM sys.all_objects AS udf INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = udf.object_id ) INNER JOIN sys.index_columns AS ic ON ( ic.column_id > 0 AND ( ic.key_ordinal > 0 OR ic.partition_ordinal = 0 OR ic.is_included_column != 0 ) ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = i.object_id ) INNER JOIN sys.columns AS clmns ON clmns.object_id = ic.object_id AND clmns.column_id = ic.column_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], clmns.name AS [Name], ( CASE ic.key_ordinal WHEN 0 THEN ic.index_column_id ELSE ic.key_ordinal END ) AS [ID], CAST(COLUMNPROPERTY(ic.object_id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], ic.is_descending_key AS [Descending], ic.is_included_column AS [IsIncluded] FROM sys.all_objects AS udf INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = udf.object_id ) INNER JOIN sys.index_columns AS ic ON ( ic.column_id > 0 AND ( ic.key_ordinal > 0 OR ic.partition_ordinal = 0 OR ic.is_included_column != 0 ) ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = i.object_id ) INNER JOIN sys.columns AS clmns ON clmns.object_id = ic.object_id AND clmns.column_id = ic.column_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], clmns.name AS [Name], ( CASE ic.key_ordinal WHEN 0 THEN ic.index_column_id ELSE ic.key_ordinal END ) AS [ID], CAST(COLUMNPROPERTY(ic.object_id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], ic.is_descending_key AS [Descending], ic.is_included_column AS [IsIncluded] FROM sys.all_objects AS udf INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = udf.object_id ) INNER JOIN sys.index_columns AS ic ON ( ic.column_id > 0 AND ( ic.key_ordinal > 0 OR ic.partition_ordinal = 0 OR ic.is_included_column != 0 ) ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = i.object_id ) INNER JOIN sys.columns AS clmns ON clmns.object_id = ic.object_id AND clmns.column_id = ic.column_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], CAST(i.index_id AS INT) AS [IndexID], clmns.name AS [Name], ( CASE ic.key_ordinal WHEN 0 THEN ic.index_column_id ELSE ic.key_ordinal END ) AS [ID], CAST(COLUMNPROPERTY(ic.object_id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], ic.is_descending_key AS [Descending], ic.is_included_column AS [IsIncluded] FROM sys.all_objects AS udf INNER JOIN sys.indexes AS i ON ( i.index_id > 0 ) AND ( i.object_id = udf.object_id ) INNER JOIN sys.index_columns AS ic ON ( ic.column_id > 0 AND ( ic.key_ordinal > 0 OR ic.partition_ordinal = 0 OR ic.is_included_column != 0 ) ) AND ( ic.index_id = CAST(i.index_id AS INT) AND ic.object_id = i.object_id ) INNER JOIN sys.columns AS clmns ON clmns.object_id = ic.object_id AND clmns.column_id = ic.column_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(udf.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(udf.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT udf.id AS [FunctionID], sudf.uid AS [SchemaID], CAST(i.indid AS INT) AS [IndexID], clmns.name AS [Name], CAST(ic.keyno AS INT) AS [ID], CAST(COLUMNPROPERTY(ic.id, clmns.name, N'IsComputed') AS BIT) AS [IsComputed], CAST(INDEXKEY_PROPERTY(ic.id, ic.indid, ic.keyno, N'IsDescending') AS BIT) AS [Descending] FROM dbo.sysobjects AS udf INNER JOIN sysusers AS sudf ON sudf.uid = udf.uid INNER JOIN dbo.sysindexes AS i ON ( i.indid > 0 AND i.indid < 255 AND 1 != INDEXPROPERTY(i.id, i.name, N'IsStatistics') ) AND ( i.id = udf.id ) INNER JOIN dbo.sysindexkeys AS ic ON CAST(ic.indid AS INT) = CAST(i.indid AS INT) AND ic.id = i.id INNER JOIN dbo.syscolumns AS clmns ON clmns.id = ic.id AND clmns.colid = ic.colid AND clmns.number = 0 WHERE ( udf.xtype IN ( 'TF', 'FN', 'IF' ) AND udf.name NOT LIKE N'#%%' ) 
"@ + $SystemObjectWhereClause
		}
	}
}

function Get-UserDefinedFunctionParameterQuery([System.Version]$ServerVersion, [String]$DatabaseEngineType, [Switch]$IncludeSystemObjects = $false) {

	if ($DatabaseEngineType -ieq $AzureDbEngine) {
		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], param.is_readonly AS [IsReadOnly], param.name AS [Name], param.parameter_id AS [ID], param.default_value AS [DefaultValue], param.has_default_value AS [HasDefaultValue], usrt.name AS [DataType], s1param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND param.max_length <> -1 THEN param.max_length / 2 ELSE param.max_length END AS INT) AS [Length], CAST(param.precision AS INT) AS [NumericPrecision], CAST(param.scale AS INT) AS [NumericScale], ISNULL(xscparam.name, N'') AS [XmlSchemaNamespace], ISNULL(s2param.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE param.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], udf.object_id AS [IDText], DB_NAME() AS [DatabaseName], param.name AS [ParamName], CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 ELSE 0 END AS BIT) AS [ParentSysObj], -1 AS [Number] FROM sys.all_objects AS udf INNER JOIN sys.all_parameters AS param ON ( param.is_output = 0 ) AND ( param.object_id = udf.object_id ) LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = param.user_type_id LEFT OUTER JOIN sys.schemas AS s1param ON s1param.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = param.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = param.system_type_id ) AND ( baset.user_type_id = param.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscparam ON xscparam.xml_collection_id = param.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2param ON s2param.schema_id = xscparam.schema_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause

	} else {

		$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
			' AND CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = udf.object_id AND minor_id = 0 AND class = 1 AND name = N''microsoft_database_tools_support'') IS NOT NULL THEN 1 ELSE 0 END AS BIT) = 0'
		} else {
			[String]::Empty
		}

		if ($ServerVersion.CompareTo($SQLServer2012) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], param.is_readonly AS [IsReadOnly], param.name AS [Name], param.parameter_id AS [ID], param.default_value AS [DefaultValue], param.has_default_value AS [HasDefaultValue], usrt.name AS [DataType], s1param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND param.max_length <> -1 THEN param.max_length / 2 ELSE param.max_length END AS INT) AS [Length], CAST(param.precision AS INT) AS [NumericPrecision], CAST(param.scale AS INT) AS [NumericScale], ISNULL(xscparam.name, N'') AS [XmlSchemaNamespace], ISNULL(s2param.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE param.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], udf.object_id AS [IDText], DB_NAME() AS [DatabaseName], param.name AS [ParamName], CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = udf.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [ParentSysObj], -1 AS [Number] FROM sys.all_objects AS udf INNER JOIN sys.all_parameters AS param ON ( param.is_output = 0 ) AND ( param.object_id = udf.object_id ) LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = param.user_type_id LEFT OUTER JOIN sys.schemas AS s1param ON s1param.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = param.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = param.system_type_id ) AND ( baset.user_type_id = param.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscparam ON xscparam.xml_collection_id = param.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2param ON s2param.schema_id = xscparam.schema_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008R2) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], param.is_readonly AS [IsReadOnly], param.name AS [Name], param.parameter_id AS [ID], param.default_value AS [DefaultValue], param.has_default_value AS [HasDefaultValue], usrt.name AS [DataType], s1param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND param.max_length <> -1 THEN param.max_length / 2 ELSE param.max_length END AS INT) AS [Length], CAST(param.precision AS INT) AS [NumericPrecision], CAST(param.scale AS INT) AS [NumericScale], ISNULL(xscparam.name, N'') AS [XmlSchemaNamespace], ISNULL(s2param.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE param.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], udf.object_id AS [IDText], DB_NAME() AS [DatabaseName], param.name AS [ParamName], CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = udf.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [ParentSysObj], -1 AS [Number] FROM sys.all_objects AS udf INNER JOIN sys.all_parameters AS param ON ( param.is_output = 0 ) AND ( param.object_id = udf.object_id ) LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = param.user_type_id LEFT OUTER JOIN sys.schemas AS s1param ON s1param.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = param.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = param.system_type_id ) AND ( baset.user_type_id = param.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscparam ON xscparam.xml_collection_id = param.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2param ON s2param.schema_id = xscparam.schema_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2008) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], param.is_readonly AS [IsReadOnly], param.name AS [Name], param.parameter_id AS [ID], param.default_value AS [DefaultValue], param.has_default_value AS [HasDefaultValue], usrt.name AS [DataType], s1param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND param.max_length <> -1 THEN param.max_length / 2 ELSE param.max_length END AS INT) AS [Length], CAST(param.precision AS INT) AS [NumericPrecision], CAST(param.scale AS INT) AS [NumericScale], ISNULL(xscparam.name, N'') AS [XmlSchemaNamespace], ISNULL(s2param.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE param.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], CASE WHEN usrt.is_table_type = 1 THEN N'structured' ELSE N'' END AS [UserType], udf.object_id AS [IDText], DB_NAME() AS [DatabaseName], param.name AS [ParamName], CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = udf.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [ParentSysObj], -1 AS [Number] FROM sys.all_objects AS udf INNER JOIN sys.all_parameters AS param ON ( param.is_output = 0 ) AND ( param.object_id = udf.object_id ) LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = param.user_type_id LEFT OUTER JOIN sys.schemas AS s1param ON s1param.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = param.system_type_id AND baset.user_type_id = baset.system_type_id ) OR ( ( baset.system_type_id = param.system_type_id ) AND ( baset.user_type_id = param.user_type_id ) AND ( baset.is_user_defined = 0 ) AND ( baset.is_assembly_type = 1 ) ) LEFT OUTER JOIN sys.xml_schema_collections AS xscparam ON xscparam.xml_collection_id = param.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2param ON s2param.schema_id = xscparam.schema_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2005) -ge 0) {
			@"
SELECT udf.object_id AS [FunctionID], udf.schema_id AS [SchemaID], param.name AS [Name], param.parameter_id AS [ID], param.default_value AS [DefaultValue], param.has_default_value AS [HasDefaultValue], usrt.name AS [DataType], s1param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'nchar', N'nvarchar' ) AND param.max_length <> -1 THEN param.max_length / 2 ELSE param.max_length END AS INT) AS [Length], CAST(param.precision AS INT) AS [NumericPrecision], CAST(param.scale AS INT) AS [NumericScale], ISNULL(xscparam.name, N'') AS [XmlSchemaNamespace], ISNULL(s2param.name, N'') AS [XmlSchemaNamespaceSchema], ISNULL(( CASE param.is_xml_document WHEN 1 THEN 2 ELSE 1 END ), 0) AS [XmlDocumentConstraint], udf.object_id AS [IDText], DB_NAME() AS [DatabaseName], param.name AS [ParamName], CAST(CASE WHEN udf.is_ms_shipped = 1 THEN 1 WHEN (SELECT major_id FROM sys.extended_properties WHERE major_id = udf.object_id AND minor_id = 0 AND class = 1 AND name = N'microsoft_database_tools_support') IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS [ParentSysObj], -1 AS [Number] FROM sys.all_objects AS udf INNER JOIN sys.all_parameters AS param ON ( param.is_output = 0 ) AND ( param.object_id = udf.object_id ) LEFT OUTER JOIN sys.types AS usrt ON usrt.user_type_id = param.user_type_id LEFT OUTER JOIN sys.schemas AS s1param ON s1param.schema_id = usrt.schema_id LEFT OUTER JOIN sys.types AS baset ON ( baset.user_type_id = param.system_type_id AND baset.user_type_id = baset.system_type_id ) LEFT OUTER JOIN sys.xml_schema_collections AS xscparam ON xscparam.xml_collection_id = param.xml_collection_id LEFT OUTER JOIN sys.schemas AS s2param ON s2param.schema_id = xscparam.schema_id WHERE ( udf.type IN ( 'TF', 'FN', 'IF', 'FS', 'FT' ) ) 
"@ + $SystemObjectWhereClause
		}
		elseif ($ServerVersion.CompareTo($SQLServer2000) -ge 0) {
			$SystemObjectWhereClause = if ($IncludeSystemObjects -ne $true) {
				' AND CAST(CASE WHEN ( OBJECTPROPERTY(udf.id, N''IsMSShipped'') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(udf.id, N''IsSystemTable'') THEN 1 ELSE 0 END AS BIT) = 0'
			} else {
				[String]::Empty
			}

			@"
SELECT udf.id AS [FunctionID], sudf.uid AS [SchemaID], param.name AS [Name], CAST(param.colid AS INT) AS [ID], NULL AS [DefaultValue], usrt.name AS [DataType], s1param.name AS [DataTypeSchema], ISNULL(baset.name, N'') AS [SystemType], CAST(CASE WHEN baset.name IN ( N'char', N'varchar', N'binary', N'varbinary', N'nchar', N'nvarchar' ) THEN param.prec ELSE param.length END AS INT) AS [Length], CAST(param.xprec AS INT) AS [NumericPrecision], CAST(param.xscale AS INT) AS [NumericScale], udf.id AS [IDText], DB_NAME() AS [DatabaseName], param.name AS [ParamName], CAST(CASE WHEN ( OBJECTPROPERTY(udf.id, N'IsMSShipped') = 1 ) THEN 1 WHEN 1 = OBJECTPROPERTY(udf.id, N'IsSystemTable') THEN 1 ELSE 0 END AS BIT) AS [ParentSysObj], -1 AS [Number] FROM dbo.sysobjects AS udf INNER JOIN sysusers AS sudf ON sudf.uid = udf.uid INNER JOIN syscolumns AS param ON ( param.number = 1 OR ( param.number = 0 AND 1 = OBJECTPROPERTY(param.id, N'IsScalarFunction') AND ISNULL(param.name, '') != '' ) ) AND ( param.id = udf.id ) LEFT OUTER JOIN systypes AS usrt ON usrt.xusertype = param.xusertype LEFT OUTER JOIN sysusers AS s1param ON s1param.uid = usrt.uid LEFT OUTER JOIN systypes AS baset ON baset.xusertype = param.xtype AND baset.xusertype = baset.xtype WHERE ( udf.xtype IN ( 'TF', 'FN', 'IF' ) AND udf.name NOT LIKE N'#%%' ) 
"@ + $SystemObjectWhereClause
		}
	}
}


<#
	.SYNOPSIS
		A brief description of the function.
	.DESCRIPTION
		A detailed description of the function.
	.PARAMETER  param1
		The description of the first parameter.
	.EXAMPLE
		PS C:\> Get-Foo -param1 'string value'
	.INPUTS
		System.String
	.OUTPUTS
		System.String
	.NOTES
		Additional information about the function goes here.
	.LINK
		about_functions_advanced
	.LINK
		about_comment_based_help
#>
function Get-NTGroupMemberList {
	[CmdletBinding(DefaultParametersetName='domain')]
	param(
		[Parameter(Position=0, Mandatory=$true, ParameterSetName='domain')]
		[ValidateLength(1,15)]
		[System.String]
		$NTDomainName
		,
		[Parameter(Position=0, Mandatory=$true, ParameterSetName='machine')]
		[ValidateLength(1,15)]
		[System.String]
		$NTMachineName
		,
		[Parameter(Position=1, Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$GroupName
		,
		[Parameter(Position=2, Mandatory=$false)]
		[Switch]
		$Recurse
	)
	begin {
		Add-Type -AssemblyName System.DirectoryServices.AccountManagement
		$NTAccountType = [System.Security.Principal.NTAccount] -as [Type]

		if ($PsCmdlet.ParameterSetName -ieq 'domain') {
			$ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
			$PrincipalContext = New-Object -TypeName System.DirectoryServices.AccountManagement.PrincipalContext -ArgumentList $ContextType, $NTDomainName
		} else {
			$ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Machine
			$PrincipalContext = New-Object -TypeName System.DirectoryServices.AccountManagement.PrincipalContext -ArgumentList $ContextType, $NTMachineName
		}
	}
	process {
		$Group = [System.DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($PrincipalContext,$GroupName)
		
		if ($Group) {
		
			try {
				$Group.GetMembers($Recurse) | ForEach-Object {
					Write-Output (
						New-Object -TypeName PSObject -Property @{
							AccountExpirationDate = $_.AccountExpirationDate
							AccountLockoutTime = $_.AccountLockoutTime
							Description = $_.Description
							DisplayName = $_.DisplayName
							Enabled = $_.Enabled # In Win32_UserAccount this is called "Disabled". Hooray for inconsistencies!
							Name = $_.Name
							SamAccountName = $_.SamAccountName
							NTDomainName = if ($_.ContextType -ieq 'domain') { 
								$_.Sid.Translate($NTAccountType).ToString().Split('\')[0] 
							} elseif ($PsCmdlet.ParameterSetName -ieq 'machine') {
								$NTMachineName
							} else { 
								$null 
							}
							NTAccountName = if ($_.ContextType -ieq 'domain') { 
								$_.Sid.Translate($NTAccountType).ToString() 
							} elseif ($PsCmdlet.ParameterSetName -ieq 'machine') {
								'{0}\{1}' -f $NTMachineName, $_.SamAccountName
							} else { 
								$null 
							}
							Sid = $_.Sid
							PrincipalType = if ($PsCmdlet.ParameterSetName -ieq 'domain') {
								$_.StructuralObjectClass
							} else {
								if ($_.ContextType.ToString() -ieq 'domain') { 
									$_.StructuralObjectClass 
								} else { 
									'user' 
								}
							}
							ContextType = $_.ContextType.ToString()
						}
					)
				}
			}
			catch {
				Write-Output (
					New-Object -TypeName PSObject -Property @{
						Description = '*Incomplete List'
						DisplayName = '*Incomplete List'
						Name = '*Incomplete List'
						NTAccountName = '*Incomplete List'
					}
				)
				throw
			}
		} else {
			throw 'Unable to find a group that matches the provided identity. Check that the group still exists; if not it may have been orphaned.'
		}
	}
	end {
		Remove-Variable -Name ContextType, NTAccountType, PrincipalContext, Group
	}
}


function Get-AgentActivationOrderValue($ActivationOrder) {

	$ActivationOrderEnum = 'Microsoft.SqlServer.Management.Smo.Agent.ActivationOrder' -as [Type]

	if (-not $ActivationOrderEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($ActivationOrder) {
				$($ActivationOrderEnum::First) {'First'}
				$($ActivationOrderEnum::None) {'None'}
				$($ActivationOrderEnum::Last) {'Last'}

				# Added the following to support bypassing SMO when retrieving database objects
				$( $ActivationOrderEnum::First).value__ {'First'}
				$( $ActivationOrderEnum::None).value__ {'None'}
				$( $ActivationOrderEnum::Last).value__ {'Last'}

				$null {$null} #$null {[String]$null}
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name ActivationOrderEnum
}

function Get-AgentCompletionActionValue($CompletionAction) {
	$CompletionActionEnum = 'Microsoft.SqlServer.Management.Smo.Agent.CompletionAction' -as [Type]

	if (-not $CompletionActionEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($CompletionAction) {
				$($CompletionActionEnum::Never) {'Never'}
				$($CompletionActionEnum::OnSuccess) {'When the job succeeds'}
				$($CompletionActionEnum::OnFailure) {'When the job fails'}
				$($CompletionActionEnum::Always) {'Whenever the job completes'}
				$null {$null} #$null {[String]$null}
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name CompletionActionEnum
}

function Get-AgentCompletionResultValue($CompletionResult) {
	$CompletionResultEnum = 'Microsoft.SqlServer.Management.Smo.Agent.CompletionResult' -as [Type]

	if (-not $CompletionResultEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($CompletionResult) {
				$($CompletionResultEnum::Failed) {'Failed'}
				$($CompletionResultEnum::Succeeded) {'Succeeded'}
				$($CompletionResultEnum::Retry) {'Retrying'}
				$($CompletionResultEnum::Cancelled) {'Cancelled'}
				$($CompletionResultEnum::InProgress) {'In Progress'}
				$($CompletionResultEnum::Unknown) {'Unknown'}
				$null {$null} #$null {[String]$null}
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name CompletionResultEnum
}

function Get-AgentLogLevelValue($AgentLogLevel) {
	# http://msdn.microsoft.com/en-us/library/microsoft.sqlserver.management.smo.agent.notifymethods
	# This is a cheat until I can figure out how to deal with FlagsAttribute in PowerShell

	if ($AgentLogLevel) {
		Write-Output $AgentLogLevel.ToString()
	} else {
		Write-Output $null #Write-Output [String]$null
	}
}

function Get-AgentMailTypeValue($AgentMailType) {
	$AgentMailTypeEnum = 'Microsoft.SqlServer.Management.Smo.Agent.AgentMailType' -as [Type]

	if (-not $AgentMailTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($AgentMailType) {
				$($AgentMailTypeEnum::SqlAgentMail) {'Agent Mail'}
				$($AgentMailTypeEnum::DatabaseMail) {'Database Mail'}
				$null {$null} #$null {[String]$null}
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name AgentMailTypeEnum
}

function Get-AgentStepCompletionActionValue($StepCompletionAction) {
	$StepCompletionActionEnum = 'Microsoft.SqlServer.Management.Smo.Agent.StepCompletionAction' -as [Type]

	if (-not $StepCompletionActionEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($StepCompletionAction) {
				$($StepCompletionActionEnum::QuitWithSuccess) {'Quit the job reporting succces'}
				$($StepCompletionActionEnum::QuitWithFailure) {'Quit the job reporting failure'}
				$($StepCompletionActionEnum::GoToNextStep) {'Go to the next step'}
				$($StepCompletionActionEnum::GoToStep) {'Go to step'}
				$null {$null} #$null {[String]$null}
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name StepCompletionActionEnum
}

function Get-AgentSubSystemValue($AgentSubSystem) {
	$AgentSubSystemEnum = 'Microsoft.SqlServer.Management.Smo.Agent.AgentSubSystem' -as [Type]

	if (-not $AgentSubSystemEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($AgentSubSystem) {
				$($AgentSubSystemEnum::TransactSql) {'Transact-SQL script (T-SQL)'}
				$($AgentSubSystemEnum::ActiveScripting) {'ActiveX Script'}
				$($AgentSubSystemEnum::CmdExec) {'Operating System (CmdExec)'}
				$($AgentSubSystemEnum::Snapshot) {'Replication Snapshot'}
				$($AgentSubSystemEnum::LogReader) {'Replication Transaction-Log Reader'}
				$($AgentSubSystemEnum::Distribution) {'Replication Distributor'}
				$($AgentSubSystemEnum::Merge) {'Replication Merge'}
				$($AgentSubSystemEnum::QueueReader) {'Replication Queue Reader'}
				$($AgentSubSystemEnum::AnalysisQuery) {'SQL Server Analysis Services Query'}
				$($AgentSubSystemEnum::AnalysisCommand) {'SQL Server Analysis Services Command'}
				$($AgentSubSystemEnum::Ssis) {'SQL Server Integration Services Package'}
				$($AgentSubSystemEnum::PowerShell) {'PowerShell'}
				$null {$null} #$null {[String]$null}
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name AgentSubSystemEnum
}

function Get-AgentWeekdaysValue($Weekdays) {
	# http://msdn.microsoft.com/en-us/library/microsoft.sqlserver.management.smo.agent.operator.pagerdays.aspx
	# This is a cheat until I can figure out how to deal with FlagsAttribute in PowerShell

	if (($Weekdays) -and ($Weekdays -gt 0)) {
		Write-Output $Weekdays.ToString()
	} else {
		Write-Output $null #Write-Output [String]$null
	}
}

function Get-AlertTypeValue($AlertType) {
	$AlertTypeEnum = 'Microsoft.SqlServer.Management.Smo.Agent.AlertType' -as [Type]

	if (-not $AlertTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($AlertType) {
				$($AlertTypeEnum::SqlServerEvent) {'SQL Server event'}
				$($AlertTypeEnum::SqlServerPerformanceCondition) {'SQL Server performance condition'}
				$($AlertTypeEnum::NonSqlServerEvent) {'Non-SQL Server event'}
				$($AlertTypeEnum::WmiEvent) {'WMI event'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name AlertTypeEnum
}

function Get-AssemblySecurityLevelValue($AssemblySecurityLevel) {
	$AssemblySecurityLevelEnum = 'Microsoft.SqlServer.Management.Smo.AssemblySecurityLevel' -as [Type]

	if (-not $AssemblySecurityLevelEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($AlertType) {
				$($AssemblySecurityLevelEnum::Safe) {'Safe'}
				$($AssemblySecurityLevelEnum::External) {'External'}
				$($AssemblySecurityLevelEnum::Unrestricted) {'Unrestricted'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name AssemblySecurityLevelEnum
}

function Get-AsymmetricKeyEncryptionAlgorithmValue($AsymmetricKeyEncryptionAlgorithm) {
	$AsymmetricKeyEncryptionAlgorithmEnum = 'Microsoft.SqlServer.Management.Smo.AsymmetricKeyEncryptionAlgorithm' -as [Type]

	if (-not $AsymmetricKeyEncryptionAlgorithmEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($AsymmetricKeyEncryptionAlgorithm) {
				$($AsymmetricKeyEncryptionAlgorithmEnum::CryptographicProviderDefined) {'Cryptographic Provider'}
				$($AsymmetricKeyEncryptionAlgorithmEnum::Rsa512) {'512-bit RSA encryption algorithm'}
				$($AsymmetricKeyEncryptionAlgorithmEnum::Rsa1024) {'1024-bit RSA encryption algorithm'}
				$($AsymmetricKeyEncryptionAlgorithmEnum::Rsa2048) {'2048-bit RSA encryption algorithm'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name AsymmetricKeyEncryptionAlgorithmEnum
}


# See http://msdn.microsoft.com/en-us/library/cc280663.aspx
function Get-AuditActionTypeValue($AuditActionType) {
	$AuditActionTypeEnum = 'Microsoft.SqlServer.Management.Smo.AuditActionType' -as [Type]

	if (-not $AuditActionTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($AuditActionType) {
				$($AuditActionTypeEnum::ApplicationRoleChangePasswordGroup) {'APPLICATION_ROLE_CHANGE_PASSWORD_GROUP'}
				$($AuditActionTypeEnum::AuditChangeGroup) {'AUDIT_CHANGE_GROUP'}
				$($AuditActionTypeEnum::BackupRestoreGroup) {'BACKUP_RESTORE_GROUP'}
				$($AuditActionTypeEnum::BrokerLoginGroup) {'BROKER_LOGIN_GROUP'}
				$($AuditActionTypeEnum::DatabaseChangeGroup) {'DATABASE_CHANGE_GROUP'}
				$($AuditActionTypeEnum::DatabaseLogoutGroup) {'DATABASE_LOGOUT_GROUP'}
				$($AuditActionTypeEnum::DatabaseMirroringLoginGroup) {'DATABASE_MIRRORING_LOGIN_GROUP'}
				$($AuditActionTypeEnum::DatabaseObjectAccessGroup) {'DATABASE_OBJECT_ACCESS_GROUP'}
				$($AuditActionTypeEnum::DatabaseObjectChangeGroup) {'DATABASE_OBJECT_CHANGE_GROUP'}
				$($AuditActionTypeEnum::DatabaseObjectOwnershipChangeGroup) {'DATABASE_OBJECT_OWNERSHIP_CHANGE_GROUP'}
				$($AuditActionTypeEnum::DatabaseObjectPermissionChangeGroup) {'DATABASE_OBJECT_PERMISSION_CHANGE_GROUP'}
				$($AuditActionTypeEnum::DatabaseOperationGroup) {'DATABASE_OPERATION_GROUP'}
				$($AuditActionTypeEnum::DatabaseOwnershipChangeGroup) {'DATABASE_OWNERSHIP_CHANGE_GROUP'}
				$($AuditActionTypeEnum::DatabasePermissionChangeGroup) {'DATABASE_PERMISSION_CHANGE_GROUP'}
				$($AuditActionTypeEnum::DatabasePrincipalChangeGroup) {'DATABASE_PRINCIPAL_CHANGE_GROUP'}
				$($AuditActionTypeEnum::DatabasePrincipalImpersonationGroup) {'DATABASE_PRINCIPAL_IMPERSONATION_GROUP'}
				$($AuditActionTypeEnum::DatabaseRoleMemberChangeGroup) {'DATABASE_ROLE_MEMBER_CHANGE_GROUP'}
				$($AuditActionTypeEnum::DbccGroup) {'DBCC_GROUP'}
				$($AuditActionTypeEnum::Delete) {'DELETE'}
				$($AuditActionTypeEnum::Execute) {'EXECUTE'}
				$($AuditActionTypeEnum::FailedDatabaseAuthenticationGroup) {'FAILED_DATABASE_AUTHENTICATION_GROUP'}
				$($AuditActionTypeEnum::FailedLoginGroup) {'FAILED_LOGIN_GROUP'}
				$($AuditActionTypeEnum::FullTextGroup) {'FULLTEXT_GROUP'}
				$($AuditActionTypeEnum::Insert) {'INSERT'}
				$($AuditActionTypeEnum::LoginChangePasswordGroup) {'LOGIN_CHANGE_PASSWORD_GROUP'}
				$($AuditActionTypeEnum::LogoutGroup) {'LOGOUT_GROUP'}
				$($AuditActionTypeEnum::Receive) {'RECEIVE'}
				$($AuditActionTypeEnum::References) {'REFERENCES'}
				$($AuditActionTypeEnum::SchemaObjectAccessGroup) {'SCHEMA_OBJECT_ACCESS_GROUP'}
				$($AuditActionTypeEnum::SchemaObjectChangeGroup) {'SCHEMA_OBJECT_CHANGE_GROUP'}
				$($AuditActionTypeEnum::SchemaObjectOwnershipChangeGroup) {'SCHEMA_OBJECT_OWNERSHIP_CHANGE_GROUP'}
				$($AuditActionTypeEnum::SchemaObjectPermissionChangeGroup) {'SCHEMA_OBJECT_PERMISSION_CHANGE_GROUP'}
				$($AuditActionTypeEnum::Select) {'SELECT'}
				$($AuditActionTypeEnum::ServerObjectChangeGroup) {'SERVER_OBJECT_CHANGE_GROUP'}
				$($AuditActionTypeEnum::ServerObjectOwnershipChangeGroup) {'SERVER_OBJECT_OWNERSHIP_CHANGE_GROUP'}
				$($AuditActionTypeEnum::ServerObjectPermissionChangeGroup) {'SERVER_OBJECT_PERMISSION_CHANGE_GROUP'}
				$($AuditActionTypeEnum::ServerOperationGroup) {'SERVER_OPERATION_GROUP'}
				$($AuditActionTypeEnum::ServerPermissionChangeGroup) {'SERVER_PERMISSION_CHANGE_GROUP'}
				$($AuditActionTypeEnum::ServerPrincipalChangeGroup) {'SERVER_PRINCIPAL_CHANGE_GROUP'}
				$($AuditActionTypeEnum::ServerPrincipalImpersonationGroup) {'SERVER_PRINCIPAL_IMPERSONATION_GROUP'}
				$($AuditActionTypeEnum::ServerRoleMemberChangeGroup) {'SERVER_ROLE_MEMBER_CHANGE_GROUP'}
				$($AuditActionTypeEnum::ServerStateChangeGroup) {'SERVER_STATE_CHANGE_GROUP'}
				$($AuditActionTypeEnum::SuccessfulDatabaseAuthenticationGroup) {'SUCCESSFUL_DATABASE_AUTHENTICATION_GROUP'}
				$($AuditActionTypeEnum::SuccessfulLoginGroup) {'SUCCESSFUL_LOGIN_GROUP'}
				$($AuditActionTypeEnum::TraceChangeGroup) {'TRACE_CHANGE_GROUP'}
				$($AuditActionTypeEnum::Update) {'UPDATE'}
				$($AuditActionTypeEnum::UserChangePasswordGroup) {'USER_CHANGE_PASSWORD_GROUP'}
				$($AuditActionTypeEnum::UserDefinedAuditGroup) {'USER_DEFINED_AUDIT_GROUP'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name AuditActionTypeEnum 
}

function Get-AuditDestinationTypeValue($AuditDestinationType) {
	$AuditDestinationTypeEnum = 'Microsoft.SqlServer.Management.Smo.AuditDestinationType' -as [Type]

	if (-not $AuditDestinationTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($AuditDestinationType) {
				$($AuditDestinationTypeEnum::File) {'File'}
				$($AuditDestinationTypeEnum::SecurityLog) {'Security Log'}
				$($AuditDestinationTypeEnum::ApplicationLog) {'Application Log'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name AuditDestinationTypeEnum 
}

function Get-AuditFileSizeUnitValue($AuditFileSizeUnit) {
	$AuditFileSizeUnitEnum = 'Microsoft.SqlServer.Management.Smo.AuditFileSizeUnit' -as [Type]

	if (-not $AuditFileSizeUnitEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($AuditFileSizeUnit) {
				$($AuditFileSizeUnitEnum::Mb) {'MB'}
				$($AuditFileSizeUnitEnum::Gb) {'GB'}
				$($AuditFileSizeUnitEnum::Tb) {'TB'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name AuditFileSizeUnitEnum 
}

function Get-AuditLevelValue($AuditLevel) {

	# Added to SMO 2008 (v10), will be $null if using SMO 2005 (v9)
	$AuditLevelEnum = 'Microsoft.SqlServer.Management.Smo.AuditLevel' -as [Type]

	if (-not $AuditLevelEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($AuditLevel) {
				$($AuditLevelEnum::None) {'None'}
				$($AuditLevelEnum::Success) {'Successful logins only'}
				$($AuditLevelEnum::Failure) {'Failed logins only'}
				$($AuditLevelEnum::All) {'Both failed and successful logins'}
				$null { $null }
				default { $_.ToString() }
			}
		)
	}

	Remove-Variable -Name AuditLevelEnum
}

function Get-AuthenticationModeValue($LoginMode) {
	$ServerLoginModeEnum = 'Microsoft.SqlServer.Management.Smo.ServerLoginMode' -as [Type]

	if (-not $ServerLoginModeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($LoginMode) {
				$($ServerLoginModeEnum::Normal) {'SQL Server Authentication'}
				$($ServerLoginModeEnum::Integrated) {'Windows Authentication'}
				$($ServerLoginModeEnum::Mixed) {'SQL Server and Windows Authentication'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name ServerLoginModeEnum 
}

function Get-AvailabilityDatabaseSynchronizationStatusValue($AvailabilityDatabaseSynchronizationStatus) {

	# Added to SMO 2012 (v11), will be $null if using SMO 2005 (v9) or SMO 2008 (v10)
	$AvailabilityDatabaseSynchronizationStateEnum = 'Microsoft.SqlServer.Management.Smo.AvailabilityDatabaseSynchronizationState' -as [Type]

	if (-not $AvailabilityDatabaseSynchronizationStateEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($AvailabilityDatabaseSynchronizationStatus) {
				$($AvailabilityDatabaseSynchronizationStateEnum::NotSynchronizing) {'Not Synchronizing'}
				$($AvailabilityDatabaseSynchronizationStateEnum::Synchronizing) {'Synchronizing'}
				$($AvailabilityDatabaseSynchronizationStateEnum::Synchronized) {'Synchronized'}
				$($AvailabilityDatabaseSynchronizationStateEnum::Reverting) {'Reverting'}
				$($AvailabilityDatabaseSynchronizationStateEnum::Initializing) {'Initializing'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name AvailabilityDatabaseSynchronizationStateEnum
}

function Get-BrokerMessageSourceValue($MessageSource) {

	$MessageSourceEnum = 'Microsoft.SqlServer.Management.Smo.Broker.MessageSource' -as [Type]

	if (-not $MessageSourceEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($MessageSource) {
				$($MessageSourceEnum::Initiator) {'Initiator'}
				$($MessageSourceEnum::Target) {'Target'}
				$($MessageSourceEnum::InitiatorAndTarget) {'Initiator And Target'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name MessageSourceEnum
}

function Get-BrokerMessageTypeValidationValue($MessageTypeValidation) {

	$MessageTypeValidationEnum = 'Microsoft.SqlServer.Management.Smo.Broker.MessageTypeValidation' -as [Type]

	if (-not $MessageTypeValidationEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($MessageTypeValidation) {
				$($MessageTypeValidationEnum::None) {'None'}
				$($MessageTypeValidationEnum::XmlSchemaCollection) {'XML Schema Collection'}
				$($MessageTypeValidationEnum::Empty) {'Empty'}
				$($MessageTypeValidationEnum::Xml) {'XML'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name MessageTypeValidationEnum
}

function Get-CatalogPopulationStatusValue($CatalogPopulationStatus) {

	$CatalogPopulationStatusEnum = 'Microsoft.SqlServer.Management.Smo.CatalogPopulationStatus' -as [Type]

	if (-not $CatalogPopulationStatusEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($CatalogPopulationStatus) {
				$($CatalogPopulationStatusEnum::Idle) {'Idle'}
				$($CatalogPopulationStatusEnum::CrawlinProgress) {'Crawl In Progress'}
				$($CatalogPopulationStatusEnum::Paused) {'Paused'}
				$($CatalogPopulationStatusEnum::Throttled) {'Throttled'}
				$($CatalogPopulationStatusEnum::Recovering) {'Recovering'}
				$($CatalogPopulationStatusEnum::Shutdown) {'Shutdown'}
				$($CatalogPopulationStatusEnum::Incremental) {'Incremental'}
				$($CatalogPopulationStatusEnum::UpdatingIndex) {'Updating Index'}
				$($CatalogPopulationStatusEnum::DiskFullPause) {'Disk Full'}
				$($CatalogPopulationStatusEnum::Notification) {'Notification'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name CatalogPopulationStatusEnum
}

function Get-ChangeTrackingValue($ChangeTracking) {

	$ChangeTrackingEnum = 'Microsoft.SqlServer.Management.Smo.ChangeTracking' -as [Type]

	if (-not $ChangeTrackingEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($ChangeTracking) {
				$($ChangeTrackingEnum::Off) {'Off'}
				$($ChangeTrackingEnum::Automatic) {'Automatic'}
				$($ChangeTrackingEnum::Manual) {'Manual'}

				# Added the following to support bypassing SMO when retrieving database objects
				$($ChangeTrackingEnum::Off).value__ {'Off'}
				$($ChangeTrackingEnum::Automatic).value__ {'Automatic'}
				$($ChangeTrackingEnum::Manual).value__ {'Manual'}

				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name ChangeTrackingEnum
}

function Get-ClusterQuorumStateValue($ClusterQuorumState) {

	# Added to SMO 2012 (v11), will be $null otherwise
	$ClusterQuorumStateEnum = 'Microsoft.SqlServer.Management.Smo.ClusterQuorumState' -as [Type]

	if (-not $ClusterQuorumStateEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($ClusterQuorumState) {
				$($ClusterQuorumStateEnum::UnknownQuorumState) {'Uknown'}
				$($ClusterQuorumStateEnum::NormalQuorum) {'Normal'}
				$($ClusterQuorumStateEnum::ForcedQuorum) {'Forced'}
				$($ClusterQuorumStateEnum::NotApplicable) {'Not Applicable'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name ClusterQuorumStateEnum 
}

function Get-ClusterQuorumTypeValue($ClusterQuorumType) {

	# Added to SMO 2012 (v11), will be $null otherwise
	$ClusterQuorumTypeEnum = 'Microsoft.SqlServer.Management.Smo.ClusterQuorumType' -as [Type]

	if (-not $ClusterQuorumTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($ClusterQuorumType) {
				$($ClusterQuorumTypeEnum::NodeMajority) {'Node Majority'}
				$($ClusterQuorumTypeEnum::NodeAndDiskMajority) {'Node and Disk Majority'}
				$($ClusterQuorumTypeEnum::NodeAndFileshareMajority) {'Node and Fileshare Majority'}
				$($ClusterQuorumTypeEnum::DiskOnly) {'Disk Only'}
				$($ClusterQuorumTypeEnum::NotApplicable) {'Not Applicable'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name ClusterQuorumTypeEnum 
}

function Get-DatabaseDdlTriggerExecutionContextValue($DatabaseDdlTriggerExecutionContext) {

	$DatabaseDdlTriggerExecutionContextEnum = 'Microsoft.SqlServer.Management.Smo.DatabaseDdlTriggerExecutionContext' -as [Type]

	if (-not $DatabaseDdlTriggerExecutionContextEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($DatabaseDdlTriggerExecutionContext) {
				$($DatabaseDdlTriggerExecutionContextEnum::Caller) {'Caller'}
				$($DatabaseDdlTriggerExecutionContextEnum::ExecuteAsUser) {'Execute As User'}
				$($DatabaseDdlTriggerExecutionContextEnum::Self) {'Self'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name DatabaseDdlTriggerExecutionContextEnum
}

function Get-DatabaseEngineTypeValue($DatabaseEngineType) {

	# Added to SMO 2008 (v10), will be $null if using SMO 2005 (v9)
	$DatabaseEngineTypeEnum = 'Microsoft.SqlServer.Management.Common.DatabaseEngineType' -as [Type]

	if (-not $DatabaseEngineTypeEnum) {
		# Assume we're dealing w\ a Standalone engine if the datatype can't be loaded
		Write-Output $StandaloneDbEngine
		#Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($DatabaseEngineType) {
				$($DatabaseEngineTypeEnum::Standalone) {$StandaloneDbEngine}
				$($DatabaseEngineTypeEnum::SqlAzureDatabase) {$AzureDbEngine}
				$($DatabaseEngineTypeEnum::Unknown) {'Unknown'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name DatabaseEngineTypeEnum
}

function Get-DataCompressionTypeValue($DataCompressionType) {

	$DataCompressionTypeEnum = 'Microsoft.SqlServer.Management.Smo.DataCompressionType' -as [Type]

	if (-not $DataCompressionTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($DataCompressionType) {
				$($DataCompressionTypeEnum::None) {'None'}
				$($DataCompressionTypeEnum::Row) {'Row'}
				$($DataCompressionTypeEnum::Page) {'Page'}
				$($DataCompressionTypeEnum::ColumnStore) {'Column Store'}

				# Added the following to support bypassing SMO when retrieving database objects
				$($DataCompressionTypeEnum::None).value__ {'None'}
				$($DataCompressionTypeEnum::Row).value__ {'Row'}
				$($DataCompressionTypeEnum::Page).value__ {'Page'}
				$($DataCompressionTypeEnum::ColumnStore).value__ {'Column Store'}

				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name DataCompressionTypeEnum
}

function Get-DatabaseUserAccessValue($UserAccess) {
	$DatabaseUserAccessEnum = 'Microsoft.SqlServer.Management.Smo.DatabaseUserAccess' -as [Type]

	if (-not $DatabaseUserAccessEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($UserAccess) {
				$($DatabaseUserAccessEnum::Single) {'Single User'}
				$($DatabaseUserAccessEnum::Restricted) {'Restricted User'}
				$($DatabaseUserAccessEnum::Multiple) {'Multi User'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name DatabaseUserAccessEnum
}

function Get-EndpointTypeValue($EndpointType) {
	$EndpointTypeEnum = 'Microsoft.SqlServer.Management.Smo.EndpointType' -as [Type]

	if (-not $EndpointTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($EndpointType) {
				$($EndpointTypeEnum::Soap) {'SOAP'}
				$($EndpointTypeEnum::TSql) {'TSQL'}
				$($EndpointTypeEnum::ServiceBroker) {'Service Broker'}
				$($EndpointTypeEnum::DatabaseMirroring) {'Database Mirroring'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name EndpointTypeEnum
}

function Get-EndpointStateValue($EndpointState) {
	$EndpointStateEnum = 'Microsoft.SqlServer.Management.Smo.EndpointState' -as [Type]

	if (-not $EndpointStateEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($EndpointState) {
				$($EndpointStateEnum::Started) {'Started'}
				$($EndpointStateEnum::Stopped) {'Stopped'}
				$($EndpointStateEnum::Disabled) {'Disabled'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name EndpointStateEnum
}

function Get-EventSeverityLevelValue([int]$EventSeverityLevel) {
	Write-Output $(
		switch ($EventSeverityLevel) {
			1 { '001 - Miscellaneous System Information' }
			2 { '002 - Reserved' }
			3 { '003 - Reserved' }
			4 { '004 - Reserved' }
			5 { '005 - Reserved' }
			6 { '006 - Reserved' }
			7 { '007 - Notification: Status Information' }
			8 { '008 - Notification: User Intervention Required' }
			9 { '009 - User Defined' }
			10 { '010 - Information' }
			11 { '011 - Specified Database Object Not Found' }
			12 { '012 - Unused' }
			13 { '013 - User Transaction Syntax Error' }
			14 { '014 - Insufficient Permission' }
			15 { '015 - Syntax Error in SQL Statements' }
			16 { '016 - Miscellaneous User Error' }
			17 { '017 - Insufficient Resources' }
			18 { '018 - Non-Fatal Internal Error' }
			19 { '019 - Fatal Error in Resource' }
			20 { '020 - Fatal Error in Current Process' }
			21 { '021 - Fatal Error in Database Processes' }
			22 { '022 - Fatal Error: Table Integrity Suspect' }
			23 { '023 - Fatal Error: Database Integrity Suspect' }
			24 { '024 - Fatal Error: Hardware Error' }
			25 { '025 - Fatal Error' }
			default { '{0:000} - Unknown' -f $_ }
		}
	)
}

function Get-ExecutionContextValue($ExecutionContext) {

	$ExecutionContextEnum = 'Microsoft.SqlServer.Management.Smo.ExecutionContext' -as [Type]

	if (-not $ExecutionContextEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($ExecutionContext) {
				$($ExecutionContextEnum::Caller) {'Caller'}
				$($ExecutionContextEnum::Owner) {'Owner'}
				$($ExecutionContextEnum::ExecuteAsUser) {'Execute As User'}
				$($ExecutionContextEnum::Self) {'Self'}

				# Added the following to support bypassing SMO when retrieving database objects
				$($ExecutionContextEnum::Caller).value__ {'Caller'}
				$($ExecutionContextEnum::Owner).value__ {'Owner'}
				$($ExecutionContextEnum::ExecuteAsUser).value__ {'Execute As User'}
				$($ExecutionContextEnum::Self).value__ {'Self'}

				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name ExecutionContextEnum
}

function Get-ForeignKeyActionValue($ForeignKeyAction) {

	$ForeignKeyActionEnum = 'Microsoft.SqlServer.Management.Smo.ForeignKeyAction' -as [Type]

	if (-not $ForeignKeyActionEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($ForeignKeyAction) {
				$($ForeignKeyActionEnum::NoAction) {'No Action'}
				$($ForeignKeyActionEnum::Cascade) {'Cascade'}
				$($ForeignKeyActionEnum::SetNull) {'Set Null'}
				$($ForeignKeyActionEnum::SetDefault) {'Set Default'}

				# Added the following to support bypassing SMO when retrieving database objects
				$($ForeignKeyActionEnum::NoAction).value__ {'No Action'}
				$($ForeignKeyActionEnum::Cascade).value__ {'Cascade'}
				$($ForeignKeyActionEnum::SetNull).value__ {'Set Null'}
				$($ForeignKeyActionEnum::SetDefault).value__ {'Set Default'}

				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name ForeignKeyActionEnum
}

function Get-FullTextCatalogUpgradeOptionValue($FullTextCatalogUpgradeOption) {

	# Added to SMO 2008 (v10), will be $null if using SMO 2005 (v9)
	$FullTextCatalogUpgradeOptionEnum = 'Microsoft.SqlServer.Management.Smo.FullTextCatalogUpgradeOption' -as [Type]

	if (-not $FullTextCatalogUpgradeOptionEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($FullTextCatalogUpgradeOption) {
				$($FullTextCatalogUpgradeOptionEnum::AlwaysRebuild) {'Rebuild'}
				$($FullTextCatalogUpgradeOptionEnum::AlwaysReset) {'Reset'}
				$($FullTextCatalogUpgradeOptionEnum::ImportWithRebuild) {'Import'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name FullTextCatalogUpgradeOptionEnum 
}

function Get-FullTextLanguageValue($Language) {
	Write-Output $(
		switch ($Language) {
			0 { 'Neutral'}
			1025 { 'Arabic'}
			1026 { 'Bulgarian'}
			1027 { 'Catalan'}
			1028 { 'Traditional Chinese'}
			1029 { 'Czech'}
			1030 { 'Danish'}
			1031 { 'German'}
			1032 { 'Greek'}
			1033 { 'English'}
			1036 { 'French'}
			1037 { 'Hebrew'}
			1039 { 'Icelandic'}
			1040 { 'Italian'}
			1041 { 'Japanese'}
			1042 { 'Korean'}
			1043 { 'Dutch'}
			1044 { 'Bokmål'}
			1045 { 'Polish'}
			1046 { 'Brazilian'}
			1048 { 'Romanian'}
			1049 { 'Russian'}
			1050 { 'Croatian'}
			1051 { 'Slovak'}
			1053 { 'Swedish'}
			1054 { 'Thai'}
			1055 { 'Turkish'}
			1056 { 'Urdu'}
			1057 { 'Indonesian'}
			1058 { 'Ukrainian'}
			1060 { 'Slovenian'}
			1062 { 'Latvian'}
			1063 { 'Lithuanian'}
			1066 { 'Vietnamese'}
			1081 { 'Hindi'}
			1086 { 'Malay - Malaysia'}
			1093 { 'Bengali (India)'}
			1094 { 'Punjabi'}
			1095 { 'Gujarati'}
			1097 { 'Tamil'}
			1098 { 'Telugu'}
			1099 { 'Kannada'}
			1100 { 'Malayalam'}
			1102 { 'Marathi'}
			2052 { 'Simplified Chinese'}
			2057 { 'British English'}
			2070 { 'Portuguese'}
			2074 { 'Serbian (Latin)'}
			3076 { 'Chinese (Hong Kong SAR, PRC)'}
			3082 { 'Spanish'}
			3098 { 'Serbian (Cyrillic)'}
			4100 { 'Chinese (Singapore)'}
			5124 { 'Chinese (Macau SAR)'}
			$null {$null} 
			default {'Unknown'} 
		}
	)
}

function Get-HadrManagerStatusValue($HadrManagerStatus) {

	# Added to SMO 2012 (v11), will be $null otherwise
	$HadrManagerStatusEnum = 'Microsoft.SqlServer.Management.Smo.HadrManagerStatus' -as [Type]

	if (-not $HadrManagerStatusEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($HadrManagerStatus) {
				$($HadrManagerStatusEnum::PendingCommunication) {'Pending Communication'}
				$($HadrManagerStatusEnum::Running) {'Running'}
				$($HadrManagerStatusEnum::Failed) {'Failed'}
				$null { $null }
				default { $_.ToString() }
			}
		)
	}

	Remove-Variable -Name HadrManagerStatusEnum 
}

function Get-ImplementationTypeValue($ImplementationType) {

	# Added to SMO 2008 (v10), will be $null if using SMO 2005 (v9)
	$ImplementationTypeEnum = 'Microsoft.SqlServer.Management.Smo.ImplementationType' -as [Type]

	if (-not $ImplementationTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($ImplementationType) {
				$($ImplementationTypeEnum::TransactSql) {'T-SQL'}
				$($ImplementationTypeEnum::SqlClr) {'CLR'}

				# Added the following to support bypassing SMO when retrieving database objects
				$($ImplementationTypeEnum::TransactSql).value__ {'T-SQL'}
				$($ImplementationTypeEnum::SqlClr).value__ {'CLR'}

				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name ImplementationTypeEnum 
}

function Get-IndexPopulationStatusValue($IndexPopulationStatus) {

	$IndexPopulationStatusEnum = 'Microsoft.SqlServer.Management.Smo.IndexPopulationStatus' -as [Type]

	if (-not $IndexPopulationStatusEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($IndexPopulationStatus) {
				$($IndexPopulationStatusEnum::None) {'None'}
				$($IndexPopulationStatusEnum::Full) {'Full'}
				$($IndexPopulationStatusEnum::Incremental) {'Incremental'}
				$($IndexPopulationStatusEnum::Manual) {'Manual'}
				$($IndexPopulationStatusEnum::Background) {'Background'}
				$($IndexPopulationStatusEnum::PausedOrThrottled) {'Paused Or Throttled'}

				# Added the following to support bypassing SMO when retrieving database objects
				$($IndexPopulationStatusEnum::None).value__ {'None'}
				$($IndexPopulationStatusEnum::Full).value__ {'Full'}
				$($IndexPopulationStatusEnum::Incremental).value__ {'Incremental'}
				$($IndexPopulationStatusEnum::Manual).value__ {'Manual'}
				$($IndexPopulationStatusEnum::Background).value__ {'Background'}
				$($IndexPopulationStatusEnum::PausedOrThrottled).value__ {'Paused Or Throttled'}

				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name IndexPopulationStatusEnum
}

function Get-IndexKeyTypeValue($IndexKeyType) {

	$IndexKeyTypeEnum = 'Microsoft.SqlServer.Management.Smo.IndexKeyType' -as [Type]

	if (-not $IndexKeyTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($IndexKeyType) {
				$($IndexKeyTypeEnum::None) {'None'}
				$($IndexKeyTypeEnum::DriPrimaryKey) {'Primary Key'}
				$($IndexKeyTypeEnum::DriUniqueKey) {'Unique Constraint'}

				# Added the following to support bypassing SMO when retrieving database objects
				$($IndexKeyTypeEnum::None).value__ {'None'}
				$($IndexKeyTypeEnum::DriPrimaryKey).value__ {'Primary Key'}
				$($IndexKeyTypeEnum::DriUniqueKey).value__ {'Unique Constraint'}

				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name IndexKeyTypeEnum
}

function Get-IndexTypeValue($IndexType) {

	$IndexTypeEnum = 'Microsoft.SqlServer.Management.Smo.IndexType' -as [Type]

	if (-not $IndexTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($IndexType) {
				$($IndexTypeEnum::ClusteredIndex) {'Clustered'}
				$($IndexTypeEnum::NonClusteredIndex) {'Nonclustered'}
				$($IndexTypeEnum::PrimaryXmlIndex) {'Primary XML'}
				$($IndexTypeEnum::SecondaryXmlIndex) {'Secondary XML'}
				$($IndexTypeEnum::SpatialIndex) {'Spatial'}
				$($IndexTypeEnum::NonClusteredColumnStoreIndex) {'Nonclustered columnstore'}
				$($IndexTypeEnum::SelectiveXmlIndex) {'Selective XML'}
				$($IndexTypeEnum::SecondarySelectiveXmlIndex) {'Secondary Selective XML'}

				# Added the following to support bypassing SMO when retrieving database objects
				$($IndexTypeEnum::ClusteredIndex).value__ {'Clustered'}
				$($IndexTypeEnum::NonClusteredIndex).value__ {'Nonclustered'}
				$($IndexTypeEnum::PrimaryXmlIndex).value__ {'Primary XML'}
				$($IndexTypeEnum::SecondaryXmlIndex).value__ {'Secondary XML'}
				$($IndexTypeEnum::SpatialIndex).value__ {'Spatial'}
				$($IndexTypeEnum::NonClusteredColumnStoreIndex).value__ {'Nonclustered columnstore'}
				$($IndexTypeEnum::SelectiveXmlIndex).value__ {'Selective XML'}
				$($IndexTypeEnum::SecondarySelectiveXmlIndex).value__ {'Secondary Selective XML'} 

				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name IndexTypeEnum
}

function Get-JobServerTypeValue($JobServerType) {
	$JobServerTypeEnum = 'Microsoft.SqlServer.Management.Smo.Agent.JobServerType' -as [Type]

	if (-not $JobServerTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($JobServerType) {
				$($JobServerTypeEnum::Standalone) {'Standalone'}
				$($JobServerTypeEnum::Tsx) {'Target Server'}
				$($JobServerTypeEnum::Msx) {'Master Server'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name JobServerTypeEnum
}


function Get-LanguageValue($Language) {
	Write-Output $(
		switch ($Language) {
			0 { 'English'}
			1 { 'German'}
			2 { 'French'}
			3 { 'Japanese'}
			4 { 'Danish'}
			5 { 'Spanish'}
			6 { 'Italian'}
			7 { 'Dutch'}
			8 { 'Norwegian'}
			9 { 'Portuguese'}
			10 { 'Finnish'}
			11 { 'Swedish'}
			12 { 'Czech'}
			13 { 'Hungarian'}
			14 { 'Polish'}
			15 { 'Romanian'}
			16 { 'Croatian'}
			17 { 'Slovak'}
			18 { 'Slovenian'}
			19 { 'Greek'}
			20 { 'Bulgarian'}
			21 { 'Russian'}
			22 { 'Turkish'}
			23 { 'British English'}
			24 { 'Estonian'}
			25 { 'Latvian'}
			26 { 'Lithuanian'}
			27 { 'Brazilian'}
			28 { 'Traditional Chinese'}
			29 { 'Korean'}
			30 { 'Simplified Chinese'}
			31 { 'Arabic'}
			32 { 'Thai'}
			33 { 'Bokmål'} 
			$null {$null}
			default {'Unknown'} 
		}
	)
}


function Get-LoginTypeValue($LoginType) {
	$LoginTypeEnum = 'Microsoft.SqlServer.Management.Smo.LoginType' -as [Type]

	if (-not $LoginTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($LoginType) {
				$($LoginTypeEnum::WindowsUser) {'Windows User'}
				$($LoginTypeEnum::WindowsGroup) {'Windows Group'}
				$($LoginTypeEnum::SqlLogin) {'SQL Login'}
				$($LoginTypeEnum::Certificate) {'Certificate'}
				$($LoginTypeEnum::AsymmetricKey) {'Asymmetric Key'}
				$null { $null }
				default { $_.ToString() }
			}
		)
	}

	Remove-Variable -Name LoginTypeEnum 
}

function Get-MirroringSafetyLevelValue($MirroringSafetyLevel) {
	$MirroringSafetyLevelEnum = 'Microsoft.SqlServer.Management.Smo.MirroringSafetyLevel' -as [Type]

	if (-not $MirroringSafetyLevelEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($MirroringSafetyLevel) {
				$($MirroringSafetyLevelEnum::None) {'None'}
				$($MirroringSafetyLevelEnum::Unknown) {'Unknown'}
				$($MirroringSafetyLevelEnum::Off) {'Off'}
				$($MirroringSafetyLevelEnum::Full) {'Full'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name MirroringSafetyLevelEnum
}

function Get-MirroringStatusValue($MirroringStatus) {
	$MirroringStatusEnum = 'Microsoft.SqlServer.Management.Smo.MirroringStatus' -as [Type]

	if (-not $MirroringStatusEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($MirroringStatus) {
				$($MirroringStatusEnum::None) {'None'}
				$($MirroringStatusEnum::Suspended) {'Suspended'}
				$($MirroringStatusEnum::Disconnected) {'Disconnected'}
				$($MirroringStatusEnum::Synchronizing) {'Synchronizing'}
				$($MirroringStatusEnum::PendingFailover) {'Pending Failover'}
				$($MirroringStatusEnum::Synchronized) {'Synchronized'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name MirroringStatusEnum
}

function Get-MirroringWitnessStatusValue($MirroringWitnessStatus) {
	$MirroringWitnessStatusEnum = 'Microsoft.SqlServer.Management.Smo.MirroringWitnessStatus' -as [Type]

	if (-not $MirroringWitnessStatusEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($MirroringWitnessStatus) {
				$($MirroringWitnessStatusEnum::None) {'None'}
				$($MirroringWitnessStatusEnum::Unknown) {'Unknown'}
				$($MirroringWitnessStatusEnum::Connected) {'Connected'}
				$($MirroringWitnessStatusEnum::Disconnected) {'Disconnected'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name MirroringWitnessStatusEnum
}

function Get-OnFailureActionValue($OnFailureAction) {
	$OnFailureActionEnum = 'Microsoft.SqlServer.Management.Smo.OnFailureAction' -as [Type]

	if (-not $OnFailureActionEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($OnFailureAction) {
				$($OnFailureActionEnum::Continue) {'Continue'}
				$($OnFailureActionEnum::Shutdown) {'Shut Down Server'}
				$($OnFailureActionEnum::FailOperation) {'Fail Operation'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name OnFailureActionEnum 
}

function Get-ObjectClassValue($ObjectClass) {

	$ObjectClassEnum = 'Microsoft.SqlServer.Management.Smo.ObjectClass' -as [Type]

	if (-not $ObjectClassEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($ObjectClass) {
				$($ObjectClassEnum::Database) {'Database'}
				$($ObjectClassEnum::ObjectOrColumn) {'Object Or Column'}
				$($ObjectClassEnum::Schema) {'Schema'}
				$($ObjectClassEnum::User) {'User'}
				$($ObjectClassEnum::DatabaseRole) {'Database Role'}
				$($ObjectClassEnum::ApplicationRole) {'Application Role'}
				$($ObjectClassEnum::SqlAssembly) {'SQL Assembly'}
				$($ObjectClassEnum::UserDefinedType) {'User Defined Type'}
				$($ObjectClassEnum::SecurityExpression) {'Security Expression'}
				$($ObjectClassEnum::XmlNamespace) {'XML Namespace'}
				$($ObjectClassEnum::MessageType) {'Message Type'}
				$($ObjectClassEnum::ServiceContract) {'Service Contract'}
				$($ObjectClassEnum::Service) {'Service'}
				$($ObjectClassEnum::RemoteServiceBinding) {'Remote Service Binding'}
				$($ObjectClassEnum::ServiceRoute) {'Service Route'}
				$($ObjectClassEnum::FullTextCatalog) {'Full-Text Catalog'}
				$($ObjectClassEnum::SymmetricKey) {'Symmetric Key'}
				$($ObjectClassEnum::Server) {'Server'}
				$($ObjectClassEnum::Login) {'Login'}
				$($ObjectClassEnum::ServerPrincipal) {'Server Principal'}
				$($ObjectClassEnum::ServerRole) {'Server Role'}
				$($ObjectClassEnum::Endpoint) {'Endpoint'}
				$($ObjectClassEnum::Certificate) {'Certificate'}
				$($ObjectClassEnum::FullTextStopList) {'Full-Text Stop List'}
				$($ObjectClassEnum::AsymmetricKey) {'Asymmetric Key'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name ObjectClassEnum
}

function Get-NotificationMethodValue($NotificationMethod) {
	# http://msdn.microsoft.com/en-us/library/microsoft.sqlserver.management.smo.agent.notifymethods
	# This is a cheat until I can figure out how to deal with FlagsAttribute in PowerShell

	if ($NotificationMethod) {
		Write-Output $NotificationMethod.ToString()
	} else {
		Write-Output $null #Write-Output [String]$null
	}
}

function Get-PageVerifyValue($PageVerify) {
	$PageVerifyEnum = 'Microsoft.SqlServer.Management.Smo.PageVerify' -as [Type]

	if (-not $PageVerifyEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($PageVerify) {
				$($PageVerifyEnum::None) {'NONE'}
				$($PageVerifyEnum::TornPageDetection) {'TORN_PAGE_DETECTION'}
				$($PageVerifyEnum::Checksum) {'CHECKSUM'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name PageVerifyEnum 
}

function Get-PermissionStateValue($PermissionState) {

	$PermissionStateEnum = 'Microsoft.SqlServer.Management.Smo.PermissionState' -as [Type]

	if (-not $PermissionStateEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($PermissionState) {
				$($PermissionStateEnum::Deny) {'Deny'}
				$($PermissionStateEnum::Revoke) {'Revoke'}
				$($PermissionStateEnum::Grant) {'Grant'}
				$($PermissionStateEnum::GrantWithGrant) {'Grant With Grant'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name PermissionStateEnum
}

function Get-PrincipalTypeValue($PrincipalType) {

	$PrincipalTypeEnum = 'Microsoft.SqlServer.Management.Smo.PrincipalType' -as [Type]

	if (-not $PrincipalTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($PrincipalType) {
				$($PrincipalTypeEnum::None) {'None'}
				$($PrincipalTypeEnum::Login) {'Login'}
				$($PrincipalTypeEnum::ServerRole) {'Server Role'}
				$($PrincipalTypeEnum::User) {'User'}
				$($PrincipalTypeEnum::DatabaseRole) {'Database Role'}
				$($PrincipalTypeEnum::ApplicationRole) {'Application Role'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name PrincipalTypeEnum
}

function Get-PrivateKeyEncryptionTypeValue($PrivateKeyEncryptionType) {
	$PrivateKeyEncryptionTypeEnum = 'Microsoft.SqlServer.Management.Smo.PrivateKeyEncryptionType' -as [Type]

	if (-not $PrivateKeyEncryptionTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($PrivateKeyEncryptionType) {
				$($PrivateKeyEncryptionTypeEnum::NoKey) {'No Key'}
				$($PrivateKeyEncryptionTypeEnum::MasterKey) {'Master Key'}
				$($PrivateKeyEncryptionTypeEnum::Password) {'User Password'}
				$($PrivateKeyEncryptionTypeEnum::Provider) {'Encryption Provider'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name PrivateKeyEncryptionTypeEnum 
}

function Get-ProtocolTypeValue($ProtocolType) {
	$ProtocolTypeEnum = 'Microsoft.SqlServer.Management.Smo.ProtocolType' -as [Type]

	if (-not $ProtocolTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($ProtocolType) {
				$($ProtocolTypeEnum::Http) {'HTTP'}
				$($ProtocolTypeEnum::Tcp) {'TCP/IP'}
				$($ProtocolTypeEnum::NamedPipes) {'Named pipes'}
				$($ProtocolTypeEnum::SharedMemory) {'Shared memory'}
				$($ProtocolTypeEnum::Via) {'VIA'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name ProtocolTypeEnum
}

function Get-RangeTypeValue($RangeType) {

	$RangeTypeEnum = 'Microsoft.SqlServer.Management.Smo.RangeType' -as [Type]

	if (-not $RangeTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($RangeType) {
				$($RangeTypeEnum::None) {'None'}
				$($RangeTypeEnum::Left) {'Left'}
				$($RangeTypeEnum::Right) {'Right'}

				# Added the following to support bypassing SMO when retrieving database objects
				$($RangeTypeEnum::None).value__ {'None'}
				$($RangeTypeEnum::Left).value__ {'Left'}
				$($RangeTypeEnum::Right).value__ {'Right'}

				$null { $null }
				default { $_.ToString() }
			} 


		)
	}

	Remove-Variable -Name RangeTypeEnum
}

function Get-ReplicationOptionsValue($ReplicationOptions) {
	if ($ReplicationOptions) {
		Write-Output $ReplicationOptions.ToString()
	} else {
		Write-Output $null #Write-Output [String]$null
	}
}

function Get-SecondaryXmlIndexTypeValue($SecondaryXmlIndexType) {

	$SecondaryXmlIndexTypeEnum = 'Microsoft.SqlServer.Management.Smo.SecondaryXmlIndexType' -as [Type]

	if (-not $SecondaryXmlIndexTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($SecondaryXmlIndexType) {
				$($SecondaryXmlIndexTypeEnum::None) {'None'}
				$($SecondaryXmlIndexTypeEnum::Path) {'Path'}
				$($SecondaryXmlIndexTypeEnum::Value) {'Value'}
				$($SecondaryXmlIndexTypeEnum::Property) {'Property'}

				# Added the following to support bypassing SMO when retrieving database objects
				$($SecondaryXmlIndexTypeEnum::None).value__ {'None'}
				$($SecondaryXmlIndexTypeEnum::Path).value__ {'Path'}
				$($SecondaryXmlIndexTypeEnum::Value).value__ {'Value'}
				$($SecondaryXmlIndexTypeEnum::Property).value__ {'Property'}

				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name SecondaryXmlIndexTypeEnum
}

function Get-SequenceCacheTypeValue($SequenceCacheType) {

	$SequenceCacheTypeEnum = 'Microsoft.SqlServer.Management.Smo.SequenceCacheType' -as [Type]

	if (-not $SequenceCacheTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($SequenceCacheType) {
				$($SequenceCacheTypeEnum::DefaultCache) {'Default Size'}
				$($SequenceCacheTypeEnum::NoCache) {'No Cache'}
				$($SequenceCacheTypeEnum::CacheWithSize) {'Cache Size'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name SequenceCacheTypeEnum
}

function Get-SnapshotIsolationStateValue($SnapshotIsolationState) {
	$SnapshotIsolationStateEnum = 'Microsoft.SqlServer.Management.Smo.SnapshotIsolationState' -as [Type]

	if (-not $SnapshotIsolationStateEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($SnapshotIsolationState) {
				$($SnapshotIsolationStateEnum::Disabled) {'Disabled'}
				$($SnapshotIsolationStateEnum::Enabled) {'Enabled'}
				$($SnapshotIsolationStateEnum::PendingOff) {'Pending Off'}
				$($SnapshotIsolationStateEnum::PendingOn) {'Pending On'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name SnapshotIsolationStateEnum 
}

function Get-SpatialGeoLevelSizeValue($SpatialGeoLevelSize) {

	$SpatialGeoLevelSizeEnum = 'Microsoft.SqlServer.Management.Smo.SpatialGeoLevelSize' -as [Type]

	if (-not $SpatialGeoLevelSizeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($SpatialGeoLevelSize) {
				$($SpatialGeoLevelSizeEnum::None) {'None'}
				$($SpatialGeoLevelSizeEnum::Low) {'Low'}
				$($SpatialGeoLevelSizeEnum::Medium) {'Medium'}
				$($SpatialGeoLevelSizeEnum::High) {'High'}

				# Added the following to support bypassing SMO when retrieving database objects
				$($SpatialGeoLevelSizeEnum::None).value__ {'None'}
				$($SpatialGeoLevelSizeEnum::Low).value__ {'Low'}
				$($SpatialGeoLevelSizeEnum::Medium).value__ {'Medium'}
				$($SpatialGeoLevelSizeEnum::High).value__ {'High'}

				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name SpatialGeoLevelSizeEnum
}

function Get-SpatialIndexTypeValue($SpatialIndexType) {

	$SpatialIndexTypeEnum = 'Microsoft.SqlServer.Management.Smo.SpatialIndexType' -as [Type]

	if (-not $SpatialIndexTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($SpatialIndexType) {
				$($SpatialIndexTypeEnum::None) {'None'}
				$($SpatialIndexTypeEnum::GeometryGrid) {'Geometry'}
				$($SpatialIndexTypeEnum::GeographyGrid) {'Geography'}

				# Added the following to support bypassing SMO when retrieving database objects
				$($SpatialIndexTypeEnum::None).value__ {'None'}
				$($SpatialIndexTypeEnum::GeometryGrid).value__ {'Geometry'}
				$($SpatialIndexTypeEnum::GeographyGrid).value__ {'Geography'}

				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name SpatialIndexTypeEnum
}

function Get-SqlDataTypeValue($SqlDataType) {

	$SqlDataTypeEnum = 'Microsoft.SqlServer.Management.Smo.SqlDataType' -as [Type]

	if (-not $SqlDataTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($SqlDataType) {
				$($SqlDataTypeEnum::None) {'None'}
				$($SqlDataTypeEnum::BigInt) {'BigInt'}
				$($SqlDataTypeEnum::Binary) {'Binary'}
				$($SqlDataTypeEnum::Bit) {'Bit'}
				$($SqlDataTypeEnum::Char) {'Char'}
				$($SqlDataTypeEnum::DateTime) {'DateTime'}
				$($SqlDataTypeEnum::Decimal) {'Decimal'}
				$($SqlDataTypeEnum::Float) {'Float'}
				$($SqlDataTypeEnum::Image) {'Image'}
				$($SqlDataTypeEnum::Int) {'Int'}
				$($SqlDataTypeEnum::Money) {'Money'}
				$($SqlDataTypeEnum::NChar) {'NChar'}
				$($SqlDataTypeEnum::NText) {'NText'}
				$($SqlDataTypeEnum::NVarChar) {'NVarChar'}
				$($SqlDataTypeEnum::NVarCharMax) {'NVarCharMax'}
				$($SqlDataTypeEnum::Real) {'Real'}
				$($SqlDataTypeEnum::SmallDateTime) {'SmallDateTime'}
				$($SqlDataTypeEnum::SmallInt) {'SmallInt'}
				$($SqlDataTypeEnum::SmallMoney) {'SmallMoney'}
				$($SqlDataTypeEnum::SmallMoney) {'SmallMoney'}
				$($SqlDataTypeEnum::Text) {'Text'}
				$($SqlDataTypeEnum::Timestamp) {'Timestamp'}
				$($SqlDataTypeEnum::TinyInt) {'TinyInt'}
				$($SqlDataTypeEnum::UniqueIdentifier) {'UniqueIdentifier'}
				$($SqlDataTypeEnum::UserDefinedDataType) {'User Defined Data Type'}
				$($SqlDataTypeEnum::UserDefinedType) {'User Defined Type'}
				$($SqlDataTypeEnum::VarBinary) {'VarBinary'}
				$($SqlDataTypeEnum::VarBinaryMax) {'VarBinaryMax'}
				$($SqlDataTypeEnum::VarChar) {'VarChar'}
				$($SqlDataTypeEnum::VarCharMax) {'VarCharMax'}
				$($SqlDataTypeEnum::Variant) {'Variant'}
				$($SqlDataTypeEnum::Xml) {'Xml'}
				$($SqlDataTypeEnum::SysName) {'SysName'}
				$($SqlDataTypeEnum::Numeric) {'Numeric'}
				$($SqlDataTypeEnum::Date) {'Date'}
				$($SqlDataTypeEnum::Time) {'Time'}
				$($SqlDataTypeEnum::DateTimeOffset) {'DateTimeOffset'}
				$($SqlDataTypeEnum::DateTime2) {'DateTime2'}
				$($SqlDataTypeEnum::UserDefinedTableType) {'User Defined Table Type'}
				$($SqlDataTypeEnum::HierarchyId) {'HierarchyId'}
				$($SqlDataTypeEnum::Geometry) {'Geometry'}
				$($SqlDataTypeEnum::Geography) {'Geography'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name SqlDataTypeEnum
}

function Get-StopListOptionValue($StopListOption) {

	$StopListOptionEnum = 'Microsoft.SqlServer.Management.Smo.StopListOption' -as [Type]

	if (-not $StopListOptionEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($StopListOption) {
				$($StopListOptionEnum::Off) {'Off'}
				$($StopListOptionEnum::System) {'System'}
				$($StopListOptionEnum::Name) {'Name'}

				# Added the following to support bypassing SMO when retrieving database objects
				$($StopListOptionEnum::Off).value__ {'Off'}
				$($StopListOptionEnum::System).value__ {'System'}
				$($StopListOptionEnum::Name).value__ {'Name'}

				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name StopListOptionEnum
}

function Get-ServerDdlTriggerExecutionContextValue($ServerDdlTriggerExecutionContext) {
	$ServerDdlTriggerExecutionContextEnum = 'Microsoft.SqlServer.Management.Smo.ServerDdlTriggerExecutionContext' -as [Type]

	if (-not $ServerDdlTriggerExecutionContextEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($ServerDdlTriggerExecutionContext) {
				$($ServerDdlTriggerExecutionContextEnum::Caller) {'Execute As Caller'}
				$($ServerDdlTriggerExecutionContextEnum::ExecuteAsLogin) {'Execute As Login'}
				$($ServerDdlTriggerExecutionContextEnum::Self) {'Execute As Self'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name ServerDdlTriggerExecutionContextEnum 
}

function Get-SymmetricKeyEncryptionAlgorithmValue($SymmetricKeyEncryptionAlgorithm) {
	$SymmetricKeyEncryptionAlgorithmEnum = 'Microsoft.SqlServer.Management.Smo.SymmetricKeyEncryptionAlgorithm' -as [Type]

	if (-not $SymmetricKeyEncryptionAlgorithmEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($SymmetricKeyEncryptionAlgorithm) {
				$($SymmetricKeyEncryptionAlgorithmEnum::CryptographicProviderDefined) {'Cryptographic Provider'}
				$($SymmetricKeyEncryptionAlgorithmEnum::RC2) {'RC2'}
				$($SymmetricKeyEncryptionAlgorithmEnum::RC4) {'RC4'}
				$($SymmetricKeyEncryptionAlgorithmEnum::Des) {'DES'}
				$($SymmetricKeyEncryptionAlgorithmEnum::TripleDes) {'Triple DES'}
				$($SymmetricKeyEncryptionAlgorithmEnum::DesX) {'DESX'}
				$($SymmetricKeyEncryptionAlgorithmEnum::Aes128) {'AES 128'}
				$($SymmetricKeyEncryptionAlgorithmEnum::Aes192) {'AES 192'}
				$($SymmetricKeyEncryptionAlgorithmEnum::Aes256) {'AES 256'}
				$($SymmetricKeyEncryptionAlgorithmEnum::TripleDes3Key) {'Triple DES 3KEY'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name SymmetricKeyEncryptionAlgorithmEnum 
}

function Get-SynonymBaseTypeValue($SynonymBaseType) {

	$SynonymBaseTypeEnum = 'Microsoft.SqlServer.Management.Smo.SynonymBaseType' -as [Type]

	if (-not $SynonymBaseTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($SynonymBaseType) {
				$($SynonymBaseTypeEnum::None) {'None'}
				$($SynonymBaseTypeEnum::Table) {'Table'}
				$($SynonymBaseTypeEnum::View) {'View'}
				$($SynonymBaseTypeEnum::SqlStoredProcedure) {'Stored Procedure'}
				$($SynonymBaseTypeEnum::SqlScalarFunction) {'Scalar Function'}
				$($SynonymBaseTypeEnum::SqlTableValuedFunction) {'Table Valued Function'}
				$($SynonymBaseTypeEnum::SqlInlineTableValuedFunction) {'Inline Table Valued Function'}
				$($SynonymBaseTypeEnum::ExtendedStoredProcedure) {'Extended Stored Procedure'}
				$($SynonymBaseTypeEnum::ReplicationFilterProcedure) {'Replication Filter Procedure'}
				$($SynonymBaseTypeEnum::ClrStoredProcedure) {'CLR Stored Procedure'}
				$($SynonymBaseTypeEnum::ClrScalarFunction) {'CLR Scalar Function'}
				$($SynonymBaseTypeEnum::ClrTableValuedFunction) {'CLR Table Valued Function'}
				$($SynonymBaseTypeEnum::ClrAggregateFunction) {'CLR Aggregate Function'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name SynonymBaseTypeEnum
}

function Get-ServiceStartModeValue($ServiceStartMode) {

	# Added to SMO 2008 (v10), will be $null if using SMO 2005 (v9)
	$ServiceStartModeEnum = 'Microsoft.SqlServer.Management.Smo.ServiceStartMode' -as [Type]

	if (-not $ServiceStartModeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($ServiceStartMode) {
				$($ServiceStartModeEnum::Boot) {'Boot'}
				$($ServiceStartModeEnum::System) {'System'}
				$($ServiceStartModeEnum::Auto) {'Automatic'}
				$($ServiceStartModeEnum::Manual) {'Manual'}
				$($ServiceStartModeEnum::Disabled) {'Disabled'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name ServiceStartModeEnum 
}

function Get-UserDefinedFunctionTypeValue($UserDefinedFunctionType) {

	$UserDefinedFunctionTypeEnum = 'Microsoft.SqlServer.Management.Smo.UserDefinedFunctionType' -as [Type]

	if (-not $UserDefinedFunctionTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($UserDefinedFunctionType) {
				$($UserDefinedFunctionTypeEnum::Inline) {'Inline'}
				$($UserDefinedFunctionTypeEnum::Scalar) {'Scalar'}
				$($UserDefinedFunctionTypeEnum::Table) {'Table'}
				$($UserDefinedFunctionTypeEnum::Unknown) {'Unknown'}

				# Added the following to support bypassing SMO when retrieving database objects
				$($UserDefinedFunctionTypeEnum::Inline).value__ {'Inline'}
				$($UserDefinedFunctionTypeEnum::Scalar).value__ {'Scalar'}
				$($UserDefinedFunctionTypeEnum::Table).value__ {'Table'}
				$($UserDefinedFunctionTypeEnum::Unknown).value__ {'Unknown'}

				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name UserDefinedFunctionTypeEnum
}

function Get-UserDefinedTypeFormatValue($UserDefinedTypeFormat) {

	$UserDefinedTypeFormatEnum = 'Microsoft.SqlServer.Management.Smo.UserDefinedTypeFormat' -as [Type]

	if (-not $UserDefinedTypeFormatEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($UserDefinedTypeFormat) {
				$($UserDefinedTypeFormatEnum::Native) {'Native'}
				$($UserDefinedTypeFormatEnum::UserDefined) {'User Defined'}
				$($UserDefinedTypeFormatEnum::SerializedData) {'Serialized Data'}
				$($UserDefinedTypeFormatEnum::SerializedDataWithMetadata) {'Serialized Data With Metadata'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name UserDefinedTypeFormatEnum
}

function Get-UserTypeValue($UserType) {
	$UserTypeEnum = 'Microsoft.SqlServer.Management.Smo.UserType' -as [Type]

	if (-not $UserTypeEnum) {
		Write-Output $null #Write-Output [String]$null
	} else {
		Write-Output $(
			switch ($UserType) {
				$($UserTypeEnum::SqlLogin) {'SQL user with login'}
				$($UserTypeEnum::Certificate) {'User mapped to a certificate'}
				$($UserTypeEnum::AsymmetricKey) {'User mapped to an asymmetric key'}
				$($UserTypeEnum::NoLogin) {'SQL user without login'}
				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name UserTypeEnum
}

function Get-XmlDocumentConstraintValue($XmlDocumentConstraint) {

	$XmlDocumentConstraintEnum = 'Microsoft.SqlServer.Management.Smo.XmlDocumentConstraint' -as [Type]

	if (-not $XmlDocumentConstraintEnum) {
		Write-Output $null #Write-Output [String]$null
	} else { 
		Write-Output $(
			switch ($XmlDocumentConstraint) {
				$($XmlDocumentConstraintEnum::Default) {'Default'}
				$($XmlDocumentConstraintEnum::Content) {'Content'}
				$($XmlDocumentConstraintEnum::Document) {'Document'}

				# Added the following to support bypassing SMO when retrieving database objects
				$($XmlDocumentConstraintEnum::Default).value__ {'Default'}
				$($XmlDocumentConstraintEnum::Content).value__ {'Content'}
				$($XmlDocumentConstraintEnum::Document).value__ {'Document'}

				$null { $null }
				default { $_.ToString() }
			} 
		)
	}

	Remove-Variable -Name XmlDocumentConstraintEnum
}

function Get-SqlServerVersionName([int]$MajorVersion, [int]$MinorVersion) {
	if (($MajorVersion -eq 11) -and ($MinorVersion -eq 0)) { '2012' }
	elseif (($MajorVersion -eq 10) -and ($MinorVersion -eq 50)) { '2008 R2' }
	elseif (($MajorVersion -eq 10) -and ($MinorVersion -eq 0)) { '2008' }
	elseif (($MajorVersion -eq 9) -and ($MinorVersion -eq 0)) { '2005' }
	elseif (($MajorVersion -eq 8) -and ($MinorVersion -eq 0)) { '2000' }
	elseif (($MajorVersion -eq 7) -and ($MinorVersion -eq 0)) { '7.0' }
	elseif (($MajorVersion -eq 6) -and ($MinorVersion -eq 50)) { '6.5' }
	elseif (($MajorVersion -eq 6) -and ($MinorVersion -eq 0)) { '6.0' }
	else { 'unknown '}
}


function Write-SqlServerDatabaseEngineInformationLog {
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

function Get-SqlConnection {
	[CmdletBinding(DefaultParametersetName='WindowsAuthentication')]
	[OutputType([System.Data.SqlClient.SqlConnection])]
	param(
		[Parameter(Mandatory=$false, ParameterSetName = 'SQLAuthentication')]
		[Parameter(Mandatory=$false, ParameterSetName = 'WindowsAuthentication')]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$Instance = '(local)'
		,
		[Parameter(Mandatory=$false)]
		[System.Net.IPAddress]
		$IpAddress = $null
		,
		[Parameter(Mandatory=$false)]
		[Int]
		$Port = $null
		, 
		[Parameter(Mandatory=$false)]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$Database = 'master'
		,
		[Parameter(Mandatory=$true, ParameterSetName = 'SQLAuthentication')]
		[ValidateNotNull()]
		[System.String]
		$Username
		,
		[Parameter(Mandatory=$true, ParameterSetName = 'SQLAuthentication')]
		[ValidateNotNull()]
		[System.String]
		$Password
		,
		[Parameter(Mandatory=$false)]
		[System.String]
		$FailoverPartner = $null
		,
		[Parameter(Mandatory=$false)]
		[System.String]
		$ApplicationName = 'Windows PowerShell' # $MyInvocation.ScriptName	
	)
	try {

		# ConnectionStringBuilder docs: http://msdn.microsoft.com/en-us/library/system.data.sqlclient.sqlconnectionstringbuilder.aspx
		# http://www.connectionstrings.com/Articles/Show/all-sql-server-connection-string-keywords

		$SQLConnection = New-Object -TypeName System.Data.SqlClient.SqlConnection
		$SQLConnectionBuilder = New-Object -TypeName system.Data.SqlClient.SqlConnectionStringBuilder

		$SQLConnectionBuilder.psBase.DataSource = if ($IpAddress) {
			if ($Port) { "$($IpAddress.ToString()),$Port" } else { $IpAddress.ToString() }
		} else {
			if ($Port) { "$Instance,$Port" } else { $Instance }
		}
		$SQLConnectionBuilder.psBase.InitialCatalog = $Database

		if ($PSCmdlet.ParameterSetName -eq 'SQLAuthentication') {
			$SQLConnectionBuilder.psBase.IntegratedSecurity = $false
			$SQLConnectionBuilder.psBase.UserID = $Username
			$SQLConnectionBuilder.psBase.Password = $Password
		} else {
			$SQLConnectionBuilder.psBase.IntegratedSecurity = $true
		}

		$SQLConnectionBuilder.psBase.FailoverPartner = $FailoverPartner
		$SQLConnectionBuilder.psBase.ApplicationName = $ApplicationName

		$SQLConnection.ConnectionString = $SQLConnectionBuilder.ConnectionString

		Write-Output $SQLConnection

		Remove-Variable -Name SQLConnection, SQLConnectionBuilder

	}
	catch {
		Throw
	}
}

function Get-InstanceVersion {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Server]
		$Server
	)
	try {
		$Version = ($Server.Information.Version)

		Write-Output (
			New-Object -TypeName psobject -Property @{
				Build = $Version.Build
				Major = $Version.Major
				Minor = $Version.Minor
				Revision = $Version.Revision
			}
		)
	}
	catch {
		Throw
	}
	finally {
		Remove-Variable -Name Version
	}
}

function Get-CheckInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)]
		[Microsoft.SqlServer.Management.Smo.CheckCollection]
		$CheckCollection
	)
	Begin {
	}
	Process {
		try {
			$CheckCollection | ForEach-Object {
				Write-Output (
					New-Object -TypeName PSObject -Property @{
						CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
						DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
						#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
						ID = $_.ID # System.Int32 ID {get;}
						IsChecked = $_.IsChecked # System.Boolean IsChecked {get;set;}
						IsEnabled = $_.IsEnabled # System.Boolean IsEnabled {get;set;}
						IsFileTableDefined = $_.IsFileTableDefined # System.Boolean IsFileTableDefined {get;}
						IsSystemNamed = $_.IsSystemNamed # System.Boolean IsSystemNamed {get;}
						Name = $_.Name # System.String Name {get;set;}
						IsNotForReplication = $_.NotForReplication # System.Boolean NotForReplication {get;set;}
						#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlSmoObject Parent {get;set;}
						#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
						#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
						#Text = $_.Text # System.String Text {get;set;}

						Definition = if (-not [String]::IsNullOrEmpty($_.Text)) {
							$_.Text.Trim()
						} else {
							$null
						}

						#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
						#UserData = $_.UserData	# System.Object UserData {get;set;}
					} 
				) 
			}
		}
		catch {
			throw
		}
	}
	End {
	}
}

function Get-CheckInformation2 {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)]
		[AllowNull()]
		[System.Data.DataRow[]]
		$CheckCollection
	)
	Begin {
	}
	Process {
		try {
			$CheckCollection | Where-Object { $_.ID} | ForEach-Object {
				Write-Output (
					New-Object -TypeName PSObject -Property @{
						CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
						DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
						#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
						ID = $_.ID # System.Int32 ID {get;}
						IsChecked = $_.IsChecked # System.Boolean IsChecked {get;set;}
						IsEnabled = $_.IsEnabled # System.Boolean IsEnabled {get;set;}
						IsFileTableDefined = $_.IsFileTableDefined # System.Boolean IsFileTableDefined {get;}
						IsSystemNamed = $_.IsSystemNamed # System.Boolean IsSystemNamed {get;}
						Name = $_.Name # System.String Name {get;set;}
						IsNotForReplication = $_.NotForReplication # System.Boolean NotForReplication {get;set;}
						#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlSmoObject Parent {get;set;}
						#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
						#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
						#Text = $_.Text # System.String Text {get;set;}

						Definition = if (-not [String]::IsNullOrEmpty($_.Definition)) {
							$_.Definition.Trim()
						} else {
							$null
						}


						#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
						#UserData = $_.UserData	# System.Object UserData {get;set;}
					} 
				) 
			}
		}
		catch {
			throw
		}
	}
	End {
	}
}

function Get-ColumnInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)]
		[Microsoft.SqlServer.Management.Smo.ColumnCollection]
		$ColumnCollection
	)
	Begin {
	}
	Process {
		try {
			$ColumnCollection | ForEach-Object {
				Write-Output (
					New-Object -TypeName PSObject -Property @{
						General = New-Object -TypeName PSObject -Property @{
							Binding = New-Object -TypeName PSObject -Property @{
								DefaultBinding = $_.Default # System.String Default {get;set;}
								DefaultSchema = $_.DefaultSchema # System.String DefaultSchema {get;set;}
								Rule = $_.Rule # System.String Rule {get;set;}
								RuleSchema = $_.RuleSchema # System.String RuleSchema {get;set;}
							}
							Computed = New-Object -TypeName PSObject -Property @{
								IsComputed = $_.Computed # System.Boolean Computed {get;set;}
								ComputedText = if ($_.Computed) { $_.ComputedText } else { $null } # System.String ComputedText {get;set;}
							}
							General = New-Object -TypeName PSObject -Property @{
								AllowNulls = $_.Nullable # System.Boolean Nullable {get;set;}
								AnsiPaddingStatus = $_.AnsiPaddingStatus # System.Boolean AnsiPaddingStatus {get;}
								DataType = $_.DataType.Name # System.String Name {get;set;}
								Length = $_.DataType.MaximumLength # System.Int32 MaximumLength {get;set;}
								ID = $_.ID # System.Int32 ID {get;}
								Name = $_.Name # System.String Name {get;set;}
								NumericPrecision = $_.DataType.NumericPrecision # System.Int32 NumericPrecision {get;set;}
								NumericScale = $_.DataType.NumericScale # System.Int32 NumericScale {get;set;}
								InPrimaryKey = $_.InPrimaryKey # System.Boolean InPrimaryKey {get;}
								SystemType = [String](Get-SqlDataTypeValue -SqlDataType $_.DataType.SqlDataType) # Microsoft.SqlServer.Management.Smo.SqlDataType SqlDataType {get;set;}
							}
							Identity = New-Object -TypeName PSObject -Property @{
								IsIdentity = $_.Identity # System.Boolean Identity {get;set;}
								IdentityIncrement = $_.IdentityIncrement # System.Int64 IdentityIncrement {get;set;}
								IdentitySeed = $_.IdentitySeed # System.Int64 IdentitySeed {get;set;}
							}
							Miscellaneous = New-Object -TypeName PSObject -Property @{
								Collation = $_.Collation # System.String Collation {get;set;}
								IsFullTextIndexed = $_.IsFullTextIndexed # System.Boolean IsFullTextIndexed {get;}
								IsNotForReplication = $_.NotForReplication # System.Boolean NotForReplication {get;set;}
								StatisticalSemantics = $_.StatisticalSemantics # System.Int32 StatisticalSemantics {get;}

								# Do these belong here?
								IsDeterministic = $_.IsDeterministic # System.Boolean IsDeterministic {get;}
								IsFileStream = $_.IsFileStream # System.Boolean IsFileStream {get;set;}
								IsForeignKey = $_.IsForeignKey # System.Boolean IsForeignKey {get;}
								IsPersisted = $_.IsPersisted # System.Boolean IsPersisted {get;set;}
								IsPrecise = $_.IsPrecise # System.Boolean IsPrecise {get;}
								IsRowGuidCol = $_.RowGuidCol # System.Boolean RowGuidCol {get;set;}								
							}
							Sparse = New-Object -TypeName PSObject -Property @{
								IsColumnSet = $_.IsColumnSet # System.Boolean IsColumnSet {get;set;}
								IsSparse = $_.IsSparse # System.Boolean IsSparse {get;set;}
							}
							XML = New-Object -TypeName PSObject -Property @{
								XmlSchemaNameSpace = $null
								XmlSchemaNameSpaceSchema = $null
							}
						}
						<#
						DataType = if ($_.DataType) { 
							New-Object -TypeName PSObject -Property @{
								Schema = $_.DataType.Schema # System.String Schema {get;set;}
								XmlDocumentConstraint = [String](Get-XmlDocumentConstraintValue -XmlDocumentConstraint $_.DataType.XmlDocumentConstraint) # Microsoft.SqlServer.Management.Smo.XmlDocumentConstraint XmlDocumentConstraint {get;set;}
							} # Microsoft.SqlServer.Management.Smo.DataType DataType {get;set;}
						} else {
							New-Object -TypeName PSObject -Property @{
								Schema = $null
								XmlDocumentConstraint = $null
							}
						}
#>

						DefaultConstraint = New-Object -TypeName PSObject -Property @{
							CreateDate = $_.DefaultConstraint.CreateDate # System.DateTime CreateDate {get;}
							DateLastModified = $_.DefaultConstraint.DateLastModified # System.DateTime DateLastModified {get;}
							#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
							ID = $_.DefaultConstraint.ID # System.Int32 ID {get;}
							IsFileTableDefined = $_.DefaultConstraint.IsFileTableDefined # System.Boolean IsFileTableDefined {get;}
							IsSystemNamed = $_.DefaultConstraint.IsSystemNamed # System.Boolean IsSystemNamed {get;}
							Name = $_.DefaultConstraint.Name # System.String Name {get;set;}
							#Parent = $_.DefaultConstraint.Parent	# Microsoft.SqlServer.Management.Smo.Column Parent {get;}
							#Properties = $_.DefaultConstraint.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
							#State = $_.DefaultConstraint.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}

							Text = if (-not [String]::IsNullOrEmpty($_.DefaultConstraint.Text)) {
								$_.DefaultConstraint.Text.Trim() # System.String Text {get;set;}
							} else {
								$null
							}

							#Urn = $_.DefaultConstraint.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
							#UserData = $_.DefaultConstraint.UserData	# System.Object UserData {get;set;}
						} # Microsoft.SqlServer.Management.Smo.DefaultConstraint DefaultConstraint {get;}
						#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}

						#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlSmoObject Parent {get;set;}
						#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
						#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
						#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
						#UserData = $_.UserData	# System.Object UserData {get;set;}

					}
				) 
			}
		}
		catch {
			throw
		}
	}
	End {
	}
}

function Get-ColumnInformation2 {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)]
		[AllowNull()]
		[System.Data.DataRow[]]
		$ColumnCollection
		,
		[Parameter(Mandatory=$true)]
		[AllowNull()]
		[System.Data.DataRow[]]
		$DefaultConstraintCollection
	)
	Begin {
	}
	Process {
		try {
			$ColumnCollection | Where-Object { $_.ID } | ForEach-Object {
				$ColumnID = $_.ID

				Write-Output (
					New-Object -TypeName PSObject -Property @{
						General = New-Object -TypeName PSObject -Property @{
							Binding = New-Object -TypeName PSObject -Property @{
								DefaultBinding = $_.Default # System.String Default {get;set;}
								DefaultSchema = $_.DefaultSchema # System.String DefaultSchema {get;set;}
								Rule = $_.Rule # System.String Rule {get;set;}
								RuleSchema = $_.RuleSchema # System.String RuleSchema {get;set;}
							}
							Computed = New-Object -TypeName PSObject -Property @{
								IsComputed = $_.Computed # System.Boolean Computed {get;set;}
								ComputedText = $_.ComputedText # System.String ComputedText {get;set;}
							}
							General = New-Object -TypeName PSObject -Property @{
								AllowNulls = $_.Nullable # System.Boolean Nullable {get;set;}
								AnsiPaddingStatus = $_.AnsiPaddingStatus # System.Boolean AnsiPaddingStatus {get;}
								DataType = $_.DataType
								Length = $_.Length # System.Int32 MaximumLength {get;set;}
								ID = $_.ID # System.Int32 ID {get;}
								Name = $_.Name # System.String Name {get;set;}
								NumericPrecision = $_.NumericPrecision # System.Int32 NumericPrecision {get;set;}
								NumericScale = $_.NumericScale # System.Int32 NumericScale {get;set;}
								InPrimaryKey = $_.InPrimaryKey # System.Boolean InPrimaryKey {get;}
								SystemType = $_.SystemType
							}
							Identity = New-Object -TypeName PSObject -Property @{
								IsIdentity = $_.Identity # System.Boolean Identity {get;set;}
								IdentityIncrement = $_.IdentityIncrement # System.Int64 IdentityIncrement {get;set;}
								IdentitySeed = $_.IdentitySeed # System.Int64 IdentitySeed {get;set;}
							}
							Miscellaneous = New-Object -TypeName PSObject -Property @{
								Collation = $_.Collation # System.String Collation {get;set;}
								IsFullTextIndexed = $_.IsFullTextIndexed # System.Boolean IsFullTextIndexed {get;}
								IsNotForReplication = $_.NotForReplication # System.Boolean NotForReplication {get;set;}
								StatisticalSemantics = $_.StatisticalSemantics # System.Int32 StatisticalSemantics {get;}

								# Do these belong here?
								IsDeterministic = $_.IsDeterministic # System.Boolean IsDeterministic {get;}
								IsFileStream = $_.IsFileStream # System.Boolean IsFileStream {get;set;}
								IsForeignKey = $_.IsForeignKey # System.Boolean IsForeignKey {get;}
								IsPersisted = $_.IsPersisted # System.Boolean IsPersisted {get;set;}
								IsPrecise = $_.IsPrecise # System.Boolean IsPrecise {get;}
								IsRowGuidCol = $_.RowGuidCol # System.Boolean RowGuidCol {get;set;}
							}
							Sparse = New-Object -TypeName PSObject -Property @{
								IsColumnSet = $_.IsColumnSet # System.Boolean IsColumnSet {get;set;}
								IsSparse = $_.IsSparse # System.Boolean IsSparse {get;set;}
							}
							XML = New-Object -TypeName PSObject -Property @{
								XmlSchemaNameSpace = $null
								XmlSchemaNameSpaceSchema = $null
							}
						}
						<#
						DataType = New-Object -TypeName PSObject -Property @{
							Schema = $_.DataTypeSchema # System.String  DataTypeSchema {get;set;}
							XmlDocumentConstraint = [String](Get-XmlDocumentConstraintValue -XmlDocumentConstraint $_.XmlDocumentConstraint) # Microsoft.SqlServer.Management.Smo.XmlDocumentConstraint XmlDocumentConstraint {get;set;}
						} # Microsoft.SqlServer.Management.Smo.DataType DataType {get;set;}

#>
						DefaultConstraint = $DefaultConstraintCollection | Where-Object { $_.ColumnID -eq $ColumnID } | ForEach-Object {
							New-Object -TypeName PSObject -Property @{
								CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
								DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
								#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
								ID = $_.ID # System.Int32 ID {get;}
								IsFileTableDefined = $_.IsFileTableDefined # System.Boolean IsFileTableDefined {get;}
								IsSystemNamed = $_.IsSystemNamed # System.Boolean IsSystemNamed {get;}
								Name = $_.Name # System.String Name {get;set;}
								#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Column Parent {get;}
								#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
								#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}

								Text = if (-not [String]::IsNullOrEmpty($_.Text)) {
									$_.Text.Trim() # System.String Text {get;set;}
								} else {
									$null
								}

								#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
								#UserData = $_.UserData	# System.Object UserData {get;set;}
							} # Microsoft.SqlServer.Management.Smo.DefaultConstraint DefaultConstraint {get;}
						}

						#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
						#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlSmoObject Parent {get;set;}
						#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
						#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
						#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
						#UserData = $_.UserData	# System.Object UserData {get;set;}
					}
				) 
			}
		}
		catch {
			throw
		}
	}
	End {
	}
}

function Get-ExtendedPropertyInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection]
		$ExtendedPropertyCollection
	)
	Begin {
	}
	Process {
		try {
			$ExtendedPropertyCollection | ForEach-Object {
				Write-Output (
					New-Object -TypeName PSObject -Property @{
						Name = $_.Name # System.String Name {get;set;}
						#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlSmoObject Parent {get;set;}
						#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
						#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
						#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
						#UserData = $_.UserData	# System.Object UserData {get;set;}
						Value = $_.Value # System.Object Value {get;set;}
					}
				) 
			}
		}
		catch {
			throw
		}
	}
	End {
	}
}


function Get-FailoverClusterMemberList {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Server]
		$Server
	)
	Begin {
		$DbEngineType = [String](Get-DatabaseEngineTypeValue -DatabaseEngineType $Server.ServerType)
	}
	Process {
		try {
			if ($Server.Information.IsClustered) {
				if (
					$($Server.Information.Version).CompareTo($SQLServer2012) -ge 0 -and
					$DbEngineType -ieq $StandaloneDbEngine # Doesn't work against Azure
				) {
					Write-Output (
						@() + (
							$Server.Databases['master'].ExecuteWithResults('SELECT NodeName, Status, is_current_owner as IsCurrentOwner FROM sys.dm_os_cluster_nodes').Tables[0].Rows | ForEach-Object {
								New-Object -TypeName psobject -Property @{
									Name = $_.NodeName
									Status = switch ($_.Status) {
										0 { 'Up' }
										1 { 'Down' }
										2 { 'Paused' }
										3 { 'Joining' }
										-1 { 'Uknown' }
										default { 'Uknown' }
									}
									IsCurrentOwner = $_.IsCurrentOwner
								}
							}
						)
					)
				} elseif (
					$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0 -and
					$DbEngineType -ieq $StandaloneDbEngine # Doesn't work against Azure
				) {
					Write-Output (
						$Server.Databases['master'].ExecuteWithResults('SELECT NodeName FROM sys.dm_os_cluster_nodes').Tables[0].Rows | ForEach-Object {
							New-Object -TypeName psobject -Property @{
								Name = $_.NodeName
								Status = $null
								IsCurrentOwner = $null
							}
						}
					)
				} elseif (
					$($Server.Information.Version).CompareTo($SQLServer2000) -ge 0 -and
					$DbEngineType -ieq $StandaloneDbEngine # Doesn't work against Azure
				) {
					Write-Output (
						$Server.Databases['master'].ExecuteWithResults('SELECT NodeName FROM ::fn_virtualservernodes()').Tables[0].Rows | ForEach-Object {
							New-Object -TypeName psobject -Property @{
								Name = $_.NodeName
								Status = $null
								IsCurrentOwner = $null
							}
						}
					)
				}
				else {
					Write-Output $null
				}
			}
			else {
				Write-Output $null
			}
		}
		catch {
			throw
		}
	}
	End {
		Remove-Variable -Name DbEngineType
	}
}



function Get-IndexInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)]
		[Microsoft.SqlServer.Management.Smo.IndexCollection]
		$IndexCollection
	)
	Begin {
	}
	Process {
		try {
			$IndexCollection | ForEach-Object {
				Write-Output (
					New-Object -TypeName PSObject -Property @{
						General = New-Object -TypeName PSObject -Property @{
							Name = $_.Name # System.String Name {get;set;}
							ID = $_.ID # System.Int32 ID {get;}
							IndexType = [String](Get-IndexTypeValue -IndexType $_.IndexType) # Microsoft.SqlServer.Management.Smo.IndexType IndexType {get;set;}
							IndexKeyType = [String](Get-IndexKeyTypeValue -IndexKeyType $_.IndexKeyType) # Microsoft.SqlServer.Management.Smo.IndexKeyType IndexKeyType {get;set;}
							ParentXmlIndex = $_.ParentXmlIndex # System.String ParentXmlIndex {get;set;}
							SecondaryXmlIndexType = [String](Get-SecondaryXmlIndexTypeValue -SecondaryXmlIndexType $_.SecondaryXmlIndexType) # Microsoft.SqlServer.Management.Smo.SecondaryXmlIndexType SecondaryXmlIndexType {get;set;}
							IndexedColumns = @() + (
								$_.IndexedColumns | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										ID = $_.ID # System.Int32 ID {get;}
										IsComputed = $_.IsComputed # System.Boolean IsComputed {get;}
										IsDescending = $_.Descending # System.Boolean Descending {get;set;}
										IsIncluded = $_.IsIncluded # System.Boolean IsIncluded {get;set;}
										Name = $_.Name # System.String Name {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Index Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							) # Microsoft.SqlServer.Management.Smo.IndexedColumnCollection IndexedColumns {get;}

							# Not part of the SSMS GUI but probably belong here
							CompactLargeObjects = $_.CompactLargeObjects # System.Boolean CompactLargeObjects {get;set;}
							HasCompressedPartitions = $_.HasCompressedPartitions # System.Boolean HasCompressedPartitions {get;}
							HasFilter = $_.HasFilter # System.Boolean HasFilter {get;}
							IsClustered = $_.IsClustered # System.Boolean IsClustered {get;set;}
							IsDisabled = $_.IsDisabled # System.Boolean IsDisabled {get;}
							IsFileTableDefined = $_.IsFileTableDefined # System.Boolean IsFileTableDefined {get;}
							IsFullTextKey = $_.IsFullTextKey # System.Boolean IsFullTextKey {get;set;}
							IsHypothetical = $_.IsHypothetical # System.Boolean IsHypothetical {get;set;}
							IsIndexOnComputed = $_.IsIndexOnComputed # System.Boolean IsIndexOnComputed {get;}
							IsIndexOnTable = $_.IsIndexOnTable # System.Boolean IsIndexOnTable {get;}
							IsPartitioned = $_.IsPartitioned # System.Boolean IsPartitioned {get;}
							IsSpatialIndex = $_.IsSpatialIndex # System.Boolean IsSpatialIndex {get;}
							IsSystemNamed = $_.IsSystemNamed # System.Boolean IsSystemNamed {get;}
							IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
							IsUnique = $_.IsUnique # System.Boolean IsUnique {get;set;}
							IsXmlIndex = $_.IsXmlIndex # System.Boolean IsXmlIndex {get;}
							SpaceUsedKB = $_.SpaceUsed # System.Double SpaceUsed {get;}
						}
						Options = New-Object -TypeName PSObject -Property @{
							General = New-Object -TypeName PSObject -Property @{
								# HA! These properties are backwards from what SSMS labels them as
								NoAutomaticRecomputation = $_.NoAutomaticRecomputation # System.Boolean NoAutomaticRecomputation {get;set;}
								IgnoreDuplicateKeys = $_.IgnoreDuplicateKeys # System.Boolean IgnoreDuplicateKeys {get;set;}
							}
							Locks = New-Object -TypeName PSObject -Property @{
								DisallowPageLocks = $_.DisallowPageLocks # System.Boolean DisallowPageLocks {get;set;}
								DisallowRowLocks = $_.DisallowRowLocks # System.Boolean DisallowRowLocks {get;set;}
							}
							Operation = New-Object -TypeName PSObject -Property @{
								OnlineIndexOperation = $_.OnlineIndexOperation # System.Boolean OnlineIndexOperation {get;set;}
								MaximumDegreeOfParallelism = $_.MaximumDegreeOfParallelism # System.Int32 MaximumDegreeOfParallelism {get;set;}
							}
							Storage = New-Object -TypeName PSObject -Property @{
								SortInTempdb = $_.SortInTempdb # System.Boolean SortInTempdb {get;set;}
								FillFactor = $_.FillFactor # System.Byte FillFactor {get;set;}
								PadIndex = $_.PadIndex # System.Boolean PadIndex {get;set;}
							}
						}
						Storage = New-Object -TypeName PSObject -Property @{
							FileGroup = $_.FileGroup # System.String FileGroup {get;set;}
							FileStreamFileGroup = $_.FileStreamFileGroup # System.String FileStreamFileGroup {get;set;}
							PartitionScheme = $_.PartitionScheme # System.String PartitionScheme {get;set;}
							FileStreamPartitionScheme = $_.FileStreamPartitionScheme # System.String FileStreamPartitionScheme {get;set;}
							PartitionSchemeParameters = @() + (
								$_.PartitionSchemeParameters | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										ID = $_.ID # System.Int32 ID {get;}
										Name = $_.Name # System.String Name {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlSmoObject Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}													
									}
								}
							) # Microsoft.SqlServer.Management.Smo.PartitionSchemeParameterCollection PartitionSchemeParameters {get;}
							PhysicalPartitions = @() + (
								$_.PhysicalPartitions | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										DataCompression = [String](Get-DataCompressionTypeValue -DataCompressionType $_.DataCompression) # Microsoft.SqlServer.Management.Smo.DataCompressionType DataCompression {get;set;}
										FileGroupName = $_.FileGroupName # System.String FileGroupName {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlSmoObject Parent {get;}
										PartitionNumber = $_.PartitionNumber # System.Int32 PartitionNumber {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										RangeType = [String](Get-RangeTypeValue -RangeType $_.RangeType) # Microsoft.SqlServer.Management.Smo.RangeType RangeType {get;set;}
										RightBoundaryValue = if ($_.RightBoundaryValue) { $_.RightBoundaryValue.ToString() } else { $null } # System.Object RightBoundaryValue {get;set;}
										RowCount = $_.RowCount # System.Double RowCount {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							) # Microsoft.SqlServer.Management.Smo.PhysicalPartitionCollection PhysicalPartitions {get;}
						}
						Spatial = New-Object -TypeName PSObject -Property @{
							BoundingBox = New-Object -TypeName PSObject -Property @{
								XMax = $_.BoundingBoxXMax # System.Double BoundingBoxXMax {get;set;}
								XMin = $_.BoundingBoxXMin # System.Double BoundingBoxXMin {get;set;}
								YMax = $_.BoundingBoxYMax # System.Double BoundingBoxYMax {get;set;}
								YMin = $_.BoundingBoxYMin # System.Double BoundingBoxYMin {get;set;}
							}
							General = New-Object -TypeName PSObject -Property @{
								SpatialIndexType = [String](Get-SpatialIndexTypeValue -SpatialIndexType $_.SpatialIndexType) # Microsoft.SqlServer.Management.Smo.SpatialIndexType SpatialIndexType {get;set;}
								CellsPerObject = $_.CellsPerObject # System.Int32 CellsPerObject {get;set;}
							}
							Grids = New-Object -TypeName PSObject -Property @{
								Level1Grid = [String](Get-SpatialGeoLevelSizeValue -SpatialGeoLevelSize $_.Level1Grid) # Microsoft.SqlServer.Management.Smo.SpatialGeoLevelSize Level1Grid {get;set;}
								Level2Grid = [String](Get-SpatialGeoLevelSizeValue -SpatialGeoLevelSize $_.Level2Grid) # Microsoft.SqlServer.Management.Smo.SpatialGeoLevelSize Level2Grid {get;set;}
								Level3Grid = [String](Get-SpatialGeoLevelSizeValue -SpatialGeoLevelSize $_.Level3Grid) # Microsoft.SqlServer.Management.Smo.SpatialGeoLevelSize Level3Grid {get;set;}
								Level4Grid = [String](Get-SpatialGeoLevelSizeValue -SpatialGeoLevelSize $_.Level4Grid) # Microsoft.SqlServer.Management.Smo.SpatialGeoLevelSize Level4Grid {get;set;}
							}

						}
						FilterDefinition = $_.FilterDefinition # System.String FilterDefinition {get;set;}

						#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.IndexEvents Events {get;}
						#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
						#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlSmoObject Parent {get;set;}
						#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
						#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
						#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
						#UserData = $_.UserData	# System.Object UserData {get;set;}

					}
				) 
			}
		}
		catch {
			throw
		}
	}
	End {
	}
}

function Get-IndexInformation2 {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)]
		[AllowNull()]
		[System.Data.DataRow[]]
		$IndexCollection
		,
		[Parameter(Mandatory=$true)]
		[AllowNull()]
		[System.Data.DataRow[]]
		$IndexedColumnCollection
		,
		[Parameter(Mandatory=$true)]
		[AllowNull()]
		[System.Data.DataRow[]]
		$PartitionSchemeParameterCollection
		,
		[Parameter(Mandatory=$true)]
		[AllowNull()]
		[System.Data.DataRow[]]
		$PhysicalPartitionCollection
	)
	Begin {
	}
	Process {
		try {
			$IndexCollection | Where-Object { $_.ID } | ForEach-Object {
				$IndexID = $_.ID

				Write-Output (
					New-Object -TypeName PSObject -Property @{
						General = New-Object -TypeName PSObject -Property @{
							Name = $_.Name # System.String Name {get;set;}
							ID = $_.ID # System.Int32 ID {get;}
							IndexType = [String](Get-IndexTypeValue -IndexType $_.IndexType) # Microsoft.SqlServer.Management.Smo.IndexType IndexType {get;set;}
							IndexKeyType = [String](Get-IndexKeyTypeValue -IndexKeyType $_.IndexKeyType) # Microsoft.SqlServer.Management.Smo.IndexKeyType IndexKeyType {get;set;}
							ParentXmlIndex = $_.ParentXmlIndex # System.String ParentXmlIndex {get;set;}
							SecondaryXmlIndexType = [String](Get-SecondaryXmlIndexTypeValue -SecondaryXmlIndexType $_.SecondaryXmlIndexType) # Microsoft.SqlServer.Management.Smo.SecondaryXmlIndexType SecondaryXmlIndexType {get;set;}
							IndexedColumns = @() + (
								$IndexedColumnCollection | Where-Object { $_.IndexID -eq $IndexID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										ID = $_.ID # System.Int32 ID {get;}
										IsComputed = $_.IsComputed # System.Boolean IsComputed {get;}
										IsDescending = $_.Descending # System.Boolean Descending {get;set;}
										IsIncluded = $_.IsIncluded # System.Boolean IsIncluded {get;set;}
										Name = $_.Name # System.String Name {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Index Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							) # Microsoft.SqlServer.Management.Smo.IndexedColumnCollection IndexedColumns {get;}

							# Not part of the SSMS GUI but probably belong here
							CompactLargeObjects = $_.CompactLargeObjects # System.Boolean CompactLargeObjects {get;set;}
							HasCompressedPartitions = $_.HasCompressedPartitions # System.Boolean HasCompressedPartitions {get;}
							HasFilter = $_.HasFilter # System.Boolean HasFilter {get;}
							IsClustered = $_.IsClustered # System.Boolean IsClustered {get;set;}
							IsDisabled = $_.IsDisabled # System.Boolean IsDisabled {get;}
							IsFileTableDefined = $_.IsFileTableDefined # System.Boolean IsFileTableDefined {get;}
							IsFullTextKey = $_.IsFullTextKey # System.Boolean IsFullTextKey {get;set;}
							IsHypothetical = $_.IsHypothetical # System.Boolean IsHypothetical {get;set;}
							IsIndexOnComputed = $_.IsIndexOnComputed # System.Boolean IsIndexOnComputed {get;}
							IsIndexOnTable = $_.IsIndexOnTable # System.Boolean IsIndexOnTable {get;}
							IsPartitioned = $_.IsPartitioned # System.Boolean IsPartitioned {get;}
							IsSpatialIndex = $_.IsSpatialIndex # System.Boolean IsSpatialIndex {get;}
							IsSystemNamed = $_.IsSystemNamed # System.Boolean IsSystemNamed {get;}
							IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
							IsUnique = $_.IsUnique # System.Boolean IsUnique {get;set;}
							IsXmlIndex = $_.IsXmlIndex # System.Boolean IsXmlIndex {get;}
							SpaceUsed = $_.SpaceUsed # System.Double SpaceUsed {get;}
						}
						Options = New-Object -TypeName PSObject -Property @{
							General = New-Object -TypeName PSObject -Property @{
								# HA! These properties are backwards from what SSMS labels them as
								NoAutomaticRecomputation = $_.NoAutomaticRecomputation # System.Boolean NoAutomaticRecomputation {get;set;}
								IgnoreDuplicateKeys = $_.IgnoreDuplicateKeys # System.Boolean IgnoreDuplicateKeys {get;set;}
							}
							Locks = New-Object -TypeName PSObject -Property @{
								DisallowPageLocks = $_.DisallowPageLocks # System.Boolean DisallowPageLocks {get;set;}
								DisallowRowLocks = $_.DisallowRowLocks # System.Boolean DisallowRowLocks {get;set;}
							}
							Operation = New-Object -TypeName PSObject -Property @{
								OnlineIndexOperation = $_.OnlineIndexOperation # System.Boolean OnlineIndexOperation {get;set;}
								MaximumDegreeOfParallelism = $_.MaximumDegreeOfParallelism # System.Int32 MaximumDegreeOfParallelism {get;set;}
							}
							Storage = New-Object -TypeName PSObject -Property @{
								SortInTempdb = $_.SortInTempdb # System.Boolean SortInTempdb {get;set;}
								FillFactor = $_.FillFactor # System.Byte FillFactor {get;set;}
								PadIndex = $_.PadIndex # System.Boolean PadIndex {get;set;}
							}
						}
						Storage = New-Object -TypeName PSObject -Property @{
							FileGroup = $_.FileGroup # System.String FileGroup {get;set;}
							FileStreamFileGroup = $_.FileStreamFileGroup # System.String FileStreamFileGroup {get;set;}
							PartitionScheme = $_.PartitionScheme # System.String PartitionScheme {get;set;}
							FileStreamPartitionScheme = $_.FileStreamPartitionScheme # System.String FileStreamPartitionScheme {get;set;}
							PartitionSchemeParameters = @() + (
								$PartitionSchemeParameterCollection | Where-Object { $_.IndexID -eq $IndexID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										ID = $_.ID # System.Int32 ID {get;}
										Name = $_.Name # System.String Name {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlSmoObject Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}													
									}
								}
							) # Microsoft.SqlServer.Management.Smo.PartitionSchemeParameterCollection PartitionSchemeParameters {get;}
							PhysicalPartitions = @() + (
								$PhysicalPartitionCollection | Where-Object { $_.IndexID -eq $IndexID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										DataCompression = [String](Get-DataCompressionTypeValue -DataCompressionType $_.DataCompression) # Microsoft.SqlServer.Management.Smo.DataCompressionType DataCompression {get;set;}
										FileGroupName = $_.FileGroupName # System.String FileGroupName {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlSmoObject Parent {get;}
										PartitionNumber = $_.PartitionNumber # System.Int32 PartitionNumber {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										RangeType = [String](Get-RangeTypeValue -RangeType $_.RangeType) # Microsoft.SqlServer.Management.Smo.RangeType RangeType {get;set;}
										RightBoundaryValue = if ($_.RightBoundaryValue) { $_.RightBoundaryValue.ToString() } else { $null } # System.Object RightBoundaryValue {get;set;}
										RowCount = $_.RowCount # System.Double RowCount {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							) # Microsoft.SqlServer.Management.Smo.PhysicalPartitionCollection PhysicalPartitions {get;}
						}
						Spatial = New-Object -TypeName PSObject -Property @{
							BoundingBox = New-Object -TypeName PSObject -Property @{
								XMax = $_.BoundingBoxXMax # System.Double BoundingBoxXMax {get;set;}
								XMin = $_.BoundingBoxXMin # System.Double BoundingBoxXMin {get;set;}
								YMax = $_.BoundingBoxYMax # System.Double BoundingBoxYMax {get;set;}
								YMin = $_.BoundingBoxYMin # System.Double BoundingBoxYMin {get;set;}
							}
							General = New-Object -TypeName PSObject -Property @{
								SpatialIndexType = [String](Get-SpatialIndexTypeValue -SpatialIndexType $_.SpatialIndexType) # Microsoft.SqlServer.Management.Smo.SpatialIndexType SpatialIndexType {get;set;}
								CellsPerObject = $_.CellsPerObject # System.Int32 CellsPerObject {get;set;}
							}
							Grids = New-Object -TypeName PSObject -Property @{
								Level1Grid = [String](Get-SpatialGeoLevelSizeValue -SpatialGeoLevelSize $_.Level1Grid) # Microsoft.SqlServer.Management.Smo.SpatialGeoLevelSize Level1Grid {get;set;}
								Level2Grid = [String](Get-SpatialGeoLevelSizeValue -SpatialGeoLevelSize $_.Level2Grid) # Microsoft.SqlServer.Management.Smo.SpatialGeoLevelSize Level2Grid {get;set;}
								Level3Grid = [String](Get-SpatialGeoLevelSizeValue -SpatialGeoLevelSize $_.Level3Grid) # Microsoft.SqlServer.Management.Smo.SpatialGeoLevelSize Level3Grid {get;set;}
								Level4Grid = [String](Get-SpatialGeoLevelSizeValue -SpatialGeoLevelSize $_.Level4Grid) # Microsoft.SqlServer.Management.Smo.SpatialGeoLevelSize Level4Grid {get;set;}
							}

						}
						FilterDefinition = $_.FilterDefinition # System.String FilterDefinition {get;set;}

						#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.IndexEvents Events {get;}
						#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
						#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlSmoObject Parent {get;set;}
						#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
						#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
						#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
						#UserData = $_.UserData	# System.Object UserData {get;set;}

					}
				) 
			}
		}
		catch {
			throw
		}
	}
	End {
	}
}

function Get-TriggerInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)]
		[Microsoft.SqlServer.Management.Smo.TriggerCollection]
		$TriggerCollection
	)
	Begin {
	}
	Process {
		try {
			$TriggerCollection | ForEach-Object {
				Write-Output (
					New-Object -TypeName PSObject -Property @{
						AnsiNullsStatus = $_.AnsiNullsStatus # System.Boolean AnsiNullsStatus {get;set;}
						AssemblyName = $_.AssemblyName # System.String AssemblyName {get;set;}
						ClassName = $_.ClassName # System.String ClassName {get;set;}
						CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
						DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
						Delete = $_.Delete # System.Boolean Delete {get;set;}
						DeleteOrder = [String](Get-AgentActivationOrderValue -ActivationOrder $_.DeleteOrder) # Microsoft.SqlServer.Management.Smo.Agent.ActivationOrder DeleteOrder {get;set;}
						#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.TriggerEvents Events {get;}
						ExecutionContext = [String](Get-ExecutionContextValue -ExecutionContext $_.ExecutionContext) # Microsoft.SqlServer.Management.Smo.ExecutionContext ExecutionContext {get;set;}
						ExecutionContextPrincipal = $_.ExecutionContextPrincipal # System.String ExecutionContextPrincipal {get;set;}
						#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
						ID = $_.ID # System.Int32 ID {get;}
						ImplementationType = [String](Get-ImplementationTypeValue -ImplementationType $_.ImplementationType) # Microsoft.SqlServer.Management.Smo.ImplementationType ImplementationType {get;set;}
						Insert = $_.Insert # System.Boolean Insert {get;set;}
						InsertOrder = [String](Get-AgentActivationOrderValue -ActivationOrder $_.InsertOrder) # Microsoft.SqlServer.Management.Smo.Agent.ActivationOrder InsertOrder {get;set;}
						InsteadOf = $_.InsteadOf # System.Boolean InsteadOf {get;set;}
						IsEnabled = $_.IsEnabled # System.Boolean IsEnabled {get;set;}
						IsEncrypted = $_.IsEncrypted # System.Boolean IsEncrypted {get;set;}
						IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
						MethodName = $_.MethodName # System.String MethodName {get;set;}
						Name = $_.Name # System.String Name {get;set;}
						NotForReplication = $_.NotForReplication # System.Boolean NotForReplication {get;set;}
						#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlSmoObject Parent {get;set;}
						#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
						QuotedIdentifierStatus = $_.QuotedIdentifierStatus # System.Boolean QuotedIdentifierStatus {get;set;}
						#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
						<#
						TextBody = $_.TextBody # System.String TextBody {get;set;}
						TextHeader = $_.TextHeader # System.String TextHeader {get;set;}
						TextMode = $_.TextMode # System.Boolean TextMode {get;set;}

#>
						Definition = $null #$_.TextBody # System.String TextBody {get;set;}
						Update = $_.Update # System.Boolean Update {get;set;}
						UpdateOrder = [String](Get-AgentActivationOrderValue -ActivationOrder $_.UpdateOrder) # Microsoft.SqlServer.Management.Smo.Agent.ActivationOrder UpdateOrder {get;set;}
						#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
						#UserData = $_.UserData	# System.Object UserData {get;set;}
					}
				) 
			}
		}
		catch {
			throw
		}
	}
	End {
	}
}

function Get-TriggerInformation2 {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)]
		[AllowNull()]
		[System.Data.DataRow[]]
		$TriggerCollection
	)
	Begin {
	}
	Process {
		try {
			$TriggerCollection | Where-Object { $_.ID } | ForEach-Object {
				Write-Output (
					New-Object -TypeName PSObject -Property @{
						AnsiNullsStatus = $_.AnsiNullsStatus # System.Boolean AnsiNullsStatus {get;set;}
						AssemblyName = $_.AssemblyName # System.String AssemblyName {get;set;}
						ClassName = $_.ClassName # System.String ClassName {get;set;}
						CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
						DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
						Delete = $_.Delete # System.Boolean Delete {get;set;}
						DeleteOrder = [String](Get-AgentActivationOrderValue -ActivationOrder $_.DeleteOrder) # Microsoft.SqlServer.Management.Smo.Agent.ActivationOrder DeleteOrder {get;set;}
						#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.TriggerEvents Events {get;}
						ExecutionContext = [String](Get-ExecutionContextValue -ExecutionContext $_.ExecutionContext) # Microsoft.SqlServer.Management.Smo.ExecutionContext ExecutionContext {get;set;}
						ExecutionContextPrincipal = $_.ExecutionContextPrincipal # System.String ExecutionContextPrincipal {get;set;}
						#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
						ID = $_.ID # System.Int32 ID {get;}
						ImplementationType = [String](Get-ImplementationTypeValue -ImplementationType $_.ImplementationType) # Microsoft.SqlServer.Management.Smo.ImplementationType ImplementationType {get;set;}
						Insert = $_.Insert # System.Boolean Insert {get;set;}
						InsertOrder = [String](Get-AgentActivationOrderValue -ActivationOrder $_.InsertOrder) # Microsoft.SqlServer.Management.Smo.Agent.ActivationOrder InsertOrder {get;set;}
						InsteadOf = $_.InsteadOf # System.Boolean InsteadOf {get;set;}
						IsEnabled = $_.IsEnabled # System.Boolean IsEnabled {get;set;}
						IsEncrypted = $_.IsEncrypted # System.Boolean IsEncrypted {get;set;}
						IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
						MethodName = $_.MethodName # System.String MethodName {get;set;}
						Name = $_.Name # System.String Name {get;set;}
						NotForReplication = $_.NotForReplication # System.Boolean NotForReplication {get;set;}
						#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlSmoObject Parent {get;set;}
						#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
						QuotedIdentifierStatus = $_.QuotedIdentifierStatus # System.Boolean QuotedIdentifierStatus {get;set;}
						#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
						<#
						TextBody = $_.TextBody # System.String TextBody {get;set;}
						TextHeader = $_.TextHeader # System.String TextHeader {get;set;}
						TextMode = $_.TextMode # System.Boolean TextMode {get;set;}
						#>

						Definition = if ($_.IsSystemObject -eq $true) {
							# Don't include definitions for system objects
							$null
						} else {
							if (-not [String]::IsNullOrEmpty($_.Definition)) {
								$_.Definition.Trim()
							} else {
								$null
							}
						}


						Update = $_.Update # System.Boolean Update {get;set;}
						UpdateOrder = [String](Get-AgentActivationOrderValue -ActivationOrder $_.UpdateOrder) # Microsoft.SqlServer.Management.Smo.Agent.ActivationOrder UpdateOrder {get;set;}
						#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
						#UserData = $_.UserData	# System.Object UserData {get;set;}
					}
				) 
			}
		}
		catch {
			throw
		}
	}
	End {
	}
}


function Get-ServerConfigurationInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Server]
		$Server
	)

	# See http://msdn.microsoft.com/en-us/library/ms189631
	try {

		# In SMO if you specify a port number on the connection $Server.Name also contains the port number
		# To get around this "feature" I'm just executing the same statement that SMO executes against the DB Engine
		$ServerName = $Server.Databases['master'].ExecuteWithResults('SELECT SERVERPROPERTY(N''servername'') AS ServerName').Tables[0].Rows[0].ServerName
		$DbEngineType = [String](Get-DatabaseEngineTypeValue -DatabaseEngineType $Server.ServerType)

		Write-Output (
			New-Object -TypeName psobject -Property @{
				General = New-Object -TypeName psobject -Property @{

					# In SMO if you specify a port number on the connection $Server.Name also contains the port number
					# To get around this "feature" I'm just executing the same statement that SMO executes against the DB Engine
					#Name = $Server.Name
					Name = $ServerName
					GlobalName = $Server.Databases['master'].ExecuteWithResults('SELECT @@servername AS ServerName').Tables[0].Rows[0].ServerName

					#Product = "$($Server.Information.Product) $(Get-SqlServerVersionName -MajorVersion $Server.Information.Version.Major -MinorVersion $Server.Information.Version.Minor) ($($Server.ProductLevel)) -  $($Server.Information.VersionString) ($($Server.Platform))"
					Product = "$($Server.Information.Product) $(Get-SqlServerVersionName -MajorVersion $Server.Information.Version.Major -MinorVersion $Server.Information.Version.Minor)"
					OperatingSystem = $Server.Information.OSVersion # System.String OSVersion {get;}
					Platform = $Server.Information.Platform # System.String Platform {get;}
					Version = $Server.Information.VersionString # System.String VersionString {get;}
					Language = $Server.Information.Language # System.String Language {get;}
					MemoryMB = $Server.Information.PhysicalMemory # System.Int32 PhysicalMemory {get;}
					ProcessorCount = $Server.Information.Processors # System.Int32 Processors {get;}
					RootDirectory = $Server.Information.RootDirectory # System.String RootDirectory {get;}
					ServerCollation = $Server.Information.Collation # System.String Collation {get;}
					IsClustered = $Server.Information.IsClustered # System.Boolean IsClustered {get;}
					IsHadrEnabled = $Server.Information.IsHadrEnabled # System.Boolean IsHadrEnabled {get;}

					IsClusteredInstance = $Server.Information.IsClustered # Add as alias ?
					IsAlwaysOnEnabled = $Server.Information.IsHadrEnabled # Add as alias ?

					# Not part of the SSMS GUI but I think these properties belong here
					ComputerName = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
						# For SQL 2005 and up use SERVERPROPERTY(N'ComputerNamePhysicalNetBIOS')
						$Server.Databases['master'].ExecuteWithResults('SELECT SERVERPROPERTY(N''ComputerNamePhysicalNetBIOS'') AS ComputerNamePhysicalNetBIOS').Tables[0].Rows[0].ComputerNamePhysicalNetBIOS
					} else {
						# For SQL 2000 use SERVERPROPERTY(N'MachineName')
						# Note that for clustered servers this will return the virtual name and not the physical machine name
						$Server.Databases['master'].ExecuteWithResults('SELECT SERVERPROPERTY(N''MachineName'') AS ComputerNamePhysicalNetBIOS').Tables[0].Rows[0].ComputerNamePhysicalNetBIOS
					}

					ProductLevel = $Server.Information.ProductLevel # System.String ProductLevel {get;}
					Edition = $Server.Information.Edition # System.String Edition {get;}

					# This is an approximation at best. It *could* be wrong if these logins
					# have been dropped and recreated since SQL Server was originally installed
					InstallDate = $Server.Logins | Where-Object {
						$_.Name -ieq 'nt authority\system' -or 
						$_.Name -ieq 'builtin\administrators'
					} | Sort-Object -Property CreateDate | Select-Object -First 1 | ForEach-Object {
						$_.CreateDate # System.DateTime CreateDate {get;}
					}
					ServerType = $DbEngineType

					# Temporarily putting this here but will move it to a performance section
					# This is not available until SQL 2008, BTW
					MemoryInUseKB = $Server.PhysicalMemoryUsageInKB # System.Int64 PhysicalMemoryUsageInKB {get;}
				}
				Memory = New-Object -TypeName psobject -Property @{
					# AWE no longer available in SQL 2012
					UseAweToAllocateMemory = if (($Server.Information.Version).CompareTo($SQLServer2012) -ge 0) { 
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = $null
							RunningValue = $null
							DefaultValue = $null
							MinimumValue = $null
							MaximumValue = $null
							IsAdvanced = $null
							IsDynamic = $null
							ConfigurationName = 'awe enabled'
							FriendlyName = 'Use AWE to allocate memory'
							Description = 'This option is no longer available starting with SQL Server 2012'
						}
					} else {
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = if ($Server.Configuration.AweEnabled.ConfigValue -gt 0) { $true } else { $false }
							RunningValue = if ($Server.Configuration.AweEnabled.RunValue -gt 0) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $Server.Configuration.AweEnabled.Minimum # System.Int32 Minimum {get;}
							MaximumValue = $Server.Configuration.AweEnabled.Maximum # System.Int32 Maximum {get;}
							IsAdvanced = $Server.Configuration.AweEnabled.IsAdvanced # System.Boolean IsAdvanced {get;}
							IsDynamic = $Server.Configuration.AweEnabled.IsDynamic # System.Boolean IsDynamic {get;}
							ConfigurationName = $Server.Configuration.AweEnabled.DisplayName #'awe enabled'	# System.String DisplayName {get;}
							FriendlyName = 'Use AWE to allocate memory'
							Description = $Server.Configuration.AweEnabled.Description # System.String Description {get;}
						}
					}
					MinServerMemoryMB = New-Object -TypeName psobject -Property @{
						ConfiguredValue = $Server.Configuration.MinServerMemory.ConfigValue
						RunningValue = $Server.Configuration.MinServerMemory.RunValue
						DefaultValue = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) { 8 } else { 0 }
						MinimumValue = $Server.Configuration.MinServerMemory.Minimum
						MaximumValue = $Server.Configuration.MinServerMemory.Maximum
						IsAdvanced = $Server.Configuration.MinServerMemory.IsAdvanced
						IsDynamic = $Server.Configuration.MinServerMemory.IsDynamic
						ConfigurationName = $Server.Configuration.MinServerMemory.DisplayName #'min server memory (MB)'
						FriendlyName = 'Minimum server memory (in MB)'
						Description = $Server.Configuration.MinServerMemory.Description
					}
					MaxServerMemoryMB = New-Object -TypeName psobject -Property @{
						ConfiguredValue = $Server.Configuration.MaxServerMemory.ConfigValue
						RunningValue = $Server.Configuration.MaxServerMemory.RunValue
						DefaultValue = 2147483647
						MinimumValue = $Server.Configuration.MaxServerMemory.Minimum
						MaximumValue = $Server.Configuration.MaxServerMemory.Maximum
						IsAdvanced = $Server.Configuration.MaxServerMemory.IsAdvanced
						IsDynamic = $Server.Configuration.MaxServerMemory.IsDynamic
						ConfigurationName = $Server.Configuration.MaxServerMemory.DisplayName #'max server memory (MB)'
						FriendlyName = 'Maximum server memory (in MB)'
						Description = $Server.Configuration.MaxServerMemory.Description
					}
					IndexCreationMemoryKB = New-Object -TypeName psobject -Property @{
						ConfiguredValue = $Server.Configuration.IndexCreateMemory.ConfigValue
						RunningValue = $Server.Configuration.IndexCreateMemory.RunValue
						DefaultValue = 0
						MinimumValue = $Server.Configuration.IndexCreateMemory.Minimum
						MaximumValue = $Server.Configuration.IndexCreateMemory.Maximum
						IsAdvanced = $Server.Configuration.IndexCreateMemory.IsAdvanced
						IsDynamic = $Server.Configuration.IndexCreateMemory.IsDynamic
						ConfigurationName = $Server.Configuration.IndexCreateMemory.DisplayName #'index create memory (KB)'
						FriendlyName = 'Index creation memory (in KB, 0 = dynamic memory)'
						Description = $Server.Configuration.IndexCreateMemory.Description
					}
					MinMemoryPerQueryKB = New-Object -TypeName psobject -Property @{
						ConfiguredValue = $Server.Configuration.MinMemoryPerQuery.ConfigValue
						RunningValue = $Server.Configuration.MinMemoryPerQuery.RunValue
						DefaultValue = 1024
						MinimumValue = $Server.Configuration.MinMemoryPerQuery.Minimum
						MaximumValue = $Server.Configuration.MinMemoryPerQuery.Maximum
						IsAdvanced = $Server.Configuration.MinMemoryPerQuery.IsAdvanced
						IsDynamic = $Server.Configuration.MinMemoryPerQuery.IsDynamic
						ConfigurationName = $Server.Configuration.MinMemoryPerQuery.DisplayName #'min memory per query (KB)'
						FriendlyName = 'Mimimum memory per query (in KB)'
						Description = $Server.Configuration.MinMemoryPerQuery.Description
					}

					# Not part of the SSMS GUI but I think these properties belong here
					SetWorkingSetSize = New-Object -TypeName psobject -Property @{
						ConfiguredValue = if ($Server.Configuration.SetWorkingSetSize.ConfigValue -gt 0) { $true } else { $false } 
						RunningValue = if ($Server.Configuration.SetWorkingSetSize.RunValue -gt 0) { $true } else { $false } 
						DefaultValue = $false
						MinimumValue = $Server.Configuration.SetWorkingSetSize.Minimum
						MaximumValue = $Server.Configuration.SetWorkingSetSize.Maximum
						IsAdvanced = $Server.Configuration.SetWorkingSetSize.IsAdvanced
						IsDynamic = $Server.Configuration.SetWorkingSetSize.IsDynamic
						ConfigurationName = $Server.Configuration.SetWorkingSetSize.DisplayName #'set working set size'
						FriendlyName = 'Set Working Set Size'
						Description = $Server.Configuration.SetWorkingSetSize.Description
					}
				}
				Processor = New-Object -TypeName psobject -Property @{
					AutoSetProcessorAffinityMask = New-Object -TypeName psobject -Property @{
						ConfiguredValue = if ($Server.Configuration.AffinityMask.ConfigValue -eq 0) { $true } else { $false }
						RunningValue = if ($Server.Configuration.AffinityMask.RunValue -eq 0) { $true } else { $false }
						DefaultValue = $true
						MinimumValue = $Server.Configuration.AffinityMask.Minimum
						MaximumValue = $Server.Configuration.AffinityMask.Maximum
						IsAdvanced = $Server.Configuration.AffinityMask.IsAdvanced
						IsDynamic = $Server.Configuration.AffinityMask.IsDynamic
						ConfigurationName = $Server.Configuration.AffinityMask.DisplayName #'affinity mask'
						FriendlyName = 'Automatically set processor affinity mask for all processors'
						Description = $Server.Configuration.AffinityMask.Description
					}
					AutoSetIoAffinityMask = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = if ($Server.Configuration.AffinityIOMask.ConfigValue -eq 0) { $true } else { $false }
							RunningValue = if ($Server.Configuration.AffinityIOMask.RunValue -eq 0) { $true } else { $false }
							DefaultValue = $true
							MinimumValue = $Server.Configuration.AffinityIOMask.Minimum
							MaximumValue = $Server.Configuration.AffinityIOMask.Maximum
							IsAdvanced = $Server.Configuration.AffinityIOMask.IsAdvanced
							IsDynamic = $Server.Configuration.AffinityIOMask.IsDynamic
							ConfigurationName = $Server.Configuration.AffinityIOMask.DisplayName #'affinity I/O mask'
							FriendlyName = 'Automatically set I/O affinity mask for all processors'
							Description = $Server.Configuration.AffinityIOMask.Description
						}
					} else {
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = $null
							RunningValue = $null
							DefaultValue = $null
							MinimumValue = $null
							MaximumValue = $null
							IsAdvanced = $null
							IsDynamic = $null
							ConfigurationName = 'affinity I/O mask'
							FriendlyName = 'Automatically set I/O affinity mask for all processors'
							Description = 'This option is available starting with SQL Server 2005'
						} 
					}
					MaxWorkerThreads = New-Object -TypeName psobject -Property @{
						ConfiguredValue = $Server.Configuration.MaxWorkerThreads.ConfigValue
						RunningValue = $Server.Configuration.MaxWorkerThreads.RunValue
						DefaultValue = 0
						MinimumValue = $Server.Configuration.MaxWorkerThreads.Minimum
						MaximumValue = $Server.Configuration.MaxWorkerThreads.Maximum
						IsAdvanced = $Server.Configuration.MaxWorkerThreads.IsAdvanced
						IsDynamic = $Server.Configuration.MaxWorkerThreads.IsDynamic
						ConfigurationName = $Server.Configuration.MaxWorkerThreads.DisplayName #'max worker threads'
						FriendlyName = 'Maximum worker threads'
						Description = $Server.Configuration.MaxWorkerThreads.Description
					}
					BoostSqlServerPriority = New-Object -TypeName psobject -Property @{
						ConfiguredValue = if ($Server.Configuration.PriorityBoost.ConfigValue -gt 0) { $true } else { $false }
						RunningValue = if ($Server.Configuration.PriorityBoost.RunValue -gt 0) { $true } else { $false }
						DefaultValue = $false
						MinimumValue = $Server.Configuration.PriorityBoost.Minimum
						MaximumValue = $Server.Configuration.PriorityBoost.Maximum
						IsAdvanced = $Server.Configuration.PriorityBoost.IsAdvanced
						IsDynamic = $Server.Configuration.PriorityBoost.IsDynamic
						ConfigurationName = $Server.Configuration.PriorityBoost.DisplayName #'priority boost'
						FriendlyName = 'Boost SQL Server priority'
						Description = $Server.Configuration.PriorityBoost.Description
					}
					UseWindowsFibers = New-Object -TypeName psobject -Property @{
						ConfiguredValue = if ($Server.Configuration.LightweightPooling.ConfigValue -gt 0) { $true } else { $false }
						RunningValue = if ($Server.Configuration.LightweightPooling.RunValue -gt 0) { $true } else { $false }
						DefaultValue = $false
						MinimumValue = $Server.Configuration.LightweightPooling.Minimum
						MaximumValue = $Server.Configuration.LightweightPooling.Maximum
						IsAdvanced = $Server.Configuration.LightweightPooling.IsAdvanced
						IsDynamic = $Server.Configuration.LightweightPooling.IsDynamic
						ConfigurationName = $Server.Configuration.LightweightPooling.DisplayName #'lightweight pooling'
						FriendlyName = 'Use Windows fibers (lightweight pooling)'
						Description = $Server.Configuration.LightweightPooling.Description
					}

					# Not part of the SSMS GUI but I think these properties belong here
					AffinityIOMask = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = $Server.Configuration.AffinityIOMask.ConfigValue
							RunningValue = $Server.Configuration.AffinityIOMask.RunValue
							DefaultValue = 0
							MinimumValue = $Server.Configuration.AffinityIOMask.Minimum
							MaximumValue = $Server.Configuration.AffinityIOMask.Maximum
							IsAdvanced = $Server.Configuration.AffinityIOMask.IsAdvanced
							IsDynamic = $Server.Configuration.AffinityIOMask.IsDynamic
							ConfigurationName = $Server.Configuration.AffinityIOMask.DisplayName #'affinity I/O mask'
							FriendlyName = 'Affinity I/O mask'
							Description = $Server.Configuration.AffinityIOMask.Description
						}
					} else {
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = $null
							RunningValue = $null
							DefaultValue = $null
							MinimumValue = $null
							MaximumValue = $null
							IsAdvanced = $null
							IsDynamic = $null
							ConfigurationName = 'affinity I/O mask'
							FriendlyName = 'Affinity I/O mask'
							Description = 'This option is available starting with SQL Server 2005'
						} 
					}
					AffinityMask = New-Object -TypeName psobject -Property @{
						ConfiguredValue = $Server.Configuration.AffinityMask.ConfigValue
						RunningValue = $Server.Configuration.AffinityMask.RunValue
						DefaultValue = 0
						MinimumValue = $Server.Configuration.AffinityMask.Minimum
						MaximumValue = $Server.Configuration.AffinityMask.Maximum
						IsAdvanced = $Server.Configuration.AffinityMask.IsAdvanced
						IsDynamic = $Server.Configuration.AffinityMask.IsDynamic
						ConfigurationName = $Server.Configuration.AffinityMask.DisplayName #'affinity mask'
						FriendlyName = 'Affinity mask'
						Description = $Server.Configuration.AffinityMask.Description
					}
					Affinity64IOMask = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = $Server.Configuration.Affinity64IOMask.ConfigValue
							RunningValue = $Server.Configuration.Affinity64IOMask.RunValue
							DefaultValue = 0
							MinimumValue = $Server.Configuration.Affinity64IOMask.Minimum
							MaximumValue = $Server.Configuration.Affinity64IOMask.Maximum
							IsAdvanced = $Server.Configuration.Affinity64IOMask.IsAdvanced
							IsDynamic = $Server.Configuration.Affinity64IOMask.IsDynamic
							ConfigurationName = $Server.Configuration.Affinity64IOMask.DisplayName #'affinity64 I/O mask'
							FriendlyName = 'Affinity64 I/O mask'
							Description = $Server.Configuration.Affinity64IOMask.Description
						}
					} else {
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = $null
							RunningValue = $null
							DefaultValue = $null
							MinimumValue = $null
							MaximumValue = $null
							IsAdvanced = $null
							IsDynamic = $null
							ConfigurationName = 'affinity64 I/O mask'
							FriendlyName = 'Affinity64 I/O mask'
							Description = 'This option is available starting with SQL Server 2005'
						}
					}
					Affinity64Mask = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = $Server.Configuration.Affinity64Mask.ConfigValue
							RunningValue = $Server.Configuration.Affinity64Mask.RunValue
							DefaultValue = 0
							MinimumValue = $Server.Configuration.Affinity64Mask.Minimum
							MaximumValue = $Server.Configuration.Affinity64Mask.Maximum
							IsAdvanced = $Server.Configuration.Affinity64Mask.IsAdvanced
							IsDynamic = $Server.Configuration.Affinity64Mask.IsDynamic
							ConfigurationName = $Server.Configuration.Affinity64Mask.DisplayName #'affinity64 mask'
							FriendlyName = 'Affinity64 mask'
							Description = $Server.Configuration.Affinity64Mask.Description
						}
					} else {
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = $null
							RunningValue = $null
							DefaultValue = $null
							MinimumValue = $null
							MaximumValue = $null
							IsAdvanced = $null
							IsDynamic = $null
							ConfigurationName = 'affinity64 mask'
							FriendlyName = 'Affinity64 mask'
							Description = 'This option is available starting with SQL Server 2005'
						}
					}
				}
				Security = New-Object -TypeName psobject -Property @{
					AuthenticationMode = [String](Get-AuthenticationModeValue -LoginMode $Server.LoginMode) # Microsoft.SqlServer.Management.Smo.ServerLoginMode LoginMode {get;set;}
					LoginAuditLevel = [String](Get-AuditLevelValue -AuditLevel $Server.Settings.AuditLevel) # Microsoft.SqlServer.Management.Smo.AuditLevel AuditLevel {get;set;}
					ServerProxyAccount = New-Object -TypeName psobject -Property @{
						Enabled = $Server.ProxyAccount.IsEnabled # System.Boolean IsEnabled {get;set;}
						Username = $Server.ProxyAccount.WindowsAccount # System.String WindowsAccount {get;set;}
						#Password = $null	# Not exposed through SMO
					}
					EnableCommonCriteriaCompliance = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = if ($Server.Configuration.CommonCriteriaComplianceEnabled.ConfigValue -gt 0) { $true } else { $false }
							RunningValue = if ($Server.Configuration.CommonCriteriaComplianceEnabled.RunValue -gt 0) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $Server.Configuration.CommonCriteriaComplianceEnabled.Minimum
							MaximumValue = $Server.Configuration.CommonCriteriaComplianceEnabled.Maximum
							IsAdvanced = $Server.Configuration.CommonCriteriaComplianceEnabled.IsAdvanced
							IsDynamic = $Server.Configuration.CommonCriteriaComplianceEnabled.IsDynamic
							ConfigurationName = $Server.Configuration.CommonCriteriaComplianceEnabled.DisplayName #'common criteria compliance enabled'
							FriendlyName = 'Common Criteria compliance enabled'
							Description = $Server.Configuration.CommonCriteriaComplianceEnabled.Description
						}
					} else {
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = $null
							RunningValue = $null
							DefaultValue = $null
							MinimumValue = $null
							MaximumValue = $null
							IsAdvanced = $null
							IsDynamic = $null
							ConfigurationName = 'common criteria compliance enabled'
							FriendlyName = 'Common Criteria compliance enabled'
							Description = 'This option is available starting with SQL Server 2005'
						}
					}
					EnableC2AuditTracing = New-Object -TypeName psobject -Property @{
						ConfiguredValue = if ($Server.Configuration.C2AuditMode.ConfigValue -gt 0) { $true } else { $false }
						RunningValue = if ($Server.Configuration.C2AuditMode.RunValue -gt 0) { $true } else { $false }
						DefaultValue = $false
						MinimumValue = $Server.Configuration.C2AuditMode.Minimum
						MaximumValue = $Server.Configuration.C2AuditMode.Maximum
						IsAdvanced = $Server.Configuration.C2AuditMode.IsAdvanced
						IsDynamic = $Server.Configuration.C2AuditMode.IsDynamic
						ConfigurationName = $Server.Configuration.C2AuditMode.DisplayName #'c2 audit mode'
						FriendlyName = 'C2 audit tracing enabled'
						Description = $Server.Configuration.C2AuditMode.Description
					}
					CrossDatabaseOwnershipChaining = New-Object -TypeName psobject -Property @{
						ConfiguredValue = if ($Server.Configuration.CrossDBOwnershipChaining.ConfigValue -gt 0) { $true } else { $false }
						RunningValue = if ($Server.Configuration.CrossDBOwnershipChaining.RunValue -gt 0) { $true } else { $false }
						DefaultValue = $false
						MinimumValue = $Server.Configuration.CrossDBOwnershipChaining.Minimum
						MaximumValue = $Server.Configuration.CrossDBOwnershipChaining.Maximum
						IsAdvanced = $Server.Configuration.CrossDBOwnershipChaining.IsAdvanced
						IsDynamic = $Server.Configuration.CrossDBOwnershipChaining.IsDynamic
						ConfigurationName = $Server.Configuration.CrossDBOwnershipChaining.DisplayName #'cross db ownership chaining'
						FriendlyName = 'Cross database ownership chaining enabled'
						Description = $Server.Configuration.CrossDBOwnershipChaining.Description
					}
				}
				Connections = New-Object -TypeName psobject -Property @{
					MaxConcurrentConnections = New-Object -TypeName psobject -Property @{
						ConfiguredValue = $Server.Configuration.UserConnections.ConfigValue
						RunningValue = $Server.Configuration.UserConnections.RunValue
						DefaultValue = 0
						MinimumValue = $Server.Configuration.UserConnections.Minimum
						MaximumValue = $Server.Configuration.UserConnections.Maximum
						IsAdvanced = $Server.Configuration.UserConnections.IsAdvanced
						IsDynamic = $Server.Configuration.UserConnections.IsDynamic
						ConfigurationName = $Server.Configuration.UserConnections.DisplayName #'user connections'
						FriendlyName = 'Maximum number of concurrent connections (0 = unlimited)'
						Description = $Server.Configuration.UserConnections.Description
					}
					QueryGovernor = New-Object -TypeName psobject -Property @{
						Enabled = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if ($Server.Configuration.QueryGovernorCostLimit.ConfigValue -gt 0) { $true } else { $false }
							RunningValue = if ($Server.Configuration.QueryGovernorCostLimit.RunValue -gt 0) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $false
							MaximumValue = $true
							IsAdvanced = $Server.Configuration.QueryGovernorCostLimit.IsAdvanced
							IsDynamic = $Server.Configuration.QueryGovernorCostLimit.IsDynamic
							ConfigurationName = $Server.Configuration.QueryGovernorCostLimit.DisplayName #'query governor cost limit'
							FriendlyName = 'Use query governor to precent long-running queries'
							Description = $Server.Configuration.QueryGovernorCostLimit.Description
						}
						TimeoutSeconds= New-Object -TypeName psobject -Property @{
							ConfiguredValue = $Server.Configuration.QueryGovernorCostLimit.ConfigValue
							RunningValue = $Server.Configuration.QueryGovernorCostLimit.RunValue
							DefaultValue = 0
							MinimumValue = $Server.Configuration.QueryGovernorCostLimit.Minimum
							MaximumValue = $Server.Configuration.QueryGovernorCostLimit.Maximum
							IsAdvanced = $Server.Configuration.QueryGovernorCostLimit.IsAdvanced
							IsDynamic = $Server.Configuration.QueryGovernorCostLimit.IsDynamic
							ConfigurationName = $Server.Configuration.QueryGovernorCostLimit.DisplayName #'query governor cost limit'
							FriendlyName = 'Query Governor cost limit'
							Description = $Server.Configuration.QueryGovernorCostLimit.Description
						}
					}

					# A.K.A. User Options....see http://msdn.microsoft.com/en-us/library/ms190763
					DefaultOptions = New-Object -TypeName psobject -Property @{
						InterimConstraintChecking = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if (($Server.Configuration.UserOptions.ConfigValue -band 1) -eq 1) { $true } else { $false }
							RunningValue = if (($Server.Configuration.UserOptions.RunValue -band 1) -eq 1) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $false
							MaximumValue = $true
							IsAdvanced = $Server.Configuration.UserOptions.IsAdvanced
							IsDynamic = $Server.Configuration.UserOptions.IsDynamic
							ConfigurationName = $Server.Configuration.UserOptions.DisplayName #'user options'
							FriendlyName = 'Interim/Deferred Constraint Checking'
							Description = $Server.Configuration.UserOptions.Description
						}
						ImplicitTransactions = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if (($Server.Configuration.UserOptions.ConfigValue -band 2) -eq 2) { $true } else { $false }
							RunningValue = if (($Server.Configuration.UserOptions.RunValue -band 2) -eq 2) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $false
							MaximumValue = $true
							IsAdvanced = $Server.Configuration.UserOptions.IsAdvanced
							IsDynamic = $Server.Configuration.UserOptions.IsDynamic
							ConfigurationName = $Server.Configuration.UserOptions.DisplayName #'user options'
							FriendlyName = 'Implicit Transactions Default'
							Description = $Server.Configuration.UserOptions.Description
						}
						CursorCloseOnCommit = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if (($Server.Configuration.UserOptions.ConfigValue -band 4) -eq 4) { $true } else { $false }
							RunningValue = if (($Server.Configuration.UserOptions.RunValue -band 4) -eq 4) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $false
							MaximumValue = $true
							IsAdvanced = $Server.Configuration.UserOptions.IsAdvanced
							IsDynamic = $Server.Configuration.UserOptions.IsDynamic
							ConfigurationName = $Server.Configuration.UserOptions.DisplayName #'user options'
							FriendlyName = 'Cursor Close On Commit Default'
							Description = $Server.Configuration.UserOptions.Description
						}
						AnsiWarnings = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if (($Server.Configuration.UserOptions.ConfigValue -band 8) -eq 8) { $true } else { $false }
							RunningValue = if (($Server.Configuration.UserOptions.RunValue -band 8) -eq 8) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $false
							MaximumValue = $true
							IsAdvanced = $Server.Configuration.UserOptions.IsAdvanced
							IsDynamic = $Server.Configuration.UserOptions.IsDynamic
							ConfigurationName = $Server.Configuration.UserOptions.DisplayName #'user options'
							FriendlyName = 'Ansi Warnings Default'
							Description = $Server.Configuration.UserOptions.Description
						}
						AnsiPadding = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if (($Server.Configuration.UserOptions.ConfigValue -band 16) -eq 16) { $true } else { $false }
							RunningValue = if (($Server.Configuration.UserOptions.RunValue -band 16) -eq 16) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $false
							MaximumValue = $true
							IsAdvanced = $Server.Configuration.UserOptions.IsAdvanced
							IsDynamic = $Server.Configuration.UserOptions.IsDynamic
							ConfigurationName = $Server.Configuration.UserOptions.DisplayName #'user options'
							FriendlyName = 'Ansi Padding Default'
							Description = $Server.Configuration.UserOptions.Description
						}
						AnsiNulls = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if (($Server.Configuration.UserOptions.ConfigValue -band 32) -eq 32) { $true } else { $false }
							RunningValue = if (($Server.Configuration.UserOptions.RunValue -band 32) -eq 32) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $false
							MaximumValue = $true
							IsAdvanced = $Server.Configuration.UserOptions.IsAdvanced
							IsDynamic = $Server.Configuration.UserOptions.IsDynamic
							ConfigurationName = $Server.Configuration.UserOptions.DisplayName #'user options'
							FriendlyName = 'ANSI NULLS Default'
							Description = $Server.Configuration.UserOptions.Description
						}
						ArithmeticAbort = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if (($Server.Configuration.UserOptions.ConfigValue -band 64) -eq 64) { $true } else { $false }
							RunningValue = if (($Server.Configuration.UserOptions.RunValue -band 64) -eq 64) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $false
							MaximumValue = $true
							IsAdvanced = $Server.Configuration.UserOptions.IsAdvanced
							IsDynamic = $Server.Configuration.UserOptions.IsDynamic
							ConfigurationName = $Server.Configuration.UserOptions.DisplayName #'user options'
							FriendlyName = 'Arithmatic Abort Default'
							Description = $Server.Configuration.UserOptions.Description
						}
						ArithmeticIgnore = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if (($Server.Configuration.UserOptions.ConfigValue -band 128) -eq 128) { $true } else { $false }
							RunningValue = if (($Server.Configuration.UserOptions.RunValue -band 128) -eq 128) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $false
							MaximumValue = $true
							IsAdvanced = $Server.Configuration.UserOptions.IsAdvanced
							IsDynamic = $Server.Configuration.UserOptions.IsDynamic
							ConfigurationName = $Server.Configuration.UserOptions.DisplayName #'user options'
							FriendlyName = 'Arithmatic Ignore Default'
							Description = $Server.Configuration.UserOptions.Description
						}
						QuotedIdentifier = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if (($Server.Configuration.UserOptions.ConfigValue -band 256) -eq 256) { $true } else { $false }
							RunningValue = if (($Server.Configuration.UserOptions.RunValue -band 256) -eq 256) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $false
							MaximumValue = $true
							IsAdvanced = $Server.Configuration.UserOptions.IsAdvanced
							IsDynamic = $Server.Configuration.UserOptions.IsDynamic
							ConfigurationName = $Server.Configuration.UserOptions.DisplayName #'user options'
							FriendlyName = 'Quoted Identifier Default'
							Description = $Server.Configuration.UserOptions.Description
						}
						NoCount = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if (($Server.Configuration.UserOptions.ConfigValue -band 512) -eq 512) { $true } else { $false }
							RunningValue = if (($Server.Configuration.UserOptions.RunValue -band 512) -eq 512) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $false
							MaximumValue = $true
							IsAdvanced = $Server.Configuration.UserOptions.IsAdvanced
							IsDynamic = $Server.Configuration.UserOptions.IsDynamic
							ConfigurationName = $Server.Configuration.UserOptions.DisplayName #'user options'
							FriendlyName = 'No Count Default'
							Description = $Server.Configuration.UserOptions.Description
						}
						AnsiNullDefaultOn = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if (($Server.Configuration.UserOptions.ConfigValue -band 1024) -eq 1024) { $true } else { $false }
							RunningValue = if (($Server.Configuration.UserOptions.RunValue -band 1024) -eq 1024) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $false
							MaximumValue = $true
							IsAdvanced = $Server.Configuration.UserOptions.IsAdvanced
							IsDynamic = $Server.Configuration.UserOptions.IsDynamic
							ConfigurationName = $Server.Configuration.UserOptions.DisplayName #'user options'
							FriendlyName = 'ANSI NULL Default On'
							Description = $Server.Configuration.UserOptions.Description
						}
						AnsiNullDefaultOff = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if (($Server.Configuration.UserOptions.ConfigValue -band 2048) -eq 2048) { $true } else { $false }
							RunningValue = if (($Server.Configuration.UserOptions.RunValue -band 2048) -eq 2048) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $false
							MaximumValue = $true
							IsAdvanced = $Server.Configuration.UserOptions.IsAdvanced
							IsDynamic = $Server.Configuration.UserOptions.IsDynamic
							ConfigurationName = $Server.Configuration.UserOptions.DisplayName #'user options'
							FriendlyName = 'ANSI NULL Default Off'
							Description = $Server.Configuration.UserOptions.Description
						}
						ConcatNullYieldsNull = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if (($Server.Configuration.UserOptions.ConfigValue -band 4096) -eq 4096) { $true } else { $false }
							RunningValue = if (($Server.Configuration.UserOptions.RunValue -band 4096) -eq 4096) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $false
							MaximumValue = $true
							IsAdvanced = $Server.Configuration.UserOptions.IsAdvanced
							IsDynamic = $Server.Configuration.UserOptions.IsDynamic
							ConfigurationName = $Server.Configuration.UserOptions.DisplayName #'user options'
							FriendlyName = 'Concat Null Yields Null Default'
							Description = $Server.Configuration.UserOptions.Description
						}
						NumericRoundAbort = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if (($Server.Configuration.UserOptions.ConfigValue -band 8192) -eq 8192) { $true } else { $false }
							RunningValue = if (($Server.Configuration.UserOptions.RunValue -band 8192) -eq 8192) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $false
							MaximumValue = $true
							IsAdvanced = $Server.Configuration.UserOptions.IsAdvanced
							IsDynamic = $Server.Configuration.UserOptions.IsDynamic
							ConfigurationName = $Server.Configuration.UserOptions.DisplayName #'user options'
							FriendlyName = 'Numeric Round Abort Default'
							Description = $Server.Configuration.UserOptions.Description
						}
						XactAbort = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if (($Server.Configuration.UserOptions.ConfigValue -band 16384) -eq 16384) { $true } else { $false }
							RunningValue = if (($Server.Configuration.UserOptions.RunValue -band 16384) -eq 16384) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $false
							MaximumValue = $true
							IsAdvanced = $Server.Configuration.UserOptions.IsAdvanced
							IsDynamic = $Server.Configuration.UserOptions.IsDynamic
							ConfigurationName = $Server.Configuration.UserOptions.DisplayName #'user options'
							FriendlyName = 'Xact Abort Default'
							Description = $Server.Configuration.UserOptions.Description
						}
					}
					AllowRemoteConnections = New-Object -TypeName psobject -Property @{
						ConfiguredValue = if ($Server.Configuration.RemoteAccess.ConfigValue -gt 0) { $true } else { $false }
						RunningValue = if ($Server.Configuration.RemoteAccess.RunValue -gt 0) { $true } else { $false }
						DefaultValue = $true
						MinimumValue = $Server.Configuration.RemoteAccess.Minimum
						MaximumValue = $Server.Configuration.RemoteAccess.Maximum
						IsAdvanced = $Server.Configuration.RemoteAccess.IsAdvanced
						IsDynamic = $Server.Configuration.RemoteAccess.IsDynamic
						ConfigurationName = $Server.Configuration.RemoteAccess.DisplayName #'remote access'
						FriendlyName = 'Allow remote connections to this server'
						Description = $Server.Configuration.RemoteAccess.Description
					}
					RemoteQueryTimeoutSeconds = New-Object -TypeName psobject -Property @{
						ConfiguredValue = $Server.Configuration.RemoteQueryTimeout.ConfigValue
						RunningValue = $Server.Configuration.RemoteQueryTimeout.RunValue
						DefaultValue = 600
						MinimumValue = $Server.Configuration.RemoteQueryTimeout.Minimum
						MaximumValue = $Server.Configuration.RemoteQueryTimeout.Maximum
						IsAdvanced = $Server.Configuration.RemoteQueryTimeout.IsAdvanced
						IsDynamic = $Server.Configuration.RemoteQueryTimeout.IsDynamic
						ConfigurationName = $Server.Configuration.RemoteQueryTimeout.DisplayName #'remote query timeout (s)'
						FriendlyName = 'Remote query timeout (in seconds, 0 = no timeout)'
						Description = $Server.Configuration.RemoteQueryTimeout.Description
					}
					RequireDistributedTransactions = New-Object -TypeName psobject -Property @{
						ConfiguredValue = if ($Server.Configuration.RemoteProcTrans.ConfigValue -gt 0) { $true } else { $false }
						RunningValue = if ($Server.Configuration.RemoteProcTrans.RunValue -gt 0) { $true } else { $false }
						DefaultValue = $false
						MinimumValue = $Server.Configuration.RemoteProcTrans.Minimum
						MaximumValue = $Server.Configuration.RemoteProcTrans.Maximum
						IsAdvanced = $Server.Configuration.RemoteProcTrans.IsAdvanced
						IsDynamic = $Server.Configuration.RemoteProcTrans.IsDynamic
						ConfigurationName = $Server.Configuration.RemoteProcTrans.DisplayName #'remote proc trans'
						FriendlyName = 'Require distributed transactions for server-to-server communication'
						Description = $Server.Configuration.RemoteProcTrans.Description
					}

					# The following options are not exposed through the SSMS GUI but this is where I think they would be if they were
					AdHocDistributedQueriesEnabled = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = if ($Server.Configuration.AdHocDistributedQueriesEnabled.ConfigValue -gt 0) { $true } else { $false }
							RunningValue = if ($Server.Configuration.AdHocDistributedQueriesEnabled.RunValue -gt 0) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $Server.Configuration.AdHocDistributedQueriesEnabled.Minimum
							MaximumValue = $Server.Configuration.AdHocDistributedQueriesEnabled.Maximum
							IsAdvanced = $Server.Configuration.AdHocDistributedQueriesEnabled.IsAdvanced
							IsDynamic = $Server.Configuration.AdHocDistributedQueriesEnabled.IsDynamic
							ConfigurationName = $Server.Configuration.AdHocDistributedQueriesEnabled.DisplayName #'Ad Hoc Distributed Queries'
							FriendlyName = 'Ad Hoc Distributed Queries'
							Description = $Server.Configuration.AdHocDistributedQueriesEnabled.Description
						}
					} else {
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = $null
							RunningValue = $null
							DefaultValue = $null
							MinimumValue = $null
							MaximumValue = $null
							IsAdvanced = $null
							IsDynamic = $null
							ConfigurationName = 'Ad Hoc Distributed Queries'
							FriendlyName = 'Ad Hoc Distributed Queries'
							Description = 'This option is available starting with SQL Server 2005'
						} 
					}
					RemoteAdminConnectionsEnabled = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = if ($Server.Configuration.RemoteDacConnectionsEnabled.ConfigValue -gt 0) { $true } else { $false }
							RunningValue = if ($Server.Configuration.RemoteDacConnectionsEnabled.RunValue -gt 0) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $Server.Configuration.RemoteDacConnectionsEnabled.Minimum
							MaximumValue = $Server.Configuration.RemoteDacConnectionsEnabled.Maximum
							IsAdvanced = $Server.Configuration.RemoteDacConnectionsEnabled.IsAdvanced
							IsDynamic = $Server.Configuration.RemoteDacConnectionsEnabled.IsDynamic
							ConfigurationName = $Server.Configuration.RemoteDacConnectionsEnabled.DisplayName #'remote admin connections'
							FriendlyName = 'Allow Remote Admin Connections'
							Description = $Server.Configuration.RemoteDacConnectionsEnabled.Description
						}
					} else {
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = $null
							RunningValue = $null
							DefaultValue = $null
							MinimumValue = $null
							MaximumValue = $null
							IsAdvanced = $null
							IsDynamic = $null
							ConfigurationName = 'remote admin connections'
							FriendlyName = 'Allow Remote Admin Connections'
							Description = 'This option is available starting with SQL Server 2005'
						}
					}
				}
				DatabaseSettings = New-Object -TypeName psobject -Property @{
					IndexFillFactor = New-Object -TypeName psobject -Property @{
						ConfiguredValue = $Server.Configuration.FillFactor.ConfigValue
						RunningValue = $Server.Configuration.FillFactor.RunValue
						DefaultValue = 0
						MinimumValue = $Server.Configuration.FillFactor.Minimum
						MaximumValue = $Server.Configuration.FillFactor.Maximum
						IsAdvanced = $Server.Configuration.FillFactor.IsAdvanced
						IsDynamic = $Server.Configuration.FillFactor.IsDynamic
						ConfigurationName = $Server.Configuration.FillFactor.DisplayName #'fill factor (%)'
						FriendlyName = 'Default Index Fill Factor'
						Description = $Server.Configuration.FillFactor.Description
					}
					BackupMediaRetentionDays = New-Object -TypeName psobject -Property @{
						ConfiguredValue = $Server.Configuration.MediaRetention.ConfigValue
						RunningValue = $Server.Configuration.MediaRetention.RunValue
						DefaultValue = 0
						MinimumValue = $Server.Configuration.MediaRetention.Minimum
						MaximumValue = $Server.Configuration.MediaRetention.Maximum
						IsAdvanced = $Server.Configuration.MediaRetention.IsAdvanced
						IsDynamic = $Server.Configuration.MediaRetention.IsDynamic
						ConfigurationName = $Server.Configuration.MediaRetention.DisplayName #'media retention'
						FriendlyName = 'Default backup media retention (in days)'
						Description = $Server.Configuration.MediaRetention.Description
					}

					# Backup compression available up starting with SQL 2008
					CompressBackup = if ((($Server.Information.Version).CompareTo($SQLServer2008) -ge 0) -and ($SmoMajorVersion -ge 10)) {
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = if ($Server.Configuration.DefaultBackupCompression.ConfigValue -gt 0) { $true } else { $false }
							RunningValue = if ($Server.Configuration.DefaultBackupCompression.RunValue -gt 0) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $Server.Configuration.DefaultBackupCompression.Minimum
							MaximumValue = $Server.Configuration.DefaultBackupCompression.Maximum
							IsAdvanced = $Server.Configuration.DefaultBackupCompression.IsAdvanced
							IsDynamic = $Server.Configuration.DefaultBackupCompression.IsDynamic
							ConfigurationName = $Server.Configuration.DefaultBackupCompression.DisplayName #'backup compression default'
							FriendlyName = 'Compress Backups Default'
							Description = $Server.Configuration.DefaultBackupCompression.Description
						}
					} else {
						New-Object -TypeName psobject -Property @{
							ConfiguredValue = $null
							RunningValue = $null
							DefaultValue = $null
							MinimumValue = $null
							MaximumValue = $null
							IsAdvanced = $null
							IsDynamic = $null
							ConfigurationName = 'backup compression default'
							FriendlyName = 'Compress Backups Default'
							Description = if (($Server.Information.Version).CompareTo($SQLServer2008) -ge 0) {
								'This option is available starting with SQL Server 2008'
							} else {
								'SMO 2008 or higher is required to retreive this value'
							}
						}

						if (($Server.Information.Version).CompareTo($SQLServer2008) -ge 0) {
							Write-SqlServerDatabaseEngineInformationLog -Message "[$ServerName] Unable to retrieve backup compression settings - Server version is higher than SMO version." -MessageLevel Warning 
						}
					}
					RecoveryIntervalMinutes = New-Object -TypeName psobject -Property @{
						ConfiguredValue = $Server.Configuration.RecoveryInterval.ConfigValue
						RunningValue = $Server.Configuration.RecoveryInterval.RunValue
						DefaultValue = 0
						MinimumValue = $Server.Configuration.RecoveryInterval.Minimum
						MaximumValue = $Server.Configuration.RecoveryInterval.Maximum
						IsAdvanced = $Server.Configuration.RecoveryInterval.IsAdvanced
						IsDynamic = $Server.Configuration.RecoveryInterval.IsDynamic
						ConfigurationName = $Server.Configuration.RecoveryInterval.DisplayName #'recovery interval (min)'
						FriendlyName = 'Recovery Interval (minutes)'
						Description = $Server.Configuration.RecoveryInterval.Description
					}
					DataPath = if (($Server.Settings.DefaultFile).Length -gt 0) { 
						$Server.Settings.DefaultFile # System.String DefaultFile {get;set;}
					} else { 
						$Server.Information.MasterDBPath # System.String MasterDBPath {get;}
					}
					LogPath = if (($Server.Settings.DefaultLog).Length -gt 0) { 
						$Server.Settings.DefaultLog # System.String DefaultLog {get;set;}
					} else { 
						$Server.Information.MasterDBPath # System.String MasterDBLogPath {get;}
					}
					BackupPath = $Server.Settings.BackupDirectory # System.String BackupDirectory {get;set;}
				}
				Advanced = New-Object -TypeName psobject -Property @{

					# Contained databases available up starting with SQL 2012
					Containment = New-Object -TypeName psobject -Property @{
						EnableContainedDatabases = if ((($Server.Information.Version).CompareTo($SQLServer2012) -ge 0) -and ($SmoMajorVersion -ge 11)) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = if ($Server.Configuration.ContainmentEnabled.ConfigValue -gt 0) { $true } else { $false }
								RunningValue = if ($Server.Configuration.ContainmentEnabled.RunValue -gt 0) { $true } else { $false }
								DefaultValue = $false
								MinimumValue = $Server.Configuration.ContainmentEnabled.Minimum
								MaximumValue = $Server.Configuration.ContainmentEnabled.Maximum
								IsAdvanced = $Server.Configuration.ContainmentEnabled.IsAdvanced
								IsDynamic = $Server.Configuration.ContainmentEnabled.IsDynamic
								ConfigurationName = $Server.Configuration.ContainmentEnabled.DisplayName #'contained database authentication'
								FriendlyName = 'Enable Contained Databases'
								Description = $Server.Configuration.ContainmentEnabled.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'contained database authentication'
								FriendlyName = 'Enable Contained Databases'
								Description = if (($Server.Information.Version).CompareTo($SQLServer2012) -ge 0) {
									'This option is available starting with SQL Server 2012'
								} else {
									'SMO 2012 or higher is required to retreive this value'
								}
							}
						}
					}

					# Filestream available up starting with SQL 2008
					Filestream = if ((($Server.Information.Version).CompareTo($SQLServer2008) -ge 0) -and ($SmoMajorVersion -ge 10)) {
						New-Object -TypeName psobject -Property @{
							FilestreamAccessLevel = @{
								ConfiguredValue = switch ($Server.Configuration.FilestreamAccessLevel.ConfigValue) {
									0 { 'Disabled' }
									1 { 'Transact-SQL access enabled' }
									2 { 'Full access enabled' }
								}
								RunningValue = switch ($Server.Configuration.FilestreamAccessLevel.ConfigValue) {
									0 { 'Disabled' }
									1 { 'Transact-SQL access enabled' }
									2 { 'Full access enabled' }
								}
								DefaultValue = 'Disabled'
								MinimumValue = $Server.Configuration.FilestreamAccessLevel.Minimum
								MaximumValue = $Server.Configuration.FilestreamAccessLevel.Maximum
								IsAdvanced = $Server.Configuration.FilestreamAccessLevel.IsAdvanced
								IsDynamic = $Server.Configuration.FilestreamAccessLevel.IsDynamic
								ConfigurationName = $Server.Configuration.FilestreamAccessLevel.DisplayName #'filestream access level'
								FriendlyName = 'FILESTREAM Access Level'
								Description = $Server.Configuration.FilestreamAccessLevel.Description
							}
						}
					} else {
						New-Object -TypeName psobject -Property @{
							FilestreamAccessLevel = @{
								ConfiguredValue = 'Not Available'
								RunningValue = 'Not Available'
								DefaultValue = 'Not Available'
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'filestream access level'
								FriendlyName = 'FILESTREAM Access Level'
								Description = if (($Server.Information.Version).CompareTo($SQLServer2008) -ge 0) {
									'This option is available starting with SQL Server 2008'
								} else {
									'SMO 2008 or higher is required to retreive this value'
								}
							}
						}
					}

					# The following category is not exposed through the SSMS GUI but this is where I think they would be if they were				
					FullText = New-Object -TypeName psobject -Property @{
						FullTextCrawlBandwidthMin = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $Server.Configuration.FullTextCrawlBandwidthMin.ConfigValue
								RunningValue = $Server.Configuration.FullTextCrawlBandwidthMin.RunValue
								DefaultValue = 0
								MinimumValue = $Server.Configuration.FullTextCrawlBandwidthMin.Minimum
								MaximumValue = $Server.Configuration.FullTextCrawlBandwidthMin.Maximum
								IsAdvanced = $Server.Configuration.FullTextCrawlBandwidthMin.IsAdvanced
								IsDynamic = $Server.Configuration.FullTextCrawlBandwidthMin.IsDynamic
								ConfigurationName = $Server.Configuration.FullTextCrawlBandwidthMin.DisplayName #'ft crawl bandwidth (min)'
								FriendlyName = 'Full-Text Crawl Bandwidth Min'
								Description = $Server.Configuration.FullTextCrawlBandwidthMin.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'ft crawl bandwidth (min)'
								FriendlyName = 'Full-Text Crawl Bandwidth Min'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						FullTextCrawlBandwidthMax = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $Server.Configuration.FullTextCrawlBandwidthMax.ConfigValue
								RunningValue = $Server.Configuration.FullTextCrawlBandwidthMax.RunValue
								DefaultValue = 100
								MinimumValue = $Server.Configuration.FullTextCrawlBandwidthMax.Minimum
								MaximumValue = $Server.Configuration.FullTextCrawlBandwidthMax.Maximum
								IsAdvanced = $Server.Configuration.FullTextCrawlBandwidthMax.IsAdvanced
								IsDynamic = $Server.Configuration.FullTextCrawlBandwidthMax.IsDynamic
								ConfigurationName = $Server.Configuration.FullTextCrawlBandwidthMax.DisplayName #'ft crawl bandwidth (max)'
								FriendlyName = 'Full-Text Crawl Bandwidth Max'
								Description = $Server.Configuration.FullTextCrawlBandwidthMax.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'ft crawl bandwidth (max)'
								FriendlyName = 'Full-Text Crawl Bandwidth Max'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						FullTextCrawlRangeMax = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $Server.Configuration.FullTextCrawlRangeMax.ConfigValue
								RunningValue = $Server.Configuration.FullTextCrawlRangeMax.RunValue
								DefaultValue = 4
								MinimumValue = $Server.Configuration.FullTextCrawlRangeMax.Minimum
								MaximumValue = $Server.Configuration.FullTextCrawlRangeMax.Maximum
								IsAdvanced = $Server.Configuration.FullTextCrawlRangeMax.IsAdvanced
								IsDynamic = $Server.Configuration.FullTextCrawlRangeMax.IsDynamic
								ConfigurationName = $Server.Configuration.FullTextCrawlRangeMax.DisplayName #'max full-text crawl range'
								FriendlyName = 'Full-Text Crawl Range Max'
								Description = $Server.Configuration.FullTextCrawlRangeMax.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'max full-text crawl range'
								FriendlyName = 'Full-Text Crawl Range Max'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						FullTextNotifyBandwidthMin = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $Server.Configuration.FullTextNotifyBandwidthMin.ConfigValue
								RunningValue = $Server.Configuration.FullTextNotifyBandwidthMin.RunValue
								DefaultValue = 0
								MinimumValue = $Server.Configuration.FullTextNotifyBandwidthMin.Minimum
								MaximumValue = $Server.Configuration.FullTextNotifyBandwidthMin.Maximum
								IsAdvanced = $Server.Configuration.FullTextNotifyBandwidthMin.IsAdvanced
								IsDynamic = $Server.Configuration.FullTextNotifyBandwidthMin.IsDynamic
								ConfigurationName = $Server.Configuration.FullTextNotifyBandwidthMin.DisplayName #'ft notify bandwidth (min)'
								FriendlyName = 'Full-Text Notify Bandwidth Min'
								Description = $Server.Configuration.FullTextNotifyBandwidthMin.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'ft notify bandwidth (min)'
								FriendlyName = 'Full-Text Notify Bandwidth Min'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						FullTextNotifyBandwidthMax = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $Server.Configuration.FullTextNotifyBandwidthMax.ConfigValue
								RunningValue = $Server.Configuration.FullTextNotifyBandwidthMax.RunValue
								DefaultValue = 100
								MinimumValue = $Server.Configuration.FullTextNotifyBandwidthMax.Minimum
								MaximumValue = $Server.Configuration.FullTextNotifyBandwidthMax.Maximum
								IsAdvanced = $Server.Configuration.FullTextNotifyBandwidthMax.IsAdvanced
								IsDynamic = $Server.Configuration.FullTextNotifyBandwidthMax.IsDynamic
								ConfigurationName = $Server.Configuration.FullTextNotifyBandwidthMax.DisplayName #'ft notify bandwidth (max)'
								FriendlyName = 'Full-Text Notify Bandwidth Max'
								Description = $Server.Configuration.FullTextNotifyBandwidthMax.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'ft notify bandwidth (max)'
								FriendlyName = 'Full-Text Notify Bandwidth Max'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						PrecomputeRank = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = if ($Server.Configuration.PrecomputeRank.ConfigValue -gt 0) { $true } else { $false }
								RunningValue = if ($Server.Configuration.PrecomputeRank.RunValue -gt 0) { $true } else { $false }
								DefaultValue = $false
								MinimumValue = $Server.Configuration.PrecomputeRank.Minimum
								MaximumValue = $Server.Configuration.PrecomputeRank.Maximum
								IsAdvanced = $Server.Configuration.PrecomputeRank.IsAdvanced
								IsDynamic = $Server.Configuration.PrecomputeRank.IsDynamic
								ConfigurationName = $Server.Configuration.PrecomputeRank.DisplayName #'precompute rank'
								FriendlyName = 'Full-Text Precompute Rank'
								Description = $Server.Configuration.PrecomputeRank.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'precompute rank'
								FriendlyName = 'Full-Text Precompute Rank'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						ProtocolHandlerTimeout = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $Server.Configuration.ProtocolHandlerTimeout.ConfigValue
								RunningValue = $Server.Configuration.ProtocolHandlerTimeout.RunValue
								DefaultValue = 60
								MinimumValue = $Server.Configuration.ProtocolHandlerTimeout.Minimum
								MaximumValue = $Server.Configuration.ProtocolHandlerTimeout.Maximum
								IsAdvanced = $Server.Configuration.ProtocolHandlerTimeout.IsAdvanced
								IsDynamic = $Server.Configuration.ProtocolHandlerTimeout.IsDynamic
								ConfigurationName = $Server.Configuration.ProtocolHandlerTimeout.DisplayName #'PH timeout (s)'
								FriendlyName = 'Full-Text Protocol Handler Timeout'
								Description = $Server.Configuration.ProtocolHandlerTimeout.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'PH timeout (s)'
								FriendlyName = 'Full-Text Protocol Handler Timeout'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						TransformNoiseWords = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = if ($Server.Configuration.TransformNoiseWords.ConfigValue -gt 0) { $true } else { $false }
								RunningValue = if ($Server.Configuration.TransformNoiseWords.RunValue -gt 0) { $true } else { $false }
								DefaultValue = $false
								MinimumValue = $Server.Configuration.TransformNoiseWords.Minimum
								MaximumValue = $Server.Configuration.TransformNoiseWords.Maximum
								IsAdvanced = $Server.Configuration.TransformNoiseWords.IsAdvanced
								IsDynamic = $Server.Configuration.TransformNoiseWords.IsDynamic
								ConfigurationName = $Server.Configuration.TransformNoiseWords.DisplayName #'transform noise words'
								FriendlyName = 'Full-Text Transform Noise Words'
								Description = $Server.Configuration.TransformNoiseWords.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'transform noise words'
								FriendlyName = 'Full-Text Transform Noise Words'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
					}

					Miscellaneous = New-Object -TypeName psobject -Property @{
						AllowTriggersToFireOthers = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if ($Server.Configuration.NestedTriggers.ConfigValue -gt 0) { $true } else { $false }
							RunningValue = if ($Server.Configuration.NestedTriggers.RunValue -gt 0) { $true } else { $false }
							DefaultValue = $true
							MinimumValue = $Server.Configuration.NestedTriggers.Minimum
							MaximumValue = $Server.Configuration.NestedTriggers.Maximum
							IsAdvanced = $Server.Configuration.NestedTriggers.IsAdvanced
							IsDynamic = $Server.Configuration.NestedTriggers.IsDynamic
							ConfigurationName = $Server.Configuration.NestedTriggers.DisplayName #'nested triggers'
							FriendlyName = 'Allow Triggers To Fire Others'
							Description = $Server.Configuration.NestedTriggers.Description
						}
						AllowUpdates = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if ($Server.Configuration.AllowUpdates.ConfigValue -gt 0) { $true } else { $false }
							RunningValue = if ($Server.Configuration.AllowUpdates.RunValue -gt 0) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $Server.Configuration.AllowUpdates.Minimum
							MaximumValue = $Server.Configuration.AllowUpdates.Maximum
							IsAdvanced = $Server.Configuration.AllowUpdates.IsAdvanced
							IsDynamic = $Server.Configuration.AllowUpdates.IsDynamic
							ConfigurationName = $Server.Configuration.AllowUpdates.DisplayName #'allow updates'
							FriendlyName = 'Allow Updates To System Tables'
							Description = $Server.Configuration.AllowUpdates.Description
						}
						BlockedProcessThresholdSeconds = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $Server.Configuration.BlockedProcessThreshold.ConfigValue
								RunningValue = $Server.Configuration.BlockedProcessThreshold.RunValue
								DefaultValue = 0
								MinimumValue = $Server.Configuration.BlockedProcessThreshold.Minimum
								MaximumValue = $Server.Configuration.BlockedProcessThreshold.Maximum
								IsAdvanced = $Server.Configuration.BlockedProcessThreshold.IsAdvanced
								IsDynamic = $Server.Configuration.BlockedProcessThreshold.IsDynamic
								ConfigurationName = $Server.Configuration.BlockedProcessThreshold.DisplayName #'blocked process threshold (s)'
								FriendlyName = 'Blocked Process Threshold (sec)'
								Description = $Server.Configuration.BlockedProcessThreshold.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'blocked process threshold (s)'
								FriendlyName = 'Blocked Process Threshold (sec)'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						CursorThreshold = New-Object -TypeName psobject -Property @{
							ConfiguredValue = $Server.Configuration.CursorThreshold.ConfigValue
							RunningValue = $Server.Configuration.CursorThreshold.RunValue
							DefaultValue = -1
							MinimumValue = $Server.Configuration.CursorThreshold.Minimum
							MaximumValue = $Server.Configuration.CursorThreshold.Maximum
							IsAdvanced = $Server.Configuration.CursorThreshold.IsAdvanced
							IsDynamic = $Server.Configuration.CursorThreshold.IsDynamic
							ConfigurationName = $Server.Configuration.CursorThreshold.DisplayName #'cursor threshold'
							FriendlyName = 'Cursor Threshold'
							Description = $Server.Configuration.CursorThreshold.Description
						}
						DefaultFullTextLanguage = New-Object -TypeName psobject -Property @{
							ConfiguredValue = [String](Get-FullTextLanguageValue -Language $Server.Configuration.DefaultFullTextLanguage.ConfigValue)
							RunningValue = [String](Get-FullTextLanguageValue -Language $Server.Configuration.DefaultFullTextLanguage.RunValue)
							DefaultValue = [String](Get-FullTextLanguageValue -Language 1033)
							MinimumValue = $Server.Configuration.DefaultFullTextLanguage.Minimum
							MaximumValue = $Server.Configuration.DefaultFullTextLanguage.Maximum
							IsAdvanced = $Server.Configuration.DefaultFullTextLanguage.IsAdvanced
							IsDynamic = $Server.Configuration.DefaultFullTextLanguage.IsDynamic
							ConfigurationName = $Server.Configuration.DefaultFullTextLanguage.DisplayName #'default full-text language'
							FriendlyName = 'Default Full-Text Language'
							Description = $Server.Configuration.DefaultFullTextLanguage.Description
						}
						DefaultLanguage = New-Object -TypeName psobject -Property @{
							ConfiguredValue = [String](Get-LanguageValue -Language $Server.Configuration.DefaultLanguage.ConfigValue)
							RunningValue = [String](Get-LanguageValue -Language $Server.Configuration.DefaultLanguage.RunValue)
							DefaultValue = [String](Get-LanguageValue -Language 0) # 0 (English) (this comes from sys.syslanguages)
							MinimumValue = $Server.Configuration.DefaultLanguage.Minimum
							MaximumValue = $Server.Configuration.DefaultLanguage.Maximum
							IsAdvanced = $Server.Configuration.DefaultLanguage.IsAdvanced
							IsDynamic = $Server.Configuration.DefaultLanguage.IsDynamic
							ConfigurationName = $Server.Configuration.DefaultLanguage.DisplayName #'default language'
							FriendlyName = 'Default Language'
							Description = $Server.Configuration.DefaultLanguage.Description
						}
						FullTextUpgradeOption = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = [String](Get-FullTextCatalogUpgradeOptionValue -FullTextCatalogUpgradeOption $Server.FullTextService.CatalogUpgradeOption)
								RunningValue = [String](Get-FullTextCatalogUpgradeOptionValue -FullTextCatalogUpgradeOption $Server.FullTextService.CatalogUpgradeOption)
								DefaultValue = 'Rebuild'
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = [String]::Empty
								FriendlyName = 'Full-Text Upgrade Option'
								Description = $null
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = [String]::Empty
								FriendlyName = 'Full-Text Upgrade Option'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						MaxTextReplicationSize = New-Object -TypeName psobject -Property @{
							ConfiguredValue = $Server.Configuration.ReplicationMaxTextSize.ConfigValue
							RunningValue = $Server.Configuration.ReplicationMaxTextSize.RunValue
							DefaultValue = 65536
							MinimumValue = $Server.Configuration.ReplicationMaxTextSize.Minimum
							MaximumValue = $Server.Configuration.ReplicationMaxTextSize.Maximum
							IsAdvanced = $Server.Configuration.ReplicationMaxTextSize.IsAdvanced
							IsDynamic = $Server.Configuration.ReplicationMaxTextSize.IsDynamic
							ConfigurationName = $Server.Configuration.ReplicationMaxTextSize.DisplayName #'max text repl size (B)'
							FriendlyName = 'Max Text Replication Size (in bytes)'
							Description = $Server.Configuration.ReplicationMaxTextSize.Description
						}
						OpenObjects = New-Object -TypeName psobject -Property @{
							ConfiguredValue = $Server.Configuration.OpenObjects.ConfigValue
							RunningValue = $Server.Configuration.OpenObjects.RunValue
							DefaultValue = 0
							MinimumValue = $Server.Configuration.OpenObjects.Minimum
							MaximumValue = $Server.Configuration.OpenObjects.Maximum
							IsAdvanced = $Server.Configuration.OpenObjects.IsAdvanced
							IsDynamic = $Server.Configuration.OpenObjects.IsDynamic
							ConfigurationName = $Server.Configuration.OpenObjects.DisplayName #'open objects'
							FriendlyName = 'Open database objects'
							Description = $Server.Configuration.OpenObjects.Description
						}
						OptimizeForAdHocWorkloads = if ((($Server.Information.Version).CompareTo($SQLServer2008) -ge 0) -and ($SmoMajorVersion -ge 10)) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = if ($Server.Configuration.OptimizeAdhocWorkloads.ConfigValue -gt 0) { $true } else { $false }
								RunningValue = if ($Server.Configuration.OptimizeAdhocWorkloads.RunValue -gt 0) { $true } else { $false }
								DefaultValue = $false
								MinimumValue = $Server.Configuration.OptimizeAdhocWorkloads.Minimum
								MaximumValue = $Server.Configuration.OptimizeAdhocWorkloads.Maximum
								IsAdvanced = $Server.Configuration.OptimizeAdhocWorkloads.IsAdvanced
								IsDynamic = $Server.Configuration.OptimizeAdhocWorkloads.IsDynamic
								ConfigurationName = $Server.Configuration.OptimizeAdhocWorkloads.DisplayName #'optimize for ad hoc workloads'
								FriendlyName = 'Optimize for Ad hoc Workloads'
								Description = $Server.Configuration.OptimizeAdhocWorkloads.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'optimize for ad hoc workloads'
								FriendlyName = 'Optimize for Ad hoc Workloads'
								Description = if (($Server.Information.Version).CompareTo($SQLServer2008) -ge 0) {
									'This option is available starting with SQL Server 2008'
								} else {
									'SMO 2008 or higher is required to retreive this value'
								}
							}
						}
						ScanForStartupProcs = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if ($Server.Configuration.ScanForStartupProcedures.ConfigValue -gt 0) { $true } else { $false }
							RunningValue = if ($Server.Configuration.ScanForStartupProcedures.RunValue -gt 0) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $Server.Configuration.ScanForStartupProcedures.Minimum
							MaximumValue = $Server.Configuration.ScanForStartupProcedures.Maximum
							IsAdvanced = $Server.Configuration.ScanForStartupProcedures.IsAdvanced
							IsDynamic = $Server.Configuration.ScanForStartupProcedures.IsDynamic
							ConfigurationName = $Server.Configuration.ScanForStartupProcedures.DisplayName #'scan for startup procs'
							FriendlyName = 'Scan for Startup Procs'
							Description = $Server.Configuration.ScanForStartupProcedures.Description
						}
						TwoDigitYearCutoff = New-Object -TypeName psobject -Property @{
							ConfiguredValue = $Server.Configuration.TwoDigitYearCutoff.ConfigValue
							RunningValue = $Server.Configuration.TwoDigitYearCutoff.RunValue
							DefaultValue = 2049
							MinimumValue = $Server.Configuration.TwoDigitYearCutoff.Minimum
							MaximumValue = $Server.Configuration.TwoDigitYearCutoff.Maximum
							IsAdvanced = $Server.Configuration.TwoDigitYearCutoff.IsAdvanced
							IsDynamic = $Server.Configuration.TwoDigitYearCutoff.IsDynamic
							ConfigurationName = $Server.Configuration.TwoDigitYearCutoff.DisplayName #'two digit year cutoff'
							FriendlyName = 'Two Digit Year Cutoff'
							Description = $Server.Configuration.TwoDigitYearCutoff.Description
						}


						# The following options are not exposed through the SSMS GUI but this is where I think they would be if they were
						ClrEnabled = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = if ($Server.Configuration.IsSqlClrEnabled.ConfigValue -gt 0) { $true } else { $false }
								RunningValue = if ($Server.Configuration.IsSqlClrEnabled.RunValue -gt 0) { $true } else { $false }
								DefaultValue = $false
								MinimumValue = $Server.Configuration.IsSqlClrEnabled.Minimum
								MaximumValue = $Server.Configuration.IsSqlClrEnabled.Maximum
								IsAdvanced = $Server.Configuration.IsSqlClrEnabled.IsAdvanced
								IsDynamic = $Server.Configuration.IsSqlClrEnabled.IsDynamic
								ConfigurationName = $Server.Configuration.IsSqlClrEnabled.DisplayName #'clr enabled'
								FriendlyName = 'CLR Enabled'
								Description = $Server.Configuration.IsSqlClrEnabled.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'clr enabled'
								FriendlyName = 'CLR Enabled'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						DatabaseMailXPsEnabled = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = if ($Server.Configuration.DatabaseMailEnabled.ConfigValue -gt 0) { $true } else { $false }
								RunningValue = if ($Server.Configuration.DatabaseMailEnabled.RunValue -gt 0) { $true } else { $false }
								DefaultValue = $false
								MinimumValue = $Server.Configuration.DatabaseMailEnabled.Minimum
								MaximumValue = $Server.Configuration.DatabaseMailEnabled.Maximum
								IsAdvanced = $Server.Configuration.DatabaseMailEnabled.IsAdvanced
								IsDynamic = $Server.Configuration.DatabaseMailEnabled.IsDynamic
								ConfigurationName = $Server.Configuration.DatabaseMailEnabled.DisplayName #'Database Mail XPs'
								FriendlyName = 'Database Mail XPs Enabled'
								Description = $Server.Configuration.DatabaseMailEnabled.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'Database Mail XPs'
								FriendlyName = 'Database Mail XPs Enabled'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						DefaultTraceEnabled = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = if ($Server.Configuration.DefaultTraceEnabled.ConfigValue -gt 0) { $true } else { $false }
								RunningValue = if ($Server.Configuration.DefaultTraceEnabled.RunValue -gt 0) { $true } else { $false }
								DefaultValue = $true
								MinimumValue = $Server.Configuration.DefaultTraceEnabled.Minimum
								MaximumValue = $Server.Configuration.DefaultTraceEnabled.Maximum
								IsAdvanced = $Server.Configuration.DefaultTraceEnabled.IsAdvanced
								IsDynamic = $Server.Configuration.DefaultTraceEnabled.IsDynamic
								ConfigurationName = $Server.Configuration.DefaultTraceEnabled.DisplayName #'default trace enabled'
								FriendlyName = 'Default Trace Enabled'
								Description = $Server.Configuration.DefaultTraceEnabled.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'default trace enabled'
								FriendlyName = 'Default Trace Enabled'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						DisallowResultsFromTriggers = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = if ($Server.Configuration.DisallowResultsFromTriggers.ConfigValue -gt 0) { $true } else { $false }
								RunningValue = if ($Server.Configuration.DisallowResultsFromTriggers.RunValue -gt 0) { $true } else { $false }
								DefaultValue = $false
								MinimumValue = $Server.Configuration.DisallowResultsFromTriggers.Minimum
								MaximumValue = $Server.Configuration.DisallowResultsFromTriggers.Maximum
								IsAdvanced = $Server.Configuration.DisallowResultsFromTriggers.IsAdvanced
								IsDynamic = $Server.Configuration.DisallowResultsFromTriggers.IsDynamic
								ConfigurationName = $Server.Configuration.DisallowResultsFromTriggers.DisplayName #'disallow results from triggers'
								FriendlyName = 'Disallow Results From Triggers'
								Description = $Server.Configuration.DisallowResultsFromTriggers.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'disallow results from triggers'
								FriendlyName = 'Disallow Results From Triggers'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						ExtensibleKeyManagementEnabled = if ((($Server.Information.Version).CompareTo($SQLServer2008) -ge 0) -and ($SmoMajorVersion -ge 10)) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = if ($Server.Configuration.ExtensibleKeyManagementEnabled.ConfigValue -gt 0) { $true } else { $false }
								RunningValue = if ($Server.Configuration.ExtensibleKeyManagementEnabled.RunValue -gt 0) { $true } else { $false }
								DefaultValue = $false
								MinimumValue = $Server.Configuration.ExtensibleKeyManagementEnabled.Minimum
								MaximumValue = $Server.Configuration.ExtensibleKeyManagementEnabled.Maximum
								IsAdvanced = $Server.Configuration.ExtensibleKeyManagementEnabled.IsAdvanced
								IsDynamic = $Server.Configuration.ExtensibleKeyManagementEnabled.IsDynamic
								ConfigurationName = $Server.Configuration.ExtensibleKeyManagementEnabled.DisplayName #'EKM provider enabled'
								FriendlyName = 'Extensible Key Management Enabled'
								Description = $Server.Configuration.ExtensibleKeyManagementEnabled.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'EKM provider enabled'
								FriendlyName = 'Extensible Key Management Enabled'
								Description = if (($Server.Information.Version).CompareTo($SQLServer2008) -ge 0) {
									'This option is available starting with SQL Server 2008'
								} else {
									'SMO 2008 or higher is required to retreive this value'
								}
							}
						}
						InDoubtTransactionResolution = if ((($Server.Information.Version).CompareTo($SQLServer2008) -ge 0) -and ($SmoMajorVersion -ge 10)) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = switch ($Server.Configuration.InDoubtTransactionResolution.ConfigValue) {
									0 { '0 (No presumption)' }
									1 { '1 (Presume commit)' }
									2 { '2 (Presume abort)' }
								}
								RunningValue = switch ($Server.Configuration.InDoubtTransactionResolution.RunValue) {
									0 { '0 (No presumption)' }
									1 { '1 (Presume commit)' }
									2 { '2 (Presume abort)' }
								}
								DefaultValue = '0 (No presumption)'
								MinimumValue = $Server.Configuration.InDoubtTransactionResolution.Minimum
								MaximumValue = $Server.Configuration.InDoubtTransactionResolution.Maximum
								IsAdvanced = $Server.Configuration.InDoubtTransactionResolution.IsAdvanced
								IsDynamic = $Server.Configuration.InDoubtTransactionResolution.IsDynamic
								ConfigurationName = $Server.Configuration.InDoubtTransactionResolution.DisplayName #'in-doubt xact resolution'
								FriendlyName = 'In Doubt Transaction Resolution'
								Description = $Server.Configuration.InDoubtTransactionResolution.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'in-doubt xact resolution'
								FriendlyName = 'In Doubt Transaction Resolution'
								Description = if (($Server.Information.Version).CompareTo($SQLServer2008) -ge 0) {
									'This option is available starting with SQL Server 2008'
								} else {
									'SMO 2008 or higher is required to retreive this value'
								}
							}
						}
						OleAutomationProceduresEnabled = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = if ($Server.Configuration.OleAutomationProceduresEnabled.ConfigValue -gt 0) { $true } else { $false }
								RunningValue = if ($Server.Configuration.OleAutomationProceduresEnabled.RunValue -gt 0) { $true } else { $false }
								DefaultValue = $false
								MinimumValue = $Server.Configuration.OleAutomationProceduresEnabled.Minimum
								MaximumValue = $Server.Configuration.OleAutomationProceduresEnabled.Maximum
								IsAdvanced = $Server.Configuration.OleAutomationProceduresEnabled.IsAdvanced
								IsDynamic = $Server.Configuration.OleAutomationProceduresEnabled.IsDynamic
								ConfigurationName = $Server.Configuration.OleAutomationProceduresEnabled.DisplayName #'Ole Automation Procedures'
								FriendlyName = 'OLE Automation Procs Enabled'
								Description = $Server.Configuration.OleAutomationProceduresEnabled.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'Ole Automation Procedures'
								FriendlyName = 'OLE Automation Procs Enabled'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						ReplicationXPsEnabled = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = if ($Server.Configuration.ReplicationXPsEnabled.ConfigValue -gt 0) { $true } else { $false }
								RunningValue = if ($Server.Configuration.ReplicationXPsEnabled.RunValue -gt 0) { $true } else { $false }
								DefaultValue = $false
								MinimumValue = $Server.Configuration.ReplicationXPsEnabled.Minimum
								MaximumValue = $Server.Configuration.ReplicationXPsEnabled.Maximum
								IsAdvanced = $Server.Configuration.ReplicationXPsEnabled.IsAdvanced
								IsDynamic = $Server.Configuration.ReplicationXPsEnabled.IsDynamic
								ConfigurationName = $Server.Configuration.ReplicationXPsEnabled.DisplayName #'Replication XPs'
								FriendlyName = 'Replication XPs Enabled'
								Description = $Server.Configuration.ReplicationXPsEnabled.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'Replication XPs'
								FriendlyName = 'Replication XPs Enabled'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						SqlAgentXPsEnabled = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = if ($Server.Configuration.AgentXPsEnabled.ConfigValue -gt 0) { $true } else { $false }
								RunningValue = if ($Server.Configuration.AgentXPsEnabled.RunValue -gt 0) { $true } else { $false }
								DefaultValue = $false
								MinimumValue = $Server.Configuration.AgentXPsEnabled.Minimum
								MaximumValue = $Server.Configuration.AgentXPsEnabled.Maximum
								IsAdvanced = $Server.Configuration.AgentXPsEnabled.IsAdvanced
								IsDynamic = $Server.Configuration.AgentXPsEnabled.IsDynamic
								ConfigurationName = $Server.Configuration.AgentXPsEnabled.DisplayName #'Agent XPs'
								FriendlyName = 'SQL Agent XPs Enabled'
								Description = $Server.Configuration.AgentXPsEnabled.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'Agent XPs'
								FriendlyName = 'SQL Agent XPs Enabled'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						SqlMailXPsEnabled = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = if ($Server.Configuration.SqlMailXPsEnabled.ConfigValue -gt 0) { $true } else { $false }
								RunningValue = if ($Server.Configuration.SqlMailXPsEnabled.RunValue -gt 0) { $true } else { $false }
								DefaultValue = $false
								MinimumValue = $Server.Configuration.SqlMailXPsEnabled.Minimum
								MaximumValue = $Server.Configuration.SqlMailXPsEnabled.Maximum
								IsAdvanced = $Server.Configuration.SqlMailXPsEnabled.IsAdvanced
								IsDynamic = $Server.Configuration.SqlMailXPsEnabled.IsDynamic
								ConfigurationName = $Server.Configuration.SqlMailXPsEnabled.DisplayName #'SQL Mail XPs'
								FriendlyName = 'SQL Mail XPs Enabled'
								Description = $Server.Configuration.SqlMailXPsEnabled.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'SQL Mail XPs'
								FriendlyName = 'SQL Mail XPs Enabled'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						ServerTriggerRecursionEnabled = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = if ($Server.Configuration.ServerTriggerRecursionEnabled.ConfigValue -gt 0) { $true } else { $false }
								RunningValue = if ($Server.Configuration.ServerTriggerRecursionEnabled.RunValue -gt 0) { $true } else { $false }
								DefaultValue = $true
								MinimumValue = $Server.Configuration.ServerTriggerRecursionEnabled.Minimum
								MaximumValue = $Server.Configuration.ServerTriggerRecursionEnabled.Maximum
								IsAdvanced = $Server.Configuration.ServerTriggerRecursionEnabled.IsAdvanced
								IsDynamic = $Server.Configuration.ServerTriggerRecursionEnabled.IsDynamic
								ConfigurationName = $Server.Configuration.ServerTriggerRecursionEnabled.DisplayName #'server trigger recursion'
								FriendlyName = 'Server Trigger Recursion Enabled'
								Description = $Server.Configuration.ServerTriggerRecursionEnabled.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'server trigger recursion'
								FriendlyName = 'Server Trigger Recursion Enabled'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}
						ShowAdvancedOptions = New-Object -TypeName psobject -Property @{
							ConfiguredValue = if ($Server.Configuration.ShowAdvancedOptions.ConfigValue -gt 0) { $true } else { $false }
							RunningValue = if ($Server.Configuration.ShowAdvancedOptions.RunValue -gt 0) { $true } else { $false }
							DefaultValue = $false
							MinimumValue = $Server.Configuration.ShowAdvancedOptions.Minimum
							MaximumValue = $Server.Configuration.ShowAdvancedOptions.Maximum
							IsAdvanced = $Server.Configuration.ShowAdvancedOptions.IsAdvanced
							IsDynamic = $Server.Configuration.ShowAdvancedOptions.IsDynamic
							ConfigurationName = $Server.Configuration.ShowAdvancedOptions.DisplayName #'show advanced options'
							FriendlyName = 'Show Advanced Options'
							Description = $Server.Configuration.ShowAdvancedOptions.Description
						}
						SmoAndDmoXPsEnabled = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = if ($Server.Configuration.SmoAndDmoXPsEnabled.ConfigValue -gt 0) { $true } else { $false }
								RunningValue = if ($Server.Configuration.SmoAndDmoXPsEnabled.RunValue -gt 0) { $true } else { $false }
								DefaultValue = $true
								MinimumValue = $Server.Configuration.SmoAndDmoXPsEnabled.Minimum
								MaximumValue = $Server.Configuration.SmoAndDmoXPsEnabled.Maximum
								IsAdvanced = $Server.Configuration.SmoAndDmoXPsEnabled.IsAdvanced
								IsDynamic = $Server.Configuration.SmoAndDmoXPsEnabled.IsDynamic
								ConfigurationName = $Server.Configuration.SmoAndDmoXPsEnabled.DisplayName #'SMO and DMO XPs'
								FriendlyName = 'SMO & DMO XPs Enabled'
								Description = $Server.Configuration.SmoAndDmoXPsEnabled.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'SMO and DMO XPs'
								FriendlyName = 'SMO & DMO XPs Enabled'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}

						# Web Assistant Procedures only available in SQL 2005 
						WebAssistantProceduresEnabled = if ((($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) -and (($Server.Information.Version).CompareTo($SQLServer2008) -lt 0)) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = if ($Server.Configuration.WebXPsEnabled.ConfigValue -gt 0) { $true } else { $false }
								RunningValue = if ($Server.Configuration.WebXPsEnabled.RunValue -gt 0) { $true } else { $false }
								DefaultValue = $false
								MinimumValue = $Server.Configuration.WebXPsEnabled.Minimum
								MaximumValue = $Server.Configuration.WebXPsEnabled.Maximum
								IsAdvanced = $Server.Configuration.WebXPsEnabled.IsAdvanced
								IsDynamic = $Server.Configuration.WebXPsEnabled.IsDynamic
								ConfigurationName = $Server.Configuration.WebXPsEnabled.DisplayName #'Web Assistant Procedures'
								FriendlyName = 'Web Assistant Procs Enabled'
								Description = $Server.Configuration.WebXPsEnabled.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'Web Assistant Procedures'
								FriendlyName = 'Web Assistant Procs Enabled'
								Description = 'This option is only available in SQL Server 2005'
							}
						}

						XPCmdShellEnabled = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = if ($Server.Configuration.XPCmdShellEnabled.ConfigValue -gt 0) { $true } else { $false }
								RunningValue = if ($Server.Configuration.XPCmdShellEnabled.RunValue -gt 0) { $true } else { $false }
								DefaultValue = $false
								MinimumValue = $Server.Configuration.XPCmdShellEnabled.Minimum
								MaximumValue = $Server.Configuration.XPCmdShellEnabled.Maximum
								IsAdvanced = $Server.Configuration.XPCmdShellEnabled.IsAdvanced
								IsDynamic = $Server.Configuration.XPCmdShellEnabled.IsDynamic
								ConfigurationName = $Server.Configuration.XPCmdShellEnabled.DisplayName #'xp_cmdshell'
								FriendlyName = 'xp_cmdshell Enabled'
								Description = $Server.Configuration.XPCmdShellEnabled.Description
							}
						} else {
							New-Object -TypeName psobject -Property @{
								ConfiguredValue = $null
								RunningValue = $null
								DefaultValue = $null
								MinimumValue = $null
								MaximumValue = $null
								IsAdvanced = $null
								IsDynamic = $null
								ConfigurationName = 'xp_cmdshell'
								FriendlyName = 'xp_cmdshell Enabled'
								Description = 'This option is available starting with SQL Server 2005'
							}
						}

						# Exposed only in Express edition. I don't really care about these.
						#UserInstancesEnabled
						#UserInstanceTimeout

					}
					Network = New-Object -TypeName psobject -Property @{
						NetworkPacketSize = New-Object -TypeName psobject -Property @{
							ConfiguredValue = $Server.Configuration.NetworkPacketSize.ConfigValue
							RunningValue = $Server.Configuration.NetworkPacketSize.RunValue
							DefaultValue = 4096
							MinimumValue = $Server.Configuration.NetworkPacketSize.Minimum
							MaximumValue = $Server.Configuration.NetworkPacketSize.Maximum
							IsAdvanced = $Server.Configuration.NetworkPacketSize.IsAdvanced
							IsDynamic = $Server.Configuration.NetworkPacketSize.IsDynamic
							ConfigurationName = $Server.Configuration.NetworkPacketSize.DisplayName #'network packet size (B)'
							FriendlyName = 'Network Packet Size'
							Description = $Server.Configuration.NetworkPacketSize.Description
						}
						RemoteLoginTimeoutSeconds = New-Object -TypeName psobject -Property @{
							ConfiguredValue = $Server.Configuration.RemoteLoginTimeout.ConfigValue
							RunningValue = $Server.Configuration.RemoteLoginTimeout.RunValue
							DefaultValue = if (($Server.Information.Version).CompareTo($SQLServer2012) -ge 0) { 10 } else { 20 }
							MinimumValue = $Server.Configuration.RemoteLoginTimeout.Minimum
							MaximumValue = $Server.Configuration.RemoteLoginTimeout.Maximum
							IsAdvanced = $Server.Configuration.RemoteLoginTimeout.IsAdvanced
							IsDynamic = $Server.Configuration.RemoteLoginTimeout.IsDynamic
							ConfigurationName = $Server.Configuration.RemoteLoginTimeout.DisplayName #'remote login timeout (s)'
							FriendlyName = 'Remote Login Timeout (sec)'
							Description = $Server.Configuration.RemoteLoginTimeout.Description
						}
					}
					Parallelism = New-Object -TypeName psobject -Property @{
						CostThresholdForParallelism = New-Object -TypeName psobject -Property @{
							ConfiguredValue = $Server.Configuration.CostThresholdForParallelism.ConfigValue
							RunningValue = $Server.Configuration.CostThresholdForParallelism.RunValue
							DefaultValue = 5
							MinimumValue = $Server.Configuration.CostThresholdForParallelism.Minimum
							MaximumValue = $Server.Configuration.CostThresholdForParallelism.Maximum
							IsAdvanced = $Server.Configuration.CostThresholdForParallelism.IsAdvanced
							IsDynamic = $Server.Configuration.CostThresholdForParallelism.IsDynamic
							ConfigurationName = $Server.Configuration.CostThresholdForParallelism.DisplayName #'cost threshold for parallelism'
							FriendlyName = 'Cost Threshold for Parallelism'
							Description = $Server.Configuration.CostThresholdForParallelism.Description
						}
						Locks = New-Object -TypeName psobject -Property @{
							ConfiguredValue = $Server.Configuration.Locks.ConfigValue
							RunningValue = $Server.Configuration.Locks.RunValue
							DefaultValue = 0
							MinimumValue = $Server.Configuration.Locks.Minimum
							MaximumValue = $Server.Configuration.Locks.Maximum
							IsAdvanced = $Server.Configuration.Locks.IsAdvanced
							IsDynamic = $Server.Configuration.Locks.IsDynamic
							ConfigurationName = $Server.Configuration.Locks.DisplayName #'locks'
							FriendlyName = 'Locks'
							Description = $Server.Configuration.Locks.Description
						}
						MaxDegreeOfParallelism = New-Object -TypeName psobject -Property @{
							ConfiguredValue = $Server.Configuration.MaxDegreeOfParallelism.ConfigValue
							RunningValue = $Server.Configuration.MaxDegreeOfParallelism.RunValue
							DefaultValue = 0
							MinimumValue = $Server.Configuration.MaxDegreeOfParallelism.Minimum
							MaximumValue = $Server.Configuration.MaxDegreeOfParallelism.Maximum
							IsAdvanced = $Server.Configuration.MaxDegreeOfParallelism.IsAdvanced
							IsDynamic = $Server.Configuration.MaxDegreeOfParallelism.IsDynamic
							ConfigurationName = $Server.Configuration.MaxDegreeOfParallelism.DisplayName #'max degree of parallelism'
							FriendlyName = 'Max Degree of Parallelism'
							Description = $Server.Configuration.MaxDegreeOfParallelism.Description
						}
						QueryWaitSeconds = New-Object -TypeName psobject -Property @{
							ConfiguredValue = $Server.Configuration.QueryWait.ConfigValue
							RunningValue = $Server.Configuration.QueryWait.RunValue
							DefaultValue = -1
							MinimumValue = $Server.Configuration.QueryWait.Minimum
							MaximumValue = $Server.Configuration.QueryWait.Maximum
							IsAdvanced = $Server.Configuration.QueryWait.IsAdvanced
							IsDynamic = $Server.Configuration.QueryWait.IsDynamic
							ConfigurationName = $Server.Configuration.QueryWait.DisplayName #'query wait (s)'
							FriendlyName = 'Query Wait (sec)'
							Description = $Server.Configuration.QueryWait.Description
						}
					}
				}

				# This is not part of the SSMS GUI but these properties aren't exposed in a nice interface anywhere else
				# and if I designed SSMS I'd have made a page for them to go here
				HighAvailability = New-Object -TypeName psobject -Property @{
					FailoverCluster = New-Object -TypeName psobject -Property @{
						IsClusteredInstance = $Server.Information.IsClustered # System.Boolean IsClustered {get;}
						Member = if ($Server.Information.IsClustered) { 
							Get-FailoverClusterMemberList -Server $Server
						} else {
							$null
						}
						SharedDrive = if ($Server.Information.IsClustered) {
							if (
								$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0 -and
								$DbEngineType -ieq $StandaloneDbEngine # Doesn't work against Azure
							) {
								@() + (
									$Server.Databases['master'].ExecuteWithResults('SELECT DriveName FROM sys.dm_io_cluster_shared_drives').Tables[0].Rows | ForEach-Object {
										$_.DriveName
									}
								)
							} elseif (
								$($Server.Information.Version).CompareTo($SQLServer2000) -ge 0 -and
								$DbEngineType -ieq $StandaloneDbEngine # Doesn't work against Azure
							) {
								$Server.Databases['master'].ExecuteWithResults('SELECT DriveName FROM ::fn_servershareddrives()').Tables[0].Rows | ForEach-Object {
									$_.DriveName
								}
							}
							else {
								$null
							}
						} else {
							$null
						}
					}
					AlwaysOn = New-Object -TypeName psobject -Property @{
						IsAlwaysOnEnabled = $Server.Information.IsHadrEnabled # System.Boolean IsHadrEnabled {get;}
						AlwaysOnManagerStatus = [String](Get-HadrManagerStatusValue -HadrManagerStatus $Server.HadrManagerStatus) # Microsoft.SqlServer.Management.Smo.HadrManagerStatus HadrManagerStatus {get;}
						WindowsFailoverCluster = New-Object -TypeName psobject -Property @{
							Name = $Server.ClusterName # System.String ClusterName {get;}
							QuorumState = [String](Get-ClusterQuorumStateValue -ClusterQuorumState $Server.ClusterQuorumState) # Microsoft.SqlServer.Management.Smo.ClusterQuorumState ClusterQuorumState {get;}
							QuorumType = [String](Get-ClusterQuorumTypeValue -ClusterQuorumType $Server.ClusterQuorumType) # Microsoft.SqlServer.Management.Smo.ClusterQuorumType ClusterQuorumType {get;}
							Member = if (
								$Server.Information.IsHadrEnabled -and
								$($Server.Information.Version).CompareTo($SQLServer2012) -ge 0 -and 
								$SmoMajorVersion -ge 11 -and
								$DbEngineType -ieq $StandaloneDbEngine # Doesn't work against Azure
							) {
								@() + (
									$Server.EnumClusterMembersState() | ForEach-Object {
										New-Object -TypeName psobject -Property @{
											Name = $_.Name
											MemberState = switch ($_.member_state) {
												0 { 'Offline' }
												1 { 'Online' }
												default { 'Uknown' }
											}
											MemberType = switch ($_.MemberType) {
												0 { 'Cluster Node' }
												1 { 'Disk Witness' }
												1 { 'File Share Witness' }
												default { 'Uknown' }
											}
											NumberOfQuorumVotes = $_.NumberOfQuorumVotes
										}
									}
								)
							} else {
								$null
							}
							Subnet = if (
								$Server.Information.IsHadrEnabled -and
								$($Server.Information.Version).CompareTo($SQLServer2012) -ge 0 -and 
								$SmoMajorVersion -ge 11 -and
								$DbEngineType -ieq $StandaloneDbEngine # Doesn't work against Azure
							) {
								@() + (
									$Server.EnumClusterSubnets() | ForEach-Object { 
										New-Object -TypeName psobject -Property @{
											Name = $_.Name
											IsIPv4 = $_.IsIPv4
											SubnetIP = $_.SubnetIP
											SubnetIPv4Mask = $_.SubnetIPv4Mask
											SubnetPrefixLength = $_.SubnetPrefixLength
										}
									}
								)
							} else {
								$null
							}
						}
					}
				}

				Permissions = if (
					$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0 -and
					$DbEngineType -ieq $StandaloneDbEngine # Doesn't work against Azure
				) {
					@() + (
						$Server.EnumServerPermissions() | ForEach-Object {
							New-Object -TypeName psobject -Property @{
								ColumnName = $_.ColumnName # System.String ColumnName {get;}
								Grantee = $_.Grantee # System.String Grantee {get;}
								GranteeType = [String](Get-PrincipalTypeValue -PrincipalType $_.GranteeType) # Microsoft.SqlServer.Management.Smo.PrincipalType GranteeType {get;}
								Grantor = $_.Grantor # System.String Grantor {get;}
								GrantorType = [String](Get-PrincipalTypeValue -PrincipalType $_.GrantorType) # Microsoft.SqlServer.Management.Smo.PrincipalType GrantorType {get;}
								ObjectClass = [String](Get-ObjectClassValue -ObjectClass $_.ObjectClass) # Microsoft.SqlServer.Management.Smo.ObjectClass ObjectClass {get;}
								ObjectID = $_.ObjectID # System.Int32 ObjectID {get;}
								ObjectName = if (($_.ObjectName -ieq $Server.Name) -and ($ServerName -ine $Server.Name)) { $ServerName } else { $_.ObjectName } # System.String ObjectName {get;}
								ObjectSchema = $_.ObjectSchema # System.String ObjectSchema {get;}
								PermissionState = [String](Get-PermissionStateValue -PermissionState $_.PermissionState) # Microsoft.SqlServer.Management.Smo.PermissionState PermissionState {get;}
								PermissionType = $_.PermissionType.ToString() # Microsoft.SqlServer.Management.Smo.ServerPermissionSet PermissionType {get;}
							}
						}
					) + (
						$Server.EnumObjectPermissions() | ForEach-Object {
							New-Object -TypeName psobject -Property @{
								ColumnName = $_.ColumnName # System.String ColumnName {get;}
								Grantee = $_.Grantee # System.String Grantee {get;}
								GranteeType = [String](Get-PrincipalTypeValue -PrincipalType $_.GranteeType) # Microsoft.SqlServer.Management.Smo.PrincipalType GranteeType {get;}
								Grantor = $_.Grantor # System.String Grantor {get;}
								GrantorType = [String](Get-PrincipalTypeValue -PrincipalType $_.GrantorType) # Microsoft.SqlServer.Management.Smo.PrincipalType GrantorType {get;}
								ObjectClass = [String](Get-ObjectClassValue -ObjectClass $_.ObjectClass) # Microsoft.SqlServer.Management.Smo.ObjectClass ObjectClass {get;}
								ObjectID = $_.ObjectID # System.Int32 ObjectID {get;}
								ObjectName = $_.ObjectName # System.String ObjectName {get;}
								ObjectSchema = $_.ObjectSchema # System.String ObjectSchema {get;}
								PermissionState = [String](Get-PermissionStateValue -PermissionState $_.PermissionState) # Microsoft.SqlServer.Management.Smo.PermissionState PermissionState {get;}
								PermissionType = $_.PermissionType.ToString() # Microsoft.SqlServer.Management.Smo.ObjectPermissionSet PermissionType {get;}
							}
						}
					)
					# Dropping object permissions from this version because including it causes a massive memory footprint...will address later
					# Possible solution - Use Gz compression to store the Permissions as a compressed string and decompress at rendering time
				} else {
					# Need to figure out how to get permissions from SQL 2000
					$null
				}
			}
		)
	}
	catch {
		Throw
	}
}

function Get-SQLTraceInformation {
	[CmdletBinding()]
	[OutputType([PSObject])]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Server]
		$Server
	)
	Begin {
	}
	Process {
		try {
			$DbEngineType = [String](Get-DatabaseEngineTypeValue -DatabaseEngineType $Server.ServerType) 

			# For now we're only getting trace data from SQL 2005 and up
			# Trace not supported in Azure
			if (
				$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0 -and
				$DbEngineType -ieq $StandaloneDbEngine
			) {
				$Server.Databases['master'].ExecuteWithResults(
					'select id, status, path, max_size, stop_time, max_files, is_rowset, is_rollover, is_shutdown, is_default,' + `
					' buffer_count, buffer_size, file_position, reader_spid, start_time, last_event_time, event_count, dropped_event_count' +`
					' from sys.traces'
				).Tables[0].Rows | ForEach-Object {
					Write-Output (
						New-Object -TypeName PSObject -Property @{
							ID = $_.id # System.Int32 id {get;set;}
							Status = switch ($_.status) {
								0 { 'Stopped' }
								1 { 'Running' }
								$null { $null }
								default { 'Unknown' }
							} # System.Int32 status {get;set;}
							Path = $_.path # System.String path {get;set;}
							MaxSizeMb = $_.max_size # System.Int64 max_size {get;set;}
							StopTime = $_.stop_time # System.DateTime stop_time {get;set;}
							MaxFileCount = $_.max_files # System.Int32 max_files {get;set;}
							IsRowsetTrace = $_.is_rowset # System.Boolean is_rowset {get;set;}
							RolloverEnabled = $_.is_rollover # System.Boolean is_rollover {get;set;}
							ShutdownEnabled = $_.is_shutdown # System.Boolean is_shutdown {get;set;}
							IsDefaultTrace = $_.is_default # System.Boolean is_default {get;set;}
							BufferCount = $_.buffer_count # System.Int32 buffer_count {get;set;}
							BufferSizeKb = $_.buffer_size # System.Int32 buffer_size {get;set;}
							FilePosition = $_.file_position # System.Int64 file_position {get;set;}
							RowsetReaderSessionId = $_.reader_spid # System.Int32 reader_spid {get;set;}
							StartTime = $_.start_time # System.DateTime start_time {get;set;}
							LastEventTime = $_.last_event_time # System.DateTime last_event_time {get;set;}
							EventCount = $_.event_count # System.Int64 event_count {get;set;}
							DroppedEventCount = $_.dropped_event_count # System.Int32 dropped_event_count {get;set;}
							Events = @() + (
								$Server.Databases['master'].ExecuteWithResults(
									"select tcat.category_id, ei.eventid, ei.columnid, tcat.name as category_name, tcat.type as category_type, te.name as event_name, tc.name as column_name" + `
									" from ::fn_trace_geteventinfo($($_.id)) as ei" + `
									" inner join sys.trace_events as te on ei.eventid = te.trace_event_id" + `
									" inner join sys.trace_categories as tcat on te.category_id = tcat.category_id" + `
									" inner join sys.trace_columns as tc on ei.columnid = tc.trace_column_id" 
								).Tables[0].Rows | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										CategoryID = $_.category_id # System.Int16 category_id {get;set;}
										EventID = $_.eventid # System.Int32 eventid {get;set;}
										ColumnID = $_.columnid # System.Int32 columnid {get;set;}
										Category = $_.category_name # System.String category_name {get;set;}
										Type = switch ($_.category_type) {
											0 { 'Normal' }
											1 { 'Connection' }
											2 { 'Error' }
											$null { [String]$null }
											default { 'Unknown' }
										} # System.Byte category_type {get;set;}
										Event = $_.event_name # System.String event_name {get;set;}
										Column = $_.column_name # System.String column_name {get;set;}
									}
								}
							) 
							Filters = @() + (
								$Server.Databases['master'].ExecuteWithResults(
									"select fi.columnid, tc.name, fi.logical_operator, fi.comparison_operator, fi.value" + `
									" from ::fn_trace_getfilterinfo($($_.id)) as fi" + `
									" inner join sys.trace_columns as tc on fi.columnid = tc.trace_column_id" 
								).Tables[0].Rows | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										ColumnID = $_.columnid # System.Int32 columnid {get;set;}
										Column = $_.name # System.String column_name {get;set;}
										LogicalOperator = switch ($_.logical_operator) {
											0 { 'AND' }
											1 { 'OR' }
											$null { [String]$null }
											default { 'Unknown' }
										} # System.Int32 logical_operator {get;set;}
										ComparisonOperator = switch ($_.comparison_operator) {
											0 { 'Equal' }
											1 { 'Not equal' }
											2 { 'Greater than' }
											3 { 'Less than' }
											4 { 'Greater than or equal' }
											5 { 'Less than or equal' }
											6 { 'Like' }
											7 { 'Not like' }
											$null { [String]$null }
											default { 'Uknown' }
										} # System.Int32 comparison_operator {get;set;}
										FilterValue = $_.value # System.Object value {get;set;}
									}
								}
							)
						}
					)
				}
			}
		}
		catch {
			Throw
		}
	}
	End {
	}

}

function Get-ServerTriggerInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Server]
		$Server
	)

	try {


		$DbEngineType = [String](Get-DatabaseEngineTypeValue -DatabaseEngineType $Server.ServerType)

		if ($DbEngineType -ieq $StandaloneDbEngine) { 
			# Normally wouldn't have to check for existence of a property but SMO returns an empty object when enumerating in SQL 2000
			# Other properties (e.g. $Server.ApplicationRoles) do not
			# Might have to file a report on Connect for this
			$Server.Triggers | Where-Object { $_.ID } | ForEach-Object {
				Write-Output (
					New-Object -TypeName psobject -Property @{

						Name = $_.Name # System.String
						ID = $_.ID # System.Int32 ID {get;}
						DdlTriggerEvents = if ($_.DdlTriggerEvents) { $_.DdlTriggerEvents.ToString() } else { $null } # Microsoft.SqlServer.Management.Smo.ServerDdlTriggerEventSet DdlTriggerEvents {get;set;}

						ExecutionContext = [String](Get-ServerDdlTriggerExecutionContextValue -ServerDdlTriggerExecutionContext $_.ExecutionContext) # Microsoft.SqlServer.Management.Smo.ServerDdlTriggerExecutionContext ExecutionContext {get;set;}
						ExecutionContextLogin = $_.ExecutionContextLogin # System.String ExecutionContextLogin {get;set;}

						IsEnabled = $_.IsEnabled # System.Boolean IsEnabled {get;set;}
						IsEncrypted = $_.IsEncrypted # System.Boolean IsEncrypted {get;set;}
						IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}

						AnsiNullsStatus = $_.AnsiNullsStatus # System.Boolean AnsiNullsStatus {get;set;}
						QuotedIdentifierStatus = $_.QuotedIdentifierStatus # System.Boolean QuotedIdentifierStatus {get;set;}

						ImplementationType = [String](Get-ImplementationTypeValue -ImplementationType $_.ImplementationType) # Microsoft.SqlServer.Management.Smo.ImplementationType ImplementationType {get;set;}

						CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
						DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}

						AssemblyName = $_.AssemblyName # System.String AssemblyName {get;set;}
						ClassName = $_.ClassName # System.String ClassName {get;set;}
						MethodName = $_.MethodName # System.String MethodName {get;set;}

						##BodyStartIndex = $_.BodyStartIndex	# System.Int32 BodyStartIndex {get;}
						#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Server Parent {get;set;}
						#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
						#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
						##Text = $_.Text	# System.String Text {get;}
						##TextBody = $_.TextBody	# System.String TextBody {get;set;}
						##TextHeader = $_.TextHeader	# System.String TextHeader {get;set;}
						##TextMode = $_.TextMode	# System.Boolean TextMode {get;set;}
						#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
						#UserData = $_.UserData	# System.Object UserData {get;set;}

					}
				)
			}
		} else {
			@()
		}
	}
	catch {
		Throw
	}
}

# function Get-ObjectInformation {
# 	[CmdletBinding()]
# 	param(
# 		[Parameter(Mandatory=$true)] 
# 		[Microsoft.SqlServer.Management.Smo.Server]
# 		$Server
# 	)
# 
# 	try {
# 		Write-Output (
# 			New-Object -TypeName psobject -Property @{
# 				
# 			}
# 		)
# 	}
# 	catch {
# 		Throw
# 	}
# }

function Get-StartupProcedureInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Server]
		$Server
	)
	try {
		$Server.EnumStartupProcedures() | ForEach-Object {
			Write-Output (
				New-Object -TypeName psobject -Property @{
					Name = $_.Name # System.String Name {get;set;}
					Schema = $_.Schema # System.String Schema {get;set;}
				}
			)
		} 
	}
	catch {
		Throw
	}
}

function Get-EndpointInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Server]
		$Server
	)
	try {

		$DbEngineType = [String](Get-DatabaseEngineTypeValue -DatabaseEngineType $Server.ServerType)

		# Endpoints not available until SQL 2005
		if (
			$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0 -and
			$DbEngineType -ieq $StandaloneDbEngine # Doesn't work against Azure
		) {
			$Server.Endpoints | ForEach-Object {
				Write-Output (
					New-Object -TypeName psobject -Property @{
						EndpointState = [String](Get-EndpointStateValue -EndpointState $_.EndpointState) # Microsoft.SqlServer.Management.Smo.EndpointState EndpointState {get;}
						EndpointType = [String](Get-EndpointTypeValue -EndpointType $_.EndpointType) # Microsoft.SqlServer.Management.Smo.EndpointType EndpointType {get;set;}
						ID = $_.ID # System.Int32 ID {get;}
						IsAdminEndpoint = $_.IsAdminEndpoint # System.Boolean IsAdminEndpoint {get;}
						IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
						Name = $_.Name # System.String Name {get;set;}
						Owner = $_.Owner # System.String Owner {get;set;}
						#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Server Parent {get;set;}
						#Payload = $_.Payload	# Microsoft.SqlServer.Management.Smo.Payload Payload {get;}
						#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
						#Protocol = $_.Protocol	# Microsoft.SqlServer.Management.Smo.Protocol Protocol {get;}
						ProtocolType = [String](Get-ProtocolTypeValue -ProtocolType $_.ProtocolType) # Microsoft.SqlServer.Management.Smo.ProtocolType ProtocolType {get;set;}
						#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
						#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
						#UserData = $_.UserData		# System.Object UserData {get;set;}				
					}
				)
			}
		} else {
			@()
		}
	}
	catch {
		Throw
	}
}

function Get-LinkedServerInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Server]
		$Server
	)
	try {
		$DbEngineType = [String](Get-DatabaseEngineTypeValue -DatabaseEngineType $Server.ServerType)

		if ($DbEngineType -ieq $StandaloneDbEngine) {
			$Server.LinkedServers | ForEach-Object {
				Write-Output (
					New-Object -TypeName psobject -Property @{
						General = New-Object -TypeName psobject -Property @{
							Name = $_.Name # System.String Name {get;set;}
							ProductName = $_.ProductName # System.String ProductName {get;set;}
							ProviderName = $_.ProviderName # System.String ProviderName {get;set;}
							DataSource = $_.DataSource # System.String DataSource {get;set;}
							ProviderString = $_.ProviderString # System.String ProviderString {get;set;}
							Location = $_.Location # System.String Location {get;set;}
							Catalog = $_.Catalog # System.String Catalog {get;set;}
						}
						Options = New-Object -TypeName psobject -Property @{
							CollationCompatible = $_.CollationCompatible # System.Boolean CollationCompatible {get;set;}
							DataAccess = $_.DataAccess # System.Boolean DataAccess {get;set;}
							Rpc = $_.Rpc # System.Boolean Rpc {get;set;}
							RpcOut = $_.RpcOut # System.Boolean RpcOut {get;set;}
							UseRemoteCollation = $_.UseRemoteCollation # System.Boolean UseRemoteCollation {get;set;}
							CollationName = $_.CollationName # System.String CollationName {get;set;}
							ConnectTimeoutSeconds = $_.ConnectTimeout # System.Int32 ConnectTimeout {get;set;}
							QueryTimeoutSeconds = $_.QueryTimeout # System.Int32 QueryTimeout {get;set;}
							Distributor = $_.Distributor # System.Boolean Distributor {get;set;}
							DistPublisher = $_.DistPublisher # System.Boolean DistPublisher {get;set;}
							Publisher = $_.Publisher # System.Boolean Publisher {get;set;}
							Subscriber = $_.Subscriber # System.Boolean Subscriber {get;set;}
							LazySchemaValidation = $_.LazySchemaValidation # System.Boolean LazySchemaValidation {get;set;}
							IsPromotionofDistributedTransactionsForRPCEnabled = $_.IsPromotionofDistributedTransactionsForRPCEnabled # System.Boolean IsPromotionofDistributedTransactionsForRPCEnabled {get;set;}
						}
						Security = @() + (
							$_.LinkedServerLogins | ForEach-Object {
								New-Object -TypeName psobject -Property @{
									LocalLogin = $_.Name # System.String Name {get;set;}
									Impersonate = $_.Impersonate # System.Boolean Impersonate {get;set;}
									RemoteUser = $_.RemoteUser # System.String RemoteUser {get;set;}
									DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
								}
							}
						)
					}
				)
			} 
		} else {
			@()
		}
	}
	catch {
		Throw
	}
}

function Get-TraceFlagInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Server]
		$Server
	)
	try {

		$DbEngineType = [String](Get-DatabaseEngineTypeValue -DatabaseEngineType $Server.ServerType) 

		# Not using SMO for this since SMO 2005 doesn't have an easy way to get Trace Flags
		# ...but DBCC TRACESTATUS is supported all the way back to SQL 2000

		# Trace Flags not supported in Azure
		if ($DbEngineType -ieq $StandaloneDbEngine) {
			$Server.Databases['master'].ExecuteWithResults('DBCC TRACESTATUS WITH NO_INFOMSGS').Tables[0].Rows | ForEach-Object {
				Write-Output (
					New-Object -TypeName psobject -Property @{
						TraceFlag = $_.TraceFlag # System.Int16 TraceFlag {get;set;}
						Status = switch ($_.Status) {
							0 { 'OFF' }
							1 { 'ON' }
							default { 'Uknown' }
						} # System.Int16 Status {get;set;}
						IsGlobal = switch ($_.Global) {
							0 { $true }
							1 { $false }
							default { $null }
						} # System.Int16 Global {get;set;}
						IsSession = switch ($_.Session) {
							0 { $true }
							1 { $false }
							default { $null }
						} # System.Int16 Session {get;set;}
					}
				)
			}
		} else {
			@()
		}

	}
	catch {
		Throw
	}
}

function Get-SqlServerServiceInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Server]
		$Server
	)
	try {
		Write-Output (
			New-Object -TypeName psobject -Property @{
				ServiceAccount = $Server.ServiceAccount # System.String ServiceAccount {get;}
				ServiceInstanceId = $Server.ServiceInstanceId # System.String ServiceInstanceId {get;}
				ServiceName = $Server.ServiceName # System.String ServiceName {get;}
				ServiceStartMode = [String](Get-ServiceStartModeValue -ServiceStartMode $Server.ServiceStartMode) # Microsoft.SqlServer.Management.Smo.ServiceStartMode ServiceStartMode {get;}

				# This is really an approximation. A more accurate start date can be retrieved from WMI for the PID of the SQL Server Process
				ServiceStartDate = $Server.Databases['tempdb'].CreateDate
			}
		) 
	}
	catch {
		Throw
	}
}

function Get-SqlAgentServiceInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Agent.JobServer]
		$JobServer
	)
	process {
		try {
			Write-Output (
				New-Object -TypeName psobject -Property @{
					AutoStart = $JobServer.SqlAgentAutoStart # System.Boolean SqlAgentAutoStart {get;set;}
					DomainGroup = $JobServer.AgentDomainGroup # System.String AgentDomainGroup {get;}
					ServiceAccount = $JobServer.ServiceAccount # System.String ServiceAccount {get;}
					#ServiceInstanceId = $JobServer.ServiceInstanceId
					ServiceName = $JobServer.Name # System.String Name {get;}
					ServiceStartMode = [String](Get-ServiceStartModeValue -ServiceStartMode $JobServer.ServiceStartMode) # Microsoft.SqlServer.Management.Smo.ServiceStartMode ServiceStartMode {get;}
				}
			)
		}
		catch {
			throw
		}
	}
}

function Get-SqlAgentAlertInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Agent.JobServer]
		$JobServer
	)
	process {
		try {

			$JobServer.Alerts | ForEach-Object { 
				Write-Output (
					New-Object -TypeName psobject -Property @{
						General = New-Object -TypeName psobject -Property @{
							ID = $_.ID # System.Int32 ID {get;}
							Name = $_.Name # System.String Name {get;set;}
							IsEnabled = $_.IsEnabled # System.Boolean IsEnabled {get;set;}
							EventSource = $_.EventSource # System.String EventSource {get;}
							CategoryName = $_.CategoryName # System.String CategoryName {get;set;}
							Type = [String](Get-AlertTypeValue -AlertType $_.AlertType) # Microsoft.SqlServer.Management.Smo.Agent.AlertType AlertType {get;}
							Definition = New-Object -TypeName psobject -Property @{
								DatabaseName = $_.DatabaseName # System.String DatabaseName {get;set;}
								ErrorNumber = $_.MessageID # System.Int32 MessageID {get;set;}
								Severity = [String](Get-EventSeverityLevelValue -EventSeverityLevel $_.Severity) # System.Int32 Severity {get;set;}
								EventDescriptionKeyword = $_.EventDescriptionKeyword # System.String EventDescriptionKeyword {get;set;}
								PerformanceCondition = $_.PerformanceCondition # System.String PerformanceCondition {get;set;}
								WmiNamespace = $_.WmiEventNamespace # System.String WmiEventNamespace {get;set;}
								WmiQuery = $_.WmiEventQuery # System.String WmiEventQuery {get;set;}
							}
						}
						Response = New-Object -TypeName psobject -Property @{
							ExecuteJobID = $_.JobID # System.Guid JobID {get;set;}
							ExecuteJobName = $_.JobName # System.String JobName {get;}
							NotifyOperators = if ($_.HasNotification -gt 0) {
								$_.EnumNotifications() | ForEach-Object {
									New-Object -TypeName psobject -Property @{
										OperatorId = $_.OperatorId # System.Int32 OperatorId {get;set;}
										OperatorName = $_.OperatorName # System.String OperatorName {get;set;}
										HasEmail = $_.HasEmail # System.Boolean HasEmail {get;set;}
										HasNetSend = $_.HasNetSend # System.Boolean HasNetSend {get;set;}
										HasPager = $_.HasPager # System.Boolean HasPager {get;set;}
										UseEmail = $_.UseEmail # System.Boolean UseEmail {get;set;}
										UseNetSend = $_.UseNetSend # System.Boolean UseNetSend {get;set;}
										UsePager = $_.UsePager # System.Boolean UsePager {get;set;}
									}
								}
							} else {
								@()
							}
						}
						Options = New-Object -TypeName psobject -Property @{
							IncludeErrorTextIn = [String](Get-NotificationMethodValue -NotificationMethod $_.IncludeEventDescription)
							NotificationMessage = $_.NotificationMessage # System.String NotificationMessage {get;set;}
							DelaySecondsBetweenResponses = $_.DelayBetweenResponses # System.Int32 DelayBetweenResponses {get;set;}
						}
						History = New-Object -TypeName psobject -Property @{
							LastAlertDate = if (($_.LastOccurrenceDate) -and ($_.LastOccurrenceDate.CompareTo($SmoEpoch) -le 0)) { $null } else { $_.LastOccurrenceDate } # System.DateTime LastOccurrenceDate {get;set;}
							LastResponseDate = if (($_.LastResponseDate) -and ($_.LastResponseDate.CompareTo($SmoEpoch) -le 0)) { $null } else { $_.LastResponseDate } # System.DateTime LastResponseDate {get;set;}
							OccurrenceCount = $_.OccurrenceCount # System.Int32 OccurrenceCount {get;}
							CountResetDate = if (($_.CountResetDate) -and ($_.CountResetDate.CompareTo($SmoEpoch) -le 0)) { $null } else { $_.CountResetDate } # System.DateTime CountResetDate {get;}
						}
					}
				)
			}
		}
		catch {
			throw
		}
	}
}

function Get-SqlAgentConfigurationInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Agent.JobServer]
		$JobServer
	)

	# See http://msdn.microsoft.com/en-us/library/ms189631
	try {
		Write-Output (
			New-Object -TypeName psobject -Property @{
				General = New-Object -TypeName psobject -Property @{
					Name = $JobServer.Name # System.String Name {get;}
					ServerType = [String](Get-JobServerTypeValue -JobServerType $JobServer.JobServerType) # Microsoft.SqlServer.Management.Smo.Agent.JobServerType JobServerType {get;}
					AgentService = New-Object -TypeName psobject -Property @{
						AutoRestartSqlServer = $JobServer.SqlServerRestart # System.Boolean SqlServerRestart {get;set;}
						AutoRestartSqlAgent = $JobServer.SqlAgentRestart # System.Boolean SqlAgentRestart {get;set;}
					}
					ErrorLog = New-Object -TypeName psobject -Property @{
						FileName = $JobServer.ErrorLogFile # System.String ErrorLogFile {get;set;}
						WriteOemFile = $JobServer.WriteOemErrorLog # System.Boolean WriteOemErrorLog {get;set;}
						NetSendRecipient = $JobServer.NetSendRecipient # System.String NetSendRecipient {get;set;}

						# Not part of the SSMS GUI but I think these properties belong here
						LogLevel = [String](Get-AgentLogLevelValue -AgentLogLevel $JobServer.AgentLogLevel) # Microsoft.SqlServer.Management.Smo.Agent.AgentLogLevels AgentLogLevel {get;set;}
					} | Add-Member -MemberType NoteProperty -Name IncludeExecutionTraceMessages -Value $(
						switch (Get-AgentLogLevelValue -AgentLogLevel $JobServer.AgentLogLevel) {
							'all' { $true }
							default { $false }
						}
					) -PassThru

				}
				Advanced = New-Object -TypeName psobject -Property @{
					EventForwarding = New-Object -TypeName psobject -Property @{
						IsEnabled = if ($JobServer.AlertSystem.ForwardingServer) { $true } else { $false }
						Server = $JobServer.AlertSystem.ForwardingServer # System.String ForwardingServer {get;set;}
						EventsToForward = if ($JobServer.AlertSystem.IsForwardedAlways) { 'All Events' } else { 'Unhandled Events' } # System.Boolean IsForwardedAlways {get;set;}
						SeverityAtOrAbove = [String](Get-EventSeverityLevelValue -EventSeverityLevel $JobServer.AlertSystem.ForwardingSeverity) # System.Int32 ForwardingSeverity {get;set;}
					}
					IdleCpuCondition = New-Object -TypeName psobject -Property @{
						IsEnabled = $JobServer.IsCpuPollingEnabled # System.Boolean IsCpuPollingEnabled {get;set;}
						AvgCpuBelowPercent = $JobServer.IdleCpuPercentage # System.Int32 IdleCpuPercentage {get;set;}
						AvgCpuRemainsForSeconds = $JobServer.IdleCpuDuration # System.Int32 IdleCpuDuration {get;set;}
					}
				}
				AlertSystem = New-Object -TypeName psobject -Property @{
					MailSession = New-Object -TypeName psobject -Property @{
						IsEnabled = if (($JobServer.SqlAgentMailProfile) -or ($JobServer.DatabaseMailProfile)) { $true } else { $false }
						MailSystem = [String](Get-AgentMailTypeValue -AgentMailType $JobServer.AgentMailType) # Microsoft.SqlServer.Management.Smo.Agent.AgentMailType AgentMailType {get;set;}
						SaveSentMessages = $JobServer.SaveInSentFolder # System.Boolean SaveInSentFolder {get;set;}
					} | Add-Member -MemberType NoteProperty -Name MailProfile -Value $(
						switch (Get-AgentMailTypeValue -AgentMailType $JobServer.AgentMailType) {
							'Agent Mail' { $JobServer.SqlAgentMailProfile } 
							'Database Mail' { $JobServer.DatabaseMailProfile }
							default { $null }
						}
					) -PassThru
					PagerEmails = New-Object -TypeName psobject -Property @{
						ToTemplate = $JobServer.AlertSystem.PagerToTemplate # System.String PagerToTemplate {get;set;}
						CcTemplate = $JobServer.AlertSystem.PagerCCTemplate # System.String PagerCCTemplate {get;set;}
						Subject = $JobServer.AlertSystem.PagerSubjectTemplate # System.String PagerSubjectTemplate {get;set;}
						IncludeBody = -not $JobServer.AlertSystem.PagerSendSubjectOnly # System.Boolean PagerSendSubjectOnly {get;set;}
					}
					FailSafeOperator = New-Object -TypeName psobject -Property @{
						Operator = $JobServer.AlertSystem.FailSafeOperator # System.String FailSafeOperator {get;set;}
						FailSafeEmailAddress = $JobServer.AlertSystem.FailSafeEmailAddress # System.String FailSafeEmailAddress {get;set;}
						FailSafeNetSendAddress = $JobServer.AlertSystem.FailSafeNetSendAddress # System.String FailSafeNetSendAddress {get;set;}
						FailSafePagerAddress = $JobServer.AlertSystem.FailSafePagerAddress # System.String FailSafePagerAddress {get;set;}
						NotificationMethod = (Get-NotificationMethodValue -NotificationMethod $JobServer.AlertSystem.NotificationMethod) # Microsoft.SqlServer.Management.Smo.Agent.NotifyMethods NotificationMethod {get;set;}
					}
					TokenReplacement = New-Object -TypeName psobject -Property @{
						ReplaceAlertTokens = $JobServer.ReplaceAlertTokensEnabled # System.Boolean ReplaceAlertTokensEnabled {get;set;}
					}
				}
				JobSystem = New-Object -TypeName psobject -Property @{
					ShutdownTimeoutIntervalSeconds = $JobServer.AgentShutdownWaitTime # System.Int32 AgentShutdownWaitTime {get;set;}
					#JobStepProxyAccount = New-Object -TypeName psobject -Property @{}		# SQL 2000 feature, not available w\ SMO
				}
				Connection = New-Object -TypeName psobject -Property @{
					AliasLocalHostServer = $JobServer.LocalHostAlias # System.String LocalHostAlias {get;set;}
					#SqlServerConnection = New-Object -TypeName psobject -Property @{}		# Only Available for multiserver administration 

					# Not part of the SSMS GUI but I think these properties belong here
					LoginTimeoutSeconds = $JobServer.LoginTimeout # System.Int32 LoginTimeout {get;set;}
				}
				History = New-Object -TypeName psobject -Property @{
					MaxJobHistoryTotalRows = $JobServer.MaximumHistoryRows # System.Int32 MaximumHistoryRows {get;set;}
					MaxJobHistoryRowsPerJob = $JobServer.MaximumJobHistoryRows # System.Int32 MaximumJobHistoryRows {get;set;}
				}
			}
		)
	}
	catch {
		Throw
	}
}

function Get-SqlAgentJobInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Agent.JobServer]
		$JobServer
		,
		[Parameter(Mandatory=$false)]
		[String[]]
		$JobName = $null
	)
	process {

		try {

			$Job = @()
			# 		$DatabaseStatus = 'Microsoft.SqlServer.Management.Smo.DatabaseStatus' -as [Type]
			# 		$CompatibilityLevel = 'Microsoft.SqlServer.Management.Smo.CompatibilityLevel' -as [Type]

			$WeeklyFrequency = @{
				1 = 'Sunday'
				2 = 'Monday'
				4 = 'Tuesday'
				8 = 'Wednesday'
				16 = 'Thursday'
				32 = 'Friday'
				64 = 'Saturday'
				62 = 'weekday'
				65 = 'weekend day'
				127 = 'every day'
			}

			if ($JobName) {
				$Job += $JobServer.Jobs | Where-Object { $JobName -icontains $_.Name }
			} else {
				$Job = $JobServer.Jobs
			}

			$Job | ForEach-Object {
				Write-Output (
					New-Object -TypeName psobject -Property @{
						General = New-Object -TypeName psobject -Property @{
							Name = $_.Name # System.String Name {get;set;}
							ID = $_.JobID # System.Guid JobID {get;}
							Owner = $_.OwnerLoginName # System.String OwnerLoginName {get;set;}
							Category = $_.Category # System.String Category {get;set;}
							Description = $_.Description # System.String Description {get;set;}
							Enabled = $_.IsEnabled # System.Boolean IsEnabled {get;set;}
							#Source = ??
							Created = $_.DateCreated # System.DateTime DateCreated {get;}
							LastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
							LastExecuted = if ($_.LastRunDate.CompareTo($SmoEpoch) -le 0) { $null } else { $_.LastRunDate } # System.DateTime LastRunDate {get;}
							LastOutcome = [String](Get-AgentCompletionResultValue -CompletionResult $_.LastRunOutcome) # Microsoft.SqlServer.Management.Smo.Agent.CompletionResult LastRunOutcome {get;}
							NextRun = if ($_.NextRunDate.CompareTo($SmoEpoch) -le 0) { $null } else { $_.NextRunDate } # System.DateTime NextRunDate {get;}
						}
						Steps = @() + (
							$_.JobSteps | ForEach-Object {
								New-Object -TypeName psobject -Property @{
									Id = $_.Id # System.Int32 ID {get;set;}
									General = New-Object -TypeName psobject -Property @{
										StepName = $_.Name # System.String Name {get;set;}
										Type = [String](Get-AgentSubSystemValue -AgentSubSystem $_.SubSystem) # Microsoft.SqlServer.Management.Smo.Agent.AgentSubSystem SubSystem {get;set;}
										RunAs = $_.DatabaseUserName # System.String DatabaseUserName {get;set;}
										Database = $_.DatabaseName # System.String DatabaseName {get;set;}
										SuccessExitCode = $_.CommandExecutionSuccessCode # System.Int32 CommandExecutionSuccessCode {get;set;}
										Command = $_.Command # System.String Command {get;set;}
									}
									Advanced = New-Object -TypeName psobject -Property @{
										OnSuccessAction = [String](Get-AgentStepCompletionActionValue -StepCompletionAction $_.OnSuccessAction) # Microsoft.SqlServer.Management.Smo.Agent.StepCompletionAction OnSuccessAction {get;set;}
										OnSuccessStep = $_.OnSuccessStep # System.Int32 OnSuccessStep {get;set;}
										OnFailAction = [String](Get-AgentStepCompletionActionValue -StepCompletionAction $_.OnFailAction) # Microsoft.SqlServer.Management.Smo.Agent.StepCompletionAction OnFailAction {get;set;}
										OnFailStep = $_.OnFailStep # System.Int32 OnFailStep {get;set;}
										RetryAttempts = $_.RetryAttempts # System.Int32 RetryAttempts {get;set;}
										RetryIntervalMinutes = $_.RetryInterval # System.Int32 RetryInterval {get;set;}
										Logging = New-Object -TypeName psobject -Property @{
											OutputFile = $_.OutputFileName # System.String OutputFileName {get;set;}
											#AppendToFile = ?
											#LogToTable = ?
											#AppendToTable = ?
											#IncludeStepOutputInHistory = ?								
										}
										#RunAsUser = ?
									}
								}
							}
						)
						Schedules = @() + (
							$_.JobSchedules | ForEach-Object {
								New-Object -TypeName psobject -Property @{
									Id = $_.ID # System.Int32 ID {get;}
									Name = $_.Name # System.String Name {get;set;}
									DateCreated = $_.DateCreated # System.DateTime DateCreated {get;}	# Not part of the GUI, but useful
									IsEnabled = $_.IsEnabled # System.Boolean IsEnabled {get;set;}
									Description = $(
										New-Object -TypeName psobject -Property @{
											Frequency = New-Object -TypeName psobject -Property @{
												Occurs = $_.FrequencyTypes.ToString() # Microsoft.SqlServer.Management.Smo.Agent.FrequencyTypes FrequencyTypes {get;set;}
												RelativeInterval = $_.FrequencyRelativeIntervals.ToString().ToLower() # Microsoft.SqlServer.Management.Smo.Agent.FrequencyRelativeIntervals FrequencyRelativeIntervals {get;set;}

												# See http://msdn.microsoft.com/en-us/library/microsoft.sqlserver.management.smo.agent.jobschedule.frequencyinterval.aspx
												# to interpret $_.FrequencyInterval												
												Interval = if ($_.FrequencyTypes.ToString() -ieq 'MonthlyRelative') {
													switch ($_.FrequencyInterval) {
														1 { 'Sunday' }
														2 { 'Monday' }
														3 { 'Tuesday' }
														4 { 'Wednesday' }
														5 { 'Thursday' }
														6 { 'Friday' }
														7 { 'Saturday' }
														8 { 'day' }
														9 { 'weekday' }
														10 { 'weekend day' }
													} 
												} elseif ($_.FrequencyTypes.ToString() -ieq 'Monthly') {
													[String]::Join(' ', @('day',$_.FrequencyInterval))
												} elseif ($_.FrequencyTypes.ToString() -ieq 'Weekly') {
													$FrequencyInterval = $_.FrequencyInterval
													[String]::Join(', ', @($WeeklyFrequency.Keys | Where-Object { ($_ -bor $FrequencyInterval) -eq $FrequencyInterval } | ForEach-Object { $WeeklyFrequency[$_] } ))
												} elseif ($_.FrequencyTypes.ToString() -ieq 'Daily') {
													[String]::Join(' ', @($_.FrequencyInterval,'day(s)'))
												} else {
													[String]::Empty
												} # Microsoft.SqlServer.Management.Smo.Agent.FrequencyTypes FrequencyTypes {get;set;}
												RecurrenceFactor = $_.FrequencyRecurrenceFactor # System.Int32 FrequencyRecurrenceFactor {get;set;}
											}
											DailyFrequency = New-Object -TypeName psobject -Property @{
												OccursEvery = $_.FrequencySubDayInterval # System.Int32 FrequencySubDayInterval {get;set;}
												OccursEveryFreq = $_.FrequencySubDayTypes.ToString() # Microsoft.SqlServer.Management.Smo.Agent.FrequencySubDayTypes FrequencySubDayTypes {get;set;}
												StartTimeOfDay = [System.DateTime]::ParseExact($_.ActiveStartTimeOfDay, 'HH:mm:ss', $null).ToString('h:mm:ss tt', [System.Globalization.CultureInfo]::CurrentCulture) # System.TimeSpan ActiveStartTimeOfDay {get;set;}
												EndTimeOfDay = [System.DateTime]::ParseExact($_.ActiveEndTimeOfDay, 'HH:mm:ss', $null).ToString('h:mm:ss tt', [System.Globalization.CultureInfo]::CurrentCulture) # System.TimeSpan ActiveEndTimeOfDay {get;set;}
											}
											Duration = New-Object -TypeName psobject -Property @{
												StartDate = $_.ActiveStartDate.ToString('d') # System.DateTime ActiveStartDate {get;set;}
												EndDate = $_.ActiveEndDate.ToString('d') # System.DateTime ActiveEndDate {get;set;}
											}
										} | ForEach-Object { 
											$(
												if ($_.Frequency.Occurs -ieq 'AutoStart') {
													'Start automatically when SQL Server Agent starts'
												} elseif ($_.Frequency.Occurs -ieq 'OnIdle') {
													'Start whenever the CPUs become idle'
												} else {
													'Occurs' + $(
														if ($_.Frequency.Interval -ieq [String]::Empty) { 
															' on ' + $_.Duration.StartDate + ' at ' + $_.DailyFrequency.StartTimeOfDay + '.'
														} else { 
															' every ' + `
															$(
																if ($_.Frequency.Occurs -ieq 'Monthly') {
																	$(
																		if ($_.Frequency.RecurrenceFactor -gt 1) {
																			$_.Frequency.RecurrenceFactor.ToString() + ' month(s) '
																		} else {
																			'month '
																		}
																	) + `
																	'on ' + $_.Frequency.Interval + ' of that month' 
																}
																elseif ($_.Frequency.Occurs -ieq 'Weekly') {
																	$(
																		if ($_.Frequency.RecurrenceFactor -gt 1) {
																			$_.Frequency.RecurrenceFactor.ToString() + ' week(s) '
																		} else {
																			'week '
																		}
																	) + `
																	'on ' + $_.Frequency.Interval
																}
																elseif ($_.Frequency.Occurs -ieq 'Daily') {
																	$_.Frequency.Interval 
																}
																else {
																	$_.Frequency.RelativeInterval + ' ' + `
																	$_.Frequency.Interval + ' of every ' + `
																	$_.Frequency.RecurrenceFactor + ' month(s)' 
																} 
															) + `
															$(
																if ($_.DailyFrequency.OccursEveryFreq -ieq 'once') {
																	' at ' + $_.DailyFrequency.StartTimeOfDay
																}
																else {
																	' every ' + $_.DailyFrequency.OccursEvery.ToString() + ' ' + `
																	$_.DailyFrequency.OccursEveryFreq + '(s) between ' + `
																	$_.DailyFrequency.StartTimeOfDay + ' and ' + $_.DailyFrequency.EndTimeOfDay
																}
															) + `
															'. Schedule will be used ' + `
															$(
																if ($_.Duration.EndDate -ieq '12/31/9999') {
																	'starting on ' + `
																	$_.Duration.StartDate
																} else {
																	'between ' + $_.Duration.StartDate + ' and ' + $_.Duration.EndDate
																}
															) + `
															'.'
														}
													)
												}
											)
										}
									)
								}
							}
						)
						Alerts = @() + (
							$_.EnumAlerts() | ForEach-Object {
								New-Object -TypeName psobject -Property @{
									ID = $_.ID # System.Int32 ID {get;set;}
								}
							}
						)
						Notifications = New-Object -TypeName psobject -Property @{
							EmailOperator = $_.OperatorToEmail # System.String OperatorToEmail {get;set;}
							EmailCondition = [String](Get-AgentCompletionActionValue -CompletionAction $_.EmailLevel) # Microsoft.SqlServer.Management.Smo.Agent.CompletionAction EmailLevel {get;set;}
							PageOperator = $_.OperatorToPage # System.String OperatorToPage {get;set;}
							PageCondition = [String](Get-AgentCompletionActionValue -CompletionAction $_.PageLevel) # Microsoft.SqlServer.Management.Smo.Agent.CompletionAction PageLevel {get;set;}
							NetSendOperator = $_.OperatorToNetSend # System.String OperatorToNetSend {get;set;}
							NetSendCondition = [String](Get-AgentCompletionActionValue -CompletionAction $_.NetSendLevel) # Microsoft.SqlServer.Management.Smo.Agent.CompletionAction NetSendLevel {get;set;}
							EventLogCondition = [String](Get-AgentCompletionActionValue -CompletionAction $_.EventLogLevel) # Microsoft.SqlServer.Management.Smo.Agent.CompletionAction EventLogLevel {get;set;}
							DeleteJobCondition = [String](Get-AgentCompletionActionValue -CompletionAction $_.DeleteLevel) # Microsoft.SqlServer.Management.Smo.Agent.CompletionAction DeleteLevel {get;set;}
						}

						#Targets = @{}
						##History = @{}	# Add this here?
					}
				)
			}
		}
		catch {
		}
		# 		finally {
		# 		}
	}
}

function Get-SqlAgentOperatorInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Agent.JobServer]
		$JobServer
	)
	process {
		try {
			$JobServer.Operators | ForEach-Object { 
				Write-Output (
					New-Object -TypeName psobject -Property @{
						General = New-Object -TypeName psobject -Property @{
							Name = $_.Name # System.String Name {get;set;}
							ID = $_.ID # System.Int32 ID {get;}
							CategoryName = $_.CategoryName # System.String CategoryName {get;set;}
							IsEnabled = $_.Enabled # System.Boolean Enabled {get;set;}
							NotificationOptions = New-Object -TypeName psobject -Property @{
								EmailAddress = $_.EmailAddress # System.String EmailAddress {get;set;}
								NetSendAddress = $_.NetSendAddress # System.String NetSendAddress {get;set;}
								PagerAddress = $_.PagerAddress # System.String PagerAddress {get;set;}
							}
							OnDutySchedule = New-Object -TypeName psobject -Property @{
								OnDutyDays = [String](Get-AgentWeekdaysValue -Weekdays $_.PagerDays) # Microsoft.SqlServer.Management.Smo.Agent.WeekDays PagerDays {get;set;}
								WeekdayStartTime = [System.DateTime]::ParseExact($_.WeekdayPagerStartTime, 'HH:mm:ss', $null).ToString('h:mm:ss tt', [System.Globalization.CultureInfo]::CurrentCulture) # System.TimeSpan WeekdayPagerStartTime {get;set;}
								WeekdayEndTime = [System.DateTime]::ParseExact($_.WeekdayPagerEndTime, 'HH:mm:ss', $null).ToString('h:mm:ss tt', [System.Globalization.CultureInfo]::CurrentCulture) # System.TimeSpan WeekdayPagerEndTime {get;set;}
								SaturdayStartTime = [System.DateTime]::ParseExact($_.SaturdayPagerStartTime, 'HH:mm:ss', $null).ToString('h:mm:ss tt', [System.Globalization.CultureInfo]::CurrentCulture) # System.TimeSpan SaturdayPagerStartTime {get;set;}
								SaturdayEndTime = [System.DateTime]::ParseExact($_.SaturdayPagerEndTime, 'HH:mm:ss', $null).ToString('h:mm:ss tt', [System.Globalization.CultureInfo]::CurrentCulture) # System.TimeSpan SaturdayPagerEndTime {get;set;}
								SundayStartTime = [System.DateTime]::ParseExact($_.SundayPagerStartTime, 'HH:mm:ss', $null).ToString('h:mm:ss tt', [System.Globalization.CultureInfo]::CurrentCulture) # System.TimeSpan SundayPagerStartTime {get;set;}
								SundayEndTime = [System.DateTime]::ParseExact($_.SundayPagerEndTime, 'HH:mm:ss', $null).ToString('h:mm:ss tt', [System.Globalization.CultureInfo]::CurrentCulture) # System.TimeSpan SundayPagerEndTime {get;set;}
							}
						}
						Notifications = New-Object -TypeName psobject -Property @{
							Alerts = @() + (
								$_.EnumNotifications() | ForEach-Object {
									New-Object -TypeName psobject -Property @{
										AlertId = $_.AlertId # System.Int32 AlertId {get;set;}
										AlertName = $_.AlertName # System.String AlertName {get;set;}
										HasEmail = $_.HasEmail # System.Boolean HasEmail {get;set;}
										HasNetSend = $_.HasNetSend # System.Boolean HasNetSend {get;set;}
										HasPager = $_.HasPager # System.Boolean HasPager {get;set;}
										UseEmail = $_.UseEmail # System.Boolean UseEmail {get;set;}
										UseNetSend = $_.UseNetSend # System.Boolean UseNetSend {get;set;}
										UsePager = $_.UsePager # System.Boolean UsePager {get;set;}
									}
								}
							)
							Jobs = @() + (
								$_.EnumJobNotifications() | ForEach-Object {
									New-Object -TypeName psobject -Property @{
										JobId = $_.JobId # System.Guid JobId {get;set;}
										JobName = $_.JobName # System.String JobName {get;set;}
										NotifyLevelEmail = [String](Get-AgentCompletionActionValue -CompletionAction $_.NotifyLevelEmail) # System.Int32 NotifyLevelEmail {get;set;}
										NotifyLevelNetSend = [String](Get-AgentCompletionActionValue -CompletionAction $_.NotifyLevelNetSend) # System.Int32 NotifyLevelNetSend {get;set;}
										NotifyLevelPage = [String](Get-AgentCompletionActionValue -CompletionAction $_.NotifyLevelPage) # System.Int32 NotifyLevelPage {get;set;}
									}
								}
							)
						}
						History = New-Object -TypeName psobject -Property @{
							LastEmailDate = if (($_.LastEmailDate) -and ($_.LastEmailDate.CompareTo($SmoEpoch) -le 0)) { $null } else { $_.LastEmailDate } # System.DateTime LastEmailDate {get;}
							LastPagerDate = if (($_.LastPagerDate) -and ($_.LastPagerDate.CompareTo($SmoEpoch) -le 0)) { $null } else { $_.LastPagerDate } # System.DateTime LastPagerDate {get;}
							LastNetSendDate = if (($_.LastNetSendDate) -and ($_.LastNetSendDate.CompareTo($SmoEpoch) -le 0)) { $null } else { $_.LastNetSendDate } # System.DateTime LastNetSendDate {get;}
						}
					}
				)
			}
		}
		catch {
			throw
		}
	}
}

function Get-SqlServerSecurityInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Server]
		$Server
	)
	try {

		$DbEngineType = [String](Get-DatabaseEngineTypeValue -DatabaseEngineType $Server.ServerType)

		$ServerRoleMemberRole = if (
			$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0 -and
			$DbEngineType -ieq $StandaloneDbEngine # Doesn't work against Azure
		) {
			$Server.Databases['master'].ExecuteWithResults(
				'SELECT r.name AS [RoleName], p.name AS [MemberRoleName]' + `
				' FROM sys.server_principals r' + `
				' INNER JOIN sys.server_role_members m ON r.principal_id = m.role_principal_id' + `
				' INNER JOIN sys.server_principals p ON p.principal_id = m.member_principal_id' + `
				' WHERE r.type = ''R''' + `
				' AND p.type = ''R'''
			).Tables[0].Rows
		} else {
			$null
		}

		$BlankPasswordLogin = if (
			$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0 -and
			$DbEngineType -ieq $StandaloneDbEngine # Doesn't work against Azure
		) {
			$Server.Databases['master'].ExecuteWithResults('SELECT name FROM sys.sql_logins WHERE pwdcompare('''', password_hash) = 1').Tables[0].Rows
		} 
		elseif (
			$($Server.Information.Version).CompareTo($SQLServer2000) -ge 0 -and
			$DbEngineType -ieq $StandaloneDbEngine # Doesn't work against Azure
		) {
			$Server.Databases['master'].ExecuteWithResults('SELECT name FROM syslogins WHERE pwdcompare('''', password) = 1').Tables[0].Rows
		}
		else {
			$null
		}

		$NameAsPasswordLogin = if (
			$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0 -and
			$DbEngineType -ieq $StandaloneDbEngine # Doesn't work against Azure
		) {
			$Server.Databases['master'].ExecuteWithResults('SELECT name FROM sys.sql_logins WHERE pwdcompare(name, password_hash) = 1').Tables[0].Rows
		} 
		elseif (
			$($Server.Information.Version).CompareTo($SQLServer2000) -ge 0 -and
			$DbEngineType -ieq $StandaloneDbEngine # Doesn't work against Azure
		) {
			$Server.Databases['master'].ExecuteWithResults('SELECT name FROM syslogins WHERE pwdcompare(name, password) = 1').Tables[0].Rows
		}
		else {
			$null
		}

		# Grab the NETBIOS machine name for use when resolving machine group members
		# For SQL 2005 and up use SERVERPROPERTY(N'ComputerNamePhysicalNetBIOS')
		# For SQL 2000 use SERVERPROPERTY(N'MachineName')
		$NTMachineName = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
			$Server.Databases['master'].ExecuteWithResults('SELECT SERVERPROPERTY(N''ComputerNamePhysicalNetBIOS'') AS ComputerNamePhysicalNetBIOS').Tables[0].Rows[0].ComputerNamePhysicalNetBIOS
		} else {
			# Note that for clustered servers this will return the virtual name and not the physical machine name
			# We'll deal with this later
			$Server.Databases['master'].ExecuteWithResults('SELECT SERVERPROPERTY(N''MachineName'') AS ComputerNamePhysicalNetBIOS').Tables[0].Rows[0].ComputerNamePhysicalNetBIOS
		}

		# If this is a clustered server get the list of cluster members
		# We'll use this when resolving machine group members
		$ClusterMember = if ($Server.Information.IsClustered -eq $true) {
			Get-FailoverClusterMemberList -Server $Server | ForEach-Object {
				Write-Output $_.Name
			}
		} else {
			$null
		}
		
		# Used for holding results when resolving group membership
		$NTGroupMemberList = $null

		Write-Output (
			New-Object -TypeName psobject -Property @{
				Logins = @() + (
					$Server.Logins | ForEach-Object {

						$LoginName = $_.Name

						New-Object -TypeName psobject -Property @{
							AsymmetricKey = $_.AsymmetricKey # System.String AsymmetricKey {get;set;}
							Certificate = $_.Certificate # System.String Certificate {get;set;}
							CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
							Credential = $_.Credential # System.String Credential {get;set;}
							#DateLastModified = $_.DateLastModified	# System.DateTime DateLastModified {get;}
							DefaultDatabase = $_.DefaultDatabase # System.String DefaultDatabase {get;set;}
							DenyWindowsLogin = $_.DenyWindowsLogin # System.Boolean DenyWindowsLogin {get;set;}
							#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.LoginEvents Events {get;}
							HasAccess = $_.HasAccess # System.Boolean HasAccess {get;}
							ID = $_.ID # System.Int32 ID {get;}
							IsDisabled = $_.IsDisabled # System.Boolean IsDisabled {get;}
							IsLocked = $_.IsLocked # System.Boolean IsLocked {get;}
							IsPasswordExpired = $_.IsPasswordExpired # System.Boolean IsPasswordExpired {get;}
							IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
							#Language = $_.Language	# System.String Language {get;set;}
							#LanguageAlias = $_.LanguageAlias	# System.String LanguageAlias {get;}
							LoginType = [String](Get-LoginTypeValue -LoginType $_.LoginType) # Microsoft.SqlServer.Management.Smo.LoginType LoginType {get;set;}
							#MustChangePassword = $_.MustChangePassword	# System.Boolean MustChangePassword {get;}
							Name = $_.Name # System.String Name {get;set;}
							#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Server Parent {get;set;}
							PasswordExpirationEnabled = $_.PasswordExpirationEnabled # System.Boolean PasswordExpirationEnabled {get;set;}
							PasswordHashAlgorithm = if ($_.PasswordHashAlgorithm) { $_.PasswordHashAlgorithm.ToString() } else { $null } # Microsoft.SqlServer.Management.Smo.PasswordHashAlgorithm PasswordHashAlgorithm {get;}
							PasswordPolicyEnforced = $_.PasswordPolicyEnforced # System.Boolean PasswordPolicyEnforced {get;set;}
							#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
							Sid = [System.BitConverter]::ToString($_.Sid) # System.Byte[] Sid {get;set;}
							#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
							#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
							#UserData = $_.UserData	# System.Object UserData {get;set;}
							WindowsLoginAccessType = if ($_.WindowsLoginAccessType) { $_.WindowsLoginAccessType.ToString() } else { $null } # Microsoft.SqlServer.Management.Smo.WindowsLoginAccessType WindowsLoginAccessType {get;}

							HasBlankPassword = if ($BlankPasswordLogin | Where-Object { $_.Name -ieq $LoginName }) { $true } else { $false }
							HasNameAsPassword = if ($NameAsPasswordLogin | Where-Object { $_.Name -ieq $LoginName }) { $true } else { $false }


							# Resolve group members for standalone instances only
							# Ignore 'NT SERVICE\*' (it's a service SID, not an actual group that you can put members into) 
							# Ignore 'NT AUTHORITY\*' (also not an actual group that you can put members into) 
							Member = $(
								if (
									[String](Get-LoginTypeValue -LoginType $_.LoginType) -ieq 'windows group' -and
									$DbEngineType -ieq $StandaloneDbEngine -and
									$LoginName -inotlike 'NT SERVICE\*' -and
									$LoginName -inotlike 'NT AUTHORITY\*'
								) {
									$NTGroupMemberList = @('Unable to resolve group members')

									try {
										$LoginDomain = $LoginName.Split('\')[0]
										if (
											$LoginDomain -ine $NTMachineName -and
											$ClusterMember -icontains $LoginDomain
										) {
											$Account = Resolve-AccountSid -Sid $_.Sid -ComputerName $LoginDomain
										} else {
											$Account = Resolve-AccountSid -Sid $_.Sid -ComputerName $NTMachineName
										}
										
										# If the SID was resolved to an account then determine if it's a domain group or a machine group
										if ($Account.IsResolved) {
											if ($Account.AccountType -ieq 'Group') {
												Get-NTGroupMemberList -NTDomainName $Account.ReferencedDomainName -GroupName $Account.AccountName -OutVariable NTGroupMemberList -Recurse | Out-Null
											} else {
												# If the login's machine name doesn't match the server's machine name and...
												# the login's machine name doesn't match one of the member names in the cluster (or there is no cluster)...
												# then use the server's machine name to resolve the group
												if (
													$LoginDomain -ine $NTMachineName -and
													$ClusterMember -icontains $LoginDomain
												) {
													Get-NTGroupMemberList -NTMachineName $NTDomainName -GroupName $Account.AccountName -OutVariable NTGroupMemberList -Recurse | Out-Null
												} else {
													Get-NTGroupMemberList -NTMachineName $NTMachineName -GroupName $Account.AccountName -OutVariable NTGroupMemberList -Recurse | Out-Null
												}												
											}											
										} else {
											Write-SqlServerDatabaseEngineInformationLog -Message "[$($Server.ConnectionContext.ServerInstance)] Unable to resolve group members for [$LoginName]: Error Code $($Account.ErrorCode)" -MessageLevel Warning
										}
									}
									catch {
										$ErrorRecord = $_.Exception.ErrorRecord
										Write-SqlServerDatabaseEngineInformationLog -Message "[$($Server.ConnectionContext.ServerInstance)] Unable to resolve group members for [$LoginName]: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
										
										# If we were able to retrieve at least a partial list of members then add a fake user called "*Partial List"
										if (($NTGroupMemberList | Measure-Object).Count -gt 0) {
											$NTGroupMemberList += New-Object -TypeName psobject -Property @{
												
											}
										}
										
									}
									finally {
										Write-Output ($NTGroupMemberList)
									}
									
								} else {
									Write-Output @()
								}
							)

						}
					}
				)
				ServerRoles = if ($DbEngineType -ieq $StandaloneDbEngine) {
					@() + (
						$Server.Roles | ForEach-Object {
							$RoleName = $_.Name

							New-Object -TypeName psobject -Property @{
								DateCreated = $_.DateCreated # System.DateTime DateCreated {get;}
								DateModified = $_.DateModified # System.DateTime DateModified {get;}
								#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.ServerRoleEvents Events {get;}
								ID = $_.ID # System.Int32 ID {get;}
								IsFixedRole = $_.IsFixedRole # System.Boolean IsFixedRole {get;}
								Name = $_.Name # System.String Name {get;set;}
								Owner = $_.Owner # System.String Owner {get;set;}
								#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Server Parent {get;set;}
								#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
								#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
								#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
								#UserData = $_.UserData	# System.Object UserData {get;set;}

								# EnumServerRoleMembers() still works but is deprecated so this might need to change in the future
								Member = @() + ($_.EnumServerRoleMembers() | ForEach-Object { $_ } ) # Is this the best way to do this?

								MemberOf = @() + (
									$ServerRoleMemberRole | Where-Object { $_.MemberRoleName -eq $RoleName } | ForEach-Object {
										$_.RoleName
									}
								)

								## EnumMemberNames() didn't show up until SQL 2012 SMO and won't work against older SQL versions
								#MemberOf = if (($Server.Information.Version).CompareTo($SQLServer2012) -ge 0) {
								#	@() + $_.EnumContainingRoleNames() | ForEach-Object { $_ }
								#} else {
								#	$null
								#	#@('Not Supported')
								#}

								# Enumerate through object permissions and report on non-default permissions?

								#Parent = $_.Parent
								#Properties = $_.Properties
								#State = $_.State
								#Urn = $_.Urn
								#UserData = $_.UserData				
							}
						}
					)
				} else {
					$null
				}

				# Not available in SQL 2000
				Credentials = if (
					$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0 -and
					$DbEngineType -ieq $StandaloneDbEngine
				) { 
					@() + (
						$Server.Credentials | ForEach-Object {
							New-Object -TypeName psobject -Property @{
								DateCreated = $_.CreateDate # System.DateTime CreateDate {get;}
								DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
								ID = $_.ID # System.Int32 ID {get;}
								Identity = $_.Identity # System.String Identity {get;set;}
								MappedClassType = if ($_.MappedClassType) { $_.MappedClassType.ToString() } else { $null }  # Microsoft.SqlServer.Management.Smo.MappedClassType MappedClassType {get;set;}
								Name = $_.Name # System.String Name {get;set;}
								#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Server Parent {get;set;}
								#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
								ProviderName = $_.ProviderName # System.String ProviderName {get;set;}
								#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
								#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
								#UserData = $_.UserData	# System.Object UserData {get;set;}
							}
						}
					)
				} else {
					$null
					#@()
				}

				#CryptographicProviders

				# Server Audits available starting with SQL 2008
				Audits = if (
					$($Server.Information.Version).CompareTo($SQLServer2008) -ge 0 -and 
					$SmoMajorVersion -ge 10 -and
					$DbEngineType -ieq $StandaloneDbEngine
				) {
					@() + (
						$Server.Audits | Where-Object { $_.ID } | ForEach-Object { 
							Write-Output (
								New-Object -TypeName psobject -Property @{
									General = New-Object -TypeName psobject -Property @{
										AuditName = $_.Name # System.String Name {get;set;}
										QueueDelayMilliseconds = $_.QueueDelay # System.Int32 QueueDelay {get;set;}
										OnAuditLogFailure = [String](Get-OnFailureActionValue -OnFailureAction $_.OnFailure) # Microsoft.SqlServer.Management.Smo.OnFailureAction OnFailure {get;set;}
										AuditDestination = [String](Get-AuditDestinationTypeValue -AuditDestinationType $_.DestinationType) # Microsoft.SqlServer.Management.Smo.AuditDestinationType DestinationType {get;set;}
										FilePath = $_.FilePath # System.String FilePath {get;set;}
										FileName = $_.FileName # System.String FileName {get;}
										MaximumRolloverFiles = $_.MaximumRolloverFiles # System.Int64 MaximumRolloverFiles {get;set;}
										MaximumFiles = $_.MaximumFiles # System.Int32 MaximumFiles {get;set;}
										MaximumFileSize = $_.MaximumFileSize # System.Int32 MaximumFileSize {get;set;}
										MaximumFileSizeUnit = [String](Get-AuditFileSizeUnitValue -AuditFileSizeUnit $_.MaximumFileSizeUnit) # Microsoft.SqlServer.Management.Smo.AuditFileSizeUnit MaximumFileSizeUnit {get;set;}
										ReserveDiskSpace = $_.ReserveDiskSpace # System.Boolean ReserveDiskSpace {get;set;}

										# Not part of the SSMS GUI
										CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
										DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
										Enabled = $_.Enabled # System.Boolean Enabled {get;}
										ID = $_.ID # System.Int32 ID {get;}
										Guid = $_.Guid # System.Guid Guid {get;set;}
									}
									'Filter' = $_.Filter # System.String Filter {get;set;}

									#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Server Parent {get;set;}
									#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
									#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
									#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
									#UserData = $_.UserData	# System.Object UserData {get;set;}
								}
							)
						}
					)
				} else {
					$null
					#@()
				}

				ServerAuditSpecifications = if (
					$($Server.Information.Version).CompareTo($SQLServer2008) -ge 0 -and 
					$SmoMajorVersion -ge 10 -and
					$DbEngineType -ieq $StandaloneDbEngine
				) {
					@() + (
						$Server.ServerAuditSpecifications | Where-Object { $_.ID } | ForEach-Object { 
							Write-Output (
								New-Object -TypeName psobject -Property @{
									Name = $_.Name # System.String Name {get;set;}
									AuditName = $_.AuditName # System.String AuditName {get;set;}
									Actions = @() + $_.EnumAuditSpecificationDetails() | ForEach-Object {
										New-Object -TypeName psobject -Property @{
											Action = [String](Get-AuditActionTypeValue -AuditActionType $_.Action) # Microsoft.SqlServer.Management.Smo.AuditActionType Action {get;}
											ObjectClass = $_.ObjectClass # System.String ObjectClass {get;}
											ObjectName = $_.ObjectName # System.String ObjectName {get;}
											ObjectSchema = $_.ObjectSchema # System.String ObjectSchema {get;}
											Principal = $_.Principal # System.String Principal {get;}
										}
									}

									# Not part of the SSMS GUI
									CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
									DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
									Enabled = $_.Enabled # System.Boolean Enabled {get;}
									Guid = $_.Guid # System.Guid Guid {get;}
									ID = $_.ID # System.Int32 ID {get;}
								}
							)
						}
					)
				} else {
					$null
					#@()
				}

			}
		)

	}
	catch {
		Throw
	}
}

function Get-ResourceGovernorInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Server]
		$Server
	)
	try {

		$DbEngineType = [String](Get-DatabaseEngineTypeValue -DatabaseEngineType $Server.ServerType) 

		# Resource Governor added in SQL 2008
		# Not supported in Azure
		if (
			$($Server.Information.Version).CompareTo($SQLServer2008) -ge 0 -and
			$DbEngineType -ieq $StandaloneDbEngine 
		) {
			Write-Output (
				New-Object -TypeName psobject -Property @{
					Enabled = $Server.ResourceGovernor.Enabled
					ClassifierFunction = $Server.ResourceGovernor.ClassifierFunction
					ReconfigurePending = $Server.ResourceGovernor.ReconfigurePending
					ResourcePools = @() + (
						$Server.ResourceGovernor.ResourcePools | Where-Object { $_.ID } | ForEach-Object {
							New-Object -TypeName psobject -Property @{
								ID = $_.ID # System.Int32 ID {get;}
								IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
								MaximumCpuPercentage = $_.MaximumCpuPercentage # System.Int32 MaximumCpuPercentage {get;set;}
								MaximumMemoryPercentage = $_.MaximumMemoryPercentage # System.Int32 MaximumMemoryPercentage {get;set;}
								MinimumCpuPercentage = $_.MinimumCpuPercentage # System.Int32 MinimumCpuPercentage {get;set;}
								MinimumMemoryPercentage = $_.MinimumMemoryPercentage # System.Int32 MinimumMemoryPercentage {get;set;}
								Name = $_.Name # System.String Name {get;set;}
								#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.ResourceGovernor Parent {get;set;}
								#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
								#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
								#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
								#UserData = $_.UserData	# System.Object UserData {get;set;}
								WorkloadGroups = @() + (
									$_.WorkloadGroups | Where-Object { $_.ID } | ForEach-Object {
										New-Object -TypeName psobject -Property @{
											GroupMaximumRequests = $_.GroupMaximumRequests # System.Int32 GroupMaximumRequests {get;set;}
											ID = $_.ID # System.Int32 ID {get;}
											Importance = if ($_.Importance) { $_.Importance.ToString() } else { $null } # Microsoft.SqlServer.Management.Smo.WorkloadGroupImportance Importance {get;set;}
											IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
											MaximumDegreeOfParallelism = $_.MaximumDegreeOfParallelism # System.Int32 MaximumDegreeOfParallelism {get;set;}
											Name = $_.Name # System.String Name {get;set;}
											#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.ResourcePool Parent {get;set;}
											#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
											RequestMaximumCpuTimeSeconds = $_.RequestMaximumCpuTimeInSeconds # System.Int32 RequestMaximumCpuTimeInSeconds {get;set;}
											RequestMaximumMemoryGrantPercentage = $_.RequestMaximumMemoryGrantPercentage # System.Int32 RequestMaximumMemoryGrantPercentage {get;set;}
											RequestMemoryGrantTimeoutSeconds = $_.RequestMemoryGrantTimeoutInSeconds # System.Int32 RequestMemoryGrantTimeoutInSeconds {get;set;}
											#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
											#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
											#UserData = $_.UserData	# System.Object UserData {get;set;}
										}
									}
								)
							}
						}
					)
				}
			)
		} else {
			Write-Output (
				New-Object -TypeName psobject -Property @{
					Enabled = $null
					ClassifierFunction = $null
					ReconfigurePending = $null
					ResourcePools = @()
				}
			)
		} 
	}
	catch {
		Throw
	}
}


function Get-DatabaseInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Server]
		$Server
		,
		[Parameter(Mandatory=$false)]
		[String[]]
		$DatabaseName = $null
		,
		[Parameter(Mandatory=$false)]
		[switch]
		$IncludeObjectPermissions = $false
		,
		[Parameter(Mandatory=$false)]
		[switch]
		$IncludeObjectInformation = $false
		,
		[Parameter(Mandatory=$false)]
		[switch]
		$IncludeSystemObjects = $false
	)
	try {

		$Database = @()
		$DbccLogInfo = $null
		$DbccCheckDbInfo = $null
		$LastKnownGoodDbccCheckDbDate = $null
		$DatabaseRoleMemberRole = $null
		$DatabaseStatus = 'Microsoft.SqlServer.Management.Smo.DatabaseStatus' -as [Type]
		$CompatibilityLevel = 'Microsoft.SqlServer.Management.Smo.CompatibilityLevel' -as [Type]
		$DataCompressionType = 'Microsoft.SqlServer.Management.Smo.DataCompressionType' -as [Type]

		$DbEngineType = [String](Get-DatabaseEngineTypeValue -DatabaseEngineType $Server.ServerType)
		$IsAccessible = $true
		$IsFileStream = $false


		$Tables = New-Object -TypeName System.Data.DataTable
		$TablePhysicalPartitions = New-Object -TypeName System.Data.DataTable
		$TablePartitionSchemeParameters = New-Object -TypeName System.Data.DataTable
		$TableChecks = New-Object -TypeName System.Data.DataTable
		$TableColumns = New-Object -TypeName System.Data.DataTable
		$TableForeignKeys = New-Object -TypeName System.Data.DataTable
		$TableForeignKeyColumns = New-Object -TypeName System.Data.DataTable
		$TableFullTextIndexes = New-Object -TypeName System.Data.DataTable
		$TableFullTextIndexColumns = New-Object -TypeName System.Data.DataTable
		$TableIndexes = New-Object -TypeName System.Data.DataTable
		$TableIndexColumns = New-Object -TypeName System.Data.DataTable
		$TableIndexPhysicalPartitions = New-Object -TypeName System.Data.DataTable
		$TableIndexPartitionSchemeParameters = New-Object -TypeName System.Data.DataTable
		$TableStatistics = New-Object -TypeName System.Data.DataTable
		$TableStatisticsColumns = New-Object -TypeName System.Data.DataTable
		$TableTriggers = New-Object -TypeName System.Data.DataTable

		$Views = New-Object -TypeName System.Data.DataTable
		$ViewColumns = New-Object -TypeName System.Data.DataTable
		$ViewFullTextIndexes = New-Object -TypeName System.Data.DataTable
		$ViewFullTextIndexColumns = New-Object -TypeName System.Data.DataTable
		$ViewIndexes = New-Object -TypeName System.Data.DataTable
		$ViewIndexColumns = New-Object -TypeName System.Data.DataTable
		$ViewIndexPartitionSchemeParameters = New-Object -TypeName System.Data.DataTable
		$ViewIndexPhysicalPartitions = New-Object -TypeName System.Data.DataTable
		$ViewStatistics = New-Object -TypeName System.Data.DataTable
		$ViewStatisticsColumns = New-Object -TypeName System.Data.DataTable
		$ViewTriggers = New-Object -TypeName System.Data.DataTable

		$StoredProcedures = New-Object -TypeName System.Data.DataTable
		$StoredProcedureParameters = New-Object -TypeName System.Data.DataTable

		$UserDefinedFunctions = New-Object -TypeName System.Data.DataTable
		$UserDefinedFunctionChecks = New-Object -TypeName System.Data.DataTable
		$UserDefinedFunctionColumns = New-Object -TypeName System.Data.DataTable
		$UserDefinedFunctionDefaultConstraints = New-Object -TypeName System.Data.DataTable
		$UserDefinedFunctionIndexes = New-Object -TypeName System.Data.DataTable
		$UserDefinedFunctionIndexColumns = New-Object -TypeName System.Data.DataTable
		$UserDefinedFunctionParameters = New-Object -TypeName System.Data.DataTable


		if ($DatabaseName) {
			$Database += $Server.Databases | Where-Object { $DatabaseName -icontains $_.Name }
		} else {
			$Database = $Server.Databases
		}

		$Database | ForEach-Object {

			# Initialize variables for current iteration
			$DatabaseRoleMemberRole = $null
			$DbccLogInfo = $null
			$DbccCheckDbInfo = $null
			$LastKnownGoodDbccDate = $DbccEpoch

			# Assume the DB can be accessed if it's an Azure Server
			$IsAccessible = if ($DbEngineType -ieq $StandaloneDbEngine) { $_.IsAccessible } else { $true }


			# Run custom queries to gather info if the DB is accessible
			if ($IsAccessible -eq $true) {

				# Turn off SetDefaultInitFields at the server level - do this if you want to make things run poorly
				#$Server.SetDefaultInitFields($false)

				if ($DbEngineType -ieq $StandaloneDbEngine) {
					$DbccLogInfo = $_.ExecuteWithResults('DBCC LOGINFO WITH NO_INFOMSGS').Tables[0].Rows | Group-Object -Property FileId
				} else {
					$DbccLogInfo = $null
				}


				# NOTE ABOUT PREFETCH: 
				# 	I've found that the best balance of performance & number of calls to server is 
				#	using $Server.SetDefaultInitFields($true) and NOT using PrefetchObjects().
				# 	Your mileage may vary so I've left the code in here but commented out. Uncomment at your own risk.
				<#
				# Prefetch User and Schema information
				$_.PrefetchObjects($('Microsoft.SqlServer.Management.Smo.User' -as [Type]))
				$_.PrefetchObjects($('Microsoft.SqlServer.Management.Smo.Schema' -as [Type]))
				#>

				if ($IncludeObjectInformation -eq $true) {
					<#
					# Prefetch specific DB object information. See the note above about prefetch and performance
					$_.PrefetchObjects($('Microsoft.SqlServer.Management.Smo.ExtendedStoredProcedure' -as [Type]))
					$_.PrefetchObjects($('Microsoft.SqlServer.Management.Smo.SqlAssembly' -as [Type]))
					$_.PrefetchObjects($('Microsoft.SqlServer.Management.Smo.UserDefinedAggregate' -as [Type]))
					$_.PrefetchObjects($('Microsoft.SqlServer.Management.Smo.UserDefinedDataType' -as [Type]))
					$_.PrefetchObjects($('Microsoft.SqlServer.Management.Smo.UserDefinedTableType' -as [Type]))
					$_.PrefetchObjects($('Microsoft.SqlServer.Management.Smo.UserDefinedType' -as [Type]))
					$_.PrefetchObjects($('Microsoft.SqlServer.Management.Smo.XmlSchemaCollection' -as [Type]))
					$_.PrefetchObjects($('Microsoft.SqlServer.Management.Smo.Rule' -as [Type]))
					$_.PrefetchObjects($('Microsoft.SqlServer.Management.Smo.Default' -as [Type]))
					$_.PrefetchObjects($('Microsoft.SqlServer.Management.Smo.Sequence' -as [Type]))
					$_.PrefetchObjects($('Microsoft.SqlServer.Management.Smo.PartitionScheme' -as [Type]))
					$_.PrefetchObjects($('Microsoft.SqlServer.Management.Smo.PartitionFunction' -as [Type]))
					#>

					$ParameterHash = @{
						ServerVersion = $Server.Information.Version
						DatabaseEngineType = $DbEngineType
						IncludeSystemObjects = $IncludeSystemObjects
					}

					$Tables = $_.ExecuteWithResults((Get-TablePropertyQuery @ParameterHash)).Tables[0]
					$TableChecks = $_.ExecuteWithResults((Get-TableCheckQuery @ParameterHash)).Tables[0]
					$TableColumns = $_.ExecuteWithResults((Get-TableColumnQuery @ParameterHash)).Tables[0]
					$TableDefaultConstraints = $_.ExecuteWithResults((Get-TableColumnDefaultConstraintQuery @ParameterHash)).Tables[0]
					$TableForeignKeys = $_.ExecuteWithResults((Get-TableForeignKeyQuery @ParameterHash)).Tables[0]
					$TableForeignKeyColumns = $_.ExecuteWithResults((Get-TableForeignKeyColumnQuery @ParameterHash)).Tables[0]
					$TableIndexes = $_.ExecuteWithResults((Get-TableIndexQuery @ParameterHash)).Tables[0]
					$TableIndexColumns = $_.ExecuteWithResults((Get-TableIndexColumnQuery @ParameterHash)).Tables[0]
					$TableStatistics = $_.ExecuteWithResults((Get-TableStatisticsQuery @ParameterHash)).Tables[0]
					$TableStatisticsColumns = $_.ExecuteWithResults((Get-TableStatisticsColumnQuery @ParameterHash)).Tables[0]
					$TableTriggers = $_.ExecuteWithResults((Get-TableTriggerQuery @ParameterHash)).Tables[0]

					$Views = $_.ExecuteWithResults((Get-ViewPropertyQuery @ParameterHash)).Tables[0]
					$ViewColumns = $_.ExecuteWithResults((Get-ViewColumnQuery @ParameterHash)).Tables[0]
					$ViewIndexes = $_.ExecuteWithResults((Get-ViewIndexQuery @ParameterHash)).Tables[0]
					$ViewIndexColumns = $_.ExecuteWithResults((Get-ViewIndexColumnQuery @ParameterHash)).Tables[0]
					$ViewStatistics = $_.ExecuteWithResults((Get-ViewStatisticsQuery @ParameterHash)).Tables[0]
					$ViewStatisticsColumns = $_.ExecuteWithResults((Get-ViewStatisticsColumnQuery @ParameterHash)).Tables[0]
					$ViewTriggers = $_.ExecuteWithResults((Get-ViewTriggerQuery @ParameterHash)).Tables[0]

					$StoredProcedures = $_.ExecuteWithResults((Get-StoredProcedureQuery @ParameterHash)).Tables[0]
					$StoredProcedureParameters = $_.ExecuteWithResults((Get-StoredProcedureParameterQuery @ParameterHash)).Tables[0]

					$UserDefinedFunctions = $_.ExecuteWithResults((Get-UserDefinedFunctionQuery @ParameterHash)).Tables[0]
					$UserDefinedFunctionChecks = $_.ExecuteWithResults((Get-UserDefinedFunctionCheckQuery @ParameterHash)).Tables[0]
					$UserDefinedFunctionColumns = $_.ExecuteWithResults((Get-UserDefinedFunctionColumnQuery @ParameterHash)).Tables[0]
					$UserDefinedFunctionDefaultConstraints = $_.ExecuteWithResults((Get-UserDefinedFunctionDefaultConstraint @ParameterHash)).Tables[0]
					$UserDefinedFunctionIndexes = $_.ExecuteWithResults((Get-UserDefinedFunctionIndexQuery @ParameterHash)).Tables[0]
					$UserDefinedFunctionIndexColumns = $_.ExecuteWithResults((Get-UserDefinedFunctionIndexColumnQuery @ParameterHash)).Tables[0]
					$UserDefinedFunctionParameters = $_.ExecuteWithResults((Get-UserDefinedFunctionParameterQuery @ParameterHash)).Tables[0]


					# Partitioning not supported in SQL 2000 or Azure
					# Partitions and FullText Indexes on views not supported in SQL 2000 or Azure
					if (
						$DbEngineType -ieq $StandaloneDbEngine
					) {

						$TableFullTextIndexes = $_.ExecuteWithResults((Get-TableFullTextIndexQuery @ParameterHash)).Tables[0]
						$TableFullTextIndexColumns = $_.ExecuteWithResults((Get-TableFullTextIndexColumnQuery @ParameterHash)).Tables[0]

						if (
							$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
						) {
							$TablePhysicalPartitions = $_.ExecuteWithResults((Get-TablePhysicalPartitionQuery @ParameterHash)).Tables[0]
							$TablePartitionSchemeParameters = $_.ExecuteWithResults((Get-TablePartitionSchemeParameterQuery @ParameterHash)).Tables[0]
							$TableIndexPhysicalPartitions = $_.ExecuteWithResults((Get-TableIndexPhysicalPartitionQuery @ParameterHash)).Tables[0]
							$TableIndexPartitionSchemeParameters = $_.ExecuteWithResults((Get-TableIndexPartitionSchemeParameterQuery @ParameterHash)).Tables[0]

							$ViewFullTextIndexes = $_.ExecuteWithResults((Get-ViewFullTextIndexQuery @ParameterHash)).Tables[0]
							$ViewFullTextIndexColumns = $_.ExecuteWithResults((Get-ViewFullTextIndexColumnQuery @ParameterHash)).Tables[0]
							$ViewIndexPhysicalPartitions = $_.ExecuteWithResults((Get-ViewIndexPhysicalPartitionQuery @ParameterHash)).Tables[0]
							$ViewIndexPartitionSchemeParameters = $_.ExecuteWithResults((Get-ViewIndexPartitionSchemeParameterQuery @ParameterHash)).Tables[0]
						}
					}

				}


				if ($($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {

					$DatabaseRoleMemberRole = $_.ExecuteWithResults(
						'SELECT ' + `
						' p1.name AS [RoleName],' + `
						' p2.name AS [MemberRoleName]' + `
						' FROM sys.database_role_members AS r' + `
						' INNER JOIN sys.database_principals AS p1 ON p1.principal_id = r.role_principal_id' + `
						' INNER JOIN sys.database_principals AS p2 ON p2.principal_id = r.member_principal_id' + `
						' WHERE p2.type = ''R''' 
					).Tables[0].Rows


					if ($DbEngineType -ieq $StandaloneDbEngine) {

						$DbccCheckDbInfo = $_.ExecuteWithResults('DBCC DBINFO WITH TABLERESULTS, NO_INFOMSGS').Tables[0].Rows
						try {
							$DbccCheckDbInfo | Where-Object { 
								$_.Field -ieq 'dbi_dbccLastKnownGood' 
							} | ForEach-Object { 
								$LastKnownGoodDbccDate = [DateTime]$_.VALUE 
							}
						} catch {
							$LastKnownGoodDbccDate = $DbccEpoch
						}
					} else {
						$LastKnownGoodDbccDate = $DbccEpoch
					}
				} 
			}


			<#
					# See http://msdn.microsoft.com/en-us/library/dd364983.aspx for more info on primary keys and datatable performance
					# In testing I found that adding primary keys didn't really make much of a difference. Your mileage may vary - uncomment at your own risk!

					$TableChecks.PrimaryKey = @(
						$TableChecks.Columns['TableID'],
						$TableChecks.Columns['SchemaID'],
						$TableChecks.Columns['ID']
					)

					$TableColumns.PrimaryKey = @(
						$TableColumns.Columns['TableID'],
						$TableColumns.Columns['SchemaID'],
						$TableColumns.Columns['ID']
					)

					$TableDefaultConstraints.PrimaryKey = @(
						$TableDefaultConstraints.Columns['TableID'],
						$TableDefaultConstraints.Columns['SchemaID'],
						$TableDefaultConstraints.Columns['ID']
					)
					
					$TableForeignKeys.PrimaryKey = @(
						$TableForeignKeys.Columns['TableID'],
						$TableForeignKeys.Columns['SchemaID'],
						$TableForeignKeys.Columns['ID']
					)					
					$TableForeignKeyColumns.PrimaryKey = @(
						$TableForeignKeyColumns.Columns['TableID'],
						$TableForeignKeyColumns.Columns['SchemaID'],
						$TableForeignKeyColumns.Columns['ForeignKeyID'],
						$TableForeignKeyColumns.Columns['ID']
					)

					$TableFullTextIndexes.PrimaryKey = @(
						$TableFullTextIndexes.Columns['TableID'],
						$TableFullTextIndexes.Columns['SchemaID']
					)
					$TableFullTextIndexColumns.PrimaryKey = @(
						$TableFullTextIndexColumns.Columns['TableID'],
						$TableFullTextIndexColumns.Columns['SchemaID']
					)

					$TableIndexes.PrimaryKey = @(
						$TableIndexes.Columns['TableID'],
						$TableIndexes.Columns['SchemaID'],
						$TableIndexes.Columns['ID']
					)
					$TableIndexColumns.PrimaryKey = @(
						$TableIndexColumns.Columns['TableID'],
						$TableIndexColumns.Columns['SchemaID'],
						$TableIndexColumns.Columns['IndexID'],
						$TableIndexColumns.Columns['ID']
					)

					$TableStatistics.PrimaryKey = @(
						$TableStatistics.Columns['TableID'],
						$TableStatistics.Columns['SchemaID'],
						$TableStatistics.Columns['ID']
					)
					$TableStatisticsColumns.PrimaryKey = @(
						$TableStatisticsColumns.Columns['TableID'],
						$TableStatisticsColumns.Columns['SchemaID'],
						$TableStatisticsColumns.Columns['StatisticID'],
						$TableStatisticsColumns.Columns['ID']
					)
					
					$TableTriggers.PrimaryKey = @(
						$TableTriggers.Columns['TableID'],
						$TableTriggers.Columns['SchemaID'],
						$TableTriggers.Columns['ID']
					)
#>
			<#
					$ViewColumns.PrimaryKey = @(
						$ViewColumns.Columns['ViewID'],
						$ViewColumns.Columns['SchemaID'],
						$ViewColumns.Columns['ID']
					)
					$ViewIndexes.PrimaryKey = @(
						$ViewIndexes.Columns['ViewID'],
						$ViewIndexes.Columns['SchemaID'],
						$ViewIndexes.Columns['ID']
					)
					$ViewIndexColumns.PrimaryKey = @(
						$ViewIndexColumns.Columns['ViewID'],
						$ViewIndexColumns.Columns['SchemaID'],
						$ViewIndexColumns.Columns['IndexID'],
						$ViewIndexColumns.Columns['ID']
					)
					$ViewStatistics.PrimaryKey = @(
						$ViewStatistics.Columns['ViewID'],
						$ViewStatistics.Columns['SchemaID'],
						$ViewStatistics.Columns['ID']
					)
					$ViewStatisticsColumns.PrimaryKey = @(
						$ViewStatisticsColumns.Columns['ViewID'],
						$ViewStatisticsColumns.Columns['SchemaID'],
						$ViewStatisticsColumns.Columns['StatisticID'],
						$ViewStatisticsColumns.Columns['ID']
					)
					$ViewTriggers.PrimaryKey = @(
						$ViewTriggers.Columns['ViewID'],
						$ViewTriggers.Columns['SchemaID'],
						$ViewTriggers.Columns['ID']
					)
#>
			<#
					$StoredProcedureParameters.PrimaryKey = @(
						$StoredProcedureParameters.Columns['StoredProcedureID'],
						$StoredProcedureParameters.Columns['SchemaID'],
						$StoredProcedureParameters.Columns['ID']
					)
#>



			Write-Output (
				New-Object -TypeName psobject -Property @{
					Name = $_.Name # System.String Name {get;set;}
					ID = $_.ID # System.Int32 ID {get;}
					Properties = New-Object -TypeName psobject -Property @{
						General = New-Object -TypeName psobject -Property @{
							Backup = New-Object -TypeName psobject -Property @{
								LastFullBackupDate = if (($_.LastBackupDate) -and ($_.LastBackupDate.CompareTo($SmoEpoch) -le 0)) { $null } else { $_.LastBackupDate } # System.DateTime LastBackupDate {get;}
								LastDifferentialBackupDate = if (($_.LastDifferentialBackupDate) -and ($_.LastDifferentialBackupDate.CompareTo($SmoEpoch) -le 0)) { $null } else { $_.LastDifferentialBackupDate } # System.DateTime LastDifferentialBackupDate {get;}
								LastLogBackupDate = if (($_.LastLogBackupDate) -and ($_.LastLogBackupDate.CompareTo($SmoEpoch) -le 0)) { $null } else { $_.LastLogBackupDate } # System.DateTime LastLogBackupDate {get;}
							}
							Database = New-Object -TypeName psobject -Property @{
								Name = $_.Name # System.String Name {get;set;}
								Status = if ($_.Status) { $_.Status.ToString() } else { $null } # Microsoft.SqlServer.Management.Smo.DatabaseStatus Status {get;}
								Owner = $_.Owner # System.String Owner {get;}
								DateCreated = $_.CreateDate # System.DateTime CreateDate {get;}
								SizeMB = $_.Size # System.Double Size {get;}
								SpaceAvailableKB = $_.SpaceAvailable # System.Double SpaceAvailable {get;}
								NumberOfUsers = $_.ActiveConnections # System.Int32 ActiveConnections {get;}

								# Not part of the SSMS GUI but I think these properties belong here
								IsFullTextEnabled = $_.IsFullTextEnabled # System.Boolean IsFullTextEnabled {get;set;}
								IsDatabaseSnapshot = $_.IsDatabaseSnapshot # System.Boolean IsDatabaseSnapshot {get;}
								IsDatabaseSnapshotBase = $_.IsDatabaseSnapshotBase # System.Boolean IsDatabaseSnapshotBase {get;}
								DatabaseSnapshotBaseName = $_.DatabaseSnapshotBaseName # System.String DatabaseSnapshotBaseName {get;set;}
								IsMailHost = $_.IsMailHost # System.Boolean IsMailHost {get;}
								IsManagementDataWarehouse = if (($Server.Information.Version).CompareTo($SQLServer2008) -ge 0) { $_.IsManagementDataWarehouse } else { $null } # System.Boolean IsManagementDataWarehouse {get;}
								IsMirroringEnabled = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) { $_.IsMirroringEnabled } else { $null } # System.Boolean IsMirroringEnabled {get;}
								IsAvailabilityGroupMember = if (($Server.Information.Version).CompareTo($SQLServer2012) -ge 0) {
									if ($_.AvailabilityGroupName) { $true } else { $false }
								} else {
									$false
								}
								IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
								ReplicationOptions = [String](Get-ReplicationOptionsValue -ReplicationOptions $_.ReplicationOptions) # Microsoft.SqlServer.Management.Smo.ReplicationOptions ReplicationOptions {get;}

								# This should probably be moved to the "performance" section (once its created)
								LastKnownGoodDbccDate = if ($LastKnownGoodDbccDate.CompareTo($DbccEpoch) -le 0) { $null } else { $LastKnownGoodDbccDate }
							}
							Maintenance = New-Object -TypeName psobject -Property @{
								Collation = $_.Collation # System.String Collation {get;set;}
							}
						}
						Files = New-Object -TypeName psobject -Property @{

							#DatabaseName = $_.Name		# Make this a script property after the fact?
							#Owner = $_.Owner			# Make this a script property after the fact?
							#UseFullTextIndexing = $null

							DatabaseFiles = if (
								$IsAccessible -eq $true -and
								$DbEngineType -ieq $StandaloneDbEngine # Can't enumerate files for Azure Databases
							) {
								$_.FileGroups | ForEach-Object {

									$FileType = if ($_.IsFileStream) { 'FILESTREAM Data' } else { 'Rows Data' }

									$_.Files | ForEach-Object {
										New-Object -TypeName psobject -Property @{
											LogicalName = $_.Name # System.String Name {get;set;}
											FileType = $FileType
											#Filegroup =  $null # Make this a script property?
											SizeKB = $_.Size # System.Double Size {get;set;}
											Growth = $_.Growth # System.Double Growth {get;set;}
											GrowthType = $_.GrowthType.ToString() # Microsoft.SqlServer.Management.Smo.FileGrowthType GrowthType {get;set;}
											MaxSizeKB = $_.MaxSize # System.Double MaxSize {get;set;}
											Path = [System.IO.Path]::GetDirectoryName($_.FileName) # System.String FileName {get;set;}
											FileName = [System.IO.Path]::GetFileName($_.FileName) # System.String FileName {get;set;}

											AvailableSpaceKB = $_.AvailableSpace
											#BytesReadFromDisk = $_.BytesReadFromDisk	# System.Int64 BytesReadFromDisk {get;}
											#BytesWrittenToDisk = $_.BytesWrittenToDisk	# System.Int64 BytesWrittenToDisk {get;}
											ID = $_.ID # System.Int32 ID {get;}
											IsOffline = $_.IsOffline # System.Boolean IsOffline {get;}
											IsPrimaryFile = $_.IsPrimaryFile # System.Boolean IsPrimaryFile {get;set;}
											IsReadOnly = $_.IsReadOnly # System.Boolean IsReadOnly {get;}
											IsReadOnlyMedia = $_.IsReadOnlyMedia # System.Boolean IsReadOnlyMedia {get;}
											IsSparse = $_.IsSparse # System.Boolean IsSparse {get;}
											#NumberOfDiskReads = $_.NumberOfDiskReads	# System.Int64 NumberOfDiskReads {get;}
											#NumberOfDiskWrites = $_.NumberOfDiskWrites	# System.Int64 NumberOfDiskWrites {get;}
											#Parent = $_.Parent
											#Properties = $_.Properties
											#State = $_.State
											#Urn = $_.Urn
											UsedSpaceKB = $_.UsedSpace # System.Double UsedSpace {get;}
											#UserData = $_.UserData
											VolumeFreeSpaceBytes = $_.VolumeFreeSpace # System.Int64 VolumeFreeSpace {get;}
											VlfCount = $null
										}
									} | Add-Member -MemberType NoteProperty -Name Filegroup -Value $_.Name -PassThru
								}

								$_.LogFiles | ForEach-Object {
									$FileId = $_.ID

									New-Object -TypeName psobject -Property @{
										LogicalName = $_.Name # System.String Name {get;set;}
										FileType = 'Log'
										Filegroup = 'Not Applicable'
										SizeKB = $_.Size # System.Double Size {get;set;}
										Growth = $_.Growth # System.Double Growth {get;set;}
										GrowthType = $_.GrowthType.ToString() # Microsoft.SqlServer.Management.Smo.FileGrowthType GrowthType {get;set;}
										MaxSizeKB = $_.MaxSize # System.Double MaxSize {get;set;}
										Path = [System.IO.Path]::GetDirectoryName($_.FileName) # System.String FileName {get;set;}
										FileName = [System.IO.Path]::GetFileName($_.FileName) # System.String FileName {get;set;}

										AvailableSpaceKB = $null
										#BytesReadFromDisk = $_.BytesReadFromDisk	# System.Int64 BytesReadFromDisk {get;}
										#BytesWrittenToDisk = $_.BytesWrittenToDisk	# System.Int64 BytesWrittenToDisk {get;}
										ID = $_.ID # System.Int32 ID {get;}
										IsOffline = $_.IsOffline # System.Boolean IsOffline {get;}
										IsPrimaryFile = $_.IsPrimaryFile # System.Boolean IsPrimaryFile {get;set;}
										IsReadOnly = $_.IsReadOnly # System.Boolean IsReadOnly {get;}
										IsReadOnlyMedia = $_.IsReadOnlyMedia # System.Boolean IsReadOnlyMedia {get;}
										IsSparse = $_.IsSparse # System.Boolean IsSparse {get;}
										#NumberOfDiskReads = $_.NumberOfDiskReads	# System.Int64 NumberOfDiskReads {get;}
										#NumberOfDiskWrites = $_.NumberOfDiskWrites	# System.Int64 NumberOfDiskWrites {get;}
										#Parent = $_.Parent
										#Properties = $_.Properties
										#State = $_.State
										#Urn = $_.Urn
										UsedSpaceKB = $_.UsedSpace # System.Double UsedSpace {get;}
										#UserData = $_.UserData
										VolumeFreeSpaceBytes = $_.VolumeFreeSpace # System.Int64 VolumeFreeSpace {get;}
										VlfCount = ($DbccLogInfo | Where-Object { $_.Name -eq $FileId } | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
									}
								}
							} else {
								$null
								#@()
							}
						}
						FileGroups = New-Object -TypeName psobject -Property @{
							Rows = if (
								$IsAccessible -eq $true -and
								$DbEngineType -ieq $StandaloneDbEngine # Can't enumerate filegroups for Azure Databases
							) {
								@() + (
									$_.FileGroups | Where-Object { $_.IsFileStream -ne $true } | ForEach-Object {
										New-Object -TypeName psobject -Property @{
											Name = $_.Name # System.String Name {get;set;}
											Files = @($_.Files).Count # Should really be called FileCount
											ReadOnly = $_.ReadOnly # System.Boolean ReadOnly {get;set;}
											IsDefault = $_.IsDefault # System.Boolean IsDefault {get;set;}
											#Files = $_.Files
											ID = $_.ID # System.Int32 ID {get;}
											#Parent = $_.Parent
											#Properties = $_.Properties
											#Size = $_.Size
											#State = $_.State
											#Urn = $_.Urn
											#UserData = $_.UserData
										}
									}
								)
							} else {
								$null
								#@()
							}
							Filestream = if (
								$IsAccessible -eq $true -and 
								$($Server.Information.Version).CompareTo($SQLServer2008) -ge 0 -and
								$DbEngineType -ieq $StandaloneDbEngine # Can't enumerate filegroups for Azure Databases
							) {
								@() + (
									$_.FileGroups | Where-Object { $_.IsFileStream -eq $true } | ForEach-Object {
										New-Object -TypeName psobject -Property @{
											Name = $_.Name # System.String Name {get;set;}
											Files = @($_.Files).Count
											ReadOnly = $_.ReadOnly # System.Boolean ReadOnly {get;set;}
											IsDefault = $_.IsDefault # System.Boolean IsDefault {get;set;}
											ID = $_.ID # System.Int32 ID {get;}
										}
									}
								)
							} else {
								$null
								#@()
							}
						}
						Options = New-Object -TypeName psobject -Property @{
							Collation = $_.Collation # System.String Collation {get;set;}	# Duplicated in the general tab
							RecoveryModel = if ($_.DatabaseOptions.RecoveryModel) { $_.DatabaseOptions.RecoveryModel.ToString() } else { $null } # Microsoft.SqlServer.Management.Smo.RecoveryModel RecoveryModel {get;set;}
							CompatibilityLevel = switch ($_.CompatibilityLevel) {
								$($CompatibilityLevel::Version60) { 'SQL Server 6.0 (60)' }
								$($CompatibilityLevel::Version65) { 'SQL Server 6.5 (65)' }
								$($CompatibilityLevel::Version70) { 'SQL Server 7.0 (70)' }
								$($CompatibilityLevel::Version80) { 'SQL Server 2000 (80)' }
								$($CompatibilityLevel::Version90) { 'SQL Server 2005 (90)' }
								$($CompatibilityLevel::Version100) { 'SQL Server 2008 (100)' }
								$($CompatibilityLevel::Version110) { 'SQL Server 2012 (110)' }
								$null { 'Unknown' }
								default { $_.ToString() }
							} # Microsoft.SqlServer.Management.Smo.CompatibilityLevel CompatibilityLevel {get;set;}
							ContainmentType = if ((($Server.Information.Version).CompareTo($SQLServer2012) -ge 0) -and ($_.ContainmentType)) { $_.ContainmentType.ToString() } else { $null } # Microsoft.SqlServer.Management.Smo.ContainmentType ContainmentType {get;set;}
							OtherOptions = New-Object -TypeName psobject -Property @{
								Automatic = New-Object -TypeName psobject -Property @{
									AutoClose = $_.AutoClose # System.Boolean AutoClose {get;set;}
									AutoCreateStatistics = $_.AutoCreateStatisticsEnabled # System.Boolean AutoCreateStatisticsEnabled {get;set;}
									AutoShrink = $_.AutoShrink # System.Boolean AutoShrink {get;set;}
									AutoUpdateStatistics = $_.AutoUpdateStatisticsEnabled # System.Boolean AutoUpdateStatisticsEnabled {get;set;}
									AutoUpdateStatisticsAsync = $_.AutoUpdateStatisticsAsync # System.Boolean AutoUpdateStatisticsAsync {get;set;}
								}
								Containment = if (($Server.Information.Version).CompareTo($SQLServer2012) -ge 0) {
									New-Object -TypeName psobject -Property @{
										DefaultFullTextLanguage = if ($_.DefaultFullTextLanguage) { $_.DefaultFullTextLanguage.Name } else { $null } # Microsoft.SqlServer.Management.Smo.DefaultLanguage DefaultFullTextLanguage {get;}
										DefaultLanguage = if ($_.DefaultLanguage) { $_.DefaultLanguage.Name } else { $_.DefaultLanguage } # Microsoft.SqlServer.Management.Smo.DefaultLanguage DefaultLanguage {get;}
										NestedTriggersEnabled = $_.NestedTriggersEnabled # System.Boolean NestedTriggersEnabled {get;set;}
										TransformNoiseWords = $_.TransformNoiseWords # System.Boolean TransformNoiseWords {get;set;}
										TwoDigitYearCutoff = $_.TwoDigitYearCutoff # System.Int32 TwoDigitYearCutoff {get;set;}
									}
								} else {
									New-Object -TypeName psobject -Property @{
										DefaultFullTextLanguage = $null
										DefaultLanguage = $null
										NestedTriggersEnabled = $null
										TransformNoiseWords = $null
										TwoDigitYearCutoff = $null
									} }
								Cursor = New-Object -TypeName psobject -Property @{
									CloseCursorsOnCommitEnabled = $_.CloseCursorsOnCommitEnabled # System.Boolean CloseCursorsOnCommitEnabled {get;set;}
									LocalCursorsDefault = $_.LocalCursorsDefault # System.Boolean LocalCursorsDefault {get;set;}
								}
								Filestream = if ((($Server.Information.Version).CompareTo($SQLServer2008) -ge 0) -and ($SmoMajorVersion -ge 10)) {
									New-Object -TypeName psobject -Property @{
										FilestreamDirectoryName = $_.FilestreamDirectoryName # System.String FilestreamDirectoryName {get;set;}
										FilestreamNonTransactedAccess = $_.FilestreamNonTransactedAccess # Microsoft.SqlServer.Management.Smo.FilestreamNonTransactedAccessType FilestreamNonTransactedAccess {get;set;}
									}
								} else {
									New-Object -TypeName psobject -Property @{
										FilestreamDirectoryName = $null
										FilestreamNonTransactedAccess = $null
									}
								}
								Miscellaneous = New-Object -TypeName psobject -Property @{
									SnapshotIsolation = [String](Get-SnapshotIsolationStateValue -SnapshotIsolationState $_.SnapshotIsolationState) # Microsoft.SqlServer.Management.Smo.SnapshotIsolationState SnapshotIsolationState {get;}
									AnsiNullDefault = $_.AnsiNullDefault # System.Boolean AnsiNullDefault {get;set;}
									AnsiNullsEnabled = $_.AnsiNullsEnabled # System.Boolean AnsiNullsEnabled {get;set;}
									AnsiPaddingEnabled = $_.AnsiPaddingEnabled # System.Boolean AnsiPaddingEnabled {get;set;}
									AnsiWarningsEnabled = $_.AnsiWarningsEnabled # System.Boolean AnsiWarningsEnabled {get;set;}
									ArithmeticAbortEnabled = $_.ArithmeticAbortEnabled # System.Boolean ArithmeticAbortEnabled {get;set;}
									ConcatenateNullYieldsNull = $_.ConcatenateNullYieldsNull # System.Boolean ConcatenateNullYieldsNull {get;set;}
									DatabaseOwnershipChaining = $_.DatabaseOwnershipChaining # System.Boolean DatabaseOwnershipChaining {get;set;}
									DateCorrelationOptimization = $_.DateCorrelationOptimization # System.Boolean DateCorrelationOptimization {get;set;}
									IsReadCommittedSnapshotOn = $_.IsReadCommittedSnapshotOn # System.Boolean IsReadCommittedSnapshotOn {get;set;}
									NumericRoundAbortEnabled = $_.NumericRoundAbortEnabled # System.Boolean NumericRoundAbortEnabled {get;set;}
									Parameterization = switch ($_.IsParameterizationForced) { 
										$true { 'Forced' }
										default { 'Simple' }
									} # System.Boolean IsParameterizationForced {get;set;}
									QuotedIdentifiersEnabled = $_.QuotedIdentifiersEnabled # System.Boolean QuotedIdentifiersEnabled {get;set;}
									RecursiveTriggersEnabled = $_.RecursiveTriggersEnabled # System.Boolean RecursiveTriggersEnabled {get;set;}
									Trustworthy = $_.Trustworthy # System.Boolean Trustworthy {get;set;}
									VarDecimalStorageFormatEnabled = $_.IsVarDecimalStorageFormatEnabled # System.Boolean IsVarDecimalStorageFormatEnabled {get;set;}
								}
								Recovery = New-Object -TypeName psobject -Property @{
									PageVerify = [String](Get-PageVerifyValue -PageVerify $_.PageVerify) # Microsoft.SqlServer.Management.Smo.PageVerify PageVerify {get;set;}
									TargetRecoveryTimeSeconds = $_.TargetRecoveryTime # System.Int32 TargetRecoveryTime {get;set;}
								}
								ServiceBroker = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
									New-Object -TypeName psobject -Property @{
										BrokerEnabled = $_.BrokerEnabled # System.Boolean BrokerEnabled {get;set;}
										HonorBrokerPriority = $_.HonorBrokerPriority # System.Boolean HonorBrokerPriority {get;set;}
										ServiceBrokerIdentifier = $_.ServiceBrokerGuid # System.Guid ServiceBrokerGuid {get;}
									}
								} else {
									New-Object -TypeName psobject -Property @{
										BrokerEnabled = $null
										HonorBrokerPriority = $null
										ServiceBrokerIdentifier = $null
									} 
								}
								State = New-Object -TypeName psobject -Property @{
									DatabaseReadOnly = $_.ReadOnly # System.Boolean ReadOnly {get;set;}
									#DatabaseState = ## NOT SURE WHERE THIS COMES FROM. PUNT FOR NOW
									EncryptionEnabled = $_.EncryptionEnabled # System.Boolean EncryptionEnabled {get;set;}
									RestrictAccess = [String](Get-DatabaseUserAccessValue -UserAccess $_.UserAccess) # Microsoft.SqlServer.Management.Smo.DatabaseUserAccess UserAccess {get;set;}
								}
							}

						}
						ChangeTracking = if (($Server.Information.Version).CompareTo($SQLServer2008) -ge 0) {
							New-Object -TypeName psobject -Property @{
								IsEnabled = $_.ChangeTrackingEnabled # System.Boolean ChangeTrackingEnabled {get;set;}
								RetentionPeriod = $_.ChangeTrackingRetentionPeriod # System.Int32 ChangeTrackingRetentionPeriod {get;set;}
								RetentionPeriodUnits = $_.ChangeTrackingRetentionPeriodUnits # Microsoft.SqlServer.Management.Smo.RetentionPeriodUnits ChangeTrackingRetentionPeriodUnits {get;set;}
								AutoCleanUp = $_.ChangeTrackingAutoCleanUp # System.Boolean ChangeTrackingAutoCleanUp {get;set;}
							}
						} else {
							New-Object -TypeName psobject -Property @{
								IsEnabled = $null
								RetentionPeriod = $null
								RetentionPeriodUnits = $null
								AutoCleanUp = $null
							}
						}

						Permissions = if (($IsAccessible -eq $true) -and (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0)) {
							@() + $(
								$_.EnumDatabasePermissions() | ForEach-Object {
									New-Object -TypeName psobject -Property @{
										ColumnName = $_.ColumnName # System.String ColumnName {get;}
										Grantee = $_.Grantee # System.String Grantee {get;}
										GranteeType = [String](Get-PrincipalTypeValue -PrincipalType $_.GranteeType) # Microsoft.SqlServer.Management.Smo.PrincipalType GranteeType {get;}
										Grantor = $_.Grantor # System.String Grantor {get;}
										GrantorType = [String](Get-PrincipalTypeValue -PrincipalType $_.GrantorType) # Microsoft.SqlServer.Management.Smo.PrincipalType GrantorType {get;}
										ObjectClass = [String](Get-ObjectClassValue -ObjectClass $_.ObjectClass) # Microsoft.SqlServer.Management.Smo.ObjectClass ObjectClass {get;}
										ObjectID = $_.ObjectID # System.Int32 ObjectID {get;}
										ObjectName = if (($_.ObjectName -ieq $Server.Name) -and ($ServerName -ine $Server.Name)) { $ServerName } else { $_.ObjectName } # System.String ObjectName {get;}
										ObjectSchema = $_.ObjectSchema # System.String ObjectSchema {get;}
										PermissionState = [String](Get-PermissionStateValue -PermissionState $_.PermissionState) # Microsoft.SqlServer.Management.Smo.PermissionState PermissionState {get;}
										PermissionType = if ($_.PermissionType) { $_.PermissionType.ToString() } else { $null } # Microsoft.SqlServer.Management.Smo.DatabasePermissionSet PermissionType {get;}
									}
								}
								if ($IncludeObjectPermissions -eq $true) {
									# SYS and INFORMATION_SCHEMA schema only included if $IncludeSystemObjects is $true
									$_.EnumObjectPermissions() | Where-Object { 
										$IncludeSystemObjects -eq $true -or
										(
											$_.ObjectSchema -ine 'sys' -and
											$_.ObjectSchema -ine 'INFORMATION_SCHEMA'
										)
									} | ForEach-Object { 
										New-Object -TypeName psobject -Property @{
											ColumnName = $_.ColumnName # System.String ColumnName {get;}
											Grantee = $_.Grantee # System.String Grantee {get;}
											GranteeType = [String](Get-PrincipalTypeValue -PrincipalType $_.GranteeType) # Microsoft.SqlServer.Management.Smo.PrincipalType GranteeType {get;}
											Grantor = $_.Grantor # System.String Grantor {get;}
											GrantorType = [String](Get-PrincipalTypeValue -PrincipalType $_.GrantorType) # Microsoft.SqlServer.Management.Smo.PrincipalType GrantorType {get;}
											ObjectClass = [String](Get-ObjectClassValue -ObjectClass $_.ObjectClass) # Microsoft.SqlServer.Management.Smo.ObjectClass ObjectClass {get;}
											ObjectID = $_.ObjectID # System.Int32 ObjectID {get;}
											ObjectName = $_.ObjectName # System.String ObjectName {get;}
											ObjectSchema = $_.ObjectSchema # System.String ObjectSchema {get;}
											PermissionState = [String](Get-PermissionStateValue -PermissionState $_.PermissionState) # Microsoft.SqlServer.Management.Smo.PermissionState PermissionState {get;}
											PermissionType = if ($_.PermissionType) { $_.PermissionType.ToString() } else { $null } # Microsoft.SqlServer.Management.Smo.ObjectPermissionSet PermissionType {get;}
										}
									} 
								}
							)
						} else {
							# TODO: Get permissions from SQL 2000
							$null
						}

						Mirroring = if (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0) {
							New-Object -TypeName psobject -Property @{
								IsEnabled = $_.IsMirroringEnabled # System.Boolean IsMirroringEnabled {get;}
								MirroringFailoverLogSequenceNumber = $_.MirroringFailoverLogSequenceNumber # System.Decimal MirroringFailoverLogSequenceNumber {get;}
								MirroringID = $_.MirroringID # System.Guid MirroringID {get;}
								MirroringPartner = $_.MirroringPartner # System.String MirroringPartner {get;set;}
								MirroringPartnerInstance = $_.MirroringPartnerInstance # System.String MirroringPartnerInstance {get;}
								MirroringRedoQueueMaxSizeKB = $_.MirroringRedoQueueMaxSize # System.Int32 MirroringRedoQueueMaxSize {get;}
								MirroringRoleSequence = $_.MirroringRoleSequence # System.Int32 MirroringRoleSequence {get;}
								MirroringSafetyLevel = [String](Get-MirroringSafetyLevelValue -MirroringSafetyLevel $_.MirroringSafetyLevel) # Microsoft.SqlServer.Management.Smo.MirroringSafetyLevel MirroringSafetyLevel {get;set;}
								MirroringSafetySequence = $_.MirroringSafetySequence # System.Int32 MirroringSafetySequence {get;}
								MirroringStatus = [String](Get-MirroringStatusValue -MirroringStatus $_.MirroringStatus) # Microsoft.SqlServer.Management.Smo.MirroringStatus MirroringStatus {get;}
								MirroringTimeoutSeconds = $_.MirroringTimeout # System.Int32 MirroringTimeout {get;set;}
								MirroringWitness = $_.MirroringWitness # System.String MirroringWitness {get;set;}
								MirroringWitnessStatus = [String](Get-MirroringWitnessStatusValue -MirroringWitnessStatus $_.MirroringWitnessStatus) # Microsoft.SqlServer.Management.Smo.MirroringWitnessStatus MirroringWitnessStatus {get;}
							}
						} else {
							New-Object -TypeName psobject -Property @{
								IsEnabled = $null
								MirroringFailoverLogSequenceNumber = $null
								MirroringID = $null
								MirroringPartner = $null
								MirroringPartnerInstance = $null
								MirroringRedoQueueMaxSizeKB = $null
								MirroringRoleSequence = $null
								MirroringSafetyLevel = $null
								MirroringSafetySequence = $null
								MirroringStatus = $null
								MirroringTimeoutSeconds = $null
								MirroringWitness = $null
								MirroringWitnessStatus = $null
							}
						}
						AlwaysOn = if ((($Server.Information.Version).CompareTo($SQLServer2012) -ge 0) -and ($SmoMajorVersion -ge 11)) {
							New-Object -TypeName psobject -Property @{
								IsAvailabilityGroupMember = if ($_.AvailabilityGroupName) { $true } else { $false }
								AvailabilityDatabaseSynchronizationState = [String](Get-AvailabilityDatabaseSynchronizationStatusValue -AvailabilityDatabaseSynchronizationStatus $_.AvailabilityDatabaseSynchronizationState) # Microsoft.SqlServer.Management.Smo.AvailabilityDatabaseSynchronizationState AvailabilityDatabaseSynchronizationState {get;}
								AvailabilityGroupName = $_.AvailabilityGroupName # System.String AvailabilityGroupName {get;}
							}
						} else {
							New-Object -TypeName psobject -Property @{
								IsAvailabilityGroupMember = $null
								AvailabilityDatabaseSynchronizationState = $null
								AvailabilityGroupName = $null
							}
						}
					}

					#region
					Tables = if (($IsAccessible -eq $true) -and ($IncludeObjectInformation -eq $true)) {
						@() + (
							#$Tables.Rows | Where-Object { $_.ID } | ForEach-Object {
							$Tables.Rows | ForEach-Object {
								New-Object -TypeName PSObject -Property @{
									Schema = $_.Schema # System.String Schema {get;set;}
									Name = $_.Name # System.String Name {get;set;}
									ID = $_.ID # System.Int32 ID {get;}

									Properties = New-Object -TypeName PSObject -Property @{
										General = New-Object -TypeName PSObject -Property @{
											Description = New-Object -TypeName PSObject -Property @{
												Schema = $_.Schema # System.String Schema {get;set;}
												Name = $_.Name # System.String Name {get;set;}
												ID = $_.ID # System.Int32 ID {get;}
												CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
												DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}

												# Properties exposed by SMO but not in the SSMS GUI...these probably belong here if they were
												FakeSystemTable = $_.FakeSystemTable # System.Boolean FakeSystemTable {get;}
												HasAfterTrigger = $_.HasAfterTrigger # System.Boolean HasAfterTrigger {get;}
												HasClusteredIndex = $_.HasClusteredIndex # System.Boolean HasClusteredIndex {get;}
												HasCompressedPartitions = $_.HasCompressedPartitions # System.Boolean HasCompressedPartitions {get;}
												HasDeleteTrigger = $_.HasDeleteTrigger # System.Boolean HasDeleteTrigger {get;}
												HasIndex = $_.HasIndex # System.Boolean HasIndex {get;}
												HasInsertTrigger = $_.HasInsertTrigger # System.Boolean HasInsertTrigger {get;}
												HasInsteadOfTrigger = $_.HasInsteadOfTrigger # System.Boolean HasInsteadOfTrigger {get;}
												HasUpdateTrigger = $_.HasUpdateTrigger # System.Boolean HasUpdateTrigger {get;}
												IsIndexable = $_.IsIndexable # System.Boolean IsIndexable {get;}
												IsSchemaOwned = $_.IsSchemaOwned # System.Boolean IsSchemaOwned {get;}
												IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}

											}
											Options = New-Object -TypeName PSObject -Property @{
												QuotedIdentifier = $_.QuotedIdentifierStatus # System.Boolean QuotedIdentifierStatus {get;set;}
												AnsiNulls = $_.AnsiNullsStatus # System.Boolean AnsiNullsStatus {get;set;}

												# Properties exposed by SMO but not in the SSMS GUI...these probably belong here if they were
												IsFileTable = $_.IsFileTable # System.Boolean IsFileTable {get;set;}
												LockEscalation = $_.LockEscalation # Microsoft.SqlServer.Management.Smo.LockEscalationType LockEscalation {get;set;}
												MaximumDegreeOfParallelism = $_.MaximumDegreeOfParallelism # System.Int32 MaximumDegreeOfParallelism {get;set;}
												OnlineHeapOperation = $_.OnlineHeapOperation # System.Boolean OnlineHeapOperation {get;set;}
											}
											Replication = New-Object -TypeName PSObject -Property @{
												IsReplicated = $_.Replicated # System.Boolean Replicated {get;}
											}
										}
										#Permissions
										ChangeTracking = New-Object -TypeName PSObject -Property @{
											IsEnabled = $_.ChangeTrackingEnabled # System.Boolean ChangeTrackingEnabled {get;set;}
											TrackColumnsUpdated = $_.TrackColumnsUpdatedEnabled # System.Boolean TrackColumnsUpdatedEnabled {get;set;}
										}
										FileTable = New-Object -TypeName PSObject -Property @{
											FileTableDirectoryName = $_.FileTableDirectoryName # System.String FileTableDirectoryName {get;set;}
											FileTableNameColumnCollation = $_.FileTableNameColumnCollation # System.String FileTableNameColumnCollation {get;set;}
											FileTableNamespaceEnabled = $_.FileTableNamespaceEnabled # System.Boolean FileTableNamespaceEnabled {get;set;}
										}
										Storage = New-Object -TypeName PSObject -Property @{
											Compression = New-Object -TypeName PSObject -Property @{
												PartitionsNotCompressed = if (
													$DbEngineType -ieq $StandaloneDbEngine -and
													$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
												) {
													$(
														$TablePhysicalPartitions.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)") | Where-Object { 
															-not $_.DataCompression -or $_.DataCompression -eq ($DataCompressionType::None).value__
														} | Sort-Object -Property PartitionNumber | ForEach-Object { $_.PartitionNumber }
													) -join ','
												} else {
													$null
												}
												PartitionsUsingRowCompression = if (
													$DbEngineType -ieq $StandaloneDbEngine -and
													$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
												) {
													$(
														$TablePhysicalPartitions.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)") | Where-Object { 
															$_.DataCompression -and $_.DataCompression -eq ($DataCompressionType::Row).value__
														} | Sort-Object -Property PartitionNumber | ForEach-Object { $_.PartitionNumber }
													) -join ','
												} else {
													$null
												}
												PartitionsUsingPageCompression = if (
													$DbEngineType -ieq $StandaloneDbEngine -and
													$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
												) {
													$(
														$TablePhysicalPartitions.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)") | Where-Object { 
															$_.DataCompression -and $_.DataCompression -eq ($DataCompressionType::Page).value__
														} | Sort-Object -Property PartitionNumber | ForEach-Object { $_.PartitionNumber }
													) -join ','
												} else {
													$null
												}
											}
											Filegroups = New-Object -TypeName PSObject -Property @{
												TextFileGroup = $_.TextFileGroup # System.String TextFileGroup {get;set;} 
												IsPartitioned = $_.IsPartitioned # System.Boolean IsPartitioned {get;} 
												FileGroup = $_.FileGroup # System.String FileGroup {get;set;}
												FileStreamFileGroup = $_.FileStreamFileGroup # System.String FileStreamFileGroup {get;set;}
											}
											General = New-Object -TypeName PSObject -Property @{
												IsVarDecimalStorageFormatEnabled = $_.IsVarDecimalStorageFormatEnabled # System.Boolean IsVarDecimalStorageFormatEnabled {get;set;}
												IndexSpaceUsedKB = $_.IndexSpaceUsed # System.Double IndexSpaceUsed {get;}
												RowCount = $_.RowCount # System.Int64 RowCount {get;}
												#RowCountAsDouble = $_.RowCountAsDouble	# System.Double RowCountAsDouble {get;}
												DataSpaceUsedKB = $_.DataSpaceUsed # System.Double DataSpaceUsed {get;}
											}
											Partitioning = New-Object -TypeName PSObject -Property @{
												PartitionScheme = $_.PartitionScheme # System.String PartitionScheme {get;set;}
												PhysicalPartitions = if (
													$DbEngineType -ieq $StandaloneDbEngine -and
													$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
												) {
													@() + (
														$TablePhysicalPartitions.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)") | ForEach-Object {
															New-Object -TypeName PSObject -Property @{
																DataCompression = [String](Get-DataCompressionTypeValue -DataCompressionType $_.DataCompression) # Microsoft.SqlServer.Management.Smo.DataCompressionType DataCompression {get;set;}
																FileGroupName = $_.FileGroupName # System.String FileGroupName {get;set;}
																#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlSmoObject Parent {get;}
																PartitionNumber = $_.PartitionNumber # System.Int32 PartitionNumber {get;set;}
																#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
																RangeType = [String](Get-RangeTypeValue -RangeType $_.RangeType) # Microsoft.SqlServer.Management.Smo.RangeType RangeType {get;set;}
																RightBoundaryValue = if ($_.RightBoundaryValue) { $_.RightBoundaryValue.ToString() } else { $null } # System.Object RightBoundaryValue {get;set;}
																RowCount = $_.RowCount # System.Double RowCount {get;}
																#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
																#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
																#UserData = $_.UserData	# System.Object UserData {get;set;}
															}
														}
													) # Microsoft.SqlServer.Management.Smo.PhysicalPartitionCollection PhysicalPartitions {get;}
												} else {
													$null
												}
												PartitionSchemeParameters = if (
													$DbEngineType -ieq $StandaloneDbEngine -and
													$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
												) {
													@() + (
														$TablePartitionSchemeParameters.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)") | ForEach-Object {
															New-Object -TypeName PSObject -Property @{
																ID = $_.ID # System.Int32 ID {get;}
																Name = $_.Name # System.String Name {get;set;}
																#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlSmoObject Parent {get;set;}
																#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
																#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
																#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
																#UserData = $_.UserData	# System.Object UserData {get;set;}													
															}
														}
													) # Microsoft.SqlServer.Management.Smo.PartitionSchemeParameterCollection PartitionSchemeParameters {get;}
												} else {
													$null
												}
												FileStreamPartitionScheme = $_.FileStreamPartitionScheme # System.String FileStreamPartitionScheme {get;set;}
											}
										}
									}

									Checks = @() + (Get-CheckInformation2 -CheckCollection $($TableChecks.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)")))

									Columns = @() + (
										Get-ColumnInformation2 `
										-ColumnCollection $($TableColumns.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)")) `
										-DefaultConstraintCollection $($TableDefaultConstraints.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)"))
									)


									# Events = $_.Events	# Microsoft.SqlServer.Management.Smo.TableEvents Events {get;}
									#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}

									ForeignKeys = @() + (
										$TableForeignKeys.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)") | ForEach-Object {
											New-Object -TypeName PSObject -Property @{
												Columns = @() + (
													$TableForeignKeyColumns.Select("TableID = $($_.TableID) and SchemaID = $($_.SchemaID) and ForeignKeyID = $($_.ID)") | ForEach-Object {
														$_.Name
													}

												) # Microsoft.SqlServer.Management.Smo.ForeignKeyColumnCollection Columns {get;}
												CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
												DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
												DeleteAction = [String](Get-ForeignKeyActionValue -ForeignKeyAction $_.DeleteAction) # Microsoft.SqlServer.Management.Smo.ForeignKeyAction DeleteAction {get;set;}
												#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
												ID = $_.ID # System.Int32 ID {get;}
												IsChecked = $_.IsChecked # System.Boolean IsChecked {get;set;}
												IsEnabled = $_.IsEnabled # System.Boolean IsEnabled {get;set;}
												IsFileTableDefined = $_.IsFileTableDefined # System.Boolean IsFileTableDefined {get;}
												IsSystemNamed = $_.IsSystemNamed # System.Boolean IsSystemNamed {get;}
												Name = $_.Name # System.String Name {get;set;}
												IsNotForReplication = $_.NotForReplication # System.Boolean NotForReplication {get;set;}
												EnforceForReplication = if ($_.NotForReplication) { $false } else { $true } # System.Boolean NotForReplication {get;set;}
												#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Table Parent {get;set;}
												#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
												ReferencedKey = $_.ReferencedKey # System.String ReferencedKey {get;}
												ReferencedTable = $_.ReferencedTable # System.String ReferencedTable {get;set;}
												ReferencedTableSchema = $_.ReferencedTableSchema # System.String ReferencedTableSchema {get;set;}
												#ScriptReferencedTable = $_.ScriptReferencedTable # System.String ScriptReferencedTable {get;set;}
												#ScriptReferencedTableSchema = $_.ScriptReferencedTableSchema # System.String ScriptReferencedTableSchema {get;set;}
												#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
												UpdateAction = [String](Get-ForeignKeyActionValue -ForeignKeyAction $_.UpdateAction) # Microsoft.SqlServer.Management.Smo.ForeignKeyAction UpdateAction {get;set;}
												#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
												#UserData = $_.UserData	# System.Object UserData {get;set;}	
											}
										}
									) # Microsoft.SqlServer.Management.Smo.ForeignKeyCollection ForeignKeys {get;}

									FullTextIndex = if ($DbEngineType -ieq $StandaloneDbEngine) {
										$TableFullTextIndexes.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)") | ForEach-Object {
											New-Object -TypeName PSObject -Property @{
												General = New-Object -TypeName PSObject -Property @{
													CatalogName = $_.CatalogName # System.String CatalogName {get;set;}
													UniqueIndexName = $_.UniqueIndexName # System.String UniqueIndexName {get;set;}
													PopulationStatus = [String](Get-IndexPopulationStatusValue -IndexPopulationStatus $_.PopulationStatus) # Microsoft.SqlServer.Management.Smo.IndexPopulationStatus PopulationStatus {get;}
													FilegroupName = $_.FilegroupName # System.String FilegroupName {get;set;}
													StopListName = $_.StopListName # System.String StopListName {get;set;}
													StopListOption = [String](Get-StopListOptionValue -StopListOption $_.StopListOption) # Microsoft.SqlServer.Management.Smo.StopListOption StopListOption {get;set;}
													SearchPropertyListName = $_.SearchPropertyListName # System.String SearchPropertyListName {get;set;}
													ItemCount = $_.ItemCount # System.Int32 ItemCount {get;}
													DocumentsProcessed = $_.DocumentsProcessed # System.Int32 DocumentsProcessed {get;}
													PendingChanges = $_.PendingChanges # System.Int32 PendingChanges {get;}
													NumberOfFailures = $_.NumberOfFailures # System.Int32 NumberOfFailures {get;}
													IsEnabled = $_.IsEnabled # System.Boolean IsEnabled {get;}
													ChangeTracking = [String](Get-ChangeTrackingValue -ChangeTracking $_.ChangeTracking) # Microsoft.SqlServer.Management.Smo.ChangeTracking ChangeTracking {get;set;}
												}
												Columns = @() + (
													$TableFullTextIndexColumns.Select("TableID = $($_.TableID) and SchemaID = $($_.SchemaID)") | ForEach-Object {
														New-Object -TypeName PSObject -Property @{
															Name = $_.Name # System.String Name {get;set;}
															LanguageForWordBreaker = $_.Language # System.String Language {get;set;}
															TypeColumnName = $_.TypeColumnName # System.String TypeColumnName {get;set;}
															StatisticalSemantics = if ($_.StatisticalSemantics -eq 0) { $false } else { $true } # System.Int32 StatisticalSemantics {get;set;}
															#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.FullTextIndex Parent {get;set;}
															#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
															#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
															#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
															#UserData = $_.UserData	# System.Object UserData {get;set;}
														}
													} 
												) # Microsoft.SqlServer.Management.Smo.FullTextIndexColumnCollection IndexedColumns {get;}

												#Parent = $_.FullTextIndex.Parent	# Microsoft.SqlServer.Management.Smo.TableViewBase Parent {get;set;}
												#Properties = $_.FullTextIndex.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
												#State = $_.FullTextIndex.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
												#Urn = $_.FullTextIndex.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
												#UserData = $_.FullTextIndex.UserData	# System.Object UserData {get;set;}

											}

										} # Microsoft.SqlServer.Management.Smo.FullTextIndex FullTextIndex {get;}
									} else {
										$null
									}



									Indexes = if (
										$DbEngineType -ieq $StandaloneDbEngine -and
										$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
									) {
										@() + (
											Get-IndexInformation2 `
											-IndexCollection $($TableIndexes.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)") ) `
											-IndexedColumnCollection $($TableIndexColumns.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)") ) `
											-PartitionSchemeParameterCollection $($TableIndexPartitionSchemeParameters.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)") ) `
											-PhysicalPartitionCollection $($TableIndexPhysicalPartitions.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)") )
										) # Microsoft.SqlServer.Management.Smo.IndexCollection Indexes {get;}
									} else {
										@() + (
											Get-IndexInformation2 `
											-IndexCollection $($TableIndexes.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)") ) `
											-IndexedColumnCollection $($TableIndexColumns.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)") ) `
											-PartitionSchemeParameterCollection $null `
											-PhysicalPartitionCollection $null
										) # Microsoft.SqlServer.Management.Smo.IndexCollection Indexes {get;}
									}


									# #Owner = $_.Owner	# System.String Owner {get;set;}
									# #Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
									# #Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
									# #State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}

									Statistics = @() + (
										$TableStatistics.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)") | ForEach-Object {
											$StatisticID = $_.ID

											New-Object -TypeName PSObject -Property @{
												General = New-Object -TypeName PSObject -Property @{
													FileGroup = $_.FileGroup # System.String FileGroup {get;set;}
													HasFilter = $_.HasFilter # System.Boolean HasFilter {get;}
													ID = $_.ID # System.Int32 ID {get;}
													IsAutoCreated = $_.IsAutoCreated # System.Boolean IsAutoCreated {get;}
													IsFromIndexCreation = $_.IsFromIndexCreation # System.Boolean IsFromIndexCreation {get;}
													IsTemporary = $_.IsTemporary # System.Boolean IsTemporary {get;}
													LastUpdated = $_.LastUpdated # System.DateTime LastUpdated {get;}
													Name = $_.Name # System.String Name {get;set;}
													NoAutomaticRecomputation = $_.NoAutomaticRecomputation # System.Boolean NoAutomaticRecomputation {get;set;}
													Columns = @() + (
														$TableStatisticsColumns.Select("TableID = $($_.TableID) and SchemaID = $($_.SchemaID) and StatisticID = $($_.ID)") | ForEach-Object {
															New-Object -TypeName PSObject -Property @{
																ID = $_.ID # System.Int32 ID {get;}
																Name = $_.Name # System.String Name {get;set;}
																#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Statistic Parent {get;set;}
																#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
																#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
																#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
																#UserData = $_.UserData	# System.Object UserData {get;set;}
															}
														}
													) # Microsoft.SqlServer.Management.Smo.StatisticColumnCollection StatisticColumns {get;}

												}
												FilterDefinition = $_.FilterDefinition # System.String FilterDefinition {get;set;}	

												#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.StatisticEvents Events {get;}
												#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlSmoObject Parent {get;set;}
												#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
												#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
												#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
												#UserData = $_.UserData	# System.Object UserData {get;set;}

											}
										}
									) # Microsoft.SqlServer.Management.Smo.StatisticCollection Statistics {get;}

									Triggers = @() + (Get-TriggerInformation2 -TriggerCollection $($TableTriggers.Select("TableID = $($_.ID) and SchemaID = $($_.SchemaID)")))

									# #Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
									# #UserData = $_.UserData	# System.Object UserData {get;set;}

									# Federation related properties (For Azure databases)...punting on these for now
									# DistributionName = $_.DistributionName # System.String DistributionName {get;set;}
									# FederationColumnID = $_.FederationColumnID	# System.Int32 FederationColumnID {get;}
									# FederationColumnName = $_.FederationColumnName	# System.String FederationColumnName {get;set;}							
								}
							}
						)
					} else {
						$null
					}
					#endregion

					#region
					Views = if (($IsAccessible -eq $true) -and ($IncludeObjectInformation -eq $true)) {
						@() + (
							$Views.Rows | ForEach-Object {
								New-Object -TypeName PSObject -Property @{
									Properties = New-Object -TypeName PSObject -Property @{
										General = New-Object -TypeName PSObject -Property @{
											Description = New-Object -TypeName PSObject -Property @{
												ID = $_.ID # System.Int32 ID {get;}
												CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
												DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
												Name = $_.Name # System.String Name {get;set;}
												Owner = $_.Owner # System.String Owner {get;set;}
												Schema = $_.Schema # System.String Schema {get;set;}
												IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}

												# Properties exposed by SMO but not in the SSMS GUI...these probably belong here if they were
												HasAfterTrigger = $_.HasAfterTrigger # System.Boolean HasAfterTrigger {get;}
												HasColumnSpecification = $_.HasColumnSpecification # System.Boolean HasColumnSpecification {get;}
												HasDeleteTrigger = $_.HasDeleteTrigger # System.Boolean HasDeleteTrigger {get;}
												HasIndex = $_.HasIndex # System.Boolean HasIndex {get;}
												HasInsertTrigger = $_.HasInsertTrigger # System.Boolean HasInsertTrigger {get;}
												HasInsteadOfTrigger = $_.HasInsteadOfTrigger # System.Boolean HasInsteadOfTrigger {get;}
												HasUpdateTrigger = $_.HasUpdateTrigger # System.Boolean HasUpdateTrigger {get;}
												IsIndexable = $_.IsIndexable # System.Boolean IsIndexable {get;}
												IsSchemaOwned = $_.IsSchemaOwned # System.Boolean IsSchemaOwned {get;}
											}
											Options = New-Object -TypeName PSObject -Property @{
												AnsiNullsStatus = $_.AnsiNullsStatus # System.Boolean AnsiNullsStatus {get;set;}
												IsEncrypted = $_.IsEncrypted # System.Boolean IsEncrypted {get;set;}
												QuotedIdentifierStatus = $_.QuotedIdentifierStatus # System.Boolean QuotedIdentifierStatus {get;set;}
												IsSchemaBound = $_.IsSchemaBound # System.Boolean IsSchemaBound {get;set;}
												ReturnsViewMetadata = $_.ReturnsViewMetadata # System.Boolean ReturnsViewMetadata {get;set;}
											}
										}
									}

									Columns = @() + (
										Get-ColumnInformation2 `
										-ColumnCollection $($ViewColumns.Select("ViewID = $($_.ID) and SchemaID = $($_.SchemaID)")) `
										-DefaultConstraintCollection $null
									)


									#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.ViewEvents Events {get;}
									#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}

									FullTextIndex = if (
										$DbEngineType -ieq $StandaloneDbEngine -and
										$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
									) {
										$ViewFullTextIndexes.Select("ViewID = $($_.ID) and SchemaID = $($_.SchemaID)") | ForEach-Object {
											New-Object -TypeName PSObject -Property @{
												General = New-Object -TypeName PSObject -Property @{
													CatalogName = $_.CatalogName # System.String CatalogName {get;set;}
													UniqueIndexName = $_.UniqueIndexName # System.String UniqueIndexName {get;set;}
													PopulationStatus = [String](Get-IndexPopulationStatusValue -IndexPopulationStatus $_.PopulationStatus) # Microsoft.SqlServer.Management.Smo.IndexPopulationStatus PopulationStatus {get;}
													FilegroupName = $_.FilegroupName # System.String FilegroupName {get;set;}
													StopListName = $_.StopListName # System.String StopListName {get;set;}
													StopListOption = [String](Get-StopListOptionValue -StopListOption $_.StopListOption) # Microsoft.SqlServer.Management.Smo.StopListOption StopListOption {get;set;}
													SearchPropertyListName = $_.SearchPropertyListName # System.String SearchPropertyListName {get;set;}
													ItemCount = $_.ItemCount # System.Int32 ItemCount {get;}
													DocumentsProcessed = $_.DocumentsProcessed # System.Int32 DocumentsProcessed {get;}
													PendingChanges = $_.PendingChanges # System.Int32 PendingChanges {get;}
													NumberOfFailures = $_.NumberOfFailures # System.Int32 NumberOfFailures {get;}
													IsEnabled = $_.IsEnabled # System.Boolean IsEnabled {get;}
													ChangeTracking = [String](Get-ChangeTrackingValue -ChangeTracking $_.ChangeTracking) # Microsoft.SqlServer.Management.Smo.ChangeTracking ChangeTracking {get;set;}
												}
												Columns = @() + (
													$ViewFullTextIndexColumns.Select("ViewID = $($_.ViewID) and SchemaID = $($_.SchemaID)") | ForEach-Object {
														New-Object -TypeName PSObject -Property @{
															Name = $_.Name # System.String Name {get;set;}
															LanguageForWordBreaker = $_.Language # System.String Language {get;set;}
															TypeColumnName = $_.TypeColumnName # System.String TypeColumnName {get;set;}
															StatisticalSemantics = if ($_.StatisticalSemantics -eq 0) { $false } else { $true } # System.Int32 StatisticalSemantics {get;set;}
															#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.FullTextIndex Parent {get;set;}
															#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
															#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
															#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
															#UserData = $_.UserData	# System.Object UserData {get;set;}
														}
													} 
												) # Microsoft.SqlServer.Management.Smo.FullTextIndexColumnCollection IndexedColumns {get;}
												#Parent = $_.FullTextIndex.Parent	# Microsoft.SqlServer.Management.Smo.TableViewBase Parent {get;set;}
												#Properties = $_.FullTextIndex.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
												#State = $_.FullTextIndex.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
												#Urn = $_.FullTextIndex.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
												#UserData = $_.FullTextIndex.UserData	# System.Object UserData {get;set;}

											} # Microsoft.SqlServer.Management.Smo.FullTextIndex FullTextIndex {get;}
										}
									} else {
										$null
									}

									Indexes = if (
										$DbEngineType -ieq $StandaloneDbEngine -and
										$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
									) {
										@() + (
											Get-IndexInformation2 `
											-IndexCollection $($ViewIndexes.Select("ViewID = $($_.ID) and SchemaID = $($_.SchemaID)") ) `
											-IndexedColumnCollection $($ViewIndexColumns.Select("ViewID = $($_.ID) and SchemaID = $($_.SchemaID)") ) `
											-PartitionSchemeParameterCollection $($ViewIndexPartitionSchemeParameters.Select("ViewID = $($_.ID) and SchemaID = $($_.SchemaID)") ) `
											-PhysicalPartitionCollection $($ViewIndexPhysicalPartitions.Select("ViewID = $($_.ID) and SchemaID = $($_.SchemaID)") )
										) # Microsoft.SqlServer.Management.Smo.IndexCollection Indexes {get;}
									} else {
										@() + (
											Get-IndexInformation2 `
											-IndexCollection $($ViewIndexes.Select("ViewID = $($_.ID) and SchemaID = $($_.SchemaID)") ) `
											-IndexedColumnCollection $($ViewIndexColumns.Select("ViewID = $($_.ID) and SchemaID = $($_.SchemaID)") ) `
											-PartitionSchemeParameterCollection $null `
											-PhysicalPartitionCollection $null
										) # Microsoft.SqlServer.Management.Smo.IndexCollection Indexes {get;}
									}


									#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
									#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
									#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}


									Statistics = @() + (
										$ViewStatistics.Select("ViewID = $($_.ID) and SchemaID = $($_.SchemaID)") | Where-Object { $_.ID } | ForEach-Object {
											New-Object -TypeName PSObject -Property @{
												General = New-Object -TypeName PSObject -Property @{
													FileGroup = $_.FileGroup # System.String FileGroup {get;set;}
													HasFilter = $_.HasFilter # System.Boolean HasFilter {get;}
													ID = $_.ID # System.Int32 ID {get;}
													IsAutoCreated = $_.IsAutoCreated # System.Boolean IsAutoCreated {get;}
													IsFromIndexCreation = $_.IsFromIndexCreation # System.Boolean IsFromIndexCreation {get;}
													IsTemporary = $_.IsTemporary # System.Boolean IsTemporary {get;}
													LastUpdated = $_.LastUpdated # System.DateTime LastUpdated {get;}
													Name = $_.Name # System.String Name {get;set;}
													NoAutomaticRecomputation = $_.NoAutomaticRecomputation # System.Boolean NoAutomaticRecomputation {get;set;}
													Columns = @() + (
														$ViewStatisticsColumns.Select("ViewID = $($_.ViewID) and SchemaID = $($_.SchemaID) and StatisticID = $($_.ID)") | ForEach-Object {
															New-Object -TypeName PSObject -Property @{
																ID = $_.ID # System.Int32 ID {get;}
																Name = $_.Name # System.String Name {get;set;}
															}
														}
													) # Microsoft.SqlServer.Management.Smo.StatisticColumnCollection StatisticColumns {get;}

												}
												FilterDefinition = $_.FilterDefinition # System.String FilterDefinition {get;set;}	

												#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.StatisticEvents Events {get;}
												#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlSmoObject Parent {get;set;}
												#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
												#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
												#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
												#UserData = $_.UserData	# System.Object UserData {get;set;}

											}
										}
									) # Microsoft.SqlServer.Management.Smo.StatisticCollection Statistics {get;}


									Definition = if ($_.IsSystemObject -eq $true) {
										# Don't include definitions for system objects
										$null
									} else {
										if (-not [String]::IsNullOrEmpty($_.Definition)) {
											$_.Definition.Trim()
										} else {
											$null
										}
									}
									#$_.Definition
									#TextBody = $_.TextBody # System.String TextBody {get;set;}
									#TextHeader = $_.TextHeader # System.String TextHeader {get;set;}
									#TextMode = $_.TextMode # System.Boolean TextMode {get;set;}

									Triggers = @() + (Get-TriggerInformation2 -TriggerCollection $($ViewTriggers.Select("ViewID = $($_.ID) and SchemaID = $($_.SchemaID)")))

									#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
									#UserData = $_.UserData	# System.Object UserData {get;set;}

								}
							}
						)
					} else {
						$null
					}
					#endregion


					#region
					# Synonyms supported starting with SQL 2005
					Synonyms = if (
						$IsAccessible -eq $true -and 
						$IncludeObjectInformation -eq $true -and
						$Server.Version.CompareTo($SQLServer2005) -ge 0
					) {
						@() + (
							$_.Synonyms | Where-Object { $_.ID } | ForEach-Object {
								New-Object -TypeName PSObject -Property @{
									BaseDatabase = $_.BaseDatabase # System.String BaseDatabase {get;set;}
									BaseObject = $_.BaseObject # System.String BaseObject {get;set;}
									BaseSchema = $_.BaseSchema # System.String BaseSchema {get;set;}
									BaseServer = $_.BaseServer # System.String BaseServer {get;set;}
									BaseType = [String](Get-SynonymBaseTypeValue -SynonymBaseType $_.BaseType) # Microsoft.SqlServer.Management.Smo.SynonymBaseType BaseType {get;}
									CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
									DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
									#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.SynonymEvents Events {get;}
									#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
									ID = $_.ID # System.Int32 ID {get;}
									IsSchemaOwned = $_.IsSchemaOwned # System.Boolean IsSchemaOwned {get;}
									Name = $_.Name # System.String Name {get;set;}
									Owner = $_.Owner # System.String Owner {get;set;}
									#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
									#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
									Schema = $_.Schema # System.String Schema {get;set;}
									#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
									#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
									#UserData = $_.UserData	# System.Object UserData {get;set;}
								}
							}
						)
					} else {
						$null
					} # Microsoft.SqlServer.Management.Smo.SynonymCollection Synonyms {get;}
					#endregion


					#region
					Programmability = New-Object -TypeName PSObject -Property @{

						#region
						StoredProcedures = if (
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true
						) {
							@() + (
								$StoredProcedures.Rows | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										Properties = New-Object -TypeName PSObject -Property @{
											General = New-Object -TypeName PSObject -Property @{
												Description = New-Object -TypeName PSObject -Property @{
													ID = $_.ID # System.Int32 ID {get;}
													Name = $_.Name # System.String Name {get;set;}
													Schema = $_.Schema # System.String Schema {get;set;}
													Owner = $_.Owner # System.String Owner {get;set;}
													AssemblyName = $_.AssemblyName # System.String AssemblyName {get;set;}
													ClassName = $_.ClassName # System.String ClassName {get;set;}
													MethodName = $_.MethodName # System.String MethodName {get;set;}
													ImplementationType = [String](Get-ImplementationTypeValue -ImplementationType $_.ImplementationType) # Microsoft.SqlServer.Management.Smo.ImplementationType ImplementationType {get;set;}
													CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
													DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
													ExecutionContext = [String](Get-ExecutionContextValue -ExecutionContext $_.ExecutionContext) # Microsoft.SqlServer.Management.Smo.ExecutionContext ExecutionContext {get;set;}
													ExecutionContextPrincipal = $_.ExecutionContextPrincipal # System.String ExecutionContextPrincipal {get;set;}
													IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}													
												}
												Options = New-Object -TypeName PSObject -Property @{
													AnsiNullsStatus = $_.AnsiNullsStatus # System.Boolean AnsiNullsStatus {get;set;}
													ForReplication = $_.ForReplication # System.Boolean ForReplication {get;set;}
													IsEncrypted = $_.IsEncrypted # System.Boolean IsEncrypted {get;set;}
													IsSchemaOwned = $_.IsSchemaOwned # System.Boolean IsSchemaOwned {get;}
													QuotedIdentifierStatus = $_.QuotedIdentifierStatus # System.Boolean QuotedIdentifierStatus {get;set;}
													Recompile = $_.Recompile # System.Boolean Recompile {get;set;}
													Startup = $_.Startup # System.Boolean Startup {get;set;}
												}
											}
										}
										#Permissions = $null
										#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}

										#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.StoredProcedureEvents Events {get;}

										# Do I even care about this??
										#NumberedStoredProcedures = $_.NumberedStoredProcedures # Microsoft.SqlServer.Management.Smo.NumberedStoredProcedureCollection NumberedStoredProcedures {get;}

										# This section adds minutes to runtime. Can we speed it up?
										Parameters = @() + (
											$StoredProcedureParameters.Select("StoredProcedureID = $($_.ID) and SchemaID = $($_.SchemaID)") | ForEach-Object {
												New-Object -TypeName PSObject -Property @{
													DataType = New-Object -TypeName PSObject -Property @{
														MaximumLength = $_.Length # System.Int32 MaximumLength {get;set;}
														Name = $_.DataType
														NumericPrecision = $_.NumericPrecision # System.Int32 NumericPrecision {get;set;}
														NumericScale = $_.NumericScale # System.Int32 NumericScale {get;set;}
														Schema = $_.DataTypeSchema # System.String  DataTypeSchema {get;set;}
														SqlDataType = $_.SystemType
														XmlDocumentConstraint = [String](Get-XmlDocumentConstraintValue -XmlDocumentConstraint $_.XmlDocumentConstraint) # Microsoft.SqlServer.Management.Smo.XmlDocumentConstraint XmlDocumentConstraint {get;set;}
													}
													DefaultValue = $_.DefaultValue # System.String DefaultValue {get;set;}
													#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
													ID = $_.ID # System.Int32 ID {get;}
													IsCursorParameter = $_.IsCursorParameter # System.Boolean IsCursorParameter {get;set;}
													IsOutputParameter = $_.IsOutputParameter # System.Boolean IsOutputParameter {get;set;}
													IsReadOnly = $_.IsReadOnly # System.Boolean IsReadOnly {get;set;}
													Name = $_.Name # System.String Name {get;set;}
													#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.StoredProcedure Parent {get;set;}
													#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
													#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
													#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
													#UserData = $_.UserData	# System.Object UserData {get;set;}												
												}
											}
										) # Microsoft.SqlServer.Management.Smo.StoredProcedureParameterCollection Parameters {get;}

										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#TextBody = $_.TextBody # System.String TextBody {get;set;}
										#TextHeader = $_.TextHeader # System.String TextHeader {get;set;}
										#TextMode = $_.TextMode # System.Boolean TextMode {get;set;}

										Definition = if ($_.IsSystemObject -eq $true) {
											# Don't include definitions for system objects
											$null
										} else {
											if (-not [String]::IsNullOrEmpty($_.Definition)) {
												$_.Definition.Trim()
											} else {
												$null
											}
										}

										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.StoredProcedureCollection StoredProcedures {get;}
						#endregion

						#region
						# Doesn't work against SQL 2000 with SMO 2012
						# Extended Stored Procedures not available in Windows Azure Databases
						ExtendedStoredProcedures = if (
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine -and 
							(
								$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0 -or
								(
									$($Server.Information.Version).CompareTo($SQLServer2000) -ge 0 -and
									$SmoMajorVersion -le 10
								)
							)
						) {
							@() + (
								$_.ExtendedStoredProcedures | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
										DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
										DllLocation = $_.DllLocation # System.String DllLocation {get;set;}
										#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										ID = $_.ID # System.Int32 ID {get;}
										IsSchemaOwned = $_.IsSchemaOwned # System.Boolean IsSchemaOwned {get;}
										IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
										Name = $_.Name # System.String Name {get;set;}
										Owner = $_.Owner # System.String Owner {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										Schema = $_.Schema # System.String Schema {get;set;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}

									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.ExtendedStoredProcedureCollection ExtendedStoredProcedures {get;}
						#endregion

						#region
						Functions = if (($IsAccessible -eq $true) -and ($IncludeObjectInformation -eq $true)) {
							@() + (
								$UserDefinedFunctions | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										Properties = New-Object -TypeName PSObject -Property @{
											General = New-Object -TypeName PSObject -Property @{
												Description = New-Object -TypeName PSObject -Property @{
													ID = $_.ID # System.Int32 ID {get;}
													Name = $_.Name # System.String Name {get;set;}
													Schema = $_.Schema # System.String Schema {get;set;}
													Owner = $_.Owner # System.String Owner {get;set;}
													AssemblyName = $_.AssemblyName # System.String AssemblyName {get;set;}
													ClassName = $_.ClassName # System.String ClassName {get;set;}
													MethodName = $_.MethodName # System.String MethodName {get;set;}
													ImplementationType = [String](Get-ImplementationTypeValue -ImplementationType $_.ImplementationType) # Microsoft.SqlServer.Management.Smo.ImplementationType ImplementationType {get;set;}
													CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
													DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
													ExecutionContext = [String](Get-ExecutionContextValue -ExecutionContext $_.ExecutionContext) # Microsoft.SqlServer.Management.Smo.ExecutionContext ExecutionContext {get;set;}
													ExecutionContextPrincipal = $_.ExecutionContextPrincipal # System.String ExecutionContextPrincipal {get;set;}
													IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}

													# Not part of the SSMS GUI but probably goes best here
													DataType = New-Object -TypeName PSObject -Property @{
														MaximumLength = $_.Length # System.Int32 MaximumLength {get;set;}
														Name = $_.DataType
														NumericPrecision = $_.NumericPrecision # System.Int32 NumericPrecision {get;set;}
														NumericScale = $_.NumericScale # System.Int32 NumericScale {get;set;}
														Schema = $_.DataTypeSchema # System.String  DataTypeSchema {get;set;}
														SqlDataType = $_.SystemType
														XmlDocumentConstraint = [String](Get-XmlDocumentConstraintValue -XmlDocumentConstraint $_.XmlDocumentConstraint) # Microsoft.SqlServer.Management.Smo.XmlDocumentConstraint XmlDocumentConstraint {get;set;}
													}
													TableVariableName = $_.TableVariableName # System.String TableVariableName {get;set;}

												}
												Options = New-Object -TypeName PSObject -Property @{
													AnsiNullsStatus = $_.AnsiNullsStatus # System.Boolean AnsiNullsStatus {get;set;}
													FunctionType = [String](Get-UserDefinedFunctionTypeValue -UserDefinedFunctionType $_.FunctionType) # Microsoft.SqlServer.Management.Smo.UserDefinedFunctionType FunctionType {get;set;}
													IsDeterministic = $_.IsDeterministic # System.Boolean IsDeterministic {get;}
													IsEncrypted = $_.IsEncrypted # System.Boolean IsEncrypted {get;set;}
													IsSchemaBound = $_.IsSchemaBound # System.Boolean IsSchemaBound {get;set;}
													IsSchemaOwned = $_.IsSchemaOwned # System.Boolean IsSchemaOwned {get;}
													QuotedIdentifierStatus = $_.QuotedIdentifierStatus # System.Boolean QuotedIdentifierStatus {get;set;}
													ReturnsNullOnNullInput = $_.ReturnsNullOnNullInput # System.Boolean ReturnsNullOnNullInput {get;set;}
												}
											}
											#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										}

										Checks = @() + (Get-CheckInformation2 -CheckCollection $($UserDefinedFunctionChecks.Select("FunctionID = $($_.ID) and SchemaID = $($_.SchemaID)")))

										Columns = @() + (
											Get-ColumnInformation2 `
											-ColumnCollection $($UserDefinedFunctionColumns.Select("FunctionID = $($_.ID) and SchemaID = $($_.SchemaID)")) `
											-DefaultConstraintCollection $($UserDefinedFunctionDefaultConstraints.Select("FunctionID = $($_.ID) and SchemaID = $($_.SchemaID)"))
										)

										Indexes = @() + (
											Get-IndexInformation2 `
											-IndexCollection $($UserDefinedFunctionIndexes.Select("FunctionID = $($_.ID) and SchemaID = $($_.SchemaID)") ) `
											-IndexedColumnCollection $($UserDefinedFunctionIndexColumns.Select("FunctionID = $($_.ID) and SchemaID = $($_.SchemaID)") ) `
											-PartitionSchemeParameterCollection $null `
											-PhysicalPartitionCollection $null
										) # Microsoft.SqlServer.Management.Smo.IndexCollection Indexes {get;}


										# TODO: Come back to this later. I don't have a CLR handy to do this right now
										<#
 										OrderColumns = @() + (
 											$_.OrderColumns | Where-Object { $_.ID } | ForEach-Object {
 												New-Object -TypeName PSObject -Property @{
 													Descending = $_.Descending # System.Boolean Descending {get;set;}
 													ID = $_.ID # System.Int32 ID {get;}
 													Name = $_.Name # System.String Name {get;set;}
 													#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.UserDefinedFunction Parent {get;set;}
 													#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
 													#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
 													#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
 													#UserData = $_.UserData	# System.Object UserData {get;set;}													
 												}
 											}
 										) # Microsoft.SqlServer.Management.Smo.OrderColumnCollection OrderColumns {get;}
										#>


										Parameters = @() + (
											$UserDefinedFunctionParameters.Select("FunctionID = $($_.ID) and SchemaID = $($_.SchemaID)") | ForEach-Object {
												New-Object -TypeName PSObject -Property @{
													DataType = if ($_.DataType) { 
														New-Object -TypeName PSObject -Property @{
															MaximumLength = $_.DataType.MaximumLength # System.Int32 MaximumLength {get;set;}
															Name = $_.DataType.Name # System.String Name {get;set;}
															NumericPrecision = $_.DataType.NumericPrecision # System.Int32 NumericPrecision {get;set;}
															NumericScale = $_.DataType.NumericScale # System.Int32 NumericScale {get;set;}
															Schema = $_.DataType.Schema # System.String Schema {get;set;}
															SqlDataType = [String](Get-SqlDataTypeValue -SqlDataType $_.DataType.SqlDataType) # Microsoft.SqlServer.Management.Smo.SqlDataType SqlDataType {get;set;}
															XmlDocumentConstraint = [String](Get-XmlDocumentConstraintValue -XmlDocumentConstraint $_.DataType.XmlDocumentConstraint) # Microsoft.SqlServer.Management.Smo.XmlDocumentConstraint XmlDocumentConstraint {get;set;}
														} # Microsoft.SqlServer.Management.Smo.DataType DataType {get;set;}
													} else {
														New-Object -TypeName PSObject -Property @{
															MaximumLength = $null
															Name = $null
															NumericPrecision = $null
															NumericScale = $null
															Schema = $null
															SqlDataType = $null
															XmlDocumentConstraint = $null
														}
													}
													DefaultValue = $_.DefaultValue # System.String DefaultValue {get;set;}
													#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
													ID = $_.ID # System.Int32 ID {get;}
													IsReadOnly = $_.IsReadOnly # System.Boolean IsReadOnly {get;set;}
													Name = $_.Name # System.String Name {get;set;}
													#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.UserDefinedFunction Parent {get;set;}
													#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
													#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
													#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
													#UserData = $_.UserData	# System.Object UserData {get;set;}
												}
											}
										) # Microsoft.SqlServer.Management.Smo.UserDefinedFunctionParameterCollection Parameters {get;}
										#TextBody = $_.TextBody # System.String TextBody {get;set;}
										#TextHeader = $_.TextHeader # System.String TextHeader {get;set;}
										#TextMode = $_.TextMode # System.Boolean TextMode {get;set;}

										Definition = if ($_.IsSystemObject -eq $true) {
											# Don't include definitions for system objects
											$null
										} else {
											if (-not [String]::IsNullOrEmpty($_.Definition)) {
												$_.Definition.Trim()
											} else {
												$null
											}
										}

										#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.UserDefinedFunctionEvents Events {get;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}

									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.UserDefinedFunctionCollection UserDefinedFunctions {get;}
						#endregion

						#region
						DatabaseTriggers = if (($IsAccessible -eq $true) -and ($IncludeObjectInformation -eq $true)) {
							@() + (
								$_.Triggers | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										AnsiNullsStatus = $_.AnsiNullsStatus # System.Boolean AnsiNullsStatus {get;set;}
										AssemblyName = $_.AssemblyName # System.String AssemblyName {get;set;}
										BodyStartIndex = $_.BodyStartIndex # System.Int32 BodyStartIndex {get;}
										ClassName = $_.ClassName # System.String ClassName {get;set;}
										CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
										DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
										DdlTriggerEvents = if ($_.DdlTriggerEvents) { $_.DdlTriggerEvents.ToString() } else { $null } # Microsoft.SqlServer.Management.Smo.DatabaseDdlTriggerEventSet DdlTriggerEvents {get;set;}
										ExecutionContext = [String](Get-DatabaseDdlTriggerExecutionContextValue -DatabaseDdlTriggerExecutionContext $_.ExecutionContext) # Microsoft.SqlServer.Management.Smo.DatabaseDdlTriggerExecutionContext ExecutionContext {get;set;}
										ExecutionContextUser = $_.ExecutionContextUser # System.String ExecutionContextUser {get;set;}
										#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										ID = $_.ID # System.Int32 ID {get;}
										ImplementationType = [String](Get-ImplementationTypeValue -ImplementationType $_.ImplementationType) # Microsoft.SqlServer.Management.Smo.ImplementationType ImplementationType {get;set;}
										IsEnabled = $_.IsEnabled # System.Boolean IsEnabled {get;set;}
										IsEncrypted = $_.IsEncrypted # System.Boolean IsEncrypted {get;set;}
										IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
										MethodName = $_.MethodName # System.String MethodName {get;set;}
										Name = $_.Name # System.String Name {get;set;}
										NotForReplication = $_.NotForReplication # System.Boolean NotForReplication {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										QuotedIdentifierStatus = $_.QuotedIdentifierStatus # System.Boolean QuotedIdentifierStatus {get;set;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}

										Definition = if ($_.IsSystemObject -eq $true) {
											# Don't include definitions for system objects
											$null
										} else {
											if (-not [String]::IsNullOrEmpty($_.Text)) {
												$_.Text.Trim()
											} else {
												$null
											}
										}

										#Text = $_.Text # System.String Text {get;}
										#TextBody = $_.TextBody # System.String TextBody {get;set;}
										#TextHeader = $_.TextHeader # System.String TextHeader {get;set;}
										#TextMode = $_.TextMode # System.Boolean TextMode {get;set;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.DatabaseDdlTriggerCollection Triggers {get;}
						#endregion

						#region
						# Assemblies not available in SQL 2000
						# Assemblies not available in Windows Azure Databases
						Assemblies = if ( 
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine -and 
							$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
						) {
							@() + (
								$_.Assemblies | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										Properties = New-Object -TypeName PSObject -Property @{
											General = New-Object -TypeName PSObject -Property @{
												ID = $_.ID # System.Int32 ID {get;}
												Name = $_.Name # System.String Name {get;set;}
												Owner = $_.Owner # System.String Owner {get;set;}
												AssemblySecurityLevel = [String](Get-AssemblySecurityLevelValue -AssemblySecurityLevel $_.AssemblySecurityLevel) # Microsoft.SqlServer.Management.Smo.AssemblySecurityLevel AssemblySecurityLevel {get;set;}											
												CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
												Culture = $_.Culture # System.String Culture {get;}
												IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
												IsVisible = $_.IsVisible # System.Boolean IsVisible {get;set;}
												PublicKey = if ($_.PublicKey) { [System.BitConverter]::ToString($_.PublicKey) } else { $null } # System.Byte[] PublicKey {get;set;}
												SqlAssemblyFiles = @() + (
													$_.SqlAssemblyFiles | ForEach-Object {
														New-Object -TypeName PSObject -Property @{
															ID = $_.ID # System.Int32 ID {get;}
															Name = $_.Name # System.String Name {get;set;}
															#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.SqlAssembly Parent {get;set;}
															#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
															#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
															#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
															#UserData = $_.UserData	# System.Object UserData {get;set;}
														}
													}
												) # Microsoft.SqlServer.Management.Smo.SqlAssemblyFileCollection SqlAssemblyFiles {get;}
												Version = if ($_.Version) { $_.Version.ToString() } else { $null } # System.Version Version {get;}
											}
											#ExtendedProperties = $_.ExtendedProperties # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										}
										#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.SqlAssemblyEvents Events {get;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.SqlAssemblyCollection Assemblies {get;}
						#endregion

						#region
						Types = New-Object -TypeName PSObject -Property @{

							#region
							# UserDefinedAggregates not available in SQL 2000
							# UserDefinedAggregates not available in Windows Azure Databases
							UserDefinedAggregates = if (
								$IsAccessible -eq $true -and 
								$IncludeObjectInformation -eq $true -and
								$DbEngineType -ieq $StandaloneDbEngine -and 
								$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
							) {
								@() + (
									$_.UserDefinedAggregates | Where-Object { $_.ID } | ForEach-Object {
										New-Object -TypeName PSObject -Property @{
											AssemblyName = $_.AssemblyName # System.String AssemblyName {get;set;}
											ClassName = $_.ClassName # System.String ClassName {get;set;}
											CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
											DataType = if ($_.DataType) { 
												New-Object -TypeName PSObject -Property @{
													MaximumLength = $_.DataType.MaximumLength # System.Int32 MaximumLength {get;set;}
													Name = $_.DataType.Name # System.String Name {get;set;}
													NumericPrecision = $_.DataType.NumericPrecision # System.Int32 NumericPrecision {get;set;}
													NumericScale = $_.DataType.NumericScale # System.Int32 NumericScale {get;set;}
													Schema = $_.DataType.Schema # System.String Schema {get;set;}
													SqlDataType = [String](Get-SqlDataTypeValue -SqlDataType $_.DataType.SqlDataType) # Microsoft.SqlServer.Management.Smo.SqlDataType SqlDataType {get;set;}
													XmlDocumentConstraint = [String](Get-XmlDocumentConstraintValue -XmlDocumentConstraint $_.DataType.XmlDocumentConstraint) # Microsoft.SqlServer.Management.Smo.XmlDocumentConstraint XmlDocumentConstraint {get;set;}
												} # Microsoft.SqlServer.Management.Smo.DataType DataType {get;set;}
											} else {
												New-Object -TypeName PSObject -Property @{
													MaximumLength = $null
													Name = $null
													NumericPrecision = $null
													NumericScale = $null
													Schema = $null
													SqlDataType = $null
													XmlDocumentConstraint = $null
												}
											}
											DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
											#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
											ID = $_.ID # System.Int32 ID {get;}
											IsSchemaOwned = $_.IsSchemaOwned # System.Boolean IsSchemaOwned {get;}
											Name = $_.Name # System.String Name {get;set;}
											Owner = $_.Owner # System.String Owner {get;set;}
											Parameters = @() + (
												$_.Parameters | Where-Object { $_.ID } | ForEach-Object {
													New-Object -TypeName PSObject -Property @{
														DataType = if ($_.DataType) { 
															New-Object -TypeName PSObject -Property @{
																MaximumLength = $_.DataType.MaximumLength # System.Int32 MaximumLength {get;set;}
																Name = $_.DataType.Name # System.String Name {get;set;}
																NumericPrecision = $_.DataType.NumericPrecision # System.Int32 NumericPrecision {get;set;}
																NumericScale = $_.DataType.NumericScale # System.Int32 NumericScale {get;set;}
																Schema = $_.DataType.Schema # System.String Schema {get;set;}
																SqlDataType = [String](Get-SqlDataTypeValue -SqlDataType $_.DataType.SqlDataType) # Microsoft.SqlServer.Management.Smo.SqlDataType SqlDataType {get;set;}
																XmlDocumentConstraint = [String](Get-XmlDocumentConstraintValue -XmlDocumentConstraint $_.DataType.XmlDocumentConstraint) # Microsoft.SqlServer.Management.Smo.XmlDocumentConstraint XmlDocumentConstraint {get;set;}
															} # Microsoft.SqlServer.Management.Smo.DataType DataType {get;set;}
														} else {
															New-Object -TypeName PSObject -Property @{
																MaximumLength = $null
																Name = $null
																NumericPrecision = $null
																NumericScale = $null
																Schema = $null
																SqlDataType = $null
																XmlDocumentConstraint = $null
															}
														}
														#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
														ID = $_.ID # System.Int32 ID {get;}
														Name = $_.Name # System.String Name {get;set;}
														#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.UserDefinedAggregate Parent {get;set;}
														#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
														#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
														#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
														#UserData = $_.UserData	# System.Object UserData {get;set;}
													}
												}
											) # Microsoft.SqlServer.Management.Smo.UserDefinedAggregateParameterCollection Parameters {get;}
											#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
											#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
											Schema = $_.Schema # System.String Schema {get;set;}
											#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
											#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
											#UserData = $_.UserData	# System.Object UserData {get;set;}
										}
									}
								)
							} else {
								$null
							} # Microsoft.SqlServer.Management.Smo.UserDefinedAggregateCollection UserDefinedAggregates {get;}
							#endregion

							#region
							UserDefinedDataTypes = if (($IsAccessible -eq $true) -and ($IncludeObjectInformation -eq $true)) {
								@() + (
									$_.UserDefinedDataTypes | Where-Object { $_.ID } | ForEach-Object {
										New-Object -TypeName PSObject -Property @{
											Properties = New-Object -TypeName PSObject -Property @{
												General = New-Object -TypeName PSObject -Property @{
													ID = $_.ID # System.Int32 ID {get;}
													Schema = $_.Schema # System.String Schema {get;set;}
													Name = $_.Name # System.String Name {get;set;}
													SystemType = $_.SystemType # System.String SystemType {get;set;}
													Length = $_.Length # System.Int32 Length {get;set;}
													MaxLength = $_.MaxLength # System.Int16 MaxLength {get;}
													NumericPrecision = $_.NumericPrecision # System.Int32 NumericPrecision {get;set;}
													NumericScale = $_.NumericScale # System.Int32 NumericScale {get;set;}
													Nullable = $_.Nullable # System.Boolean Nullable {get;set;}

													Rule = $_.Rule # System.String Rule {get;set;}
													RuleSchema = $_.RuleSchema # System.String RuleSchema {get;set;}
													Default = $_.Default # System.String Default {get;set;}
													DefaultSchema = $_.DefaultSchema # System.String DefaultSchema {get;set;}

													# Facets not part of the properties dialog
													Collation = $_.Collation # System.String Collation {get;}
													AllowIdentity = $_.AllowIdentity # System.Boolean AllowIdentity {get;}
													Owner = $_.Owner # System.String Owner {get;set;}
													IsSchemaOwned = $_.IsSchemaOwned # System.Boolean IsSchemaOwned {get;}
													VariableLength = $_.VariableLength # System.Boolean VariableLength {get;}
												}
											}
											#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}

											#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
											#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
											#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
											#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
											#UserData = $_.UserData	# System.Object UserData {get;set;}
										}
									}
								)
							} else {
								$null
							} # Microsoft.SqlServer.Management.Smo.UserDefinedDataTypeCollection UserDefinedDataTypes {get;}
							#endregion

							#region
							UserDefinedTableTypes = if (($IsAccessible -eq $true) -and ($IncludeObjectInformation -eq $true)) {
								@() + (
									$_.UserDefinedTableTypes | Where-Object { $_.ID } | ForEach-Object {
										New-Object -TypeName PSObject -Property @{
											Properties = New-Object -TypeName PSObject -Property @{
												General = New-Object -TypeName PSObject -Property @{
													Description = New-Object -TypeName PSObject -Property @{
														Schema = $_.Schema # System.String Schema {get;set;}
														Name = $_.Name # System.String Name {get;set;}
														ID = $_.ID # System.Int32 ID {get;}
														CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
														DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}

														# Properties exposed by SMO but not in the SSMS GUI...these probably belong here if they were
														Collation = $_.Collation # System.String Collation {get;}														
														Owner = $_.Owner # System.String Owner {get;set;}
														IsSchemaOwned = $_.IsSchemaOwned # System.Boolean IsSchemaOwned {get;}
														IsUserDefined = $_.IsUserDefined # System.Boolean IsUserDefined {get;set;}
													}
													# SSMS doesn't have an Options section in the dialog but to keep it consistent
													# with the tables dialog I'm putting it in here
													Options = New-Object -TypeName PSObject -Property @{
														MaxLength = $_.MaxLength # System.Int16 MaxLength {get;}
														Nullable = $_.Nullable # System.Boolean Nullable {get;set;}
													}
												}
												#Permissions
												#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}												
											} 
											Checks = @() + (Get-CheckInformation -CheckCollection $_.Checks) # Microsoft.SqlServer.Management.Smo.CheckCollection Checks {get;}
											Columns = @() + (Get-ColumnInformation -ColumnCollection $_.Columns) # Microsoft.SqlServer.Management.Smo.ColumnCollection Columns {get;}
											Indexes = @() + (Get-IndexInformation -IndexCollection $_.Indexes) # Microsoft.SqlServer.Management.Smo.IndexCollection Indexes {get;}
											#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
											#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
											#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
											#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
											#UserData = $_.UserData	# System.Object UserData {get;set;}
										}
									}
								)
							} else {
								$null
							} # Microsoft.SqlServer.Management.Smo.UserDefinedTableTypeCollection UserDefinedTableTypes {get;}
							#endregion

							#region
							# No GUI for this in SSMS?
							# UserDefinedTypes not available in Windows Azure Databases
							UserDefinedTypes = if (
								$IsAccessible -eq $true -and 
								$IncludeObjectInformation -eq $true -and
								$DbEngineType -ieq $StandaloneDbEngine -and 
								$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
							) {
								@() + (
									$_.UserDefinedTypes | Where-Object { $_.ID } | ForEach-Object {
										New-Object -TypeName PSObject -Property @{
											AssemblyName = $_.AssemblyName # System.String AssemblyName {get;set;}
											BinaryTypeIdentifier = if ($_.BinaryTypeIdentifier) { [System.BitConverter]::ToString($_.BinaryTypeIdentifier) } else { $null } # System.Byte[] BinaryTypeIdentifier {get;}
											ClassName = $_.ClassName # System.String ClassName {get;set;}
											Collation = $_.Collation # System.String Collation {get;}
											#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.UserDefinedTypeEvents Events {get;}
											#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
											ID = $_.ID # System.Int32 ID {get;}
											IsBinaryOrdered = $_.IsBinaryOrdered # System.Boolean IsBinaryOrdered {get;}
											IsComVisible = $_.IsComVisible # System.Boolean IsComVisible {get;}
											IsFixedLength = $_.IsFixedLength # System.Boolean IsFixedLength {get;}
											IsNullable = $_.IsNullable # System.Boolean IsNullable {get;}
											IsSchemaOwned = $_.IsSchemaOwned # System.Boolean IsSchemaOwned {get;}
											MaxLength = $_.MaxLength # System.Int32 MaxLength {get;}
											Name = $_.Name # System.String Name {get;set;}
											NumericPrecision = $_.NumericPrecision # System.Int32 NumericPrecision {get;}
											NumericScale = $_.NumericScale # System.Int32 NumericScale {get;}
											Owner = $_.Owner # System.String Owner {get;set;}
											#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
											#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
											Schema = $_.Schema # System.String Schema {get;set;}
											#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
											#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
											#UserData = $_.UserData	# System.Object UserData {get;set;}
											UserDefinedTypeFormat = [String](Get-UserDefinedTypeFormatValue -UserDefinedTypeFormat $_.UserDefinedTypeFormat) # Microsoft.SqlServer.Management.Smo.UserDefinedTypeFormat UserDefinedTypeFormat {get;}
										}
									}
								)
							} else {
								$null
							} # Microsoft.SqlServer.Management.Smo.UserDefinedTypeCollection UserDefinedTypes {get;}
							#endregion

							#region
							# No GUI for this in SSMS?
							# XmlSchemaCollections not available in SQL 2000
							# XmlSchemaCollections not available in Windows Azure Databases
							XmlSchemaCollections = if (
								$IsAccessible -eq $true -and 
								$IncludeObjectInformation -eq $true -and
								$DbEngineType -ieq $StandaloneDbEngine -and 
								$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
							) {
								@() + (
									$_.XmlSchemaCollections | Where-Object { $_.ID } | ForEach-Object {
										New-Object -TypeName PSObject -Property @{
											CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
											DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
											#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
											ID = $_.ID # System.Int32 ID {get;}
											Name = $_.Name # System.String Name {get;set;}
											#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
											#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
											Schema = $_.Schema # System.String Schema {get;set;}
											#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}

											Definition = if (-not [String]::IsNullOrEmpty($_.Text)) {
												$_.Text.Trim()
											} else {
												$null
											}

											#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
											#UserData = $_.UserData	# System.Object UserData {get;set;}
										}
									}
								)
							} else {
								$null
							} # Microsoft.SqlServer.Management.Smo.XmlSchemaCollectionCollection XmlSchemaCollections {get;}
							#endregion
						}
						#endregion

						#region
						# No GUI for this in SSMS?
						# Rules not available in Windows Azure Databases
						Rules = if (
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine
						) {
							@() + (
								$_.Rules | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
										DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
										#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										ID = $_.ID # System.Int32 ID {get;}
										Name = $_.Name # System.String Name {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										Schema = $_.Schema # System.String Schema {get;set;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}

										Definition = [String]::Concat($_.TextHeader, $_.TextBody).Trim()

										#TextBody = $_.TextBody # System.String TextBody {get;set;}
										#TextHeader = $_.TextHeader # System.String TextHeader {get;set;}
										#TextMode = $_.TextMode # System.Boolean TextMode {get;set;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.RuleCollection Rules {get;}
						#endregion

						#region
						# No GUI for this in SSMS?
						# Defaults not available in Windows Azure Databases
						Defaults = if (
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine 
						) {
							@() + (
								$_.Defaults | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
										#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										ID = $_.ID # System.Int32 ID {get;}
										Name = $_.Name # System.String Name {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										Schema = $_.Schema # System.String Schema {get;set;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}

										Definition = [String]::Concat($_.TextHeader, $_.TextBody).Trim()

										#TextBody = $_.TextBody # System.String TextBody {get;set;}
										#TextHeader = $_.TextHeader # System.String TextHeader {get;set;}
										#TextMode = $_.TextMode # System.Boolean TextMode {get;set;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.DefaultCollection Defaults {get;}
						#endregion

						#region
						# There's a GUI for this in SSMS but it's a single page - very straightforward
						# PlanGuides not available in SQL 2000
						# PlanGuides not available in Windows Azure Databases
						PlanGuides = if (
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine -and 
							$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
						) {
							@() + (
								$_.PlanGuides | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										Hints = $_.Hints # System.String Hints {get;set;}
										ID = $_.ID # System.Int32 ID {get;}
										IsDisabled = $_.IsDisabled # System.Boolean IsDisabled {get;set;}
										Name = $_.Name # System.String Name {get;set;}
										Parameters = $_.Parameters # System.String Parameters {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										ScopeBatch = $_.ScopeBatch # System.String ScopeBatch {get;set;}
										ScopeObjectName = $_.ScopeObjectName # System.String ScopeObjectName {get;set;}
										ScopeSchemaName = $_.ScopeSchemaName # System.String ScopeSchemaName {get;set;}
										ScopeType = $_.ScopeType # Microsoft.SqlServer.Management.Smo.PlanGuideType ScopeType {get;set;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										Statement = $_.Statement # System.String Statement {get;set;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.PlanGuideCollection PlanGuides {get;}
						#endregion

						#region
						# Sequences not available in SQL 2000
						# Sequences not available in SQL 2005
						# Sequences not available in SQL 2008
						# Sequences not available in SQL 2008 R2
						# Sequences not available in Windows Azure Databases
						Sequences = if (
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine -and 
							$($Server.Information.Version).CompareTo($SQLServer2012) -ge 0
						) {
							@() + (
								$_.Sequences | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										Properties = New-Object -TypeName PSObject -Property @{
											General = New-Object -TypeName PSObject -Property @{
												CacheSize = $_.CacheSize # System.Int32 CacheSize {get;set;}
												CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
												CurrentValue = $_.CurrentValue # System.Object CurrentValue {get;}
												DataType = if ($_.DataType) { 
													New-Object -TypeName PSObject -Property @{
														MaximumLength = $_.DataType.MaximumLength # System.Int32 MaximumLength {get;set;}
														Name = $_.DataType.Name # System.String Name {get;set;}
														NumericPrecision = $_.DataType.NumericPrecision # System.Int32 NumericPrecision {get;set;}
														NumericScale = $_.DataType.NumericScale # System.Int32 NumericScale {get;set;}
														Schema = $_.DataType.Schema # System.String Schema {get;set;}
														SqlDataType = [String](Get-SqlDataTypeValue -SqlDataType $_.DataType.SqlDataType) # Microsoft.SqlServer.Management.Smo.SqlDataType SqlDataType {get;set;}
														XmlDocumentConstraint = [String](Get-XmlDocumentConstraintValue -XmlDocumentConstraint $_.DataType.XmlDocumentConstraint) # Microsoft.SqlServer.Management.Smo.XmlDocumentConstraint XmlDocumentConstraint {get;set;}
													} # Microsoft.SqlServer.Management.Smo.DataType DataType {get;set;}
												} else {
													New-Object -TypeName PSObject -Property @{
														MaximumLength = $null
														Name = $null
														NumericPrecision = $null
														NumericScale = $null
														Schema = $null
														SqlDataType = $null
														XmlDocumentConstraint = $null
													}
												}
												DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
												ID = $_.ID # System.Int32 ID {get;}
												IncrementValue = $_.IncrementValue # System.Object IncrementValue {get;set;}
												IsCycleEnabled = $_.IsCycleEnabled # System.Boolean IsCycleEnabled {get;set;}
												IsExhausted = $_.IsExhausted # System.Boolean IsExhausted {get;}
												IsSchemaOwned = $_.IsSchemaOwned # System.Boolean IsSchemaOwned {get;}
												MaxValue = $_.MaxValue # System.Object MaxValue {get;set;}
												MinValue = $_.MinValue # System.Object MinValue {get;set;}
												Name = $_.Name # System.String Name {get;set;}
												Owner = $_.Owner # System.String Owner {get;set;}
												Schema = $_.Schema # System.String Schema {get;set;}
												SequenceCacheType = [String](Get-SequenceCacheTypeValue -SequenceCacheType $_.SequenceCacheType) # Microsoft.SqlServer.Management.Smo.SequenceCacheType SequenceCacheType {get;set;}
												StartValue = $_.StartValue # System.Object StartValue {get;set;}

											}
											#Permissions
											#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										}

										#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.SequenceEvents Events {get;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}

									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.SequenceCollection Sequences {get;}
						#endregion
					}
					#endregion

					#region
					# ServiceBroker not available in SQL 2000
					# ServiceBroker not available in Windows Azure Databases
					ServiceBroker = New-Object -TypeName psobject -Property @{

						#region
						MessageTypes = if (
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine -and 
							$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
						) {
							@() + (
								$_.ServiceBroker.MessageTypes | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.Broker.MessageTypeEvents Events {get;}
										#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										ID = $_.ID # System.Int32 ID {get;}
										IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
										MessageTypeValidation = [String](Get-BrokerMessageTypeValidationValue -MessageTypeValidation $_.MessageTypeValidation) # Microsoft.SqlServer.Management.Smo.Broker.MessageTypeValidation MessageTypeValidation {get;set;}
										Name = $_.Name # System.String Name {get;set;}
										Owner = $_.Owner # System.String Owner {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Broker.ServiceBroker Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
										ValidationXmlSchemaCollection = $_.ValidationXmlSchemaCollection # System.String ValidationXmlSchemaCollection {get;set;}
										ValidationXmlSchemaCollectionSchema = $_.ValidationXmlSchemaCollectionSchema # System.String ValidationXmlSchemaCollectionSchema {get;set;}
									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.Broker.MessageTypeCollection MessageTypes {get;}
						#endregion

						#region
						ServiceContracts = if (
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine -and 
							$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
						) {
							@() + (
								$_.ServiceBroker.ServiceContracts | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.Broker.ServiceContractEvents Events {get;}
										#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										ID = $_.ID # System.Int32 ID {get;}
										IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
										MessageTypeMappings = @() + (
											$_.MessageTypeMappings | ForEach-Object {
												New-Object -TypeName PSObject -Property @{
													MessageSource = [String](Get-BrokerMessageSourceValue -MessageSource $_.MessageSource) # Microsoft.SqlServer.Management.Smo.Broker.MessageSource MessageSource {get;set;}
													Name = $_.Name # System.String Name {get;set;}
													#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Broker.ServiceContract Parent {get;set;}
													#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
													#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
													#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
													#UserData = $_.UserData	# System.Object UserData {get;set;}									
												}
											}
										) # Microsoft.SqlServer.Management.Smo.Broker.MessageTypeMappingCollection MessageTypeMappings {get;}
										Name = $_.Name # System.String Name {get;set;}
										Owner = $_.Owner # System.String Owner {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Broker.ServiceBroker Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.Broker.ServiceContractCollection ServiceContracts {get;}
						#endregion

						#region
						Queues = if (
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine -and 
							$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
						) {
							@() + (
								$_.ServiceBroker.Queues | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										ActivationExecutionContext = $_.ActivationExecutionContext # Microsoft.SqlServer.Management.Smo.ActivationExecutionContext ActivationExecutionContext {get;set;}
										CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
										DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
										#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.Broker.ServiceQueueEvents Events {get;}
										ExecutionContextPrincipal = $_.ExecutionContextPrincipal # System.String ExecutionContextPrincipal {get;set;}
										#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										FileGroup = $_.FileGroup # System.String FileGroup {get;set;}
										ID = $_.ID # System.Int32 ID {get;}
										IsActivationEnabled = $_.IsActivationEnabled # System.Boolean IsActivationEnabled {get;set;}
										IsEnqueueEnabled = $_.IsEnqueueEnabled # System.Boolean IsEnqueueEnabled {get;set;}
										IsPoisonMessageHandlingEnabled = $_.IsPoisonMessageHandlingEnabled # System.Boolean IsPoisonMessageHandlingEnabled {get;set;}
										IsRetentionEnabled = $_.IsRetentionEnabled # System.Boolean IsRetentionEnabled {get;set;}
										IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
										MaxReaders = $_.MaxReaders # System.Int16 MaxReaders {get;set;}
										Name = $_.Name # System.String Name {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Broker.ServiceBroker Parent {get;set;}
										ProcedureDatabase = $_.ProcedureDatabase # System.String ProcedureDatabase {get;set;}
										ProcedureName = $_.ProcedureName # System.String ProcedureName {get;set;}
										ProcedureSchema = $_.ProcedureSchema # System.String ProcedureSchema {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										RowCount = $_.RowCount # System.Int64 RowCount {get;}
										RowCountAsDouble = $_.RowCountAsDouble # System.Double RowCountAsDouble {get;}
										Schema = $_.Schema # System.String Schema {get;set;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}									
									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.Broker.ServiceQueueCollection Queues {get;}
						#endregion

						#region
						Services = if (
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine -and 
							$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
						) {
							@() + (
								$_.ServiceBroker.Services | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.Broker.BrokerServiceEvents Events {get;}
										#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										ID = $_.ID # System.Int32 ID {get;}
										IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
										Name = $_.Name # System.String Name {get;set;}
										Owner = $_.Owner # System.String Owner {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Broker.ServiceBroker Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										QueueName = $_.QueueName # System.String QueueName {get;set;}
										QueueSchema = $_.QueueSchema # System.String QueueSchema {get;set;}
										ServiceContractMappings = @() + (
											$_.ServiceContractMappings | ForEach-Object {
												New-Object -TypeName PSObject -Property @{
													Name = $_.Name # System.String Name {get;set;}
													#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Broker.BrokerService Parent {get;set;}
													#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
													#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
													#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
													#UserData = $_.UserData	# System.Object UserData {get;set;}
												}
											}
										) # Microsoft.SqlServer.Management.Smo.Broker.ServiceContractMappingCollection ServiceContractMappings {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.Broker.BrokerServiceCollection Services {get;}
						#endregion

						#region
						Routes = if (
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine -and 
							$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
						) {
							@() + (
								$_.ServiceBroker.Routes | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										Address = $_.Address # System.String Address {get;set;}
										BrokerInstance = $_.BrokerInstance # System.String BrokerInstance {get;set;}
										#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.Broker.ServiceRouteEvents Events {get;}
										ExpirationDate = $_.ExpirationDate # System.DateTime ExpirationDate {get;set;}
										#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										ID = $_.ID # System.Int32 ID {get;}
										MirrorAddress = $_.MirrorAddress # System.String MirrorAddress {get;set;}
										Name = $_.Name # System.String Name {get;set;}
										Owner = $_.Owner # System.String Owner {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Broker.ServiceBroker Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										RemoteService = $_.RemoteService # System.String RemoteService {get;set;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.Broker.ServiceRouteCollection Routes {get;}
						#endregion

						#region
						RemoteServiceBindings = if (
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine -and 
							$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
						) {
							@() + (
								$_.ServiceBroker.RemoteServiceBindings | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										CertificateUser = $_.CertificateUser # System.String CertificateUser {get;set;}
										#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.Broker.RemoteServiceBindingEvents Events {get;}
										#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										ID = $_.ID # System.Int32 ID {get;}
										IsAnonymous = $_.IsAnonymous # System.Boolean IsAnonymous {get;set;}
										Name = $_.Name # System.String Name {get;set;}
										Owner = $_.Owner # System.String Owner {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Broker.ServiceBroker Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										RemoteService = $_.RemoteService # System.String RemoteService {get;set;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.Broker.RemoteServiceBindingCollection RemoteServiceBindings {get;}
						#endregion

						#region
						Priorities = if (
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine -and 
							$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
						) {
							@() + (
								$_.ServiceBroker.Priorities | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										ContractName = $_.ContractName # System.String ContractName {get;set;}
										ID = $_.ID # System.Int32 ID {get;}
										LocalServiceName = $_.LocalServiceName # System.String LocalServiceName {get;set;}
										Name = $_.Name # System.String Name {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Broker.ServiceBroker Parent {get;set;}
										PriorityLevel = $_.PriorityLevel # System.Byte PriorityLevel {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										RemoteServiceName = $_.RemoteServiceName # System.String RemoteServiceName {get;set;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.Broker.BrokerPriorityCollection Priorities {get;}
						#endregion

						#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;}
						#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
						#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
						#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
						#UserData = $_.UserData	# System.Object UserData {get;set;}

					}
					#endregion

					#region
					Storage = New-Object -TypeName psobject -Property @{

						#region
						# FullTextCatalogs not available in Windows Azure Databases
						FullTextCatalogs = if (
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine
						) {
							@() + (
								$_.FullTextCatalogs | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										ErrorLogSizeBytes = $_.ErrorLogSize # System.Int32 ErrorLogSize {get;}
										FileGroup = $_.FileGroup # System.String FileGroup {get;set;}
										FullTextIndexSizeMB = $_.FullTextIndexSize # System.Int32 FullTextIndexSize {get;}
										HasFullTextIndexedTables = $_.HasFullTextIndexedTables # System.Boolean HasFullTextIndexedTables {get;}
										ID = $_.ID # System.Int32 ID {get;}
										IsAccentSensitive = $_.IsAccentSensitive # System.Boolean IsAccentSensitive {get;set;}
										IsDefault = $_.IsDefault # System.Boolean IsDefault {get;set;}
										ItemCount = $_.ItemCount # System.Int32 ItemCount {get;}
										Name = $_.Name # System.String Name {get;set;}
										Owner = $_.Owner # System.String Owner {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										PopulationCompletionAgeSeconds = if ($_.PopulationCompletionAge) { $_.PopulationCompletionAge.TotalSeconds } else { $null } # System.TimeSpan PopulationCompletionAge {get;}
										PopulationCompletionDate = $_.PopulationCompletionDate # System.DateTime PopulationCompletionDate {get;}
										PopulationStatus = [String](Get-CatalogPopulationStatusValue -CatalogPopulationStatus $_.PopulationStatus) # Microsoft.SqlServer.Management.Smo.CatalogPopulationStatus PopulationStatus {get;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										RootPath = $_.RootPath # System.String RootPath {get;set;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										UniqueKeyCount = $_.UniqueKeyCount # System.Int32 UniqueKeyCount {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.FullTextCatalogCollection FullTextCatalogs {get;}
						#endregion

						#region
						# FullTextStopLists not available in Windows Azure Databases
						FullTextStopLists = if (
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine
						) {
							@() + (
								$_.FullTextStopLists | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										ID = $_.ID # System.Int32 ID {get;}
										Name = $_.Name # System.String Name {get;set;}
										Owner = $_.Owner # System.String Owner {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}

										StopWords = @() + (
											$_.EnumStopWords() | ForEach-Object {
												New-Object -TypeName PSObject -Property @{
													Language = $_.language # System.String language {get;set;}
													StopWord = $_.stopword # System.String stopword {get;set;}
												}
											}
										)
									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.FullTextStopListCollection FullTextStopLists {get;}
						#endregion

						#region
						# PartitionSchemes not available in SQL 2000
						# PartitionSchemes not available in Windows Azure Databases
						PartitionSchemes = if (
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine -and 
							$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
						) {
							@() + (
								$_.PartitionSchemes | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.PartitionSchemeEvents Events {get;}
										#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										FileGroups = @() + ($_.FileGroups | ForEach-Object { $_ } ) # System.Collections.Specialized.StringCollection FileGroups {get;}
										ID = $_.ID # System.Int32 ID {get;}
										Name = $_.Name # System.String Name {get;set;}
										NextUsedFileGroup = $_.NextUsedFileGroup # System.String NextUsedFileGroup {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										PartitionFunction = $_.PartitionFunction # System.String PartitionFunction {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.PartitionSchemeCollection PartitionSchemes {get;}
						#endregion

						#region
						# PartitionSchemes not available in SQL 2000
						# PartitionFunctions not available in Windows Azure Databases
						PartitionFunctions = if (
							$IsAccessible -eq $true -and 
							$IncludeObjectInformation -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine -and 
							$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0
						) {
							@() + (
								$_.PartitionFunctions | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
										#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.PartitionFunctionEvents Events {get;}
										#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										ID = $_.ID # System.Int32 ID {get;}
										Name = $_.Name # System.String Name {get;set;}
										NumberOfPartitions = $_.NumberOfPartitions # System.Int32 NumberOfPartitions {get;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										PartitionFunctionParameters = @() + (
											$_.PartitionFunctionParameters | Where-Object { $_.ID } | ForEach-Object {
												New-Object -TypeName PSObject -Property @{
													Collation = $_.Collation # System.String Collation {get;set;}
													ID = $_.ID # System.Int32 ID {get;}
													Length = $_.Length # System.Int32 Length {get;set;}
													Name = $_.Name # System.String Name {get;set;}
													NumericPrecision = $_.NumericPrecision # System.Int32 NumericPrecision {get;set;}
													NumericScale = $_.NumericScale # System.Int32 NumericScale {get;set;}
													#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.PartitionFunction Parent {get;set;}
													#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
													#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
													#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
													#UserData = $_.UserData	# System.Object UserData {get;set;}									
												}
											}
										) # Microsoft.SqlServer.Management.Smo.PartitionFunctionParameterCollection PartitionFunctionParameters {get;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										RangeType = [String](Get-RangeTypeValue -RangeType $_.RangeType) # Microsoft.SqlServer.Management.Smo.RangeType RangeType {get;set;}
										RangeValues = @() + (
											$_.RangeValues | ForEach-Object {
												[String]$_
											}
										) # System.Object[] RangeValues {get;set;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
						} # Microsoft.SqlServer.Management.Smo.PartitionFunctionCollection PartitionFunctions {get;}
						#endregion

						<#
						# Search Property Lists not part of SMO
						SearchPropertyLists = if (($IsAccessible -eq $true) -and ($IncludeObjectInformation -eq $true)) {
							@() + (
								$_.SearchPropertyLists | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
									}
								}
							)
						} else {
							$null
						}	# Microsoft.SqlServer.Management.Smo.FederationCollection Federations {get;}
	
	
						Federations = if (($IsAccessible -eq $true) -and ($IncludeObjectInformation -eq $true)) {
							@() + (
								$_.Federations | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
									}
								}
							)
						} else {
							$null
						}	# Microsoft.SqlServer.Management.Smo.FederationCollection Federations {get;}
						#>

					}
					#endregion

					#region
					Security = New-Object -TypeName psobject -Property @{
						User = if ($IsAccessible -eq $true) {
							@() + (
								$_.Users | ForEach-Object {
									New-Object -TypeName psobject -Property @{
										Name = $_.Name # System.String Name {get;set;}
										ID = $_.ID # System.Int32 ID {get;}
										Sid = if ($_.Sid) { [System.BitConverter]::ToString($_.Sid) } else { $null } # System.Byte[] Sid {get;} 
										IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
										HasDbAccess = $_.HasDBAccess # System.Boolean HasDBAccess {get;}

										# AuthenticationType introduced in SQL 2012
										AuthenticationType = if ($_.AuthenticationType) { $_.AuthenticationType.ToString() } else { $null } # Microsoft.SqlServer.Management.Smo.AuthenticationType AuthenticationType {get;}
										LoginType = [String](Get-LoginTypeValue -LoginType $_.LoginType) # Microsoft.SqlServer.Management.Smo.LoginType LoginType {get;}
										UserType = [String](Get-UserTypeValue -UserType $_.UserType) # Microsoft.SqlServer.Management.Smo.UserType UserType {get;set;}

										Login = $_.Login # System.String Login {get;set;}
										Certificate = $_.Certificate # System.String Certificate {get;set;}
										AsymmetricKey = $_.AsymmetricKey # System.String AsymmetricKey {get;set;}
										DefaultSchema = $_.DefaultSchema # System.String DefaultSchema {get;set;}

										# DefaultLanguage introduced in SQL 2012
										DefaultLanguage = if ($_.DefaultLanguage) { $_.DefaultLanguage.Name } else { $null } # Microsoft.SqlServer.Management.Smo.DefaultLanguage DefaultLanguage {get;}
										CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
										DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}										
									}
								}
							)
						} else {
							$null
							#@()
						}
						ApplicationRole = if (
							$IsAccessible -eq $true -and
							$DbEngineType -ieq $StandaloneDbEngine # Not supported in Azure
						) {
							@() + (
								$_.ApplicationRoles | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName psobject -Property @{
										CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
										DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
										DefaultSchema = $_.DefaultSchema # System.String DefaultSchema {get;set;}
										#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.ApplicationRoleEvents Events {get;}
										#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										ID = $_.ID # System.Int32 ID {get;}
										Name = $_.Name # System.String Name {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
							#@()
						}
						DatabaseRole = if ($IsAccessible -eq $true) {
							@() + (
								$_.Roles | ForEach-Object {
									$RoleName = $_.Name

									New-Object -TypeName psobject -Property @{
										CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
										DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
										#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										ID = $_.ID # System.Int32 ID {get;}
										IsFixedRole = $_.IsFixedRole # System.Boolean IsFixedRole {get;}
										Name = $_.Name # System.String Name {get;set;}
										Owner = $_.Owner # System.String Owner {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}

										Member = @() + ($_.EnumMembers() | ForEach-Object { $_ } ) # Is this the best way to do this?
										MemberOf = @() + (
											$DatabaseRoleMemberRole | Where-Object { $_.MemberRoleName -ieq $RoleName } | ForEach-Object { $_.RoleName }
										)

										## EnumRoles not introduced until SMO 2008 but WILL work against SQL 2005
										#MemberOf = if (($SmoMajorVersion -ge 10) -and (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0)) {
										#	@() + ($_.EnumRoles() | ForEach-Object { $_ } ) # Is this the best way to do this?
										#} else { 
										#	#@('Not Supported')
										#} # Is this the best way to do this? 
									}
								}
							)
						} else {
							$null
							#@()
						}
						Certificates = if (
							$IsAccessible -eq $true -and 
							$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0 -and
							$DbEngineType -ieq $StandaloneDbEngine # Not supported in Azure
						) {
							@() + (
								# Normally wouldn't have to check for existence of a property but SMO returns an empty object when enumerating in SQL 2000
								# Other properties (e.g. $_.ApplicationRoles) do not
								# Might have to file a report on Connect for this
								$_.Certificates | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										ActiveForServiceBrokerDialog = $_.ActiveForServiceBrokerDialog # System.Boolean ActiveForServiceBrokerDialog {get;set;}
										#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.CertificateEvents Events {get;}
										ExpirationDate = if ($_.ExpirationDate.CompareTo($SmoEpoch) -le 0) { $null } else { $_.ExpirationDate } # System.DateTime ExpirationDate {get;set;}
										ID = $_.ID # System.Int32 ID {get;}
										Issuer = $_.Issuer # System.String Issuer {get;}

										# $_.LastBackupDate not available in SQL 2005
										LastBackupDate = if (($_.LastBackupDate) -and ($_.LastBackupDate.CompareTo($SmoEpoch) -le 0)) { $null } else { $_.LastBackupDate } # System.DateTime LastBackupDate {get;}
										Name = $_.Name # System.String Name {get;set;}
										Owner = $_.Owner # System.String Owner {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										PrivateKeyEncryptionType = [String](Get-PrivateKeyEncryptionTypeValue -PrivateKeyEncryptionType $_.PrivateKeyEncryptionType) # Microsoft.SqlServer.Management.Smo.PrivateKeyEncryptionType PrivateKeyEncryptionType {get;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										SerialNumber = $_.Serial # System.String Serial {get;}
										Sid = [System.BitConverter]::ToString($_.Sid) # System.Byte[] Sid {get;}
										StartDate = if ($_.StartDate.CompareTo($SmoEpoch) -le 0) { $null } else { $_.StartDate } # System.DateTime StartDate {get;set;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										Subject = $_.Subject # System.String Subject {get;set;}
										Thumbprint = [System.BitConverter]::ToString($_.Thumbprint) # System.Byte[] Thumbprint {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
						}

						# Schemas added in SQL 2005
						Schema = if (($IsAccessible -eq $true) -and (($Server.Information.Version).CompareTo($SQLServer2005) -ge 0)) {
							@() + (
								$_.Schemas | ForEach-Object {
									New-Object -TypeName psobject -Property @{
										#Events = $_.Events	# Microsoft.SqlServer.Management.Smo.SchemaEvents Events {get;}
										#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
										ID = $_.ID # System.Int32 ID {get;}
										IsSystemObject = $_.IsSystemObject # System.Boolean IsSystemObject {get;}
										Name = $_.Name # System.String Name {get;set;}
										Owner = $_.Owner # System.String Owner {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
							#@()
						}

						# Asymmetric Keys added in SQL 2005
						AsymmetricKeys = if (
							$IsAccessible -eq $true -and 
							$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0 -and
							$DbEngineType -ieq $StandaloneDbEngine # Not supported in Azure
						) {
							@() + (
								$_.AsymmetricKeys | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										ID = $_.ID # System.Int32 ID {get;}
										KeyEncryptionAlgorithm = [String](Get-AsymmetricKeyEncryptionAlgorithmValue -AsymmetricKeyEncryptionAlgorithm $_.KeyEncryptionAlgorithm) # Microsoft.SqlServer.Management.Smo.AsymmetricKeyEncryptionAlgorithm KeyEncryptionAlgorithm {get;}
										KeyLength = $_.KeyLength # System.Int32 KeyLength {get;}
										Name = $_.Name # System.String Name {get;set;}
										Owner = $_.Owner # System.String Owner {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										PrivateKeyEncryptionType = [String](Get-PrivateKeyEncryptionTypeValue -PrivateKeyEncryptionType $_.PrivateKeyEncryptionType) # Microsoft.SqlServer.Management.Smo.PrivateKeyEncryptionType PrivateKeyEncryptionType {get;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										ProviderName = $_.ProviderName # System.String ProviderName {get;set;}
										PublicKey = [System.BitConverter]::ToString($_.PublicKey) # System.Byte[] PublicKey {get;}
										Sid = [System.BitConverter]::ToString($_.Sid) # System.Byte[] Sid {get;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										Thumbprint = [System.BitConverter]::ToString($_.Thumbprint) # System.Byte[] Thumbprint {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
						}

						# Symmetric Keys added in SQL 2005
						SymmetricKeys = if (
							$IsAccessible -eq $true -and 
							$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0 -and
							$DbEngineType -ieq $StandaloneDbEngine # Not supported in Azure
						) {
							@() + (
								$_.SymmetricKeys | Where-Object { $_.ID } | ForEach-Object {
									New-Object -TypeName PSObject -Property @{
										CreateDate = $_.CreateDate # System.DateTime CreateDate {get;}
										DateLastModified = $_.DateLastModified # System.DateTime DateLastModified {get;}
										EncryptionAlgorithm = [String](Get-SymmetricKeyEncryptionAlgorithmValue -SymmetricKeyEncryptionAlgorithm $_.EncryptionAlgorithm) # Microsoft.SqlServer.Management.Smo.SymmetricKeyEncryptionAlgorithm EncryptionAlgorithm {get;}
										ID = $_.ID # System.Int32 ID {get;}
										IsOpen = $_.IsOpen # System.Boolean IsOpen {get;}
										KeyGuid = $_.KeyGuid # System.Guid KeyGuid {get;}
										KeyLength = $_.KeyLength # System.Int32 KeyLength {get;}
										Name = $_.Name # System.String Name {get;set;}
										Owner = $_.Owner # System.String Owner {get;set;}
										#Parent = $_.Parent	# Microsoft.SqlServer.Management.Smo.Database Parent {get;set;}
										#Properties = $_.Properties	# Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}
										ProviderName = $_.ProviderName # System.String ProviderName {get;set;}
										#State = $_.State	# Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}
										#Urn = $_.Urn	# Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}
										#UserData = $_.UserData	# System.Object UserData {get;set;}
									}
								}
							)
						} else {
							$null
						}

						# TODO: Finish this!
						#DatabaseAuditSpecifications = $_.DatabaseAuditSpecifications

					}
					#endregion

					## OTHER PROPERTIES THAT MIGHT BE USEFUL
					#ActiveDirectory = $_.ActiveDirectory
					#CaseSensitive = $_.CaseSensitive
					#DatabaseEncryptionKey = $_.DatabaseEncryptionKey
					#DatabaseGuid = $_.DatabaseGuid
					#DatabaseOptions = $_.DatabaseOptions
					##DataSpaceUsageKB = $_.DataSpaceUsage		# Maybe a performance metric to record?
					#DboLogin = $_.DboLogin
					##DefaultFileGroup = $_.DefaultFileGroup
					##DefaultFileStreamFileGroup = $_.DefaultFileStreamFileGroup
					##DefaultFullTextCatalog = $_.DefaultFullTextCatalog
					##DefaultSchema = $_.DefaultSchema
					#Events = $_.Events
					#ExtendedProperties = @() + (Get-ExtendedPropertyInformation -ExtendedPropertyCollection $_.ExtendedProperties) # Microsoft.SqlServer.Management.Smo.ExtendedPropertyCollection ExtendedProperties {get;}
					#IndexSpaceUsage = $_.IndexSpaceUsage
					##IsAccessible = $_.IsAccessible
					##IsDbAccessAdmin = $_.IsDbAccessAdmin
					##IsDbBackupOperator = $_.IsDbBackupOperator
					##IsDbDatareader = $_.IsDbDatareader
					##IsDbDatawriter = $_.IsDbDatawriter
					##IsDbDdlAdmin = $_.IsDbDdlAdmin
					##IsDbDenyDatareader = $_.IsDbDenyDatareader
					##IsDbDenyDatawriter = $_.IsDbDenyDatawriter
					##IsDbManager = $_.IsDbManager
					##IsDbOwner = $_.IsDbOwner
					##IsDbSecurityAdmin = $_.IsDbSecurityAdmin
					#IsFederationMember = $_.IsFederationMember		## FOR AZURE
					##IsLoginManager = $_.IsLoginManager
					#IsUpdateable = $_.IsUpdateable
					#LogReuseWaitStatus = if ($_.LogReuseWaitStatus) { $_.LogReuseWaitStatus.ToString() } else { $null }
					#MasterKey = $_.MasterKey
					##PrimaryFilePath = $_.PrimaryFilePath
					#RecoveryForkGuid = $_.RecoveryForkGuid
					#SearchPropertyLists = $_.SearchPropertyLists
					##UserName = $_.UserName
					##Version = $_.Version

				}
			) 

			# Turn on prefetch at the server level
			#$Server.SetDefaultInitFields($true)
		}
	}
	catch {
		throw
	}
}

function Get-DatabaseMailInformation {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)] 
		[Microsoft.SqlServer.Management.Smo.Server]
		$Server
	)
	try {

		$DbEngineType = [String](Get-DatabaseEngineTypeValue -DatabaseEngineType $Server.ServerType) 

		# Database Mail added in SQL 2005
		# Not available in Azure
		if (
			$($Server.Information.Version).CompareTo($SQLServer2005) -ge 0 -and
			$DbEngineType -ieq $StandaloneDbEngine
		) {
			Write-Output (
				New-Object -TypeName psobject -Property @{
					Accounts = @() + (
						$Server.Mail.Accounts | ForEach-Object {
							$Account = $_
							$Account.MailServers | ForEach-Object {
								New-Object -TypeName psobject -Property @{
									ID = $Account.ID # System.Int32 ID {get;}
									AccountName = $Account.Name # System.String Name {get;set;}
									Description = $Account.Description # System.String Description {get;set;}
									OutgoingSmtpServer = New-Object -TypeName psobject -Property @{
										EmailAddress = $Account.EmailAddress # System.String EmailAddress {get;set;}
										DisplayName = $Account.DisplayName # System.String DisplayName {get;set;}
										ReplyToAddress = $_.ReplyToAddress # System.String ReplyToAddress {get;set;}
										ServerName = $_.Name # System.String Name {get;set;}
										ServerType = $_.ServerType # System.String ServerType {get;}
										PortNumber = $_.Port # System.Int32 Port {get;set;}
										SslConnectionRequired = $_.EnableSsl # System.Boolean EnableSsl {get;set;}
									}
									SmtpAuthentication = New-Object -TypeName psobject -Property @{
										AuthenticationType = if ($_.UseDefaultCredentials -eq $true) { 
											'Windows Authentication using Database Engine service credentials'
										} elseif ($_.UserName -ieq [String]::Empty) {
											'Anonymous authentication'
										} else {
											'Basic authentication'
										}
										UseDefaultCredentials = $_.UseDefaultCredentials # System.Boolean UseDefaultCredentials {get;set;}
										UserName = $_.UserName # System.String UserName {get;set;}
									}
								}
							}
						}
					)
					Profiles = @() + (
						$Server.Mail.Profiles | Where-Object { $_.ID } | ForEach-Object {
							New-Object -TypeName psobject -Property @{
								ProfileName = $_.Name # System.String Name {get;set;}
								Description = $_.Description # System.String Description {get;set;}
								ForceDeleteForActiveProfiles = $_.ForceDeleteForActiveProfiles # System.Boolean ForceDeleteForActiveProfiles {get;set;}
								ID = $_.ID # System.Int32 ID {get;}
								IsBusyProfile = $_.IsBusyProfile # System.Boolean IsBusyProfile {get;}
								Accounts = @() + ($_.EnumAccounts() | ForEach-Object { $_.AccountName } )
								Security = @() + (
									$_.EnumPrincipals() | ForEach-Object {
										New-Object -TypeName psobject -Property @{
											PrincipalName = $_.PrincipalName # System.String PrincipalName {get;set;}
											PrincipalID = $_.PrincipalID # System.Int32 PrincipalID {get;set;}
											IsDefault = $_.IsDefault # System.Boolean IsDefault {get;set;}
											IsPublic = if ($_.PrincipalName -ieq 'guest') { $true } else { $false }
										}
									}
								)
							}
						}
					)

					# Originally I had added this as a ScriptProperty, but when you deserialize and reserialize
					# using Export-Clixml and Import-Clixml it gets turned into a NoteProperty that isn't populated as expected
					ProfileSecurity = @() + (
						$Server.Mail.Profiles | Where-Object { $_.ID } | ForEach-Object {
							$_.EnumPrincipals() | ForEach-Object {
								New-Object -TypeName psobject -Property @{
									ProfileName = $_.ProfileName # System.String PrincipalName {get;set;}
									ProfileId = $_.ProfileID # System.Int32 ProfileID {get;set;}
									PrincipalName = $_.PrincipalName # System.String PrincipalName {get;set;}
									PrincipalID = $_.PrincipalID # System.Int32 PrincipalID {get;set;}
									IsDefault = $_.IsDefault # System.Boolean IsDefault {get;set;}
									IsPublic = if ($_.PrincipalName -ieq 'guest') { $true } else { $false }
								}
							}
						})

					# For some reason SMO exposes these as name\value pairs instead of as properties
					# As of SQL 2012 there's no mechanism to add new configuration values so I decided to go ahead and transform them here
					# Return NULLs if server is Express edition
					ConfigurationValues = if ($Server.Information.Edition -inotlike 'Express*') {
						New-Object -TypeName psobject -Property @{
							AccountRetryAttempts = $Server.Mail.ConfigurationValues['AccountRetryAttempts'].Value # System.String Value {get;set;}
							AccountRetryDelaySeconds = $Server.Mail.ConfigurationValues['AccountRetryDelay'].Value # System.String Value {get;set;}
							DatabaseMailExeMinimumLifeTimeSeconds = $Server.Mail.ConfigurationValues['DatabaseMailExeMinimumLifeTime'].Value # System.String Value {get;set;}
							DefaultAttachmentEncoding = $Server.Mail.ConfigurationValues['DefaultAttachmentEncoding'].Value # System.String Value {get;set;}
							LoggingLevel = $Server.Mail.ConfigurationValues['LoggingLevel'].Value # System.String Value {get;set;}
							MaxFileSizeBytes = $Server.Mail.ConfigurationValues['MaxFileSize'].Value # System.String Value {get;set;}
							ProhibitedExtensions = $Server.Mail.ConfigurationValues['ProhibitedExtensions'].Value # System.String Value {get;set;}
						}
					} else {
						New-Object -TypeName psobject -Property @{
							AccountRetryAttempts = $null
							AccountRetryDelaySeconds = $null
							DatabaseMailExeMinimumLifeTimeSeconds = $null
							DefaultAttachmentEncoding = $null
							LoggingLevel = $null
							MaxFileSizeBytes = $null
							ProhibitedExtensions = $null
						}
					}

					# 				$Server.Mail.ConfigurationValues | ForEach-Object {
					# 					New-Object -TypeName psobject -Property @{
					# 						Name = $_.Name # System.String Name {get;set;}
					# 						Description = $_.Description # System.String Description {get;set;}
					# 						Value = $_.Value # System.String Value {get;set;}
					# 					}
					# 				}
				} 
			)
		} else {
			Write-Output (
				New-Object -TypeName psobject -Property @{
					Accounts = @()
					Profiles = @()
					ConfigurationValues = @()
				}
			)
		}

		# 		$DatabaseMailInformation | Add-Member -MemberType ScriptProperty -Name ProfileSecurity -Value {
		# 			$this.Profiles | ForEach-Object {
		# 				$CurrentProfile = $_
		# 				$CurrentProfile.Security | ForEach-Object {
		# 					New-Object -TypeName psobject -Property @{
		# 						PrincipalName = $_.PrincipalName
		# 						PrincipalID = $_.PrincipalID
		# 						ProfileName = $CurrentProfile.ProfileName
		# 						ProfileID = $CurrentProfile.ID
		# 						IsDefault = $_.IsDefault
		# 						IsPublic = $_.IsPublic
		# 					}
		# 				}
		# 			}
		# 		}

		Write-Output -InputObject $DatabaseMailInformation
	}
	catch {
		Throw
	}
}



######################
# PUBLIC FUNCTIONS
######################

function Get-SqlServerDatabaseEngineInformation {
	<#
	.SYNOPSIS
		Gets information about a Windows Azure or Standalone SQL Server database engine instance.

	.DESCRIPTION
		The Get-SqlServerDatabaseEngineInformation function uses SQL Server Shared Management Objects (SMO) to retrieve comprehensive information about SQL Server database engine instance.
		
		This function is compatible with all versions of SMO and can retrieve information from SQL Server 2000 or higher and Windows Azure SQL Database (if using SMO 2008 or higher).
		
		This function works best when using a version of SMO that matches or is higher than the version of the SQL Server instance that information is being collected from.
		
		The latest version of SMO can be downloaded from http://www.microsoft.com/en-us/download/details.aspx?id=29065
		Note that SMO also requires the Microsoft SQL Server System CLR Types which can be downloaded from the same page
		
		Get-SqlServerDatabaseEngineInformation returns details about the following items (when available): 
			
			Database Engine service configuration
			Server Configuration:
				General details, Memory, Processors, Security, Connections, Database Settings, Advanced Settings, High Availability (Clustering & AlwaysOn)
			Security: 
				Logins, Roles, Credentials, Audits, Audit Specifications
			Server Objects:
				Endpoints, Linked Servers, Startup Procedures, Server Triggers
			Server Management:
				Resource Governor, SQL Trace, Trace Flags, Database Mail (Accounts, Profiles, Security, Configuration)
			Databases:
				Configuration:
					General details, Files, Filegroups, Options, Change Tracking, Permissions, Mirroring, AlwaysOn
				Tables (Properties, Checks, Columns, Foreign Keys, Full Text Indexes, Indexes, Statistics, Triggers)
				Views (Properties, Columns, Full Text Indexes, Indexes, Statistics, Triggers, Definition)
				Synonyms (Properties, Definition)
				Programmability:
					Stored Procedures (Properties, Parameters, Definition)
					Extended Stored Procedures (Properties)
					Functions (Properties, Checks, Columns, Indexes, Parameters, Definition)
					Database Triggers (Properties, Definition)
					Assemblies
					Types:
						User Defined Aggregates (Properties, Parameters)
						User Defined Data Types (Properties)
						User Defined Table Types (Properties, Checks, Columns, Indexes)
						User Defined Types (Properties)
						XML Schema Collections (Properties, Definition)
						Rules (Properties, Definition)
						Defaults (Properties, Definition)
						Plan Guides (Properties, Definition)
						Sequences (Properties)
				Service Broker:
					Message Types
					Service Contracts
					Queues
					Services
					Routes
					Remote Service Bindings
					Priorities
				Storage:
					Full Text Catalogs
					Full Text Stop Lists
					Partition Schemes
					Partition Functions
				Security:
					Users
					Application Roles
					Database Roles
					Certificates
					Schemas
					Asymmetric Keys
					Symmetric Keys
	
			SQL Agent service configuration
			SQL Agent Configuration:
				General details, Advanced Settings, Alert System, Job System, Connections, History
			SQL Agent Jobs:
				General Details, Steps, Schedules, Alerts, Notifications
			SQL Agent Alerts:
				General Details, Responses, Options, History
			SQL Agent Operators: 
				General Details, Notifications, History

				
	.PARAMETER  InstanceName
		The database engine instance to connect to.

	.PARAMETER  IpAddress
		The IP Address to use when connecting to the database engine.
		
		If not provided, the value provided for InstanceName will be used to connect instead.

	.PARAMETER  Port
		The port to use when connecting to the database engine.
		
		If not provided, port 1433 (the default port for SQL Server) will be used.

	.PARAMETER  Username
		SQL Server username to use when connecting to an instance. 
		
		Windows authentication will be used to connect if this parameter is not provided.

	.PARAMETER  Password
		SQL Server password to use when connecting to an instance.

	.PARAMETER  StopAtErrorCount
		The number of errors that can be encountered before this function calls it quits and returns.

	.PARAMETER  IncludeDatabaseObjectPermissions
		Includes database object level permissions (System object permissions included only if -IncludeDatabaseSystemObjects is also provided)

	.PARAMETER  IncludeDatabaseObjectInformation
		Include information about the following database objects:
			Tables
			Views
			Synonyms
			Stored Procedures
			Extended Stored Procedures
			Functions
			Database Triggers
			Assemblies
			User Defined Aggregates
			User Defined Data Types
			User Defined Table Types
			User Defined Types
			XML Schema Collections
			Rules
			Defaults
			Plan Guides
			Sequences
			Service Broker Message Types
			Service Broker Service Contracts
			Service Broker Queues
			Service Broker Services
			Service Broker Routes
			Service Broker Remote Service Bindings
			Service Broker Priorities
			Full Text Catalogs
			Full Text Stop Lists
			Partition Schemes
			Partition Functions
			
		System objects are included only if -IncludeDatabaseSystemObjects is also provided

	.PARAMETER  IncludeDatabaseSystemObjects
		Include system objects when retrieving database object information. 
		
		This has no effect if neither -IncludeDatabaseObjectInformation nor -IncludeDatabaseObjectPermissions are specified.

	.EXAMPLE
		Get-SqlServerDatabaseEngineInformation -InstanceName $env:COMPUTERNAME
		
		Description
		-----------
		This will retrieve information about the default instance on the local machine.
		
		Windows Authentication will be used to connect to the instance.
		
		Database objects will NOT be included in the results.

	.EXAMPLE
		Get-SqlServerDatabaseEngineInformation -InstanceName $env:COMPUTERNAME\SQL2012 -Username sa -Password BetterNotBeBlank
		
		Description
		-----------
		This will retrieve information about the named instance "SQL2012" on the local machine.
		
		SQL authentication (username = "sa", password = "BetterNotBeBlank") will be used to connect to the instance.
		
		Database objects will NOT be included in the results.

	.EXAMPLE
		Get-SqlServerDatabaseEngineInformation -InstanceName $env:COMPUTERNAME -Port 1344
		
		Description
		-----------
		This will retrieve information about the default instance running on port 1344 on the local machine.
		
		Windows Authentication will be used to connect to the instance.
		
		Database objects will NOT be included in the results.

	.EXAMPLE
		Get-SqlServerDatabaseEngineInformation -InstanceName $env:COMPUTERNAME -StopAtErrorCount 5
		
		Description
		-----------
		This will retrieve information about the default instance on the local machine.
		
		Windows Authentication will be used to connect to the instance.
		
		Get-SqlServerDatabaseEngineInformation will stop collecting information after encountering 5 errors.
		
		Database objects will NOT be included in the results.

	.EXAMPLE
		Get-SqlServerDatabaseEngineInformation -InstanceName $env:COMPUTERNAME -IncludeDatabaseObjectInformation
		
		Description
		-----------
		This will retrieve information about the default instance on the local machine.
		
		Windows Authentication will be used to connect to the instance.
		
		Database objects (EXCLUDING system objects) will be included in the results.

	.EXAMPLE
		Get-SqlServerDatabaseEngineInformation -InstanceName $env:COMPUTERNAME -IncludeDatabaseObjectInformation -IncludeDatabaseSystemObjects
		
		Description
		-----------
		This will retrieve information about the default instance on the local machine.
		
		Windows Authentication will be used to connect to the instance.
		
		Database objects (INCLUDING system objects) will be included in the results.
		

	.OUTPUTS
		System.Management.Automation.PSObject

	.NOTES

#>
	[CmdletBinding(DefaultParametersetName='WindowsAuthentication')]
	param(
		[Parameter(Mandatory=$false, ParameterSetName = 'SQLAuthentication')] 
		[Parameter(Mandatory=$false, ParameterSetName = 'WindowsAuthentication')] 
		[alias('instance')]
		[string]
		$InstanceName = '(local)'
		,
		[Parameter(Mandatory=$false)] 
		[System.Net.IPAddress]
		$IpAddress = $null
		,
		[Parameter(Mandatory=$false)] 
		[int]
		$Port = $null
		,
		[Parameter(Mandatory=$true, ParameterSetName = 'SQLAuthentication')]
		[ValidateNotNull()]
		[System.String]
		$Username
		,
		[Parameter(Mandatory=$true, ParameterSetName = 'SQLAuthentication')]
		[ValidateNotNull()]
		[System.String]
		$Password
		,
		[Parameter(Mandatory=$false)] 
		[alias('errors')]
		[int]
		$StopAtErrorCount = 3
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

		$Server = $null
		$Connection = $null
		$IsConnected = $false
		$ErrorCount = 0
		$LastErrorActionPreference = $ErrorActionPreference
		$ParameterHash = $null

		$ServerInformation = New-Object -TypeName psobject -Property @{
			Server = New-Object -TypeName psobject -Property @{
				Service = $null
				Configuration = $null
				Databases = @()
				Security = $null
				ServerObjects = New-Object -TypeName psobject -Property @{
					LinkedServers = @()
					Endpoints = @()
					StartupProcedures = @()
					Triggers = @()
				}
				Management = New-Object -TypeName psobject -Property @{
					ResourceGovernor = $null
					SQLTrace = @()
					TraceFlags = @()
					DatabaseMail = $null
				}
			}
			Agent = New-Object -TypeName psobject -Property @{
				Service = $null
				Configuration = $null
				Jobs = @()
				Alerts = @()
				Operators = @()
			}
			HasDatabaseObjectInformation = $IncludeDatabaseObjectInformation
			ServerName = $InstanceName
			ScanDateUTC = [DateTime]::UtcNow
			ScanErrorCount = 0
		}


		if ($PSCmdlet.ParameterSetName -eq 'SQLAuthentication') {
			$Connection = Get-SqlConnection -Instance $InstanceName -IpAddress $IpAddress -Port $Port -Username $Username -Password $Password
		} else {
			$Connection = Get-SqlConnection -Instance $InstanceName -IpAddress $IpAddress -Port $Port
		}


		$ErrorActionPreference = 'Stop'

		try {
			$Server = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList $Connection

			while ($IsConnected -ne $true) {
				try {
					$Server.ConnectionContext.Connect()
					$IsConnected = $true
				}
				catch {
					$ServerInformation.ScanErrorCount++
					Write-SqlServerDatabaseEngineInformationLog -Message "Failed to connect to $InstanceName" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { Throw }
				}
			}

			if ($IsConnected -eq $true) {

				# Instruct SMO to fully instantiate objects
				# See http://blogs.msdn.com/b/mwories/archive/2005/04/22/smoperf1.aspx
				#	and http://blogs.msdn.com/b/mwories/archive/2005/05/02/smoperf2.aspx
				$Server.SetDefaultInitFields($true)

				# Disable retrieval of all fields for Data Files - otherwise SMO no likey!
				$Server.SetDefaultInitFields($('Microsoft.SqlServer.Management.Smo.DataFile' -as [Type]), $false)

				# Disable retrieval of all fields for Extended Stored Procedures - otherwise SMO no likey!
				$Server.SetDefaultInitFields($('Microsoft.SqlServer.Management.Smo.ExtendedStoredProcedure' -as [Type]), $false)



				######################
				# SERVER
				######################

				# Service
				######################

				# Get-SqlServerServiceInformation
				#region
				try {
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Gathering Server service information" -MessageLevel Verbose
					$ServerInformation.Server.Service = Get-SqlServerServiceInformation -Server $Server
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering Server service information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { Throw } 
				}
				#endregion


				# Configuration
				######################

				# Get-ServerConfigurationInformation
				#region
				try {
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Gathering Server configuration information" -MessageLevel Verbose
					$ServerInformation.Server.Configuration = Get-ServerConfigurationInformation -Server $Server

					# Update ServerName to match name retrieved from instance
					$ServerInformation.ServerName = $ServerInformation.Server.Configuration.General.Name

					# Add a NoteProperty for ComputerName
					# Note: I previously had created this as a ScriptProperty refrencing $this.Server.Configuration.General.ComputerName
					# but it blew up when this module was used to take a SQL Server Inventory.
					# Since this is just a string it's not that big of a deal to just copy it
					$ServerInformation | Add-Member -MemberType NoteProperty -Name ComputerName -Value $ServerInformation.Server.Configuration.General.ComputerName 

				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering Server configuration information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { Throw } 
				}
				#endregion


				# Databases
				######################

				# Get-DatabaseInformation
				#region
				try {
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Gathering Database information" -MessageLevel Verbose

					$ParameterHash = @{
						Server = $Server
						IncludeObjectPermissions = $IncludeDatabaseObjectPermissions
						IncludeObjectInformation = $IncludeDatabaseObjectInformation
						IncludeSystemObjects = $IncludeDatabaseSystemObjects
					}

					$ServerInformation.Server.Databases += Get-DatabaseInformation @ParameterHash
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering Database information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { Throw } 
				}
				#endregion


				# Security
				######################

				# Get-SqlServerSecurityInformation
				#region
				try {
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Gathering Server security information" -MessageLevel Verbose
					$ServerInformation.Server.Security = Get-SqlServerSecurityInformation -Server $Server
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering Server security information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { Throw } 
				}
				#endregion



				# Server Objects
				######################				

				# Get-EndPointInformation
				#region
				try {
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Gathering Endpoint information" -MessageLevel Verbose
					$ServerInformation.Server.ServerObjects.Endpoints += Get-EndPointInformation -Server $Server
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering Endpoint information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { Throw } 
				}
				#endregion


				# Get-LinkedServerInformation
				#region
				try {
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Gathering Server linked server information" -MessageLevel Verbose
					$ServerInformation.Server.ServerObjects.LinkedServers += Get-LinkedServerInformation -Server $Server
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering Server linked server information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { Throw } 
				}
				#endregion 


				# Get-StartupProcedureInformation
				#region
				try {
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Gathering Server startup procedure information" -MessageLevel Verbose
					$ServerInformation.Server.ServerObjects.StartupProcedures += Get-StartupProcedureInformation -Server $Server
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering Server startup procedure information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { Throw } 
				}
				#endregion


				# Get-ServerTriggerInformation
				#region
				try {
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Gathering Server Trigger information" -MessageLevel Verbose
					$ServerInformation.Server.ServerObjects.Triggers += Get-ServerTriggerInformation -Server $Server
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering Server Trigger information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { Throw } 
				}
				#endregion 



				# Management
				######################	

				# Get-ResourceGovernorInformation
				#region
				try {
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Gathering Resource Governor information" -MessageLevel Verbose
					$ServerInformation.Server.Management.ResourceGovernor = Get-ResourceGovernorInformation -Server $Server
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering Resource Governor information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { Throw } 
				}
				#endregion


				# Get-SQLTraceInformation
				#region
				try {
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Gathering SQL Trace information" -MessageLevel Verbose
					$ServerInformation.Server.Management.SQLTrace += Get-SQLTraceInformation -Server $Server
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering SQL Trace information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { Throw } 
				}
				#endregion


				# Get-TraceFlagInformation
				#region
				try {
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Gathering Trace Flag information" -MessageLevel Verbose
					$ServerInformation.Server.Management.TraceFlags += Get-TraceFlagInformation -Server $Server
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering Trace Flag information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { Throw } 
				}
				#endregion


				# Get-DatabaseMailInformation
				#region
				try {
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Gathering Database Mail information" -MessageLevel Verbose
					$ServerInformation.Server.Management.DatabaseMail = Get-DatabaseMailInformation -Server $Server
				}
				catch {
					$ErrorRecord = $_.Exception.ErrorRecord
					Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering Database Mail information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
					if (++$ErrorCount -ge $StopAtErrorCount) { Throw } 
				}
				#endregion



				######################
				# AGENT
				######################

				if (
					$ServerInformation.Server.Configuration.General.ServerType -ieq 'standalone' -and
					$Server.Information.Edition -inotlike 'Express*' 
				) {

					# Service
					######################

					# Get-SqlAgentServiceInformation
					#region
					try {
						Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Gathering SQL Agent service information" -MessageLevel Verbose
						$ServerInformation.Agent.Service = Get-SqlAgentServiceInformation -JobServer $Server.JobServer
					}
					catch {
						$ErrorRecord = $_.Exception.ErrorRecord
						Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering SQL Agent service information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
						if (++$ErrorCount -ge $StopAtErrorCount) { Throw } 
					}
					#endregion


					# Configuration
					######################

					# Get-SqlAgentConfigurationInformation
					#region
					try {
						Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Gathering SQL Agent configuration information" -MessageLevel Verbose
						$ServerInformation.Agent.Configuration = Get-SqlAgentConfigurationInformation -JobServer $Server.JobServer
					}
					catch {
						$ErrorRecord = $_.Exception.ErrorRecord
						Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering SQL Agent configuration information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
						if (++$ErrorCount -ge $StopAtErrorCount) { Throw } 
					}
					#endregion


					# Jobs
					######################

					# Get-SqlAgentJobInformation
					#region
					try {
						Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Gathering SQL Agent job information" -MessageLevel Verbose
						$ServerInformation.Agent.Jobs += Get-SqlAgentJobInformation -JobServer $Server.JobServer
					}
					catch {
						$ErrorRecord = $_.Exception.ErrorRecord
						Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering SQL Agent job information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
						if (++$ErrorCount -ge $StopAtErrorCount) { Throw } 
					}
					#endregion


					# Alerts
					######################

					# Get-SqlAgentAlertInformation
					#region
					try {
						Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Gathering SQL Agent alert information" -MessageLevel Verbose
						$ServerInformation.Agent.Alerts += Get-SqlAgentAlertInformation -JobServer $Server.JobServer

						# Update references to alerts for each job
						$ServerInformation.Agent.Jobs | Where-Object { $_.Alerts } | ForEach-Object {
							$_.Alerts = $_.Alerts | ForEach-Object {
								$AlertId = $_.ID
								$ServerInformation.Agent.Alerts | Where-Object { $_.General.ID -eq $AlertId }
							}
						} 
					}
					catch {
						$ErrorRecord = $_.Exception.ErrorRecord
						Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering SQL Agent alert information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
						if (++$ErrorCount -ge $StopAtErrorCount) { Throw } 
					}
					#endregion


					# Operators
					######################

					# Get-SqlAgentOperatorInformation
					#region
					try {
						Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Gathering SQL Agent operator information" -MessageLevel Verbose
						$ServerInformation.Agent.Operators += Get-SqlAgentOperatorInformation -JobServer $Server.JobServer
					}
					catch {
						$ErrorRecord = $_.Exception.ErrorRecord
						Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering SQL Agent operator information: $($ErrorRecord.Exception.Message) ($([System.IO.Path]::GetFileName($ErrorRecord.InvocationInfo.ScriptName)) line $($ErrorRecord.InvocationInfo.ScriptLineNumber), char $($ErrorRecord.InvocationInfo.OffsetInLine))" -MessageLevel Warning
						if (++$ErrorCount -ge $StopAtErrorCount) { Throw }
					}
					#endregion

				}
			}
		}
		catch {
			# If we hit this point we've reached the max error threshold
			Write-SqlServerDatabaseEngineInformationLog -Message "[$InstanceName] Error gathering instance information - max error threshold reached ($StopAtErrorCount)" -MessageLevel Warning
		}
		finally {
			# Record the number of scan errors
			$ServerInformation.ScanErrorCount = $ErrorCount

			# Reset the $ErrorActionPreference and return the $ServerInformation object
			$ErrorActionPreference = $LastErrorActionPreference

			Write-Output $ServerInformation

			if ($IsConnected -eq $true) {
				$Server.ConnectionContext.Disconnect()
			}
		}

	}
}



######################
# RUN WHEN MODULE IS LOADED
######################

# Load SMO assembly, and if we're running SQL 2008 DLLs or higher load the SMOExtended and SQLWMIManagement libraries
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | ForEach-Object {
	$SmoMajorVersion = $_.GetName().Version.Major
	[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SQLEnum') | Out-Null
	if ($SmoMajorVersion -ge 10) {
		[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMOExtended') | Out-Null
		[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SQLWMIManagement') | Out-Null
	}
}

# Now check which SMO assemblies are loaded and set $SmoMajorVersion to the lowest version
[System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.FullName -ilike 'Microsoft.SqlServer.SMO, Version=*' } | ForEach-Object {
	if ($_.GetName().Version.Major -lt $SmoMajorVersion) {
		$SmoMajorVersion = $_.GetName().Version.Major
		Write-SqlServerDatabaseEngineInformationLog -Message "Multiple versions of Microsoft.SqlServer.SMO are loaded; reverting to the lowest version to avoid problems" -MessageLevel Warning
	}
}