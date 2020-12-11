# Class for SQL Server and Port
# Used for altering ServerNames from input
class CGISQLServerShort {
    [string]$ServerName
    [string]$AgencyName
    [string]$Operational
    [string]$Description
    [string]$ProductVersion
    [string]$Edition
    [string]$SQLInstallDate; `
`
    CGISQLServerShort(
        [string]$ServerName,
        [string]$AgencyName,
        [string]$Operational,
        [string]$Description,
        [string]$ProductVersion,
        [string]$Edition,
        [string]$SQLInstallDate
    ){ 
        $this.ServerName = $ServerName
        $this.AgencyName = $AgencyName
        $this.Operational = $Operational
        $this.Description = $Description
        $this.ProductVersion = $ProductVersion
        $this.Edition = $Edition
        $this.SQLInstallDate = $SQLInstallDate
    }
    CGISQLServerShort(
        [string]$ServerName
    ){ 
        $this.ServerName = $ServerName
    }
}
# Class for SQL Job
class CGISQLJob {
    [string]$InstanceName
    [string]$AgencyName
    [string]$JobName
    [string]$FirstStepCommand
    [string]$JobEnabled
    [string]$NotifyEventLog
    [string]$NotifyEmail
    [string]$JobCreatedDateTime
    [string]$JobModifiedDateTime
    [string]$Status
    [string]$RunDateTime
    [string]$RunDurationInMinutes
    [string]$RunDurationInSeconds; `
`
    CGISQLJob(
        [string]$InstanceName,
        [string]$AgencyName,
        [string]$JobName,
        [string]$FirstStepCommand,
        [string]$JobEnabled,
        [string]$NotifyEventLog,
        [string]$NotifyEmail,
        [string]$JobCreatedDateTime,
        [string]$JobModifiedDateTime,
        [string]$Status,
        [string]$RunDateTime,
        [string]$RunDurationInMinutes,
        [string]$RunDurationInSeconds
    ){ 
        $this.InstanceName = $InstanceName
        $this.AgencyName = $AgencyName
        $this.JobName = $JobName
        $this.FirstStepCommand = $FirstStepCommand
        $this.JobEnabled = $JobEnabled
        $this.NotifyEventLog = $NotifyEventLog
        $this.NotifyEmail = $NotifyEmail
        $this.JobCreatedDateTime = $JobCreatedDateTime
        $this.JobModifiedDateTime = $JobModifiedDateTime
        $this.Status = $Status
        $this.RunDateTime = $RunDateTime
        $this.RunDurationInMinutes = $RunDurationInMinutes
        $this.RunDurationInSeconds = $RunDurationInSeconds
    }
    CGISQLJob(
        [string]$InstanceName,
        [string]$AgencyName,
        [string]$JobName
     ){
        $this.InstanceName = $InstanceName
        $this.AgencyName = $AgencyName
        $this.JobName = $JobName
       }
}
#Class for SQL Alerts
class CGISQLAlert {
    [string]$InstanceName 
    [string]$AgencyName
    [string]$Name
    [string]$EventSource 
    [string]$MessageID
    [string]$Severity 
    [string]$Enabled  
    [string]$LastOccurrenceDate 
    [string]$LastResponseDate
    [string]$NotificationMessage
    [string]$DatabaseName 
    [string]$OccurrenceCounter  
    [string]$OccurrenceResetCounterDate
    [string]$OccurrenceResetCounterTime 
    [string]$NumberofNotificationContacts ; `
`
    CGISQLAlert(
        [string]$InstanceName, 
        [string]$AgencyName,
        [string]$Name,
        [string]$EventSource, 
        [string]$MessageID,
        [string]$Severity,
        [string]$Enabled, 
        [string]$LastOccurrenceDate, 
        [string]$LastResponseDate,
        [string]$NotificationMessage,
        [string]$DatabaseName, 
        [string]$OccurrenceCounter,  
        [string]$OccurrenceResetCounterDate,
        [string]$OccurrenceResetCounterTime, 
        [string]$NumberofNotificationContacts
    ){
        $this.InstanceName = $InstanceName
        $this.AgencyName = $AgencyName
        $this.Name = $Name
        $this.EventSource = $EventSource
        $this.MessageID = $MessageID
        $this.Severity = $Severity
        $this.Enabled = $Enabled
        $this.LastOccurrenceDate = $LastOccurrenceDate 
        $this.LastResponseDate = $LastResponseDate
        $this.NotificationMessage = $NotificationMessage
        $this.DatabaseName = $DatabaseName
        $this.OccurrenceCounter = $OccurrenceCounter
        $this.OccurrenceResetCounterDate = $OccurrenceResetCounterDate
        $this.OccurrenceResetCounterTime = $OccurrenceResetCounterTime
        $this.NumberofNotificationContacts = $NumberofNotificationContacts
    }
    CGISQLAlert(
        [string]$InstanceName,
        [string]$AgencyName,
        [string]$Name
     ){
        $this.InstanceName = $InstanceName
        $this.AgencyName = $AgencyName
        $this.Name = $Name
    }
}
#SQL to get information on SQL Servers
$sqlsmtsrv = "
SELECT  
    SERVERPROPERTY('ServerName') AS InstanceName,
    '/ReplaceText/' as 'AgencyName',
    left(left(@@VERSION,charindex(' - ',@@VERSION,1)-1),charindex(' (',@@VERSION,1)-1) as ProductName,
    SERVERPROPERTY('ProductVersion') AS ProductVersion,  
    SERVERPROPERTY('ProductLevel') AS ProductLevel,
    SERVERPROPERTY('ProductBuild') AS ProductBuild,
    SERVERPROPERTY('ProductLevel') AS SP,
    CASE SERVERPROPERTY('EngineEdition')
        WHEN 1 THEN 'PERSONAL'
        WHEN 2 THEN 'STANDARD'
        WHEN 3 THEN 'ENTERPRISE'
        WHEN 4 THEN 'EXPRESS'
        WHEN 5 THEN 'SQL DATABASE'
        WHEN 6 THEN 'SQL DATAWAREHOUSE'
    END AS Edition, 
    (SELECT convert(varchar(2),month(create_date)) + '/' + convert(varchar(2),day(create_date))  + '/' + convert(varchar(4),year(create_date))FROM sys.server_principals WHERE sid = 0x010100000000000512000000) as 'SQLInstallDate',
    CASE SERVERPROPERTY('IsHadrEnabled')
        WHEN 0 THEN ''
        WHEN 1 THEN 'Enabled'
        ELSE ''
    END AS AlwaysOn,
	CASE SERVERPROPERTY('IsHadrEnabled') 
		WHEN 1 THEN 
			CASE SERVERPROPERTY('HadrManagerStatus')
				WHEN 0 THEN 'Not started, pending communication'
				WHEN 1 THEN 'Started and running'
				WHEN 2 THEN 'Not started and failed'
				ELSE ''
			END
		ELSE ''
    END AS AlwaysOnStatus,
	case
		when SERVERPROPERTY('IsHadrEnabled') = 1 or SERVERPROPERTY('IsClustered') = 1 then 'Clustered'
		else 'Single'
	end as 'ServerType',
    (SELECT count(*) FROM sys.databases) as Databases
"
# SQL to get SQL DataBase Information on databases
$SQLsmtdb = "
SELECT  SERVERPROPERTY('ServerName') AS InstanceName
,'/ReplaceText/' as 'AgencyName'
,name as 'DBName'
,create_date as 'CreatedDate'
,state_desc as 'CurrentStatus'
,recovery_model_desc as 'RecoveryModel'
FROM sys.databases"
# SQL to get SQL DataBase Information on alerts
$SQLsmtal = "
SELECT  SERVERPROPERTY('ServerName') AS InstanceName
,'/ReplaceText/' as 'AgencyName'
,name as 'Name'
,event_source as 'EventSource'
,message_id as 'MessageID'
,case 
	when Severity = 1   then '1 Misc System Info'
	when Severity =  2  then '2 Reserved'
	when Severity =  3  then '3 Reserved'
	when Severity =  4  then '4 Reserved'
	when Severity =  5  then '5 Reserved'
	when Severity =  6  then '6 Reserved'
	when Severity =  7  then '7 Notification: Status Info'
	when Severity =  8  then '8 Notification: User Intervention Rquired'
	when Severity =  9  then '9 User Defined'
	when Severity =  10  then '10 Information'
	when Severity =  11  then '11 Specified Database Object Not Found'
	when Severity =  12  then '12 Unused'
	when Severity =  13  then '13 User Transation Systax Error'
	when Severity =  14  then '14 Insufficient Permission'
	when Severity =  15  then '15 Syntax Error in SQL Statements'
	when Severity =  16  then '16 Miscellaneous User Error'
	when Severity =  17  then '17 Insufficient Resources'
	when Severity =  18  then '18 Nonfatal Internal Error'
	when Severity =  19  then '19 Fatal Error in Resource'
	when Severity =  20  then '20 Fatal Error in Current Process'
	when Severity =  21  then '21 Fatal Error in Database Process'
	when Severity =  22  then '22 Fatal Error: Table Integrity Suspect'
	when Severity =  23  then '23 Fatal Error: Database Integrity Suspect'
	when Severity =  24  then '24 Fatal Error: Hardware Error'
	when Severity =  25  then '25 Fatal Error'
	else ''
end as 'Severity'
,case
    when enabled = 1 then'Yes'
    else 'No'
end as 'Enabled'
,case 
    when last_occurrence_date != 0 then
	    convert(date,
	    SUBSTRING(convert(char(8),last_occurrence_date),5,2) + '/' 
	    + SUBSTRING(convert(char(8),last_occurrence_date),7,2) 
	    + '/' + SUBSTRING(convert(char(8),last_occurrence_date),1,4)
	    )
	else null
end as 'LastOccurrenceDate'
,case 
    when last_Response_date != 0 then
	    convert(date,
	    SUBSTRING(convert(char(8),last_Response_date),5,2) + '/' 
	    + SUBSTRING(convert(char(8),last_Response_date),7,2) 
	    + '/' + SUBSTRING(convert(char(8),last_Response_date),1,4)
	    )
	else null
end as 'LastResponseDate'
,left(notification_message,255) as 'NotificationMessage'
,isnull(database_name,'') as 'DatabaseName'
,occurrence_count as 'OccurrenceCounter'
,count_reset_date as 'OccurrenceResetCounterDate'
,count_reset_time as 'OccurrenceResetCounterTime' 
,has_notification as 'NumberofNotificationContacts'
FROM msdb.dbo.sysalerts 
order by name
"
$SQLsmtJob = "
select  @@servername as 'InstanceName', 
'/ReplaceText/' as 'AgencyName',
 j.name as 'JobName',
 (select top 1 '""' + sx.command + '""' from msdb.dbo.sysjobsteps sx where j.job_id = sx.job_id order by step_id) as 'FirstStepCommand',
 Case
	when j.enabled = 1 then 'Yes'
	else 'No'
end as 'JobEnabled',
case	
	when j.notify_level_eventlog = 0 then 'Never'
	when j.notify_level_eventlog = 1 then 'On Success'
	when j.notify_level_eventlog = 2 then 'On Failure'
	when j.notify_level_eventlog = 3 then 'Always'
	else ''
end as 'NotifyEventLog',
case	
	when j.notify_level_email = 0 then 'Never'
	when j.notify_level_email = 1 then 'On Success'
	when j.notify_level_email = 2 then 'On Failure'
	when j.notify_level_email = 3 then 'Always'
	else ''
end as 'NotifyEmail',
j.date_created as 'JobCreatedDateTime',
j.date_modified as 'JobModifiedDateTime',
 case
	when h.run_status = 0 then 'Failed'
 	when h.run_status = 1 then 'Successful'
	when h.run_status = 2 then 'Retry'
	when h.run_status = 3 then 'Cancelled'
	when h.run_status = 4 then 'In Progress'
	else '' --should never be this, just making sure there is a default response dra
end as 'Status',
convert(datetime,msdb.dbo.agent_datetime(run_date, run_time)) as 'RunDateTime',
 ((run_duration/10000*3600 + (run_duration/100)%100*60 + run_duration%100 + 31 ) / 60) 
         as 'RunDurationInMinutes' ,
	run_duration as 'RunDurationInSeconds' 
From msdb.dbo.sysjobs j 
left outer JOIN msdb.dbo.sysjobsteps s  ON j.job_id = s.job_id
left outer JOIN msdb.dbo.sysjobhistory h  ON s.job_id = h.job_id 
											 AND s.step_id = h.step_id 
											 AND h.step_id <> 0
where j.enabled = 1 and
        (j.name = 'DatabaseBackup - USER_DATABASES - LOG' 
			  or j.name = 'DatabaseBackup - USER_DATABASES - FULL'
              or j.name = 'DatabaseBackup - SYSTEM_DATABASES - FULL'
              or j.name = 'Backup System Databases.Subplan_1'
              or j.name = 'Backup User Databases - FULL.Subplan_1'
              or j.name = 'CESC-DBA - DatabaseBackup - SYSTEM_DATABASES - FULL'
              or j.name = 'CESC-DBA - DatabaseBackup - USER_DATABASES - FULL'
              or j.name = 'DatabaseBackup - Selected - DATABASES - FULL'
              or j.name = 'DatabaseBackup_System_Database-Full') 
        and h.run_status in (0,1)
order by j.name
"

function Get-CGISQL 
{           
    param ($SQLServer,$DBName,$SQLStmt)            
    Process 
    {
        Invoke-Sqlcmd -ServerInstance $SQLServer -Database $DBName -Query $SQLStmt -DisableVariables
    }
}

function Get-CGISQLInfoSrv
{           
    param ($ServerList)            
    Process{ 
    # Get SQL Server Information
    $TotServers = $ServerList.count
    $IndexServer = 1
    $temp = $serverlist |foreach-object {Write-Progress -Activity ("Getting Server Info: " + $_.ServerName + " " + $IndexServer + " of " + $TotServers)
        ;$IndexServer += 1
        ;IF ($_.Description -eq "SQL Server") {
            ;$sqlsmtsrvx = $sqlsmtsrv -replace "/ReplaceText/",$_.AgencyName 
            ;$x = Get-CGISQL $_.ServerName "master" $sqlsmtsrvx 
            if (($null -eq $x.count) -or  (0 -lt $x.count )) {$x} else {[pscustomobject]@{'InstanceName'=(get-cgiserver $_.ServerName); 'ProductName'="Error:Not Found"; 'AgencyName'=$_.AgencyName}}}
        else {[pscustomobject]@{'InstanceName'=(get-cgiserver $_.ServerName); 'AgencyName'=$_.AgencyName; 'ProductName'="SSRS Microsoft SQL Server "+$_.ProductVersion; 'ProductVersion'=$_.ProductVersion; 'Edition'=$_.Edition; 'SQLInstallDate'=$_.SQLInstallDate; 'SP'='n/a'; 'AlwaysOn'=$_.AlwaysOn;'AlwaysOnStatus'=$_.AlwaysOnStatus; 'ServerType'='Single'; 'Databases'='0'}}}
    $temp
    }  # End Process
}
function Get-CGISQLInfodb
{           
    param ($ServerList)            
    Process{ 
    # Get list of DB from SQL Server
    $TotServers = $ServerList.count
    $IndexServer = 1
    $temp =  $ServerList |foreach-object {Write-Progress -Activity ("Getting DataBase Info: " + $_.ServerName + " " + $IndexServer + " of " + $TotServers)
        ;$IndexServer += 1
        ;IF ($_.Description -eq "SQL Server") {
            ;$sqlsmtdbx = $sqlsmtdb -replace "/ReplaceText/",$_.AgencyName 
            ;$x = Get-CGISQL $_.ServerName "master" $sqlsmtdbx
            ;if ($x.count -gt 0) {$x} else {[pscustomobject]@{'InstanceName'=(get-cgiserver $_.ServerName); 'DBName'="Error:Not Found"}}
        } else {[pscustomobject]@{'InstanceName'=(get-cgiserver $_.ServerName) + " SSRS - No databases on this server "; 'AgencyName'=$_.AgencyName}}}
    $temp
    }  # End Process
}
function Get-CGISQLInfoJobs
{           
    param ($ServerList,
            [datetime]$StartDate,
            [datetime]$EndDate)            
    Process{ 
    # Get list of DB from SQL Server
    $newSQLsmtJob = $SQLsmtJob 
    if (($null -ne $startdate) -and ($null -ne $enddate))
        {$enddate = $enddate.adddays(1);$newSQLsmtJob = $SQLsmtJob -replace "where j.enabled = 1 and","where j.enabled = 1 and  msdb.dbo.agent_datetime(run_date, run_time) between '$StartDate' and '$EndDate' and"}
    #For each server run sql if count is > 0 then write data else write error.
    $ServerList |foreach-object {$TotServers = $ServerList.count;$IndexServer = 1; Write-Progress -Activity ("Getting Job Info: " + $_.ServerName + " " + $IndexServer + " of " + $TotServers)
        IF ($_.Description -eq "SQL Server") {
        $newsqlsmtjobx = $newsqlsmtjob -replace "/ReplaceText/",$_.AgencyName ; $x = Get-CGISQL $_.ServerName "master" $newsqlsmtjobx
        if ($null -ne $x) 
            {
            $x |ForEach-Object { $IndexServer += 1;
                $temp +=@([CGISQLJob]::new($_.InstanceName,$_.AgencyName,$_.JobName,$_.FirstStepCommand,$_.JobEnabled,$_.NotifyEventLog,$_.NotifyEmail,$_.JobCreatedDateTime,$_.JobModifiedDateTime,$_.Status,$_.RunDateTime,$_.RunDurationInMinutes,$_.RunDurationInSeconds))}
            } #End If
        else 
            {
            $Temp +=@([CGISQLJob]::new((get-cgiserver $_.ServerName),$_.AgencyName,"ERROR: No Jobs Found"))
            } #End Else
        } else {$Temp +=@([CGISQLJob]::new((get-cgiserver $_.ServerName),$_.AgencyName,"SSRS - No jobs on this server"))}
        } #End For-each
     return  $temp #Return object
    }  # End Process
}
function Get-CGISQLInfoalerts
{           
    param ($ServerList)            
    Process{ 
    # Get Alerts from each server in $Serverlist
    $temp = $null #erase previous data
    #For each server run sql if count is > 0 then write data else write error.
    $ServerList |foreach-object {$TotServers = $ServerList.count;$IndexServer = 1;Write-Progress -Activity ("Getting Alerts Info: " + $_.ServerName + " " + $IndexServer + " of " + $TotServers); $IndexServer += 1
        IF ($_.Description -eq "SQL Server") {
            ;$sqlsmtalx = $sqlsmtal -replace "/ReplaceText/",$_.AgencyName ;$x = Get-CGISQL $_.ServerName "master" $sqlsmtalx;`
            if ($null -ne $x)  
                {
                $x |ForEach-Object {
                $temp +=@([CGISQLAlert]::new($_.InstanceName,$_.AgencyName,$_.Name,$_.EventSource,$_.MessageID,$_.Severity,$_.Enabled,$_.LastOccurrenceDate,$_.LastResponseDate,$_.NotificationMessage,$_.DatabaseName,$_.OccurrenceCounter,$_.OccurrenceResetCounterDate,$_.OccurrenceResetCounterTime,$_.NumberofNotificationContacts))}
                } #End If
            else 
                {
                $Temp +=@([CGISQLAlert]::new((get-cgiserver $_.ServerName),$_.AgencyName,"ERROR: No Alerts Found"))
                } #End Else
            } else {$Temp +=@([CGISQLAlert]::new((get-cgiserver $_.ServerName),$_.AgencyName,"SSRS - No Alerts on this server"))}  
        } #End For-each
     return  $temp #Return object
    }  # End Process
}
# function Get-CGIServers ($AllTemp)
# {
#     Process
#     {
#     $ServerListTempFix = $null
#     if ($AllTemp -ne "All") {$ServerListTemp = import-csv -header 'ServerName' -Delimiter '|' ".\servers.txt"| Out-GridView -PassThru -Title "Select Servers"}
#     else {$ServerListTemp = import-csv -header 'ServerName' -Delimiter '|' ".\input\servers.txt"}
#     $ServerListTemp|ForEach-Object {$s=Get-CGIServer $_.ServerName;$p=Get-CGIPort $_.ServerName;
#                                     if ((Test-NetConnection -computername $s -port $p -informationlevel quiet) -eq $true)
#                                             {$ServerListTempFix +=@([CGISQLServerShort]::new($s+","+$p))}
#                                     else {write-host "Error Connecting to " $S","$P  -ForegroundColor red -BackgroundColor white}
#     }
#     Write-Output $ServerListTempFix
#     } # End Process
# }
function Get-CGIServers ($AllTemp)
{
    Process
    {
    $ServerListTempFix = $null
    # $ServerListTemp = import-csv -header 'ServerName' -Delimiter '|' ".\input\servers.txt"
    $header = 'Effective Date Changed',
            'CIName',
            'ProductVersion',
            'Operational Status',
            'Category',
            'Agency',
            'TCP (Ports)',
            'Installed Date',
            'Description',
            'InstanceName',
            'ServerName',
            'Edition',
            'No. of Databases',
            'Date End of Normal Support',	
            'Purchased Extended Support (Y/N)',	
            'Date Extended Support Ends',	
            'Agency Notification of Expiration (Contact Name)',	
            'Notes',
            'InstanceNameFor Reports',
            'OS Version'
    
    $ServerListTemp = $null
    $ServerListTemp = import-csv  ".\input\servers.csv" -Header $header | Where-Object {$_.'CIName' -ne ''} | Where-Object {$_.'CIName' -ne "CIName"} | Where-Object {$_.'Operational Status' -ne 'Non-Operational'}
    $ServerListTemp|ForEach-Object {$ServerListTempFix +=@([CGISQLServerShort]::new($_.'InstanceNameFor Reports'+','+$_.'TCP (Ports)',$_.Agency,$_.'Operational Status',$_.Description,$_.ProductVersion,$_.Edition,$_.'Installed Date'))}
    Write-Output $ServerListTempFix
    } # End Process
}
function Get-CGIServer { # Get Server name from string
    param ([String]$Connection)
    process
    {
        $idx = $Connection.IndexOf(",") #find position of , in string
        $idx2 = $Connection.IndexOf("\") #find position of \ in string
        $idx3 = if ($idx2 -eq -1) {$idx} else {$idx2} #choose where to trim string on the \ or ,
        if ($idx3 -le 0) {return} #avoid bad data 
        write-output $Connection.Substring(0,$idx3) #Get only the Server
    }
}
function Get-CGIPort { # Get Port Number from String
    param ([String]$Connection)
    process
    {
        $idx = $Connection.IndexOf(",") + 1 #find position of , in string
        $len = $connection.length - $idx #calc the length for substring
        if ($Connection.Substring($idx,$len -le 0)) {Return} #avoid bad data
        write-output $Connection.Substring($idx,$len) #Return only the value after the ,
    }
}
function Get-CGITestSQL { # Test the server and port from input return object
    param ([String]$Connection)
    process
    {
        if ($null -eq $Connection) {Return} #avoid blank records 
        $ServerName =  (Get-CGIServer $Connection);$PortNum =  (Get-CGIport $Connection) #break string down to server and port
        $SQLOpen = if (Test-NetConnection -computername $ComputerName -port $PortNum -informationlevel quiet) {"Successful"}
                     else {"Unreachable"} # Test port for network connection
        $PSOpen = if (Test-NetConnection -computername $ComputerName -port 445 -informationlevel quiet) {"Successful"}
                     else {"Unreachable"} # Test port for network connection
        $obj = [pscustomobject]@{'ServerName'=$ServerName; 
                                'PortNum'=$PortNum;
                                'DateTimeChecked'=$RunDate;
                                'SQLOpenPort'=$SQLOpen;
                                'PSOpenPort'=$PSOpen
                                } #return an object
        $obj
    }
}
