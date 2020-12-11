#Set master location for server.txt and output files
$whereami = $env:computername
if ($whereami -eq "USC-W-26K6SQ2"){set-location  C:\Users\david.r.allen\Documents\VITA\NewJobs;$LogServer = "localhost"} else {set-location  H:\_SQLTEAMAPPS\_Dev\;$LogServer = "WDB03892"}

Import-Module .\cgisql.psm1

#Get all servers in .\Input\Servers.txt
$CGIServersReport = Get-CGIServers "All" 

<# Get data from Servers #>
# $CGIServersReport # List Jobs
$CGIInfoSvr = Get-CGISQLInfoSrv $CGIServersReport
$CGIInfoDB= Get-CGISQLInfodb  $CGIServersReport 
$CGIInfoAlerts = Get-CGISQLInfoalerts $CGIServersReport
$CGIInfoJobs = CGISQLInfojobs $CGIServersReport 

<# Reporting #>
#Report Totals Totals on 1 don't work because arrays only show 2 entries, I need more code to fix that.
write-host "Total Servers Processed: "  $CGIInfoSvr.count -ForegroundColor red -BackgroundColor white
#Report Average DB for Non Clustered or Clustered DB in set
write-host "Server Type Single Average" -ForegroundColor red -BackgroundColor white
$CGIInfoSVR | Where-Object {$_.ServerType -eq "Single"} |   Measure-Object -Property Databases -Average 
write-host "Server Type Cluster Average" -ForegroundColor red -BackgroundColor white
$CGIInfoSVR | Where-Object {$_.ServerType -eq "Clustered"} |   Measure-Object -Property Databases -Average 
#write-host "Total Databases Processed: "  $CGIInfoDB.count -ForegroundColor red -BackgroundColor white

<# Export lists to CSV #>
$CGIInfoSvr | export-csv -path ".\Output\SQLServersInfo.csv" -NoTypeInformation
$CGIInfoDB | export-csv -path ".\Output\SQLDBInfo.csv" -NoTypeInformation
$CGIInfoAlerts | export-csv  -path ".\Output\SQLAlertsInfo.csv" -NoTypeInformation
$CGIInfoJobs | export-csv  -path ".\Output\SQLJobsInfo.csv" -NoTypeInformation

<# Copy to SQL Server#>
<#Reimport the input file with all columns and insert records into LogDB#>
#$LogServer = "localhost"

$LogDB = "CGISQLInfo"
$LogDate = get-date
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
            'Notes',
            'InstanceNameFor Reports',
            'OS Version'
    import-csv  ".\input\servers.csv" -Header $header | Where-Object {$_.'CIName' -ne ''} | Where-Object {$_.'CIName' -ne "CIName"} | Where-Object {$_.'Operational Status' -ne 'Non-Operational'} | `
    ForEach-Object {$sql = "insert into CGIInstanceSpreadsheet values('" + $logdate + "','" `
                                                                         + $_.'Effective Date Changed'  + "','" `
                                                                         + $_.'CIName'  + "','" `
                                                                         + $_.'ProductVersion'  + "','" `
                                                                         + $_.'Operational Status'  + "','" `
                                                                         + $_.'Category'  + "','" `
                                                                         + $_.'Agency'  + "','" `
                                                                         + $_.'TCP (Ports)'  + "','" `
                                                                         + $_.'Installed Date'  + "','" `
                                                                         + $_.'Description'  + "','" `
                                                                         + $_.'InstanceName'  + "','" `
                                                                         + $_.'ServerName'  + "','" `
                                                                         + $_.'Edition'  + "','" `
                                                                         + $_.'No. of Databases'  + "','" `
                                                                         + $_.'Notes'  + "','" `
                                                                         + $_.'InstanceNameFor Reports'  + "','" `
                                                                         + $_.'OS Version'  + "')" `
                    ;Get-CGISQL $LogServer $LogDB $sql}

$CGIInfoSvr |  ForEach-Object {$sql = "insert into CGIServerInfo values('" + $logdate + "','" `
                    + $_.'InstanceName'  + "','" `
                    + $_.'AgencyName'  + "','" `
                    + $_.'ProductName'  + "','" `
                    + $_.'ProductVersion'  + "','" `
                    + $_.'ProductLevel'  + "','" `
                    + $_.'ProductBuild'  + "','" `
                    + $_.'SP'  + "','" `
                    + $_.'Edition'  + "','" `
                    + $_.'SQLInstallDate'  + "','" `
                    + $_.'AlwaysOn'  + "','" `
                    + $_.'AlwaysOnStatus'  + "','" `
                    + $_.'ServerType'  + "','" `
                    + $_.'Databases'  + "')" 
                    ;Get-CGISQL $LogServer $LogDB $sql}

$CGIInfoDB |  ForEach-Object {$sql = "insert into CGIInfoDB values('" + $logdate + "','" `
                    + $_.'InstanceName'  + "','" `
                    + $_.'AgencyName'  + "','" `
                    + $_.'DBName'  + "','" `
                    + $_.'CreateDate'  + "','" `
                    + $_.'CurrentStatus'  + "','" `
                    + $_.'RecoveryModel'  + "')" 
                    ;Get-CGISQL $LogServer $LogDB $sql}
$CGIInfoAlerts |  ForEach-Object {$sql = "insert into CGIInfoAlerts values('" + $logdate + "','" `
                    + $_.'InstanceName'  + "','" `
                    + $_.'AgencyName'  + "','" `
                    + $_.'Name'  + "','" `
                    + $_.'EventSource'  + "','" `
                    + $_.'MessageID'  + "','" `
                    + $_.'Severity'  + "','" `
                    + $_.'Enabled'  + "','" `
                    + $_.'LastOccurrenceDate'  + "','" `
                    + $_.'LastresponseDate'  + "','" `
                    + $_.'NotificationMessage'  + "','" `
                    + $_.'DatabaseName'  + "','" `
                    + $_.'OccurrenceCounter'  + "','" `
                    + $_.'OccurrenceresetCounterDate'  + "','" `
                    + $_.'OccurrenceResetCounterTime'  + "','" `
                    + $_.'NumberofNotificationsContacts'  + "')" 
                    ;Get-CGISQL $LogServer $LogDB $sql}
$CGIInfoJobs |  ForEach-Object {$sql = "insert into CGIInfoJobs values('" + $logdate + "','" `
                    + $_.'InstanceName'  + "','" `
                    + $_.'AgencyName'  + "','" `
                    + $_.'JobName'  + "','" `
                    + $_.'FirstStepCommand'  + "','" `
                    + $_.'JobEnabled'  + "','" `
                    + $_.'NotifyEventLog'  + "','" `
                    + $_.'NotifyEmail'  + "','" `
                    + $_.'JobCreatedDateTime'  + "','" `
                    + $_.'JobModifiedDateTime'  + "','" `
                    + $_.'Status'  + "','" `
                    + $_.'RunDateTime'  + "','" `
                    + $_.'RunDurationInMinutes'  + "','" `
                    + $_.'RunDurationInSeconds'  + "')" 
                    ;Get-CGISQL $LogServer $LogDB $sql}

#Export Servers for Exec report.
$CGIInfoSvr | Select-Object InstanceName,AgencyName,ProductName,SP,Edition,SQLInstallDate,AlwaysOn,ServerType,Databases `
    | export-csv -path ".\Output\Unisys - SQLServersInfo.csv" -NoTypeInformation  
$CGIInfoSvr     | ForEach-Object {$sql = "insert into UnisysServerInfo values('" + $logdate + "','" `
                    + $_.'InstanceName'  + "','" `
                    + $_.'AgencyName'  + "','" `
                    + $_.'ProductName'  + "','" `
                    + $_.'SP'  + "','" `
                    + $_.'Edition'  + "','" `
                    + $_.'SQLInstallDate'  + "','" `
                    + $_.'AlwaysOn'  + "','" `
                    + $_.'ServerType'  + "','" `
                    + $_.'Databases'  + "')" 
                    ;Get-CGISQL $LogServer $LogDB $sql}

#Alter Alert Name for Exec report
class CGISQLAlert2 {
    [string]$InstanceName 
    [string]$AgencyName
    [string]$Name
    [string]$Enabled; `
`
    CGISQLAlert2(
        [string]$InstanceName, 
        [string]$AgencyName,
        [string]$Name,
        [string]$Enabled
    ){
        $this.InstanceName = $InstanceName
        $this.AgencyName = $AgencyName
        $this.Name = $Name
        $this.Enabled = $Enabled
    }
}

$CGIInfoAlerts | foreach-Object {
    switch ($_.name) {
        "825Error-TransientCorruption"  {$NewAlert = "01. 825Error-TransientCorruption - Soft I/O error, has occurred but SQL Server was able to retry and read the data." }
        "Sev17Alert"                    {$NewAlert = "02. Sev17Alert Resource alert check the server logs." }
        "Sev19Alert"                    {$NewAlert = "03. Sev19Alert SQL Server Error in Resource " }
        "Sev20Alert"                    {$NewAlert = "04. Sev20Alert No Kerberos SPN SQL authentication defaults to NTLM." }
        "Sev21Alert"                    {$NewAlert = "05. Sev21Alert SQL Server Fatal Error in Database ID Process"   }
        "Sev22Alert"                    {$NewAlert = "06. Sev22Alert SQL Server Fatal Error Table Integrity Suspect" }
        "Sev23Alert"                    {$NewAlert = "07. Sev23Alert SQL Server Fatal Error: Database Integrity Suspect"   }
        "Sev24Alert"                    {$NewAlert = "08. Sev24Alert Hardware Error"  }
        "Sev25Alert"                    {$NewAlert = "09. Sev25Alert SQL send an internal message."  }
        "SQLMemPageOut"                 {$NewAlert = "10. SQLMemPageOut A significant part of SQL server process memory has been paged out." }
        "SlowIORequests"                {$NewAlert = "11. SlowIORequests I/O requests taking longer than 15 seconds." }
        "Insufficient Permissions"      {$NewAlert = "12. Insufficient Permissions Account does not have permissions." }
        "SSRS - No Alerts on this server"   {$NewAlert = "SSRS - No Alerts on this server"}
        "ERROR: No Alerts Found"        {$NewAlert = "ERROR: No Alerts Found"}
        Default                         {$NewAlert = "Skip"}
        }
 ;$CGIInfoAlerts2 +=@([CGISQLAlert2]::new($_.instancename,$_.AgencyName,$NewAlert,$_.Enabled))    
}
#Export Alerts for Exec report
$CGIInfoAlerts2 | where-object {$_.name -ne "Skip"}  | sort-object InstanceName,AgencyName,Name,Enabled | Select-Object InstanceName,AgencyName,Name,Enabled | 
    export-csv  -path ".\Output\Unisys - SQLAlertsInfo.csv" -NoTypeInformation
$CGIInfoAlerts2  |  ForEach-Object {$sql = "insert into UnisysInfoAlerts values('" + $logdate + "','" `
                + $_.'InstanceName'  + "','" `
                + $_.'AgencyName'  + "','" `
                + $_.'Name'  + "','" `
                + $_.'Enabled'  + "')" 
                ;Get-CGISQL $LogServer $LogDB $sql}

#Export Jobs to Exec report, needs to be 3 exports to get data in the required format.
$UnisysInfoJobs = $CGIInfoJobs | where-object {$_.RunDateTime.todatetime($x) -ge (get-date).date.adddays(-1)} | 
    Select-Object InstanceName,AgencyName,JobName,JobEnabled,JobCreatedDateTime,Status,RunDateTime,RunDurationInMinutes,RunDurationInSeconds 

$UnisysInfoJobs += $CGIInfoJobs | where-object {$_.JobName -eq "SSRS - No jobs on this server"} | 
    Select-Object InstanceName,AgencyName,JobName,JobEnabled,JobCreatedDateTime,Status,RunDateTime,RunDurationInMinutes,RunDurationInSeconds   

$UnisysInfoJobs += $CGIInfoJobs | where-object  {$_.JobName -eq "ERROR: No Jobs Found"} | 
    Select-Object InstanceName,AgencyName,JobName,JobEnabled,JobCreatedDateTime,Status,RunDateTime,RunDurationInMinutes,RunDurationInSeconds  

$UnisysInfoJobs |  export-csv  -path ".\Output\Unisys - SQLJobsInfo.csv" -NoTypeInformation 

$UnisysInfoJobs |  ForEach-Object {$sql = "insert into UnisysInfoJobs values('" + $logdate + "','" `
                    + $_.'InstanceName'  + "','" `
                    + $_.'AgencyName'  + "','" `
                    + $_.'JobName'  + "','" `
                    + $_.'JobEnabled'  + "','" `
                    + $_.'JobCreatedDateTime'  + "','" `
                    + $_.'Status'  + "','" `
                    + $_.'RunDateTime'  + "','" `
                    + $_.'RunDurationInMinutes'  + "','" `
                    + $_.'RunDurationInSeconds'  + "')" 
                    ;Get-CGISQL $LogServer $LogDB $sql}

write-host "Data Exported to .CSV" -ForegroundColor red -BackgroundColor white
