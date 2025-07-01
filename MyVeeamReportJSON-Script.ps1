<#====================================================================
Author        : Tiago DA SILVA - ATHEO INGENIERIE
Version       : 1.0.1
Creation Date : 2025-07-01
Last Update   : 2025-07-01
GitHub Repo   : https://github.com/TiagoDSLV/MyVeeamMonitoring
====================================================================

DESCRIPTION:
My Veeam Report is a flexible reporting script for Veeam Backup and
Replication. This report can be customized to report on Backup, Replication,
Backup Copy, Tape Backup, SureBackup and Agent Backup jobs as well as
infrastructure details like repositories, proxies and license status. 

====================================================================#>

#Region Update Script
param (
    [Parameter(Mandatory = $true)]
    [string]$ConfigFileName
)

# Load Configuration
$ConfigPath = Join-Path -Path $PSScriptRoot -ChildPath $ConfigFileName
if (Test-Path $ConfigPath) {
    . $ConfigPath
} else {
    Write-Warning "Config file '$ConfigPath' not found."
    exit 1
}

#Region Update Script
function Get-VersionFromScript {
  param ([string]$Content)
  if ($Content -match "Version\s*:\s*([\d\.]+)") {
      return $matches[1]  # Return the version string if found
  }
  return $null  # Return null if no version is found
}

$OutputPath = ".\MyVeeamReportJSON-Script.ps1"
$FileURL = "https://raw.githubusercontent.com/TiagoDSLV/MyVeeamReportJSON/refs/heads/main/MyVeeamReportJSON-Script.ps1"

# Lire le contenu local et la version
$localScriptContent = Get-Content -Path $OutputPath -Raw
$localVersion = Get-VersionFromScript -Content $localScriptContent

# Initialiser $remoteScriptContent et $remoteVersion
$remoteScriptContent = $null
$remoteVersion = $null

try {
    # Essayer de récupérer le script distant
    $remoteScriptContent = Invoke-RestMethod -Uri $FileURL -UseBasicParsing
    $remoteVersion = Get-VersionFromScript -Content $remoteScriptContent
} catch {
    Write-Warning "Failed to retrieve remote script content: $_"
    # Optionnel : continuer avec l'ancien script sans mise à jour
    return
}

if ($localVersion -ne $remoteVersion) {
    try {
        $remoteScriptContent | Set-Content -Path $OutputPath -Encoding UTF8 -Force
        Write-Host "Script updated."
        Write-Host "Restarting script..."
        Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$OutputPath`""
        exit
    } catch {
        Write-Warning "Update error: $_"
    }
} else {
    Write-Host "Script is up to date."
}

#enregion

# Set ReportVersion
$localScriptContent = Get-Content -Path $OutputPath -Raw  # Read local file content
$localVersion = Get-VersionFromScript -Content $localScriptContent  # Extract local version
$reportVersion = $localVersion

#Region Connect
# Connect to VBR server
$OpenConnection = (Get-VBRServerSession).Server
If ($OpenConnection -ne $vbrServer){
Disconnect-VBRServer
Try {
Connect-VBRServer -server $vbrServer -ErrorAction Stop
} Catch {
Write-Host "Unable to connect to VBR server - $vbrServer" -ForegroundColor Red
exit
}
}
#endregion

#region NonUser-Variables
# Get all Backup/Backup Copy/Replica Jobs
$allJobs = @()
If ($showSummaryBk + $showJobsBk + $showFileJobsBk + $showAllSessBk + $showAllTasksBk + $showRunningBk +
$showRunningTasksBk + $showWarnFailBk + $showTaskWFBk + $showSuccessBk + $showTaskSuccessBk +
$showSummaryRp + $showJobsRp + $showAllSessRp + $showAllTasksRp + $showRunningRp +
$showRunningTasksRp + $showWarnFailRp + $showTaskWFRp + $showSuccessRp + $showTaskSuccessRp +
$showSummaryBc + $showJobsBc + $showAllSessBc + $showAllTasksBc + $showIdleBc +
$showPendingTasksBc + $showRunningBc + $showRunningTasksBc + $showWarnFailBc +
$showTaskWFBc + $showSuccessBc + $showTaskSuccessBc) {
$allJobs = Get-VBRJob -WarningAction SilentlyContinue
}

#Other version where FileBackup is just added to normal backup job sessions.
#$allJobsBk = @($allJobs | Where-Object {$_.JobType -eq "Backup" -or $_.JobType -eq"NasBackup" })
# Get all Backup Jobs
$allJobsBk = @($allJobs | Where-Object {$_.JobType -eq "Backup"})
# Get all File Backup Jobs
$allFileJobsBk = @($allJobs | Where-Object {$_.JobType -eq "NasBackup"})
# Get all Replication Jobs
$allJobsRp = @($allJobs | Where-Object {$_.JobType -eq "Replica"})
# Get all Backup Copy Jobs
$allJobsBc = @($allJobs | Where-Object {$_.JobType -eq "BackupSync" -or $_.JobType -eq "SimpleBackupCopyPolicy"})
# Get all Tape Jobs
$allJobsTp = @()
If ($showSummaryTp + $showJobsTp + $showAllSessTp + $showAllTasksTp +
$showWaitingTp + $showIdleTp + $showPendingTasksTp + $showRunningTp + $showRunningTasksTp +
$showWarnFailTp + $showTaskWFTp + $showSuccessTp + $showTaskSuccessTp) {
$allJobsTp = @(Get-VBRTapeJob)
}
# Get all Agent Backup Jobs
$allJobsEp = @()
If ($showSummaryEp + $showJobsEp + $showAllSessEp + $showRunningEp +
$showWarnFailEp + $showSuccessEp) {
$allJobsEp = @(Get-VBRComputerBackupJob)
}
# Get all SureBackup Jobs
$allJobsSb = @()
If ($showSummarySb + $showJobsSb + $showAllSessSb + $showAllTasksSb +
$showRunningSb + $showRunningTasksSb + $showWarnFailSb + $showTaskWFSb +
$showSuccessSb + $showTaskSuccessSb) {
$allJobsSb = @(Get-VBRSureBackupJob)
}

# Get all Backup/Backup Copy/Replica Sessions
$allSess = @()
If ($allJobs) {
$allSess = Get-VBRBackupSession
}
# Get all File / NAS Backup Sessions
$allFileSess = @()
If ($allFileJobs) {
$allFileSess = Get-VBRNASBackupSession -Name *
}

# Get all Tape Backup Sessions
$allSessTp = @()
If ($allJobsTp) {
Foreach ($tpJob in $allJobsTp){
$tpSessions = [veeam.backup.core.cbackupsession]::GetByJob($tpJob.id)
$allSessTp += $tpSessions
}
}
# Get all Agent Backup Sessions
$allSessEp = @()
If ($allJobsEp) {
$allSessEp = Get-VBRComputerBackupJobSession
}
# Get all SureBackup Sessions
$allSessSb = @()
If ($allJobsSb) {
$allSessSb = Get-VBRSureBackupSession
}

# Get all Backups
$jobBackups = @()
If ($showBackupSizeBk + $showBackupSizeBc + $showBackupSizeEp) {
$jobBackups = Get-VBRBackup
}
# Get Backup Job Backups
$backupsBk = @($jobBackups | Where-Object { $_.JobType -in @("Backup", "PerVmParentBackup") })
# Get Backup Copy Job Backups
$backupsBc = @($jobBackups | Where-Object { $_.JobType -in @("BackupSync", "SimpleBackupCopyPolicy") })
# Get Agent Backup Job Backups
$backupsEp = @($jobBackups | Where-Object {$_.JobType -eq "EndpointBackup" -or $_.JobType -eq "EpAgentBackup" -or $_.JobType -eq "EpAgentPolicy"})

# Get all Media Pools
$mediaPools = Get-VBRTapeMediaPool
# Get all Media Vaults
Try {
$mediaVaults = Get-VBRTapeVault
} Catch {
Write-Host "Tape possibly not licensed."
}
# Get all Tapes
$mediaTapes = Get-VBRTapeMedium
# Get all Tape Libraries
$mediaLibs = Get-VBRTapeLibrary
# Get all Tape Drives
$mediaDrives = Get-VBRTapeDrive

# Get Configuration Backup Info
$configBackup = Get-VBRConfigurationBackupJob
# Get all Proxies
$proxyList = Get-VBRViProxy
# Get all Repositories
$repoList = Get-VBRBackupRepository | Where-Object { $_.Name -notin $excludedRepositories }
$repoListSo = Get-VBRBackupRepository -ScaleOut | Where-Object { $_.Name -notin $excludedRepositories }

# Convert mode (timeframe) to hours
If ($reportMode -eq "Monthly") {
$HourstoCheck = 720
} Elseif ($reportMode -eq "Weekly") {
$HourstoCheck = 168
} Else {
$HourstoCheck = $reportMode
}

# Gather all VMs in VBRViEntity 
$allVMsVBRVi = Find-VBRViEntity | Where-Object { $_.Type -eq "Vm" }

# Gather all Backup Sessions within timeframe
$sessListBk = @($allSess | Where-Object {($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working") -and $_.JobType -eq "Backup"})
If ($null -ne $backupJob -and $backupJob -ne "") {
$allJobsBkTmp = @()
$sessListBkTmp = @()
$backupsBkTmp = @()
Foreach ($bkJob in $backupJob) {
$allJobsBkTmp += $allJobsBk | Where-Object {$_.Name -like $bkJob}
$sessListBkTmp += $sessListBk | Where-Object {$_.JobName -like $bkJob}
$backupsBkTmp += $backupsBk | Where-Object {$_.JobName -like $bkJob}
}
$allJobsBk = $allJobsBkTmp | Sort-Object Id -Unique
$sessListBk = $sessListBkTmp | Sort-Object Id -Unique
$backupsBk = $backupsBkTmp | Sort-Object Id -Unique
}
If ($onlyLastBk) {
$tempSessListBk = $sessListBk
$sessListBk = @()
Foreach($job in $allJobsBk) {
$sessListBk += $tempSessListBk | Where-Object {$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
}
}
# Get Backup Session information
$totalXferBk = 0
$totalReadBk = 0

$sessListBk | ForEach-Object {$totalXferBk += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListBk | ForEach-Object {$totalReadBk += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$successSessionsBk = @($sessListBk | Where-Object {$_.Result -eq "Success"})
$warningSessionsBk = @($sessListBk | Where-Object {$_.Result -eq "Warning"})
$failsSessionsBk = @($sessListBk | Where-Object {$_.Result -eq "Failed"})
$runningSessionsBk = @($sessListBk | Where-Object {$_.State -eq "Working"})
$failedSessionsBk = @($sessListBk | Where-Object {($_.Result -eq "Failed") -and ($_.WillBeRetried -ne "True")})

# File Backup Session Section Start

$fileSessListBk = @($allFileSess | Where-Object {($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working") -and $_.JobType -eq "Backup"})
If ($null -ne $backupJob -and $backupJob -ne "") {
$allFileJobsBkTmp = @()
$fileSessListBkTmp = @()
$fileBackupsBkTmp = @()
Foreach ($bkJob in $backupJob) {
$allFileJobsBkTmp += $allFileJobsBk | Where-Object {$_.Name -like $bkJob}
$fileSessListBkTmp += $fileSessListBk | Where-Object {$_.JobName -like $bkJob}
$fileBackupsBkTmp += $fileBackupsBk | Where-Object {$_.JobName -like $bkJob}
}
$allFileJobsBk = $allFileJobsBkTmp | Sort-Object Id -Unique
$fileSessListBk = $fileSessListBkTmp | Sort-Object Id -Unique
$fileBackupsBk = $fileBackupsBkTmp | Sort-Object Id -Unique
}
If ($onlyLastBk) {
$tempFileSessListBk = $fileSessListBk
$fileSessListBk = @()
Foreach($job in $allFileJobsBk) {
$fileSessListBk += $tempFileSessListBk | Where-Object {$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
}
}
# Get Backup Session information
$totalXferFileBk = 0
$totalReadFileBk = 0

$fileSessListBk | ForEach-Object {$totalXferFileBk += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$fileSessListBk | ForEach-Object {$totalReadFileBk += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
# End File Backup Session Section End

# Gather all Replication Sessions within timeframe
$sessListRp = @($allSess | Where-Object {($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working") -and $_.JobType -eq "Replica"})
If ($null -ne $replicaJob -and $replicaJob -ne "") {
$allJobsRpTmp = @()
$sessListRpTmp = @()
Foreach ($rpJob in $replicaJob) {
$allJobsRpTmp += $allJobsRp | Where-Object {$_.Name -like $rpJob}
$sessListRpTmp += $sessListRp | Where-Object {$_.JobName -like $rpJob}
}
$allJobsRp = $allJobsRpTmp | Sort-Object Id -Unique
$sessListRp = $sessListRpTmp | Sort-Object Id -Unique
}
If ($onlyLastRp) {
$tempSessListRp = $sessListRp
$sessListRp = @()
Foreach($job in $allJobsRp) {
$sessListRp += $tempSessListRp | Where-Object {$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
}
}
# Get Replication Session information
$totalXferRp = 0
$totalReadRp = 0
$sessListRp | ForEach-Object {$totalXferRp += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListRp | ForEach-Object {$totalReadRp += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$successSessionsRp = @($sessListRp | Where-Object {$_.Result -eq "Success"})
$warningSessionsRp = @($sessListRp | Where-Object {$_.Result -eq "Warning"})
$failsSessionsRp = @($sessListRp | Where-Object {$_.Result -eq "Failed"})
$runningSessionsRp = @($sessListRp | Where-Object {$_.State -eq "Working"})
$failedSessionsRp = @($sessListRp | Where-Object {($_.Result -eq "Failed") -and ($_.WillBeRetried -ne "True")})

# Gather all Backup Copy Sessions within timeframe
$sessListBc = @($allSess | Where-Object {($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -match "Working|Idle") -and ($_.JobType -eq "BackupSync" -or $_.JobType -eq "SimpleBackupCopyWorker")})
If ($null -ne $bcopyJob -and $bcopyJob -ne "") {
$allJobsBcTmp = @()
$sessListBcTmp = @()
$backupsBcTmp = @()
Foreach ($bcJob in $bcopyJob) {
$allJobsBcTmp += $allJobsBc | Where-Object {$_.'Job Name'-like $bcJob}
$sessListBcTmp += $sessListBc | Where-Object {$_.'Job Name' -like $bcJob}
$backupsBcTmp += $backupsBc | Where-Object {$_.'Job Name' -like $bcJob}
}
$allJobsBc = $allJobsBcTmp | Sort-Object Id -Unique
$sessListBc = $sessListBcTmp | Sort-Object Id -Unique
$backupsBc = $backupsBcTmp | Sort-Object Id -Unique
}
If ($onlyLastBc) {
$tempSessListBc = $sessListBc
$sessListBc = @()
Foreach($job in $allJobsBc) {
$sessListBc += $tempSessListBc | Where-Object {($_.JobName -split '\\')[0] -eq $job.Name -and $_.BaseProgress -eq 100}
}
}
# Get Backup Copy Session information
$totalXferBc = 0
$totalReadBc = 0
$sessListBc | ForEach-Object {$totalXferBc += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListBc | ForEach-Object {$totalReadBc += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$idleSessionsBc = @($sessListBc | Where-Object {$_.State -eq "Idle"})
$successSessionsBc = @($sessListBc | Where-Object {$_.Result -eq "Success"})
$warningSessionsBc = @($sessListBc | Where-Object {$_.Result -eq "Warning"})
$failsSessionsBc = @($sessListBc | Where-Object {$_.Result -eq "Failed"})
$workingSessionsBc = @($sessListBc | Where-Object {$_.State -eq "Working"})

# Gather all Tape Backup Sessions within timeframe
$sessListTp = @($allSessTp | Where-Object {$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -match "Working|Idle"})
If ($null -ne $tapeJob -and $tapeJob -ne "") {
$allJobsTpTmp = @()
$sessListTpTmp = @()
Foreach ($tpJob in $tapeJob) {
$allJobsTpTmp += $allJobsTp | Where-Object {$_.Name -like $tpJob}
$sessListTpTmp += $sessListTp | Where-Object {$_.JobName -like $tpJob}
}
$allJobsTp = $allJobsTpTmp | Sort-Object Id -Unique
$sessListTp = $sessListTpTmp | Sort-Object Id -Unique
}
If ($onlyLastTp) {
$tempSessListTp = $sessListTp
$sessListTp = @()
Foreach($job in $allJobsTp) {
$sessListTp += $tempSessListTp | Where-Object {$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
}
}
# Get Tape Backup Session information
$totalXferTp = 0
$totalReadTp = 0
$sessListTp | ForEach-Object {$totalXferTp += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListTp | ForEach-Object {$totalReadTp += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$idleSessionsTp = @($sessListTp | Where-Object {$_.State -eq "Idle"})
$successSessionsTp = @($sessListTp | Where-Object {$_.Result -eq "Success"})
$warningSessionsTp = @($sessListTp | Where-Object {$_.Result -eq "Warning"})
$failsSessionsTp = @($sessListTp | Where-Object {$_.Result -eq "Failed"})
$workingSessionsTp = @($sessListTp | Where-Object {$_.State -eq "Working"})
$waitingSessionsTp = @($sessListTp | Where-Object {$_.State -eq "WaitingTape"})

# Gather all Agent Backup Sessions within timeframe
$sessListEp = $allSessEp | Where-Object {($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working")}
If ($null -ne $epbJob -and $epbJob -ne "") {
$allJobsEpTmp = @()
$sessListEpTmp = @()
$backupsEpTmp = @()
Foreach ($eJob in $epbJob) {
$allJobsEpTmp += $allJobsEp | Where-Object {$_.Name -like $eJob}
$backupsEpTmp += $backupsEp | Where-Object {$_.JobName -like $eJob}
}
Foreach ($job in $allJobsEpTmp) {
$sessListEpTmp += $sessListEp | Where-Object {$_.JobId -eq $job.Id}
}
$allJobsEp = $allJobsEpTmp | Sort-Object Id -Unique
$sessListEp = $sessListEpTmp | Sort-Object Id -Unique
$backupsEp = $backupsEpTmp | Sort-Object Id -Unique
}
If ($onlyLastEp) {
$tempSessListEp = $sessListEp
$sessListEp = @()
Foreach($job in $allJobsEp) {
$sessListEp += $tempSessListEp | Where-Object {$_.JobId -eq $job.Id} | Sort-Object EndTime -Descending | Select-Object -First 1
}
}
# Get Agent Backup Session information
$successSessionsEp = @($sessListEp | Where-Object {$_.Result -eq "Success"})
$warningSessionsEp = @($sessListEp | Where-Object {$_.Result -eq "Warning"})
$failsSessionsEp = @($sessListEp | Where-Object {$_.Result -eq "Failed"})
$runningSessionsEp = @($sessListEp | Where-Object {$_.State -eq "Working"})

# Gather all SureBackup Sessions within timeframe
$sessListSb = @($allSessSb | Where-Object {$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -ne "Stopped"})
If ($null -ne $surebJob -and $surebJob -ne "") {
$allJobsSbTmp = @()
$sessListSbTmp = @()
Foreach ($SbJob in $surebJob) {
$allJobsSbTmp += $allJobsSb | Where-Object {$_.Name -like $SbJob}
$sessListSbTmp += $sessListSb | Where-Object {$_.JobName -like $SbJob}
}
$allJobsSb = $allJobsSbTmp | Sort-Object Id -Unique
$sessListSb = $sessListSbTmp | Sort-Object Id -Unique
}
If ($onlyLastSb) {
$tempSessListSb = $sessListSb
$sessListSb = @()
Foreach($job in $allJobsSb) {
$sessListSb += $tempSessListSb | Where-Object {$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
}
}
# Get SureBackup Session information
$successSessionsSb = @($sessListSb | Where-Object {$_.Result -eq "Success"})
$warningSessionsSb = @($sessListSb | Where-Object {$_.Result -eq "Warning"})
$failsSessionsSb = @($sessListSb | Where-Object {$_.Result -eq "Failed"})
$runningSessionsSb = @($sessListSb | Where-Object {$_.State -ne "Stopped"})


# Append Report Mode to Email subject
If ($modeSubject) {
If (($reportMode -ne "Weekly") -And ($reportMode -ne "Monthly")) {
$emailSubject = "$emailSubject (Last $reportMode Hrs)"
} Else {
$emailSubject = "$emailSubject ($reportMode)"
}
}

# Append VBR Server to Email subject
If ($vbrSubject) {
$emailSubject = "$emailSubject - $vbrServer"
}

# Append Date and Time to Email subject
If ($dtSubject) {
$emailSubject = "$emailSubject - $(Get-Date -format g)"
}
#endregion

#region Functions

Function Get-VBRProxyInfo {
[CmdletBinding()]
param (
[Parameter(Position=0, ValueFromPipeline=$true)]
[PSObject[]]$Proxy
)
Begin {
$outputAry = @()
Function Build-Object {param ([PsObject]$inputObj)
  $ping = New-Object system.net.networkinformation.ping
  $isIP = '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b'
  If ($inputObj.Host.Name -match $isIP) {
    $IPv4 = $inputObj.Host.Name
  } Else {
    $DNS = [Net.DNS]::GetHostEntry("$($inputObj.Host.Name)")
    $IPv4 = ($DNS.get_AddressList() | Where-Object {$_.AddressFamily -eq "InterNetwork"} | Select-Object -First 1).IPAddressToString
  }
  $pinginfo = $ping.send("$($IPv4)")
  If ($pinginfo.Status -eq "Success") {
    $hostAlive = "Success"
    $response = $pinginfo.RoundtripTime
  } Else {
    $hostAlive = "Failed"
    $response = $null
  }
  If ($inputObj.IsDisabled) {
    $enabled = "False"
  } Else {
    $enabled = "True"
  }
  $tMode = switch ($inputObj.Options.TransportMode) {
    "Auto" {"Automatic"}
    "San" {"Direct SAN"}
    "HotAdd" {"Hot Add"}
    "Nbd" {"Network"}
    default {"Unknown"}
  }
  $vPCFuncObject = New-Object PSObject -Property @{
    ProxyName = $inputObj.Name
    RealName = $inputObj.Host.Name.ToLower()
    Disabled = $inputObj.IsDisabled
    pType = $inputObj.ChassisType
    Status  = $hostAlive
    IP = $IPv4
    Response = $response
    Enabled = $enabled
    maxtasks = $inputObj.Options.MaxTasksCount
    tMode = $tMode
  }
  Return $vPCFuncObject
}
}
Process {
Foreach ($p in $Proxy) {
  $outputObj = Build-Object $p
}
$outputAry += $outputObj
}
End {
$outputAry
}
}

Function Get-VBRRepoInfo {
[CmdletBinding()]
param (
[Parameter(Position=0, ValueFromPipeline=$true)]
[PSObject[]]$Repository
)
Begin {
$outputAry = @()
Function Build-Object {param($name, $repohost, $path, $free, $total, $maxtasks, $rtype)
  $repoObj = New-Object -TypeName PSObject -Property @{
    Target = $name
    RepoHost = $repohost
    Storepath = $path
    StorageFree = [Math]::Round([Decimal]$free/1GB,2)
    StorageTotal = [Math]::Round([Decimal]$total/1GB,2)
    FreePercentage = [Math]::Round(($free/$total)*100)
    StorageBackup = [Math]::Round([Decimal]$rBackupsize/1GB,2)
    StorageOther = [Math]::Round([Decimal]($total-$rBackupsize-$free)/1GB-0.5,2)
    MaxTasks = $maxtasks
    rType = $rtype
  }
  Return $repoObj
}
}
Process {
Foreach ($r in $Repository) {
  # Refresh Repository Size Info
  [Veeam.Backup.Core.CBackupRepositoryEx]::SyncSpaceInfoToDb($r, $true)
  $rBackupSize = [Veeam.Backup.Core.CBackupRepository]::GetRepositoryBackupsSize($r.Id.Guid)
  $rType = switch ($r.Type) {
    "WinLocal" {"Windows Local"}
    "LinuxLocal" {"Linux Local"}
    "LinuxHardened" {"Hardened"}
    "CifsShare" {"CIFS Share"}
    "AzureStorage"{"Azure Storage"}
    "DataDomain" {"Data Domain"}
    "ExaGrid" {"ExaGrid"}
    "HPStoreOnce" {"HP StoreOnce"}
    "Nfs" {"NFS Direct"}
    default {"Unknown"}
  }
  $outputObj = Build-Object $r.Name $($r.GetHost()).Name.ToLower() $r.Path $r.GetContainer().CachedFreeSpace.InBytes $r.GetContainer().CachedTotalSpace.InBytes $r.Options.MaxTaskCount $rType
}
$outputAry += $outputObj
}
End {
$outputAry
}
}

Function Get-VBRSORepoInfo {
[CmdletBinding()]
param (
[Parameter(Position=0, ValueFromPipeline=$true)]
[PSObject[]]$Repository
)
Begin {
$outputAry = @()
Function Build-Object {param($name, $rname, $repohost, $path, $free, $total, $maxtasks, $rtype, $capenabled)
  $repoObj = New-Object -TypeName PSObject -Property @{
    SoTarget = $name
    Target = $rname
    RepoHost = $repohost
    Storepath = $path
    StorageFree = [Math]::Round([Decimal]$free/1GB,2)
    StorageTotal = [Math]::Round([Decimal]$total/1GB,2)
    FreePercentage = [Math]::Round(($free/$total)*100)
    MaxTasks = $maxtasks
    rType = $rtype
    capEnabled = $capenabled
  }
  Return $repoObj
}
}
Process {
Foreach ($rs in $Repository) {
  ForEach ($rp in $rs.Extent) {
    $r = $rp.Repository
    # Refresh Repository Size Info
    [Veeam.Backup.Core.CBackupRepositoryEx]::SyncSpaceInfoToDb($r, $true)
$rBackupSize = [Veeam.Backup.Core.CBackupRepository]::GetRepositoryBackupsSize($r.Id.Guid)
    $rType = switch ($r.Type) {
      "WinLocal" {"Windows Local"}
      "LinuxLocal" {"Linux Local"}
      "LinuxHardened" {"Hardened"}
      "CifsShare" {"CIFS Share"}
      "AzureStorage"{"Azure Storage"}
      "DataDomain" {"Data Domain"}
      "ExaGrid" {"ExaGrid"}
      "HPStoreOnce" {"HPE StoreOnce"}
      "Nfs" {"NFS Direct"}
      "SanSnapshotOnly" {"SAN Snapshot"}
      "Cloud" {"VCSP Cloud"}
      default {"Unknown"}
    }
if ($rtype -eq "SAN Snapshot" -or $rtype -eq "VCSP Cloud") {$maxTaskCount="N/A"}
else {$maxTaskCount=$r.Options.MaxTaskCount}
    $outputObj = Build-Object $rs.Name $r.Name $($r.GetHost()).Name.ToLower() $r.Path $r.GetContainer().CachedFreeSpace.InBytes $r.GetContainer().CachedTotalSpace.InBytes $maxTaskCount $rType $rBackupSize
    $outputAry += $outputObj
  }
<# #Added for capacity tier begin ToDo
if($rs.CapacityExtent.Repository.Name.Length -gt 0) {
    $ce = $rs.CapacityExtent
    $outputObj = Build-Object $rs.Name $ce.Repository.Name $ce.Repository.ServicePoint $ce.Repository.AmazonS3Folder
    $outputAry += $outputObj
}
#Added for capacity tier end #>
}
}
End {
$outputAry
}
}

Function Get-VBRReplicaTarget {
[CmdletBinding()]
param(
[Parameter(ValueFromPipeline=$true)]
[PSObject[]]$InputObj
)
BEGIN {
$outputAry = @()
$dsAry = @()
If (($null -ne $Name) -and ($null -ne $InputObj)) {
  $InputObj = Get-VBRJob -Name $Name
}
}
PROCESS {
Foreach ($obj in $InputObj) {
  If (($dsAry -contains $obj.ViReplicaTargetOptions.DatastoreName) -eq $false) {
    $esxi = $obj.GetTargetHost()
    $dtstr =  $esxi | Find-VBRViDatastore -Name $obj.ViReplicaTargetOptions.DatastoreName
    $objoutput = New-Object -TypeName PSObject -Property @{
      Target = $esxi.Name
      Datastore = $obj.ViReplicaTargetOptions.DatastoreName
      StorageFree = [Math]::Round([Decimal]$dtstr.FreeSpace/1GB,2)
      StorageTotal = [Math]::Round([Decimal]$dtstr.Capacity/1GB,2)
      FreePercentage = [Math]::Round(($dtstr.FreeSpace/$dtstr.Capacity)*100)
    }
    $dsAry = $dsAry + $obj.ViReplicaTargetOptions.DatastoreName
    $outputAry = $outputAry + $objoutput
  } Else {
    return
  }
}
}
END {
$outputAry | Select-Object Target, Datastore, StorageFree, StorageTotal, FreePercentage
}
}

Function Get-VeeamVersion {
Try {
$veeamCore = Get-Item -Path $veeamCorePath
$VeeamVersion = [single]($veeamCore.VersionInfo.ProductVersion).substring(0,4)
$productVersion=[string]$veeamCore.VersionInfo.ProductVersion
$productHotfix=[string]$veeamCore.VersionInfo.Comments
$objectVersion = New-Object -TypeName PSObject -Property @{
      VeeamVersion = $VeeamVersion
      productVersion = $productVersion
      productHotfix = $productHotfix
}

Return $objectVersion
} Catch {
    Write-Host "Unable to Locate Veeam Core, check path - $veeamCorePath" -ForegroundColor Red
exit
}
}

Function Get-VeeamSupportDate {
# Query for license info
$licenseInfo = Get-VBRInstalledLicense

$type = $licenseinfo.Type

switch ( $type ) {
    'Perpetual' {
        $date = $licenseInfo.SupportExpirationDate
    }
    'Evaluation' {
        $date = Get-Date
    }
    'Subscription' {
        $date = $licenseInfo.ExpirationDate
    }
    'Rental' {
        $date = $licenseInfo.ExpirationDate
    }
    'NFR' {
        $date = $licenseInfo.ExpirationDate
    }

}

[PSCustomObject]@{
   LicType    = $type
   ExpDate    = $date.ToShortDateString()
   DaysRemain = ($date - (Get-Date)).Days
}
}

Function Get-VMsBackupStatus {
    $outputAry = @()
    $excludevms_regex = ('(?i)^(' + (($script:excludeVMs | ForEach-Object {[regex]::escape($_)}) -join "|") + ')$') -replace "\\\*", ".*"
    $excludefolder_regex = ('(?i)^(' + (($script:excludeFolder | ForEach-Object {[regex]::escape($_)}) -join "|") + ')$') -replace "\\\*", ".*"
    $excludecluster_regex = ('(?i)^(' + (($script:excludeCluster | ForEach-Object {[regex]::escape($_)}) -join "|") + ')$') -replace "\\\*", ".*"
    $excludeTags_regex = ('(?i)^(' + (($script:excludeTags | ForEach-Object {[regex]::escape($_)}) -join "|") + ')$') -replace "\\\*", ".*"
    $excludedc_regex = ('(?i)^(' + (($script:excludeDC | ForEach-Object {[regex]::escape($_)}) -join "|") + ')$') -replace "\\\*", ".*"
    $vms = @{}
    $tagMapping = @{}
    $vmTags = Find-VBRViEntity -Tags | Where-Object { $_.Type -eq "Vm" }
    foreach ($tag in $vmTags) {
        $tagMapping[$tag.Id] = ($tag.Path -split "\\")[-2]
    }

    $allVMsVBRVi |
        Where-Object { $_.VmFolderName -notmatch $excludefolder_regex } |
        Where-Object { $_.Name -notmatch $excludevms_regex } |
        Where-Object { $_.Path.Split("\")[2] -notmatch $excludecluster_regex } |
        Where-Object { $_.Path.Split("\")[1] -notmatch $excludedc_regex } |
        ForEach-Object {
            $vmId = ($_.FindObject().Id, $_.Id -ne $null)[0]
            $tag = if ($tagMapping[$_.Id]) { $tagMapping[$_.Id] } else { "None" }
            if ($tag -notmatch $excludeTags_regex) {
                $vms[$vmId] = @("!", $_.Path.Split("\")[0], $_.Path.Split("\")[1], $_.Path.Split("\")[2], $_.Name, "1/11/1911", "1/11/1911", "", $_.VmFolderName, $tag)
            }
        }

    if (!$script:excludeTemp) {
        Find-VBRViEntity -VMsandTemplates |
            Where-Object { $_.Type -eq "Vm" -and $_.IsTemplate -eq $true -and $_.VmFolderName -notmatch $excludefolder_regex } |
            Where-Object { $_.Name -notmatch $excludevms_regex } |
            Where-Object { $_.Path.Split("\")[2] -notmatch $excludecluster_regex } |
            Where-Object { $_.Path.Split("\")[1] -notmatch $excludedc_regex } |
            ForEach-Object {
                $vmId = ($_.FindObject().Id, $_.Id -ne $null)[0]
                $tag = if ($tagMapping[$_.Id]) { $tagMapping[$_.Id] } else { "None" }
                if ($tag -notmatch $excludeTags_regex) {
                    $vms[$vmId] = @("!", $_.Path.Split("\")[0], $_.Path.Split("\")[1], $_.VmHostName, "[template] $($_.Name)", "1/11/1911", "1/11/1911", "", $_.VmFolderName, $tag)
                }
            }
    }

    $vbrtasksessions = (Get-VBRBackupSession |
        Where-Object {($_.JobType -eq "Backup") -and ($_.EndTime -ge (Get-Date).AddHours(-$script:HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$script:HourstoCheck) -or $_.State -eq "Working")}) |
        Get-VBRTaskSession | Where-Object {$_.Status -notmatch "InProgress|Pending"}

    if ($vbrtasksessions) {
        foreach ($vmtask in $vbrtasksessions) {
            if ($vms.ContainsKey($vmtask.Info.ObjectId)) {
                if ((Get-Date $vmtask.Progress.StartTimeLocal) -ge (Get-Date $vms[$vmtask.Info.ObjectId][5])) {
                    if ($vmtask.Status -eq "Success") {
                        $vms[$vmtask.Info.ObjectId][0]=$vmtask.Status
                        $vms[$vmtask.Info.ObjectId][5]=$vmtask.Progress.StartTimeLocal
                        $vms[$vmtask.Info.ObjectId][6]=$vmtask.Progress.StopTimeLocal
                        $vms[$vmtask.Info.ObjectId][7]=""
                    } elseif ($vms[$vmtask.Info.ObjectId][0] -ne "Success") {
                        $vms[$vmtask.Info.ObjectId][0]=$vmtask.Status
                        $vms[$vmtask.Info.ObjectId][5]=$vmtask.Progress.StartTimeLocal
                        $vms[$vmtask.Info.ObjectId][6]=$vmtask.Progress.StopTimeLocal
                        $vms[$vmtask.Info.ObjectId][7]=($vmtask.GetDetails()).Replace("<br />","ZZbrZZ")
                    }
                } elseif ($vms[$vmtask.Info.ObjectId][0] -match "Warning|Failed" -and $vmtask.Status -eq "Success") {
                    $vms[$vmtask.Info.ObjectId][0]=$vmtask.Status
                    $vms[$vmtask.Info.ObjectId][5]=$vmtask.Progress.StartTimeLocal
                    $vms[$vmtask.Info.ObjectId][6]=$vmtask.Progress.StopTimeLocal
                    $vms[$vmtask.Info.ObjectId][7]=""
                }
            }
        }
    }

    foreach ($vm in $vms.GetEnumerator()) {
        $objoutput = [PSCustomObject]@{
            Status     = $vm.Value[0]
            Name       = $vm.Value[4]
            vCenter    = $vm.Value[1]
            Datacenter = $vm.Value[2]
            Cluster    = $vm.Value[3]
            StartTime  = $vm.Value[5]
            StopTime   = $vm.Value[6]
            Details    = $vm.Value[7]
            Folder     = $vm.Value[8]
            Tags       = $vm.Value[9]
        }
        $outputAry += $objoutput
    }

    return $outputAry
}

function Get-Duration {
param ($ts)
$days = ""
If ($ts.Days -gt 0) {
$days = "{0}:" -f $ts.Days
}
"{0}{1}:{2,2:D2}:{3,2:D2}" -f $days,$ts.Hours,$ts.Minutes,$ts.Seconds
}

function Get-BackupSize {
param ($backups)
$outputObj = @()
Foreach ($backup in $backups) {
$backupSize = 0
$dataSize = 0
$logSize = 0
$files = $backup.GetAllStorages()
Foreach ($file in $Files) {
  $backupSize += [math]::Round([long]$file.Stats.BackupSize/1GB, 2)
  $dataSize += [math]::Round([long]$file.Stats.DataSize/1GB, 2)
}
#Added Log Backup Reporting
$childBackups = $backup.FindChildBackups()
if($childBackups.count -gt 0) {
  $logFiles = $childBackups.GetAllStorages()
  Foreach ($logFile in $logFiles) {
    $logSize += [math]::Round([long]$logFile.Stats.BackupSize/1GB, 2)
  }
}
$repo = If ($($script:repoList | Where-Object {$_.Id -eq $backup.RepositoryId}).Name) {
          $($script:repoList | Where-Object {$_.Id -eq $backup.RepositoryId}).Name
        } Else {
          $($script:repoListSo | Where-Object {$_.Id -eq $backup.RepositoryId}).Name
        }
$vbrMasterHash = @{
  JobName = $backup.JobName
  VMCount = $backup.VmCount
  Repo = $repo
  DataSize = $dataSize
  BackupSize = $backupSize
  LogSize = $logSize
}
$vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
$outputObj += $vbrMasterObj
}
$outputObj
}

Function Get-MultiJob {
$outputAry = @()
$vmMultiJobs = (Get-VBRBackupSession |
Where-Object {($_.JobType -eq "Backup") -and ($_.EndTime -ge (Get-Date).addhours(-$script:HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$script:HourstoCheck) -or $_.State -eq "Working")}) |
Get-VBRTaskSession | Select-Object Name, @{Name="VMID"; Expression = {$_.Info.ObjectId}}, JobName -Unique | Group-Object Name, VMID | Where-Object {$_.Count -gt 1} | Select-Object -ExpandProperty Group
ForEach ($vm in $vmMultiJobs) {
$objID = $vm.VMID
$viEntity = Find-VBRViEntity -name $vm.Name | Where-Object {$_.FindObject().Id -eq $objID}
If ($null -ne $viEntity) {
  $objoutput = New-Object -TypeName PSObject -Property @{
    Name = $vm.Name
    vCenter = $viEntity.Path.Split("\")[0]
    Datacenter = $viEntity.Path.Split("\")[1]
    Cluster = $viEntity.Path.Split("\")[2]
    Folder = $viEntity.VMFolderName
    JobName = $vm.JobName
  }
  $outputAry += $objoutput
} Else { #assume Template
  $viEntity = Find-VBRViEntity -VMsAndTemplates -name $vm.Name | Where-Object {$_.FindObject().Id -eq $objID}
  If ($null -ne $viEntity) {
    $objoutput = New-Object -TypeName PSObject -Property @{
      Name = "[template] " + $vm.Name
      vCenter = $viEntity.Path.Split("\")[0]
      Datacenter = $viEntity.Path.Split("\")[1]
      Cluster = $viEntity.VmHostName
      Folder = $viEntity.VMFolderName
      JobName = $vm.JobName
    }
  }
  If ($objoutput) {
    $outputAry += $objoutput
  }
}
}
$outputAry
}
#endregion

#region Report
# Get Veeam Version
$objectVersion = (Get-VeeamVersion).productVersion

# CrÃ©ation dâ€™un hashtable
$jsonHash = [ordered]@{}
$jsonHash.Add("reportVersion", $reportVersion)
$jsonHash.Add("generationDate", $date_title)
$jsonHash.Add("client", $Client)
$jsonHash.Add("reportMode", $reportMode)
$jsonHash.Add("vbrServerName", $vbrServer)
$jsonHash.Add("vbrServerVersion", $objectVersion)

#region Get VM Backup Status
$vmStatus = @()
If ($showSummaryProtect + $showUnprotectedVMs + $showProtectedVMs) {
$vmStatus = Get-VMsBackupStatus
}

# VMs Missing Backups
$missingVMs = @($vmStatus | Where-Object {$_.Status -match "!|Failed"})
ForEach ($VM in $missingVMs) {
If ($VM.Status -eq "!") {
$VM.Details = "No Backup Task has completed"
$VM.StartTime = ""
$VM.StopTime = ""
}
}
# VMs Successfuly Backed Up
$successVMs = @($vmStatus | Where-Object {$_.Status -eq "Success"})
# VMs Backed Up w/Warning
$warnVMs = @($vmStatus | Where-Object {$_.Status -eq "Warning"})
#endregion

#region Get VM Backup Protection Summary
If ($showSummaryProtect) {
If (@($successVMs).Count -ge 1) {
$percentProt = 1
}
If (@($missingVMs).Count -ge 1) {
$percentProt = (@($warnVMs).Count + @($successVMs).Count) / (@($warnVMs).Count + @($successVMs).Count + @($missingVMs).Count)
    }
}
$vbrMasterHash = @{
    WarningVM = @($warnVMs).Count
    ProtectedVM = @($successVMs).Count
    UnprotectedVM = @($missingVMs).Count
    ExcludedVM = ($allVMsVBRVi).count - ($vmStatus).count
    PercentProtected = [Math]::Floor($percentProt * 100)
  }

$jsonHash["VMBackupProtectionSummary"] = $vbrMasterHash
#endregion

#region Get VMs Missing Backups
If ($showUnprotectedVMs) {
  If ($missingVMs.count -gt 0) {
    $missingVMs = $missingVMs | Sort-Object vCenter, Datacenter, Cluster, Name | ForEach-Object {$_ | Select-Object Name, vCenter, Datacenter, Cluster, Folder, Tags,
        @{Name="StartTime"; Expression = { $_.StartTime.ToString("dd/MM/yyyy HH:mm") }},
        @{Name="StopTime"; Expression = { $_.StopTime.ToString("dd/MM/yyyy HH:mm") }}}
      $jsonHash["missingVms"] = $missingVMs
  }
}
#endregion

#region Get VMs Backed Up w/Warnings
If ($showProtectedVMs) {
  If ($warnVMs.Count -gt 0) {
  $warnVMs = $warnVMs | Sort-Object vCenter, Datacenter, Cluster, Name | ForEach-Object {
    $_ | Select-Object Name, vCenter, Datacenter, Cluster, Folder, Tags,
        @{Name="StartTime"; Expression = { $_.StartTime.ToString("dd/MM/yyyy HH:mm") }},
        @{Name="StopTime"; Expression = { $_.StopTime.ToString("dd/MM/yyyy HH:mm") }}}
    $jsonHash["warnVMs"] = $warnVMs
  }
}


# Get VMs Successfuly Backed Up
If ($showProtectedVMs) {
  If ($successVMs.Count -gt 0) {
  $successVMs = $successVMs | Sort-Object vCenter, Datacenter, Cluster, Name | ForEach-Object {
    $_ | Select-Object Name, vCenter, Datacenter, Cluster, Folder, Tags,
        @{Name="StartTime"; Expression = { $_.StartTime.ToString("dd/MM/yyyy HH:mm") }},
        @{Name="StopTime"; Expression = { $_.StopTime.ToString("dd/MM/yyyy HH:mm") }}}
}
    $jsonHash["successVMs"] = $successVMs
  }
#endregion


# Get VMs Backed Up by Multiple Jobs
If ($showMultiJobs) {
$multiJobs = @(Get-MultiJob)
If ($multiJobs.Count -gt 0) {
    $multiJobs = $multiJobs | Sort-Object vCenter, Datacenter, Cluster, Name | Select-Object Name, vCenter, Datacenter, Cluster, Folder,
      @{Name="Job Name"; Expression = {$_.JobName}}
    $jsonHash["multiJobs"] = $multiJobs
}
}

# Get Backup Summary Info
$arrSummaryBk = $null
If ($showSummaryBk) {
$vbrMasterHash = @{
"Failed" = @($failedSessionsBk).Count
"Sessions" = If ($sessListBk) {@($sessListBk).Count} Else {0}
"Read" = $totalReadBk
"Transferred" = $totalXferBk
"Successful" = @($successSessionsBk).Count
"Warning" = @($warningSessionsBk).Count
"Fails" = @($failsSessionsBk).Count
"Running" = @($runningSessionsBk).Count
}
$vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
If ($onlyLastBk) {
$total = "Jobs Run"
} Else {
$total = "Total Sessions"
}
$arrSummaryBk =  $vbrMasterObj | Select-Object @{Name=$total; Expression = {$_.Sessions}},
@{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
@{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
@{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}},
@{Name="Failed"; Expression = {$_.Failed}}
  $jsonHash["SummaryBk"] = $arrSummaryBk
}

# Get Backup Job Status
if ($showJobsBk -and $allJobsBk.Count -gt 0) {
  $bodyJobsBk = @()
  foreach ($bkJob in $allJobsBk) {
    $bodyJobsBk += ($bkJob | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.IsScheduleEnabled}},
        @{Name="State"; Expression = {
          if ($bkJob.IsRunning) {
            $s = $runningSessionsBk | Where-Object {$_.JobName -eq $bkJob.Name}
            if ($s) {"$($s.Progress.Percents)% completed at $([Math]::Round($s.Progress.AvgSpeed/1MB,2)) MB/s"}
            else {"Running (no session info)"}
          } else {"Stopped"}
        }},
        @{Name="Target Repo"; Expression = {
          ($repoList + $repoListSo | Where-Object {$_.Id -eq $bkJob.Info.TargetRepositoryId}).Name
        }},
        @{Name="Next Run"; Expression = {
          try {
            $s = Get-VBRJobScheduleOptions -Job $bkJob
            if (-not $bkJob.IsScheduleEnabled) {"Disabled"}
            elseif ($s.RunManually) {"Not Scheduled"}
            elseif ($s.IsContinious) {"Continious"}
            elseif ($s.OptionsScheduleAfterJob.IsEnabled) {
              "After [$(($allJobs + $allJobsTp | Where-Object {$_.Id -eq $bkJob.Info.ParentScheduleId}).Name)]"
            } else { $s.NextRun }
          } catch { "Unavailable" }
        }},
        @{Name="Status"; Expression = {
          if ($_.Info.LatestStatus -eq "None") {"Unknown"} else { $_.Info.LatestStatus.ToString() }
        }}
    )
  }
  $jsonHash["JobsBk"] = $bodyJobsBk
}

# Get Backup Job Status Begin
$bodyFileJobsBk = $null
if ($showFileJobsBk -and $allFileJobsBk.Count -gt 0) {
$bodyFileJobsBk = @()
  foreach ($bkJob in $allFileJobsBk) {
    $bodyFileJobsBk += ($bkJob | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Enabled"; Expression = {$_.IsScheduleEnabled}},
    @{Name="Status"; Expression = {
          if ($bkJob.IsRunning) {
            $s = $runningSessionsBk | Where-Object {$_.JobName -eq $bkJob.Name}
            if ($s) {"$($s.Progress.Percents)% completed at $([Math]::Round($s.Progress.AvgSpeed/1MB,2)) MB/s"}
            else {"Running (no session info)"}
          } else {"Stopped"}
    }},
    @{Name="Target Repo"; Expression = {
          ($repoList + $repoListSo | Where-Object {$_.Id -eq $bkJob.Info.TargetRepositoryId}).Name
    }},
    @{Name="Next Run"; Expression = {
          try {
            $s = Get-VBRJobScheduleOptions -Job $bkJob
            if (-not $bkJob.IsScheduleEnabled) {"Disabled"}
            elseif ($s.RunManually) {"Not Scheduled"}
            elseif ($s.IsContinious) {"Continious"}
            elseif ($s.OptionsScheduleAfterJob.IsEnabled) {
              "After [$(($allJobs + $allJobsTp | Where-Object {$_.Id -eq $bkJob.Info.ParentScheduleId}).Name)]"
            } else { $s.NextRun }
          } catch { "Unavailable" }
    }},
        @{Name="Status"; Expression = {If ($_.Info.LatestStatus -eq "None"){"Unknown"}Else{$_.Info.LatestStatus.ToString()}}})}
$jsonHash["FileJobsBk"] = $bodyFileJobsBk
}
# Get File Backup Job Status End

# Get all Backup Sessions
$arrAllSessBk = $null
If ($showAllSessBk) {
If ($sessListBk.count -gt 0) {
If ($showDetailedBk) {
  $arrAllSessBk = $sessListBk | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="State"; Expression = {$_.State.ToString()}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
    @{Name="Dedupe"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
      Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
    @{Name="Compression"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
      Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}},
    @{Name="Result"; Expression = {($_.Result.ToString())}}
    $jsonHash["AllSessBk"] = $arrAllSessBk
} Else {
  $arrAllSessBk = $sessListBk | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="State"; Expression = {$_.State.ToString()}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}},
    @{Name="Result"; Expression = {($_.Result.ToString())}}
    $jsonHash["AllSessBk"] = $arrAllSessBk
}
}
}

# Get Running Backup Jobs
$bodyRunningBk = $null
If ($showRunningBk) {
If ($runningSessionsBk.count -gt 0) {
$bodyRunningBk = $runningSessionsBk | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
  @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
  @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
  @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
  @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
  @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
  @{Name="% Complete"; Expression = {$_.Progress.Percents}}
  $jsonHash["RunningBk"] = $bodyRunningBk
}
}

# Get Backup Sessions with Warnings or Failures
$arrSessWFBk = $null
If ($showWarnFailBk) {
$sessWF = @($warningSessionsBk + $failsSessionsBk)
If ($sessWF.count -gt 0) {
If ($showDetailedBk) {
  $arrSessWFBk = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
    @{Name="Dedupe"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
      Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
    @{Name="Compression"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
      Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
    @{Name="Details"; Expression = {
      If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
      Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
    $jsonHash["SessWFBk"] = $arrSessWFBk
} Else {
  $arrSessWFBk = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {
      If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
      Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
    $jsonHash["SessWFBk"] = $arrSessWFBk
}
}
}

# Get Successful Backup Sessions
$bodySessSuccBk = $null
If ($showSuccessBk) {
If ($successSessionsBk.count -gt 0) {
If ($showDetailedBk) {
  $bodySessSuccBk = $successSessionsBk | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
    @{Name="Dedupe"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
      Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
    @{Name="Compression"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
      Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
    @{Name="Result"; Expression = {($_.Result.ToString())}}
    $jsonHash["SessSuccBk"] = $bodySessSuccBk
} Else {
  $bodySessSuccBk = $successSessionsBk | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},@{Name="Result"; Expression = {($_.Result.ToString())}}
    $jsonHash["SessSuccBk"] = $bodySessSuccBk
}
}
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Backup Tasks from Sessions within time frame
$taskListBk = @()
$taskListBk += $sessListBk | Get-VBRTaskSession
$successTasksBk = @($taskListBk | Where-Object {$_.Status -eq "Success"})
$wfTasksBk = @($taskListBk | Where-Object {$_.Status -match "Warning|Failed"})
$runningTasksBk = @()
$runningTasksBk += $runningSessionsBk | Get-VBRTaskSession | Where-Object {$_.Status -match "Pending|InProgress"}

# Get all Backup Tasks
$bodyAllTasksBk = $null
If ($showAllTasksBk) {
If ($taskListBk.count -gt 0) {
If ($showDetailedBk) {
  $arrAllTasksBk = $taskListBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, @{Name="Status"; Expression = {($_.Status.ToString())}}
      $bodyAllTasksBk = $arrAllTasksBk | Sort-Object "Start Time"
      $jsonHash["AlltaskBk"] = $bodyAllTasksBk
} Else {
  $arrAllTasksBk = $taskListBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, @{Name="Status"; Expression = {($_.Status.ToString())}}
      $bodyAllTasksBk = $arrAllTasksBk | Sort-Object "Start Time"
      $jsonHash["AlltaskBk"] = $bodyAllTasksBk
}
}
}

# Get Running Backup Tasks
$bodyTasksRunningBk = $null
If ($showRunningTasksBk) {
If ($runningTasksBk.count -gt 0) {
$bodyTasksRunningBk = $runningTasksBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort-Object "Start Time" 
    $jsonHash["TasksRunningBk"] = $bodyTasksRunningBk
}
}

# Get Backup Tasks with Warnings or Failures
$bodyTaskWFBk = $null
If ($showTaskWFBk) {
If ($wfTasksBk.count -gt 0) {
If ($showDetailedBk) {
  $arrTaskWFBk = $wfTasksBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, @{Name="Status"; Expression = {($_.Status.ToString())}}
      $bodyTaskWFBk = $arrTaskWFBk | Sort-Object "Start Time"
      $jsonHash["TaskWFBk"] = $bodyTaskWFBk
} Else {
  $arrTaskWFBk = $wfTasksBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFBk = $arrTaskWFBk | Sort-Object "Start Time"
      $jsonHash["TaskWFBk"] = $bodyTaskWFBk
}
}
}

# Get Successful Backup Tasks
$bodyTaskSuccBk = $null
If ($showTaskSuccessBk) {
If ($successTasksBk.count -gt 0) {
If ($showDetailedBk) {
  $bodyTaskSuccBk = $successTasksBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},@{Name="Status"; Expression = {($_.Status.ToString())}} | Sort-Object "Start Time"
      $jsonHash["TaskSuccBk"] = $bodyTaskSuccBk
} Else {
  $bodyTaskSuccBk = $successTasksBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}}, @{Name="Status"; Expression = {($_.Status.ToString())}} | Sort-Object "Start Time"
      $jsonHash["TaskSuccBk"] = $bodyTaskSuccBk
    }
}
}

# Get Replication Summary Info
$arrSummaryRp = $null
If ($showSummaryRp) {
$vbrMasterHash = @{
"Failed" = @($failedSessionsRp).Count
"Sessions" = If ($sessListRp) {@($sessListRp).Count} Else {0}
"Read" = $totalReadRp
"Transferred" = $totalXferRp
"Successful" = @($successSessionsRp).Count
"Warning" = @($warningSessionsRp).Count
"Fails" = @($failsSessionsRp).Count
"Running" = @($runningSessionsRp).Count
}
$vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
If ($onlyLastRp) {
$total = "Jobs Run"
} Else {
$total = "Total Sessions"
}
$arrSummaryRp =  $vbrMasterObj | Select-Object @{Name=$total; Expression = {$_.Sessions}},
@{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
@{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
@{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Fails"; Expression = {$_.Fails}},
@{Name="Failed"; Expression = {$_.Failed}}
  $jsonHash["SummaryRp"] = $arrSummaryRp
}

# Get Replication Job Status
$bodyJobsRp = $null
if ($showJobsRp -and $allJobsRp.Count -gt 0) {
$bodyJobsRp = @()
  foreach ($rpJob in $allJobsRp) {
    $bodyJobsRp += (
      $rpJob | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Enabled"; Expression = {$_.Info.IsScheduleEnabled}},
    @{Name="State"; Expression = {
          if ($rpJob.IsRunning) {
            $s = $runningSessionsRp | Where-Object {$_.JobName -eq $rpJob.Name}
            if ($s) {
              "$($s.Progress.Percents)% completed at $([Math]::Round($s.Info.Progress.AvgSpeed/1MB,2)) MB/s"
            } else {"Running (no session info)"
            }
          } else {"Stopped"}}},
        @{Name="Target"; Expression = {(Get-VBRServer | Where-Object {$_.Id -eq $rpJob.Info.TargetHostId}).Name}},
        @{Name="Target Repo"; Expression = {($repoList + $repoListSo | Where-Object {$_.Id -eq $rpJob.Info.TargetRepositoryId}).Name}},
    @{Name="Next Run"; Expression = {
      try {
        $s = Get-VBRJobScheduleOptions -Job $rpJob
        if (-not $rpJob.IsScheduleEnabled) {"Disabled"}
        elseif ($s.RunManually) {"Not Scheduled"}
        elseif ($s.IsContinious) {"Continious"}
        elseif ($s.OptionsScheduleAfterJob.IsEnabled) {
          "After [$(($allJobs + $allJobsTp | Where-Object {$_.Id -eq $rpJob.Info.ParentScheduleId}).Name)]"
        } else {$s.NextRun}
      } catch {"Unavailable"}}},
        @{Name="Status"; Expression = {$result = $_.GetLastResult()
            if ($result -eq "None") { "" } else { $result.ToString() }}})}
    $bodyJobsRp = $bodyJobsRp | Sort-Object "Next Run" 
    $jsonHash["JobsRp"] = $bodyJobsRp
}

# Get Replication Sessions
$arrAllSessRp = $null
If ($showAllSessRp) {
If ($sessListRp.count -gt 0) {
If ($showDetailedRp) {
  $arrAllSessRp = $sessListRp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="State"; Expression = {$_.State.ToString()}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
    @{Name="Dedupe"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
      Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
    @{Name="Compression"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
      Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
      $jsonHash["AllSessRp"] = $arrAllSessRp
} Else {
  $arrAllSessRp = $sessListRp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="State"; Expression = {$_.State.ToString()}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
      $jsonHash["AllSessRp"] = $arrAllSessRp
}
}
}

# Get Running Replication Jobs
$bodyRunningRp = $null
If ($showRunningRp) {
If ($runningSessionsRp.count -gt 0) {
$bodyRunningRp = $runningSessionsRp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
  @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
  @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
  @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
  @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
  @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
  @{Name="% Complete"; Expression = {$_.Progress.Percents}}
  $jsonHash["RunningRp"] = $bodyRunningRp
}
}

# Get Replication Sessions with Warnings or Failures
$arrSessWFRp = $null
if ($showWarnFailRp) {
    $sessWF = @($warningSessionsRp + $failsSessionsRp)
    if ($showDetailedRp) {
        $arrSessWFRp = $sessWF | Sort-Object CreationTime | Select-Object `
            @{Name = "Job Name"; Expression = { $_.Name }},
            @{Name = "Start Time"; Expression = { $_.CreationTime.ToString("dd/MM/yyyy HH:mm") }},
            @{Name = "Stop Time"; Expression = { $_.EndTime.ToString("dd/MM/yyyy HH:mm") }},
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration }},
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Info.Progress.AvgSpeed / 1MB, 2) }},
            @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedSize / 1GB, 2) }},
            @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Info.Progress.ProcessedUsedSize / 1GB, 2) }},
            @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Info.Progress.ReadSize / 1GB, 2) }},
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Info.Progress.TransferedSize / 1GB, 2) }},
            @{Name = "Dedupe"; Expression = {
                if ($_.Progress.ReadSize -eq 0) { 0 }
                else { [string]([Math]::Round($_.BackupStats.GetDedupeX(), 1)) + "x" }}},
            @{Name = "Compression"; Expression = {
                if ($_.Progress.ReadSize -eq 0) { 0 }
                else { [string]([Math]::Round($_.BackupStats.GetCompressX(), 1)) + "x" }}},
            @{Name = "Details"; Expression = {
                if ($_.GetDetails() -eq "") {$_ | Get-VBRTaskSession | ForEach-Object {
                        if ($_.GetDetails()) {$_.Name + ": " + ($_.GetDetails()).Replace("<br />", "ZZbrZZ")}}
                } else {($_.GetDetails()).Replace("<br />", "ZZbrZZ")}}}
        $jsonHash["SessWFRp"] = $arrSessWFRp
    } else {
        $arrSessWFRp = $sessWF | Sort-Object CreationTime | Select-Object `
            @{Name = "Job Name"; Expression = { $_.Name }},
            @{Name = "Start Time"; Expression = { $_.CreationTime.ToString("dd/MM/yyyy HH:mm") }},
            @{Name = "Stop Time"; Expression = { $_.EndTime.ToString("dd/MM/yyyy HH:mm") }},
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration }},
            @{Name = "Details"; Expression = {
                if ($_.GetDetails() -eq "") {$_ | Get-VBRTaskSession | ForEach-Object {
                        if ($_.GetDetails()) {$_.Name + ": " + ($_.GetDetails()).Replace("<br />", "ZZbrZZ")}}
                } else {($_.GetDetails()).Replace("<br />", "ZZbrZZ")}}},
            @{Name = "Result"; Expression = { $_.Result.ToString() }}
        $jsonHash["SessWFRp"] = $arrSessWFRp
    }
}


# Get Successful Replication Sessions
$bodySessSuccRp = $null
If ($showSuccessRp) {
If ($successSessionsRp.count -gt 0) {
If ($showDetailedRp) {
  $bodySessSuccRp = $successSessionsRp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
    @{Name="Dedupe"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
      Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
    @{Name="Compression"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
      Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Result"; Expression = {($_.Result.ToString())}}
        $jsonHash["SessSuccRp"] = $bodySessSuccRp
} Else {
  $bodySessSuccRp = $successSessionsRp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
    $jsonHash["SessSuccRp"] = $bodySessSuccRp
}
}
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Replication Tasks from Sessions within time frame
$taskListRp = @()
$taskListRp += $sessListRp | Get-VBRTaskSession
$successTasksRp = @($taskListRp | Where-Object {$_.Status -eq "Success"})
$wfTasksRp = @($taskListRp | Where-Object {$_.Status -match "Warning|Failed"})
$runningTasksRp = @()
$runningTasksRp += $runningSessionsRp | Get-VBRTaskSession | Where-Object {$_.Status -match "Pending|InProgress"}

# Get Replication Tasks
$bodyAllTasksRp = $null
If ($showAllTasksRp) {
If ($taskListRp.count -gt 0) {
If ($showDetailedRp) {
  $arrAllTasksRp = $taskListRp | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, @{Name="Status"; Expression = {($_.Status.ToString())}}
    $bodyAllTasksRp = $arrAllTasksRp | Sort-Object "Start Time"
    $jsonHash["AllTasksRp"] = $bodyAllTasksRp
} Else {
  $arrAllTasksRp = $taskListRp | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
    $bodyAllTasksRp = $arrAllTasksRp | Sort-Object "Start Time"
    $jsonHash["AllTasksRp"] = $bodyAllTasksRp
}
}
}

# Get Running Replication Tasks
$bodyTasksRunningRp = $null
If ($showRunningTasksRp) {
If ($runningTasksRp.count -gt 0) {
$bodyTasksRunningRp = $runningTasksRp | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort-Object "Start Time"
    $jsonHash["TasksRunningRp"] = $bodyTasksRunningRp
}
}

# Get Replication Tasks with Warnings or Failures
$bodyTaskWFRp = $null
If ($showTaskWFRp) {
If ($wfTasksRp.count -gt 0) {
If ($showDetailedRp) {
  $arrTaskWFRp = $wfTasksRp | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, 
    @{Name="Status"; Expression = {($_.Status.ToString())}}
      $bodyTaskWFRp = $arrTaskWFRp | Sort-Object "Start Time"
      $jsonHash["TaskWFRp"] = $bodyTaskWFRp
} Else {
  $arrTaskWFRp = $wfTasksRp | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}},
    @{Name="Status"; Expression = {($_.Status.ToString())}}
      $bodyTaskWFRp = $arrTaskWFRp | Sort-Object "Start Time" 
      $jsonHash["TaskWFRp"] = $bodyTaskWFRp
}
}
}

# Get Successful Replication Tasks
$bodyTaskSuccRp = $null
If ($showTaskSuccessRp) {
If ($successTasksRp.count -gt 0) {
If ($showDetailedRp) {
  $bodyTaskSuccRp = $successTasksRp | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
    @{Name="Status"; Expression = {($_.Status.ToString())}} | Sort-Object "Start Time" 
       $jsonHash["TaskSuccRp"] = $bodyTaskSuccRp
} Else {
  $bodyTaskSuccRp = $successTasksRp | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Status"; Expression = {($_.Status.ToString())}} | Sort-Object "Start Time" 
      $jsonHash["TaskSuccRp"] = $bodyTaskSuccRp
}
}
}

# Get Backup Copy Summary Info
$arrSummaryBc = $null
If ($showSummaryBc) {
$vbrMasterHash = @{
"Sessions" = If ($sessListBc) {@($sessListBc).Count} Else {0}
"Read" = $totalReadBc
"Transferred" = $totalXferBc
"Successful" = @($successSessionsBc).Count
"Warning" = @($warningSessionsBc).Count
"Fails" = @($failsSessionsBc).Count
"Working" = @($workingSessionsBc).Count
"Idle" = @($idleSessionsBc).Count
}
$vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
If ($onlyLastBc) {
$total = "Jobs Run"
} Else {
$total = "Total Sessions"
}
$arrSummaryBc =  $vbrMasterObj | Select-Object @{Name=$total; Expression = {$_.Sessions}},
@{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
@{Name="Idle"; Expression = {$_.Idle}},
@{Name="Working"; Expression = {$_.Working}}, @{Name="Successful"; Expression = {$_.Successful}},
@{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $jsonHash["SummaryBc"] = $arrSummaryBc
}

# Get Backup Copy Job Status
$bodyJobsBc = $null
if ($showJobsBc -and $allJobsBc.Count -gt 0) {
$bodyJobsBc = @()
  foreach ($BcJob in $allJobsBc) {
    $bodyJobsBc += ($BcJob | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Enabled"; Expression = {$_.Info.IsScheduleEnabled}},
    @{Name="Type"; Expression = {$_.TypeToString}},
    @{Name="State"; Expression = {
          if ($BcJob.IsRunning) {
        $currentSess = $BcJob.FindLastSession()
            if ($currentSess.State -eq "Working") {
              "$($currentSess.Progress.Percents)% completed at $([Math]::Round($currentSess.Progress.AvgSpeed/1MB,2)) MB/s"
            } else {
          $currentSess.State
        }
          } else {"Stopped"}}},
        @{Name="Target Repo"; Expression = {
          ($repoList + $repoListSo | Where-Object {$_.Id -eq $BcJob.Info.TargetRepositoryId}).Name
    }},
    @{Name="Next Run"; Expression = {
          try {
            $s = Get-VBRJobScheduleOptions -Job $BcJob
            if (-not $BcJob.IsScheduleEnabled) {"Disabled"}
            elseif ($s.RunManually) {"Not Scheduled"}
            elseif ($s.IsContinious) {"Continious"}
            elseif ($s.OptionsScheduleAfterJob.IsEnabled) {
              "After [$(($allJobs + $allJobsTp | Where-Object {$_.Id -eq $BcJob.Info.ParentScheduleId}).Name)]"
            } else {
              $s.NextRun
            }
          } catch {
            "Unavailable"
          }
        }},
        @{Name="Status"; Expression = {
          if ($_.Info.LatestStatus -eq "None") {""} else { $_.Info.LatestStatus.ToString() }
        }}
    )
}
    $bodyJobsBc = $bodyJobsBc | Sort-Object "Next Run", "Job Name"
    $jsonHash["JobsBc"] = $bodyJobsBc
}

# Get All Backup Copy Sessions
$arrAllSessBc = $null
If ($showAllSessBc) {
If ($sessListBc.count -gt 0) {
If ($showDetailedBc) {
  $arrAllSessBc = $sessListBc | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
    @{Name="Dedupe"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
      Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
    @{Name="Compression"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
      Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
      $jsonHash["AllSessBc"] = $arrAllSessBc
} Else {
  $arrAllSessBc = $sessListBc | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="State"; Expression = {$_.State.ToString()}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, 
    @{Name="Result"; Expression = {($_.Result.ToString())}}
    $jsonHash["AllSessBc"] = $arrAllSessBc
}
}
}

# Get Idle Backup Copy Sessions
$bodySessIdleBc = $null
If ($showIdleBc) {
If ($idleSessionsBc.count -gt 0) {
If ($showDetailedBc) {
  $bodySessIdleBc = $idleSessionsBc | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
    @{Name="Dedupe"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
      Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
    @{Name="Compression"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}}
      $jsonHash["SessIdleBc"] = $bodySessIdleBc
} Else {
  $bodySessIdleBc = $idleSessionsBc | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}}
      $jsonHash["SessIdleBc"] = $bodySessIdleBc
}
}
}

# Get Working Backup Copy Jobs
$bodyRunningBc = $null
If ($showRunningBc) {
If ($workingSessionsBc.count -gt 0) {
$bodyRunningBc = $workingSessionsBc | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
  @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
  @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Progress.StartTimeLocal $(Get-Date))}},
  @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
  @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
  @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}}
    $jsonHash["RunningBc"] = $bodyRunningBc
}
}

# Get Backup Copy Sessions with Warnings or Failures
$arrSessWFBc = $null
If ($showWarnFailBc) {
$sessWF = @($warningSessionsBc + $failsSessionsBc)
If ($sessWF.count -gt 0) {
If ($showDetailedBc) {
  $arrSessWFBc = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
    @{Name="Dedupe"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
      Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
    @{Name="Compression"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
      Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
    @{Name="Details"; Expression = {
      If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
      Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
      $jsonHash["SessWFBc"] = $arrSessWFBc
} Else {
  $arrSessWFBc = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {
      If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
      Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
      $jsonHash["SessWFBc"] = $arrSessWFBc
}
}
}

# Get Successful Backup Copy Sessions
$bodySessSuccBc = $null
If ($showSuccessBc) {
If ($successSessionsBc.count -gt 0) {
If ($showDetailedBc) {
  $bodySessSuccBc = $successSessionsBc | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
    @{Name="Dedupe"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
      Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
    @{Name="Compression"; Expression = {
      If ($_.Progress.ReadSize -eq 0) {0}
      Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Result"; Expression = {($_.Result.ToString())}}
      $jsonHash["SessSuccBc"] = $bodySessSuccBc
} Else {
  $bodySessSuccBc = $successSessionsBc | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Result
      $jsonHash["SessSuccBc"] = $bodySessSuccBc
}
}
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Backup Copy Tasks from Sessions within time frame
$taskListBc = @()
$taskListBc += $sessListBc | Get-VBRTaskSession
$successTasksBc = @($taskListBc | Where-Object {$_.Status -eq "Success"})
$wfTasksBc = @($taskListBc | Where-Object {$_.Status -match "Warning|Failed"})
$pendingTasksBc = @($taskListBc | Where-Object {$_.Status -eq "Pending"})
$runningTasksBc = @($taskListBc | Where-Object {$_.Status -eq "InProgress"})

# Get All Backup Copy Tasks
$bodyAllTasksBc = $null
If ($showAllTasksBc) {
If ($taskListBc.count -gt 0) {
If ($showDetailedBc) {
  $arrAllTasksBc = $taskListBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, 
    @{Name="Status"; Expression = {($_.Status.ToString())}}
      $bodyAllTasksBc = $arrAllTasksBc | Sort-Object "Start Time"
      $jsonHash["AllTasksBc"] = $bodyAllTasksBc
} Else {
  $arrAllTasksBc = $taskListBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, 
    @{Name="Status"; Expression = {($_.Status.ToString())}}
      $bodyAllTasksBc = $arrAllTasksBc | Sort-Object "Start Time"
      $jsonHash["AllTasksBc"] = $bodyAllTasksBc
}
}
}

# Get Pending Backup Copy Tasks
$bodyTasksPendingBc = $null
If ($showPendingTasksBc) {
If ($pendingTasksBc.count -gt 0) {
$bodyTasksPendingBc = $pendingTasksBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort-Object "Start Time"
    $jsonHash["TasksPendingBc"] = $bodyTasksPendingBc
}
}

# Get Working Backup Copy Tasks
$bodyTasksRunningBc = $null
If ($showRunningTasksBc) {
If ($runningTasksBc.count -gt 0) {
$bodyTasksRunningBc = $runningTasksBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort-Object "Start Time"
    $jsonHash["TasksRunningBc"] = $bodyTasksRunningBc
}
}

# Get Backup Copy Tasks with Warnings or Failures
$arrTaskWFBc = $null
If ($showTaskWFBc) {
If ($wfTasksBc.count -gt 0) {
If ($showDetailedBc) {
  $arrTaskWFBc = $wfTasksBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, 
    @{Name="Status"; Expression = {($_.Status.ToString())}}
      $jsonHash["TaskWFBc"] = $arrTaskWFBc
} Else {
  $arrTaskWFBc = $wfTasksBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, 
    @{Name="Status"; Expression = {($_.Status.ToString())}}
      $jsonHash["TaskWFBc"] = $arrTaskWFBc
}
}
}

# Get Successful Backup Copy Tasks
$bodyTaskSuccBc = $null
If ($showTaskSuccessBc) {
If ($successTasksBc.count -gt 0) {
If ($showDetailedBc) {
  $bodyTaskSuccBc = $successTasksBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {
      If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
      Else {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}
    }},
    @{Name="Duration (HH:MM:SS)"; Expression = {
      If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
      Else {Get-Duration -ts $_.Progress.Duration}
    }},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
    @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
    @{Name="Status"; Expression = {($_.Status.ToString())}} | Sort-Object "Start Time"
        $jsonHash["TaskSuccBc"] = $bodyTaskSuccBc
} Else {
  $bodyTaskSuccBc = $successTasksBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {
      If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
      Else {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}
    }},
    @{Name="Duration (HH:MM:SS)"; Expression = {
      If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
      Else {Get-Duration -ts $_.Progress.Duration}
    }},
    @{Name="Status"; Expression = {($_.Status.ToString())}} | Sort-Object "Start Time"
        $jsonHash["TaskSuccBc"] = $bodyTaskSuccBc
}
}
}

# Get Tape Backup Summary Info
$arrSummaryTp = $null
If ($showSummaryTp) {
$vbrMasterHash = @{
"Sessions" = If ($sessListTp) {@($sessListTp).Count} Else {0}
"Read" = $totalReadTp
"Transferred" = $totalXferTp
"Successful" = @($successSessionsTp).Count
"Warning" = @($warningSessionsTp).Count
"Fails" = @($failsSessionsTp).Count
"Working" = @($workingSessionsTp).Count
"Idle" = @($idleSessionsTp).Count
"Waiting" = @($waitingSessionsTp).Count
}
$vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
If ($onlyLastTp) {
$total = "Jobs Run"
} Else {
$total = "Total Sessions"
}
$arrSummaryTp =  $vbrMasterObj | Select-Object @{Name=$total; Expression = {$_.Sessions}},
@{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
@{Name="Idle"; Expression = {$_.Idle}}, @{Name="Waiting"; Expression = {$_.Waiting}},
@{Name="Working"; Expression = {$_.Working}}, @{Name="Successful"; Expression = {$_.Successful}},
@{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $jsonHash["SummaryTp"] = $arrSummaryTp
}

# Get Tape Backup Job Status
$bodyJobsTp = $null
if ($showJobsTp -and $allJobsTp.Count -gt 0) {
$bodyJobsTp = @()
  foreach ($tpJob in $allJobsTp) {
    $bodyJobsTp += (
      $tpJob | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Job Type"; Expression = {$_.Type.ToString()}},
        @{Name="Media Pool"; Expression = {$_.Target}},
    @{Name="State"; Expression = {$_.LastState.ToString()}},
    @{Name="Next Run"; Expression = {
          try {
            $s = Get-VBRJobScheduleOptions -Job $tpJob
            if ($s.Type -eq "AfterNewBackup") {"Continious"
            } elseif ($s.Type -eq "AfterJob") {"After [$(($allJobs + $allJobsTp | Where-Object {$_.Id -eq $tpJob.ScheduleOptions.JobId}).Name)]"
            } elseif ($tpJob.NextRun) {$tpJob.NextRun.ToString("dd/MM/yyyy HH:mm")} else {"Not Scheduled"}
          } catch {"Unavailable"}
        }},
    @{Name="Status"; Expression = {
          if ($_.LastResult -eq "None") {""} else { $_.LastResult.ToString()}
}}
    )
}
    $bodyJobsTp = $bodyJobsTp | Sort-Object "Next Run", "Job Name"
    $jsonHash["JobsTp"] = $bodyJobsTp
}


# Get Tape Backup Sessions
$arrAllSessTp = $null
If ($showAllSessTp) {
If ($sessListTp.count -gt 0) {
If ($showDetailedTp) {
  $arrAllSessTp = $sessListTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="State"; Expression = {$_.State.ToString()}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
      $jsonHash["AllSessTp"] = $arrAllSessTp
} Else {
  $arrAllSessTp = $sessListTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="State"; Expression = {$_.State.ToString()}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
    $jsonHash["AllSessTp"] = $arrAllSessTp
}

# Due to issue with getting details on tape sessions, we may need to get session info again :-(
If (($showWaitingTp -or $showIdleTp -or $showRunningTp -or $showWarnFailTp -or $showSuccessTp) -and $showDetailedTp) {
  # Get all Tape Backup Sessions
  $allSessTp = @()
  Foreach ($tpJob in $allJobsTp){
    $tpSessions = [veeam.backup.core.cbackupsession]::GetByJob($tpJob.id)
    $allSessTp += $tpSessions
  }
  # Gather all Tape Backup Sessions within timeframe
  $sessListTp = @($allSessTp | Where-Object {$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -match "Working|Idle"})
  If ($null -ne $tapeJob -and $tapeJob -ne "") {
    $allJobsTpTmp = @()
    $sessListTpTmp = @()
    Foreach ($tpJob in $tapeJob) {
      $allJobsTpTmp += $allJobsTp | Where-Object {$_.Name -like $tpJob}
      $sessListTpTmp += $sessListTp | Where-Object {$_.JobName -like $tpJob}
    }
    $allJobsTp = $allJobsTpTmp | Sort-Object Id -Unique
    $sessListTp = $sessListTpTmp | Sort-Object Id -Unique
  }
  If ($onlyLastTp) {
    $tempSessListTp = $sessListTp
    $sessListTp = @()
    Foreach($job in $allJobsTp) {
      $sessListTp += $tempSessListTp | Where-Object {$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
    }
  }
  # Get Tape Backup Session information
  $idleSessionsTp = @($sessListTp | Where-Object {$_.State -eq "Idle"})
  $successSessionsTp = @($sessListTp | Where-Object {$_.Result -eq "Success"})
  $warningSessionsTp = @($sessListTp | Where-Object {$_.Result -eq "Warning"})
  $failsSessionsTp = @($sessListTp | Where-Object {$_.Result -eq "Failed"})
  $workingSessionsTp = @($sessListTp | Where-Object {$_.State -eq "Working"})
  $waitingSessionsTp = @($sessListTp | Where-Object {$_.State -eq "WaitingTape"})
}
}
}

# Get Waiting Tape Backup Jobs
$bodyWaitingTp = $null
If ($showWaitingTp) {
If ($waitingSessionsTp.count -gt 0) {
$bodyWaitingTp = $waitingSessionsTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
  @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
  @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Progress.StartTimeLocal $(Get-Date))}},
  @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
  @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
  @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}}
      $jsonHash["WaitingTp"] = $bodyWaitingTp
}
}

# Get Idle Tape Backup Sessions
$bodySessIdleTp = $null
If ($showIdleTp) {
If ($idleSessionsTp.count -gt 0) {
If ($showDetailedTp) {
  $bodySessIdleTp = $idleSessionsTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}}
        $jsonHash["SessIdleTp"] = $bodySessIdleTp
} Else {
  $bodySessIdleTp = $idleSessionsTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}}
      $jsonHash["SessIdleTp"] = $bodySessIdleTp
}
}
}

# Get Working Tape Backup Jobs
$bodyRunningTp = $null
If ($showRunningTp) {
If ($workingSessionsTp.count -gt 0) {
$bodyRunningTp = $workingSessionsTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
  @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
  @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Progress.StartTimeLocal $(Get-Date))}},
  @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
  @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
  @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}}
    $jsonHash["RunningTp"] = $bodyRunningTp
}
}

# Get Tape Backup Sessions with Warnings or Failures
$arrSessWFTp = $null
If ($showWarnFailTp) {
$sessWF = @($warningSessionsTp + $failsSessionsTp)
If ($sessWF.count -gt 0) {
If ($showDetailedTp) {
  $arrSessWFTp = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
    @{Name="Details"; Expression = {
      If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
      Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
      $jsonHash["SessWFTp"] = $arrSessWFTp
} Else {
  $arrSessWFTp = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {
      If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
      Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
      $jsonHash["SessWFTp"] = $arrSessWFTp
}
}
}

# Get Successful Tape Backup Sessions
$bodySessSuccTp = $null
If ($showSuccessTp) {
If ($successSessionsTp.count -gt 0) {
If ($showDetailedTp) {
  $bodySessSuccTp = $successSessionsTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
    @{Name="Details"; Expression = {
      If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
      Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}},
    @{Name="Result"; Expression = {($_.Result.ToString())}}
        $jsonHash["SessSuccTp"] = $bodySessSuccTp
} Else {
  $bodySessSuccTp = $successSessionsTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {
      If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
      Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}},
    @{Name="Result"; Expression = {($_.Result.ToString())}}
        $jsonHash["SessSuccTp"] = $bodySessSuccTp
}
}
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Tape Backup Tasks from Sessions within time frame
$taskListTp = @()
$taskListTp += $sessListTp | Get-VBRTaskSession
$successTasksTp = @($taskListTp | Where-Object {$_.Status -eq "Success"})
$wfTasksTp = @($taskListTp | Where-Object {$_.Status -match "Warning|Failed"})
$pendingTasksTp = @($taskListTp | Where-Object {$_.Status -eq "Pending"})
$runningTasksTp = @($taskListTp | Where-Object {$_.Status -eq "InProgress"})

# Get Tape Backup Tasks
$bodyAllTasksTp = $null
If ($showAllTasksTp) {
If ($taskListTp.count -gt 0) {
If ($showDetailedTp) {
  $arrAllTasksTp = $taskListTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, 
    @{Name="Status"; Expression = {($_.Status.ToString())}}
      $bodyAllTasksTp = $arrAllTasksTp | Sort-Object "Start Time"
      $jsonHash["AllTasksTp"] = $bodyAllTasksTp
} Else {
  $arrAllTasksTp = $taskListTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, 
    @{Name="Status"; Expression = {($_.Status.ToString())}}
      $bodyAllTasksTp = $arrAllTasksTp | Sort-Object "Start Time"
      $jsonHash["AllTasksTp"] = $bodyAllTasksTp
}
}
}

# Get Pending Tape Backup Tasks
$bodyTasksPendingTp = $null
If ($showPendingTasksTp) {
If ($pendingTasksTp.count -gt 0) {
$bodyTasksPendingTp = $pendingTasksTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort-Object "Start Time"
        $jsonHash["TasksPendingTp"] = $bodyTasksPendingTp
}
}

# Get Working Tape Backup Tasks
$bodyTasksRunningTp = $null
If ($showRunningTasksTp) {
If ($runningTasksTp.count -gt 0) {
$bodyTasksRunningTp = $runningTasksTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort-Object "Start Time"
    $jsonHash["TasksRunningTp"] = $bodyTasksRunningTp
}
}

# Get Tape Backup Tasks with Warnings or Failures
$bodyTaskWFTp = $null
If ($showTaskWFTp) {
If ($wfTasksTp.count -gt 0) {
If ($showDetailedTp) {
  $arrTaskWFTp = $wfTasksTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, 
    @{Name="Status"; Expression = {($_.Status.ToString())}}
      $bodyTaskWFTp = $arrTaskWFTp | Sort-Object "Start Time"
      $jsonHash["TaskWFTp"] = $bodyTaskWFTp
} Else {
  $arrTaskWFTp = $wfTasksTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, 
    @{Name="Status"; Expression = {($_.Status.ToString())}}
      $bodyTaskWFTp = $arrTaskWFTp | Sort-Object "Start Time"
      $jsonHash["TaskWFTp"] = $bodyTaskWFTp
}
}
}

# Get Successful Tape Backup Tasks
$bodyTaskSuccTp = $null
If ($showTaskSuccessTp) {
If ($successTasksTp.count -gt 0) {
If ($showDetailedTp) {
  $bodyTaskSuccTp = $successTasksTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {
      If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
      Else {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}
    }},
    @{Name="Duration (HH:MM:SS)"; Expression = {
      If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
      Else {Get-Duration -ts $_.Progress.Duration}
    }},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
    @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
    @{Name="Status"; Expression = {($_.Status.ToString())}} | Sort-Object "Start Time"
        $jsonHash["TaskSuccTp"] = $bodyTaskSuccTp
} Else {
  $bodyTaskSuccTp = $successTasksTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {
      If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
      Else {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}
    }},
    @{Name="Duration (HH:MM:SS)"; Expression = {
      If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
      Else {Get-Duration -ts $_.Progress.Duration}
    }},
    @{Name="Status"; Expression = {($_.Status.ToString())}} | Sort-Object "Start Time"
        $jsonHash["TaskSuccTp"] = $bodyTaskSuccTp
}
}
}

# Get all Expired Tapes
$expTapes = $null
If ($showExpTp) {
$expTapes = @($mediaTapes | Where-Object {($_.IsExpired -eq $True)})
If ($expTapes.Count -gt 0) {
$expTapes = $expTapes | Select-Object Name, Barcode,
@{Name="Media Pool"; Expression = {
    $poolId = $_.MediaPoolId
    ($mediaPools | Where-Object {$_.Id -eq $poolId}).Name
}},
@{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
@{Name="Location"; Expression = {
    switch ($_.Location) {
      "None" {"Offline"}
      "Slot" {
        $lId = $_.LibraryId
        $lName = $($mediaLibs | Where-Object {$_.Id -eq $lId}).Name
        [int]$slot = $_.SlotAddress + 1
        "{0} : {1} {2}" -f $lName,$_,$slot
      }
      "Drive" {
        $lId = $_.LibraryId
        $dId = $_.DriveId
        $lName = $($mediaLibs | Where-Object {$_.Id -eq $lId}).Name
        $dName = $($mediaDrives | Where-Object {$_.Id -eq $dId}).Name
        [int]$dNum = $_.Location.DriveAddress + 1
        "{0} : {1} {2} (Drive ID: {3})" -f $lName,$_,$dNum,$dName
      }
      "Vault" {
        $vId = $_.VaultId
        $vName = $($mediaVaults | Where-Object {$_.Id -eq $vId}).Name
      "{0}: {1}" -f $_,$vName}
      default {"Lost in Space"}
    }
}},
@{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
@{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
    @{Name="Last Write"; Expression = {$_.LastWriteTime}} | Sort-Object Name 
    $jsonHash["expTapes"] = $expTapes
}
}

# Get Agent Backup Summary Info
$arrSummaryEp = $null
If ($showSummaryEp) {
$vbrEpHash = @{
"Sessions" = If ($sessListEp) {@($sessListEp).Count} Else {0}
"Successful" = @($successSessionsEp).Count
"Warning" = @($warningSessionsEp).Count
"Fails" = @($failsSessionsEp).Count
"Running" = @($runningSessionsEp).Count
}
$vbrEPObj = New-Object -TypeName PSObject -Property $vbrEpHash
If ($onlyLastEp) {
$total = "Jobs Run"
} Else {
$total = "Total Sessions"
}
$arrSummaryEp =  $vbrEPObj | Select-Object @{Name=$total; Expression = {$_.Sessions}},
@{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
@{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $jsonHash["SummaryAg"] = $arrSummaryEp
}

# Get Agent Backup Job Status
$bodyJobsEp = $null
if ($showJobsEp -and $allJobsEp.Count -gt 0) {
  $bodyJobsEp = $allJobsEp | Sort-Object Name | Select-Object
    @{Name="Job Name"; Expression = {$_.Name}},
  @{Name="Description"; Expression = {$_.Description}},
  @{Name="Enabled"; Expression = {$_.JobEnabled}},
  @{Name="State"; Expression = {(Get-VBRComputerBackupJobSession -Name $_.Name)[0].state}},
  @{Name="Target Repo"; Expression = {$_.BackupRepository.Name}},
  @{Name="Next Run"; Expression = {
      try {
        if (-not $_.ScheduleEnabled) { "Not Scheduled" }
        else { (Get-VBRJobScheduleOptions -Job $_).NextRun }
      } catch {"Unavailable"}}},
      @{Name="Status"; Expression = {(Get-VBRComputerBackupJobSession -Name $_.Name)[0].result}}
      $jsonHash["JobsAg"] = $bodyJobsEp
}

# Get Agent Backup Job Size
$JobsSizeAgent = $null
If ($showBackupSizeEp) {
If ($backupsEp.count -gt 0) {
    $JobsSizeAgent = Get-BackupSize -backups $backupsEp | Sort-Object JobName | Select-Object @{Name="Job Name"; Expression = {$_.JobName}},
  @{Name="VM Count"; Expression = {$_.VMCount}},
  @{Name="Repository"; Expression = {$_.Repo}},
      @{Name="Backup Size (GB)"; Expression = {$_.LogSize}}
     $jsonHash["JobsSizeAg"] = $JobsSizeAgent
}
}

# Get Agent Backup Sessions
$bodyAllSessEp = @()
$arrAllSessEp = @()
If ($showAllSessEp) {
If ($sessListEp.count -gt 0) {
Foreach($job in $allJobsEp) {
  $arrAllSessEp += $sessListEp | Where-Object {$_.JobId -eq $job.Id} | Select-Object @{Name="Job Name"; Expression = {$job.Name}},
    @{Name="State"; Expression = {$_.State.ToString()}},@{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {
      If ($_.EndTime -eq "1/1/1900 12:00:00 AM") {
        Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))
      } Else {
        Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)
      }
    }}, @{Name="Result"; Expression = {($_.Result.ToString())}}
}
    $bodyAllSessEp = $arrAllSessEp | Sort-Object "Start Time"
    $jsonHash["AllSessAg"] = $bodyAllSessEp
}
}

# Get Running Agent Backup Jobs
$bodyRunningEp = @()
If ($showRunningEp) {
If ($runningSessionsEp.count -gt 0) {
Foreach($job in $allJobsEp) {
  $bodyRunningEp += $runningSessionsEp | Where-Object {$_.JobId -eq $job.Id} | Select-Object @{Name="Job Name"; Expression = {$job.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}}
}
    $bodyRunningEp = $bodyRunningEp | Sort-Object "Start Time"
    $jsonHash["AllRunningAg"] = $bodyRunningEp
}
}

# Get Agent Backup Sessions with Warnings or Failures
$bodySessWFEp = @()
$arrSessWFEp = @()
If ($showWarnFailEp) {
$sessWFEp = @($warningSessionsEp + $failsSessionsEp)
If ($sessWFEp.count -gt 0) {
Foreach($job in $allJobsEp) {
  $arrSessWFEp += $sessWFEp | Where-Object {$_.JobId -eq $job.Id} | Select-Object @{Name="Job Name"; Expression = {$job.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}}, @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
    @{Name="Result"; Expression = {($_.Result.ToString())}}
}
    $jsonHash["SessWFAg"] = $bodySessWFEp
}
}

# Get Successful Agent Backup Sessions
$bodySessSuccEp = @()
If ($showSuccessEp) {
If ($successSessionsEp.count -gt 0) {
Foreach($job in $allJobsEp) {
  $bodySessSuccEp += $successSessionsEp | Where-Object {$_.JobId -eq $job.Id} | Select-Object @{Name="Job Name"; Expression = {$job.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}}, @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
    @{Name="Result"; Expression = {($_.Result.ToString())}}
}
    $bodySessSuccEp = $bodySessSuccEp | Sort-Object "Start Time"
    $jsonHash["SessSuccAg"] = $bodySessSuccEp
}
}

# Get SureBackup Summary Info
$arrSummarySb = $null
If ($showSummarySb) {
$vbrMasterHash = @{
"Sessions" = If ($sessListSb) {@($sessListSb).Count} Else {0}
"Successful" = @($successSessionsSb).Count
"Warning" = @($warningSessionsSb).Count
"Fails" = @($failsSessionsSb).Count
"Running" = @($runningSessionsSb).Count
}
$vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
If ($onlyLastSb) {
$total = "Jobs Run"
} Else {
$total = "Total Sessions"
}
$arrSummarySb =  $vbrMasterObj | Select-Object @{Name=$total; Expression = {$_.Sessions}},
@{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
@{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $jsonHash["SummarySb"] = $arrSummarySb
}

# Get SureBackup Job Status
$bodyJobsSb = $null
if ($showJobsSb -and $allJobsSb.Count -gt 0) {
$bodyJobsSb = @()
  foreach ($SbJob in $allJobsSb) {
    $bodyJobsSb += $SbJob | Select-Object @{Name = "Job Name"; Expression = { $_.Name }},
      @{Name = "Enabled"; Expression = { $_.IsEnabled }},
      @{Name = "State"; Expression = {
        if ($_.LastState -eq "Working") {$currentSess = $_.FindLastSession()
          "$($currentSess.CompletionPercentage)% completed"
        } else {$_.LastState.ToString()}}},
      @{Name = "Virtual Lab"; Expression = {$_.VirtualLab}},
      @{Name = "Linked Jobs"; Expression = {$_.LinkedJob}},
      @{Name = "Next Run"; Expression = {
        try {
          if (-not $_.ScheduleEnabled) { "Disabled" }
          else {$_.NextRun.ToString("dd/MM/yyyy HH:mm")}
        } catch {"Unavailable"}
      }},
      @{Name = "Last Result"; Expression = {$_.LastResult.ToString()}}}
    $bodyJobsSb = $bodyJobsSb | Sort-Object "Next Run" 
    $jsonHash["JobsSb"] = $bodyJobsSb
}

# Get SureBackup Sessions
$arrAllSessSb = $null
If ($showAllSessSb) {
If ($sessListSb.count -gt 0) {
$arrAllSessSb = $sessListSb | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="State"; Expression = {$_.State.ToString()}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}}},

    @{Name="Duration (HH:MM:SS)"; Expression = {
      If ($_.EndTime -eq "1/1/1900 12:00:00 AM") {
        Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))
      } Else {
        Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)
      }
    }}, @{Name="Result"; Expression = {($_.Result.ToString())}}
    $jsonHash["AllSessSb"] = $arrAllSessSb
}
}

# Get Running SureBackup Jobs
$runningSessionsSb = $null
If ($showRunningSb) {
If ($runningSessionsSb.count -gt 0) {
    $runningSessionsSb = $runningSessionsSb | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
  @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
  @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}},
  @{Name="% Complete"; Expression = {$_.Progress}}
  $jsonHash["SessionsSb"] = $runningSessionsSb
}
}

# Get SureBackup Sessions with Warnings or Failures
$arrSessWFSb = $null
If ($showWarnFailSb) {
$sessWF = @($warningSessionsSb + $failsSessionsSb)
If ($sessWF.count -gt 0) {
$arrSessWFSb = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
    $jsonHash["SessWFSb"] = $arrSessWFSb
}
}

# Get Successful SureBackup Sessions
$bodySessSuccSb = $null
If ($showSuccessSb) {
If ($successSessionsSb.count -gt 0) {
$bodySessSuccSb = $successSessionsSb | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
        @{Name="Result"; Expression = {($_.Result.ToString())}}
        $jsonHash["SessSuccSb"] = $bodySessSuccSb
}
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all SureBackup Tasks from Sessions within time frame
If ($showRunningSb) {
  $taskListSb = @()
  $taskListSb += $sessListSb | Get-VSBTaskSession
  $successTasksSb = @($taskListSb | Where-Object {$_.Info.Result -eq "Success"})
  $wfTasksSb = @($taskListSb | Where-Object {$_.Info.Result -match "Warning|Failed"})
  $runningTasksSb = @()
  $runningTasksSb += $runningSessionsSb | Get-VSBTaskSession | Where-Object {$_.Status -ne "Stopped"}
}

# Get SureBackup Tasks
$bodyAllTasksSb = $null
If ($showAllTasksSb) {
If ($taskListSb.count -gt 0) {
$arrAllTasksSb = $taskListSb | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
  @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
  @{Name="Status"; Expression = {$_.Status}},
  @{Name="Start Time"; Expression = {$_.Info.StartTime}},
  @{Name="Stop Time"; Expression = {If ($_.Info.FinishTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Info.FinishTime}}},
  @{Name="Duration (HH:MM:SS)"; Expression = {
    If ($_.Info.FinishTime -eq "1/1/1900 12:00:00 AM") {
      Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $(Get-Date))
    } Else {
      Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $_.Info.FinishTime)
    }
  }},
  @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
  @{Name="Ping Test"; Expression = {$_.PingStatus}},
  @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
  @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
  @{Name="Result"; Expression = {
      If ($_.Info.Result -eq "notrunning") {
        "None"
      } Else {
        $_.Info.Result
      }
  }}
    $bodyAllTasksSb = $arrAllTasksSb | Sort-Object "Start Time"
    $jsonHash["AllTasksSb"] = $bodyAllTasksSb
}
}

# Get Running SureBackup Tasks
$bodyTasksRunningSb = $null
If ($showRunningTasksSb) {
If ($runningTasksSb.count -gt 0) {
$bodyTasksRunningSb = $runningTasksSb | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
  @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
  @{Name="Start Time"; Expression = {$_.Info.StartTime}},
  @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $(Get-Date))}},
  @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
  @{Name="Ping Test"; Expression = {$_.PingStatus}},
  @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
  @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
      Status | Sort-Object "Start Time"
      $jsonHash["TasksRunningSb"] = $bodyTasksRunningSb
}
}

# Get SureBackup Tasks with Warnings or Failures
$bodyTaskWFSb = $null
If ($showTaskWFSb) {
If ($wfTasksSb.count -gt 0) {
$arrTaskWFSb = $wfTasksSb | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
  @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
  @{Name="Start Time"; Expression = {$_.Info.StartTime}},
  @{Name="Stop Time"; Expression = {$_.Info.FinishTime}},
  @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $_.Info.FinishTime)}},
  @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
  @{Name="Ping Test"; Expression = {$_.PingStatus}},
  @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
  @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
  @{Name="Result"; Expression = {$_.Info.Result}}
    $bodyTaskWFSb = $arrTaskWFSb | Sort-Object "Start Time"
    $jsonHash["TaskWFSb"] = $bodyTaskWFSb
}
}

# Get Successful SureBackup Tasks
$bodyTaskSuccSb = $null
If ($showTaskSuccessSb) {
If ($successTasksSb.count -gt 0) {
$bodyTaskSuccSb = $successTasksSb | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
  @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
  @{Name="Start Time"; Expression = {$_.Info.StartTime}},
  @{Name="Stop Time"; Expression = {$_.Info.FinishTime}},
  @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $_.Info.FinishTime)}},
  @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
  @{Name="Ping Test"; Expression = {$_.PingStatus}},
  @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
  @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
      @{Name="Result"; Expression = {$_.Info.Result}} | Sort-Object "Start Time"
    $jsonHash["TaskSuccSb"] = $bodyTaskSuccSb
}
}

# Get Configuration Backup Summary Info
$bodySummaryConfig = $null
If ($showSummaryConfig) {
$vbrConfigHash = @{
  "Enabled" = $configBackup.Enabled
  "State" = $configBackup.LastState.ToString()
  "Target" = $configBackup.Target
  "Schedule" = $configBackup.ScheduleOptions.ToString()
  "Restore Points" = $configBackup.RestorePointsToKeep
  "Encrypted" = $configBackup.EncryptionOptions.Enabled
  "Status" = $configBackup.LastResult.ToString()
  "Next Run" = $configBackup.NextRun.ToString()
}
$vbrConfigObj = New-Object -TypeName PSObject -Property $vbrConfigHash
$bodySummaryConfig = $vbrConfigObj | Select-Object Enabled, State, Target, Schedule, "Restore Points", "Next Run", Encrypted, "Status"
$jsonHash["SummaryConfig"] = $bodySummaryConfig
}

# Get Proxy Info
$bodyProxy = $null
If ($showProxy) {
If ($proxyList.count -gt 0) {
$arrProxy = $proxyList | Get-VBRProxyInfo | Select-Object @{Name="Proxy Name"; Expression = {$_.ProxyName}},
  @{Name="Transport Mode"; Expression = {$_.tMode}}, @{Name="Max Tasks"; Expression = {$_.MaxTasks}},
  @{Name="Proxy Host"; Expression = {$_.RealName}}, @{Name="Host Type"; Expression = {$_.pType.ToString()}},
  @{Name = "Enabled"; Expression = { $_.Enabled }}, @{Name="IP Address"; Expression = {$_.IP}},
  @{Name="RT (ms)"; Expression = {$_.Response}}, @{Name="Status"; Expression = {($_.Status.ToString())}}
    $bodyProxy = $arrProxy | Sort-Object "Proxy Host"
    $jsonHash["Proxy"] = $bodyProxy
}
}

# Get Repository Info
$bodyRepo = $null
If ($showRepo) {
If ($repoList.count -gt 0) {
$arrRepo = $repoList | Get-VBRRepoInfo | Select-Object @{Name="Repository Name"; Expression = {$_.Target}},
  @{Name="Type"; Expression = {$_.rType}},
  @{Name="Max Tasks"; Expression = {$_.MaxTasks}},
  @{Name="Host"; Expression = {$_.RepoHost}},
  @{Name="Path"; Expression = {$_.Storepath}},
  @{Name="Backups (GB)"; Expression = {$_.StorageBackup}},
  @{Name="Other data (GB)"; Expression = {$_.StorageOther}},
  @{Name="Free (GB)"; Expression = {$_.StorageFree}},
  @{Name="Total (GB)"; Expression = {$_.StorageTotal}},
  @{Name="Free (%)"; Expression = {$_.FreePercentage}},
  @{Name="Status"; Expression = {
    If ($_.FreePercentage -lt $repoCritical) {"Critical"}
    ElseIf ($_.StorageTotal -eq 0 -and $_.rtype -ne "SAN Snapshot")  {"Warning"}
    ElseIf ($_.StorageTotal -eq 0) {"NoData"}
    ElseIf ($_.FreePercentage -lt $repoWarn) {"Warning"}
    ElseIf ($_.FreePercentage -eq "Unknown") {"Unknown"}
    Else {"OK"}}
  }
    $bodyRepo = $arrRepo | Sort-Object "Repository Name"
    $jsonHash["Repo"] = $bodyRepo
}
}
# Get Scale Out Repository Info
$bodySORepo = $null
If ($showRepo) {
If ($repoListSo.count -gt 0) {
$arrSORepo = $repoListSo | Get-VBRSORepoInfo | Select-Object @{Name="Scale Out Repository Name"; Expression = {$_.SOTarget}},
  @{Name="Member Name"; Expression = {$_.Target}},
@{Name="Type"; Expression = {$_.rType}},
  @{Name="Max Tasks"; Expression = {$_.MaxTasks}},
@{Name="Host"; Expression = {$_.RepoHost}},
  @{Name="Path"; Expression = {$_.Storepath}},
@{Name="Free (GB)"; Expression = {$_.StorageFree}},
  @{Name="Total (GB)"; Expression = {$_.StorageTotal}},
@{Name="Free (%)"; Expression = {$_.FreePercentage}},
  @{Name="Status"; Expression = {
    If ($_.FreePercentage -lt $repoCritical) {"Critical"}
    ElseIf ($_.StorageTotal -eq 0)  {"Warning"}
    ElseIf ($_.FreePercentage -lt $repoWarn) {"Warning"}
    ElseIf ($_.FreePercentage -eq "Unknown") {"Unknown"}
    Else {"OK"}}

  }
    $bodySORepo = $arrSORepo | Sort-Object "Scale Out Repository Name", "Member Repository Name"
    $jsonHash["SORepo"] = $bodySORepo
}
}

# Get Replica Target Info
$repTargets = $null
If ($showReplicaTarget) {
If ($allJobsRp.count -gt 0) {
$repTargets = $allJobsRp | Get-VBRReplicaTarget | Select-Object @{Name="Replica Target"; Expression = {$_.Target}}, Datastore,
  @{Name="Free (GB)"; Expression = {$_.StorageFree}}, @{Name="Total (GB)"; Expression = {$_.StorageTotal}},
  @{Name="Free (%)"; Expression = {$_.FreePercentage}},
  @{Name="Status"; Expression = {
    If ($_.FreePercentage -lt $replicaCritical) {"Critical"}
    ElseIf ($_.StorageTotal -eq 0)  {"Warning"}
    ElseIf ($_.FreePercentage -lt $replicaWarn) {"Warning"}
    ElseIf ($_.FreePercentage -eq "Unknown") {"Unknown"}
    Else {"OK"}
    }
  } | Sort-Object "Replica Target"
    $jsonHash["repTargets"] = $repTargets
}
}

#region license info
# Get License Info
$arrLicense = $null
If ($showLicExp) {
  $arrLicense = Get-VeeamSupportDate $vbrServer | Select-Object @{Name = "Type"; Expression = { $_.LicType.ToString() } },
@{Name="Expiry Date"; Expression = {$_.ExpDate}},
@{Name="Days Remaining"; Expression = {$_.DaysRemain}}, `
@{Name="Status"; Expression = {
  If ($_.LicType -eq "Evaluation") {"OK"}
  ElseIf ($_.DaysRemain -lt $licenseCritical) {"Critical"}
  ElseIf ($_.DaysRemain -lt $licenseWarn) {"Warning"}
  ElseIf ($_.DaysRemain -eq "Failed") {"Failed"}
  Else {"OK"}}
}
  $jsonHash["License"] = $arrLicense
}
#endregion

#region JSON Output
$jsonOutput = $jsonHash | ConvertTo-Json
If ($saveJSON) {
  $jsonOutput | Out-File $pathJSON -Encoding UTF8
If ($launchJSON) {
Invoke-Item $pathJSON
}
}
#endregion

#region Output
# Send Report via Email
$smtp = New-Object System.Net.Mail.SmtpClient($emailHost, $emailPort)
$smtp.Credentials = New-Object System.Net.NetworkCredential($emailUser, $emailPass)
$smtp.EnableSsl = $emailEnableSSL
$msg = New-Object System.Net.Mail.MailMessage($emailFrom, $emailTo)
$msg.Subject = $emailSubject
$attachment = New-Object System.Net.Mail.Attachment $pathJSON
$msg.Attachments.Add($attachment)
$body = @"
Bonjour,

Vous trouverez en pièce jointe le dernier rapport de sauvegarde Veeam au format JSON.

Cordialement,
"@
$msg.Body = $body
$msg.IsBodyHtml = $false
$smtp.Send($msg)
#endregion

#region purge

Get-childitem -path $pathJSON -Recurse | where-object {($_.LastWriteTime -lt (get-date).adddays(-$JPurge))} | Remove-Item

#endregion
