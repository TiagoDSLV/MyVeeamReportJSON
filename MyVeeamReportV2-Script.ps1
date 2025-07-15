<#====================================================================
Author        : Tiago DA SILVA - ATHEO INGENIERIE
Version       : 1.0.1
Creation Date : 2025-07-01
Last Update   : 2025-07-01
GitHub Repo   : https://github.com/TiagoDSLV/MyVeeamReportV2/
====================================================================

DESCRIPTION:
My Veeam Report is a flexible reporting script for Veeam Backup and
Replication. This report can be customized to report on Backup, Replication,
Backup Copy, Tape Backup, SureBackup and Agent Backup jobs as well as
infrastructure details like repositories, proxies and license status. 

====================================================================#>

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

$OutputPath = ".\MyVeeamReportV2-Script.ps1"
$FileURL = "https://raw.githubusercontent.com/TiagoDSLV/MyVeeamReportV2/refs/heads/main/MyVeeamReportV2-Script.ps1"

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
$sessListBc += $tempSessListBc | Where-Object {($_.JobName -split '\\')[0] -eq $job.Name -and $_.BaseProgress -eq 100} | Sort-Object EndTime -Descending | Select-Object -First 1
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

# Format Report Mode for header
If (($reportMode -ne "Weekly") -And ($reportMode -ne "Monthly")) {
  $rptMode = "RPO: $reportMode Hrs"
} Else {
  $rptMode = "RPO: $reportMode"
}

# Toggle VBR Server name in report header
If ($showVBR) {
  $vbrName = "VBR Server - $vbrServer"
} Else {
  $vbrName = $null
}

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

# Création d'un hashtable
$jsonHash = [ordered]@{}
$jsonHash.Add("reportVersion", $reportVersion)
$jsonHash.Add("generationDate", $(Get-Date -format g))
$jsonHash.Add("client", $Client)
$jsonHash.Add("mailGLPI", $MailGLPI)
$jsonHash.Add("reportMode", $reportMode)
$jsonHash.Add("vbrServerName", $vbrServer)
$jsonHash.Add("vbrServerVersion", $objectVersion)

# HTML Stuff
$headerObj = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <title>$rptTitle</title>
            <style type="text/css">
              body {font-family: Tahoma; background-color: #ffffff;}
              table {font-family: Tahoma; width: $($rptWidth)%; font-size: 12px; border-collapse: collapse; margin-left: auto; margin-right: auto;}
              table tr:nth-child(odd) td {background: $oddColor;}
              th {background-color: #e2e2e2; border: 1px solid #a7a9ac;border-bottom: none;}
              td {background-color: #ffffff; border: 1px solid #a7a9ac;padding: 2px 3px 2px 3px;}
            </style>
    </head>
"@

$bodyTop = @"
    <body>
          <table>
              <tr>
                  <td style="width: 50%;height: 14px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 10px;vertical-align: bottom;text-align: left;padding: 2px 0px 0px 5px;"></td>
                  <td style="width: 50%;height: 14px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 2px 5px 0px 0px;">Report generated on $(Get-Date -format g)</td>
              </tr>
              <tr>
                  <td style="width: 50%;height: 24px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 24px;vertical-align: bottom;text-align: left;padding: 0px 0px 0px 15px;">$rptTitle</td>
                  <td style="width: 50%;height: 24px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 0px 5px 2px 0px;">$vbrName</td>
              </tr>
              <tr>
                  <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: left;padding: 0px 0px 0px 5px;"></td>
                  <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 0px 5px 0px 0px;">VBR v$objectVersion</td>
              </tr>
              <tr>
                  <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: left;padding: 0px 0px 2px 5px;">$rptMode</td>
                  <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 0px 5px 2px 0px;">MVR v$MVRversion</td>
              </tr>
          </table>
"@

$subHead01 = @"
<table>
                <tr>
                    <td style="height: 35px;background-color: #f3f4f4;color: #626365;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead01suc = @"
<table>
                 <tr>
                    <td style="height: 35px;background-color: #00b050;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead01war = @"
<table>
                 <tr>
                    <td style="height: 35px;background-color: #ffd96c;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead01err = @"
<table>
                <tr>
                    <td style="height: 35px;background-color: #FB9895;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead01inf = @"
<table>
                <tr>
                    <td style="height: 35px;background-color: #3399FF;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead02 = @"
</td>
                </tr>
             </table>
"@

$HTMLbreak = @"
<table>
                <tr>
                    <td style="height: 10px;background-color: #626365;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;"></td>
						    </tr>
            </table>
"@

$footerObj = @"
            <table>
                <tr>
                </tr>
            </table>
    </body>
</html>
"@
#endregion

#region Get VM Backup Status
$vmStatus = @()
If ($showSummaryProtect + $showUnprotectedVMs + $showUnprotectedVMsInfo + $showProtectedVMs) {
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
$bodySummaryProtect = $null
$sumprotectHead = $subHead01
$percentProt = 0

If ($showSummaryProtect) {
  if (@($successVMs).Count -ge 1) {
    $percentProt = 1
    $sumprotectHead = $subHead01suc
  }
  if (@($warnVMs).Count -ge 1) {
    $sumprotectHead = $subHead01war
  }
  if (@($missingVMs).Count -ge 1) {
    $totalVMs = @($warnVMs).Count + @($successVMs).Count + @($missingVMs).Count
    if ($totalVMs -gt 0) {
      $percentProt = (@($warnVMs).Count + @($successVMs).Count) / $totalVMs
    }
    $sumprotectHead = if ($showUnprotectedVMsInfo) { $subHead01inf } else { $subHead01err }
  }
}

$vbrMasterHash = @{
  WarningVM = @($warnVMs).Count
  ProtectedVM = @($successVMs).Count
  UnprotectedVM = @($missingVMs).Count
  ExcludedVM = ($allVMsVBRVi).Count - ($vmStatus).Count
  PercentProtected = [Math]::Floor($percentProt * 100)
}

$jsonHash["VMBackupProtectionSummary"] = $vbrMasterHash

$vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
$summaryProtect = $vbrMasterObj | Select-Object @{Name="% Protected"; Expression = {$_.PercentProtected}},
  @{Name="Fully Protected VMs"; Expression = {$_.ProtectedVM}},
  @{Name="Protected VMs w/Warnings"; Expression = {$_.WarningVM}},
  @{Name="Unprotected VMs"; Expression = {$_.UnprotectedVM}}

$bodySummaryProtect = $summaryProtect | ConvertTo-HTML -Fragment
$bodySummaryProtect = $sumprotectHead + "VM Backup Protection Summary" + $subHead02 + $bodySummaryProtect
#endregion

#region Get VMs Missing Backups
$bodyMissing = $null
If ($showUnprotectedVMs -Or $showUnprotectedVMsInfo) {
  If ($missingVMs.Count -gt 0) {
    $missingVMs = $missingVMs | Sort-Object vCenter, Datacenter, Cluster, Name | Select-Object Name, vCenter, Datacenter, Cluster, Folder, Tags,
      @{Name = "StartTime"; Expression = { $_.StartTime.ToString("dd/MM/yyyy HH:mm") }},
      @{Name = "StopTime"; Expression = { $_.StopTime.ToString("dd/MM/yyyy HH:mm") }},
      @{Name = "Details"; Expression = { $_.Details }}
    $jsonHash["missingVms"] = $missingVMs
    $missingVMs = $missingVMs | ConvertTo-HTML -Fragment
    if ($showUnprotectedVMsInfo) {
      $bodyMissing = $subHead01inf + "Unprotected VMs within RPO" + $subHead02 + $missingVMs
    } else {
      $bodyMissing = $subHead01err + "VMs with No Successful Backups within RPO" + $subHead02 + $missingVMs
    }
  }
}
#endregion

#region Get VMs Backed Up w/Warnings
$bodyWarning = $null
If ($showProtectedVMs) {
  If ($warnVMs.Count -gt 0) {
  $warnVMs = $warnVMs | Sort-Object vCenter, Datacenter, Cluster, Name | ForEach-Object {
    $_ | Select-Object Name, vCenter, Datacenter, Cluster, Folder, Tags,
        @{Name="StartTime"; Expression = { $_.StartTime.ToString("dd/MM/yyyy HH:mm") }},
        @{Name="StopTime"; Expression = { $_.StopTime.ToString("dd/MM/yyyy HH:mm") }}}
    $jsonHash["warnVMs"] = $warnVMs
    $warnVMs = $warnVMs | ConvertTo-HTML -Fragment
    $bodyWarning = $subHead01war + "VMs with only Backups with Warnings within RPO" + $subHead02 + $warnVMs
  }
}
#endregion

# Get VMs Successfully Backed Up
$bodySuccess = $null
If ($showProtectedVMs) {
  If ($successVMs.Count -gt 0) {
    $successVMs = $successVMs | Sort-Object vCenter, Datacenter, Cluster, Name | ForEach-Object {
      $_ | Select-Object Name, vCenter, Datacenter, Cluster, Folder, Tags,
        @{Name = "StartTime"; Expression = { $_.StartTime.ToString("dd/MM/yyyy HH:mm") } },
        @{Name = "StopTime";  Expression = { $_.StopTime.ToString("dd/MM/yyyy HH:mm") } }
    }

    $jsonHash["successVMs"] = $successVMs
    $successVMs = $successVMs | ConvertTo-HTML -Fragment
    $bodySuccess = $subHead01suc + "VMs with Successful Backups within RPO" + $subHead02 + $successVMs
  }
}
#endregion


# Get VMs Backed Up by Multiple Jobs
$bodyMultiJobs = $null
If ($showMultiJobs) {
$multiJobs = @(Get-MultiJob)
If ($multiJobs.Count -gt 0) {
    $multiJobs = $multiJobs | Sort-Object vCenter, Datacenter, Cluster, Name | Select-Object Name, vCenter, Datacenter, Cluster, Folder,
      @{Name="Job Name"; Expression = {$_.JobName}}
    $jsonHash["multiJobs"] = $multiJobs
    $bodyMultiJobs = $multiJobs | ConvertTo-HTML -Fragment
    $bodyMultiJobs = $subHead01err + "VMs Backed Up by Multiple Jobs within RPO" + $subHead02 + $bodyMultiJobs
}
}

# Get Backup Summary Info
$bodySummaryBk = $null
If ($showSummaryBk) {
  $vbrMasterHash = @{
    "Failed"     = @($failedSessionsBk).Count
    "Sessions"   = If ($sessListBk) { @($sessListBk).Count } Else { 0 }
    "Read"       = $totalReadBk
    "Transferred"= $totalXferBk
    "Successful" = @($successSessionsBk).Count
    "Warning"    = @($warningSessionsBk).Count
    "Fails"      = @($failsSessionsBk).Count
    "Running"    = @($runningSessionsBk).Count
  }

  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastBk) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }

  $arrSummaryBk = $vbrMasterObj | Select-Object `
    @{Name = $total; Expression = { $_.Sessions } },
    @{Name = "Read (GB)"; Expression = { $_.Read } },
    @{Name = "Transferred (GB)"; Expression = { $_.Transferred } },
    @{Name = "Running"; Expression = { $_.Running } },
    @{Name = "Successful"; Expression = { $_.Successful } },
    @{Name = "Warnings"; Expression = { $_.Warning } },
    @{Name = "Failures"; Expression = { $_.Fails } },
    @{Name = "Failed"; Expression = { $_.Failed } }

  $jsonHash["SummaryBk"] = $arrSummaryBk
  $bodySummaryBk = $arrSummaryBk | ConvertTo-HTML -Fragment

  If ($arrSummaryBk.Failed -gt 0) {
    $summaryBkHead = $subHead01err
  } ElseIf ($arrSummaryBk.Warnings -gt 0) {
    $summaryBkHead = $subHead01war
  } ElseIf ($arrSummaryBk.Successful -gt 0) {
    $summaryBkHead = $subHead01suc
  } Else {
    $summaryBkHead = $subHead01
  }

  $bodySummaryBk = $summaryBkHead + "Backup Results Summary" + $subHead02 + $bodySummaryBk
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
  $bodyJobsBk = $bodyJobsBk | Sort-Object "Next Run" | ConvertTo-HTML -Fragment
  $bodyJobsBk = $subHead01 + "Backup Job Status" + $subHead02 + $bodyJobsBk
}

# Get Backup Job Status Begin
$bodyFileJobsBk = $null
if ($showFileJobsBk -and $allFileJobsBk.Count -gt 0) {
    $bodyFileJobsBk = @()
    foreach ($bkJob in $allFileJobsBk) {
        $bodyFileJobsBk += ($bkJob | Select-Object `
            @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Enabled"; Expression = { $_.IsScheduleEnabled } },
            @{Name = "Status"; Expression = {
                if ($bkJob.IsRunning) {
                    $s = $runningSessionsBk | Where-Object { $_.JobName -eq $bkJob.Name }
                    if ($s) {
                        "$($s.Progress.Percents)% completed at $([Math]::Round($s.Progress.AvgSpeed / 1MB, 2)) MB/s"
                    } else {
                        "Running (no session info)"
                    }
                } else {
                    "Stopped"
                }
            }},
            @{Name = "Target Repo"; Expression = {
                ($repoList + $repoListSo | Where-Object { $_.Id -eq $bkJob.Info.TargetRepositoryId }).Name
            }},
            @{Name = "Next Run"; Expression = {
                try {
                    $s = Get-VBRJobScheduleOptions -Job $bkJob
                    if (-not $bkJob.IsScheduleEnabled) {
                        "Disabled"
                    } elseif ($s.RunManually) {
                        "Not Scheduled"
                    } elseif ($s.IsContinious) {
                        "Continious"
                    } elseif ($s.OptionsScheduleAfterJob.IsEnabled) {
                        "After [$(($allJobs + $allJobsTp | Where-Object { $_.Id -eq $bkJob.Info.ParentScheduleId }).Name)]"
                    } else {
                        $s.NextRun
                    }
                } catch {
                    "Unavailable"
                }
            }},
            @{Name = "Status"; Expression = {
                if ($_.Info.LatestStatus -eq "None") { "Unknown" } else { $_.Info.LatestStatus.ToString() }
            }}
        )
    }
    $jsonHash["FileJobsBk"] = $bodyFileJobsBk
    $bodyFileJobsBk = $bodyFileJobsBk | Sort-Object "Next Run" | ConvertTo-HTML -Fragment
    $bodyFileJobsBk = $subHead01 + "File Backup Job Status" + $subHead02 + $bodyFileJobsBk
}
# Get File Backup Job Status End

# Get all Backup Sessions
$bodyAllSessBk = $null
if ($showAllSessBk) {
    if ($sessListBk.Count -gt 0) {
        if ($showDetailedBk) {
            $arrAllSessBk = $sessListBk | Sort-Object CreationTime | Select-Object `
                @{Name = "Job Name"; Expression = { $_.Name } },
                @{Name = "State"; Expression = { $_.State.ToString() } },
                @{Name = "Start Time"; Expression = { $_.CreationTime.ToString("dd/MM/yyyy HH:mm") } },
                @{Name = "Stop Time"; Expression = { if ($_.EndTime -eq "1/1/1900 12:00:00 AM") { "-" } else { $_.EndTime.ToString("dd/MM/yyyy HH:mm") } } },
                @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
                @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
                @{Name = "Total (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedSize / 1GB, 2) } },
                @{Name = "Processed (GB)"; Expression = { [Math]::Round($_.Progress.ProcessedUsedSize / 1GB, 2) } },
                @{Name = "Data Read (GB)"; Expression = { [Math]::Round($_.Progress.ReadSize / 1GB, 2) } },
                @{Name = "Transferred (GB)"; Expression = { [Math]::Round($_.Progress.TransferedSize / 1GB, 2) } },
                @{Name = "Dedupe"; Expression = { if ($_.Progress.ReadSize -eq 0) { 0 } else { ([string][Math]::Round($_.BackupStats.GetDedupeX(), 1)) + "x" } } },
                @{Name = "Compression"; Expression = { if ($_.Progress.ReadSize -eq 0) { 0 } else { ([string][Math]::Round($_.BackupStats.GetCompressX(), 1)) + "x" } } },
                @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } },
                @{Name = "Result"; Expression = { $_.Result.ToString() } }
        } else {
            $arrAllSessBk = $sessListBk | Sort-Object CreationTime | Select-Object `
                @{Name = "Job Name"; Expression = { $_.Name } },
                @{Name = "State"; Expression = { $_.State.ToString() } },
                @{Name = "Start Time"; Expression = { $_.CreationTime.ToString("dd/MM/yyyy HH:mm") } },
                @{Name = "Stop Time"; Expression = { if ($_.EndTime -eq "1/1/1900 12:00:00 AM") { "-" } else { $_.EndTime.ToString("dd/MM/yyyy HH:mm") } } },
                @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
                @{Name = "Details"; Expression = { ($_.GetDetails()).Replace("<br />", "ZZbrZZ") } },
                @{Name = "Result"; Expression = { $_.Result.ToString() } }
        }
        $jsonHash["AllSessBk"] = $arrAllSessBk
        $bodyAllSessBk = $arrAllSessBk | ConvertTo-Html -Fragment
        if ($arrAllSessBk.Result -match "Failed") {
            $allSessBkHead = $subHead01err
        } elseif ($arrAllSessBk.Result -match "Warning") {
            $allSessBkHead = $subHead01war
        } elseif ($arrAllSessBk.Result -match "Success") {
            $allSessBkHead = $subHead01suc
        } else {
            $allSessBkHead = $subHead01
        }
        $bodyAllSessBk = $allSessBkHead + "Backup Sessions" + $subHead02 + $bodyAllSessBk
    }
}


# Get Running Backup Jobs
$bodyRunningBk = $null
if ($showRunningBk) {
    if ($runningSessionsBk.Count -gt 0) {
        $bodyRunningBk = $runningSessionsBk | Sort-Object CreationTime | Select-Object `
            @{Name = "Job Name"; Expression = { $_.Name } },
            @{Name = "Start Time"; Expression = { $_.CreationTime.ToString("dd/MM/yyyy HH:mm") } },
            @{Name = "Duration (HH:MM:SS)"; Expression = { Get-Duration -ts $_.Progress.Duration } },
            @{Name = "Avg Speed (MB/s)"; Expression = { [Math]::Round($_.Progress.AvgSpeed / 1MB, 2) } },
            @{Name = "Read (GB)"; Expression = { [Math]::Round([Decimal]$_.Progress.ReadSize / 1GB, 2) } },
            @{Name = "Transferred (GB)"; Expression = { [Math]::Round([Decimal]$_.Progress.TransferedSize / 1GB, 2) } },
            @{Name = "% Complete"; Expression = { $_.Progress.Percents } }
        $jsonHash["RunningBk"] = $bodyRunningBk
        $bodyRunningBk = $bodyRunningBk | ConvertTo-Html -Fragment
        $bodyRunningBk = $subHead01 + "Running Backup Jobs" + $subHead02 + $bodyRunningBk
    }
}

# Get Backup Sessions with Warnings or Failures
$bodySessWFBk = $null
If ($showWarnFailBk) {
$sessWF = @($warningSessionsBk + $failsSessionsBk)
If ($sessWF.count -gt 0) {
      If ($onlyLastBk) {
      $headerWF = "Backup Jobs with Warnings or Failures"
    } Else {
      $headerWF = "Backup Sessions with Warnings or Failures"
    }
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
    $bodySessWFBk = $arrSessWFBk | ConvertTo-HTML -Fragment
      If ($arrSessWFBk.Result -match "Failed") {
        $sessWFBkHead = $subHead01err
      } ElseIf ($arrSessWFBk.Result -match "Warning") {
        $sessWFBkHead = $subHead01war
      } ElseIf ($arrSessWFBk.Result -match "Success") {
        $sessWFBkHead = $subHead01suc
      } Else {
        $sessWFBkHead = $subHead01
      }
      $bodySessWFBk = $sessWFBkHead + $headerWF + $subHead02 + $bodySessWFBk
} Else {
  $arrSessWFBk = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {
      If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
      Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
    $jsonHash["SessWFBk"] = $arrSessWFBk
    $bodySessWFBk = $arrSessWFBk | ConvertTo-HTML -Fragment
      If ($arrSessWFBk.Result -match "Failed") {
        $sessWFBkHead = $subHead01err
      } ElseIf ($arrSessWFBk.Result -match "Warning") {
        $sessWFBkHead = $subHead01war
      } ElseIf ($arrSessWFBk.Result -match "Success") {
        $sessWFBkHead = $subHead01suc
      } Else {
        $sessWFBkHead = $subHead01
      }
      $bodySessWFBk = $sessWFBkHead + $headerWF + $subHead02 + $bodySessWFBk
  }
}
}

# Get Successful Backup Sessions
$bodySessSuccBk = $null
If ($showSuccessBk) {
If ($successSessionsBk.count -gt 0) {
      If ($onlyLastBk) {
      $headerSucc = "Successful Backup Jobs"
    } Else {
      $headerSucc = "Successful Backup Sessions"
    }
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
    $bodySessSuccBk = $bodySessSuccBk | ConvertTo-HTML -Fragment
    $bodySessSuccBk = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBk
} Else {
  $bodySessSuccBk = $successSessionsBk | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},@{Name="Result"; Expression = {($_.Result.ToString())}}
    $jsonHash["SessSuccBk"] = $bodySessSuccBk
    $bodySessSuccBk = $bodySessSuccBk | ConvertTo-HTML -Fragment
    $bodySessSuccBk = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBk
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
      $arrAllTasksBk  = $arrAllTasksBk | Sort-Object "Start Time"
      $jsonHash["AlltaskBk"] = $arrAllTasksBk 
      $bodyAllTasksBk = $arrAllTasksBk | ConvertTo-HTML -Fragment
      If ($arrAllTasksBk.Status -match "Failed") {
        $allTasksBkHead = $subHead01err
      } ElseIf ($arrAllTasksBk.Status -match "Warning") {
        $allTasksBkHead = $subHead01war
      } ElseIf ($arrAllTasksBk.Status -match "Success") {
        $allTasksBkHead = $subHead01suc
      } Else {
        $allTasksBkHead = $subHead01
      }
      $bodyAllTasksBk = $allTasksBkHead + "Backup Tasks" + $subHead02 + $bodyAllTasksBk
} Else {
  $arrAllTasksBk = $taskListBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, @{Name="Status"; Expression = {($_.Status.ToString())}}
      $arrAllTasksBk  = $arrAllTasksBk | Sort-Object "Start Time"
      $jsonHash["AlltaskBk"] = $arrAllTasksBk 
      $bodyAllTasksBk = $arrAllTasksBk | ConvertTo-HTML -Fragment
      If ($arrAllTasksBk.Status -match "Failed") {
        $allTasksBkHead = $subHead01err
      } ElseIf ($arrAllTasksBk.Status -match "Warning") {
        $allTasksBkHead = $subHead01war
      } ElseIf ($arrAllTasksBk.Status -match "Success") {
        $allTasksBkHead = $subHead01suc
      } Else {
        $allTasksBkHead = $subHead01
      }
      $bodyAllTasksBk = $allTasksBkHead + "Backup Tasks" + $subHead02 + $bodyAllTasksBk
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
    $bodyTasksRunningBk = $bodyTasksRunningBk | ConvertTo-HTML -Fragment
    $bodyTasksRunningBk = $subHead01 + "Running Backup Tasks" + $subHead02 + $bodyTasksRunningBk
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
      $bodyTaskWFBk = $bodyTaskWFBk | ConvertTo-HTML -Fragment
      If ($arrTaskWFBk.Status -match "Failed") {
        $taskWFBkHead = $subHead01err
      } ElseIf ($arrTaskWFBk.Status -match "Warning") {
        $taskWFBkHead = $subHead01war
      } ElseIf ($arrTaskWFBk.Status -match "Success") {
        $taskWFBkHead = $subHead01suc
      } Else {
        $taskWFBkHead = $subHead01
      }
      $bodyTaskWFBk = $taskWFBkHead + "Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBk
} Else {
  $arrTaskWFBk = $wfTasksBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFBk = $arrTaskWFBk | Sort-Object "Start Time"
      $jsonHash["TaskWFBk"] = $bodyTaskWFBk
      $bodyTaskWFBk = $bodyTaskWFBk | ConvertTo-HTML -Fragment
      If ($arrTaskWFBk.Status -match "Failed") {
        $taskWFBkHead = $subHead01err
      } ElseIf ($arrTaskWFBk.Status -match "Warning") {
        $taskWFBkHead = $subHead01war
      } ElseIf ($arrTaskWFBk.Status -match "Success") {
        $taskWFBkHead = $subHead01suc
      } Else {
        $taskWFBkHead = $subHead01
      }
      $bodyTaskWFBk = $taskWFBkHead + "Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBk
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
      $bodyTaskSuccBk = $bodyTaskSuccBk | ConvertTo-HTML -Fragment
      $bodyTaskSuccBk = $subHead01suc + "Successful Backup Tasks" + $subHead02 + $bodyTaskSuccBk
} Else {
  $bodyTaskSuccBk = $successTasksBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}}, @{Name="Status"; Expression = {($_.Status.ToString())}} | Sort-Object "Start Time"
      $jsonHash["TaskSuccBk"] = $bodyTaskSuccBk
      $bodyTaskSuccBk = $bodyTaskSuccBk | ConvertTo-HTML -Fragment
      $bodyTaskSuccBk = $subHead01suc + "Successful Backup Tasks" + $subHead02 + $bodyTaskSuccBk
    }
}
}

# Get Replication Summary Info
$bodySummaryRp = $null
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
  $bodySummaryRp = $arrSummaryRp | ConvertTo-HTML -Fragment
  If ($arrSummaryRp.Failed -gt 0) {
      $summaryRpHead = $subHead01err
  } ElseIf ($arrSummaryRp.Warnings -gt 0) {
      $summaryRpHead = $subHead01war
  } ElseIf ($arrSummaryRp.Successful -gt 0) {
      $summaryRpHead = $subHead01suc
  } Else {
      $summaryRpHead = $subHead01
  }
  $bodySummaryRp = $summaryRpHead + "Replication Results Summary" + $subHead02 + $bodySummaryRp
}

# Get Replication Job Status
$bodyJobsRp = $null
if ($showJobsRp -and $allJobsRp.Count -gt 0) {
  $bodyJobsRp = foreach ($rpJob in $allJobsRp) {
    $rpJob | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Enabled"; Expression = {$_.Info.IsScheduleEnabled}},
      @{Name="State"; Expression = {
        if ($rpJob.IsRunning) {
          $s = $runningSessionsRp | Where-Object {$_.JobName -eq $rpJob.Name}
          if ($s) { "$($s.Progress.Percents)% completed at $([Math]::Round($s.Info.Progress.AvgSpeed/1MB,2)) MB/s" }
          else { "Running (no session info)" }
        } else { "Stopped" }}},
      @{Name="Target"; Expression = {(Get-VBRServer | Where-Object {$_.Id -eq $rpJob.Info.TargetHostId}).Name}},
      @{Name="Target Repo"; Expression = {($repoList + $repoListSo | Where-Object {$_.Id -eq $rpJob.Info.TargetRepositoryId}).Name}},
      @{Name="Next Run"; Expression = {
        try {
          $s = Get-VBRJobScheduleOptions -Job $rpJob
          if (-not $rpJob.IsScheduleEnabled) { "Disabled" }
          elseif ($s.RunManually) { "Not Scheduled" }
          elseif ($s.IsContinious) { "Continious" }
          elseif ($s.OptionsScheduleAfterJob.IsEnabled) { "After [$(($allJobs + $allJobsTp | Where-Object {$_.Id -eq $rpJob.Info.ParentScheduleId}).Name)]" }
          else { $s.NextRun }
        } catch { "Unavailable" }}},
      @{Name="Status"; Expression = {
        $result = $_.GetLastResult()
        if ($result -eq "None") { "" } else { $result.ToString() } }}
  }
  $bodyJobsRp = $bodyJobsRp | Sort-Object "Next Run"
  $jsonHash["JobsRp"] = $bodyJobsRp
  $bodyJobsRp = $subHead01 + "Replication Job Status" + $subHead02 + ($bodyJobsRp | ConvertTo-HTML -Fragment)
}

# Get Replication Sessions
$bodyAllSessRp = $null
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
      $bodyAllSessRp = $arrAllSessRp | ConvertTo-HTML -Fragment
      If ($arrAllSessRp.Result -match "Failed") {
        $allSessRpHead = $subHead01err
      } ElseIf ($arrAllSessRp.Result -match "Warning") {
        $allSessRpHead = $subHead01war
      } ElseIf ($arrAllSessRp.Result -match "Success") {
        $allSessRpHead = $subHead01suc
      } Else {
        $allSessRpHead = $subHead01
      }
      $bodyAllSessRp = $allSessRpHead + "Replication Sessions" + $subHead02 + $bodyAllSessRp
} Else {
  $arrAllSessRp = $sessListRp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="State"; Expression = {$_.State.ToString()}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
      $jsonHash["AllSessRp"] = $arrAllSessRp
      $bodyAllSessRp = $arrAllSessRp | ConvertTo-HTML -Fragment
      If ($arrAllSessRp.Result -match "Failed") {
        $allSessRpHead = $subHead01err
      } ElseIf ($arrAllSessRp.Result -match "Warning") {
        $allSessRpHead = $subHead01war
      } ElseIf ($arrAllSessRp.Result -match "Success") {
        $allSessRpHead = $subHead01suc
      } Else {
        $allSessRpHead = $subHead01
      }
      $bodyAllSessRp = $allSessRpHead + "Replication Sessions" + $subHead02 + $bodyAllSessRp
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
  $bodyRunningRp = $bodyRunningRp | ConvertTo-HTML -Fragment
    $bodyRunningRp = $subHead01 + "Running Replication Jobs" + $subHead02 + $bodyRunningRp
  }
}

# Get Replication Sessions with Warnings or Failures
$bodySessWFRp = $null
if ($showWarnFailRp) {
    $sessWF = @($warningSessionsRp + $failsSessionsRp)
    if ($sessWF.Count -gt 0) {
        $headerWF = if ($onlyLastRp) { "Replication Jobs with Warnings or Failures" } else { "Replication Sessions with Warnings or Failures" }

        $arrSessWFRp = $sessWF | Sort-Object CreationTime | Select-Object `
            @{Name="Job Name";Expression={ $_.Name }},
            @{Name="Start Time";Expression={ $_.CreationTime.ToString("dd/MM/yyyy HH:mm") }},
            @{Name="Stop Time";Expression={ $_.EndTime.ToString("dd/MM/yyyy HH:mm") }},
            @{Name="Duration (HH:MM:SS)";Expression={ Get-Duration -ts $_.Progress.Duration }},
            $(if ($showDetailedRp) {
                @(
                    @{Name="Avg Speed (MB/s)";Expression={ [Math]::Round($_.Info.Progress.AvgSpeed / 1MB, 2) }},
                    @{Name="Total (GB)";Expression={ [Math]::Round($_.Info.Progress.ProcessedSize / 1GB, 2) }},
                    @{Name="Processed (GB)";Expression={ [Math]::Round($_.Info.Progress.ProcessedUsedSize / 1GB, 2) }},
                    @{Name="Data Read (GB)";Expression={ [Math]::Round($_.Info.Progress.ReadSize / 1GB, 2) }},
                    @{Name="Transferred (GB)";Expression={ [Math]::Round($_.Info.Progress.TransferedSize / 1GB, 2) }},
                    @{Name="Dedupe";Expression={ if ($_.Progress.ReadSize -eq 0) { 0 } else { [string]([Math]::Round($_.BackupStats.GetDedupeX(), 1)) + "x" }}},
                    @{Name="Compression";Expression={ if ($_.Progress.ReadSize -eq 0) { 0 } else { [string]([Math]::Round($_.BackupStats.GetCompressX(), 1)) + "x" }}}
                )
            }),
            @{Name="Details";Expression={
                if ($_.GetDetails() -eq "") {
                    $_ | Get-VBRTaskSession | ForEach-Object {
                        if ($_.GetDetails()) { $_.Name + ": " + ($_.GetDetails()).Replace("<br />", "ZZbrZZ") }
                    }
                } else {
                    ($_.GetDetails()).Replace("<br />", "ZZbrZZ")
                }
            }},
            @{Name="Result";Expression={ $_.Result.ToString() }}

        $jsonHash["SessWFRp"] = $arrSessWFRp
        $bodySessWFRp = $arrSessWFRp | ConvertTo-Html -Fragment

        if ($arrSessWFRp.Result -match "Failed") {
            $sessWFRpHead = $subHead01err
        } elseif ($arrSessWFRp.Result -match "Warning") {
            $sessWFRpHead = $subHead01war
        } elseif ($arrSessWFRp.Result -match "Success") {
            $sessWFRpHead = $subHead01suc
        } else {
            $sessWFRpHead = $subHead01
        }
        $bodySessWFRp = $sessWFRpHead + $headerWF + $subHead02 + $bodySessWFRp
    }
}

# Get Successful Replication Sessions
$bodySessSuccRp = $null
if ($showSuccessRp -and $successSessionsRp.Count -gt 0) {
    $headerSucc = if ($onlyLastRp) { "Successful Replication Jobs" } else { "Successful Replication Sessions" }

    $bodySessSuccRp = $successSessionsRp | Sort-Object CreationTime | Select-Object `
        @{Name="Job Name";Expression={ $_.Name }},
        @{Name="Start Time";Expression={ $_.CreationTime.ToString("dd/MM/yyyy HH:mm") }},
        @{Name="Stop Time";Expression={ $_.EndTime.ToString("dd/MM/yyyy HH:mm") }},
        @{Name="Duration (HH:MM:SS)";Expression={ Get-Duration -ts $_.Progress.Duration }},
        $(if ($showDetailedRp) {
            @(
                @{Name="Avg Speed (MB/s)";Expression={ [Math]::Round($_.Info.Progress.AvgSpeed / 1MB, 2) }},
                @{Name="Total (GB)";Expression={ [Math]::Round($_.Info.Progress.ProcessedSize / 1GB, 2) }},
                @{Name="Processed (GB)";Expression={ [Math]::Round($_.Info.Progress.ProcessedUsedSize / 1GB, 2) }},
                @{Name="Data Read (GB)";Expression={ [Math]::Round($_.Info.Progress.ReadSize / 1GB, 2) }},
                @{Name="Transferred (GB)";Expression={ [Math]::Round($_.Info.Progress.TransferedSize / 1GB, 2) }},
                @{Name="Dedupe";Expression={ if ($_.Progress.ReadSize -eq 0) { 0 } else { [string]([Math]::Round($_.BackupStats.GetDedupeX(), 1)) + "x" }}},
                @{Name="Compression";Expression={ if ($_.Progress.ReadSize -eq 0) { 0 } else { [string]([Math]::Round($_.BackupStats.GetCompressX(), 1)) + "x" }}}
            )
        }),
        @{Name="Result";Expression={ $_.Result.ToString() }}

    $jsonHash["SessSuccRp"] = $bodySessSuccRp
    $bodySessSuccRp = $bodySessSuccRp | ConvertTo-Html -Fragment
    $bodySessSuccRp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccRp
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
    $arrAllTasksRp = $arrAllTasksRp | Sort-Object "Start Time"
    $jsonHash["AllTasksRp"] = $arrAllTasksRp
    $bodyAllTasksRp = $arrAllTasksRp | ConvertTo-HTML -Fragment
      If ($arrAllTasksRp.Status -match "Failed") {
        $allTasksRpHead = $subHead01err
      } ElseIf ($arrAllTasksRp.Status -match "Warning") {
        $allTasksRpHead = $subHead01war
      } ElseIf ($arrAllTasksRp.Status -match "Success") {
        $allTasksRpHead = $subHead01suc
      } Else {
        $allTasksRpHead = $subHead01
      }
      $bodyAllTasksRp = $allTasksRpHead + "Replication Tasks" + $subHead02 + $bodyAllTasksRp
} Else {
  $arrAllTasksRp = $taskListRp | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
    $arrAllTasksRp = $arrAllTasksRp | Sort-Object "Start Time"
    $jsonHash["AllTasksRp"] = $arrAllTasksRp
    $bodyAllTasksRp = $arrAllTasksRp | ConvertTo-HTML -Fragment
      If ($arrAllTasksRp.Status -match "Failed") {
        $allTasksRpHead = $subHead01err
      } ElseIf ($arrAllTasksRp.Status -match "Warning") {
        $allTasksRpHead = $subHead01war
      } ElseIf ($arrAllTasksRp.Status -match "Success") {
        $allTasksRpHead = $subHead01suc
      } Else {
        $allTasksRpHead = $subHead01
      }
      $bodyAllTasksRp = $allTasksRpHead + "Replication Tasks" + $subHead02 + $bodyAllTasksRp
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
    $bodyTasksRunningRp = $bodyTasksRunningRp | ConvertTo-HTML -Fragment
    $bodyTasksRunningRp = $subHead01 + "Running Replication Tasks" + $subHead02 + $bodyTasksRunningRp
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
      $arrTaskWFRp = $arrTaskWFRp | Sort-Object "Start Time" 
      $jsonHash["TaskWFRp"] = $arrTaskWFRp
      $bodyTaskWFRp = $arrTaskWFRp | ConvertTo-HTML -Fragment
      If ($arrTaskWFRp.Status -match "Failed") {
        $taskWFRpHead = $subHead01err
      } ElseIf ($arrTaskWFRp.Status -match "Warning") {
        $taskWFRpHead = $subHead01war
      } ElseIf ($arrTaskWFRp.Status -match "Success") {
        $taskWFRpHead = $subHead01suc
      } Else {
        $taskWFRpHead = $subHead01
      }
      $bodyTaskWFRp = $taskWFRpHead + "Replication Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFRp
} Else {
  $arrTaskWFRp = $wfTasksRp | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}},
    @{Name="Status"; Expression = {($_.Status.ToString())}}
      $arrTaskWFRp = $arrTaskWFRp | Sort-Object "Start Time" 
      $jsonHash["TaskWFRp"] = $arrTaskWFRp
      $bodyTaskWFRp = $arrTaskWFRp | ConvertTo-HTML -Fragment
      If ($arrTaskWFRp.Status -match "Failed") {
        $taskWFRpHead = $subHead01err
      } ElseIf ($arrTaskWFRp.Status -match "Warning") {
        $taskWFRpHead = $subHead01war
      } ElseIf ($arrTaskWFRp.Status -match "Success") {
        $taskWFRpHead = $subHead01suc
      } Else {
        $taskWFRpHead = $subHead01
      }
      $bodyTaskWFRp = $taskWFRpHead + "Replication Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFRp
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
      $bodyTaskSuccRp = $bodyTaskSuccRp | ConvertTo-HTML -Fragment
      $bodyTaskSuccRp = $subHead01suc + "Successful Replication Tasks" + $subHead02 + $bodyTaskSuccRp
} Else {
  $bodyTaskSuccRp = $successTasksRp | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Status"; Expression = {($_.Status.ToString())}} | Sort-Object "Start Time" 
      $jsonHash["TaskSuccRp"] = $bodyTaskSuccRp
      $bodyTaskSuccRp = $bodyTaskSuccRp | ConvertTo-HTML -Fragment
      $bodyTaskSuccRp = $subHead01suc + "Successful Replication Tasks" + $subHead02 + $bodyTaskSuccRp
    }
}
}

# Get Backup Copy Summary Info
$bodySummaryBc = $null
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
    $bodySummaryBc = $arrSummaryBc | ConvertTo-HTML -Fragment
  If ($arrSummaryBc.Failures -gt 0) {
      $summaryBcHead = $subHead01err
  } ElseIf ($arrSummaryBc.Warnings -gt 0) {
      $summaryBcHead = $subHead01war
  } ElseIf ($arrSummaryBc.Successful -gt 0) {
      $summaryBcHead = $subHead01suc
  } Else {
      $summaryBcHead = $subHead01
  }
  $bodySummaryBc = $summaryBcHead + "Backup Copy Results Summary" + $subHead02 + $bodySummaryBc
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
    $bodyJobsBc = $bodyJobsBc | ConvertTo-HTML -Fragment
    $bodyJobsBc = $subHead01 + "Backup Copy Job Status" + $subHead02 + $bodyJobsBc
}

# Get All Backup Copy Sessions
$bodyAllSessBc = $null
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
          $bodyAllSessBc = $arrAllSessBc | ConvertTo-HTML -Fragment
      If ($arrAllSessBc.Result -match "Failed") {
        $allSessBcHead = $subHead01err
      } ElseIf ($arrAllSessBc.Result -match "Warning") {
        $allSessBcHead = $subHead01war
      } ElseIf ($arrAllSessBc.Result -match "Success") {
        $allSessBcHead = $subHead01suc
      } Else {
        $allSessBcHead = $subHead01
      }
      $bodyAllSessBc = $allSessBcHead + "Backup Copy Sessions" + $subHead02 + $bodyAllSessBc
} Else {
  $arrAllSessBc = $sessListBc | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="State"; Expression = {$_.State.ToString()}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, 
    @{Name="Result"; Expression = {($_.Result.ToString())}}
    $jsonHash["AllSessBc"] = $arrAllSessBc
    $bodyAllSessBc = $arrAllSessBc | ConvertTo-HTML -Fragment
      If ($arrAllSessBc.Result -match "Failed") {
        $allSessBcHead = $subHead01err
      } ElseIf ($arrAllSessBc.Result -match "Warning") {
        $allSessBcHead = $subHead01war
      } ElseIf ($arrAllSessBc.Result -match "Success") {
        $allSessBcHead = $subHead01suc
      } Else {
        $allSessBcHead = $subHead01
      }
      $bodyAllSessBc = $allSessBcHead + "Backup Copy Sessions" + $subHead02 + $bodyAllSessBc
}
}
}

# Get Idle Backup Copy Sessions
$bodySessIdleBc = $null
If ($showIdleBc) {
If ($idleSessionsBc.count -gt 0) {
      If ($onlyLastBc) {
      $headerIdle = "Idle Backup Copy Jobs"
    } Else {
      $headerIdle = "Idle Backup Copy Sessions"
    }
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
      $bodySessIdleBc = $bodySessIdleBc | ConvertTo-HTML -Fragment
      $bodySessIdleBc = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleBc
} Else {
  $bodySessIdleBc = $idleSessionsBc | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}}
      $jsonHash["SessIdleBc"] = $bodySessIdleBc
      $bodySessIdleBc = $bodySessIdleBc | ConvertTo-HTML -Fragment
      $bodySessIdleBc = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleBc
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
    $bodyRunningBc = $bodyRunningBc | ConvertTo-HTML -Fragment
    $bodyRunningBc = $subHead01 + "Working Backup Copy Sessions" + $subHead02 + $bodyRunningBc
  }
}

# Get Backup Copy Sessions with Warnings or Failures
$bodySessWFBc = $null
If ($showWarnFailBc) {
$sessWF = @($warningSessionsBc + $failsSessionsBc)
If ($sessWF.count -gt 0) {
    If ($onlyLastBc) {
      $headerWF = "Backup Copy Jobs with Warnings or Failures"
    } Else {
      $headerWF = "Backup Copy Sessions with Warnings or Failures"
    }
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
      $bodySessWFBc = $arrSessWFBc | ConvertTo-HTML -Fragment
      If ($arrSessWFBc.Result -match "Failed") {
        $sessWFBcHead = $subHead01err
      } ElseIf ($arrSessWFBc.Result -match "Warning") {
        $sessWFBcHead = $subHead01war
      } ElseIf ($arrSessWFBc.Result -match "Success") {
        $sessWFBcHead = $subHead01suc
      } Else {
        $sessWFBcHead = $subHead01
      }
      $bodySessWFBc = $sessWFBcHead + $headerWF + $subHead02 + $bodySessWFBc
} Else {
  $arrSessWFBc = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {
      If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
      Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
      $jsonHash["SessWFBc"] = $arrSessWFBc
      $bodySessWFBc = $arrSessWFBc | ConvertTo-HTML -Fragment
      If ($arrSessWFBc.Result -match "Failed") {
        $sessWFBcHead = $subHead01err
      } ElseIf ($arrSessWFBc.Result -match "Warning") {
        $sessWFBcHead = $subHead01war
      } ElseIf ($arrSessWFBc.Result -match "Success") {
        $sessWFBcHead = $subHead01suc
      } Else {
        $sessWFBcHead = $subHead01
      }
      $bodySessWFBc = $sessWFBcHead + $headerWF + $subHead02 + $bodySessWFBc
    }
}
}

# Get Successful Backup Copy Sessions
$bodySessSuccBc = $null
If ($showSuccessBc) {
If ($successSessionsBc.count -gt 0) {
      If ($onlyLastBc) {
      $headerSucc = "Successful Backup Copy Jobs"
    } Else {
      $headerSucc = "Successful Backup Copy Sessions"
    }
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
      $bodySessSuccBc = $bodySessSuccBc | ConvertTo-HTML -Fragment
      $bodySessSuccBc = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBc
} Else {
  $bodySessSuccBc = $successSessionsBc | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Result
      $jsonHash["SessSuccBc"] = $bodySessSuccBc
      $bodySessSuccBc = $bodySessSuccBc | ConvertTo-HTML -Fragment
      $bodySessSuccBc = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBc
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
      $arrAllTasksBc = $arrAllTasksBc | Sort-Object "Start Time"
      $jsonHash["AllTasksBc"] = $arrAllTasksBc
      $bodyAllTasksBc = $arrAllTasksBc | ConvertTo-HTML -Fragment
      If ($arrAllTasksBc.Status -match "Failed") {
        $allTasksBcHead = $subHead01err
      } ElseIf ($arrAllTasksBc.Status -match "Warning") {
        $allTasksBcHead = $subHead01war
      } ElseIf ($arrAllTasksBc.Status -match "Success") {
        $allTasksBcHead = $subHead01suc
      } Else {
        $allTasksBcHead = $subHead01
      }
      $bodyAllTasksBc = $allTasksBcHead + "Backup Copy Tasks" + $subHead02 + $bodyAllTasksBc
} Else {
  $arrAllTasksBc = $taskListBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, 
    @{Name="Status"; Expression = {($_.Status.ToString())}}
      $arrAllTasksBc = $arrAllTasksBc | Sort-Object "Start Time"
      $jsonHash["AllTasksBc"] = $arrAllTasksBc
      $bodyAllTasksBc = $arrAllTasksBc | ConvertTo-HTML -Fragment
      If ($arrAllTasksBc.Status -match "Failed") {
        $allTasksBcHead = $subHead01err
      } ElseIf ($arrAllTasksBc.Status -match "Warning") {
        $allTasksBcHead = $subHead01war
      } ElseIf ($arrAllTasksBc.Status -match "Success") {
        $allTasksBcHead = $subHead01suc
      } Else {
        $allTasksBcHead = $subHead01
      }
      $bodyAllTasksBc = $allTasksBcHead + "Backup Copy Tasks" + $subHead02 + $bodyAllTasksBc
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
    $bodyTasksPendingBc = $bodyTasksPendingBc | ConvertTo-HTML -Fragment
    $bodyTasksPendingBc = $subHead01 + "Pending Backup Copy Tasks" + $subHead02 + $bodyTasksPendingBc
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
    $bodyTasksRunningBc = $bodyTasksRunningBc | ConvertTo-HTML -Fragment
    $bodyTasksRunningBc = $subHead01 + "Working Backup Copy Tasks" + $subHead02 + $bodyTasksRunningBc
}
}

# Get Backup Copy Tasks with Warnings or Failures
$bodyTaskWFBc = $null
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
    @{Name="Status"; Expression = {($_.Status.ToString())}} | Sort-Object "Start Time"
      $jsonHash["TaskWFBc"] = $arrTaskWFBc
      $bodyTaskWFBc = $arrTaskWFBc | ConvertTo-HTML -Fragment
      If ($arrTaskWFBc.Status -match "Failed") {
        $taskWFBcHead = $subHead01err
      } ElseIf ($arrTaskWFBc.Status -match "Warning") {
        $taskWFBcHead = $subHead01war
      } ElseIf ($arrTaskWFBc.Status -match "Success") {
        $taskWFBcHead = $subHead01suc
      } Else {
        $taskWFBcHead = $subHead01
      }
      $bodyTaskWFBc = $taskWFBcHead + "Backup Copy Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBc
} Else {
  $arrTaskWFBc = $wfTasksBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, 
    @{Name="Status"; Expression = {($_.Status.ToString())}} | Sort-Object "Start Time"
      $jsonHash["TaskWFBc"] = $arrTaskWFBc
      $bodyTaskWFBc = $arrTaskWFBc | ConvertTo-HTML -Fragment
      If ($arrTaskWFBc.Status -match "Failed") {
        $taskWFBcHead = $subHead01err
      } ElseIf ($arrTaskWFBc.Status -match "Warning") {
        $taskWFBcHead = $subHead01war
      } ElseIf ($arrTaskWFBc.Status -match "Success") {
        $taskWFBcHead = $subHead01suc
      } Else {
        $taskWFBcHead = $subHead01
      }
      $bodyTaskWFBc = $taskWFBcHead + "Backup Copy Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBc
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
      $bodyTaskSuccBc = $bodyTaskSuccBc | ConvertTo-HTML -Fragment
      $bodyTaskSuccBc = $subHead01suc + "Successful Backup Copy Tasks" + $subHead02 + $bodyTaskSuccBc
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
      $bodyTaskSuccBc = $bodyTaskSuccBc | ConvertTo-HTML -Fragment
      $bodyTaskSuccBc = $subHead01suc + "Successful Backup Copy Tasks" + $subHead02 + $bodyTaskSuccBc
}
}
}

# Get Tape Backup Summary Info
$bodySummaryTp = $null
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
  $bodySummaryTp = $arrSummaryTp | ConvertTo-HTML -Fragment
  If ($arrSummaryTp.Failures -gt 0) {
      $summaryTpHead = $subHead01err
  } ElseIf ($arrSummaryTp.Warnings -gt 0 -or $arrSummaryTp.Waiting -gt 0) {
      $summaryTpHead = $subHead01war
  } ElseIf ($arrSummaryTp.Successful -gt 0) {
      $summaryTpHead = $subHead01suc
  } Else {
      $summaryTpHead = $subHead01
  }
  $bodySummaryTp = $summaryTpHead + "Tape Backup Results Summary" + $subHead02 + $bodySummaryTp
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
    $bodyJobsTp = $bodyJobsTp | ConvertTo-HTML -Fragment
    $bodyJobsTp = $subHead01 + "Tape Backup Job Status" + $subHead02 + $bodyJobsTp
}

# Get Tape Backup Sessions
$bodyAllSessTp = $null
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
  $bodyAllSessTp = $arrAllSessTp | ConvertTo-HTML -Fragment
      If ($arrAllSessTp.Result -match "Failed") {
        $allSessTpHead = $subHead01err
      } ElseIf ($arrAllSessTp.Result -match "Warning" -or $arrAllSessTp.State -match "WaitingTape") {
        $allSessTpHead = $subHead01war
      } ElseIf ($arrAllSessTp.Result -match "Success") {
        $allSessTpHead = $subHead01suc
      } Else {
        $allSessTpHead = $subHead01
      }
      $bodyAllSessTp = $allSessTpHead + "Tape Backup Sessions" + $subHead02 + $bodyAllSessTp
  } Else {
  $arrAllSessTp = $sessListTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="State"; Expression = {$_.State.ToString()}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
    $jsonHash["AllSessTp"] = $arrAllSessTp
    $bodyAllSessTp = $arrAllSessTp | ConvertTo-HTML -Fragment
      If ($arrAllSessTp.Result -match "Failed") {
        $allSessTpHead = $subHead01err
      } ElseIf ($arrAllSessTp.Result -match "Warning" -or $arrAllSessTp.State -match "WaitingTape") {
        $allSessTpHead = $subHead01war
      } ElseIf ($arrAllSessTp.Result -match "Success") {
        $allSessTpHead = $subHead01suc
      } Else {
        $allSessTpHead = $subHead01
      }
      $bodyAllSessTp = $allSessTpHead + "Tape Backup Sessions" + $subHead02 + $bodyAllSessTp
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
    $bodyWaitingTp = $bodyWaitingTp | ConvertTo-HTML -Fragment
    $bodyWaitingTp = $subHead01war + "Waiting Tape Backup Sessions" + $subHead02 + $bodyWaitingTp
}
}

# Get Idle Tape Backup Sessions
$bodySessIdleTp = $null
If ($showIdleTp) {
If ($idleSessionsTp.count -gt 0) {
    If ($onlyLastTp) {
      $headerIdle = "Idle Tape Backup Jobs"
    } Else {
      $headerIdle = "Idle Tape Backup Sessions"
    }
If ($showDetailedTp) {
  $bodySessIdleTp = $idleSessionsTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}},
    @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
    @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
    @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}}
      $jsonHash["SessIdleTp"] = $bodySessIdleTp
      $bodySessIdleTp = $bodySessIdleTp | ConvertTo-HTML -Fragment
      $bodySessIdleTp = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleTp
} Else {
  $bodySessIdleTp = $idleSessionsTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}}
      $jsonHash["SessIdleTp"] = $bodySessIdleTp
      $bodySessIdleTp = $bodySessIdleTp | ConvertTo-HTML -Fragment
      $bodySessIdleTp = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleTp
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
    $bodyRunningTp = $bodyRunningTp | ConvertTo-HTML -Fragment
    $bodyRunningTp = $subHead01 + "Working Tape Backup Sessions" + $subHead02 + $bodyRunningTp
}
}

# Get Tape Backup Sessions with Warnings or Failures
$bodySessWFTp = $null
If ($showWarnFailTp) {
$sessWF = @($warningSessionsTp + $failsSessionsTp)
If ($sessWF.count -gt 0) {
      If ($onlyLastTp) {
      $headerWF = "Tape Backup Jobs with Warnings or Failures"
    } Else {
      $headerWF = "Tape Backup Sessions with Warnings or Failures"
    }
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
            $bodySessWFTp =  $arrSessWFTp | ConvertTo-HTML -Fragment
      If ($arrSessWFTp.Result -match "Failed") {
        $sessWFTpHead = $subHead01err
      } ElseIf ($arrSessWFTp.Result -match "Warning") {
        $sessWFTpHead = $subHead01war
      } ElseIf ($arrSessWFTp.Result -match "Success") {
        $sessWFTpHead = $subHead01suc
      } Else {
        $sessWFTpHead = $subHead01
      }
      $bodySessWFTp = $sessWFTpHead + $headerWF + $subHead02 + $bodySessWFTp
} Else {
  $arrSessWFTp = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {
      If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
      Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
      $jsonHash["SessWFTp"] = $arrSessWFTp
           $bodySessWFTp =  $arrSessWFTp | ConvertTo-HTML -Fragment
      If ($arrSessWFTp.Result -match "Failed") {
        $sessWFTpHead = $subHead01err
      } ElseIf ($arrSessWFTp.Result -match "Warning") {
        $sessWFTpHead = $subHead01war
      } ElseIf ($arrSessWFTp.Result -match "Success") {
        $sessWFTpHead = $subHead01suc
      } Else {
        $sessWFTpHead = $subHead01
      }
      $bodySessWFTp = $sessWFTpHead + $headerWF + $subHead02 + $bodySessWFTp
}
}
}

# Get Successful Tape Backup Sessions
$bodySessSuccTp = $null
If ($showSuccessTp) {
If ($successSessionsTp.count -gt 0) {
      If ($onlyLastTp) {
      $headerSucc = "Successful Tape Backup Jobs"
    } Else {
      $headerSucc = "Successful Tape Backup Sessions"
    }
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
      $bodySessSuccTp = $bodySessSuccTp | ConvertTo-HTML -Fragment
      $bodySessSuccTp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccTp
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
      $bodySessSuccTp = $bodySessSuccTp | ConvertTo-HTML -Fragment
      $bodySessSuccTp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccTp
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
      $arrAllTasksTp = $arrAllTasksTp | Sort-Object "Start Time"
      $jsonHash["AllTasksTp"] = $bodyAllTasksTp
      $bodyAllTasksTp = $arrAllTasksTp | ConvertTo-HTML -Fragment
      If ($arrAllTasksTp.Status -match "Failed") {
        $allTasksTpHead = $subHead01err
      } ElseIf ($arrAllTasksTp.Status -match "Warning") {
        $allTasksTpHead = $subHead01war
      } ElseIf ($arrAllTasksTp.Status -match "Success") {
        $allTasksTpHead = $subHead01suc
      } Else {
        $allTasksTpHead = $subHead01
      }
      $bodyAllTasksTp = $allTasksTpHead + "Tape Backup Tasks" + $subHead02 + $bodyAllTasksTp
} Else {
  $arrAllTasksTp = $taskListTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, 
    @{Name="Status"; Expression = {($_.Status.ToString())}}
      $arrAllTasksTp = $arrAllTasksTp | Sort-Object "Start Time"
      $jsonHash["AllTasksTp"] = $bodyAllTasksTp
      $bodyAllTasksTp = $arrAllTasksTp | ConvertTo-HTML -Fragment
      If ($arrAllTasksTp.Status -match "Failed") {
        $allTasksTpHead = $subHead01err
      } ElseIf ($arrAllTasksTp.Status -match "Warning") {
        $allTasksTpHead = $subHead01war
      } ElseIf ($arrAllTasksTp.Status -match "Success") {
        $allTasksTpHead = $subHead01suc
      } Else {
        $allTasksTpHead = $subHead01
      }
      $bodyAllTasksTp = $allTasksTpHead + "Tape Backup Tasks" + $subHead02 + $bodyAllTasksTp
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
        $bodyTasksPendingTp = $bodyTasksPendingTp | ConvertTo-HTML -Fragment
        $bodyTasksPendingTp = $subHead01 + "Pending Tape Backup Tasks" + $subHead02 + $bodyTasksPendingTp
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
    $bodyTasksRunningTp = $bodyTasksRunningTp | ConvertTo-HTML -Fragment
    $bodyTasksRunningTp = $subHead01 + "Working Tape Backup Tasks" + $subHead02 + $bodyTasksRunningTp
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
    @{Name="Status"; Expression = {($_.Status.ToString())}} | Sort-Object "Start Time"
      $jsonHash["TaskWFTp"] = $arrTaskWFTp
      $bodyTaskWFTp = $arrTaskWFTp | ConvertTo-HTML -Fragment      
      If ($arrTaskWFTp.Status -match "Failed") {
        $taskWFTpHead = $subHead01err
      } ElseIf ($arrTaskWFTp.Status -match "Warning") {
        $taskWFTpHead = $subHead01war
      } ElseIf ($arrTaskWFTp.Status -match "Success") {
        $taskWFTpHead = $subHead01suc
      } Else {
        $taskWFTpHead = $subHead01
      }
      $bodyTaskWFTp = $taskWFTpHead + "Tape Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFTp
} Else {
  $arrTaskWFTp = $wfTasksTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
    @{Name="Job Name"; Expression = {$_.JobSess.Name}},
    @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
    @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, 
    @{Name="Status"; Expression = {($_.Status.ToString())}} | Sort-Object "Start Time"
      $jsonHash["TaskWFTp"] = $arrTaskWFTp
      $bodyTaskWFTp = $arrTaskWFTp | ConvertTo-HTML -Fragment      
      If ($arrTaskWFTp.Status -match "Failed") {
        $taskWFTpHead = $subHead01err
      } ElseIf ($arrTaskWFTp.Status -match "Warning") {
        $taskWFTpHead = $subHead01war
      } ElseIf ($arrTaskWFTp.Status -match "Success") {
        $taskWFTpHead = $subHead01suc
      } Else {
        $taskWFTpHead = $subHead01
      }
      $bodyTaskWFTp = $taskWFTpHead + "Tape Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFTp
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
      $bodyTaskSuccTp = $bodyTaskSuccTp | ConvertTo-HTML -Fragment
      $bodyTaskSuccTp = $subHead01suc + "Successful Tape Backup Tasks" + $subHead02 + $bodyTaskSuccTp
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
      $bodyTaskSuccTp = $bodyTaskSuccTp | ConvertTo-HTML -Fragment
      $bodyTaskSuccTp = $subHead01suc + "Successful Tape Backup Tasks" + $subHead02 + $bodyTaskSuccTp
}
}
}

# Get all Expired Tapes
$bodyExpTp = $null
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
    $bodyExpTp = $expTapes | ConvertTo-HTML -Fragment
    $bodyExpTp = $subHead01 + "All Expired Tapes" + $subHead02 + $expTapes
}
}

# Get Agent Backup Summary Info
$bodySummaryEp = $null
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
  $bodySummaryEp = $arrSummaryEp | ConvertTo-HTML -Fragment
  If ($arrSummaryEp.Failures -gt 0) {
      $summaryEpHead = $subHead01err
  } ElseIf ($arrSummaryEp.Warnings -gt 0) {
      $summaryEpHead = $subHead01war
  } ElseIf ($arrSummaryEp.Successful -gt 0) {
      $summaryEpHead = $subHead01suc
  } Else {
      $summaryEpHead = $subHead01
  }
  $bodySummaryEp = $summaryEpHead + "Agent Backup Results Summary" + $subHead02 + $bodySummaryEp
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
      $bodyJobsEp = $bodyJobsEp | ConvertTo-HTML -Fragment
      $bodyJobsEp = $subHead01 + "Agent Backup Job Status" + $subHead02 + $bodyJobsEp
}

# Get Agent Backup Job Size
$bodyJobSizeEp = $null
If ($showBackupSizeEp) {
If ($backupsEp.count -gt 0) {
    $bodyJobSizeEp = Get-BackupSize -backups $backupsEp | Sort-Object JobName | Select-Object @{Name="Job Name"; Expression = {$_.JobName}},
  @{Name="VM Count"; Expression = {$_.VMCount}},
  @{Name="Repository"; Expression = {$_.Repo}},
      @{Name="Backup Size (GB)"; Expression = {$_.LogSize}}
    $jsonHash["JobsSizeAg"] = $bodyJobSizeEp
    $bodyJobSizeEp = $bodyJobSizeEp | ConvertTo-HTML -Fragment
    $bodyJobSizeEp = $subHead01 + "Agent Backup Job Size" + $subHead02 + $bodyJobSizeEp
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
  $arrAllSessEp = $arrAllSessEp | Sort-Object "Start Time"    
  $jsonHash["AllSessAg"] = $bodyAllSessEp
  $bodyAllSessEp = $arrAllSessEp | ConvertTo-HTML -Fragment
      If ($arrAllSessEp.Result -match "Failed") {
        $allSessEpHead = $subHead01err
      } ElseIf ($arrAllSessEp.Result -match "Warning") {
        $allSessEpHead = $subHead01war
      } ElseIf ($arrAllSessEp.Result -match "Success") {
        $allSessEpHead = $subHead01suc
      } Else {
        $allSessEpHead = $subHead01
      }
    $bodyAllSessEp = $allSessEpHead + "Agent Backup Sessions" + $subHead02 + $bodyAllSessEp
  }
}

# Get Running Agent Backup Jobs
$bodyRunningEp = @()
If ($showRunningEp) {
If ($runningSessionsEp.count -gt 0) {
Foreach($job in $allJobsEp) {
  $arrrunningSessionsEp += $runningSessionsEp | Where-Object {$_.JobId -eq $job.Id} | Select-Object @{Name="Job Name"; Expression = {$job.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}}
}
    $arrrunningSessionsEp = $arrrunningSessionsEp | Sort-Object "Start Time"
    $jsonHash["AllRunningAg"] = $arrrunningSessionsEp
    $bodyRunningEp = $arrrunningSessionsEp | ConvertTo-HTML -Fragment
    $bodyRunningEp = $subHead01 + "Running Agent Backup Jobs" + $subHead02 + $bodyRunningEp
  }
}

# Get Agent Backup Sessions with Warnings or Failures
$bodySessWFEp = @()
$arrSessWFEp = @()
If ($showWarnFailEp) {
$sessWFEp = @($warningSessionsEp + $failsSessionsEp)
If ($sessWFEp.count -gt 0) {
      If ($onlyLastEp) {
      $headerWFEp = "Agent Backup Jobs with Warnings or Failures"
    } Else {
      $headerWFEp = "Agent Backup Sessions with Warnings or Failures"
    }
Foreach($job in $allJobsEp) {
  $arrSessWFEp += $sessWFEp | Where-Object {$_.JobId -eq $job.Id} | Select-Object @{Name="Job Name"; Expression = {$job.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}}, @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
    @{Name="Result"; Expression = {($_.Result.ToString())}}
  }
  $jsonHash["SessWFAg"] = $arrSessWFEp
    $bodySessWFEp = $arrSessWFEp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    If ($arrSessWFEp.Result -match "Failed") {
        $sessWFEpHead = $subHead01err
      } ElseIf ($arrSessWFEp.Result -match "Warning") {
        $sessWFEpHead = $subHead01war
      } ElseIf ($arrSessWFEp.Result -match "Success") {
        $sessWFEpHead = $subHead01suc
      } Else {
        $sessWFEpHead = $subHead01
      }
    $bodySessWFEp = $sessWFEpHead + $headerWFEp + $subHead02 + $bodySessWFEp
  }
}

# Get Successful Agent Backup Sessions
$bodySessSuccEp = @()
If ($showSuccessEp) {
If ($successSessionsEp.count -gt 0) {
  If ($onlyLastEp) {
      $headerSuccEp = "Successful Agent Backup Jobs"
    } Else {
      $headerSuccEp = "Successful Agent Backup Sessions"
    }
Foreach($job in $allJobsEp) {
  $bodySessSuccEp += $successSessionsEp | Where-Object {$_.JobId -eq $job.Id} | Select-Object @{Name="Job Name"; Expression = {$job.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}}, @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
    @{Name="Result"; Expression = {($_.Result.ToString())}}
}
    $bodySessSuccEp = $bodySessSuccEp | Sort-Object "Start Time"
    $jsonHash["SessSuccAg"] = $bodySessSuccEp
    $bodySessSuccEp = $bodySessSuccEp | ConvertTo-HTML -Fragment
    $bodySessSuccEp = $subHead01suc + $headerSuccEp + $subHead02 + $bodySessSuccEp
}
}

# Get SureBackup Summary Info
$bodySummarySb = $null
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
  $bodySummarySb = $arrSummarySb | ConvertTo-HTML -Fragment
  If ($arrSummarySb.Failures -gt 0) {
      $summarySbHead = $subHead01err
  } ElseIf ($arrSummarySb.Warnings -gt 0) {
      $summarySbHead = $subHead01war
  } ElseIf ($arrSummarySb.Successful -gt 0) {
      $summarySbHead = $subHead01suc
  } Else {
      $summarySbHead = $subHead01
  }
  $bodySummarySb = $summarySbHead + "SureBackup Results Summary" + $subHead02 + $bodySummarySb
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
    $bodyJobsSb = $bodyJobsSb | ConvertTo-HTML -Fragment
    $bodyJobsSb = $subHead01 + "SureBackup Job Status" + $subHead02 + $bodyJobsSb
}

# Get SureBackup Sessions
$bodyAllSessSb = $null
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
        $bodyAllSessSb = $arrAllSessSb | ConvertTo-HTML -Fragment
    If ($arrAllSessSb.Result -match "Failed") {
        $allSessSbHead = $subHead01err
      } ElseIf ($arrAllSessSb.Result -match "Warning") {
        $allSessSbHead = $subHead01war
      } ElseIf ($arrAllSessSb.Result -match "Success") {
        $allSessSbHead = $subHead01suc
      } Else {
        $allSessSbHead = $subHead01
      }
    $bodyAllSessSb = $allSessSbHead + "SureBackup Sessions" + $subHead02 + $bodyAllSessSb
    }
}

# Get Running SureBackup Jobs
$bodyRunningSb = $null
If ($showRunningSb) {
If ($runningSessionsSb.count -gt 0) {
    $runningSessionsSb = $runningSessionsSb | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
  @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
  @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}},
  @{Name="% Complete"; Expression = {$_.Progress}}
  $jsonHash["SessionsSb"] = $runningSessionsSb
  $bodyRunningSb = $runningSessionsSb | ConvertTo-HTML -Fragment
  $bodyRunningSb = $subHead01 + "Running SureBackup Jobs" + $subHead02 + $bodyRunningSb
}
}

# Get SureBackup Sessions with Warnings or Failures
$bodySessWFSb = $null
If ($showWarnFailSb) {
$sessWF = @($warningSessionsSb + $failsSessionsSb)
If ($sessWF.count -gt 0) {
      If ($onlyLastSb) {
      $headerWF = "SureBackup Jobs with Warnings or Failures"
    } Else {
      $headerWF = "SureBackup Sessions with Warnings or Failures"
    }
$arrSessWFSb = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}}, @{Name="Result"; Expression = {($_.Result.ToString())}}
    $jsonHash["SessWFSb"] = $arrSessWFSb
    $bodySessWFSb = $arrSessWFSb | ConvertTo-HTML -Fragment
    If ($arrSessWFSb.Result -match "Failed") {
        $sessWFSbHead = $subHead01err
      } ElseIf ($arrSessWFSb.Result -match "Warning") {
        $sessWFSbHead = $subHead01war
      } ElseIf ($arrSessWFSb.Result -match "Success") {
        $sessWFSbHead = $subHead01suc
      } Else {
        $sessWFSbHead = $subHead01
      }
    $bodySessWFSb = $sessWFSbHead + $headerWF + $subHead02 + $bodySessWFSb
}
}

# Get Successful SureBackup Sessions
$bodySessSuccSb = $null
If ($showSuccessSb) {
If ($successSessionsSb.count -gt 0) {
      If ($onlyLastSb) {
      $headerSucc = "Successful SureBackup Jobs"
    } Else {
      $headerSucc = "Successful SureBackup Sessions"
    }
$arrSessSuccSb = $successSessionsSb | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
    @{Name="Start Time"; Expression = {$_.CreationTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Stop Time"; Expression = {$_.EndTime.ToString("dd/MM/yyyy HH:mm")}},
    @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
        @{Name="Result"; Expression = {($_.Result.ToString())}}
        $jsonHash["SessSuccSb"] = $arrSessSuccSb
        $bodySessSuccSb = $arrSessSuccSb | ConvertTo-HTML -Fragment
        $bodySessSuccSb = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccSb
}
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all SureBackup Tasks from Sessions within time frame
$taskListSb = @()
$taskListSb += $sessListSb | Get-VSBTaskSession
$successTasksSb = @($taskListSb | Where-Object {$_.Info.Result -eq "Success"})
$wfTasksSb = @($taskListSb | Where-Object {$_.Info.Result -match "Warning|Failed"})
If ($showRunningSb) {
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
    $arrAllTasksSb = $arrAllTasksSb | Sort-Object "Start Time"
    $jsonHash["AllTasksSb"] = arrAllTasksSb
    $bodyAllTasksSb = $arrAllTasksSb | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    If ($arrAllTasksSb.Result -match "Failed") {
        $allTasksSbHead = $subHead01err
      } ElseIf ($arrAllTasksSb.Result -match "Warning") {
        $allTasksSbHead = $subHead01war
      } ElseIf ($arrAllTasksSb.Result -match "Success") {
        $allTasksSbHead = $subHead01suc
      } Else {
        $allTasksSbHead = $subHead01
      }
    $bodyAllTasksSb = $allTasksSbHead + "SureBackup Tasks" + $subHead02 + $bodyAllTasksSb
}
}

# Get Running SureBackup Tasks
$bodyTasksRunningSb = $null
If ($showRunningTasksSb) {
If ($runningTasksSb.count -gt 0) {
$arrTasksRunningSb = $runningTasksSb | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
  @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
  @{Name="Start Time"; Expression = {$_.Info.StartTime}},
  @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $(Get-Date))}},
  @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
  @{Name="Ping Test"; Expression = {$_.PingStatus}},
  @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
  @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
      Status | Sort-Object "Start Time"
      $jsonHash["TasksRunningSb"] = $arrTasksRunningSb
      $bodyTasksRunningSb = $arrTasksRunningSb | ConvertTo-HTML -Fragment
      $bodyTasksRunningSb = $subHead01 + "Running SureBackup Tasks" + $subHead02 + $bodyTasksRunningSb
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
  $jsonHash["TaskWFSb"] = $arrTaskWFSb
  $bodyTaskWFSb = $arrTaskWFSb | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    If ($arrTaskWFSb.Result -match "Failed") {
        $taskWFSbHead = $subHead01err
      } ElseIf ($arrTaskWFSb.Result -match "Warning") {
        $taskWFSbHead = $subHead01war
      } ElseIf ($arrTaskWFSb.Result -match "Success") {
        $taskWFSbHead = $subHead01suc
      } Else {
        $taskWFSbHead = $subHead01
      }
    $bodyTaskWFSb = $taskWFSbHead + "SureBackup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFSb
}
}

# Get Successful SureBackup Tasks
$bodyTaskSuccSb = $null
If ($showTaskSuccessSb) {
If ($successTasksSb.count -gt 0) {
  $arrTaskSuccSb = $successTasksSb | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
  @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
  @{Name="Start Time"; Expression = {$_.Info.StartTime}},
  @{Name="Stop Time"; Expression = {$_.Info.FinishTime}},
  @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $_.Info.FinishTime)}},
  @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
  @{Name="Ping Test"; Expression = {$_.PingStatus}},
  @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
  @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
  @{Name="Result"; Expression = {$_.Info.Result}} | Sort-Object "Start Time"
  $jsonHash["TaskSuccSb"] = $arrTaskSuccSb
  $bodyTaskSuccSb = $arrTaskSuccSb | ConvertTo-HTML -Fragment
  $bodyTaskSuccSb = $subHead01suc + "Successful SureBackup Tasks" + $subHead02 + $bodyTaskSuccSb
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
$vbrConfigObj = $vbrConfigObj | Select-Object Enabled, State, Target, Schedule, "Restore Points", "Next Run", Encrypted, Status
$jsonHash["SummaryConfig"] = $vbrConfigObj
$bodySummaryConfig = $vbrConfigObj | ConvertTo-HTML -Fragment
  If ($configBackup.LastResult -eq "Warning" -or !$configBackup.Enabled) {
    $configHead = $subHead01war
  } ElseIf ($configBackup.LastResult -eq "Success") {
    $configHead = $subHead01suc
  } ElseIf ($configBackup.LastResult -eq "Failed") {
    $configHead = $subHead01err
  } Else {
    $configHead = $subHead01
  }
  $bodySummaryConfig = $configHead + "Configuration Backup Status" + $subHead02 + $bodySummaryConfig
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
  $arrProxy = $arrProxy | Sort-Object "Proxy Host"
  $jsonHash["Proxy"] = $bodyProxy  
  $bodyProxy = $arrProxy | ConvertTo-HTML -Fragment
  If ($arrProxy.Status -match "Failed") {
      $proxyHead = $subHead01err
    } ElseIf ($arrProxy -match "Success") {
      $proxyHead = $subHead01suc
    } Else {
      $proxyHead = $subHead01
    }
    $bodyProxy = $proxyHead + "Proxy Details" + $subHead02 + $bodyProxy
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
  $arrRepo = $arrRepo | Sort-Object "Repository Name"
  $jsonHash["Repo"] = $arrRepo
  $bodyRepo = $arrRepo | ConvertTo-HTML -Fragment
      If ($arrRepo.status -match "Critical") {
      $repoHead = $subHead01err
    } ElseIf ($arrRepo.status -match "Warning|Unknown") {
      $repoHead = $subHead01war
    } ElseIf ($arrRepo.status -match "OK|NoData") {
      $repoHead = $subHead01suc
    } Else {
      $repoHead = $subHead01
    }
    $bodyRepo = $repoHead + "Repository Details" + $subHead02 + $bodyRepo
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
  $arrSORepo = $arrSORepo | Sort-Object "Scale Out Repository Name", "Member Repository Name"
  $jsonHash["SORepo"] = $arrSORepo
  $bodySORepo = $arrSORepo | ConvertTo-HTML -Fragment
    If ($arrSORepo.status -match "Critical") {
      $sorepoHead = $subHead01err
    } ElseIf ($arrSORepo.status -match "Warning|Unknown") {
      $sorepoHead = $subHead01war
    } ElseIf ($arrSORepo.status -match "OK") {
      $sorepoHead = $subHead01suc
    } Else {
      $sorepoHead = $subHead01
    }
    $bodySORepo = $sorepoHead + "Scale Out Repository Details" + $subHead02 + $bodySORepo
}
}

# Get Replica Target Info
$repTargets = $null
$bodyReplica = $null
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
    $bodyReplica = $repTargets | ConvertTo-HTML -Fragment
    If ($repTargets.status -match "Critical") {
      $reptarHead = $subHead01err
    } ElseIf ($repTargets.status -match "Warning|Unknown") {
      $reptarHead = $subHead01war
    } ElseIf ($repTargets.status -match "OK") {
      $reptarHead = $subHead01suc
    } Else {
      $reptarHead = $subHead01
    }
    $bodyReplica = $reptarHead + "Replica Target Details" + $subHead02 + $bodyReplica
  }
}

#region license info
# Get License Info
$bodyLicense = $null
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
    $bodyLicense = $arrLicense | ConvertTo-HTML -Fragment
    If ($arrLicense.Type -eq "Evaluation") {
        $licHead = $subHead01inf
    } Else {
      If ($arrLicense.Status -eq "OK") {
        $licHead = $subHead01suc
      } ElseIf ($arrLicense.Status -eq "Warning") {
        $licHead = $subHead01war
      } Else {
        $licHead = $subHead01err
      }
    }
  $bodyLicense = $licHead + "License/Support Renewal Date" + $subHead02 + $bodyLicense
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

#region HTML Output
$htmlOutput = $headerObj + $bodyTop + $bodySummaryProtect + $bodySummaryBK + $bodySummaryRp + $bodySummaryBc + $bodySummaryTp + $bodySummaryEp + $bodySummarySb

If ($bodySummaryProtect + $bodySummaryBK + $bodySummaryRp + $bodySummaryBc + $bodySummaryTp + $bodySummaryEp + $bodySummarySb) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyMissing + $bodyWarning + $bodySuccess

If ($bodyMissing + $bodySuccess + $bodyWarning) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyMultiJobs

If ($bodyMultiJobs) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsBk + $bodyJobSizeBk + $bodyFileJobsBk + $bodyFileJobSizeBk + $bodyAllSessBk + $bodyAllTasksBk + $bodyRunningBk + $bodyTasksRunningBk + $bodySessWFBk + $bodyTaskWFBk + $bodySessSuccBk + $bodyTaskSuccBk

If ($bodyJobsBk + $bodyJobSizeBk + $bodyAllSessBk + $bodyAllTasksBk + $bodyRunningBk + $bodyTasksRunningBk + $bodySessWFBk + $bodyTaskWFBk + $bodySessSuccBk + $bodyTaskSuccBk) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyRestoRunVM + $bodyRestoreVM

If ($bodyRestoRunVM + $bodyRestoreVM) {
  $htmlOutput += $HTMLbreak
  }

$htmlOutput += $bodyJobsRp + $bodyAllSessRp + $bodyAllTasksRp + $bodyRunningRp + $bodyTasksRunningRp + $bodySessWFRp + $bodyTaskWFRp + $bodySessSuccRp + $bodyTaskSuccRp

If ($bodyJobsRp + $bodyAllSessRp + $bodyAllTasksRp + $bodyRunningRp + $bodyTasksRunningRp + $bodySessWFRp + $bodyTaskWFRp + $bodySessSuccRp + $bodyTaskSuccRp) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsBc + $bodyJobSizeBc + $bodyAllSessBc + $bodyAllTasksBc + $bodySessIdleBc + $bodyTasksPendingBc + $bodyRunningBc + $bodyTasksRunningBc + $bodySessWFBc + $bodyTaskWFBc + $bodySessSuccBc + $bodyTaskSuccBc

If ($bodyJobsBc + $bodyJobSizeBc + $bodyAllSessBc + $bodyAllTasksBc + $bodySessIdleBc + $bodyTasksPendingBc + $bodyRunningBc + $bodyTasksRunningBc + $bodySessWFBc + $bodyTaskWFBc + $bodySessSuccBc + $bodyTaskSuccBc) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsTp + $bodyAllSessTp + $bodyAllTasksTp + $bodyWaitingTp + $bodySessIdleTp + $bodyTasksPendingTp + $bodyRunningTp + $bodyTasksRunningTp + $bodySessWFTp + $bodyTaskWFTp + $bodySessSuccTp + $bodyTaskSuccTp

If ($bodyJobsTp + $bodyAllSessTp + $bodyAllTasksTp + $bodyWaitingTp + $bodySessIdleTp + $bodyTasksPendingTp + $bodyRunningTp + $bodyTasksRunningTp + $bodySessWFTp + $bodyTaskWFTp + $bodySessSuccTp + $bodyTaskSuccTp) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyTapes + $bodyTpPool + $bodyTpVlt + $bodyExpTp + $bodyTpExpPool + $bodyTpExpVlt + $bodyTpWrt

If ($bodyTapes + $bodyTpPool + $bodyTpVlt + $bodyExpTp + $bodyTpExpPool + $bodyTpExpVlt + $bodyTpWrt) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsEp + $bodyJobSizeEp + $bodyAllSessEp + $bodyRunningEp + $bodySessWFEp + $bodySessSuccEp

If ($bodyJobsEp + $bodyJobSizeEp + $bodyAllSessEp + $bodyRunningEp + $bodySessWFEp + $bodySessSuccEp) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsSb + $bodyAllSessSb + $bodyAllTasksSb + $bodyRunningSb + $bodyTasksRunningSb + $bodySessWFSb + $bodyTaskWFSb + $bodySessSuccSb + $bodyTaskSuccSb

If ($bodyJobsSb + $bodyAllSessSb + $bodyAllTasksSb + $bodyRunningSb + $bodyTasksRunningSb + $bodySessWFSb + $bodyTaskWFSb + $bodySessSuccSb + $bodyTaskSuccSb) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodySummaryConfig + $bodyProxy + $bodyRepo + $bodySORepo + $bodyRepoPerms + $bodyReplica + $bodyServices + $bodyLicense + $footerObj

# Fix Details
$htmlOutput = $htmlOutput.Replace("ZZbrZZ","<br />")
# Remove trailing HTMLbreak
$htmlOutput = $htmlOutput.Replace("$($HTMLbreak + $footerObj)","$($footerObj)")
# Add color to output depending on results
#Green
$htmlOutput = $htmlOutput.Replace("<td>Running<","<td style=""color: #00b051;"">Running<")
$htmlOutput = $htmlOutput.Replace("<td>OK<","<td style=""color: #00b051;"">OK<")
$htmlOutput = $htmlOutput.Replace("<td>Alive<","<td style=""color: #00b051;"">Alive<")
$htmlOutput = $htmlOutput.Replace("<td>Success<","<td style=""color: #00b051;"">Success<")
#Yellow
$htmlOutput = $htmlOutput.Replace("<td>Warning<","<td style=""color: #ffc000;"">Warning<")
#Red
$htmlOutput = $htmlOutput.Replace("<td>Not Running<","<td style=""color: #ff0000;"">Not Running<")
$htmlOutput = $htmlOutput.Replace("<td>Failed<","<td style=""color: #ff0000;"">Failed<")
$htmlOutput = $htmlOutput.Replace("<td>Critical<","<td style=""color: #ff0000;"">Critical<")
$htmlOutput = $htmlOutput.Replace("<td>Dead<","<td style=""color: #ff0000;"">Dead<")
# Color Report Header and Tag Email Subject
If ($htmlOutput -match "#FB9895") {
  # If any errors paint report header red
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ","#FB9895")
  $emailSubject = "[Failed] $emailSubject"
} ElseIf ($htmlOutput -match "#ffd96c") {
  # If any warnings paint report header yellow
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ","#ffd96c")
  $emailSubject = "[Warning] $emailSubject"
} ElseIf ($htmlOutput -match "#00b050") {
  # If any success paint report header green
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ","#00b050")
  $emailSubject = "[Success] $emailSubject"
} Else {
  # Else paint gray
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ","#626365")
}
#endregion

# Save HTML Report to File
If ($saveHTML) {
  $htmlOutput | Out-File $pathHTML
  If ($launchHTML) {
    Invoke-Item $pathHTML
  }
}

#region Output
# Send Report via Email
$smtp = New-Object System.Net.Mail.SmtpClient($emailHost, $emailPort)
$smtp.Credentials = New-Object System.Net.NetworkCredential($emailUser, $emailPass)
$smtp.EnableSsl = $emailEnableSSL
$msg = New-Object System.Net.Mail.MailMessage($emailFrom, $emailTo)
$msg.Subject = $emailSubject
$attachment = New-Object System.Net.Mail.Attachment $pathJSON
$msg.Attachments.Add($attachment)
$body = $htmlOutput
$msg.Body = $body
$msg.IsBodyHtml = $true
$smtp.Send($msg)
#endregion

#region purge

Get-childitem -path $pathJSON -Recurse | where-object {($_.LastWriteTime -lt (get-date).adddays(-$JPurge))} | Remove-Item

#endregion
