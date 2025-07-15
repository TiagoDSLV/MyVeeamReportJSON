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

# Configuration du Rapport
    $SDBackup = $true
    $SDCopy = $true
    $SDReplication = $true
    $SDTape = $true
    $SDAgent = $true
    $SDSure =$true

# Nom du client
    $Client = ""
  # Mail Collecteur GLPI
    $MailGLPI = ""
# Chemin du dossier où stocker les rapports
    $path = "C:\Tools\MyVeeamReport\Reports"
# Rétention des fichiers de rapports JSON en jours
    $JPurge = 60
	
# Report mode (RPO) - valid modes: any number of hours, Weekly or Monthly
# 24, 48, "Weekly", "Monthly"
    $reportMode = 24
	
# Configuration de l'envoi d'email
    $emailHost = ""
    $emailPort = 25
    $emailEnableSSL = $false
    $emailUser = ""
    $emailPass = ""
    $emailFrom = ""
    $emailTo = ""

# Exclusions
    # Exclure les VMs des sections des sauvegardes manquantes et réussies
    $excludevms = @("*_replica")
    # Exclure les VMs des sections des sauvegardes manquantes et réussies dans les dossiers suivants
    $excludeFolder = @("")
    # Exclure les VMs des sections des sauvegardes manquantes et réussies dans les datacenter suivants
    $excludeDC = @("")
    # Exclure les VMs des sections des sauvegardes manquantes et réussies dans les clusters suivants
    $excludeCluster = @("")
    # Exclure les VMs des sections des sauvegardes manquantes et réussies dans les tags suivants
    $excludeTags = @("")
    # Exclure les modèles des sections des sauvegardes manquantes et réussies
    $excludeTemp = $true
    # Exclure les repositories
    $excludedRepositories = @("")


# Seuils d'alerte
  # Espace libre sur les Repository en %
  $repoCritical = 10
  $repoWarn = 20
  # Espace libre sur les cibles de réplication en %
  $replicaCritical = 10
  $replicaWarn = 20
  # Jours restant sur la licence
  $licenseCritical = 30
  $licenseWarn = 90


# Variable à ne pas modifier sauf si nécessaire
# Dates
  $date_file = Get-Date -format ddMMyy_HHmm

# VBR Server (Server Name, FQDN, IP or localhost)
$vbrServer = $env:computername
# Report Title
$rptTitle = "$Client - Rapport de sauvegarde"
# Show VBR Server name in report header
$showVBR = $true
# HTML Report Width (Percent)
$rptWidth = 97
# HTML Table Odd Row color
$oddColor = "#f0f0f0"

# Location of Veeam Core dll  
$VeeamCorePath = "C:\Program Files\Veeam\Backup and Replication\Backup\Veeam.Backup.Core.dll"

# JSON File output path and filename
$pathJSON = $pathJSON + $Client + "_Rapport_Veeam_" + $date_file + ".json"
# Launch JSON file after creation
$launchJSON = $false
# Save JSON output to a file
$saveJSON = $true
# JSON File output path and filename
$pathHTML = $path + $Client + "_Rapport_Veeam_" + $date_file + ".html"
# Launch HTML file after creation
$launchHTML = $false
# Save HTML output to a file
$saveHTML = $true
# Email Subject 
$emailSubject = $rptTitle
# Append Report Mode to Email Subject E.g. My Veeam Report (Last 24 Hours)
$modeSubject = $true
# Append VBR Server name to Email Subject
$vbrSubject = $false
# Append Date and Time to Email Subject
$dtSubject = $true

#--------------------- Disable reports you do not need by setting them to "$false" below:                                                                                        
# Show VM Backup Protection Summary (across entire infrastructure)
$showSummaryProtect = $true
# Show VMs with No Successful Backups within RPO ($reportMode)
$showUnprotectedVMs = $true
# Show unprotected VMs for informational purposes only
$showUnprotectedVMsInfo = $true
# Show VMs with Successful Backups within RPO ($reportMode)
# Also shows VMs with Only Backups with Warnings within RPO ($reportMode)
$showProtectedVMs = $true


# Show VMs Backed Up by Multiple Jobs within time frame ($reportMode)
$showMultiJobs = $true

# Show Backup Session Summary
$showSummaryBk = $SDBackup
# Show Backup Job Status
$showJobsBk = $False
# Show File Backup Job Status
$showFileJobsBk = $SDBackup
# Show detailed information for Backup Jobs/Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedBk = $False
# Show all Backup Sessions within time frame ($reportMode)
$showAllSessBk = $False
# Show all Backup Tasks from Sessions within time frame ($reportMode)
$showAllTasksBk = $False
# Show Running Backup Jobs
$showRunningBk = $False
# Show Running Backup Tasks
$showRunningTasksBk = $False
# Show Backup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailBk = $False
# Show Backup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFBk = $SDBackup
# Show Successful Backup Sessions within time frame ($reportMode)
$showSuccessBk = $False
# Show Successful Backup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessBk = $False
# Only show last Session for each Backup Job
$onlyLastBk = $SDBackup
# Only report on the following Backup Job(s)
#$backupJob = @("Backup Job 1","Backup Job 3","Backup Job *")
$backupJob = @("")

# Show Replication Session Summary
$showSummaryRp = $SDReplication
# Show Replication Job Status
$showJobsRp = $False
# Show detailed information for Replication Jobs/Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedRp = $False
# Show all Replication Sessions within time frame ($reportMode)
$showAllSessRp = $False
# Show all Replication Tasks from Sessions within time frame ($reportMode)
$showAllTasksRp = $False
# Show Running Replication Jobs
$showRunningRp = $False
# Show Running Replication Tasks
$showRunningTasksRp = $False
# Show Replication Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailRp = $False
# Show Replication Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFRp = $SDReplication
# Show Successful Replication Sessions within time frame ($reportMode)
$showSuccessRp = $False
# Show Successful Replication Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessRp = $False
# Only show last session for each Replication Job
$onlyLastRp = $SDReplication
# Only report on the following Replication Job(s)
#$replicaJob = @("Replica Job 1","Replica Job 3","Replica Job *")
$replicaJob = @("")

# Show Backup Copy Session Summary
$showSummaryBc = $SDCopy
# Show Backup Copy Job Status
$showJobsBc = $False
# Show detailed information for Backup Copy Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedBc = $False
# Show all Backup Copy Sessions within time frame ($reportMode)
$showAllSessBc = $False
# Show all Backup Copy Tasks from Sessions within time frame ($reportMode)
$showAllTasksBc = $False
# Show Idle Backup Copy Sessions
$showIdleBc = $False
# Show Pending Backup Copy Tasks
$showPendingTasksBc = $False
# Show Working Backup Copy Jobs
$showRunningBc = $False
# Show Working Backup Copy Tasks
$showRunningTasksBc = $False
# Show Backup Copy Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailBc = $False
# Show Backup Copy Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFBc = $SDCopy
# Show Successful Backup Copy Sessions within time frame ($reportMode)
$showSuccessBc = $False
# Show Successful Backup Copy Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessBc = $False
# Only show last Session for each Backup Copy Job
$onlyLastBc = $SDCopy
# Only report on the following Backup Copy Job(s)
#$bcopyJob = @("Backup Copy Job 1","Backup Copy Job 3","Backup Copy Job *")
$bcopyJob = @("")

# Show Tape Backup Session Summary
$showSummaryTp = $SDTape
# Show Tape Backup Job Status
$showJobsTp = $False
# Show detailed information for Tape Backup Sessions (Avg Speed, Total(GB), Read(GB), Transferred(GB))
$showDetailedTp = $False
# Show all Tape Backup Sessions within time frame ($reportMode)
$showAllSessTp = $False
# Show all Tape Backup Tasks from Sessions within time frame ($reportMode)
$showAllTasksTp = $False
# Show Waiting Tape Backup Sessions
$showWaitingTp = $False
# Show Idle Tape Backup Sessions
$showIdleTp = $False
# Show Pending Tape Backup Tasks
$showPendingTasksTp = $False
# Show Working Tape Backup Jobs
$showRunningTp = $False
# Show Working Tape Backup Tasks
$showRunningTasksTp = $False
# Show Tape Backup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailTp = $False
# Show Tape Backup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFTp = $SDTape
# Show Successful Tape Backup Sessions within time frame ($reportMode)
$showSuccessTp = $False
# Show Successful Tape Backup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessTp = $False
# Only show last Session for each Tape Backup Job
$onlyLastTp = $SDTape
# Only report on the following Tape Backup Job(s)
#$tapeJob = @("Tape Backup Job 1","Tape Backup Job 3","Tape Backup Job *")
$tapeJob = @("")

# Show Agent Backup Session Summary
$showSummaryEp = $SDAgent
# Show Agent Backup Job Status
$showJobsEp = $False
# Show Agent Backup Job Size (total)
$showBackupSizeEp = $False
# Show all Agent Backup Sessions within time frame ($reportMode)
$showAllSessEp = $False
# Show Running Agent Backup jobs
$showRunningEp = $False
# Show Agent Backup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailEp = $SDAgent
# Show Successful Agent Backup Sessions within time frame ($reportMode)
$showSuccessEp = $False
# Only show last session for each Agent Backup Job
$onlyLastEp = $SDAgent
# Only report on the following Agent Backup Job(s)
#$epbJob = @("Agent Backup Job 1","Agent Backup Job 3","Agent Backup Job *")
$epbJob = @("")

# Show SureBackup Session Summary
$showSummarySb = $SDSure
# Show SureBackup Job Status
$showJobsSb = $False
# Show all SureBackup Sessions within time frame ($reportMode)
$showAllSessSb = $False
# Show all SureBackup Tasks from Sessions within time frame ($reportMode)
$showAllTasksSb = $False
# Show Running SureBackup Jobs
$showRunningSb = $False
# Show Running SureBackup Tasks
$showRunningTasksSb = $False
# Show SureBackup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailSb = $False
# Show SureBackup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFSb = $SDSure
# Show Successful SureBackup Sessions within time frame ($reportMode)
$showSuccessSb = $False
# Show Successful SureBackup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessSb = $False
# Only show last Session for each SureBackup Job
$onlyLastSb = $SDSure
# Only report on the following SureBackup Job(s)
#$surebJob = @("SureBackup Job 1","SureBackup Job 3","SureBackup Job *")
$surebJob = @("")

# Show Configuration Backup Summary
$showSummaryConfig = $true
# Show Proxy Info
$showProxy = $true
# Show Repository Info
$showRepo = $true
# Show Replica Target Info
$showReplicaTarget = $SDReplication
# Show License expiry info
$showLicExp = $true
