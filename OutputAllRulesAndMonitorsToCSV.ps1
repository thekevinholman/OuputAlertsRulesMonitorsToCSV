#=================================================================================
#  Get all Rule and Monitors from SCOM and their properties
#
#  Author: Kevin Holman
#  v1.1
#=================================================================================

# Output Directory
$OutDir = "C:\Report"

Write-Host `n"Starting Script to get all rules and monitors in SCOM" -ForegroundColor Green

IF (!(Test-Path $OutDir))
{
  Write-Host `n"Output folder not found.  Creating folder..." -ForegroundColor Magenta
  md $OutDir
}

Write-Host `n"Output path is ($OutDir)" -ForegroundColor Green

$RuleReport = @()

# Connect to SCOM
Write-Host `n"Connecting to local SCOM Management Server..." -ForegroundColor Green
$MG = Get-SCOMManagementGroup -ComputerName "localhost"

#Get all the SCOM Rules
Write-Host `n"Getting all Rules in SCOM..." -ForegroundColor Green
$Rules = Get-SCOMRule

#Create a hashtable of all the SCOM classes for faster retreival based on Class ID
$Classes = Get-SCOMClass
$ClassHT = @{}
FOREACH ($Class in $Classes)
{
  $ClassHT.Add("$($Class.Id)",$Class)
}

#Get GenerateAlert WriteAction module
$HealthMP = Get-SCOMManagementPack -Name "System.Health.Library"
$AlertWA = $HealthMP.GetModuleType("System.Health.GenerateAlert")
$AlertWAID = $AlertWA.Id

Write-Host `n"Getting Properties from Each Rule..." -ForegroundColor Green
FOREACH ($Rule in $Rules)
{
  [string]$RuleDisplayName = $Rule.DisplayName
  [string]$RuleName = $Rule.Name
  [string]$Target = ($ClassHT.($Rule.Target.Id.Guid)).DisplayName
  [string]$Category = $Rule.Category
  [string]$Enabled = $Rule.Enabled
    IF ($Enabled -eq "onEssentialMonitoring") {$Enabled = "TRUE"}
    IF ($Enabled -eq "onStandardMonitoring") {$Enabled = "TRUE"}
  $MP = $Rule.GetManagementPack()
  [string]$MPDisplayName = $MP.DisplayName
  [string]$MPName = $Rule.ManagementPackName
  [string]$RuleDS = $Rule.DataSourceCollection.TypeID.Identifier.Path
  [string]$Description = $Rule.Description
 
  #WriteAction Section
  $GenAlert = $false
  $AlertDisplayName = ""
  $AlertPriority = ""
  $AlertSeverity = ""
  $WA = $Rule.writeactioncollection
 
  #Inspect each WA module to see if it contains a System.Health.GenerateAlert module
  FOREACH ($WAModule in $WA)
  {
    $WAId = $WAModule.TypeId.Id
    IF ($WAId -eq $AlertWAID)
    {
      #this rule generates alert using System.Health.GenerateAlert module
      $GenAlert = $true
      #Get the module configuration
      [string]$WAModuleConfig = $WAModule.Configuration
      #Assign the module configuration the XML type and encapsulate it to make it easy to retrieve values
      [xml]$WAModuleConfigXML = "<Root>" + $WAModuleConfig + "</Root>"
      $WAXMLRoot = $WAModuleConfigXML.Root
      #Get the Alert Display Name from the AlertMessageID and MP
      $AlertName = $WAXMLRoot.AlertMessageId.Split('"')[1]
      IF (!($AlertName))
      {
        $AlertName = $WAXMLRoot.AlertMessageId.Split("'")[1]
      }
      $AlertDisplayName = $MP.GetStringResource($AlertName).DisplayName
      #Get Alert Priority and Severity
      $AlertPriority = $WAXMLRoot.Priority
      $AlertPriority = switch($AlertPriority)
      {
        "0" {"Low"}
        "1" {"Medium"} 
        "2" {"High"}
      }
      $AlertSeverity = $WAXMLRoot.Severity
      $AlertSeverity = switch($AlertSeverity)
      {
        "0" {"Information"}
        "1" {"Warning"} 
        "2" {"Critical"}
      }
    } 
    ELSE 
    {
      #need to detect if it's using a Custom Composite WA which contains System.Health.GenerateAlert module
      $WASource = $MG.GetMonitoringModuleType($WAId)

      #Check each write action member modules in the customized write action module...
      FOREACH ($Item in $WASource.WriteActionCollection)
      {
        $ItemId = $Item.TypeId.Id
        IF ($ItemId -eq $AlertWAId)
        {
          $GenAlert = $true
          #Get the module configuration
          [string]$WAModuleConfig = $WAModule.Configuration
          #Assign the module configuration the XML type and encapsulate it to make it easy to retrieve values
          [xml]$WAModuleConfigXML = "<Root>" + $WAModuleConfig + "</Root>"
          $WAXMLRoot = $WAModuleConfigXML.Root
          #Get the Alert Display Name from the AlertMessageID and MP
          $AlertName = $WAXMLRoot.AlertMessageId.Split('"')[1]
          IF (!($AlertName))
          {
            $AlertName = $WAXMLRoot.AlertMessageId.Split("'")[1]
          }
          $AlertDisplayName = $MP.GetStringResource($AlertName).DisplayName
          #Get Alert Priority and Severity
          $AlertPriority = $WAXMLRoot.Priority
          $AlertPriority = switch($AlertPriority)
          {
            "0" {"Low"}
            "1" {"Medium"} 
            "2" {"High"}
          }
          $AlertSeverity = $WAXMLRoot.Severity
          $AlertSeverity = switch($AlertSeverity)
          {
            "0" {"Information"}
            "1" {"Warning"} 
            "2" {"Critical"}
          }
        }
      }
    }
  }

  #Create generic object and assign values  
  $obj = New-Object -TypeName psobject
  $obj | Add-Member -Type NoteProperty -Name "DisplayName" -Value $RuleDisplayName
  $obj | Add-Member -Type NoteProperty -Name "Name" -Value $RuleName
  $obj | Add-Member -Type NoteProperty -Name "Target" -Value $Target
  $obj | Add-Member -Type NoteProperty -Name "Category" -Value $Category 
  $obj | Add-Member -Type NoteProperty -Name "Enabled" -Value $Enabled
  $obj | Add-Member -Type NoteProperty -Name "Alert" -Value $GenAlert
  $obj | Add-Member -Type NoteProperty -Name "AlertName" -Value $AlertDisplayName
  $obj | Add-Member -Type NoteProperty -Name "AlertPriority" -Value $AlertPriority
  $obj | Add-Member -Type NoteProperty -Name "AlertSeverity" -Value $AlertSeverity
  $obj | Add-Member -Type NoteProperty -Name "MPDisplayName" -Value $MPDisplayName
  $obj | Add-Member -Type NoteProperty -Name "MPName" -Value $MPName
  $obj | Add-Member -Type NoteProperty -Name "DataSource" -Value $RuleDS
  $obj | Add-Member -Type NoteProperty -Name "Description" -Value $Description
  $RuleReport += $obj
}

Write-Host `n"Generating Rules.csv at ($OutDir)..." -ForegroundColor Green
$RuleReport | Export-Csv $OutDir\Rules.csv -NotypeInformation



####################
# Monitor Section:
####################

$MonitorReport = @()

#Get all the SCOM Monitors
Write-Host `n"Getting all Monitors in SCOM..." -ForegroundColor Green
$Monitors =  Get-SCOMMonitor

#Loop through each monitor and get properties
Write-Host `n"Getting Properties from Each Monitor..." -ForegroundColor Green
FOREACH ($Monitor in $Monitors)
{
  [string]$MonitorDisplayName = $Monitor.DisplayName
  [string]$MonitorName = $Monitor.Name
  [string]$Target = ($ClassHT.($Monitor.Target.Id.Guid)).DisplayName
  [string]$Category = $Monitor.Category
  [string]$Enabled = $Monitor.Enabled
    IF ($Enabled -eq "onEssentialMonitoring") {$Enabled = "TRUE"}
    IF ($Enabled -eq "onStandardMonitoring") {$Enabled = "TRUE"}
  $MP = $Monitor.GetManagementPack()
  [string]$MPDisplayName = $MP.DisplayName
  [string]$MPName = $MP.Name
  [string]$MonitorClassification = $Monitor.XmlTag
  [string]$MonitorType = $Monitor.TypeID.Identifier.Path
  [string]$Description = $Monitor.Description

  # Get the Alert Settings for the Monitor
  $AlertSettings = $Monitor.AlertSettings
  $GenAlert = ""
  $AlertDisplayName = ""
  $AlertSeverity = ""
  $AlertPriority = ""
  $AutoResolve = ""

  IF (!($AlertSettings))
  {
    $GenAlert = $false
  }
  ELSE
  {
    $GenAlert = $true
    #Get the Alert Display Name from the AlertMessageID and MP
    $AlertName =  $AlertSettings.AlertMessage.Identifier.Path
    $AlertDisplayName = $MP.GetStringResource($AlertName).DisplayName    
    $AlertSeverity = $AlertSettings.AlertSeverity
      IF ($AlertSeverity -eq "MatchMonitorHealth") {$AlertSeverity = $AlertSettings.AlertOnState}
      IF ($AlertSeverity -eq "Error") {$AlertSeverity = "Critical"}
    $AlertPriority = $AlertSettings.AlertPriority
      IF ($AlertPriority -eq "Normal") {$AlertPriority = "Medium"}
    $AutoResolve = $AlertSettings.AutoResolve
  }

  #Create generic object and assign values  
  $obj = New-Object -TypeName psobject
  $obj | Add-Member -Type NoteProperty -Name "DisplayName" -Value $MonitorDisplayName
  $obj | Add-Member -Type NoteProperty -Name "Name" -Value $MonitorName
  $obj | Add-Member -Type NoteProperty -Name "Target" -Value $Target
  $obj | Add-Member -Type NoteProperty -Name "Category" -Value $Category 
  $obj | Add-Member -Type NoteProperty -Name "Enabled" -Value $Enabled
  $obj | Add-Member -Type NoteProperty -Name "Alert" -Value $GenAlert
  $obj | Add-Member -Type NoteProperty -Name "AlertName" -Value $AlertDisplayName
  $obj | Add-Member -Type NoteProperty -Name "AlertPriority" -Value $AlertPriority
  $obj | Add-Member -Type NoteProperty -Name "AlertSeverity" -Value $AlertSeverity
  $obj | Add-Member -Type NoteProperty -Name "MPDisplayName" -Value $MPDisplayName
  $obj | Add-Member -Type NoteProperty -Name "MPName" -Value $MPName
  $obj | Add-Member -Type NoteProperty -Name "MonitorClassification" -Value $MonitorClassification
  $obj | Add-Member -Type NoteProperty -Name "MonitorType" -Value $MonitorType
  $obj | Add-Member -Type NoteProperty -Name "Description" -Value $Description
  $MonitorReport += $obj
}

Write-Host `n"Generating Monitors.csv at ($OutDir)..." -ForegroundColor Green
$MonitorReport | Export-Csv $OutDir\Monitors.csv -NotypeInformation
#End of Script