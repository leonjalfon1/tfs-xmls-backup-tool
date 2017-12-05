﻿<# 
  .SYNOPSIS
  Blah Blah Blah

  .DESCRIPTION
  Blah Blah Blah

  .PARAMETER BuildDefinitionName
  Blah Blah Blah

  .PARAMETER CollectionUrl
  Blah Blah Blah

  .USING
  Blah Blah Blah

  .EXAMPLE
  Blah Blah Blah

  .CONTRIBUTING
  Please feel free to contribute.

  .NOTES
  Leon Jalfon 
  DevOps & ALM Architect
  Sela Software Labs Ltd
#>

param
(
    [Parameter(Mandatory=$false)]
    [string]$WitadminPath,
    [Parameter(Mandatory=$false)]
    [string]$TfsCollectionUrl,
    [Parameter(Mandatory=$false)]
    [string]$TfsTeamProjects,
    [Parameter(Mandatory=$false)]
    [string]$BackupRepository
)


########################### FUNCTIONS ###########################

function Run-CmdCommand
{
    param
    (
        [Parameter(Mandatory=$true)]
        [string]$Command,
        [Parameter(Mandatory=$false)]
        [string]$WorkingDirectory = (Get-Location),
        [Parameter(Mandatory=$false)]
        [int]$WaitProcessToFinish = $true,
        [Parameter(Mandatory=$false)]
        [int]$CheckRateSeconds = 1,
        [Parameter(Mandatory=$false)]
        [int]$TimeoutMinutes = 10
    )

    try
    {
       # Setup the Process startup info
       $pinfo = New-Object System.Diagnostics.ProcessStartInfo
       $pinfo.FileName = "cmd"
       $pinfo.Arguments = "/c $Command"
       $pinfo.UseShellExecute = $false
       $pinfo.CreateNoWindow = $true
       $pinfo.RedirectStandardOutput = $true
       $pinfo.RedirectStandardError = $true
       $pinfo.WorkingDirectory = $WorkingDirectory


       # Create a process object using the startup info
       $process = New-Object System.Diagnostics.Process
       $process.StartInfo = $pinfo


       # Start the process
       $process.Start() | Out-Null


       # Wait process to finish = TRUE
       if($WaitProcessToFinish -eq $true)
       {
           # Set timeout
           $timeout = new-timespan -Minutes $TimeoutMinutes
           $stopwatch = [diagnostics.stopwatch]::StartNew()
           $count = 0

           while ($stopwatch.elapsed -lt $timeout)
           {
              Write-Host "Process Running... {$count Sec}" -ForegroundColor Gray
              $count = $count + $CheckRateSeconds
              sleep -Seconds $CheckRateSeconds

              if($process.HasExited -eq $true)
              {
                   # Get output from stdout and stderr
                   $stdexit = $process.ExitCode.ToString()
                   $stderr = $process.StandardError.ReadToEnd()
                   $stdout = $process.StandardOutput.ReadToEnd()
                   return $stdexit, $stdout, $stderr
              }
           }

           Write-Host "Timeout reached... {$TimeoutMinutes Min}" -ForegroundColor Red
           return $null
       }

       # Wait process to finish = FALSE
       else
       {
           Write-Host "Process Running..." -ForegroundColor Gray

           sleep -Seconds 5
           if($process.HasExited)
           {
               # Get output from stdout and stderr
               $stdexit = $process.ExitCode.ToString()
               $stderr = $process.StandardError.ReadToEnd()
               $stdout = $process.StandardOutput.ReadToEnd()

               return $stdexit, $stdout, $stderr
           }

           
           return 0,"The process continues executing in another thread...`n",$null
       }
    }

    catch
    {
        Write-Host "Error running command {$Command}, Exception $_" -ForegroundColor Red
        return $null
    }
}

function Get-WorkItemTypes
{
    param
    (
        [Parameter(Mandatory=$true)]
        [string]$WorkingDirectory,
        [Parameter(Mandatory=$true)]
        [string]$Collection,
        [Parameter(Mandatory=$true)]
        [string]$TeamProject,
        [Parameter(Mandatory=$true)]
        [string]$WitadminPath
    )

    try
    {
        $Command = "$WitadminPath listwitd /collection:$Collection /p:$TeamProject"
        $CmdResult = Run-CmdCommand -Command $Command -WorkingDirectory $WorkingDirectory -WaitProcessToFinish $true -CheckRateSeconds "3" -TimeoutMinutes "5"
    }
    catch
    {
        Write-Host "Error running command {$Command}, Exception $_" -ForegroundColor Red
        return $null
    }

}

function Test-WitadminPath
{
    param
    (
        [string]$Path
    )

    if(Test-Path "$Path\witadmin.exe"){return $Path}
    elseif(Test-Path "C:\Program Files (x86)\Microsoft Visual Studio 15.0\Common7\IDE\witadmin.exe"){ return "C:\Program Files (x86)\Microsoft Visual Studio 15.0\Common7\IDE"}
    elseif(Test-Path "C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\witadmin.exe"){ return "C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE"}
    elseif(Test-Path "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\witadmin.exe"){ return "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE"}
    elseif(Test-Path "C:\Program Files (x86)\Microsoft Visual Studio 11.0\Common7\IDE\witadmin.exe"){ return "C:\Program Files (x86)\Microsoft Visual Studio 11.0\Common7\IDE"}
    elseif(Test-Path "C:\Program Files (x86)\Microsoft Visual Studio 10.0\Common7\IDE\witadmin.exe"){ return "C:\Program Files (x86)\Microsoft Visual Studio 10.0\Common7\IDE"}
    else
    {
        Write-Host "Invalid Witadmin Path, witadmin.exe not found in the specified path" -ForegroundColor Red
        Exit 500
    }
}

############################ SCRIPT #############################


# Test or Assign WitadminPath

$WitadminPath = Test-WitadminPath -Path $WitadminPath


# Initialize Workspace (Clone Repository)

$workspace = (Get-Location).ToString() + "\workspace-" + (Get-Date -Format "yyyyMMddHHmmss")
Write-Host "Cloning Backup Repository to {$workspace}" -ForegroundColor Yellow
$CloneRepositoryCommand = "git clone $BackupRepository `"$workspace`""
$CommandResponse = Run-CmdCommand -Command $CloneRepositoryCommand -WorkingDirectory (Get-Location) -WaitProcessToFinish $true -CheckRateSeconds "5" -TimeoutMinutes "5"
Remove-Item $workspace\* -exclude .git -Recurse


# Get Backups for Each Project

foreach($TfsProject in $TfsTeamProjects.Split(','))
{

    # Create Directory for Backups

    if(-not(Test-Path $workspace\$TfsProject))
    {
        Write-Host "Creating directory for project backups {$workspace\$TfsProject}" -ForegroundColor Yellow
        $Output = New-Item -ItemType Directory $workspace\$TfsProject
    }


    # Get WorkItems Types

    $GetWorkItemsTypesCommand = "witadmin listwitd /collection:$TfsCollectionUrl /p:$TfsProject"
    Write-Host "Running {$GetWorkItemsTypesCommand}" -ForegroundColor Yellow
    $CommandResponse = Run-CmdCommand -Command $GetWorkItemsTypesCommand -WorkingDirectory $WitadminPath -WaitProcessToFinish $true -CheckRateSeconds "5" -TimeoutMinutes "5"
    $WorkItemsTypes = $CommandResponse[1]


    # Export WorkItems Configurations

    foreach($WorkItemsType in $WorkItemsTypes.Split("`n"))
    {
        if($WorkItemsType)
        {
            $ExportWorkItemCommand = "witadmin exportwitd /collection:$TfsCollectionUrl /p:`"$TfsProject`" /n:`"$WorkItemsType`" /f:`"$workspace\$TfsProject\$WorkItemsType.xml`""
            Write-Host "Exporting $WorkItemsType, Running {$ExportWorkItemCommand}" -ForegroundColor Yellow
            $CommandResponse = Run-CmdCommand -Command $ExportWorkItemCommand -WorkingDirectory $WitadminPath -WaitProcessToFinish $true -CheckRateSeconds "5" -TimeoutMinutes "5"
        } 
    }

    # Export Project Configurations

    $ExportCategoriesCommand = "witadmin exportcategories /collection:$TfsCollectionUrl /p:`"$TfsProject`" /f:`"$workspace\$TfsProject\Categories.xml`""
    Write-Host "Exporting $WorkItemsType, Running {$ExportCategoriesCommand}" -ForegroundColor Yellow
    $CommandResponse = Run-CmdCommand -Command $ExportCategoriesCommand -WorkingDirectory $WitadminPath -WaitProcessToFinish $true -CheckRateSeconds "5" -TimeoutMinutes "5"

    $ExportProcessConfigCommand = "witadmin exportprocessconfig /collection:$TfsCollectionUrl /p:`"$TfsProject`" /f:`"$workspace\$TfsProject\ProcessConfig.xml`""
    Write-Host "Exporting $WorkItemsType, Running {$ExportProcessConfigCommand}" -ForegroundColor Yellow
    $CommandResponse = Run-CmdCommand -Command $ExportProcessConfigCommand -WorkingDirectory $WitadminPath -WaitProcessToFinish $true -CheckRateSeconds "5" -TimeoutMinutes "5"
}


# Save Backup

$GitAddCommand = "git add -A"
Write-Host "Running Command {$GitAddCommand}" -ForegroundColor Yellow
$CommandResponse = Run-CmdCommand -Command $GitAddCommand -WorkingDirectory $workspace -WaitProcessToFinish $true -CheckRateSeconds "5" -TimeoutMinutes "5"
    
$GitCommitCommand = "git commit -m `"Backup $(Get-Date -Format "dd-MM-yyyy HH:mm")`""
Write-Host "Running Command {$GitCommitCommand}" -ForegroundColor Yellow
$CommandResponse = Run-CmdCommand -Command $GitCommitCommand -WorkingDirectory $workspace -WaitProcessToFinish $true -CheckRateSeconds "5" -TimeoutMinutes "5"

if($CommandResponse[0] -ne 1)
{
    $GitPushCommand = "git push"
    Write-Host "Running Command {$GitPushCommand}" -ForegroundColor Yellow
    $CommandResponse = Run-CmdCommand -Command $GitPushCommand -WorkingDirectory $workspace -WaitProcessToFinish $true -CheckRateSeconds "5" -TimeoutMinutes "5"

    if($CommandResponse[0] -eq 0) { Write-Host "Backup Success" -ForegroundColor Green }
    else { Write-Host "Backup Failed" -ForegroundColor Red }
}
else
{
    Write-Host "No changes since the last backup" -ForegroundColor Green
}


# Remove Temporal Workspace

Remove-Item $workspace -Recurse -Force
