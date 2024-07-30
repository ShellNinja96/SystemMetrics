#================================================================================================================================================#
#=CHANGELOG======================================================================================================================================#

<# 30.07.2024
 - ADDED VARIABLE TO SPECIFY WHICH GPU CORE TO MONITOR
#>

<# 28.03.2024
 - CHANGED DISCLAIMER TO SHOW OHM REPOSITORY INSTEAD OF WEBSITE SINCE IT HAS BEEN DOWN FOR DAYS
 - ADDED OPTION TO DISABLE LINE 2
 - MADE CODE MORE CONCISE BY TURNING ALL OHM QUERY FUNCTIONS (EXCEPT CPU CLOCK) INTO A SINGLE FUNCTION
 - MADE THE RETRIEVAL OF METRICS DEPENDENT ON THEIR VISIBILITY SETTINGS, AVOIDING UNNECESSARY CODE EXECUTIONS
 - ADDED NETWORK ADAPTER ERROR CHECKING SECTION
 - TURNED GETDOWNLOADSPEED AND GETUPLOADSPEED INTO A SINGLE FUNCTION AVOIDING CODE REPETITION AND MINIMIZING CONSEQUENT SLEEP TIME TO 1 SECOND
 - ADDED MAINLOOPSLEEP FUNCTION TO AVOID HIGH CPU DEMAND
 - IMPROVED ON CONSOLE SIZE MANAGEMENT FOR SITUATIONS IN WHICH THE USER PREFERS THE SCRIPT TO NOT CHANGE CONSOLE SIZE DURING MAIN LOOP EXECUTION
#>

#================================================================================================================================================#
#=DISCLAIMER=====================================================================================================================================#

$disclaimer = @"
Disclaimer:
This script is provided as-is, without any warranty or support. 
The user assumes all responsibility for its use and any consequences that may arise from it. 
The script author takes no responsibility for any damages, data loss, or other issues caused by the script's use.

Usage:
This script intends to provide a customizable hardware monitoring tool in a console environment.
This script was made to be run on Windows Console Host as administrator.
Some functionality like the automatic adjustement of the Console Host window size is not compatible with Windows Terminal.
Users are encouraged to review and modify the script's settings in the Settings section according to their requirements. 
Modifying anything beyond the Settings section may break the script's functionality and is generally not recommended.

Requirements:
This script requires the OpenHardwareMonitor application to be installed on the system. 
Please ensure that OpenHardwareMonitor is installed before using this script. 
OpenHardwareMonitor can be downloaded from its official repository: https://github.com/openhardwaremonitor

Repository:
The official repository for this script can be found on GitHub: https://github.com/ShellNinja96/SystemMetrics

"@

$shellNinja96 = @"
  ___ _        _ _ _  _ _       _      ___  __ 
 / __| |_  ___| | | \| (_)_ _  (_)__ _/ _ \/ / 
 \__ \ ' \/ -_) | | .' | | ' \ | / _' \_, / _ \
 |___/_||_\___|_|_|_|\_|_|_||_|/ \__,_|/_/\___/
                             |__/            
"@

#================================================================================================================================================#
#=SETTINGS=======================================================================================================================================#

<# SYSTEM SPECIFIC #>
Set-ExecutionPolicy      Restricted                      # Reset system script execution policy
$networkAdapterName      = "vEthernet (External Switch)" # Name of the Network Adapter whose speeds should be monitored
$gpuName                 = "/nvidiagpu/0"                # Name of the gpu you want to monitor. In case of doubt run the following command -> Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class "Sensor" | Select-Object Name, Parent, SensorType, Value
$createStartupShortcut   = $false                        # Creates a shortcut for this script in the startup folder
$replaceExistingShortcut = $false                        # If the shortcut already exists it gets replaced by a new one. $createStartupShortcut must also be set to true
$mainLoopSleepDuration   = 1                             # Sleep time in seconds for each iteration of the main loop

<# FLAG THRESHOLDS #>
$cpuTempThreshold = 90.0                                 # Temperature in celsius at which the CPU temperature flag should be triggered
$cpuLoadThreshold = 90.0                                 # Load percentage at which the CPU load flag should be triggered
$gpuTempThreshold = 80.0                                 # Temperature in celsius at which the GPU temperature flag should be triggered
$gpuCoreThreshold = 90.0                                 # Load percentage at which the GPU core load flag should be triggered
$gpuMemThreshold  = 90.0                                 # Load percentage at which the GPU memory load flag should be triggered
$ramLoadThreshold = 60                                   # Load in GB at which the flag should be triggered

<# PERSONALIZATION #>
$indentation             = "  "                          # Indentation displayed at the beginning of each line
$spacing                 = " "                           # Spacing displayed between each metric
$flagColor               = "Red"                         # Color representing an exceeded metric threshold
$showLine2               = $true                         # Enable/disable line of dashes that separates the metric names from their values
$cpuTempMetricName       = "CPUtemp"                     # String displayed above the CPU temperature metric
$cpuTempMetricShow       = $true                         # Show the CPU temperature metric
$cpuLoadMetricName       = "CPUload"                     # String displayed above the CPU load metric
$cpuLoadMetricShow       = $true                         # Show the CPU load metric
$cpuClockMetricName      = "CPUclck"                     # String displayed above the CPU clock speed metric
$cpuClockMetricShow      = $true                         # Show the CPU clock speed metric
$gpuTempMetricName       = "GPUtemp"                     # String displayed above the GPU temperature metric
$gpuTempMetricShow       = $true                         # Show the GPU temperature metric
$gpuCoreMetricName       = "GPUcore"                     # String displayed above the GPU core load metric
$gpuCoreMetricShow       = $true                         # Show the GPU core load metric
$gpuMemMetricName        = "GPUmem"                      # String displayed above the GPU memory load metric
$gpuMemMetricShow        = $true                         # Show the GPU memory load metric
$ramLoadMetricName       = "RAMload"                     # String displayed above the RAM memory load metric
$ramLoadMetricShow       = $true                         # Show the RAM memory load metric
$downloadSpeedMetricName = "NETdown"                     # String displayed above the download speed metric
$downloadSpeedMetricShow = $true                         # Show the download speed metric
$uploadSpeedMetricName   = "NETuplo"                     # String displayed above the upload speed metric
$uploadSpeedMetricShow   = $true                         # Show the upload speed metric
$showDisclaimer          = $false                        # Enable/disable the disclaimer displayment
$showSexySplashScreen    = $false                        # Enable/disable a very sexy splash screen :)

<# CONSOLE SETTINGS #>
$consoleWindowTitle = "ShellNinja96 SystemMetrics"       # Change the console host window title in which this script is being executed
$changeConsoleSize  = $true                              # Change to $false if you don't want the script to change the console window/buffer size during main loop execution
$forceConsoleSize   = $false                             # Change to $true if you'd like to force the console window/buffer size with the values below. $changeConsoleSize must also be set to $true
$forceConsoleWidth  = 74                                 # Force the console window/buffer width. $changeConsoleSize and $forceConsoleSize must be set to $true
$forceConsoleHeight = 3                                  # Force the console window/buffer height. $changeConsoleSize and $forceConsoleSize must be set to $true

<# ADVANCED SETTINGS #>     
$performVariableValidation = $true                       # Enable/disable variable type validation
$debug                     = $false                      # Enable/disable debugging 


#================================================================================================================================================#
#=ERROR=CHECKING=================================================================================================================================#

# Declares $errorFound variable as false
$errorFound = $false

# Validates all settings variable types, throws all found errors, pauses and exits if any was found 
if ( $performVariableValidation ) {

    <# BOOLEAN TYPE SECTION #>
    if (-not($createStartupShortcut   -is [bool]))   { Write-Host "Settings Error: `$createStartupShortcut must be of type bool"     -ForegroundColor Red; $errorFound = $true }
    if (-not($replaceExistingShortcut -is [bool]))   { Write-Host "Settings Error: `$replaceExistingShortcut must be of type bool"   -ForegroundColor Red; $errorFound = $true }
    if (-not($cpuTempMetricShow       -is [bool]))   { Write-Host "Settings Error: `$cpuTempMetricShow must be of type bool"         -ForegroundColor Red; $errorFound = $true }
    if (-not($cpuLoadMetricShow       -is [bool]))   { Write-Host "Settings Error: `$cpuLoadMetricShow must be of type bool"         -ForegroundColor Red; $errorFound = $true }
    if (-not($cpuClockMetricShow      -is [bool]))   { Write-Host "Settings Error: `$cpuClockMetricShow must be of type bool"        -ForegroundColor Red; $errorFound = $true }
    if (-not($gpuTempMetricShow       -is [bool]))   { Write-Host "Settings Error: `$gpuTempMetricShow must be of type bool"         -ForegroundColor Red; $errorFound = $true }
    if (-not($gpuCoreMetricShow       -is [bool]))   { Write-Host "Settings Error: `$gpuCoreMetricShow must be of type bool"         -ForegroundColor Red; $errorFound = $true }
    if (-not($gpuMemMetricShow        -is [bool]))   { Write-Host "Settings Error: `$gpuMemMetricShow must be of type bool"          -ForegroundColor Red; $errorFound = $true }
    if (-not($ramLoadMetricShow       -is [bool]))   { Write-Host "Settings Error: `$ramLoadMetricShow must be of type bool"         -ForegroundColor Red; $errorFound = $true }
    if (-not($downloadSpeedMetricShow -is [bool]))   { Write-Host "Settings Error: `$downloadSpeedMetricShow must be of type bool"   -ForegroundColor Red; $errorFound = $true }
    if (-not($uploadSpeedMetricShow   -is [bool]))   { Write-Host "Settings Error: `$uploadSpeedMetricShow must be of type bool"     -ForegroundColor Red; $errorFound = $true }
    if (-not($changeConsoleSize       -is [bool]))   { Write-Host "Settings Error: `$changeConsoleSize must be of type bool"         -ForegroundColor Red; $errorFound = $true }
    if (-not($forceConsoleSize        -is [bool]))   { Write-Host "Settings Error: `$forceConsoleSize must be of type bool"         -ForegroundColor Red; $errorFound = $true }
    if (-not($showDisclaimer          -is [bool]))   { Write-Host "Settings Error: `$showDisclaimer must be of type bool"            -ForegroundColor Red; $errorFound = $true }
    if (-not($showSexySplashScreen    -is [bool]))   { Write-Host "Settings Error: `$showSexySplashScreen must be of type bool"      -ForegroundColor Red; $errorFound = $true }
    if (-not($debug                   -is [bool]))   { Write-Host "Settings Error: `$debug must be of type bool"                     -ForegroundColor Red; $errorFound = $true }

    <# STRING TYPE SECTION #>
    if (-not($networkAdapterName      -is [string])) { Write-Host "Settings Error: `$networkAdapterName must be of type string"      -ForegroundColor Red; $errorFound = $true }
    if (-not($indentation             -is [string])) { Write-Host "Settings Error: `$indentation must be of type string"             -ForegroundColor Red; $errorFound = $true }
    if (-not($spacing                 -is [string])) { Write-Host "Settings Error: `$spacing must be of type string"                 -ForegroundColor Red; $errorFound = $true }
    if (-not($cpuTempMetricName       -is [string])) { Write-Host "Settings Error: `$cpuTempMetricName must be of type string"       -ForegroundColor Red; $errorFound = $true }
    if (-not($cpuLoadMetricName       -is [string])) { Write-Host "Settings Error: `$cpuLoadMetricName must be of type string"       -ForegroundColor Red; $errorFound = $true }
    if (-not($cpuClockMetricName      -is [string])) { Write-Host "Settings Error: `$cpuClockMetricName must be of type string"      -ForegroundColor Red; $errorFound = $true }
    if (-not($gpuTempMetricName       -is [string])) { Write-Host "Settings Error: `$gpuTempMetricName must be of type string"       -ForegroundColor Red; $errorFound = $true }
    if (-not($gpuCoreMetricName       -is [string])) { Write-Host "Settings Error: `$gpuCoreMetricName must be of type string"       -ForegroundColor Red; $errorFound = $true }
    if (-not($gpuMemMetricName        -is [string])) { Write-Host "Settings Error: `$gpuMemMetricName must be of type string"        -ForegroundColor Red; $errorFound = $true }
    if (-not($ramLoadMetricName       -is [string])) { Write-Host "Settings Error: `$ramLoadMetricName must be of type string"       -ForegroundColor Red; $errorFound = $true }
    if (-not($downloadSpeedMetricName -is [string])) { Write-Host "Settings Error: `$downloadSpeedMetricName must be of type string" -ForegroundColor Red; $errorFound = $true }
    if (-not($uploadSpeedMetricName   -is [string])) { Write-Host "Settings Error: `$uploadSpeedMetricName must be of type string"   -ForegroundColor Red; $errorFound = $true }
    if (-not($consoleWindowTitle      -is [string])) { Write-Host "Settings Error: `$consoleWindowTitle must be of type string"      -ForegroundColor Red; $errorFound = $true }

    <# DOUBLE TYPE SECTION #>
    if (-not($cpuTempThreshold        -is [double])) { Write-Host "Settings Error: `$cpuTempThreshold must be of type double"        -ForegroundColor Red; $errorFound = $true }
    if (-not($cpuLoadThreshold        -is [double])) { Write-Host "Settings Error: `$cpuLoadThreshold must be of type double"        -ForegroundColor Red; $errorFound = $true }
    if (-not($gpuTempThreshold        -is [double])) { Write-Host "Settings Error: `$gpuTempThreshold must be of type double"        -ForegroundColor Red; $errorFound = $true }
    if (-not($gpuCoreThreshold        -is [double])) { Write-Host "Settings Error: `$gpuCoreThreshold must be of type double"        -ForegroundColor Red; $errorFound = $true }
    if (-not($gpuMemThreshold         -is [double])) { Write-Host "Settings Error: `$gpuMemThreshold must be of type double"         -ForegroundColor Red; $errorFound = $true }

    <# INTEGER TYPE SECTION #>
    if (-not($mainLoopSleepDuration   -is [int]))    { Write-Host "Settings Error: `$mainLoopSleepDuration must be of type integer"  -ForegroundColor Red; $errorFound = $true }
    if (-not($ramLoadThreshold        -is [int]))    { Write-Host "Settings Error: `$ramLoadThreshold must be of type integer"       -ForegroundColor Red; $errorFound = $true }
    if (-not($forceConsoleWidth       -is [int]))    { Write-Host "Settings Error: `$forceConsoleWidth must be of type integer"     -ForegroundColor Red; $errorFound = $true }
    if (-not($forceConsoleHeight      -is [int]))    { Write-Host "Settings Error: `$forceConsoleHeight must be of type integer"    -ForegroundColor Red; $errorFound = $true }

    <# COLOR TYPE SECTION #>
    if (-not(([System.Enum]::GetNames([System.ConsoleColor]) -contains $flagColor))) { Write-Host "Settings Error: `$flagColor must be a valid color string" -ForegroundColor Red; Write-Host "Valid color strings: $([System.Enum]::GetNames([System.ConsoleColor]) -join ', ')"; $errorFound = $true }

}

# Validates $networkAdapterName and assigns $networkAdapter for later usage if network related metrics visibility is enabled
if ( $downloadSpeedMetricShow -or $uploadSpeedMetricShow) {
    $networkAdapter = Get-NetAdapter | Where-Object {$_.Name -eq $networkAdapterName}
    if (-not $networkAdapter) {
        Write-Host "Settings Error: `$networkAdapterName must be a valid NetAdapter.Name string" -ForegroundColor Red
        Write-Host "Valid NetAdapter.Name strings: $((Get-NetAdapter | Select-Object -ExpandProperty Name) -join ", ")"
        $errorFound = $true
    }
}

# If an error was found, pauses the script and then exits
if ($errorFound) { pause; exit }

#================================================================================================================================================#
#=FUNCTIONS======================================================================================================================================#

function SetConsoleSize {
    Param(
        [parameter(Mandatory=$true)]
        [int]$Height,
        [parameter(Mandatory=$true)]
        [int]$Width
    )

    # Retrieves the current buffer and window size of the console window
    $ConBuffer  = $host.ui.RawUI.BufferSize
    $ConSize    = $host.ui.RawUI.WindowSize
    $currWidth  = $ConSize.Width
    $currHeight = $ConSize.Height

    # if height is too large, set to max allowed size
    if ($Height -gt $host.UI.RawUI.MaxPhysicalWindowSize.Height) {
        $Height = $host.UI.RawUI.MaxPhysicalWindowSize.Height
    }

    # if width is too large, set to max allowed size
    if ($Width -gt $host.UI.RawUI.MaxPhysicalWindowSize.Width) {
        $Width = $host.UI.RawUI.MaxPhysicalWindowSize.Width
    }

    # If the Buffer is wider than the new console setting, first reduce the width
    If ($ConBuffer.Width -gt $Width ) {
    $currWidth = $Width
    }
    # If the Buffer is higher than the new console setting, first reduce the height
    If ($ConBuffer.Height -gt $Height ) {
        $currHeight = $Height
    }
    # initial resizing if needed
    $host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.size($currWidth,$currHeight)

    # Set the Buffer
    $host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.size($Width,$Height)

    # Now set the WindowSize
    $host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.size($Width,$Height)

}

#================================================================================================================================================#

function GetAutoConsoleWidth {
    Param(
        [parameter(Mandatory=$true)]
        [array]$hardwareMetrics
    )

    # Initialize the variable to hold the total length of line2
    $line2Length = 0

    # Loop through each hardware metric object in the array
    for ($i = 0; $i -lt $hardwareMetrics.Length; $i++) {
        # If the current index of $hardwareMetrics Show property is $true, add the length of the Line2 property string to $line2Length
        if ( $hardwareMetrics[$i].Show ) { $line2Length += $hardwareMetrics[$i].Line2.Length }
    }

    # Add the additional length required for indentation on both sides of Line2
    $line2Length += ($indentation.Length * 2)

    # Return the calculated total console width required for correctly displaying all hardware metrics
    return $line2Length

}

#================================================================================================================================================#

function ShowDisclaimer {

    Clear-Host
    SetConsoleSize -Height 22 -Width 122
    Write-Host $disclaimer
    pause

}

#================================================================================================================================================#

function ShowSexySplashScreen {

    Clear-Host
    SetConsoleSize -Height 5 -Width 48
    
    $sexyColors = @( "Blue", "Green", "Cyan", "Red", "Magenta", "Yellow" )

    for ($i = 1; $i -le 7; $i++) {
        foreach ($color in $sexyColors) {
            [Console]::SetCursorPosition(0, 0)
            Write-Host $shellNinja96 -ForegroundColor $color -NoNewLine
            Start-Sleep -Milliseconds 75
        }
    }

}

#================================================================================================================================================#

function CreateStartupShortcut {
    param (
        [parameter(Mandatory=$true)]
        [string]$targetPath
    )

    SetConsoleSize -Height 9 -Width 120

    # Get the Startup folder path
    $startupFolderPath = [Environment]::GetFolderPath("Startup")

    # Create the Startup folder if it doesn't exist
    if ($startupFolderPath -eq "") {
        $startupFolderPath = ([Environment]::GetFolderPath("Programs"))
        $startupFolderPath += "\Startup"
        New-Item -ItemType Directory -Path $startupFolderPath
        Clear-Host
    }

    # Define the path for the shortcut file
    $shortcutFilePath = [System.IO.Path]::Combine($startupFolderPath, "ShellNinja96SystemMetrics.lnk")

    # If the shorcut doesn't exist, or if the shortcut exists and $replaceExistingShortcut set to $true
    if ((-not (Test-Path -Path $shortcutFilePath)) -or ((Test-Path -Path $shortcutFilePath) -and $replaceExistingShortcut)) {

        # Define the command to be executed by PowerShell
        $command = "-NoExit -Command `"Set-ExecutionPolicy Bypass; & '$targetPath'; Clear-Host;`""

        # Create a WScript Shell object
        $shell = New-Object -ComObject WScript.Shell

        # Create a shortcut object
        $shortcut = $shell.CreateShortcut($shortcutFilePath)

        # Set the properties of the shortcut and save it
        $shortcut.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $shortcut.Arguments = $command
        $shortcut.IconLocation = "%SystemRoot%\System32\SHELL32.dll,12"
        $shortcut.Save()

        # Let user know that a shortcut has been created and that it should be ran as Administrator
        Write-Host "Shortcut created in:" -ForegroundColor Yellow
        Write-Host "$startupFolderPath`n"
        Write-Host "Modify its properties so that it runs as administrator." -ForegroundColor Yellow
        Write-Host "Right Mouse Click > Properties > Advanced > Run as administrator > OK > OK`n"
        Write-Host "Launch the script from the shortcut." -ForegroundColor Yellow
        Start-Process explorer.exe "/select,$startupFolderPath\ShellNinja96SystemMetrics.lnk"
        pause
        exit
    }

}

#================================================================================================================================================#

function StartOpenHardwareMonitor {

    # Checks if OpenHardwareMonitor process is running and proceeds accordingly
    $process = Get-Process -Name OpenHardwareMonitor*

    if ($process) {

        # Nothing to do here...
        if ($debug) { Write-Host "OpenHardwareMonitor is already running." -ForegroundColor Yellow; pause}

    } else {

        # Define the path to the OpenHardwareMonitor software
        $openHardwareMonitorPath = "C:\Program Files\OpenHardwareMonitor\OpenHardwareMonitor.exe"

        # Test the path
        if (-not(Test-Path -Path $openHardwareMonitorPath)) {
            Write-Host "OpenHardwareMonitor not found!" -ForegroundColor Red
            Write-Host $openHardwareMonitorPath
            pause
            exit
        }

        # Starts the OpenHardwareMonitor process
        Start-Process -FilePath $openHardwareMonitorPath -Verb RunAs

        # Waits for OpenHardwareMonitor process to be up and running
        while ($true) {

            $process = Get-Process -Name OpenHardwareMonitor*
            if ($process) {
                if ($debug) { Write-Host "OpenHardwareMonitor has been started." -ForegroundColor Yellow; pause }
                Clear-Host
                break
            }
            else {
                if ($debug) { Write-Host "Waiting for OpenHardwareMonitor process to start" -ForegroundColor Yellow }
                Start-Sleep -Seconds 1
            }

        }

    }

}

#================================================================================================================================================#

function GetOHMMetric {
    param (
        [parameter(Mandatory=$true)]
        [string]$query,
        [parameter(Mandatory=$true)]
        [int]$roundTo        
    )

    # Get the object from the passed $query parameter and then extract its value
    $value = (Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Query $query) | Select-Object -ExpandProperty Value

    # Return the rounded value based on $roundTo parameter.
    return [Math]::Round($value, $roundTo)

}

#===============================================================================================================================================#

function GetCPUClock {
    # Get the clock objects for CPU cores
    $clocks = Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Query "SELECT * FROM Sensor WHERE SensorType='Clock' AND Name LIKE '%CPU Core%'"

    # Extract the clock values
    $clockValues = $clocks | Select-Object -ExpandProperty Value

    # Get the highest clock value
    $clockValue = ($clockValues | Measure-Object -Maximum).Maximum

    # Return the highest clock value rounded to integer
    return [Math]::Round($clockValue)
}

#===============================================================================================================================================#

function GetNetworkSpeed {

    # Get-NetAdapterStatistics sleep 1 second and Get-NetAdapterStatistics again.
    $networkStatistics1 = Get-NetAdapterStatistics -Name $networkAdapter.Name
    Start-Sleep -Seconds 1
    $networkStatistics2 = Get-NetAdapterStatistics -Name $networkAdapter.Name

    # Calculate speeds and convert them from bytes per second to megabits per second if metric visibility enabled in settings, else skip calculation, metric is $null
    if ( $downloadSpeedMetricShow ) { $downloadSpeed = [Math]::Round((($networkStatistics2.ReceivedBytes - $networkStatistics1.ReceivedBytes) * 8 / 1e6)) } else { $downloadSpeed = $null }
    if ( $uploadSpeedMetricShow   ) { $uploadSpeed   = [Math]::Round((($networkStatistics2.SentBytes     - $networkStatistics1.SentBytes    ) * 8 / 1e6)) } else { $uploadSpeed   = $null }

    # Return the NetworkSpeed as a custom object
    return [PSCustomObject]@{ Download = $downloadSpeed ; Upload = $uploadSpeed }

}

#================================================================================================================================================#

function GetHardwareMetrics {

    # Ensure OpenHardwareMonitor is running
    StartOpenHardwareMonitor

    # Get the OpenHardwareMonitor metrics whose visibility is enabled in settings 
    if ( $cpuClockMetricShow ) { $cpuClock = GetCPUClock                                                                                                        } else { $cpuClock = $null }
    if ( $cpuTempMetricShow  ) { $cpuTemp  = GetOHMMetric -query "SELECT * FROM Sensor WHERE SensorType='Temperature' AND Name LIKE '%CPU Package%'" -roundTo 2 } else { $cpuTemp  = $null }
    if ( $cpuLoadMetricShow  ) { $cpuLoad  = GetOHMMetric -query "SELECT * FROM Sensor WHERE SensorType='Load' AND Name LIKE '%CPU Total%'"          -roundTo 2 } else { $cpuLoad  = $null }        
    if ( $gpuTempMetricShow  ) { $gpuTemp  = GetOHMMetric -query "SELECT * FROM Sensor WHERE SensorType='Temperature' AND Name LIKE '%GPU%'"         -roundTo 2 } else { $gpuTemp  = $null }
    if ( $gpuMemMetricShow   ) { $gpuMem   = GetOHMMetric -query "SELECT * FROM Sensor WHERE SensorType='Load' AND Name LIKE '%GPU Memory%'"         -roundTo 2 } else { $gpuMem   = $null }
    if ( $ramLoadMetricShow  ) { $ramLoad  = GetOHMMetric -query "SELECT * FROM Sensor WHERE SensorType='Data' AND Name LIKE '%Used Memory%'"        -roundTo 2 } else { $ramLoad  = $null }
    if ( $gpuCoreMetricShow  ) {
        $query    = "SELECT * FROM Sensor WHERE SensorType='Load' AND Name LIKE '%GPU Core%' AND Parent LIKE '" + $gpuName + "'"
        $gpuCore  = GetOHMMetric -query $query -roundTo 0
    } else { $gpuCore  = $null }

    # Get the $networkSpeed custom object if network related metrics visibility is enabled in settings. Else the metrics are $null.
    if ( $downloadSpeedMetricShow -or $uploadSpeedMetricShow) { 
        $networkSpeed  = GetNetworkSpeed
        $downloadSpeed = $networkSpeed.Download
        $uploadSpeed   = $networkSpeed.Upload
    } else { $downloadSpeed = $null ; $uploadSpeed = $null }

    # Populate the $hardwareMetrics array
    $hardwareMetrics = @(
    <#[0]#> [PSCustomObject]@{ Space = $spacing; Name = $cpuTempMetricName;       Value = $cpuTemp;       Unit = "C";    Flag = $cpuTemp -ge $cpuTempThreshold; Line1 = ""; Line2 = ""; Line3 = ""; Show = $cpuTempMetricShow       },
    <#[1]#> [PSCustomObject]@{ Space = $spacing; Name = $cpuLoadMetricName;       Value = $cpuLoad;       Unit = "%";    Flag = $cpuLoad -ge $cpuLoadThreshold; Line1 = ""; Line2 = ""; Line3 = ""; Show = $cpuLoadMetricShow       },
    <#[2]#> [PSCustomObject]@{ Space = $spacing; Name = $cpuClockMetricName;      Value = $cpuClock;      Unit = "MHz";  Flag = $false;                         Line1 = ""; Line2 = ""; Line3 = ""; Show = $cpuClockMetricShow      },
    <#[3]#> [PSCustomObject]@{ Space = $spacing; Name = $gpuTempMetricName;       Value = $gpuTemp;       Unit = "C";    Flag = $gpuTemp -ge $gpuTempThreshold; Line1 = ""; Line2 = ""; Line3 = ""; Show = $gpuTempMetricShow       },
    <#[4]#> [PSCustomObject]@{ Space = $spacing; Name = $gpuCoreMetricName;       Value = $gpuCore;       Unit = "%";    Flag = $gpuCore -ge $gpuCoreThreshold; Line1 = ""; Line2 = ""; Line3 = ""; Show = $gpuCoreMetricShow       },
    <#[5]#> [PSCustomObject]@{ Space = $spacing; Name = $gpuMemMetricName;        Value = $gpuMem;        Unit = "%";    Flag = $gpuMem  -ge $gpuMemThreshold;  Line1 = ""; Line2 = ""; Line3 = ""; Show = $gpuMemMetricShow        },
    <#[6]#> [PSCustomObject]@{ Space = $spacing; Name = $ramLoadMetricName;       Value = $ramLoad;       Unit = "GB";   Flag = $ramLoad -ge $ramLoadThreshold; Line1 = ""; Line2 = ""; Line3 = ""; Show = $ramLoadMetricShow       },
    <#[7]#> [PSCustomObject]@{ Space = $spacing; Name = $downloadSpeedMetricName; Value = $downloadSpeed; Unit = "Mbps"; Flag = $false;                         Line1 = ""; Line2 = ""; Line3 = ""; Show = $downloadSpeedMetricShow },
    <#[8]#> [PSCustomObject]@{ Space = $spacing; Name = $uploadSpeedMetricName;   Value = $uploadSpeed;   Unit = "Mbps"; Flag = $false;                         Line1 = ""; Line2 = ""; Line3 = ""; Show = $uploadSpeedMetricShow   }
    )
    if ( $debug ) { Write-Host "`$hardwareMetrics populated:" -ForegroundColor Yellow ; Write-Host $hardwareMetrics ; pause }

    # Declare the $firstShownMetricFlag as true
    $firstShownMetricFlag = $true

    # Adjust and format the properties in the $hardwareMetrics array
    for ($i = 0; $i -lt $hardwareMetrics.Length; $i++) {

        # Iterate only on the metrics that have the Show property set to $true
        if ( $hardwareMetrics[$i].Show ) {

            # Populate the Line3 property with the Value and Unit property strings respectively
            $hardwareMetrics[$i].Line3 = $hardwareMetrics[$i].Value.ToString() + $hardwareMetrics[$i].Unit
            if ( $debug ) { Write-Host "`$hardwareMetrics[$i].Line3 populated:" -ForegroundColor Yellow ; Write-Host $hardwareMetrics[$i].Line3 }

            # Populate the Line1 property string with whitespaces if the length of Line3 property string is bigger than the Name property string
            if (($hardwareMetrics[$i].Line3).Length -gt ($hardwareMetrics[$i].Name).Length) {      

                # Calculates how many whitespaces should be populated and proceeds to populate them
                $missingSpaces = ($hardwareMetrics[$i].Line3).Length - ($hardwareMetrics[$i].Name).Length
                $hardwareMetrics[$i].Line1 = ( " " * $missingSpaces )
                if ( $debug ) { Write-Host "`$hardwareMetrics[$i].Line1 was populated with $missingSpaces missing spaces." -ForegroundColor Yellow }
            } else {
                if ( $debug ) { Write-Host "`$hardwareMetrics[$i].Line1 was not populated with missing spaces." -ForegroundColor Yellow }
            }

            # Add the Name property string to the beginning of the Line1 property string
            $hardwareMetrics[$i].Line1 = $hardwareMetrics[$i].Name + $hardwareMetrics[$i].Line1
            if ( $debug ) { Write-Host "`$hardwareMetrics[$i].Line1 was changed to:" -ForegroundColor Yellow ; Write-Host $hardwareMetrics[$i].Line1 }

            # Add whitespaces to the end of the Line3 property string if the length of Line1 property string is bigger than the Line3 property string
            if (($hardwareMetrics[$i].Line1).Length -gt ($hardwareMetrics[$i].Line3).Length) {
                $missingSpaces = ($hardwareMetrics[$i].Line1).Length - ($hardwareMetrics[$i].Line3).Length
                $hardwareMetrics[$i].Line3 += ( " " * $missingSpaces )
                if ( $debug ) { Write-Host "$missingSpaces missing spaces were added to the end of `$hardwareMetrics[$i].Line3" -ForegroundColor Yellow ; Write-Host $hardwareMetrics[$i].Line3 }
            } else {
                if ( $debug ) { Write-Host "No white spaces were added to the end of `$hardwareMetrics[$i].Line3" -ForegroundColor Yellow }
            }

            # Populate Line3 property with dashes. The amount of dashes should be equal to the length of the Line1 property string
            $hardwareMetrics[$i].Line2 = ( "-" * ($hardwareMetrics[$i].Line1).Length )
            if ( $debug ) { Write-Host "`$hardwareMetrics[$i].Line2 was populated with dashes:" -ForegroundColor Yellow ; Write-Host $hardwareMetrics[$i].Line2 }

            # Ensures that the first metric in $hardwareMetrics array that has the property Show set to $true has no spacing behind it
            if ( $firstShownMetricFlag ) { $hardwareMetrics[$i].Space = "" ; $firstShownMetricFlag = $false }

            # Adds the Space property string to the beginning of Line1, Line2 and Line3 property strings
            $hardwareMetrics[$i].Line1 = $hardwareMetrics[$i].Space + $hardwareMetrics[$i].Line1
            $hardwareMetrics[$i].Line2 = $hardwareMetrics[$i].Space + $hardwareMetrics[$i].Line2
            $hardwareMetrics[$i].Line3 = $hardwareMetrics[$i].Space + $hardwareMetrics[$i].Line3
            if ( $debug ) { Write-Host "Indentation/spacing was added to the beginning of `$hardwareMetrics[$i].Line1, `$hardwareMetrics[$i].Line2 and `$hardwareMetrics[$i].Line3" -ForegroundColor Yellow }

        }

    } if ( $debug ) { Write-Host "`n`$hardwareMetrics was adjusted and formatted:" -ForegroundColor Yellow ; Write-Host $hardwareMetrics ; pause ; Clear-Host }

    # Returns the fully populated and formatted $hardwareMetrics array
    return $hardwareMetrics

}

#================================================================================================================================================#

function DisplayHardwareMetrics {
    param (
        [parameter(Mandatory=$true)]
        [array]$hardwareMetrics
    )

    # Write Indentation
    Write-Host $indentation -NoNewLine

    # Write Line1
    for ($i = 0; $i -lt $hardwareMetrics.Length; $i++) {
        if ( $hardwareMetrics[$i].Show ) {
            if ( $hardwareMetrics[$i].Flag ) {
                Write-Host $hardwareMetrics[$i].Line1 -NoNewLine -ForegroundColor $flagColor 
            } else {
                Write-Host $hardwareMetrics[$i].Line1 -NoNewLine
            }
        }
    }

    # New Line
    Write-Host ""

    # if $showLine2 enabled in settings
    if ( $showLine2 ) {

        # Write Indentation
        Write-Host $indentation -NoNewLine

        # Write Line2
        for ($i = 0; $i -lt $hardwareMetrics.Length; $i++) {
            if ( $hardwareMetrics[$i].Show ) {
                if ( $hardwareMetrics[$i].Flag ) {
                    Write-Host $hardwareMetrics[$i].Line2 -NoNewLine -ForegroundColor $flagColor 
                } else {
                    Write-Host $hardwareMetrics[$i].Line2 -NoNewLine
                }
            }
        }

        # New Line
        Write-Host ""

    }

    # Write Indentation
    Write-Host $indentation -NoNewLine

    # Write Line3
    for ($i = 0; $i -lt $hardwareMetrics.Length; $i++) {
        if ( $hardwareMetrics[$i].Show ) {
            if ( $hardwareMetrics[$i].Flag ) {
                Write-Host $hardwareMetrics[$i].Line3 -NoNewLine -ForegroundColor $flagColor 
            } else {
                Write-Host $hardwareMetrics[$i].Line3 -NoNewLine
            }
        }
    }

}

#================================================================================================================================================#

function MainLoopSleep {
    param (
        [parameter(Mandatory=$true)]
        [int]$sleepDuration
    )

    # Subtracts 1 second from $sleepDuration if network related metrics are enabled since GetNetworkSpeed function already has a 1 second sleep session
    if ( $downloadSpeedMetricShow -or $uploadSpeedMetricShow ) { $sleepDuration -=1 }

    # Make sure that the $sleepDuration is never less than 0 seconds.
    if ( $sleepDuration -lt 0 ) { $sleepDuration = 0}

    # Sleepy time
    Start-Sleep -Seconds $sleepDuration

}

#================================================================================================================================================#
#=MAIN=SCRIPT====================================================================================================================================#

# Set console window title
$host.ui.RawUI.WindowTitle = $consoleWindowTitle

# If the user prefers that the console size is not changed by the script during main loop execution, store the console size now and restore it before the first iteration
if (-not $changeConsoleSize) { $storedConsoleSize = $host.ui.rawui.WindowSize }

# Shows the Disclaimer for this script if enabled in settings
if ( $showDisclaimer ) { ShowDisclaimer }

# Create startup shortcut for this scripts if enabled in settings
if ( $createStartupShortcut ) { CreateStartupShortcut -targetPath $MyInvocation.MyCommand.Definition }

# Shows the Sexy Splash Screen script if enabled in settings
if ( $showSexySplashScreen ) { ShowSexySplashScreen }

# Restore console size stored in the beginning of the MAIN SCRIPT section
if (-not $changeConsoleSize) { SetConsoleSize -Height $storedConsoleSize.Height -Width $storedConsoleSize.Width }

# Clean Start :)
Clear-Host

# Main Loop
while ($true) {

    # Populates and Formats the $hardwareMetrics array
    $hardwareMetrics = GetHardwareMetrics

    # Changes the console size either manually or automatically based on user settings
    if ($changeConsoleSize) { 
        if ($forceConsoleSize) {
            # Manual console size, as defined in settings
            try { SetConsoleSize -Height $forceConsoleHeight -Width $forceConsoleWidth } catch {}
        } else {
            # Automatic console size
            SetConsoleSize -Height 3 -Width (GetAutoConsoleWidth -hardwareMetrics $hardwareMetrics)
        }
    } 

    # Displays the formatted $hardwareMetrics array.
    DisplayHardwareMetrics -hardwareMetrics $hardwareMetrics
    if ( $debug ) { pause }
    
    # Reset the cursor position for the next iteration of the Main Loop, produces a result similiar to Clear-Host but without the distracting blinking.
    [Console]::SetCursorPosition(0, 0)

    # Run the MainLoopSleep function
    MainLoopSleep -sleepDuration $mainLoopSleepDuration
    
}

#================================================================================================================================================#