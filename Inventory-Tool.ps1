# Ask the user to choose output format
$userSelection = Read-Host "Choose the output format:`n1. JSON (Default)`n2. CSV`nPlease enter the number of your selection, or leave empty for default:"

switch ($userSelection) {
    "2" { $OutputFormat = "CSV" }
    default { $OutputFormat = "JSON" }
}


function Write-Info {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Message
    )
    process {
        Write-Host "[INFO] $Message" -ForegroundColor Blue
    }
}

function Write-Success {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Message
    )
    process {
        Write-Host "[SUCCESS] $Message" -ForegroundColor Green
    }
}

function Write-Error {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Message
    )
    process {
        Write-Host "[ERROR] $Message" -ForegroundColor Red
    }
}

# Get the Major version of PowerShell
$PSVersion = $PSVersionTable.PSVersion.Major

if ($PSVersion -ge 3) {
    Write-Success "Running on PowerShell version $PSVersion. Ordered hashtable can be used."
    [Environment]::NewLine
}
else {
    Write-Error "Running on PowerShell version $PSVersion. Ordered hashtable is not supported. The CSV Option wont work this way."
}


<#
.SYNOPSIS
Displays an inline progress bar while executing a script block.

.DESCRIPTION
The Write-InlineProgress function is used to display an inline progress bar in the console while a script block is being executed. The duration of the progress bar is calculated based on the execution time of the script block.

.PARAMETER Task
A script block that represents the task to be executed.

.PARAMETER Activity
A string that represents the name of the activity. This name is displayed before the progress bar.

.EXAMPLE
Write-InlineProgress -Task { Start-Sleep -Seconds 5 } -Activity "Performing operation"

This command executes a sleep operation for 5 seconds and displays a progress bar with the activity name "Performing operation".

.NOTES
The progress bar is displayed with a delay calculated based on the execution time of the task. The delay ensures that the progress bar completes roughly at the same time as the task.

The function also includes a one second delay after the progress bar completes before it returns the result of the task.
#>


function Write-InlineProgress([scriptblock]$Task, [string]$Activity = "Processing") {
    $activityMaxWidth = 20
    $progressMaxWidth = 50
    $progressBar = ''

    Write-Host -NoNewline ("{0} " -f $Activity.PadRight($activityMaxWidth))

    # Run task first to get execution time
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $taskResult = & $Task
    } catch {
        Write-Host "`r[$($progressBar.PadRight($progressMaxWidth, '#'))] Failed!" -ForegroundColor Red
        Write-Error "Failed to execute task: $_"
        return
    }
    $stopwatch.Stop()

    # Calculate sleep interval based on task execution time
    $sleepInterval = $stopwatch.Elapsed.TotalMilliseconds / $progressMaxWidth

    for ($i = 1; $i -le $progressMaxWidth; $i++) {
        $progressBar += '#'
        Write-Host -NoNewline "`r[$($progressBar.PadRight($progressMaxWidth))]"
        Start-Sleep -Milliseconds $sleepInterval
    }

    Write-Host "`r[$($progressBar.PadRight($progressMaxWidth, '#'))] Completed!" -ForegroundColor Green

    return $taskResult
}


<#
.SYNOPSIS
Retrieves various information about the local computer.
#>
function Get-ComputerInfo {
    Write-Info "Fetching computer information:"
    Write-InlineProgress -Task {
        try {
            Get-CimInstance -ClassName Win32_ComputerSystem
        } catch {
            Write-Error "Failed to fetch computer information: $_"
        }
    }
}

function Get-CPUInfo {
    Write-Info "Fetching CPU information:"
    Write-InlineProgress -Task {
        try {
            Get-CimInstance -ClassName Win32_Processor
        } catch {
            Write-Error "Failed to fetch CPU information: $_"
        }
    }
}

function Get-RAMInfo {
    Write-Info "Fetching RAM information:"
    Write-InlineProgress -Task {
        try {
            $RAM = Get-CimInstance -ClassName CIM_PhysicalMemory
            $totalRAM = "{0:N2}" -f (($RAM | Measure-Object -Property Capacity -Sum).Sum / 1GB)
            $totalRAM.Trim()
        } catch {
            Write-Error "Failed to fetch RAM information: $_"
        }
    }
}

function Get-MotherboardInfo {
    Write-Info "Fetching Motherboard information:"
    Write-InlineProgress -Task {
        try {
            Get-CimInstance -ClassName Win32_BaseBoard
        } catch {
            Write-Error "Failed to fetch motherboard information: $_"
        }
    }
}

function Get-DriveInfo {
    Write-Info "Fetching Drive Size and Type..."
    Write-InlineProgress -Task {
        try {
            $PhysicalDisks = Get-PhysicalDisk
            $DriveInfo = foreach ($PhysicalDisk in $PhysicalDisks) {
                $DriveType = $PhysicalDisk.MediaType
                $DriveSizeGB = "{0:N2}" -f ($PhysicalDisk.Size / 1GB)
                $Drive = "{0} {1}" -f $DriveSizeGB, $DriveType
                return $Drive.Trim()
            }
            $DriveInfo -join ", "
        } catch {
            Write-Error "Failed to fetch drive information: $_"
        }
    }
}


<#
.SYNOPSIS
Updates or creates an inventory file in either JSON or CSV format.

.DESCRIPTION
The Update-InventoryFile function updates an existing inventory file or creates a new one if it does not exist. 
The inventory file can be in either JSON or CSV format.

.PARAMETER output
The data to be added to the inventory file.

.PARAMETER filePath
The path to the inventory file.

.PARAMETER OutputFormat
The format of the inventory file. It can be either 'JSON' or 'CSV'.

.EXAMPLE
Update-InventoryFile -output $data -filePath "C:\inventory.json" -OutputFormat "JSON"
#>
function Update-InventoryFile {
    Param(
        $output,
        $filePath,
        [ValidateSet('JSON','CSV')]
        $OutputFormat = 'JSON'
    )

    [Environment]::NewLine
    Write-Info "Checking if inventory file exists..."

    $existingContent = New-Object System.Collections.ArrayList

    if (Test-Path -Path $filePath) {
        Write-Success "Inventory file exists. Reading existing content..."

        if ($OutputFormat -eq "JSON") {
            $jsonContent = Get-Content -Path $filePath -Raw | ConvertFrom-Json
            if ($jsonContent -isnot [array]) {
                $existingContent.Add($jsonContent) | Out-Null
            } else {
                foreach($item in $jsonContent){
                    $existingContent.Add($item) | Out-Null
                }
            }
        } else {
            $csvContent = Import-Csv -Path $filePath
            foreach($item in $csvContent){
                $existingContent.Add($item) | Out-Null
            }
        }

        Write-Info "Adding new system information to inventory..."

    } else {
        Write-Info "Inventory file doesn't exist. Creating new file..."
        Write-Info "Inventory file created sucessfully!"
    }
    
    $existingContent.Add($output) | Out-Null

    try {
        if ($OutputFormat -eq "JSON") {
            Write-Info "Converting data to JSON format..."
            $json = $existingContent | ConvertTo-Json
            $json | Out-File -FilePath $filePath
        } else {
            Write-Info "Converting data to CSV format..."
            $existingContent | Export-Csv -Path $filePath -NoTypeInformation
        }
        Write-Success "Inventory file updated successfully."
    } catch {
        Write-Error "Failed to write to inventory file: $_"
    }

    Start-Sleep -Seconds 1
}



<#
.SYNOPSIS
This script collects system information and updates an inventory file in either JSON or CSV format.

.DESCRIPTION
This script collects the following system information: 
- Computer Information
- CPU Information
- RAM Information
- Motherboard Information
- Drive Information

The gathered information is stored in a PSObject. The object's properties are Timestamp, ComputerName, CPU, TotalRAM, Motherboard, and Drive.

The inventory file is updated with this new information. The file's format (JSON or CSV) is determined by the value of the $OutputFormat variable. The filename is "inventory" with the extension as the output format.

Finally, the script writes a success message to the console, stating "Inventory update complete!"

.NOTES
This script depends on several custom functions, including Get-ComputerInfo, Get-CPUInfo, Get-RAMInfo, Get-MotherboardInfo, Get-DriveInfo, Update-InventoryFile, and Write-Success. Ensure these functions are defined and accessible before running this script.
#>

Write-Info "Fetching Data:"

# Get system information
try {
    $Computer = Get-ComputerInfo
    $CPU = Get-CPUInfo
    $TotalRAM = Get-RAMInfo
    $Motherboard = Get-MotherboardInfo
    $Drive = Get-DriveInfo
} catch {
    Write-Error "Failed to fetch system information: $_"
    return
}

Write-Info "System information fetched successfully."

# Define output data as an ordered hashtable
$output = [ordered]@{
    'TimeStamp' = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    'ComputerName' = $Computer.Name
    'CPU' = $CPU.Name
    'TotalRAM' = $TotalRAM
    'Motherboard' = $Motherboard.Product
    'Drive' = $Drive
}

# Convert the hashtable to a PSObject
$output = New-Object PSObject -Property $output

# Validate output format
if ($OutputFormat -notin @('JSON', 'CSV')) {
    Write-Error "Invalid output format: $OutputFormat"
    return
}

# Update the inventory file
$filePath = ".\inventory.$OutputFormat"
try {
    Update-InventoryFile -output $output -filePath $filePath
    Write-Success "Script Execution Successfully"
} catch {
    Write-Error "Failed to update inventory file: $_"
}
