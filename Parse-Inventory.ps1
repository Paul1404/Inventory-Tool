<#
.SYNOPSIS
A function to create a new button.

.DESCRIPTION
This function creates a new System.Windows.Forms.Button object with the provided properties and returns it.

.PARAMETER text
The text to display on the button.

.PARAMETER onClick
The scriptblock to execute when the button is clicked.

.EXAMPLE
$button = New-Button "Click Me!" { Write-Host "Button Clicked!" }
#>
function New-Button {
    param([string]$text, [scriptblock]$onClick)

    return New-Object System.Windows.Forms.Button -Property @{
        Text = $text
        Dock = $settings.Button.Dock
        BackColor = $settings.Button.BackColor
        ForeColor = $settings.Button.ForeColor
        Font = New-Object System.Drawing.Font($settings.Button.Font, $settings.Button.FontSize)  # Set the font and size
        Add_Click = $onClick
    }
}

<#
.SYNOPSIS
A function to convert a PSCustomObject to a Hashtable.

.DESCRIPTION
This function iterates over the properties of a PSCustomObject and adds them to a new Hashtable. The Hashtable is then returned.

.PARAMETER InputObject
The PSCustomObject to convert to a Hashtable.

.EXAMPLE
$hash = ConvertTo-Hashtable $customObject
#>
function ConvertTo-Hashtable {
    param($InputObject)

    $hashTable = @{}

    $InputObject.PSObject.Properties | ForEach-Object {
        $hashTable[$_.Name] = $_.Value
    }

    return $hashTable
}

<#
.SYNOPSIS
A function to load and display data from a JSON file in a DataGridView.

.DESCRIPTION
This function clears the existing data from the DataTable, loads new data from the specified JSON file, adds the data to the DataTable, and sets the DataSource of the DataGridView to the DataTable.

.EXAMPLE
Get-Data
#>
function Get-Data {
    # Clear existing data
    $dataTable.Clear()

    if (Test-Path $filePath) {
        # Load and parse JSON data
        $jsonData = Get-Content -Path $filePath | ConvertFrom-Json

        # Add JSON data to the DataTable
        foreach ($item in $jsonData) {
            $row = $dataTable.NewRow()
            $row['TimeStamp'] = $item.TimeStamp
            $row['ComputerName'] = $item.ComputerName
            $row['CPU'] = $item.CPU
            $row['TotalRAM'] = $item.TotalRAM
            $row['Motherboard'] = $item.Motherboard
            $row['Drive'] = $item.Drive.Trim('"')
            $row['BasicWindowsVersion'] = $item.BasicWindowsVersion
            $row['WindowsUpdateVersion'] = $item.WindowsUpdateVersion
            $row['IPAddress'] = $item.IPAddress
            $row['MACAddress'] = $item.MACAddress
            $dataTable.Rows.Add($row)
        }
    }

    $dataGridView.DataSource = $dataTable
}



# Add necessary .NET assembly
Add-Type -AssemblyName System.Windows.Forms

# Define the file path
$filePath = ".\inventory.json"

# Load settings from JSON file
$settings = Get-Content -Path '.\settings.json' | ConvertFrom-Json

# Create a DataTable
$dataTable = New-Object System.Data.DataTable

# Define columns
$dataTable.Columns.Add("Timestamp")
$dataTable.Columns.Add('ComputerName')
$dataTable.Columns.Add('CPU')
$dataTable.Columns.Add('TotalRAM')
$dataTable.Columns.Add('Motherboard')
$dataTable.Columns.Add('Drive')
$dataTable.Columns.Add('BasicWindowsVersion')
$dataTable.Columns.Add('WindowsUpdateVersion')
$dataTable.Columns.Add('IPAddress')
$dataTable.Columns.Add('MACAddress')


# Create a DataGridView and add data
$dataGridView = New-Object System.Windows.Forms.DataGridView -Property (ConvertTo-Hashtable $settings.DataGridView)
# Set AutoSizeColumnsMode to Fill for the DataGridView
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill

# Create buttons
$clearButton = New-Button 'Clear Inventory' {
    $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to clear the inventory?", "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

    if ($result -eq 'Yes') {
        if (Test-Path $filePath) {
            Remove-Item -Path $filePath
        }

        Get-Data
    }
} (ConvertTo-Hashtable $settings.Button)

$refreshButton = New-Button 'Refresh Data' {
    Get-Data
} (ConvertTo-Hashtable $settings.Button)

$closeButton = New-Button 'Close Application' {
    $form.Close()
} (ConvertTo-Hashtable $settings.Button)

# Create a panel and add buttons
$buttonPanel = New-Object System.Windows.Forms.TableLayoutPanel
$buttonPanel.RowCount = 1
$buttonPanel.ColumnCount = 3
$buttonPanel.Dock = 'Bottom'
$buttonPanel.Controls.Add($clearButton, 0, 0)
$buttonPanel.Controls.Add($refreshButton, 1, 0)
$buttonPanel.Controls.Add($closeButton, 2, 0)
$buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
$buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
$buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))

# Create a form and add DataGridView and button panel
$form = New-Object System.Windows.Forms.Form -Property @{
    Size = New-Object System.Drawing.Size(1500, 1000)
    Text = "Inventory-Tool"  # Set the title of the form
}

# Set the form icon
$iconPath = Join-Path -Path $PSScriptRoot -ChildPath 'icon.ico'
if (Test-Path $iconPath) {
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
} else {
    Write-Warning "Icon file not found at path: $iconPath"
}

$form.Controls.Add($dataGridView)
$form.Controls.Add($buttonPanel)

# Load and display data
Get-Data

# Show the form
$form.ShowDialog()