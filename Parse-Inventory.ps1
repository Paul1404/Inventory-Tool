# Add necessary .NET assembly
Add-Type -AssemblyName System.Windows.Forms

# Define the file path
$filePath = ".\inventory.json"

function ConvertTo-Hashtable {
    param($InputObject)

    $hashTable = @{}

    $InputObject.PSObject.Properties | ForEach-Object {
        $hashTable[$_.Name] = $_.Value
    }

    return $hashTable
}


# Load settings from JSON file
$settings = Get-Content -Path '.\settings.json' | ConvertFrom-Json

# Create a DataTable
$dataTable = New-Object System.Data.DataTable

# Define columns
$dataTable.Columns.Add("TimeStamp")
$dataTable.Columns.Add('ComputerName')
$dataTable.Columns.Add('CPU')
$dataTable.Columns.Add('TotalRAM')
$dataTable.Columns.Add('Motherboard')
$dataTable.Columns.Add('Drive')

# Function to load and display data
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
            $row['Drive'] = $item.Drive.Trim('"')  # Trim the double quotes from the Drive property
            $dataTable.Rows.Add($row)
        }
    }

    $dataGridView.DataSource = $dataTable
}

# Create a DataGridView and add data
$dataGridView = New-Object System.Windows.Forms.DataGridView -Property (ConvertTo-Hashtable $settings.DataGridView)

# Function to create a button
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
$buttonPanel.ColumnCount = 3  # Increase column count to 3
$buttonPanel.Dock = 'Bottom'
$buttonPanel.Controls.Add($clearButton, 0, 0)
$buttonPanel.Controls.Add($refreshButton, 1, 0)
$buttonPanel.Controls.Add($closeButton, 2, 0)  # Add close button in the third column
$buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))  # Set width to 33.33%
$buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))  # Set width to 33.33%
$buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))  # Add a new ColumnStyle for the third column

# Create a form and add DataGridView and button panel
$form = New-Object System.Windows.Forms.Form -Property @{
    Size = New-Object System.Drawing.Size(1000, 1000)
}
$form.Controls.Add($dataGridView)
$form.Controls.Add($buttonPanel)

# Load and display data
Get-Data

# Show the form
$form.ShowDialog()