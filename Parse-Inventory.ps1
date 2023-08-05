# Add necessary .NET assembly
Add-Type -AssemblyName System.Windows.Forms

# Define the file path
$filePath = ".\inventory.json"

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
$dataGridView = New-Object System.Windows.Forms.DataGridView -Property @{
    AutoSizeColumnsMode = 'AllCells'
    Dock = 'Fill'
    BackgroundColor = 'White'
    GridColor = 'LightGray'
    ColumnHeadersDefaultCellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle -Property @{
        BackColor = 'Navy'
        ForeColor = 'White'
    }
    DefaultCellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle -Property @{
        BackColor = 'WhiteSmoke'
        ForeColor = 'Black'
    }
    AlternatingRowsDefaultCellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle -Property @{
        BackColor = 'Lavender'
        ForeColor = 'Black'
    }
    RowHeadersVisible = $false
}

# Create a button to clear inventory
$clearButton = New-Object System.Windows.Forms.Button -Property @{
    Text = 'Clear Inventory'
    Dock = 'Fill'
    BackColor = 'Navy'
    ForeColor = 'White'
    Font = New-Object System.Drawing.Font("Arial", 15)  # Set the font and size
}

# Create a button to refresh the data
$refreshButton = New-Object System.Windows.Forms.Button -Property @{
    Text = 'Refresh Data'
    Dock = 'Fill'
    BackColor = 'Navy'
    ForeColor = 'White'
    Font = New-Object System.Drawing.Font("Arial", 15)  # Set the font and size
}

# Create a button to close the application
$closeButton = New-Object System.Windows.Forms.Button -Property @{
    Text = 'Close Application'
    Dock = 'Fill'
    BackColor = 'Navy'
    ForeColor = 'White'
    Font = New-Object System.Drawing.Font("Arial", 15)  # Set the font and size
}


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


# Define what happens when the clear button is clicked
$clearButton.Add_Click({
    # Confirm with the user
    $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to clear the inventory?", "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

    if ($result -eq 'Yes') {
        # Clear the file
        if (Test-Path $filePath) {
            Remove-Item -Path $filePath
        }

        # Refresh the data
        Get-Data
    }
})

# Define what happens when the refresh button is clicked
$refreshButton.Add_Click({
    Get-Data
})

# Define what happens when the close button is clicked
$closeButton.Add_Click({
    $form.Close()
})

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
