Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName "System.Data"

# Destination folder to save the images
$destinationFolder = "C:\Users\Public\Desktop\Extracted_ID_Photos"

# Throttling settings
$batchSize = 10  # Number of images to process in each batch
$pauseDurationSeconds = 1  # Duration to pause between batches (in seconds)

# Create the destination folder if it doesn't exist
if (-not (Test-Path -Path $destinationFolder)) {
    New-Item -ItemType Directory -Path $destinationFolder | Out-Null
}

# Create the window and set properties
$window = New-Object System.Windows.Window
$window.Title = "Database Connection Details"
$window.Width = 400  # Increased width for better layout
$window.Height = 675  # Increased height for more room
$window.WindowStartupLocation = "CenterScreen"
$window.ResizeMode = "CanResizeWithGrip"
$window.MinWidth = 400
$window.MinHeight = 675

# Create a grid for layout
$grid = New-Object System.Windows.Controls.Grid

# Define rows and columns
$rows = 21  # Increased rows for new fields
for ($i = 0; $i -lt $rows; $i++) {
    $row = New-Object -Typ System.Windows.Controls.RowDefinition
    $row.Height = [System.Windows.GridLength]::Auto
    $grid.RowDefinitions.Add($row) | Out-Null  # Suppress output
}

# Define the number of columns
$columns = 2

for ($i = 0; $i -lt $columns; $i++) {
    $column = New-Object System.Windows.Controls.ColumnDefinition
    if ($i -eq 0) {
        $column.Width = New-Object System.Windows.GridLength -ArgumentList 200  # Fixed width for the first column
    } else {
        $column.Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)  # Star width for the second column
    }
    $grid.ColumnDefinitions.Add($column) | Out-Null  # Suppress output
}

# Title
$title = New-Object -TypeName System.Windows.Controls.TextBlock
$title.Text = "Advanced Deblobber Tool"
$title.FontSize = 20
$title.FontWeight = [System.Windows.FontWeights]::Bold
$title.Foreground = [System.Windows.Media.Brushes]::Black
$title.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
$title.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
$title.Margin = [System.Windows.Thickness]::new(10, 10, 0, 0)
$title.SetValue([System.Windows.Controls.Grid]::RowProperty, 0)
$title.SetValue([System.Windows.Controls.Grid]::ColumnSpanProperty, 2)
$grid.Children.Add($title) | Out-Null  # Suppress output

# Description
$description = New-Object -TypeName System.Windows.Controls.TextBlock
$description.Text = "This tool connects to a SQL Server database to extract image files from BLOB fields. Use the provided fields to enter your database connection details and authentication information."
$description.FontSize = 12
$description.Foreground = [System.Windows.Media.Brushes]::Black
$description.TextWrapping = [System.Windows.TextWrapping]::Wrap
$description.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
$description.Margin = [System.Windows.Thickness]::new(10, 5, 0, 10)
$description.SetValue([System.Windows.Controls.Grid]::RowProperty, 1)
$description.SetValue([System.Windows.Controls.Grid]::ColumnSpanProperty, 2)
$grid.Children.Add($description) | Out-Null  # Suppress output

# Helper function to create descriptions above fields
function CreateDescription {
    param (
        [string]$text,
        [int]$row
    )
    
    $description = New-Object -TypeName System.Windows.Controls.TextBlock
    $description.Text = $text
    $description.Foreground = [System.Windows.Media.Brushes]::Red
    $description.TextWrapping = [System.Windows.TextWrapping]::Wrap
    $description.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    $description.Margin = [System.Windows.Thickness]::new(10, 5, 0, 5)
    $description.SetValue([System.Windows.Controls.Grid]::RowProperty, $row)
    $description.SetValue([System.Windows.Controls.Grid]::ColumnSpanProperty, 2)
    return $description
}

# Data Source
$lblDataSourceDesc = CreateDescription "Enter the name of your SQL Server instance." 2
$grid.Children.Add($lblDataSourceDesc) | Out-Null  # Suppress output

$lblDataSource = New-Object -TypeName System.Windows.Controls.Label
$lblDataSource.Content = "Data Source:"
$lblDataSource.SetValue([System.Windows.Controls.Grid]::RowProperty, 3)
$lblDataSource.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 0)
$grid.Children.Add($lblDataSource) | Out-Null  # Suppress output

$txtDataSource = New-Object -TypeName System.Windows.Controls.TextBox
$txtDataSource.SetValue([System.Windows.Controls.Grid]::RowProperty, 3)
$txtDataSource.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 1)
$txtDataSource.Margin = [System.Windows.Thickness]::new(0, 0, 10, 0)  # Right margin for alignment
$grid.Children.Add($txtDataSource) | Out-Null  # Suppress output

# Database Name
$lblDatabaseDesc = CreateDescription "Enter the name of the database you want to connect to." 4
$grid.Children.Add($lblDatabaseDesc) | Out-Null  # Suppress output

$lblDatabase = New-Object -TypeName System.Windows.Controls.Label
$lblDatabase.Content = "Database Name:"
$lblDatabase.SetValue([System.Windows.Controls.Grid]::RowProperty, 5)
$lblDatabase.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 0)
$grid.Children.Add($lblDatabase) | Out-Null  # Suppress output

$txtDatabase = New-Object -TypeName System.Windows.Controls.TextBox
$txtDatabase.SetValue([System.Windows.Controls.Grid]::RowProperty, 5)
$txtDatabase.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 1)
$txtDatabase.Margin = [System.Windows.Thickness]::new(0, 0, 10, 0)  # Right margin for alignment
$grid.Children.Add($txtDatabase) | Out-Null  # Suppress output

# Table Name
$lblTableNameDesc = CreateDescription "Enter the name of the table containing the BLOB data." 6
$grid.Children.Add($lblTableNameDesc) | Out-Null  # Suppress output

$lblTableName = New-Object -TypeName System.Windows.Controls.Label
$lblTableName.Content = "Table Name:"
$lblTableName.SetValue([System.Windows.Controls.Grid]::RowProperty, 7)
$lblTableName.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 0)
$grid.Children.Add($lblTableName) | Out-Null  # Suppress output

$txtTableName = New-Object -TypeName System.Windows.Controls.TextBox
$txtTableName.SetValue([System.Windows.Controls.Grid]::RowProperty, 7)
$txtTableName.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 1)
$txtTableName.Margin = [System.Windows.Thickness]::new(0, 0, 10, 0)  # Right margin for alignment
$grid.Children.Add($txtTableName) | Out-Null  # Suppress output

# Identifier Column
$lblIdentifierColumnDesc = CreateDescription "Enter the name of the column used as the identifier." 8
$grid.Children.Add($lblIdentifierColumnDesc) | Out-Null  # Suppress output

$lblIdentifierColumn = New-Object -TypeName System.Windows.Controls.Label
$lblIdentifierColumn.Content = "Identifier Column:"
$lblIdentifierColumn.SetValue([System.Windows.Controls.Grid]::RowProperty, 9)
$lblIdentifierColumn.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 0)
$grid.Children.Add($lblIdentifierColumn) | Out-Null  # Suppress output

$txtIdentifierColumn = New-Object -TypeName System.Windows.Controls.TextBox
$txtIdentifierColumn.SetValue([System.Windows.Controls.Grid]::RowProperty, 9)
$txtIdentifierColumn.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 1)
$txtIdentifierColumn.Margin = [System.Windows.Thickness]::new(0, 0, 10, 0)  # Right margin for alignment
$grid.Children.Add($txtIdentifierColumn) | Out-Null  # Suppress output

# Photo Column
$lblPhotoColumnDesc = CreateDescription "Enter the name of the column containing the photo BLOB." 10
$grid.Children.Add($lblPhotoColumnDesc) | Out-Null  # Suppress output

$lblPhotoColumn = New-Object -TypeName System.Windows.Controls.Label
$lblPhotoColumn.Content = "Photo Column:"
$lblPhotoColumn.SetValue([System.Windows.Controls.Grid]::RowProperty, 11)
$lblPhotoColumn.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 0)
$grid.Children.Add($lblPhotoColumn) | Out-Null  # Suppress output

$txtPhotoColumn = New-Object -TypeName System.Windows.Controls.TextBox
$txtPhotoColumn.SetValue([System.Windows.Controls.Grid]::RowProperty, 11)
$txtPhotoColumn.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 1)
$txtPhotoColumn.Margin = [System.Windows.Thickness]::new(0, 0, 10, 0)  # Right margin for alignment
$grid.Children.Add($txtPhotoColumn) | Out-Null  # Suppress output

# Authentication Type
$lblAuth = CreateDescription "Select the database authentication type." 12
$grid.Children.Add($lblAuth) | Out-Null  # Suppress output

$rbWindowsAuth = New-Object -TypeName System.Windows.Controls.RadioButton
$rbWindowsAuth.Content = "Windows Authentication"
$rbWindowsAuth.GroupName = "AuthType"
$rbWindowsAuth.IsChecked = $true
$rbWindowsAuth.SetValue([System.Windows.Controls.Grid]::RowProperty, 13)
$rbWindowsAuth.SetValue([System.Windows.Controls.Grid]::ColumnSpanProperty, 2)
$grid.Children.Add($rbWindowsAuth) | Out-Null  # Suppress output

$rbSqlAuth = New-Object -TypeName System.Windows.Controls.RadioButton
$rbSqlAuth.Content = "SQL Server Authentication"
$rbSqlAuth.GroupName = "AuthType"
$rbSqlAuth.SetValue([System.Windows.Controls.Grid]::RowProperty, 14)
$rbSqlAuth.SetValue([System.Windows.Controls.Grid]::ColumnSpanProperty, 2)
$grid.Children.Add($rbSqlAuth) | Out-Null  # Suppress output

# Username
$lblUsernameDesc = CreateDescription "Enter the username for SQL Server Authentication." 15
$lblUsernameDesc.Visibility = [System.Windows.Visibility]::Collapsed  # Hide initially
$grid.Children.Add($lblUsernameDesc) | Out-Null  # Suppress output

$lblUsername = New-Object -TypeName System.Windows.Controls.Label
$lblUsername.Content = "Username:"
$lblUsername.SetValue([System.Windows.Controls.Grid]::RowProperty, 16)
$lblUsername.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 0)
$lblUsername.Visibility = [System.Windows.Visibility]::Collapsed  # Hide initially
$grid.Children.Add($lblUsername) | Out-Null  # Suppress output

$txtUsername = New-Object -TypeName System.Windows.Controls.TextBox
$txtUsername.SetValue([System.Windows.Controls.Grid]::RowProperty, 16)
$txtUsername.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 1)
$txtUsername.Margin = [System.Windows.Thickness]::new(0, 0, 10, 0)  # Right margin for alignment
$txtUsername.Visibility = [System.Windows.Visibility]::Collapsed  # Hide initially
$grid.Children.Add($txtUsername) | Out-Null  # Suppress output

# Password
$lblPasswordDesc = CreateDescription "Enter the password for SQL Server Authentication." 17
$lblPasswordDesc.Visibility = [System.Windows.Visibility]::Collapsed  # Hide initially
$grid.Children.Add($lblPasswordDesc) | Out-Null  # Suppress output

$lblPassword = New-Object -TypeName System.Windows.Controls.Label
$lblPassword.Content = "Password:"
$lblPassword.SetValue([System.Windows.Controls.Grid]::RowProperty, 18)
$lblPassword.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 0)
$lblPassword.Visibility = [System.Windows.Visibility]::Collapsed  # Hide initially
$grid.Children.Add($lblPassword) | Out-Null  # Suppress output

$txtPassword = New-Object -TypeName System.Windows.Controls.PasswordBox
$txtPassword.SetValue([System.Windows.Controls.Grid]::RowProperty, 18)
$txtPassword.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 1)
$txtPassword.Margin = [System.Windows.Thickness]::new(0, 0, 10, 0)  # Right margin for alignment
$txtPassword.Visibility = [System.Windows.Visibility]::Collapsed  # Hide initially
$grid.Children.Add($txtPassword) | Out-Null  # Suppress output

# Event handler for authentication type selection
$rbSqlAuth.Add_Checked({
    $lblUsername.Visibility = [System.Windows.Visibility]::Visible
    $txtUsername.Visibility = [System.Windows.Visibility]::Visible
    $lblPassword.Visibility = [System.Windows.Visibility]::Visible
    $txtPassword.Visibility = [System.Windows.Visibility]::Visible
    $lblUsernameDesc.Visibility = [System.Windows.Visibility]::Visible
    $lblPasswordDesc.Visibility = [System.Windows.Visibility]::Visible
})

$rbWindowsAuth.Add_Checked({
    $lblUsername.Visibility = [System.Windows.Visibility]::Collapsed
    $txtUsername.Visibility = [System.Windows.Visibility]::Collapsed
    $lblPassword.Visibility = [System.Windows.Visibility]::Collapsed
    $txtPassword.Visibility = [System.Windows.Visibility]::Collapsed
    $lblUsernameDesc.Visibility = [System.Windows.Visibility]::Collapsed
    $lblPasswordDesc.Visibility = [System.Windows.Visibility]::Collapsed
})

# Submit Button
$btnSubmit = New-Object -TypeName System.Windows.Controls.Button
$btnSubmit.Content = "Extract Photos"
$btnSubmit.Width = 120
$btnSubmit.Margin = [System.Windows.Thickness]::new(10, 10, 10, 10)
$btnSubmit.SetValue([System.Windows.Controls.Grid]::RowProperty, 19)
$btnSubmit.SetValue([System.Windows.Controls.Grid]::ColumnSpanProperty, 2)
$btnSubmit.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
$grid.Children.Add($btnSubmit) | Out-Null  # Suppress output

# Status TextBlock
$txtStatus = New-Object -TypeName System.Windows.Controls.TextBlock
$txtStatus.Text = ""
$txtStatus.Foreground = [System.Windows.Media.Brushes]::Green
$txtStatus.TextWrapping = [System.Windows.TextWrapping]::Wrap
$txtStatus.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
$txtStatus.Margin = [System.Windows.Thickness]::new(10, 5, 0, 5)
$txtStatus.SetValue([System.Windows.Controls.Grid]::RowProperty, 20)
$txtStatus.SetValue([System.Windows.Controls.Grid]::ColumnSpanProperty, 2)
$grid.Children.Add($txtStatus) | Out-Null  # Suppress output

# Add grid to window
$window.Content = $grid

# Event handler for button click
$btnSubmit.Add_Click({
    $txtStatus.Text = "Extracting photos, please wait..."
    $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [System.Action]{})
    
    $dataSource = $txtDataSource.Text
    $database = $txtDatabase.Text
    $tableName = $txtTableName.Text
    $identifierColumn = $txtIdentifierColumn.Text
    $photoColumn = $txtPhotoColumn.Text
    $authType = if ($rbWindowsAuth.IsChecked) { "Windows" } else { "SQL" }
    $username = $txtUsername.Text
    $password = $txtPassword.Password
    
    # Build the connection string
    if ($authType -eq "Windows") {
        $connectionString = "Server=$dataSource;Database=$database;Integrated Security=True;"
    } else {
        $connectionString = "Server=$dataSource;Database=$database;User Id=$username;Password=$password;"
    }

    try {
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        $connection.Open() | Out-Null  # Suppress output

        $query = "SELECT $identifierColumn, $photoColumn FROM $tableName"
        $command = $connection.CreateCommand()
        $command.CommandText = $query
        $reader = $command.ExecuteReader()

        $count = 0
        $batchCount = 0
        while ($reader.Read()) {
            $identifier = $reader.GetValue(0).ToString()
            $photo = $reader.GetValue(1)
            
            if ($photo -ne [System.DBNull]::Value) {
                $fileName = Join-Path -Path $destinationFolder -ChildPath ("{0}.jpg" -f $identifier)
                [System.IO.File]::WriteAllBytes($fileName, $photo) | Out-Null  # Suppress output
                $count++
            }
            
            $batchCount++
            if ($batchCount -ge $batchSize) {
                Start-Sleep -Seconds $pauseDurationSeconds
                $batchCount = 0
            }
        }

        $reader.Close() | Out-Null  # Suppress output
        $connection.Close() | Out-Null  # Suppress output

        $txtStatus.Text = "Extraction completed. $count photos extracted."
    } catch {
        $txtStatus.Text = "An error occurred: $($_.Exception.Message)"
    }
})

# Show the window
$window.ShowDialog() | Out-Null  # Suppress output
