# Create a global variable to contain the dataTable 
# which contains all the data regarding filepath and length
# This is accessed by various functions
$Global:dataTable = New-Object System.Data.DataTable

$ModulesExist = Test-Path $PSScriptRoot\Modules

if (!($ModulesExist)) {
    Install-Module -Name PSWriteHTML -AllowClobber -Force
    Install-Module -Name PowerForensicsV2 -AllowClobber -Force
}
else {

    $modules = Get-ChildItem -Path $PSScriptRoot\Modules -Recurse -Include "*.psd1"
    foreach ($module in $modules) {
        Import-Module -Name $module.Fullname -Verbose
    }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

$Form = New-Object system.Windows.Forms.Form
$Form.ClientSize = '700,520'
$Form.text = "Long Path Checker"
$Form.TopMost = $false
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $false

$iconBase64 = 'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAABmJLR0QA/wD/AP+gvaeTAAAA2klEQVRYhe2WQQrCMBBFn+IRVBDc6JX0CF6gvZ9SEC9gF55CN+7VhRFKSJtknBIr+TAEksnvY5gMhawf1lMQF2CREkAV4mMoAVaBkALUjXWeAmAGnFGohBQALYhvesCOuu3SyGPoy3HlR31rEmgeojbQTrCxIoBIGSAWwO7upk7AsW+ALj2Ie7Zexc4Bkc/gegBgB6wc+2tzpiZX6UqzdwUKYGmiBG7mrAjwEQNMgQPtM39vcnoDgPfI3QIVcDdRARvc41gdQNVnkK/gvwBC/gdUx6ut5BXISq4XO9py76gKcG0AAAAASUVORK5CYII='
$iconBytes = [Convert]::FromBase64String($iconBase64)
$stream = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
$stream.Write($iconBytes, 0, $iconBytes.Length);
$iconImage = [System.Drawing.Image]::FromStream($stream, $true)
$Form.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())


$lbDriveSelection = New-Object system.Windows.Forms.Label
$lbDriveSelection.text = "Select Drive"
$lbDriveSelection.AutoSize = $true
$lbDriveSelection.width = 25
$lbDriveSelection.height = 10
$lbDriveSelection.location = New-Object System.Drawing.Point(25, 12)
$lbDriveSelection.Font = 'Microsoft Sans Serif,10'

$lbTotalItems = New-Object system.Windows.Forms.Label
$lbTotalItems.text = ""
$lbTotalItems.AutoSize = $true
$lbTotalItems.width = 25
$lbTotalItems.height = 10
$lbTotalItems.location = New-Object System.Drawing.Point(25, 495)
$lbTotalItems.Font = 'Microsoft Sans Serif,10'

$lbNumPath = New-Object system.Windows.Forms.Label
$lbNumPath.text = "Path Length > than"
$lbNumPath.AutoSize = $true
$lbNumPath.width = 25
$lbNumPath.height = 10
$lbNumPath.location = New-Object System.Drawing.Point(248, 12)
$lbNumPath.Font = 'Microsoft Sans Serif,10'

$lbStopWatch = New-Object system.Windows.Forms.Label
$lbStopWatch.text = ""
$lbStopWatch.AutoSize = $true
$lbStopWatch.width = 25
$lbStopWatch.height = 10
$lbStopWatch.location = New-Object System.Drawing.Point(550, 495)
$lbStopWatch.Font = 'Microsoft Sans Serif,10'

$numPathLength = New-Object System.Windows.Forms.NumericUpDown
$numPathLength.width = 55
$numPathLength.height = 10
$numPathLength.location = New-Object System.Drawing.Point(380, 10)
$numPathLength.Font = 'Microsoft Sans Serif,10'
$numPathLength.Minimum = 200
$numPathLength.Maximum = 2500

$lBoxDrives = New-Object system.Windows.Forms.ComboBox
$lBoxDrives.width = 133
$lBoxDrives.height = 30
$lBoxDrives.DropDownStyle = 'DropDownList'
$lBoxDrives.location = New-Object System.Drawing.Point(115, 10) 

$btScanDrives = New-Object system.Windows.Forms.Button
$btScanDrives.text = "Scan Drive"
$btScanDrives.width = 80
$btScanDrives.height = 30
$btScanDrives.location = New-Object System.Drawing.Point(25, 45)
$btScanDrives.Font = 'Microsoft Sans Serif,10'

$btExportCSV = New-Object system.Windows.Forms.Button
$btExportCSV.text = "Export to CSV"
$btExportCSV.width = 100
$btExportCSV.height = 30
$btExportCSV.location = New-Object System.Drawing.Point(150, 45)
$btExportCSV.Font = 'Microsoft Sans Serif,10'

$btExportHTML = New-Object system.Windows.Forms.Button
$btExportHTML.text = "Export to HTML"
$btExportHTML.width = 120
$btExportHTML.height = 30
$btExportHTML.location = New-Object System.Drawing.Point(290, 45)
$btExportHTML.Font = 'Microsoft Sans Serif,10'

$dgvFilePaths = New-Object system.Windows.Forms.DataGridView
$dgvFilePaths.width = 650 
$dgvFilePaths.height = 400 
$dgvFilePaths.location = New-Object System.Drawing.Point(25, 90)
$dgvFilePaths.RowHeadersVisible = $false; 

$Form.controls.AddRange(@($lbDriveSelection, $lBoxDrives, $btScanDrives, $dgvFilePaths, $btExportCSV, $lbStopWatch, $lbTotalItems, $numPathLength, $btExportHTML, $lbNumPath))


$Form.Add_Shown(
    {
        Hide-Console
        $drives = Get-Drives
        foreach ($drive in $drives) {
            if ($drive.VolumeName -eq "" -or $drive.VolumeName -eq $null -or !($drive.VolumeName)) {
                $lBoxDrives.Items.Add($drive.DeviceID)
            }
            else {
                $lBoxDrives.Items.Add("$($drive.DeviceID) - $($drive.VolumeName)")
            }
        }
    
    })

$btScanDrives.Add_Click(
    {
        if ($lBoxDrives.SelectedItem -eq $null) {
            $lbStopWatch.Text = "Please Select A Drive"
            $lbStopWatch.ForeColor = 'Red'
        }
        else {
        
            Add-DataTable2DGV

        }
   
    }
)

$dgvFilePaths.Add_cellclick( { gridClick })

function Add-DataTable2DGV {
    # The string that Get-ForensicFileRecord accepts as Volume Name is in the format of C: or D: without \
    # The following 3 commands are making sure that the string that will be passed to the command has that
    # specific pattern 
    $selectedDrive = $lBoxDrives.SelectedItem.ToString()
    $selectedDrive = $selectedDrive.Trim()
    $selectedDrive = $selectedDrive.Substring(0, 2)

    # start a stopwatch to count how many seconds it takes to finish the task
    $stopwatch = [system.diagnostics.stopwatch]::StartNew()
   
    # call the Get-FileNames function
    $files = Get-Filenames -Drive $selectedDrive

    # Add 2 columns on the DataTable
    [void]$Global:dataTable.Columns.Add("File Path")
    [void]$Global:dataTable.Columns.Add("Length")

    # Hide column headers for faster dataBinding
    $dgvFilePaths.ColumnHeadersVisible = $false
    # Initiate variable only once and convert the string to int32
    $length = $numPathLength.Value.ToInt32($null)
    foreach ($file in $files) {                
        if ($file.FullName.Length -ge $length) {               
            # Fill DataTable with the path and length values
            $Global:dataTable.Rows.Add($file.FullName, $file.FullName.Length)
        }
    } 
    # Bind the dataTable to the DataGrivView as DataSource
    $dgvFilePaths.DataSource = $global:dataTable
    # Make Column Headers visible
    $dgvFilePaths.ColumnHeadersVisible = $true

    # Sort the dataGridView by path length
    $dgvFilePaths.Sort($dgvFilePaths.Columns[1], 'Descending')

    # Set column width
    $dgvFilePaths.Columns[0].Width = 574
    $dgvFilePaths.Columns[1].Width = 55
    
    # Set dataGridView to ReadOnly
    $dgvFilePaths.ReadOnly = $true

    # Show the files that exceed the path limit and all the paths that were returned by Get-ForensicFileRecord
    $lbTotalItems.Text = "Files: $($dgvFilePaths.RowCount.ToString()) / $($files.Count)" 
    # Empty the variable
    $files = $null
    # Stop the clock
    $lbStopWatch.ForeColor = 'Black'
    $lbStopWatch.Text = "Finished: $($stopwatch.Elapsed.ToString('mm\:ss'))"
}

function gridClick {
    # This is called when the user clicks on the DataGridView
    # It gets the row and column index
    # and if it's a specific column index, it launches explorer 
    # on the value of the row on path column
    $rowIndex = $dgvFilePaths.CurrentRow.Index
    $columnIndex = $dgvFilePaths.CurrentCell.ColumnIndex
    if ($columnIndex -eq 0) {
        Write-Host $location
        $location = Split-Path $dgvFilePaths.Rows[$rowIndex].Cells[0].value -Parent
        explorer $location
        $columnIndex = 0
        $columnIndex = $null
    }
}

$btExportCSV.Add_Click(
    {
        Export-DGV2CSV
    }
)

$btExportHTML.Add_Click(
    {

        Export-DG2HTML
    }
) 

function Export-DG2HTML {
    # Call function to show save dialog
    $File = Save-File -initialDirectory "C:\" -fileType "html" -fileTypeDescription "HTML File" 
    if ( $File -ne "" ) {
           
    } 
    else {
            
    }
 
    # Need to transform the data to only show File Path and Length instead of includign TypeInformation as well.
    $dataTable = $global:dataTable | Select-Object "File Path", Length
    $PagingOptions = @(50, 100, 250, 500, 1000)
    New-HTML -TitleText 'Long Path File Names' -UseCssLinks:$true -UseJavaScriptLinks:$true  -FilePath $file {
        New-HTMLContent -HeaderText 'Long Path File Names' {
            New-HTMLPanel {
                New-HTMLTable -DataTable $dataTable -PagingOptions $PagingOptions -HideFooter -PagingStyle "full_numbers" {
                    New-HTMLTableHeader -Title 'Long Path Report' 
                }  
            }
        }
    }   

    # Show Messagebox asking if the user would like to open the file immediatelly.
    Show-MessageBoxExportFinished -path $File
}

function Show-Console {
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #ref https://stackoverflow.com/questions/40617800/opening-powershell-script-and-hide-command-prompt-but-not-the-gui
    # Hide = 0,
    # ShowNormal = 1,
    # ShowMinimized = 2,
    # ShowMaximized = 3,
    # Maximize = 3,
    # ShowNormalNoActivate = 4,
    # Show = 5,
    # Minimize = 6,
    # ShowMinNoActivate = 7,
    # ShowNoActivate = 8,
    # Restore = 9,
    # ShowDefault = 10,
    # ForceMinimized = 11

    [Console.Window]::ShowWindow($consolePtr, 4)
}

function Hide-Console {
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}

# Get the attached drives on the computer
function Get-Drives {
    # Return only Removable and Local Drives, excluding network drives as can't read MFT Table.
    $drives = Get-WmiObject -Class Win32_logicaldisk | Where-Object { $_.DriveType -eq 3 -or $_.DriveType -eq 4 } | Select-Object DeviceID, VolumeName
    return $drives
}

function Show-MessageBoxExportFinished {
    param(
        [string]$path
    )
    Add-Type -AssemblyName PresentationFramework
    $msgBoxInput = [System.Windows.MessageBox]::Show('The export has finished. Would you like to open the exported file?', 'Export Finished', 'YesNo', 'Info')

    switch ($msgBoxInput) {
        'Yes' {
            Invoke-Item $path
        }
        'No' {
            ## Do something
        }
    }
}
function Save-File {
    
    param([string] $initialDirectory,
        [string]$fileType,
        [string]$fileTypeDescription)

    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "$filetypeDescription (*.$filetype) | *.$filetype"
    $OpenFileDialog.ShowDialog() | Out-Null
    
    return $OpenFileDialog.filename
} 
 

function Export-DGV2CSV {
    $File = Save-File -initialDirectory "C:\" -fileType "csv" -fileTypeDescription "CSV File" 
    # Build logic to handle empty files
    if ( $File -ne "" ) {
    
    } 
    else {
   
    }

    $Global:dataTable | Export-CSV $File -NoTypeInformation
    Show-MessageBoxExportFinished -path $File
}

function Get-Filenames {
    param([string]$Drive)

    $MaxThreads = 5

    $ScriptBlock = {
        Param (
            [string]$Drive
        )
        $RunResult = Get-ForensicFileRecord -VolumeName $Drive | Select-Object FullName
        Return $RunResult
    }
        
    $RunspacePool = [RunspaceFactory ]::CreateRunspacePool(1, $MaxThreads)
    # More info here https://docs.microsoft.com/en-us/dotnet/api/system.threading.apartmentstate?redirectedfrom=MSDN&view=netframework-4.8
    $RunspacePool.ApartmentState = "STA"
    # More info here https://docs.microsoft.com/en-us/dotnet/api/system.management.automation.runspaces.psthreadoptions?view=pscore-6.2.0
    $RunspacePool.ThreadOptions - "ReuseThread"
    $RunspacePool.Open()

    $Job = [powershell]::Create().AddScript($ScriptBlock).AddArgument($Drive)
    $Job.RunspacePool = $RunspacePool
    $Jobs += New-Object PSObject -Property @{
        Pipe   = $Job
        Result = $Job.BeginInvoke()
    }

    $lbStopWatch.ForeColor = 'Black'

    Do {
        # Update clock
        $lbStopWatch.Text = "Running: $($stopwatch.Elapsed.ToString('mm\:ss'))"
        # Force Form to Update
        [System.Windows.Forms.Application]::DoEvents() 
    } While ( $Jobs.Result.IsCompleted -contains $false)
    Write-Host "All jobs completed!"
    
    $Results = @()
    ForEach ($Job in $Jobs) {
        $Results += $Job.Pipe.EndInvoke($Job.Result)
    }
    
    return $results
 
}   

# Show Form
$Form.ShowDialog()