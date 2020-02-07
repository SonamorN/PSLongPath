# Create a global variable to contain the dataTable 
# which contains all the data regarding filepath and length
# This is accessed by various functions
$Global:dataTable = New-Object System.Data.DataTable

# Create Icon Extractor Assembly
$code = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System
{
	public class IconExtractor
	{

	 public static Icon Extract(string file, int number, bool largeIcon)
	 {
	  IntPtr large;
	  IntPtr small;
	  ExtractIconEx(file, number, out large, out small, 1);
	  try
	  {
	   return Icon.FromHandle(largeIcon ? large : small);
	  }
	  catch
	  {
	   return null;
	  }

	 }
	 [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
	 private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

	}
}
"@
Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing

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

# Enable Visual Styles
[Windows.Forms.Application]::EnableVisualStyles()

$mainForm = New-Object System.Windows.Forms.Form
$menuMain = New-Object System.Windows.Forms.MenuStrip
$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem
$menuAbout = New-Object System.Windows.Forms.ToolStripMenuItem
$menuExportCSV = New-Object System.Windows.Forms.ToolStripMenuItem
$menuSaveAs = New-Object System.Windows.Forms.ToolStripMenuItem
$menuAboutInfo = New-Object System.Windows.Forms.ToolStripMenuItem
$menuExit = New-Object System.Windows.Forms.ToolStripMenuItem
$menuAbout = New-Object System.Windows.Forms.ToolStripMenuItem
$mainToolStrip = New-Object System.Windows.Forms.ToolStrip

$Form = New-Object system.Windows.Forms.Form
$Form.ClientSize = '700,540'
$Form.text = "Long Path Checker v1.2"
$Form.TopMost = $false
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $false
$Form.MainMenuSTrip = $menuMain

# Main ToolStrip
[void]$mainForm.Controls.Add($mainToolStrip)
 
# Main Menu Bar
[void]$mainForm.Controls.Add($menuMain)
 
# Menu Options - File
$menuFile.Text = "File"
$menuAbout.Text = "About"
[void]$menuMain.Items.AddRange(@($menuFile,$menuAbout))

# Menu Options - File / Export to CSV
$menuExportCSV.Image = [System.IconExtractor]::Extract("ieframe.dll", 2, $true)
$menuExportCSV.ShortcutKeys = "Control, S"
$menuExportCSV.Text = "Export to CSV"
$menuExportCSV.Add_Click( { Export-DGV2CSV })
[void]$menuFile.DropDownItems.Add($menuExportCSV)
 
# Menu Options - File / Export to HTML
$menuSaveAs.Image = [System.IconExtractor]::Extract("inetcpl.cpl", 25, $true)
$menuSaveAs.ShortcutKeys = "Control, H"
$menuSaveAs.Text = "Export to HTML"
$menuSaveAs.Add_Click( { Export-DGV2HTML })
[void]$menuFile.DropDownItems.Add($menuSaveAs)
 
# Menu Options - File / Exit
$menuExit.Image = [System.IconExtractor]::Extract("shell32.dll", 10, $true)
$menuExit.ShortcutKeys = "Control, X"
$menuExit.Text = "Exit"
$menuExit.Add_Click( { $Form.Close() })
[void]$menuFile.DropDownItems.Add($menuExit)

# Menu Options - About / Info

$menuAboutInfo.Image = [System.IconExtractor]::Extract("imageres.dll", 76, $true)
$menuAboutInfo.ShortcutKeys = "Control, I"
$menuAboutInfo.Text = "Info"
$menuAboutInfo.Add_Click( { $FormAbout.ShowDialog() })
[void]$menuAbout.DropDownItems.Add($menuAboutInfo)


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
$lbDriveSelection.location = New-Object System.Drawing.Point(25, 32)
$lbDriveSelection.Font = 'Microsoft Sans Serif,10'

$lbTotalItems = New-Object system.Windows.Forms.Label
$lbTotalItems.text = ""
$lbTotalItems.AutoSize = $true
$lbTotalItems.width = 25
$lbTotalItems.height = 10
$lbTotalItems.location = New-Object System.Drawing.Point(25, 515)
$lbTotalItems.Font = 'Microsoft Sans Serif,10'

$lbNumPath = New-Object system.Windows.Forms.Label
$lbNumPath.text = "Path Length > than"
$lbNumPath.AutoSize = $true
$lbNumPath.width = 25
$lbNumPath.height = 10
$lbNumPath.location = New-Object System.Drawing.Point(255, 32)
$lbNumPath.Font = 'Microsoft Sans Serif,10'

$lbStopWatch = New-Object system.Windows.Forms.Label
$lbStopWatch.text = ""
$lbStopWatch.AutoSize = $true
$lbStopWatch.width = 25
$lbStopWatch.height = 10
$lbStopWatch.location = New-Object System.Drawing.Point(587, 515)
$lbStopWatch.Font = 'Microsoft Sans Serif,10'

$numPathLength = New-Object System.Windows.Forms.NumericUpDown
$numPathLength.width = 55
$numPathLength.height = 10
$numPathLength.location = New-Object System.Drawing.Point(380, 30)
$numPathLength.Font = 'Microsoft Sans Serif,10'
$numPathLength.Minimum = 200
$numPathLength.Maximum = 2500

$lBoxDrives = New-Object system.Windows.Forms.ComboBox
$lBoxDrives.width = 133
$lBoxDrives.height = 30
$lBoxDrives.DropDownStyle = 'DropDownList'
$lBoxDrives.location = New-Object System.Drawing.Point(115, 30) 

$btScanDrives = New-Object system.Windows.Forms.Button
$btScanDrives.text = "Scan Drive"
$btScanDrives.width = 80
$btScanDrives.height = 30
$btScanDrives.location = New-Object System.Drawing.Point(25, 65)
$btScanDrives.Font = 'Microsoft Sans Serif,10'

$dgvFilePaths = New-Object system.Windows.Forms.DataGridView
$dgvFilePaths.width = 650 
$dgvFilePaths.height = 400 
$dgvFilePaths.location = New-Object System.Drawing.Point(25, 110)
$dgvFilePaths.RowHeadersVisible = $false; 

$Form.controls.AddRange(@($lbDriveSelection, $menuMain, $lBoxDrives, $btScanDrives, $dgvFilePaths, $lbStopWatch, $lbTotalItems, $numPathLength, $lbNumPath))

$FormAbout                       = New-Object system.Windows.Forms.Form
$FormAbout.ClientSize            = '400,400'
$FormAbout.text                  = "About"
$FormAbout.TopMost               = $true
$FormAbout.FormBorderStyle       = 'Fixed3D'
$FormAbout.MaximizeBox           = $false
$FormAbout.Icon = [System.IconExtractor]::Extract("imageres.dll", 76, $true)

$lbCreator                       = New-Object system.Windows.Forms.LinkLabel
$lbCreator.text                  = "Romanos Nianios"
$lbCreator.LinkColor             = "Blue"
$lbCreator.AutoSize              = $true
$lbCreator.width                 = 25
$lbCreator.height                = 10
$lbCreator.location              = New-Object System.Drawing.Point(97,70)
$lbCreator.Font                  = 'Microsoft Sans Serif,10'
$lbCreator.add_click({[system.Diagnostics.Process]::start("https://romanos.nianios.gr")})

$lbYear                          = New-Object system.Windows.Forms.Label
$lbYear.text                     = "2020"
$lbYear.AutoSize                 = $true
$lbYear.width                    = 25
$lbYear.height                   = 10
$lbYear.location                 = New-Object System.Drawing.Point(97,114)
$lbYear.Font                     = 'Microsoft Sans Serif,10'

$lbProductName                   = New-Object system.Windows.Forms.Label
$lbProductName.text              = "PSLongPath v1.2"
$lbProductName.AutoSize          = $true
$lbProductName.width             = 25
$lbProductName.height            = 10
$lbProductName.location          = New-Object System.Drawing.Point(97,91)
$lbProductName.Font              = 'Microsoft Sans Serif,10'

$lbInformation                   = New-Object system.Windows.Forms.Label
$lbInformation.text              = "Script Information"
$lbInformation.AutoSize          = $true
$lbInformation.width             = 25
$lbInformation.height            = 10
$lbInformation.location          = New-Object System.Drawing.Point(85,49)
$lbInformation.Font              = 'Microsoft Sans Serif,10,style=Bold'

$lbThirdParty                    = New-Object system.Windows.Forms.Label
$lbThirdParty.text               = "Third  Party Software/Icon Mentions"
$lbThirdParty.AutoSize           = $true
$lbThirdParty.width              = 25
$lbThirdParty.height             = 10
$lbThirdParty.location           = New-Object System.Drawing.Point(80,158)
$lbThirdParty.Font               = 'Microsoft Sans Serif,10,style=Bold'

$lbPSWriteHTML                   = New-Object system.Windows.Forms.LinkLabel
$lbPSWriteHTML.text              = "PSWriteHTML"
$lbPSWriteHTML.AutoSize          = $true
$lbPSWriteHTML.LinkColor         = "Blue"
$lbPSWriteHTML.width             = 25
$lbPSWriteHTML.height            = 10
$lbPSWriteHTML.location          = New-Object System.Drawing.Point(97,182)
$lbPSWriteHTML.add_click({[system.Diagnostics.Process]::start("https://github.com/EvotecIT/PSWriteHTML")})
$lbPSWriteHTML.Font              = 'Microsoft Sans Serif,10'

$lbPowerForensicsV2              = New-Object system.Windows.Forms.LinkLabel
$lbPowerForensicsV2.text         = "PowerForensicsV2"
$lbPowerForensicsV2.AutoSize     = $true
$lbPowerForensicsV2.LinkColor    = "Blue"
$lbPowerForensicsV2.width        = 25
$lbPowerForensicsV2.height       = 10
$lbPowerForensicsV2.location     = New-Object System.Drawing.Point(97,202)
$lbPowerForensicsV2.add_click({[system.Diagnostics.Process]::start("https://github.com/Invoke-IR/PowerForensics")})
$lbPowerForensicsV2.Font         = 'Microsoft Sans Serif,10'

$lbIcon                          = New-Object system.Windows.Forms.LinkLabel
$lbIcon.text                     = "Winking Document icon by Icons8"
$lbIcon.AutoSize                 = $true
$lbIcon.LinkColor                = "Blue"
$lbIcon.width                    = 25
$lbIcon.height                   = 10
$lbIcon.location                 = New-Object System.Drawing.Point(97,225)
$lbIcon.Font                     = 'Microsoft Sans Serif,10'
$lbIcon.add_click({[system.Diagnostics.Process]::start("https://icons8.com/icons/set/happy-document")})

$lbGitHub                        = New-Object system.Windows.Forms.LinkLabel
$lbGitHub.text                   = "Github"
$lbGitHub.LinkColor              = "Blue"
$lbGitHub.AutoSize               = $true
$lbGitHub.width                  = 25
$lbGitHub.height                 = 10
$lbGitHub.location               = New-Object System.Drawing.Point(97,135)
$lbGitHub.Font                   = 'Microsoft Sans Serif,10'
$lbGitHub.add_click({[system.Diagnostics.Process]::start("https://github.com/rNianios/PSLongPath")})


$FormAbout.controls.AddRange(@($lbCreator,$lbYear,$lbProductName,$lbInformation,$lbThirdParty,$lbPSWriteHTML,$lbPowerForensicsV2,$lbIcon,$lbGitHub))

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

function Get-DGVRowCount {

 $dgvFilePaths.RowCount.ToInt32($null)

   return 
}
function Export-DGV2CSV {
    if ((Get-DGVRowCount) -gt 0) {
        $File = Save-File -initialDirectory "C:\" -fileType "csv" -fileTypeDescription "CSV File" 
        # Build logic to handle empty files
        if ( $File -ne "" ) {
    
        } 
        else {
   
        }

        $Global:dataTable | Export-CSV $File -NoTypeInformation
        Show-MessageBoxExportFinished -path $File
    }
    else {
        [System.Windows.MessageBox]::Show('No data to export, please scan a drive', 'Error', 'OK', 'Error')
    }
}

function Export-DGV2HTML {
    # Call function to show save dialog
 
    if ((Get-DGVRowCount) -gt 0) {
        $File = Save-File -initialDirectory "C:\" -fileType "html" -fileTypeDescription "HTML File" 
        if ( $File -ne "" ) {
           
        } 
        else {
            [System.Windows.MessageBox]::Show('No data to export, please scan a drive', 'Error', 'OK', 'Error')

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
    else {
        [System.Windows.MessageBox]::Show('No data to export, please scan a drive', 'Error', 'OK', 'Error')
    }
    
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