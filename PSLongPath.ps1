# Create a global variable to contain the dataTable 
# which contains all the data regarding filepath and length
# This is accessed by various functions
$Global:dataTable = $null

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
Add-Type -AssemblyName PresentationFramework #messagebox
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
$menuPreferences = New-Object System.Windows.Forms.ToolStripMenuItem
$menuAbout = New-Object System.Windows.Forms.ToolStripMenuItem
$menuOpenDisk = New-Object System.Windows.Forms.ToolStripMenuItem
$menuOpenFolder = New-Object System.Windows.Forms.ToolStripMenuItem
$menuExportCSV = New-Object System.Windows.Forms.ToolStripMenuItem
$menuExportHTML = New-Object System.Windows.Forms.ToolStripMenuItem
$menuOptions = New-Object System.Windows.Forms.ToolStripMenuItem
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
$menuPreferences.Text = "Preferences"
[void]$menuMain.Items.AddRange(@($menuFile, $menuPreferences, $menuAbout))

$menuOpenDisk.Image = [System.IconExtractor]::Extract("shell32.dll", 4, $true)
$menuOpenDisk.ShortcutKeys = "Control, O"
$menuOpenDisk.Text = "Scan Drive"
#$menuOpen.Add_Click({OpenFile})
[void]$menuFile.DropDownItems.Add($menuOpenDisk)

$menuOpenFolder.Image = [System.IconExtractor]::Extract("shell32.dll", 4, $true)
$menuOpenFolder.ShortcutKeys = "Control, F"
$menuOpenFolder.Text = "Scan Folder"
$menuOpenFolder.Add_Click( {
        $folderToScan = Get-ScanFolder
        Add-DataTable2DGV $folderToScan.SubString(0, 2) $folderToScan
    })
[void]$menuFile.DropDownItems.Add($menuOpenFolder)

# Menu Options - File / Export to CSV
$menuExportCSV.Image = [System.IconExtractor]::Extract("ieframe.dll", 2, $true)
$menuExportCSV.ShortcutKeys = "Control, S"
$menuExportCSV.Text = "Export to CSV"
$menuExportCSV.Add_Click( { Export-DGV2CSV })
[void]$menuFile.DropDownItems.Add($menuExportCSV)
 
# Menu Options - File / Export to HTML
$menuExportHTML.Image = [System.IconExtractor]::Extract("inetcpl.cpl", 25, $true)
$menuExportHTML.ShortcutKeys = "Control, H"
$menuExportHTML.Text = "Export to HTML"
$menuExportHTML.Add_Click( { Export-DGV2HTML })
[void]$menuFile.DropDownItems.Add($menuExportHTML)
 
# Menu Options - File / Exit
$menuExit.Image = [System.IconExtractor]::Extract("shell32.dll", 10, $true)
$menuExit.ShortcutKeys = "Control, X"
$menuExit.Text = "Exit"
$menuExit.Add_Click( { $Form.Close() })
[void]$menuFile.DropDownItems.Add($menuExit)

# Menu Option - Preferences / Options
$menuOptions.Image = [System.IconExtractor]::Extract("shell32.dll", 207, $true)
$menuOptions.ShortcutKeys = "Control, P"
$menuOptions.Text = "Options"
$menuOptions.Add_Click( { $FormOptions.ShowDialog() })
[void]$menuPreferences.DropDownItems.Add($menuOptions)


<# Menu Option - About / Info #>
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

$lbTotalItems = New-Object system.Windows.Forms.Label
$lbTotalItems.text = ""
$lbTotalItems.AutoSize = $true
$lbTotalItems.width = 25
$lbTotalItems.height = 10
$lbTotalItems.location = New-Object System.Drawing.Point(150, 515)
$lbTotalItems.Font = 'Microsoft Sans Serif,10'

$lbPathLength = New-Object system.Windows.Forms.Label
$lbPathLength.text = ""
$lbPathLength.AutoSize = $true
$lbPathLength.width = 25
$lbPathLength.height = 10
$lbPathLength.location = New-Object System.Drawing.Point(25, 515)
$lbPathLength.Font = 'Microsoft Sans Serif,10'

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

$dgvFilePaths = New-Object system.Windows.Forms.DataGridView
$dgvFilePaths.width = 670 
$dgvFilePaths.height = 475 
$dgvFilePaths.location = New-Object System.Drawing.Point(15, 35)
$dgvFilePaths.RowHeadersVisible = $false; 

$Form.controls.AddRange(@($menuMain, $dgvFilePaths, $lbStopWatch, $lbTotalItems, $lbPathLength))

$FormOptions = New-Object system.Windows.Forms.Form
$FormOptions.ClientSize = '400,400'
$FormOptions.text = "Options"
$FormOptions.TopMost = $false

$lbOptionsPathLength = New-Object system.Windows.Forms.Label
$lbOptionsPathLength.text = "Path Length"
$lbOptionsPathLength.AutoSize = $true
$lbOptionsPathLength.width = 25
$lbOptionsPathLength.height = 10
$lbOptionsPathLength.location = New-Object System.Drawing.Point(38, 60)
$lbOptionsPathLength.Font = 'Microsoft Sans Serif,10'

$tltOptionsPathLength = New-Object system.Windows.Forms.ToolTip

$tltOptionsPathLength.SetToolTip($lbOptionsPathLength, 'This Options allows you to set the minimum path length which you want to appear on your report.')

$numPathLength = New-Object System.Windows.Forms.NumericUpDown
$numPathLength.width = 55
$numPathLength.height = 10
$numPathLength.location = New-Object System.Drawing.Point(150, 58)
$numPathLength.Font = 'Microsoft Sans Serif,10'
$numPathLength.Minimum = 200
$numPathLength.Maximum = 2500

$btOptionsSave = New-Object system.Windows.Forms.Button
$btOptionsSave.text = "Save"
$btOptionsSave.width = 60
$btOptionsSave.height = 30
$btOptionsSave.location = New-Object System.Drawing.Point(159, 346)
$btOptionsSave.Font = 'Microsoft Sans Serif,10'

$FormOptions.controls.AddRange(@($lbOptionsPathLength, $numPathLength, $btOptionsSave))

$FormAbout = New-Object system.Windows.Forms.Form
$FormAbout.ClientSize = '400,400'
$FormAbout.text = "About"
$FormAbout.TopMost = $true
$FormAbout.FormBorderStyle = 'Fixed3D'
$FormAbout.MaximizeBox = $false
$FormAbout.Icon = [System.IconExtractor]::Extract("imageres.dll", 76, $true)

$lbCreator = New-Object system.Windows.Forms.LinkLabel
$lbCreator.text = "Romanos Nianios"
$lbCreator.LinkColor = "Blue"
$lbCreator.AutoSize = $true
$lbCreator.width = 25
$lbCreator.height = 10
$lbCreator.location = New-Object System.Drawing.Point(97, 70)
$lbCreator.Font = 'Microsoft Sans Serif,10'
$lbCreator.add_click( { [system.Diagnostics.Process]::start("https://romanos.nianios.gr") })

$lbYear = New-Object system.Windows.Forms.Label
$lbYear.text = "2020"
$lbYear.AutoSize = $true
$lbYear.width = 25
$lbYear.height = 10
$lbYear.location = New-Object System.Drawing.Point(97, 114)
$lbYear.Font = 'Microsoft Sans Serif,10'

$lbProductName = New-Object system.Windows.Forms.Label
$lbProductName.text = "PSLongPath v1.2"
$lbProductName.AutoSize = $true
$lbProductName.width = 25
$lbProductName.height = 10
$lbProductName.location = New-Object System.Drawing.Point(97, 91)
$lbProductName.Font = 'Microsoft Sans Serif,10'

$lbInformation = New-Object system.Windows.Forms.Label
$lbInformation.text = "Script Information"
$lbInformation.AutoSize = $true
$lbInformation.width = 25
$lbInformation.height = 10
$lbInformation.location = New-Object System.Drawing.Point(85, 49)
$lbInformation.Font = 'Microsoft Sans Serif,10,style=Bold'

$lbThirdParty = New-Object system.Windows.Forms.Label
$lbThirdParty.text = "Third  Party Software/Icon Mentions"
$lbThirdParty.AutoSize = $true
$lbThirdParty.width = 25
$lbThirdParty.height = 10
$lbThirdParty.location = New-Object System.Drawing.Point(80, 158)
$lbThirdParty.Font = 'Microsoft Sans Serif,10,style=Bold'

$lbPSWriteHTML = New-Object system.Windows.Forms.LinkLabel
$lbPSWriteHTML.text = "PSWriteHTML"
$lbPSWriteHTML.AutoSize = $true
$lbPSWriteHTML.LinkColor = "Blue"
$lbPSWriteHTML.width = 25
$lbPSWriteHTML.height = 10
$lbPSWriteHTML.location = New-Object System.Drawing.Point(97, 182)
$lbPSWriteHTML.add_click( { [system.Diagnostics.Process]::start("https://github.com/EvotecIT/PSWriteHTML") })
$lbPSWriteHTML.Font = 'Microsoft Sans Serif,10'

$lbPSWriteHTMLLicense = New-Object system.Windows.Forms.LinkLabel
$lbPSWriteHTMLLicense.text = "License"
$lbPSWriteHTMLLicense.AutoSize = $true
$lbPSWriteHTMLLicense.LinkColor = "Blue"
$lbPSWriteHTMLLicense.width = 25
$lbPSWriteHTMLLicense.height = 10
$lbPSWriteHTMLLicense.location = New-Object System.Drawing.Point(220, 182)
$lbPSWriteHTMLLicense.add_click( { [system.Diagnostics.Process]::start("https://github.com/EvotecIT/PSWriteHTML/blob/master/LICENSE") })
$lbPSWriteHTMLLicense.Font = 'Microsoft Sans Serif,10'

$lbPowerForensicsV2 = New-Object system.Windows.Forms.LinkLabel
$lbPowerForensicsV2.text = "PowerForensicsV2"
$lbPowerForensicsV2.AutoSize = $true
$lbPowerForensicsV2.LinkColor = "Blue"
$lbPowerForensicsV2.width = 25
$lbPowerForensicsV2.height = 10
$lbPowerForensicsV2.location = New-Object System.Drawing.Point(97, 202)
$lbPowerForensicsV2.add_click( { [system.Diagnostics.Process]::start("https://github.com/Invoke-IR/PowerForensics") })
$lbPowerForensicsV2.Font = 'Microsoft Sans Serif,10'

$lbPowerForensicsV2License = New-Object system.Windows.Forms.LinkLabel
$lbPowerForensicsV2License.text = "License"
$lbPowerForensicsV2License.AutoSize = $true
$lbPowerForensicsV2License.LinkColor = "Blue"
$lbPowerForensicsV2License.width = 25
$lbPowerForensicsV2License.height = 10
$lbPowerForensicsV2License.location = New-Object System.Drawing.Point(220, 202)
$lbPowerForensicsV2License.add_click( { [system.Diagnostics.Process]::start("https://github.com/Invoke-IR/PowerForensics/blob/master/LICENSE.md") })
$lbPowerForensicsV2License.Font = 'Microsoft Sans Serif,10'

$lbIcon = New-Object system.Windows.Forms.LinkLabel
$lbIcon.text = "Winking Document icon by Icons8"
$lbIcon.AutoSize = $true
$lbIcon.LinkColor = "Blue"
$lbIcon.width = 25
$lbIcon.height = 10
$lbIcon.location = New-Object System.Drawing.Point(97, 225)
$lbIcon.Font = 'Microsoft Sans Serif,10'
$lbIcon.add_click( { [system.Diagnostics.Process]::start("https://icons8.com/icons/set/happy-document") })

$lbGitHub = New-Object system.Windows.Forms.LinkLabel
$lbGitHub.text = "Github"
$lbGitHub.LinkColor = "Blue"
$lbGitHub.AutoSize = $true
$lbGitHub.width = 25
$lbGitHub.height = 10
$lbGitHub.location = New-Object System.Drawing.Point(97, 135)
$lbGitHub.Font = 'Microsoft Sans Serif,10'
$lbGitHub.add_click( { [system.Diagnostics.Process]::start("https://github.com/rNianios/PSLongPath") })

$FormAbout.controls.AddRange(@($lbCreator, $lbYear, $lbProductName, $lbInformation, $lbThirdParty, $lbPSWriteHTML,$lbPSWriteHTMLLicense, $lbPowerForensicsV2,$lbPowerForensicsV2License, $lbIcon, $lbGitHub))

$Form.Add_Shown(
    {
        Hide-Console
        $pathLengthvalue = Get-IniFileSettings -SettingsCategory "Settings" -SettingValue "pathlength"
        if ($pathLengthvalue -eq $false) {
            $lbPathLength.Text = "Path Length Option N/A"
        }
        else {
            $lbPathLength.Text = "Path Length > $pathLengthvalue"
        }
        $drives = Get-Drives 
        foreach ($d in $drives) {
            if ($d.VolumeName -eq "" -or $d.VolumeName -eq $null -or !($d.VolumeName)) {
                Add-DriveToMenu -driveL $d.DeviceID
            }
            else {
                Add-DriveToMenu -driveL $d.DeviceID -VolumeName $d.VolumeName
            }
        } 
    
    })

$FormOptions.Add_Shown( {

        $iniExists = Test-IniFile 
        if ($iniExists -eq $false) {

            $numPathLength.Text = 200
      
        }
        else {
            $settingsIni = Get-IniFile -FilePath $PSScriptRoot\settings.ini
            $numPathLength.Text = $settingsIni."Settings".pathlength
        }
    })

$btOptionsSave.Add_Click(
    {
        $iniExists = Test-IniFile 
        if ($iniExists -eq $false) {
            Set-DefaultSettings
        }
        else {
            Set-Settings
            $lbPathLength.Text = "Path Length >  $(Get-IniFileSettings -SettingsCategory "Settings" -SettingValue "pathlength")"
        }
        $msgBoxInput = [System.Windows.MessageBox]::Show('Settings Saved', 'Settings Saved', 'OK', 'Info')
        $FormOptions.Close()
    }
)

$dgvFilePaths.Add_cellclick( { gridClick })

function Get-IniFileSettings {
    param(
        [string]$SettingsCategory,
        [string]$SettingValue
    )
    $iniExists = Test-IniFile 
    if ($iniExists -eq $false) {

        return $false
      
    }
    else {
        $settingsIni = Get-IniFile -FilePath $PSScriptRoot\settings.ini
        $result = $settingsIni."$($SettingsCategory)".$SettingValue
        return $result
    }
}
function Set-DefaultSettings {
    $settings = [ordered] @{ }
    $settings.Add("Settings", (New-Object System.Collections.Specialized.OrderedDictionary))
    $settings."Settings".Add("pathlength", "200")
    $settings | Out-IniFile -FilePath $PSScriptRoot\settings.ini
}

function Set-Settings {
    $settings = [ordered] @{ }
    $settings.Add("Settings", (New-Object System.Collections.Specialized.OrderedDictionary))
    $settings."Settings".Add("pathlength", $numPathLength.Text)
    $settings | Out-IniFile -FilePath $PSScriptRoot\settings.ini
}

function Test-IniFile {
    return Test-Path("$PSScriptRoot\settings.ini")
}
# Handle iniFile https://www.remkoweijnen.nl/blog/2014/07/29/handling-ini-files-powershell/ 
function Get-IniFile {
    param (
        [parameter(mandatory = $true, position = 0, valuefrompipelinebypropertyname = $true, valuefrompipeline = $true)][string]$FilePath
    )

    $ini = New-Object System.Collections.Specialized.OrderedDictionary
    $currentSection = New-Object System.Collections.Specialized.OrderedDictionary
    $curSectionName = "default"

    switch -regex (gc $FilePath) {
        "^\[(?<section>.*)\]" {
            $ini.Add($curSectionName, $currentSection)
			
            $curSectionName = $Matches['Section']
            $currentSection = New-Object System.Collections.Specialized.OrderedDictionary	
        }
        "(?<key>\w+)\=(?<value>\w+)" {
            # add to current section Hash Set
            $currentSection.Add($Matches['Key'], $Matches['Value'])
        }
        "^$" {
            # ignore blank line
        }
		 
        "(?<key>\;)(?<value>.*)" {
            $currentSection.Add($Matches['Key'], $Matches['Value'])	  
        }
        default {
            throw "Unidentified: $_"  # should not happen
        }
    }
    if ($ini.Keys -notcontains $curSectionName) { $ini.Add($curSectionName, $currentSection) }
	
    return $ini
}

function Out-IniFile {
    param (
        [parameter(mandatory = $true, position = 0, valuefrompipelinebypropertyname = $true, valuefrompipeline = $true)][System.Collections.Specialized.OrderedDictionary]$ini,
        [parameter(mandatory = $false, position = 1, valuefrompipelinebypropertyname = $true, valuefrompipeline = $false)][String]$FilePath
    )

    $output = ""
    ForEach ($section in $ini.GetEnumerator()) {
        if ($section.Name -ne "default") { 
            # insert a blank line after a section
            $sep = @{$true = ""; $false = "`r`n" }[[String]::IsNullOrWhiteSpace($output)]
            $output += "$sep[$($section.Name)]`r`n" 
        }
        ForEach ($entry in $section.Value.GetEnumerator()) {
            $sep = @{$true = ""; $false = "=" }[$entry.Name -eq ";"]
            $output += "$($entry.Name)$sep$($entry.Value)`r`n"
        }
    }

    $output = $output.TrimEnd("`r`n")
    if ([String]::IsNullOrEmpty($FilePath)) {
        return $output
    }
    else {
        $output | Out-File -FilePath $FilePath -Encoding:ASCII
    }
}

function Get-ScanFolder {
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.Description = "Select Folder to Scan"
    $FolderBrowser.ShowNewFolderButton = $false
    $FolderBrowser.ShowDialog() | Out-Null
    return $FolderBrowser.SelectedPath
}

function Add-DriveToMenu {
    param(
        [string]$driveL,
        [string]$VolumeName = $null
    )
    if ($VolumeName) {
        $menuText = "$driveL - $volumeName"
    }
    else {
        $menuText = $driveL
    }

    $menuDrive = $menuOpenDisk.DropDownItems.Add($menuText)
    $menuDrive.Image = [System.IconExtractor]::Extract("imageres.dll", 30, $true)
    $Script:driveL = $driveL

    $menuDrive.Add_Click( { Add-DataTable2DGV $script:driveL })
    [void]$menuOpenDisk.DropDownItems.Add($menuDrive)

}

# The parameters of the function are being called next to the functio nanme 
function Add-DataTable2DGV($driveLetter, $folderToScan) {
    # The string that Get-ForensicFileRecord accepts as Volume Name is in the format of C: or D: without \
    # The following 3 commands are making sure that the string that will be passed to the command has that
    # specific pattern 
    
    if ($folderToScan) { #if folder to Scan Exists
        $folderToScan = $folderToScan.ToString().SubString(12, $folderToScan.Length - 1)
        $selectedDrive = $folderToScan.ToString().Substring(0, 2)

    }
    else {
        $selectedDrive = $this.ToString().Trim()
        $selectedDrive = $selectedDrive.Substring(0, 2)
    }

     # start a stopwatch to count how many seconds it takes to finish the task
     $stopwatch = [system.diagnostics.stopwatch]::StartNew()

    # Get FilesNmes from MFT.
    $files = Get-Filenames -Drive $selectedDrive

    # Add 2 columns on the DataTable
    $Global:dataTable = $null
    $Global:dataTable = New-Object System.Data.DataTable
    [void]$Global:dataTable.Columns.Add("File Path")
    [void]$Global:dataTable.Columns.Add("Length")

    # Hide column headers for faster dataBinding
    $dgvFilePaths.ColumnHeadersVisible = $false
    # Initiate variable only once and convert the string to int32
    $lengthSettingValue = Get-IniFileSettings -SettingsCategory "Settings" -SettingValue "pathlength"

    if ($lengthSettingValue -eq $false) {
        $length = 200
    }
    else {
        $length = $lengthSettingValue
    }
   
    # For each files in the 
    if (!$folderToScan) { #Not scanning for a folder
        foreach ($file in $files) {                         
            if ($file.FullName.length -ge $length) {
                $Global:dataTable.Rows.Add($file.FullName, $file.FullName.Length)
            }
        }
    }
    else { #scanning a folder
       <#  foreach ($file in $files) {                         
            if ($file.FullName.length -ge $length -and $file.FullName -like "$folderToScan*") {
                $Global:dataTable.Rows.Add($file.FullName, $file.FullName.Length)
            }

        }  #>
        $folderToScan = $folderToScan.Replace("\","\\")
        [regex]$folderToScanRegex = "($folderToScan)"
        foreach ($file in $files) {                         
            if ($file.FullName.length -ge $length -and $file.FullName -match $folderToScanRegex -eq $true) {
                $Global:dataTable.Rows.Add($file.FullName, $file.FullName.Length)
            }

        } 
    }

        $dgvFilePaths.DataSource = $Global:dataTable

        $dgvFilePaths.ColumnHeadersVisible = $true
        $dgvFilePaths.AllowUserToAddRows = $false
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
        #$sw.Stop()

    }

    function gridClick {
        # This is called when the user clicks on the DataGridView
        # It gets the row and column index
        # and if it's a specific column index, it launches explorer 
        # on the value of the row on path column
        $rowIndex = $dgvFilePaths.CurrentRow.Index

        $columnIndex = $dgvFilePaths.CurrentCell.ColumnIndex
        if ($columnIndex -eq 0) {

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
    
        param([string]$initialDirectory,
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
        
    
        $Results = @()
        ForEach ($Job in $Jobs) {
            $Results += $Job.Pipe.EndInvoke($Job.Result)
        }
    
        return $results
 
    }   

    # Show Form
    $Form.ShowDialog()

