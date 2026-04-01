# Active Directory Export Tool v1.1
# Author: Nikolaos Karanikolas — https://karanik.gr
# Requires RSAT (ActiveDirectory) and GroupPolicy modules where applicable.

$script:AppVersion = "1.1.0"

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Try to load Windows API Code Pack for modern folder picker (if DLLs exist next to script)
$global:UseModernDialog = $false
try {
    if ($PSCommandPath) {
        $basePath = Split-Path -Parent $PSCommandPath
    }
    else {
        # Fallback for ps2exe-compiled .exe or interactive session
        try {
            $basePath = [System.IO.Path]::GetDirectoryName(
                [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
            )
        }
        catch {
            $basePath = Get-Location
        }
    }
    $dll1 = Join-Path $basePath "Microsoft.WindowsAPICodePack.dll"
    $dll2 = Join-Path $basePath "Microsoft.WindowsAPICodePack.Shell.dll"
    if ((Test-Path $dll1 -PathType Leaf) -and (Test-Path $dll2 -PathType Leaf)) {
        Add-Type -Path $dll1
        Add-Type -Path $dll2
        $global:UseModernDialog = $true
    }
}
catch {
    $global:UseModernDialog = $false
}

# XAML UI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Active Directory Export Tool"
        Width="800" Height="650"
        WindowStartupLocation="CenterScreen"
        Background="#f7f7f7"
        ResizeMode="CanResize">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>   <!-- Menu -->
            <RowDefinition Height="*"/>      <!-- Main content -->
            <RowDefinition Height="Auto"/>   <!-- Footer -->
        </Grid.RowDefinitions>

        <!-- Menu bar -->
        <Menu Grid.Row="0">
            <MenuItem Header="_File">
                <MenuItem x:Name="menuSelectFolder" Header="Select output folder"/>
                <Separator/>
                <MenuItem x:Name="menuSaveLog" Header="Save log..."/>
                <Separator/>
                <MenuItem x:Name="menuExit" Header="Exit"/>
            </MenuItem>
            <MenuItem Header="_Info">
                <MenuItem x:Name="menuModules" Header="Modules status"/>
                <MenuItem x:Name="menuAbout" Header="About"/>
            </MenuItem>
        </Menu>

        <!-- Main content -->
        <Grid Grid.Row="1" Margin="10,10,10,5">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>   <!-- Title -->
                <RowDefinition Height="Auto"/>   <!-- File type -->
                <RowDefinition Height="Auto"/>   <!-- Folder -->
                <RowDefinition Height="Auto"/>   <!-- Progress -->
                <RowDefinition Height="Auto"/>   <!-- Options checkboxes -->
                <RowDefinition Height="Auto"/>   <!-- Buttons -->
                <RowDefinition Height="*"/>      <!-- Log -->
            </Grid.RowDefinitions>

            <!-- Title -->
            <TextBlock Grid.Row="0"
                       Text="Active Directory Export Tool"
                       FontSize="24"
                       FontWeight="Bold"
                       Margin="0,0,0,10"/>

            <!-- File type selection -->
            <StackPanel Grid.Row="1"
                        Orientation="Horizontal"
                        Margin="0,5,0,5">
                <TextBlock Text="File type:"
                           VerticalAlignment="Center"
                           Margin="0,0,10,0"/>
                <ComboBox x:Name="comboType"
                          Width="140">
                    <ComboBoxItem Content="CSV" IsSelected="True"/>
                    <ComboBoxItem Content="TXT"/>
                </ComboBox>
            </StackPanel>

            <!-- Folder selection -->
            <StackPanel Grid.Row="2"
                        Orientation="Horizontal"
                        Margin="0,5,0,5">
                <Button x:Name="btnFolder"
                        Content="Select output folder"
                        Height="32"
                        Width="220"
                        Background="#e0e0e0"
                        BorderBrush="#d0d0d0"
                        Margin="0,0,10,0"/>
                <TextBlock x:Name="lblFolder"
                           Text="No folder selected"
                           VerticalAlignment="Center"
                           Foreground="Gray"
                           TextTrimming="CharacterEllipsis"
                           Width="420"/>
            </StackPanel>

            <!-- Progress area -->
            <StackPanel Grid.Row="3"
                        Orientation="Vertical"
                        Margin="0,10,0,10">
                <TextBlock Text="Progress:"
                           Margin="0,0,0,5"/>
                <ProgressBar x:Name="pbStatus"
                             Height="18"
                             Minimum="0"
                             Maximum="100"
                             Value="0"/>
                <TextBlock x:Name="lblStatus"
                           Text="Ready."
                           Margin="0,5,0,0"
                           Foreground="Gray"/>
            </StackPanel>

            <!-- Options checkboxes -->
            <StackPanel Grid.Row="4"
                        Orientation="Vertical"
                        Margin="0,5,0,0">
                <TextBlock Text="Options:"
                           Margin="0,0,0,5"/>
                <WrapPanel Margin="0,0,0,10">
                    <CheckBox x:Name="chkNested"
                              Content="Include nested members"
                              Margin="0,0,20,0"
                              ToolTip="When checked, group exports will recursively resolve all nested group memberships.&#x0a;&#x0a;Example: If Group A contains Group B, and Group B contains User1,&#x0a;the export will list User1 as a member of Group A.&#x0a;&#x0a;Without this option, only direct members are exported&#x0a;(Group B would appear as a member instead of its users)."/>
                    <CheckBox x:Name="chkFsmoFull"
                              Content="Include full Domain/Forest details"
                              ToolTip="When checked, FSMO export will include the complete&#x0a;Get-ADDomain and Get-ADForest output in addition to FSMO roles.&#x0a;&#x0a;Without this option, only the FSMO role holders table&#x0a;and Domain Controllers list are exported."/>
                </WrapPanel>
            </StackPanel>

            <!-- Export buttons -->
            <StackPanel Grid.Row="5"
                        Orientation="Vertical"
                        Margin="0,0,0,10">
                <TextBlock Text="Exports:"
                           Margin="0,0,0,5"/>
                <UniformGrid Columns="2"
                             Rows="3">
                    <Button x:Name="btnExportUsers"
                            Content="Export AD Users"
                            Margin="0,0,5,5"
                            Height="40"
                            Background="#d8eaff"/>

                    <Button x:Name="btnExportGroups"
                            Content="Export AD Groups"
                            Margin="5,0,0,5"
                            Height="40"
                            Background="#ffd8d8"/>

                    <Button x:Name="btnExportComputers"
                            Content="Export AD Computers"
                            Margin="0,0,5,5"
                            Height="40"
                            Background="#fff0c2"/>

                    <Button x:Name="btnExportOuTree"
                            Content="Export OU Tree"
                            Margin="5,0,0,5"
                            Height="40"
                            Background="#e8ffd8"/>

                    <Button x:Name="btnExportGpos"
                            Content="Export GPOs"
                            Margin="0,0,5,5"
                            Height="40"
                            Background="#e0e0ff"/>

                    <Button x:Name="btnExportFsmo"
                            Content="Export FSMO Roles + Domain Info"
                            Margin="5,0,0,5"
                            Height="40"
                            Background="#f0e0ff"/>
                </UniformGrid>
            </StackPanel>

            <!-- Log area -->
            <Grid Grid.Row="6"
                  Margin="0,5,0,0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <StackPanel Grid.Row="0"
                            Orientation="Horizontal"
                            Margin="0,0,0,5">
                    <TextBlock Text="Log:"
                               Margin="0,0,10,0"
                               VerticalAlignment="Center"/>
                    <Button x:Name="btnClearLog"
                            Content="Clear log"
                            Height="24"
                            Margin="0,0,5,0"/>
                    <Button x:Name="btnSaveLog"
                            Content="Save log..."
                            Height="24"/>
                </StackPanel>

                <TextBox Grid.Row="1"
                         x:Name="txtLog"
                         IsReadOnly="True"
                         VerticalScrollBarVisibility="Auto"
                         HorizontalScrollBarVisibility="Auto"
                         TextWrapping="Wrap"
                         Background="White"/>
            </Grid>
        </Grid>

        <!-- Footer -->
        <TextBlock Grid.Row="2"
                   Text="Note: Run this script on a domain-joined machine with RSAT / AD / Group Policy modules installed."
                   FontSize="11"
                   Foreground="Gray"
                   TextWrapping="Wrap"
                   Margin="10,0,10,5"/>
    </Grid>
</Window>
"@

# Load XAML
$reader = New-Object System.Xml.XmlNodeReader $xaml
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Bind controls
$comboType          = $Window.FindName("comboType")
$btnFolder          = $Window.FindName("btnFolder")
$lblFolder          = $Window.FindName("lblFolder")
$pbStatus           = $Window.FindName("pbStatus")
$lblStatus          = $Window.FindName("lblStatus")
$btnExportUsers     = $Window.FindName("btnExportUsers")
$btnExportGroups    = $Window.FindName("btnExportGroups")
$btnExportComputers = $Window.FindName("btnExportComputers")
$btnExportOuTree    = $Window.FindName("btnExportOuTree")
$btnExportGpos      = $Window.FindName("btnExportGpos")
$btnExportFsmo      = $Window.FindName("btnExportFsmo")
$chkNested          = $Window.FindName("chkNested")
$chkFsmoFull        = $Window.FindName("chkFsmoFull")
$txtLog             = $Window.FindName("txtLog")
$btnClearLog        = $Window.FindName("btnClearLog")
$btnSaveLog         = $Window.FindName("btnSaveLog")

# Menu controls
$menuSelectFolder   = $Window.FindName("menuSelectFolder")
$menuSaveLog        = $Window.FindName("menuSaveLog")
$menuExit           = $Window.FindName("menuExit")
$menuModules        = $Window.FindName("menuModules")
$menuAbout          = $Window.FindName("menuAbout")

[xml]$aboutXaml = @"
<Window
    xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
    Title='About'
    Width='420'
    Height='220'
    WindowStartupLocation='CenterOwner'
    ResizeMode='NoResize'
    Background='#fafafa'>

    <Grid Margin='20'>
        <StackPanel>

            <TextBlock FontSize='20'
                       FontWeight='Bold'
                       Margin='0,0,0,10'>
                Active Directory Export Tool
            </TextBlock>

            <TextBlock Margin='0,0,0,10'
                       TextWrapping='Wrap'>
                <Run Text='Tool created by '/>
                <Hyperlink x:Name='lnkSite'
                           NavigateUri='https://karanik.gr'>
                    karanik
                </Hyperlink>
            </TextBlock>

            <TextBlock TextWrapping='Wrap'>
                Exports Users, Groups, Computers, OUs, GPOs, FSMO roles.
                Supports CSV/TXT and nested group members.
            </TextBlock>

        </StackPanel>
    </Grid>
</Window>
"@



function Get-SelectedFileType {
    param([System.Windows.Controls.ComboBox]$Combo)
    $item = [System.Windows.Controls.ComboBoxItem]$Combo.SelectedItem
    return $item.Content.ToString()
}

function Ensure-AdModule {
    param([string]$Name)
    try {
        if (-not (Get-Module -Name $Name)) {
            Import-Module $Name -ErrorAction Stop
        }
        Write-Log "Module '$Name' is loaded."
        return $true
    }
    catch {
        [System.Windows.MessageBox]::Show("Module '$Name' is not available. Please install the required RSAT / management tools.")
        Write-Log "ERROR: Module '$Name' could not be loaded. $($_.Exception.Message)"
        return $false
    }
}

function Require-Folder {
    if ($lblFolder.Text -eq "No folder selected" -or [string]::IsNullOrWhiteSpace($lblFolder.Text)) {
        [System.Windows.MessageBox]::Show("Please select an output folder first.")
        return $false
    }
    return $true
}

function Set-ProgressState {
    param(
        [int]$Value,
        [string]$Text
    )
    $pbStatus.Value = $Value
    $lblStatus.Text = $Text
    Write-Log $Text
    # Pump WPF dispatcher to keep UI responsive during long operations
    [System.Windows.Forms.Application]::DoEvents()
}

function Write-Log {
    param([string]$Message)
    if (-not $txtLog) { return }
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$timestamp - $Message"
    $txtLog.AppendText($line + [Environment]::NewLine)
    $txtLog.ScrollToEnd()
}

# Folder selection (button + menu)
$folderSelectAction = {
    if ($global:UseModernDialog) {
        $dialog = New-Object Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog
        $dialog.IsFolderPicker = $true
        $dialog.Title = "Select output folder"
        if ($dialog.ShowDialog() -eq "OK") {
            $lblFolder.Text = $dialog.FileName
            Write-Log "Output folder selected: $($dialog.FileName)"
        }
    }
    else {
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Select output folder"
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $lblFolder.Text = $dialog.SelectedPath
            Write-Log "Output folder selected: $($dialog.SelectedPath)"
        }
    }
}

$btnFolder.Add_Click($folderSelectAction)
$menuSelectFolder.Add_Click($folderSelectAction)

# Clear log
$btnClearLog.Add_Click({
    $txtLog.Clear()
    Write-Log "Log cleared."
})

# Save log (button + menu)
$saveLogAction = {
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title = "Save log"
    $dlg.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $dlg.FileName = "AD_Export_Log.txt"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        [System.IO.File]::WriteAllText($dlg.FileName, $txtLog.Text)
        Write-Log "Log saved to $($dlg.FileName)"
        [System.Windows.MessageBox]::Show("Log saved to:`n$($dlg.FileName)")
    }
}

$btnSaveLog.Add_Click($saveLogAction)
$menuSaveLog.Add_Click($saveLogAction)

# Menu: Exit
$menuExit.Add_Click({
    Write-Log "Exit selected from menu."
    $Window.Close()
})

# Menu: Modules status
$menuModules.Add_Click({
    $adOK  = Get-Command Get-ADUser -ErrorAction SilentlyContinue
    $gpOK  = Get-Command Get-GPO    -ErrorAction SilentlyContinue

    $adText = if ($adOK) { "ActiveDirectory module: AVAILABLE" } else { "ActiveDirectory module: NOT FOUND" }
    $gpText = if ($gpOK) { "GroupPolicy module: AVAILABLE" } else { "GroupPolicy module: NOT FOUND" }

    Write-Log $adText
    Write-Log $gpText
    [System.Windows.MessageBox]::Show("$adText`n$gpText","Modules status")
})

# Menu: About
$menuAbout.Add_Click({
    # Re-parse XAML each time to avoid consumed XmlNodeReader on 2nd open
    [xml]$localAboutXaml = $aboutXaml.OuterXml
    $reader = New-Object System.Xml.XmlNodeReader $localAboutXaml
    $aboutWin = [Windows.Markup.XamlReader]::Load($reader)

    $link = $aboutWin.FindName("lnkSite")
    if ($link -ne $null) {
        $link.add_RequestNavigate({
            param($sender, $e)
            Start-Process $e.Uri.AbsoluteUri
            $e.Handled = $true
        })
    }

    # Show version in About window title
    $aboutWin.Title = "About — v$($script:AppVersion)"
    $aboutWin.Owner = $Window
    $aboutWin.ShowDialog() | Out-Null
})

# Export AD Users
$btnExportUsers.Add_Click({
    if (-not (Require-Folder)) { return }
    if (-not (Ensure-AdModule -Name "ActiveDirectory")) { return }

    try {
        Set-ProgressState -Value 0 -Text "Starting users export..."
        Set-ProgressState -Value 10 -Text "Collecting users..."
        $fileType = Get-SelectedFileType -Combo $comboType
        $fileName = "AD_Users." + $fileType.ToLower()
        $path = Join-Path $lblFolder.Text $fileName

        $users = Get-ADUser -Filter * -Properties SamAccountName,DisplayName,Mail,Enabled,WhenCreated,Department,Title

        $selection = $users | Select-Object SamAccountName, Name, DisplayName, Mail, Enabled, WhenCreated, Department, Title

        Set-ProgressState -Value 70 -Text "Exporting users to $fileName..."
        if ($fileType -eq "CSV") {
            $selection | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
        }
        else {
            $selection | Out-File -FilePath $path -Encoding UTF8
        }

        Set-ProgressState -Value 100 -Text "Users export completed."
        [System.Windows.MessageBox]::Show("Users exported to:`n$path")
    }
    catch {
        Set-ProgressState -Value 0 -Text "Error exporting users."
        Write-Log "ERROR exporting users: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Error exporting users: $($_.Exception.Message)")
    }
})

# Export AD Groups (with or without nested members)
$btnExportGroups.Add_Click({
    if (-not (Require-Folder)) { return }
    if (-not (Ensure-AdModule -Name "ActiveDirectory")) { return }

    try {
        Set-ProgressState -Value 0 -Text "Starting groups export..."
        $includeNested = $chkNested.IsChecked -eq $true
        if ($includeNested) {
            Write-Log "Export AD Groups WITH nested members started."
        }
        else {
            Write-Log "Export AD Groups (groups only) started."
        }

        Set-ProgressState -Value 5 -Text "Collecting groups..."
        $fileType = Get-SelectedFileType -Combo $comboType
        $groups = Get-ADGroup -Filter * -Properties GroupCategory,GroupScope,Description,Mail,ManagedBy,WhenCreated,WhenChanged

        if (-not $includeNested) {
            # Groups only
            $fileName = "AD_Groups." + $fileType.ToLower()
            $path = Join-Path $lblFolder.Text $fileName

            $selection = $groups | Select-Object Name,SamAccountName,GroupCategory,GroupScope,Description,Mail,ManagedBy,WhenCreated,WhenChanged

            Set-ProgressState -Value 70 -Text "Exporting groups to $fileName..."
            if ($fileType -eq "CSV") {
                $selection | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
            }
            else {
                $selection | Out-File -FilePath $path -Encoding UTF8
            }

            Set-ProgressState -Value 100 -Text "Groups export completed."
            [System.Windows.MessageBox]::Show("Groups exported to:`n$path")
        }
        else {
            # Groups + nested members, all in one file
            $fileName = "AD_Groups_With_Members." + $fileType.ToLower()
            $path = Join-Path $lblFolder.Text $fileName

            $results = New-Object System.Collections.Generic.List[object]
            $total = $groups.Count
            if ($total -lt 1) { $total = 1 }
            $index = 0

            foreach ($g in $groups) {
                $index++
                $percent = [int](($index / $total) * 100)
                Set-ProgressState -Value $percent -Text ("Processing group {0} of {1}: {2}" -f $index,$total,$g.Name)

                try {
                    $members = Get-ADGroupMember -Identity $g -Recursive -ErrorAction Stop
                }
                catch {
                    Write-Log "WARNING: Could not enumerate members of '$($g.Name)': $($_.Exception.Message)"
                    $members = $null
                }

                if ($members) {
                    foreach ($m in $members) {
                        $results.Add([PSCustomObject]@{
                            GroupName           = $g.Name
                            MemberSamAccountName= $m.SamAccountName
                            MemberName          = $m.Name
                            ObjectClass         = $m.objectClass
                        })
                    }
                }
                else {
                    # Group with no members
                    $results.Add([PSCustomObject]@{
                        GroupName           = $g.Name
                        MemberSamAccountName= ""
                        MemberName          = ""
                        ObjectClass         = ""
                    })
                }
            }

            Set-ProgressState -Value 95 -Text "Exporting groups with members to $fileName..."
            if ($fileType -eq "CSV") {
                $results | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
            }
            else {
                $results | Out-File -FilePath $path -Encoding UTF8
            }

            Set-ProgressState -Value 100 -Text "Groups with members export completed."
            [System.Windows.MessageBox]::Show("Groups with members exported to:`n$path")
        }
    }
    catch {
        Set-ProgressState -Value 0 -Text "Error exporting groups."
        Write-Log "ERROR exporting groups: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Error exporting groups: $($_.Exception.Message)")
    }
})

# Export AD Computers
$btnExportComputers.Add_Click({
    if (-not (Require-Folder)) { return }
    if (-not (Ensure-AdModule -Name "ActiveDirectory")) { return }

    try {
        Set-ProgressState -Value 0 -Text "Starting computers export..."
        Set-ProgressState -Value 10 -Text "Collecting computers..."
        $fileType = Get-SelectedFileType -Combo $comboType
        $fileName = "AD_Computers." + $fileType.ToLower()
        $path = Join-Path $lblFolder.Text $fileName

        $computers = Get-ADComputer -Filter * -Properties DNSHostName,OperatingSystem,OperatingSystemVersion,Enabled,LastLogonDate,WhenCreated

        $selection = $computers | Select-Object Name,SamAccountName,DNSHostName,OperatingSystem,OperatingSystemVersion,Enabled,LastLogonDate,WhenCreated

        Set-ProgressState -Value 70 -Text "Exporting computers to $fileName..."
        if ($fileType -eq "CSV") {
            $selection | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
        }
        else {
            $selection | Out-File -FilePath $path -Encoding UTF8
        }

        Set-ProgressState -Value 100 -Text "Computers export completed."
        [System.Windows.MessageBox]::Show("Computers exported to:`n$path")
    }
    catch {
        Set-ProgressState -Value 0 -Text "Error exporting computers."
        Write-Log "ERROR exporting computers: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Error exporting computers: $($_.Exception.Message)")
    }
})

# Export OU Tree
$btnExportOuTree.Add_Click({
    if (-not (Require-Folder)) { return }
    if (-not (Ensure-AdModule -Name "ActiveDirectory")) { return }

    try {
        Set-ProgressState -Value 0 -Text "Starting OU tree export..."
        Set-ProgressState -Value 10 -Text "Collecting OUs..."
        $fileName = "AD_OU_Tree.txt"
        $path = Join-Path $lblFolder.Text $fileName

        # Sort by CanonicalName for correct parent→child order
        $ous = Get-ADOrganizationalUnit -Filter * -Properties CanonicalName |
               Sort-Object CanonicalName

        $lines = New-Object System.Collections.Generic.List[string]

        foreach ($ou in $ous) {
            $dn = $ou.DistinguishedName
            $ouParts = ($dn -split ',') | Where-Object { $_ -like 'OU=*' }
            $level = $ouParts.Count
            if ($level -lt 1) { $level = 1 }
            $indent = ' ' * (($level - 1) * 2)
            $lines.Add(("{0}{1}" -f $indent, $ou.Name))
        }

        Set-ProgressState -Value 70 -Text "Exporting OU tree to $fileName..."
        $lines | Out-File -FilePath $path -Encoding UTF8

        Set-ProgressState -Value 100 -Text "OU tree export completed."
        [System.Windows.MessageBox]::Show("OU tree exported to:`n$path")
    }
    catch {
        Set-ProgressState -Value 0 -Text "Error exporting OU tree."
        Write-Log "ERROR exporting OU tree: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Error exporting OU tree: $($_.Exception.Message)")
    }
})

# Export GPOs
$btnExportGpos.Add_Click({
    if (-not (Require-Folder)) { return }
    if (-not (Ensure-AdModule -Name "GroupPolicy")) { return }

    try {
        Set-ProgressState -Value 0 -Text "Starting GPO export..."
        Set-ProgressState -Value 10 -Text "Collecting GPOs..."
        $fileType = Get-SelectedFileType -Combo $comboType
        $fileName = "GPO_List." + $fileType.ToLower()
        $path = Join-Path $lblFolder.Text $fileName

        $gpos = Get-GPO -All

        $selection = $gpos | Select-Object DisplayName,Id,Owner,CreationTime,ModificationTime,GpoStatus,UserVersion,ComputerVersion,WmiFilter

        Set-ProgressState -Value 70 -Text "Exporting GPOs to $fileName..."
        if ($fileType -eq "CSV") {
            $selection | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
        }
        else {
            $selection | Out-File -FilePath $path -Encoding UTF8
        }

        Set-ProgressState -Value 100 -Text "GPO export completed."
        [System.Windows.MessageBox]::Show("GPOs exported to:`n$path")
    }
    catch {
        Set-ProgressState -Value 0 -Text "Error exporting GPOs."
        Write-Log "ERROR exporting GPOs: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Error exporting GPOs: $($_.Exception.Message)")
    }
})

# Export FSMO roles + domain / forest info
$btnExportFsmo.Add_Click({
    if (-not (Require-Folder)) { return }
    if (-not (Ensure-AdModule -Name "ActiveDirectory")) { return }

    try {
        Set-ProgressState -Value 0 -Text "Starting FSMO / domain info export..."
        Set-ProgressState -Value 10 -Text "Collecting domain / forest / FSMO info..."
        $includeFull = $chkFsmoFull.IsChecked -eq $true
        $fileName = "AD_FSMO_and_Domain_Info.txt"
        $path = Join-Path $lblFolder.Text $fileName

        $domain = Get-ADDomain
        $forest = Get-ADForest
        $dcs    = Get-ADDomainController -Filter *

        if ($includeFull) {
            Write-Log "FSMO export WITH full Domain/Forest details."
            "Domain Information" | Out-File -FilePath $path -Encoding UTF8
            "------------------" | Out-File -FilePath $path -Encoding UTF8 -Append
            ($domain | Format-List * | Out-String) | Out-File -FilePath $path -Encoding UTF8 -Append
            "" | Out-File -FilePath $path -Encoding UTF8 -Append

            "Forest Information" | Out-File -FilePath $path -Encoding UTF8 -Append
            "------------------" | Out-File -FilePath $path -Encoding UTF8 -Append
            ($forest | Format-List * | Out-String) | Out-File -FilePath $path -Encoding UTF8 -Append
            "" | Out-File -FilePath $path -Encoding UTF8 -Append
        }
        else {
            Write-Log "FSMO export (roles and DCs only)."
            "Domain: $($domain.DNSRoot)" | Out-File -FilePath $path -Encoding UTF8
            "Forest: $($forest.Name)" | Out-File -FilePath $path -Encoding UTF8 -Append
            "" | Out-File -FilePath $path -Encoding UTF8 -Append
        }

        "FSMO Roles" | Out-File -FilePath $path -Encoding UTF8 -Append
        "----------" | Out-File -FilePath $path -Encoding UTF8 -Append

        $fsmoInfo = @(
            [PSCustomObject]@{ Role = "PDCEmulator";          Holder = $domain.PDCEmulator }
            [PSCustomObject]@{ Role = "RIDMaster";            Holder = $domain.RIDMaster }
            [PSCustomObject]@{ Role = "InfrastructureMaster"; Holder = $domain.InfrastructureMaster }
            [PSCustomObject]@{ Role = "SchemaMaster";         Holder = $forest.SchemaMaster }
            [PSCustomObject]@{ Role = "DomainNamingMaster";   Holder = $forest.DomainNamingMaster }
        )

        ($fsmoInfo | Format-Table -AutoSize | Out-String) | Out-File -FilePath $path -Encoding UTF8 -Append
        "" | Out-File -FilePath $path -Encoding UTF8 -Append

        "Domain Controllers" | Out-File -FilePath $path -Encoding UTF8 -Append
        "------------------" | Out-File -FilePath $path -Encoding UTF8 -Append
        ($dcs | Format-Table Name,IPv4Address,OperatingSystem -AutoSize | Out-String) | Out-File -FilePath $path -Encoding UTF8 -Append

        Set-ProgressState -Value 100 -Text "FSMO export completed."
        [System.Windows.MessageBox]::Show("FSMO roles and domain info exported to:`n$path")
    }
    catch {
        Set-ProgressState -Value 0 -Text "Error exporting FSMO roles."
        Write-Log "ERROR exporting FSMO roles: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Error exporting FSMO roles: $($_.Exception.Message)")
    }
})

# Show window
$Window.Topmost = $false
$Window.ShowDialog() | Out-Null
