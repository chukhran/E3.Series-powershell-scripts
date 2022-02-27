<#
    ===========================================================================
    Name:           Batch Device Renaming
    
    Description:    BatchDeviceRenaming is a script for batch processing used
                    to rename selected devices in E3.Series to save time and 
                    reduce the amount of work involved.
    
    Requirements:   E3.Series 2016 and above
                    PowerShell 5.0 and above
    -------------------------------------------------------------------------
     Filename:      BatchDeviceRenaming.ps1
     Created by:    Dmytro Chukhran
     Support:       https://github.com/chukhran/E3.Series-powershell-scripts
     Created with:  Windows PowerShell ISE
    -------------------------------------------------------------------------
     Changes:
        27/02/2022 Initial Version

    ===========================================================================
#>


# It requires for [Windows.Markup.XamlReader]
Add-Type -Assembly PresentationFramework

#
 Class E3Wrapper {

    [__ComObject] $e3series
    [__ComObject] $project
    [__ComObject] $device

    [String] $separator
    [Array] $deviceIds


    # Constructor
    E3Wrapper(){
        $this.e3series = New-Object -ComObject 'CT.Application'
        $this.project = $this.e3series.CreateJobObject()
        $this.device = $this.project.CreateDeviceObject()
        $this.separator = $this.project.GetDeviceNameSeparator()
    }

    #
    [Void] Init(){
        # Is there an open project?
        if ($this.project.GetId() -eq 0) {
            throw 'There is not any open project'
        }

    }

    # Get all devices selected in a sheet and a device tree
    [Array] GetSelectedDeviceIds() {
        
        # Get devices selected in a sheet
        #--------------------------------
        $devShtIds = $null
        $null = $this.project.GetSelectedAllDeviceIds([ref]$devShtIds)


        # Sort devices by name
        # https://www.powershelladmin.com/wiki/Sort_strings_with_numbers_more_humanely_in_PowerShell.php
        # https://stackoverflow.com/questions/48292081/how-can-i-sort-by-column-in-powershell
        # https://stackoverflow.com/questions/1408042/output-data-with-no-column-headings-using-powershell
        $devShtIds = $devShtIds | ForEach-Object {
            $nulle = $this.device.SetId($_)
            $Field =  $this.device.GetName()

            # To get correct alphanumeric sorting, numbers in the string 
            # should be replaced with zero-pad numbers 
            $MaxDigitCount = 10
            $Field = [Regex]::Replace(
                $Field,
                '(\d+)',
                {"{0:D$MaxDigitCount}" -f [Int] $Args[0].Value}
            )

            New-Object -TypeName PSCustomObject -Property @{
                Id = $_
                Field = $Field
            }
        } | Sort-Object Field | Select Id -ExpandProperty Id


        # Get all devices selected in the device tree
        #--------------------------------------------
        $devTreeIds = $null
        $null = $this.project.GetTreeSelectedAllDeviceIds([ref]$devTreeIds)
        
        #
        $devIds = $devShtIds + $devTreeIds

        # Check terminals as devices !!!
        

        # Keep unique values only
        #------------------------
        $devIds = $devIds | Select-Object -Unique


        $this.deviceIds = $devIds
        return $devIds
    }


    # Get a new device name based on a mask name
    Hidden [String] GetNewDevName (
        [String]$nameMask,
        [string]$devName,
        [string]$counterValue)
    {
        # Initial new device name is as the mask value
        $newDevName = $nameMask

        # Exclude the separator from beginning of device name
        $devName = $devName -replace ('^'+$this.separator),''

        # Process [N#-#] placeholders
        #----------------------------
        $pattern = '\[N(\d+)-(\d+)\]'
        $matches = ([regex]$pattern).Matches($newDevName)
        #$matches[1].Groups

        # Go through found placeholders
        foreach($match in $matches) {
            $found = $match.Value # found pattern, for example, [N2-13]
            $from = [int]$match.Groups[1].Value - 1 # 2 form [N2-13]
            $to = [int]$match.Groups[2].Value - 1 # 13 from [N2-13]

            $subDevName = $devName[$from..$to] -join ''
            $newDevName = $newDevName.Replace($found, $subDevName)
        }

        # Process [N] placeholders
        #-------------------------
        $newDevName = $newDevName.Replace('[N]', $devName)

        # Process [C] placeholders
        #-------------------------
        $newDevName = $newDevName.Replace('[C]', $counterValue)

        # Return back the separator back
        $newDevName = $this.separator + $newDevName
        return $newDevName
    }

    # Get a string with preview names of the 1st and last devices
    [String] GetNamePreview() {

        $mask = $script:frmTxtMask.Text -as [string]
        
        $devCnt = $this.deviceIds.Length
        
        if($devCnt -eq 0) {
            return 'There are not any selected devices'
        }

        # The name of the first device
        $null = $this.device.SetId($this.deviceIds[0])
        $devName = $this.device.GetName()

        # Get a number of digits value from the main window
        $digits = $script:frmComDigits.SelectedIndex -as [int]
        $digits += 1
        $start = $script:frmTxtStart.Text -as [int]
        $step = $script:frmTxtStep.Text -as [int]
        $counterText = "{0:d$digits}" -f $start

        $return = $this.GetNewDevName($mask, $devName, $counterText)
        
        # If there are more than one device
        if($devCnt -gt 1) {
            
            # The name of the last device
            $null = $this.device.SetId($this.deviceIds[-1])
            $devName = $this.device.GetName()
            
            $counterText = "{0:d$digits}" -f ($start + ($step * ($devCnt - 1)))
            
            # Get formated name of the last device
            $devName = $this.GetNewDevName($mask, $devName, $counterText)
            
            # The final resuls as a name range
            $return = "{0} ... {1}" -f $return, $devName
            
        }        
        
        return $return
    }

    # Rename all selected devices
    [Void] RenameDevices() {

        # Generate a random string
        #-------------------------
        # ... from the upper case letters of the ASCII table
        $randomStr = -join ((65..90) | Get-Random -Count 7 | % {[char]$_})
        $randomStr = '-[TMP-{0}]' -f $randomStr

        # Data from the main window
        #--------------------------
        $mask = $script:frmTxtMask.Text -as [string]
        $counter = $script:frmTxtStart.Text -as [int]
        $step = $script:frmTxtStep.Text -as [int]
        $digits = $script:frmComDigits.SelectedIndex -as [int]
        $digits += 1

        # The hash of original devices' names
        $hashOrgNames = @{}

        # Error status: 0 is a renamming error
        $renameErr = 1


        # Rename devices to temporary names
        #----------------------------------
        foreach($id in $this.deviceIds){
            $null = $this.device.SetId($id)
            $devName = $this.device.GetName()
            $hashOrgNames.Add($id, $devName)

            $counterText = "{0:d$digits}" -f $counter
            $devTmpName = $this.GetNewDevName($mask, $devName, $counterText)
            $devTmpName += $randomStr

            $result = $this.device.SetName($devTmpName)

            if ($result -eq 0) {
                $renameErr = 0
            }

            $counter += $step
        }

        # Rename devices to target names
        #-------------------------------
        $counter = $script:frmTxtStart.Text -as [int]
        foreach($id in $this.deviceIds){
            $null = $this.device.SetId($id)
            #$devName = $this.device.GetName()
            $devOrgName = $hashOrgNames[$id]

            $counterText = "{0:d$digits}" -f $counter
            $devTrgName = $this.GetNewDevName($mask, $devOrgName, $counterText)

            $result = $this.device.SetName($devTrgName)

            if ($result -eq 0) {
                $renameErr = 0
            }

            $counter += $step
        }


        if ($renameErr -eq 0) {
            #[System.Windows.Forms.MessageBox]::Show('Hello')
            # does not work, see 
            # https://stackoverflow.com/a/56569543/17806220

            $this.e3series.PutError(1, 'Renaming errors exist! Please check an output window.')

        } else {

            $this.e3series.PutInfo(1, ("{0} device(s) renamed successfully!" -f $this.deviceIds.Count))
        }

    }

    # Add a placeholder into the mask field of the main window
    Static [Void] AddPlaceholder([String] $placeholder){
        $mask = $script:frmTxtMask

        # If there is not any selected text in the frmTxtMask field
        if ($mask.SelectedText -eq '') {
            $mask.Text = $mask.Text.Insert($mask.CaretIndex, $placeholder)
        } else {
            # Replace the selected text with the name placeholder
            $mask.SelectedText = $placeholder
        }

        # Move the caret after the inserted text
        $mask.CaretIndex += $placeholder.Length
        $mask.SelectionLength = 0
    }

 }
#-----


# The main window
# ---------------
[xml]$xaml = @"
<Window
 xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
 xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
 xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 ResizeMode="NoResize"
 Topmost = "True"
 WindowStartupLocation="CenterScreen"
Title="Batch devices renaming" Height="250" Width="450">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="12" />
            <ColumnDefinition Width="*" />
            <ColumnDefinition Width="8" />
            <ColumnDefinition Width="180" />
            <ColumnDefinition Width="12" />
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="12" />
            <RowDefinition Height="auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="12" />
            <RowDefinition Height="auto" />
        </Grid.RowDefinitions>

        <GroupBox Header="Device designation mask" Grid.Column="1" Grid.Row="1" Padding="4" Margin="0,0,0,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="6" />
                    <ColumnDefinition Width="*" />
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="auto" />
                    <RowDefinition Height="auto" />
                    <RowDefinition Height="auto" />
                </Grid.RowDefinitions>

                <TextBox x:Name="frmTxtMask" Text="[N]" Grid.ColumnSpan="3" Grid.Row="0" Padding="2" Margin="0,0,0,6" />
                <Button x:Name="frmBtnName" Content="[N] Name" Grid.Column="0" Grid.Row="1" Padding="2" />
                <Button x:Name="frmBtnCounter" Content="[C] Counter" Grid.Column="2" Grid.Row="1" Padding="2" />
                <Button x:Name="frmBtnRange" Content="[N#-#] Name range" Grid.ColumnSpan="3" Grid.Row="2" Padding="2" Margin="0,6,0,0" />
            </Grid>
        </GroupBox>

        <GroupBox Header="Define counter [C]" Grid.Column="3" Grid.Row="1" Padding="4" Margin="0,0,0,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="70" />
                    <ColumnDefinition Width="6" />
                    <ColumnDefinition Width="*" />
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="auto" />
                    <RowDefinition Height="auto" />
                    <RowDefinition Height="auto" />
                </Grid.RowDefinitions>

                <TextBlock Text="Start at:" Grid.Column="0" Grid.Row="0" Padding="2" Margin="0,0,0,6" HorizontalAlignment="Right" />
                <TextBlock Text="Step by:" Grid.Column="0" Grid.Row="1" Padding="2" HorizontalAlignment="Right" />
                <TextBlock Text="Digits:" Grid.Column="0" Grid.Row="2" Padding="2" Margin="0,6,0,0" HorizontalAlignment="Right" />

                <TextBox x:Name="frmTxtStart" Text="1" Grid.Column="2" Grid.Row="0" MaxLength="5" Padding="2" Margin="0,0,0,6" />
                <TextBox x:Name="frmTxtStep" Text="1" Grid.Column="2" Grid.Row="1" MaxLength="5" Padding="2" />

                <ComboBox x:Name="frmComDigits" Grid.Column="2" Grid.Row="2" Margin="0,6,0,0">
                    <ComboBoxItem IsSelected="True">1</ComboBoxItem>
                    <ComboBoxItem>2</ComboBoxItem>
                    <ComboBoxItem>3</ComboBoxItem>
                    <ComboBoxItem>4</ComboBoxItem>
                    <ComboBoxItem>5</ComboBoxItem>
                    <ComboBoxItem>6</ComboBoxItem>
                    <ComboBoxItem>7</ComboBoxItem>
                    <ComboBoxItem>8</ComboBoxItem>
                    <ComboBoxItem>9</ComboBoxItem>
                    <ComboBoxItem>10</ComboBoxItem>
                </ComboBox>
            </Grid>
        </GroupBox>

        <DockPanel Grid.Column="1" Grid.ColumnSpan="3" Grid.Row="2" >
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                </Grid.ColumnDefinitions>
<!--
                <StackPanel Grid.Column="0" Orientation="Horizontal" HorizontalAlignment="Left" VerticalAlignment="Center">
                    <TextBlock  Margin="8,0,0,0" Foreground="DarkBlue" ToolTip="Help" Cursor="Hand">           
                        <Hyperlink x:Name="Help" NavigateUri="https://github.com/">
                            Help
                        </Hyperlink>
                    </TextBlock>
                </StackPanel>
-->
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <Button x:Name="frmBtnRename" Content="Rename" Padding="4" Width="100" Margin="0,0,20,0" FontWeight="Bold" />
                    <Button Content="Close" Padding="4" Width="100" IsCancel="True" />
                </StackPanel>
            </Grid>
        </DockPanel>

    <DockPanel Grid.ColumnSpan="5" Grid.Row="4">
            <Border BorderBrush="DarkGray" BorderThickness="0,0.5,0,0">
                <StatusBar >
                    <StatusBar.ItemsPanel>
                        <ItemsPanelTemplate>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="12"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                            </Grid>
                        </ItemsPanelTemplate>
                    </StatusBar.ItemsPanel>

                    <StatusBarItem Grid.Column="1">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock>Selected items:</TextBlock>
                            <TextBlock x:Name="frmLblItems" Margin="6,0,0,0">0</TextBlock>
                        </StackPanel>
                    </StatusBarItem>
                    <Separator Grid.Column="2"/>
                    <StatusBarItem Grid.Column="3">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock>Preview:</TextBlock>
                            <TextBlock x:Name="frmLblPreview" Margin="6,0,0,0">-X-001</TextBlock>
                        </StackPanel>
                    </StatusBarItem>
                </StatusBar>
            </Border>
        </DockPanel>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
# ---------------


# Icon of the main window
#------------------------
[string]$base64=@"
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
"@
$bitmap = New-Object System.Windows.Media.Imaging.BitMapImage
$bitmap.BeginInit()
$bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($base64)
$bitmap.EndInit()
$bitmap.Freeze()
$window.Icon = $bitmap
#-----


# Store objects of window form as PowerShell variables
#-----------------------------------------------------
#$xaml.SelectNodes("//*[@Name]")
$xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | %{
    #$_.Name
    Set-Variable -Name ($_.Name) -Value $window.FindName($_.Name) -Scope Global
}

#------

try {

    #
    $e3 = [E3Wrapper]::new()
    $e3.Init()

    # Get unique selected devices
    $devIds = $e3.GetSelectedDeviceIds()

    # Set a value of Selected unique items field in the main window
    $frmLblItems.Text = $devIds.Count -as [string]

    # Set a value of Preview field in the main window
    # See Data binding section below


    # Data binding
    #-------------
    # Create a datacontext for the textbox and set it
    $DataContext = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
    $Text = $e3.GetNamePreview()
    $DataContext.Add($Text)
    $frmLblPreview.DataContext = $DataContext

    # Create and set a binding on the textbox object
    $Binding = New-Object System.Windows.Data.Binding # -ArgumentList "[0]"
    $Binding.Path = "[0]"
    $Binding.Mode = [System.Windows.Data.BindingMode]::OneWay
    [void][System.Windows.Data.BindingOperations]::SetBinding($frmLblPreview,[System.Windows.Controls.TextBlock]::TextProperty, $Binding)

    #-----



    #
    $frmBtnRename.Add_Click{

        $null = $e3.RenameDevices()
    }

    #
    $frmBtnName.Add_Click{
        $null = [E3Wrapper]::AddPlaceholder('[N]')
    }

    #
    $frmBtnCounter.Add_Click{
        $null = [E3Wrapper]::AddPlaceholder('[C]')
    }

    #
    $frmBtnRange.Add_Click{
        $null = [E3Wrapper]::AddPlaceholder('[N#-#]')
    }



    #Allow numbers input only
    $frmTxtStart.Add_KeyUp{
        $frmTxtStart.text = $frmTxtStart.text -replace '[^0-9]',''
    }

    #Allow numbers input only
    $frmTxtStep.Add_KeyUp{
        $frmTxtStep.text = $frmTxtStep.text -replace '[^0-9]',''
    }
    <#
    #>


    #
    $frmTxtStart.Add_TextChanged{
        # Update a value of Preview field in the main window
        $DataContext[0] = $e3.GetNamePreview()
    }

    #
    $frmTxtStep.Add_TextChanged{
        # Update a value of Preview field in the main window
        $DataContext[0] = $e3.GetNamePreview()
    }

    #
    $frmTxtMask.Add_TextChanged{
        # Update a value of Preview field in the main window
        $DataContext[0] = $e3.GetNamePreview()
    }

    #
    $frmComDigits.Add_SelectionChanged{
        # Update a value of Preview field in the main window
        $DataContext[0] = $e3.GetNamePreview()
    }

    #
    $window.Add_Activated{
        # Update a value of Selected unique items field in the main window
        $devIds = $e3.GetSelectedDeviceIds()
        $frmLblItems.Text = $devIds.Count -as [string]

        # Update a value of Preview field in the main window
        $DataContext[0] = $e3.GetNamePreview()
    <##>
    }


    #$window.WindowStyle = "ToolWindow"
    $window.ShowDialog() | Out-Null

} catch {
    $null = $e3.e3series.PutError(1, $_.Exception.Message)
    Write-Host $_.Exception.Message
}