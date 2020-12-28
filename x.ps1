
#$SW_HIDE, $SW_SHOW = 0, 5
#$TypeDef = '[DllImport("User32.dll")]public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'
#Add-Type -MemberDefinition $TypeDef -Namespace Win32 -Name Functions
#$hWnd = (Get-Process -Id $PID).MainWindowHandle
#$Null = [Win32.Functions]::ShowWindow($hWnd,$SW_HIDE)
#####
#
#
#
# Revise Codes
#     WPF - remove from code, add to WPF
#     Func - Seperate to function
#     Rm - Remove / not needed anymore


Remove-Variable * -ErrorAction SilentlyContinue
$ConfirmPreference = "None"
Copy-Item -Path "\\labtop\TempData\v3\v3\MainWindow.xaml"  -Destination C:\TempData\MainWindow.xaml -Force

$xamlPath = "C:\TempData\MainWindow.xaml"
$psexec = "C:\TempData\asm\PSExec.exe" # REVISE - REMOVE

######

$savedConfig = "C:\TempData\config.json"

$segChar = @{} # REVISE - WPF
$segChar.Add("segWarn", "")
$segChar.Add("segCaution", "")
$segChar.Add("segCheck", "")
$segChar.Add("segUp", "")
$segChar.Add("segDown", "")

Import-Module C:\TempData\func\func.psm1
Remove-Module internal
Import-Module C:\TempData\internal\internal.psm1

# generated hash tables used throughout tool
New-HashTables

# Import from JSON and add to hash table
Set-Config -ConfigPath $savedConfig -Type Import -ConfigHash $configHash

# process loaded data or creates initial item templates for various config datagrids
@('userPropList','compPropList','contextConfig','objectToolConfig','nameMapList', 'netMapList') | Set-InitialValues -ConfigHash $configHash -PullDefaults
@('userLogMapping','compLogMapping') | Set-InitialValues -ConfigHash $configHash

# matches config'd user/comp logins with default headers, creates new headers from custom values
$defaultList = @('User', 'DateRaw', 'LoginDc', 'ClientName', 'ComputerName', 'Ignore', 'Custom')
@('userLogMapping','compLogMapping') | Set-LoggingStructure -DefaultList $DefaultList -ConfigHash $configHash

# Add default values if they are missing
@('MSRA','MSTSC') | Set-RTDefaults -ConfigHash $configHash

# loaded required DLLs
foreach ($dll in ((Get-ChildItem C:\TempData\asm\ -Filter *.dll).FullName)) {
    [System.Reflection.Assembly]::LoadFrom($dll) | Out-Null   
}

# read xaml and load wpf controls into synchash (named synchash)
Set-WPFControls $syncHash -XAMLPath $xamlPath

# builds custom WPF controls from whatever was defined and saved in ConfigHash
Add-CustomRTControls -SyncHash $syncHash -ConfigHash $configHash
Add-CustomItemBoxControls -SyncHash $syncHash -ConfigHash $configHash
Add-CustomToolControls -SyncHash $syncHash -ConfigHash $configHash
$sysCheckHash.missingCheckFail = $false


# parse custom RT settings



$syncHash.itemToolGridItemsGrid.Add_SelectionChanged({

    if ($syncHash.itemToolGridItemsGrid.SelectedItem.Image) {
        $syncHash.itemToolImageSource.Source = [byte[]]($syncHash.itemToolGridItemsGrid.SelectedItem.Image)
    }

})

$syncHash.itemToolGridItemsGrid.Add_AutoGeneratingColumn({

    if ($_.Column.Header -eq 'Image') {
        $_.Cancel = $true 
        $syncHash.itemToolImageBorder.Visibility = "Visible"
    }

})


# REVISE - FUNCTION
$syncHash.ItemToolADSelectionButton.Add_Click({



 Start-RSJob -Name PopulateListboxAD -ArgumentList $configHash, $syncHash, $syncHash.itemToolListSelectConfirmButton.Tag -FunctionsToImport Select-ADObject -ScriptBlock {
 param($configHash, $syncHash, $toolID)


    $syncHash.Window.Dispatcher.Invoke([Action]{
        $syncHash.itemToolListSelectListBox.ItemsSource = $null
    })

     $selectedObject = (Select-ADObject -Type All -MultiSelect $false).FetchedAttributes -replace '$'
     $syncHash.Window.Dispatcher.Invoke([Action]{
        $syncHash.itemToolADSelectedItem.Content = $selectedObject
        $syncHash.itemTooListBoxProgress.Visibility = "Visible"
    })
     
     $list = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
     
    Invoke-Expression ($configHash.objectToolConfig[$toolId - 1].toolFetchCmd) | ForEach-Object {$list.Add([PSCUstomObject]@{'Name' = $_})}

    $syncHash.Window.Dispatcher.Invoke([Action]{
        $syncHash.itemToolListSelectListBox.ItemsSource = [System.Windows.Data.ListCollectionView]$list
        $syncHash.itemTooListBoxProgress.Visibility = "Collapsed"
     })
 }      

})

# REVISE - FUNCTION
$syncHash.ItemToolGridADSelectionButton.Add_Click({

 Start-RSJob -Name PopulateGridbox -ArgumentList $configHash, $syncHash, $syncHash.itemToolGridSelectConfirmButton.Tag -FunctionsToImport Select-ADObject -ScriptBlock {
 param($configHash, $syncHash, $toolID)

     $selectedObject = (Select-ADObject -Type All -MultiSelect $false).FetchedAttributes -replace '$'
     $syncHash.Window.Dispatcher.Invoke([Action]{
        $syncHash.itemToolGridADSelectedItem.Content = $selectedObject
        $syncHash.itemToolGridProgress.Visibility = "Visible"
    })
     
     $list = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
     
    Invoke-Expression ($configHash.objectToolConfig[$toolId - 1].toolFetchCmd) | ForEach-Object {$list.Add($_)}
     
    $syncHash.Window.Dispatcher.Invoke([Action]{
        $syncHash.itemToolGridItemsGrid.ItemsSource = [System.Windows.Data.ListCollectionView]$list
        $syncHash.itemToolGridProgress.Visibility = "Collapsed"
     })
 }      

})

$syncHash.itemToolListSearchBox.Add_TextChanged({
    $syncHash.itemToolListSelectListBox.ItemsSource.Filter = $null
    $syncHash.itemToolListSelectListBox.ItemsSource.Filter = {param ($item) $item -match $syncHash.itemToolListSearchBox.Text}

})

$synchash.itemToolGridSearchBox.Add_TextChanged({

    $syncHash.itemToolGridItemsGrid.ItemsSource.Filter = $null
    $syncHash.itemToolGridItemsGrid.ItemsSource.Filter = {param ($item) $item -match $syncHash.itemToolGridSearchBox.Text}

})
        

# REVISE - FUNCTION
$syncHash.itemToolListSelectConfirmButton.Add_Click({

    $itemList = $syncHash.itemToolListSelectListBox.SelectedItems.Name

    Start-RSJob -ArgumentList $configHash, $itemList, $syncHash.snackMsg.MessageQueue, $syncHash.itemToolListSelectConfirmButton.Tag -ScriptBlock {
    param($configHash, $itemList, $queue, $toolID) 

        $toolName = $configHash.objectToolConfig[$toolID - 1].toolActionToolTip
        $target = $configHash.currentTabItem

        try {

            Invoke-Expression -Command $configHash.objectToolConfig[$toolID - 1].toolTargetFetchCmd

            foreach ($selectedItem in $itemList) {
                Invoke-Expression -Command $configHash.objectToolConfig[$toolID - 1].toolAction
            }

            $queue.Enqueue("[$toolName]: SUCCESS: tool ran on $target")

        }

        catch {
             $queue.Enqueue("[$toolName]: FAIL: tool incomplete on $target")

        }
    }

    $syncHash.itemToolDialog.IsOpen = $false

})

$syncHash.itemToolListSelectConfirmCancel.Add_Click({
    $syncHash.itemToolDialog.IsOpen = $false
})
# REVISE - FUNCTION
$syncHash.itemToolGridSelectConfirmButton.Add_Click({

    $itemList = $syncHash.itemToolGridItemsGrid.SelectedItems

    Start-RSJob -ArgumentList $configHash, $itemList, $syncHash.snackMsg.MessageQueue, $syncHash.itemToolGridSelectConfirmButton.Tag -ScriptBlock {
    param($configHash, $itemList, $queue, $toolID) 

        $toolName = $configHash.objectToolConfig[$toolID - 1].toolActionToolTip
        $target = $configHash.currentTabItem

        try {

            Invoke-Expression -Command $configHash.objectToolConfig[$toolID - 1].toolTargetFetchCmd

            foreach ($selectedItem in $itemList) {
                Invoke-Expression -Command $configHash.objectToolConfig[$toolID - 1].toolAction
            }

            $queue.Enqueue("[$toolName]: SUCCESS: tool ran on $target")

        }

        catch {
             $queue.Enqueue("[$toolName]: FAIL: tool incomplete on $target")

        }
    }

    $syncHash.itemToolDialog.IsOpen = $false

})

$syncHash.itemToolGridSelectConfirmCancel.Add_Click({
    $syncHash.itemToolDialog.IsOpen = $false
})

# REVISE - FUNCTION
$syncHash.itemToolListSelectAllButton.Add_Click({
    
    
    if ($syncHash.itemToolListSelectListBox.Items.Count -eq  $syncHash.itemToolListSelectListBox.SelectedItems.Count) {
       $syncHash.itemToolListSelectListBox.UnselectAll()
    } 

    else {
        $syncHash.itemToolListSelectListBox.SelectAll()
    }
        
})
 
 # REVISE - FUNCTION
$syncHash.itemToolGridSelectAllButton.Add_Click({
    
    
    if ($syncHash.itemToolGridItemsGrid.Items.Count -eq  $syncHash.itemToolGridItemsGrid.SelectedItems.Count) {
       $syncHash.itemToolGridItemsGrid.UnselectAll()
    } 

    else {
        $syncHash.itemToolGridItemsGrid.SelectAll()
    }
        
})
  
  



$syncHash.settingRemoteClick.add_Click({
    Set-ChildWindow -Panel settingRTContent -Title "Configure Remote Connection Clients" -SyncHash $syncHash
})

$syncHash.settingNetworkClick.add_Click({
    Set-ChildWindow -Panel settingNetContent -Title "Networking Mappings" -SyncHash $syncHash -Height 275
})

$syncHash.settingNamingClick.add_Click({
    Set-ChildWindow -Panel settingNameContent -Title "Device Categorization" -SyncHash $syncHash -Height 275
})

$syncHash.settingUserPropClick.add_Click({
    $syncHash.settingPropContent.Visibility = "Visible"
    Set-ChildWindow -Panel settingUserPropContent -Title "User Property Mappings" -SyncHash $syncHash  
})

$syncHash.settingCompPropClick.add_Click({
    $syncHash.settingPropContent.Visibility = "Visible"
    Set-ChildWindow -Panel settingCompPropContent -Title "Computer Property Mappings" -SyncHash $syncHash 
})

$syncHash.settingObjectToolsClick.add_Click({
    $syncHash.settingPropContent.Visibility = "Visible"
    Set-ChildWindow -Panel settingItemToolsContent -Title "Object Tools Mappings" -SyncHash $syncHash 
})

$syncHash.settingContextClick.add_Click({
    Set-ChildWindow -Panel settingContextPropContent -Title "Contextual Actions Mappings" -SyncHash $syncHash
})


$syncHash.settingLoggingClick.add_Click( {
    if (!($configHash.pcLogPath) -or !(Test-Path $configHash.pcLogPath)) {$syncHash.compLogPopupButton.IsEnabled = $false}
    if (!($configHash.userLogPath) -or !(Test-Path $configHash.userLogPath)) {$syncHash.userLogPopupButton.IsEnabled = $false}

    Set-ChildWindow -Panel settingLoggingContent -Title "Login Log Paths" -SyncHash $syncHash -Height 300 -Width 480
})

### Glyph load
$glyphs = Get-Content C:\TempData\segoeGlyphs.txt

$configHash.buttonGlyphs = [System.Collections.ArrayList]@()

$glyphs | ForEach-Object { $configHash.buttonGlyphs.Add($_) | Out-Null }

#region settingload / systems check
$syncHash.settingLogo.add_Loaded( {
        $syncHash.Window.Activate()
        
        $sysCheckHash.sysChecks = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
            $sysCheckHash.sysChecks.Add([PSCustomObject]@{
                'ADModule'           = 'False'
                'RSModule'           = 'False'
                'ADMember'           = 'False'
                'ADDCConnectivity'   = 'False'
                'IsInAdmin'          = 'False'
                'IsDomainAdmin'      = 'False'
                'Admin'              = 'False'
                'Modules'            = 'False'
                'ADDS'               = 'False'
         })

        $modList = Get-Module -ListAvailable ActiveDirectory, PoshRSJob | Sort-Object -Unique

        if ($modList.Name -contains "ActiveDirectory") { $sysCheckHash.sysChecks[0].ADModule = 'True' }

        if ($modList.Name -contains "PoshRSJob") {
            
            $sysCheckHash.sysChecks[0].RSModule = 'True'

            Start-RSJob -Name init -ArgumentList $syncHash, $psexec, $segChar, $sysCheckHash, $configHash, $savedConfig -ModulesToImport ActiveDirectory, C:\TempData\internal\internal.psm1 -ScriptBlock {        
                Param($syncHash, $psExec, $segChar, $sysCheckHash, $configHash, $savedConfig)
             
                Start-BasicADCheck -SysCheckHash $sysCheckHash
                
                Start-AdminCheck -SysCheckHash $sysCheckhash
                   
                # Check individual checks; mark parent categories as true is children are true       
                switch ($sysCheckHash.sysChecks) {
                    {$_.ADModule -eq $true -and $_.RSModule -eq $true} {$sysCheckHash.sysChecks[0].Modules = 'True'}
                    {$_.ADMember -eq $true -and $_.ADDCConnectivity -eq $true} {$sysCheckHash.sysChecks[0].ADDS = 'True'}
                    {$_.IsInAdmin -eq $true -and $_.IsDomainAdmin -eq $true} {$sysCheckHash.sysChecks[0].Admin = 'True'}
                }

                @('settingADMemberLabel','settingADDCLabel','settingModADLabel','settingModRSLabel','settingDomainAdminLabel',
                'settingLocalAdminLabel','settingPermLabel','settingADLabel','settingModLabel') | 
                Set-RSDataContext -SyncHash $syncHash -DataContext $sysCheckHash.sysChecks
               

                $sysCheckHash.checkComplete = $true


                ############ Begin Loading Window
                Start-Sleep -Seconds 2

                if ($sysCheckHash.sysChecks[0].ADDS -eq $false -or $sysCheckHash.sysChecks[0].Modules -eq $false -or $sysCheckHash.sysChecks[0].Admin -eq $false) {

                    Suspend-FailedItems -SyncHash $syncHash -CheckedItems SysCheck

                }
                
                else {

                    Start-PropBoxPopulate -ConfigHash $configHash  
                    
                }  
                
                $syncHash.Window.Dispatcher.invoke([action] {                       
                    $syncHash.windowContent.Visibility = "Visible"
                    $syncHash.Window.MinWidth = "1000"
                    $syncHash.Window.MinHeight = "700"
                    $syncHash.Window.ResizeMode = "CanResizeWithGrip"
                    $syncHash.Window.ShowTitleBar = $true
                    $syncHash.Window.ShowCloseButton = $true                   
                    $syncHash.splashLoad.Visibility = "Collapsed" 
                })           
            }
        }
        else {                                   
            Suspend-FailedItems -SyncHash $syncHash -CheckedItems RSCheck

            $syncHash.windowContent.Visibility = "Visible"
            $syncHash.Window.MinWidth = "1000"
            $syncHash.Window.MinHeight = "700"
            $syncHash.Window.ResizeMode = "CanResizeWithGrip"
            $syncHash.Window.ShowTitleBar = $true
            $syncHash.Window.ShowCloseButton = $true                   
            $syncHash.splashLoad.Visibility = "Collapsed" 
        }
    })

#endregion 

#region Configurable popup items


#endregion

#region Config events/logic

$syncHash.settingRtRDPClick.Add_Click( {

        $syncHash.settingRtExeSelect.Visibility = "Hidden"
        $syncHash.settingRtPathSelect.Visibility = "Hidden"
        $syncHash.rtSettingRequiresOnline.Visibility = "Hidden"
        $syncHash.rtSettingRequiresUser.Visibility = "Hidden"
        $syncHash.rtDock.DataContext = $configHash.rtConfig.mstsc
        $syncHash.settingRemoteFlyout.isOpen = $true
    })

$syncHash.settingRtMSRAClick.Add_Click( {
        $syncHash.settingRtExeSelect.Visibility = "Hidden"
        $syncHash.settingRtPathSelect.Visibility = "Hidden"
        $syncHash.rtSettingRequiresOnline.Visibility = "Hidden"
        $syncHash.rtSettingRequiresUser.Visibility = "Hidden"
        $syncHash.rtDock.DataContext = $configHash.rtConfig.MSRA
        $syncHash.settingRemoteFlyout.isOpen = $true
     

    })

$syncHash.settingRemoteFlyout.Add_OpeningFinished( {

   
        $syncHash.settingRemoteListTypes.ItemsSource = $configHash.nameMapList
 
        switch ($syncHash.settingRALabel.Content) {
            
            'MSTSC' {
                if ($configHash.rtConfig.MSTSC.Icon) {
                    $syncHash.settingRTIcon.Source = ([Convert]::FromBase64String($configHash.rtConfig.MSTSC.Icon))
                }

                $syncHash.settingRemoteListTypes.Items | Where-Object {$_.Name -in $configHash.rtConfig.MSTSC.Types} | ForEach-Object { 
                    $syncHash.settingRemoteListTypes.SelectedItems.Add(($_))
                }
                break
            }
            'MSRA' {
                if ($configHash.rtConfig.MSRA.Icon) {
                    $syncHash.settingRTIcon.Source = ([Convert]::FromBase64String($configHash.rtConfig.MSRA.Icon))
                }
                 $syncHash.settingRemoteListTypes.Items | Where-Object {$_.Name -in $configHash.rtConfig.MSRA.Types} | ForEach-Object { 
                    $syncHash.settingRemoteListTypes.SelectedItems.Add(($_))
                }
                break
            }
            Default {
                $rtID = 'rt' + [string]($syncHash.settingRALabel.Content -replace ".[A-Z]* ")

                if ($configHash.rtConfig.$rtID.Icon) {
                    $syncHash.settingRTIcon.Source = ([Convert]::FromBase64String($configHash.rtConfig.$rtID.Icon))
                }
               
                 $syncHash.settingRemoteListTypes.Items | Where-Object {$_.Name -in $configHash.rtConfig.$rtID.Types} | ForEach-Object { 
                    $syncHash.settingRemoteListTypes.SelectedItems.Add(($_))
                }

                break
            }
        }
    

   

  
    })

$syncHash.settingNetFlyoutExit.Add_Click( {
    
        $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.Window.BorderBrush.Color).ToString()
        $syncHash.settingChildWindow.Title = "Map Comp Login Logs"
        $syncHash.settingChildWindow.ShowCloseButton = $true
        $syncHash.settingLoggingCompFlyout.IsOpen = $true
        $syncHash.settingNetFlyout.IsOpen = $false
       
        [Array]$configHash.netMapList | ForEach-Object {
            if ($_.ValidMask -ne $true -and $_.ValidNetwork -ne $true) {
               [Array]$configHash.netMapList.RemoveAt([Array]::IndexOf($configHash.netMapList.ID,$_.ID))
            }
        }

    


    })

$syncHash.settingRemoteFlyoutExit.Add_Click( {
        $syncHash.settingRemoteFlyout.isOpen = $false
    
        $syncHash.settingRTIcon.Source = $null

        switch ($syncHash.settingRALabel.Content) {
            'MSTSC' { 
                $configHash.rtConfig.MSTSC.Types = @()
                $syncHash.settingRemoteListTypes.SelectedItems | ForEach-Object { $configHash.rtConfig.MSTSC.Types += $_.Name }
                break 
            }
            'MSRA' {
                $configHash.rtConfig.MSRA.Types = @()
                $syncHash.settingRemoteListTypes.SelectedItems | ForEach-Object { $configHash.rtConfig.MSRA.Types += $_.Name }
                break 
            }
            Default { 
                $rtID = 'rt' + [string]($syncHash.settingRALabel.Content -replace ".[A-Z]* ")
                $configHash.rtConfig.$rtID.Types = @()
                $syncHash.settingRemoteListTypes.SelectedItems | ForEach-Object { $configHash.rtConfig.$rtID.Types += $_.Name }
                break
            }


        }

        $syncHash.settingRemoteListTypes.SelectedItems.Clear()   

    })

$syncHash.settingRtExeSelect.Add_Click( {
    
    $rtID = 'rt' + [string]($syncHash.settingRALabel.Content -replace ".[A-Z]* ")
    $customSelection = New-Object System.Windows.Forms.OpenFileDialog
    $customSelection.initialDirectory = [Environment]::GetFolderPath('ProgramFilesx86')
    $customSelection.Title = "Select Custom RT Executable"
    $customSelection.ShowDialog() | Out-Null

    if (![string]::IsNullOrEmpty($customSelection.fileName)) {

        if (Test-Path $customSelection.fileName) {
            $configHash.rtConfig.$rtID.Path = $customSelection.fileName
            $syncHash.settingRtPathSelect.Text = $customSelection.fileName
            $configHash.rtConfig.$rtID.Icon = Get-Icon -Path $customSelection.fileName -ToBase64
            $syncHash.settingRTIcon.Source = ([Convert]::FromBase64String(($configHash.rtConfig.$rtID.Icon)))
        }
    }
})


$syncHash.userLogPopupButton.Add_Click( {
     
        if (!($configHash.userLogMapping)) { 

            $testLog = Get-Content ((Get-ChildItem -Path $confighash.UserLogPath | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName) | Where-Object { $_ -notmatch ",.[(\W)]$" -and $_ -notmatch ",$" } | Select-Object -First 1
            $fieldCount = ($testLog.ToCharArray() | Where-Object { $_ -eq ',' } | Measure-Object).Count + 1

            $header = @()
            $configHash.userLogMapping = New-Object System.Collections.ObjectModel.ObservableCollection[Object]

            for ($i = 1; $i -le $fieldCount; $i++) {
                $header += $i
               
            }

            $csv = $testLog | ConvertFrom-Csv -Header $header
            
            for ($i = 1; $i -le $fieldCount; $i++) {
                $configHash.userLogMapping.Add([PSCustomObject]@{
                        ID              = $i
                        Field           = $csv.$i
                        FieldSelList    = @('User', 'DateRaw', 'LoginDc', 'ClientName', 'ComputerName', 'Ignore', 'Custom')
                        FieldSel        = $null
                        CustomFieldName = $null
                        Ignore          = $false
                    }  
                )                                                 
               
            }

        }

        else {

            $testLog = Get-Content ((Get-ChildItem -Path $confighash.UserLogPath | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName) | Where-Object { $_ -notmatch ",.[(\W)]$" -and $_ -notmatch ",$"  -and $_.Trim() -ne ''} | Select-Object -Last 1
            $fieldCount = ($testLog.ToCharArray() | Where-Object { $_ -eq ',' } | Measure-Object).Count + 1

            $header = @()
            $userLogTemp = $configHash.userLogMapping
            $configHash.userLogMapping = New-Object System.Collections.ObjectModel.ObservableCollection[Object]

            for ($i = 1; $i -le $fieldCount; $i++) {
                $header += $i
               
            }

            $csv = $testLog | ConvertFrom-Csv -Header $header
            
            for ($i = 1; $i -le $fieldCount; $i++) {
                $configHash.userLogMapping.Add([PSCustomObject]@{
                        ID              = $i
                        Field           = $csv.$i
                        FieldSelList    = @('User', 'DateRaw', 'LoginDc', 'ClientName', 'ComputerName', 'Ignore', 'Custom')
                        FieldSel        = $userLogTemp[$i - 1].FieldSel
                        CustomFieldName = $userLogTemp[$i - 1].CustomFieldName
                        Ignore          = $false
                    }  
                )                                                 
               
            }

        }

        $syncHash.userLogListView.ItemsSource = $configHash.userLogMapping
        $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.settingLoggingUserFlyout.Background.Color).ToString()
        $syncHash.settingChildWindow.Title = "Map User Login Logs"
        $syncHash.settingChildWindow.ShowCloseButton = $false
        $syncHash.settingLoggingUserFlyout.IsOpen = $true
    

    
    })

$syncHash.compLogPopupButton.Add_Click( {
   
        if (!($configHash.compLogMapping)) { 

            $testLog = Get-Content ((Get-ChildItem -Path $configHash.pcLogPath | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName) | Where-Object { $_ -notmatch ",.[(\W)]$" -and $_ -notmatch ",$" -and $_.Trim() -ne '' } | Select-Object -First 1
            $fieldCount = ($testLog.ToCharArray() | Where-Object { $_ -eq ',' } | Measure-Object).Count + 1

            $header = @()
            $configHash.compLogMapping = New-Object System.Collections.ObjectModel.ObservableCollection[Object]

            for ($i = 1; $i -le $fieldCount; $i++) {
                $header += $i
               
            }

            $csv = $testLog | ConvertFrom-Csv -Header $header
            
            for ($i = 1; $i -le $fieldCount; $i++) {
                $configHash.compLogMapping.Add([PSCustomObject]@{
                        ID              = $i
                        Field           = $csv.$i
                        FieldSelList    = @('User', 'DateRaw', 'LoginDc', 'ClientName', 'ComputerName', 'Ignore', 'Custom')
                        FieldSel        = $null
                        CustomFieldName = $null
                        Ignore          = $false
                    }  
                )                                                 
               
            }

        }

        else {

            $testLog = Get-Content ((Get-ChildItem -Path $configHash.pcLogPath | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName) | Where-Object { $_ -notmatch ",.[(\W)]$" -and $_ -notmatch ",$" } | Select-Object -Last 1
            $fieldCount = ($testLog.ToCharArray() | Where-Object { $_ -eq ',' } | Measure-Object).Count + 1

            $header = @()
            $compLogTemp = $configHash.compLogMapping
            $configHash.compLogMapping = New-Object System.Collections.ObjectModel.ObservableCollection[Object]

            for ($i = 1; $i -le $fieldCount; $i++) {
                $header += $i
               
            }

            $csv = $testLog | ConvertFrom-Csv -Header $header
            
            for ($i = 1; $i -le $fieldCount; $i++) {
                $configHash.compLogMapping.Add([PSCustomObject]@{
                        ID              = $i
                        Field           = $csv.$i
                        FieldSelList    = @('User', 'DateRaw', 'LoginDc', 'ClientName', 'ComputerName', 'Ignore', 'Custom')
                        FieldSel        = $compLogTemp[$i - 1].FieldSel
                        CustomFieldName = $compLogTemp[$i - 1].CustomFieldName
                        Ignore          = $false
                    }  
                )                                                 
               
            }

        }

        $syncHash.compLogListView.ItemsSource = $configHash.compLogMapping
        $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.settingLoggingCompFlyout.Background.Color).ToString()
        $syncHash.settingChildWindow.Title = "Map Comp Login Logs"
        $syncHash.settingChildWindow.ShowCloseButton = $false
        $syncHash.settingLoggingCompFlyout.IsOpen = $true
    

    })

$syncHash.settingLoggingCompFlyoutExit.Add_Click( {

        $syncHash.settingLoggingCompFlyout.IsOpen = $false
        $syncHash.settingChildWindow.ShowCloseButton = $true
        $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.Window.BorderBrush.Color).ToString()
        $syncHash.settingChildWindow.Title = "Login Log Paths"

    })

$syncHash.settingLoggingUserFlyoutExit.Add_Click( {

        $syncHash.settingLoggingUserFlyout.IsOpen = $false
        $syncHash.settingChildWindow.ShowCloseButton = $true
        $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.Window.BorderBrush.Color).ToString()
        $syncHash.settingChildWindow.Title = "Login Log Paths"

    })

$syncHash.settingObjectToolFlyoutExit.Add_Click( {

        $syncHash.settingObjectToolDefFlyout.IsOpen = $false
        $syncHash.settingChildWindow.ShowCloseButton = $true
        $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.Window.BorderBrush.Color).ToString()
        $syncHash.settingChildWindow.Title = "Object Tools Mappings"

    })




$syncHash.settingNameMapClick.Add_Click( {
    
    if ($configHash.nameMapList.GetType().Name -ne 'ListCollectionView') {
        $configHash.nameMapListView = [System.Windows.Data.ListCollectionView]$configHash.nameMapList 
        $configHash.nameMapListView.IsLiveSorting = $true
        
        
        $syncHash.settingNameDataGrid.ItemsSource = $configHash.nameMapListView         
        $configHash.nameMapListView.Refresh()
    }
        $configHash.nameMapListView.LiveSortingProperties.Add("Id")   
        $configHash.nameMapListView.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription("Id", "Descending")))
        $syncHash.settingNameDataGrid.Visibility = "Visible"
        
        Set-ChildWindow -SyncHash $syncHash -Title "Computer Categorization Rules" -HideCloseButton -BackGround Flyout

        $syncHash.settingNameFlyout.IsOpen = $true
    })

$syncHash.settingNameAddClick.Add_Click( {
    
        if (($configHash.nameMapList | Measure-Object).Count -gt 1) {
            ($configHash.nameMapList | Sort-Object -Property ID -Descending | Select-Object -First 1).TopPos = $false
        }

        $configHash.nameMapList.Add([PSCustomObject]@{
                Id        = ($configHash.nameMapList.ID | Sort-Object -Descending | Select-Object -First 1) + 1
                Name      = $null
                Condition = $null
                topPos    = $true
            })    


        $syncHash.settingNameDataGrid.Items.Refresh()                   

})

$syncHash.settingNameFlyoutExit.Add_Click( {
    Reset-ChildWindow -SyncHash $syncHash -Title "Computer Categorization" -SkipContentPaneReset 
})


$syncHash.settingContextDefFlyout.Add_OpeningFinished({
    $syncHash.settingContextListTypes.ItemsSource = $configHash.nameMapList
    $syncHash.settingContextListTypes.Items | Where-Object {$_.Name -in $configHash.contextConfig[$syncHash.settingContextPropGrid.SelectedItem.IDNUM - 1].Types} | ForEach-Object { 
    $syncHash.settingContextListTypes.SelectedItems.Add(($_))
    }
        
})

$syncHash.settingContextFlyoutExit.Add_Click( {
$configHash.contextConfig[$syncHash.settingContextPropGrid.SelectedItem.IDNUM - 1].Types = @()
$syncHash.settingContextListTypes.SelectedItems | ForEach-Object { $configHash.contextConfig[$syncHash.settingContextPropGrid.SelectedItem.IDNUM - 1].Types += $_.Name } 

        $syncHash.settingContextDefFlyout.IsOpen = $false
        $syncHash.settingChildWindow.ShowCloseButton = $true
        $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.Window.BorderBrush.Color).ToString()
        $syncHash.settingChildWindow.Title = "Contextual Action Buttons"
        $syncHash.settingRemoteListTypes.SelectedItems.Clear()   


    })


     
$syncHash.settingNetMapClick.Add_Click( {

      
        $syncHash.settingNetDataGrid.Visibility = "Visible"
        $syncHash.settingNetDataGrid.ItemsSource = $configHash.netMapList 
        $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.settingNetFlyout.Background.Color).ToString()
        $syncHash.settingChildWindow.Title = "Network Mappings"
        $syncHash.settingChildWindow.ShowCloseButton = $false
        $syncHash.settingNetFlyout.IsOpen = $true
    })




$syncHash.settingNetAddClick.Add_Click( {
    

        $configHash.netMapList.Add([PSCustomObject]@{
                ID           = ($configHash.netMapList.ID | Sort-Object -Descending | Select-Object -First 1) + 1
                Network      = $null
                ValidNetwork = $false
                Mask         = $null
                ValidMask    = $false
                Location     = "New"
            })                       

    })

$syncHash.settingNetImportClick.Add_Click( {

        $configHash.netMapList = (New-Object System.Collections.ObjectModel.ObservableCollection[Object])
        $subnets = Get-ADReplicationSubnet -Filter * -Properties * | Select-Object Name, Site, Location, Description

        for ($i = 1; $i -le (($subnets | Measure-Object).Count); $i++) {
                
            $configHash.netMapList.Add([PSCustomObject]@{
                    Id           = $i
                    Network      = ($subnets[$i - 1].Name -replace "//*.*", "")
                    ValidNetwork = $true
                    Mask         = ($subnets[$i - 1].Name -replace ".*/", "")
                    ValidMask    = $true
                    Location     = if ($subnets[$i - 1].Location -ne $null) { $subnets[$i - 1].Location }
                    elseif ($subnets[$i - 1].Description -ne $null) { $subnets[$i - 1].Description }
                    
                })                                                               
        }


       

        $syncHash.settingNetDataGrid.ItemsSource = $configHash.netMapList
    

    })



$syncHash.settingUserPropMapClick.Add_Click( {

        Start-RSJob -ArgumentList $syncHash, $configHash -ScriptBlock {
            Param($syncHash, $configHash)

            $syncHash.Window.Dispatcher.invoke([action] {
        
                    if ($configHash.UserPropList -ne $null) {
                        $syncHash.settingUserPropGrid.ItemsSource = $null
                        $syncHash.settingUserPropGrid.ItemsSource = $configHash.UserPropList
                    }

                    if ($syncHash.settinguserPropGrid.Visibility -eq "Visible") {
                        $syncHash.settingUserPropSaveClick.Visibility = "Hidden"
                        $syncHash.settingUserPropGrid.Visibility = "Hidden"          
                        $syncHash.settingUserPropImportClick.Visibility = "Hidden"
                        $syncHash.settingUserAddItemClick.Visibility = "Hidden"
                        $syncHash.settingChildHeight.Height = 225     
                        $syncHash.settingChildHeight.Width = 390
                        $syncHash.userPropGrid.HorizontalAlignment = "Left"
                               
                    }
            
                    else {
                        $syncHash.settingUserPropGrid.Visibility = "Visible"
                        $syncHash.settinguserPropSaveClick.Visibility = "Visible"
                        $syncHash.settingUserPropImportClick.Visibility = "Visible"
                        $syncHash.settingUserAddItemClick.Visibility = "Visible"
                        $syncHash.settingChildHeight.Height = 365 #330
                        $syncHash.settingChildHeight.Width = 600
                        $syncHash.userPropGrid.HorizontalAlignment = "Stretch"
                    }           
                })  
        }
    })

$syncHash.settingCompPropMapClick.Add_Click( {

        Start-RSJob -ArgumentList $syncHash, $configHash -ScriptBlock {
            Param($syncHash, $configHash)

            $syncHash.Window.Dispatcher.invoke([action] {
        
                    if ($configHash.CompPropList -ne $null) {
                        $syncHash.settingCompPropGrid.ItemsSource = $null
                        $syncHash.settingCompPropGrid.ItemsSource = $configHash.CompPropList
                    }

                    if ($syncHash.settingCompPropGrid.Visibility -eq "Visible") {
                        $syncHash.settingCompPropGrid.Visibility = "Hidden"          
                        $syncHash.settingCompPropImportClick.Visibility = "Hidden"
                        $syncHash.settingCompAddItemClick.Visibility = "Hidden"
                        $syncHash.settingChildHeight.Height = 225     
                        $syncHash.settingChildHeight.Width = 390
                        $syncHash.CompPropGrid.HorizontalAlignment = "Left"
                               
                    }
            
                    else {
                        $syncHash.settingCompPropGrid.Visibility = "Visible"
                        $syncHash.settingCompPropImportClick.Visibility = "Visible"
                        $syncHash.settingCompAddItemClick.Visibility = "Visible"
                        $syncHash.settingChildHeight.Height = 365 #330
                        $syncHash.settingChildHeight.Width = 600
                        $syncHash.CompPropGrid.HorizontalAlignment = "Stretch"
                    }           
                })  
        }
    })

$syncHash.settingObjectToolsMapClick.Add_Click( {

        Start-RSJob -ArgumentList $syncHash, $configHash -ScriptBlock {
            Param($syncHash, $configHash)

            $syncHash.Window.Dispatcher.invoke([action] {
        
                    if ($configHash.objectToolConfig -ne $null) {
                        $syncHash.settingObjectToolsPropGrid.ItemsSource = $null
                        $syncHash.settingObjectToolsPropGrid.ItemsSource = $configHash.objectToolConfig
                    }

                    if ($syncHash.settingObjectToolsPropGrid.Visibility -eq "Visible") {
                        $syncHash.settingObjectToolsPropGrid.Visibility = "Hidden"          
                        $syncHash.settingObjectToolsAddItemClick.Visibility = "Hidden"
                        $syncHash.settingChildHeight.Height = 225     
                        $syncHash.settingChildHeight.Width = 390
                        $syncHash.objectToolsPropGrid.HorizontalAlignment = "Left"
                               
                    }
            
                    else {
                        $syncHash.settingObjectToolsPropGrid.Visibility = "Visible"
                        $syncHash.settingObjectToolsAddItemClick.Visibility = "Visible"
                        $syncHash.settingChildHeight.Height = 365 #330
                        $syncHash.settingChildHeight.Width = 600
                        $syncHash.objectToolsPropGrid.HorizontalAlignment = "Stretch"
                    }           
                })  
        }
    })

$syncHash.settingContextPropMapClick.Add_Click( {

        Start-RSJob -ArgumentList $syncHash, $configHash -ScriptBlock {
            Param($syncHash, $configHash)

            $syncHash.Window.Dispatcher.invoke([action] {
        
                    if ($configHash.contextConfig -ne $null) {
                        $syncHash.settingContextPropGrid.ItemsSource = $null
                        $syncHash.settingContextPropGrid.ItemsSource = $configHash.contextConfig
                    }

                    if ($syncHash.settingContextPropGrid.Visibility -eq "Visible") {
                        $syncHash.settingContextPropGrid.Visibility = "Hidden"          
                        $syncHash.settingContextPropImportClick.Visibility = "Hidden"
                        $syncHash.settingContextAddItemClick.Visibility = "Hidden"
                        $syncHash.settingChildHeight.Height = 225     
                        $syncHash.settingChildHeight.Width = 390
                        $syncHash.contextPropGrid.HorizontalAlignment = "Left"
                               
                    }
            
                    else {
                        $syncHash.settingContextPropGrid.Visibility = "Visible"
                        $syncHash.settingContextPropImportClick.Visibility = "Visible"
                        $syncHash.settingContextAddItemClick.Visibility = "Visible"
                        $syncHash.settingChildHeight.Height = 365 #330
                        $syncHash.settingChildHeight.Width = 390
                        $syncHash.contextPropGrid.HorizontalAlignment = "Stretch"
                    }           
                })  
        }
    })
#$syncHash.settingUserPropSaveClick.Add_Click({

#   $int = 0
#   $configHash.UserPropListSelection = @()
#   $configHash.UserPropList = [System.Collections.ArrayList]@()
   
    
#   $syncHash.settingUserPropGrid.Items | ForEach-Object {

#       $int++

#       if ($null -notlike $_.fieldName -and $null -notlike $_.PropName) {
#           $configHash.UserPropList.Add([PSCustomObject]@{Field = $int; FieldName = $_.FieldName; PropName = $_.PropName; ActionName = $_.ActionName})
#       }
#   }
# })
        
$syncHash.settingLoggingPcPathClick.Add_Click( {
 
        $pcLogSelection = New-FolderSelection -Title "Select client logging directory"

        if (![string]::IsNullOrEmpty($pcLogSelection)) {

            if (Test-Path $pcLogSelection) {
                $configHash.pcLogPath = $pcLogSelection
                $configHash.pcLogInUse = $true
                $syncHash.compLogPopupButton.IsEnabled = $true
                $syncHash.settingLoggingPC.Content = $segChar.segCheck
                $syncHash.settingLoggingPC.Foreground = "Green"
                $syncHash.settingLoggingPC.Tooltip = "Client logging path found"
            }

            else {
                $syncHash.compLogPopupButton.IsEnabled = $false
            }
        }
    })

$syncHash.settingLoggingUserPathClick.Add_Click( {
 
        $UserLogSelection = New-FolderSelection -Title "Select user logging directory"

        if (![string]::IsNullOrEmpty($UserLogSelection)) {

            if (Test-Path $UserLogSelection) {
                $configHash.UserLogPath = $UserLogSelection
                $configHash.UserLogInUse = $true
                $syncHash.settingLoggingUser.Content = $segChar.segCheck
                $syncHash.settingLoggingUser.Foreground = "Green"
                $syncHash.userLogPopupButton.IsEnabled = $true
                $syncHash.settingLoggingUser.Tooltip = "Client logging path found"
            
           

            }

            else {
                $syncHash.userLogPopupButton.IsEnabled = $false
            }
        }
    })

 
    
#endregion

$syncHash.settingChildWindow.add_ClosingFinished( {
    Reset-ChildWindow -SyncHash $syncHash    
})

#region Category button events
$syncHash.settingModClick.add_Click( {
        $syncHash.settingModContent.Visibility = "Visible"
        $syncHash.settingChildWindow.Title = "Required PS Modules"
        $syncHash.settingChildWindow.IsOpen = $true
    })

$syncHash.settingToolClick.add_Click( {
        $syncHash.settingToolContent.Visibility = "Visible"
        $syncHash.settingChildWindow.Title = "Required Tools"
        $syncHash.settingChildWindow.IsOpen = $true
    })

$syncHash.settingADClick.add_Click( {
        $syncHash.settingADContent.Visibility = "Visible"
        $syncHash.settingChildWindow.Title = "ADDS Config"
        $syncHash.settingChildWindow.IsOpen = $true
    })

$syncHash.settingPermClick.add_Click( {
        $syncHash.settingAdminContent.Visibility = "Visible"
        $syncHash.settingChildWindow.Title = "Admin Permissions"
        $syncHash.settingChildWindow.IsOpen = $true
    })

#endregion Category button events
#region Action pane exit events

$syncHash.settingCloseClick.add_Click( {
        $syncHash.Window.Close()
    })
    
$syncHash.settingConfigCancelClick.add_Click( {
        $syncHash.Window.Close()
    })

$syncHash.settingImportClick.add_Click( {

        $configSelection = New-Object System.Windows.Forms.OpenFileDialog
        $configSelection.initialDirectory = [Environment]::GetFolderPath('MyComputer')
        $configSelection.filter = "config|config.json"
        $configSelection.Title = "Select config.json"
        $configSelection.ShowDialog() | Out-Null

        if (![string]::IsNullOrEmpty($configSelection.fileName)) {
            Copy-Item -Path $configSelection.fileName -Destination $PSScriptRoot -Force
            $syncHash.Window.Close()
            Start-Process -WindowStyle Minimized -FilePath "$PSHOME\powershell.exe" -ArgumentList " -ExecutionPolicy Bypass -NonInteractive -File $($PSCommandPath)"
            exit
       
        }


    }) 

$syncHash.settingConfigClick.add_Click( {
       
       Set-Config -ConfigPath $savedConfig -Type Export -ConfigHash $configHash


        $syncHash.Window.Close()
        Start-Process -WindowStyle Hidden -FilePath "$PSHOME\powershell.exe" -ArgumentList " -ExecutionPolicy Bypass -NonInteractive -File $($PSCommandPath)"
        exit
    })

#endregion

###### MAIN WINDOW
$syncHash.TabMenu.add_SelectionChanged( {
        $syncHash.itemHeader.Content = $syncHash.tabMenu.SelectedItem.Tag
        $syncHash.TabMenu.Items | ForEach-Object { $_.Background = "#FF444444" }
        $syncHash.tabMenu.SelectedItem.Background = "#576573"
        $syncHash.TabMenu.Items.Header | ForEach-Object { $_.Foreground = "Gray" }
        $syncHash.tabMenu.SelectedItem.Header.Foreground = "AliceBlue"

        if ($syncHash.tabMenu.SelectedIndex -eq 4) {         
           $syncHash.consoleControl.StartProcess(("$PSHOME\powershell.exe"))            
        }

        elseif ($syncHash.consoleControl.IsProcessRunning) {
            $syncHash.consoleControl.StopProcess()
        }


    })

$syncHash.searchBoxHelp.add_MouseLeftButtonUp( {
        $syncHash.childHelp.isOpen = $true       
    })

$syncHash.searchBoxHelp.add_MouseLeftButtonDown( {
        $syncHash.searchBoxHelp.Background = "#576573"
        $syncHash.searchBoxHelp.Foreground = "Gray"
    })

$syncHash.searchBoxHelp.add_MouseEnter( {
        $syncHash.searchBoxHelp.Foreground = "#576573"
        $syncHash.searchBoxHelp.Background = "#454545"
    })

$syncHash.searchBoxHelp.add_MouseLeave( {
        $syncHash.searchBoxHelp.Foreground = "LightGray"
        $syncHash.searchBoxHelp.Background = "Transparent"
    })


$syncHash.tabControl.ItemsSource = New-Object System.Collections.ObjectModel.ObservableCollection[Object]

$syncHash.tabMenu.add_Loaded( {
        # Need an delay here, otherwise the default tabindex cannot be selected
        if ($sysCheckHash.sysChecks.RSModule -eq $true) {

            Start-RSJob -Name menuLoad -ArgumentList $syncHash, $savedConfig, $sysCheckHash -ModulesToImport C:\TempData\internal\internal.psm1 -ScriptBlock {        
                Param($syncHash, $savedConfig, $sysCheckHash)

                do {} until ($sysCheckHash.checkComplete)
           

                if ($sysCheckHash.missingCheckFail -or $sysCheckHash.adCheckFail -or $sysCheckHash.adminCheckFail) {
                    Suspend-FailedItems -SyncHash $syncHash -CheckedItems SysCheck  
                }

                elseif (!(Test-Path $savedConfig)) {
                    Suspend-FailedItems -SyncHash $syncHash -CheckedItems Config
                }
        
                else {  
                    $syncHash.Window.Dispatcher.invoke([action] { $syncHash.tabMenu.SelectedIndex = 0 })  
                }
            }

            #userbox structure populate
   
            # if (!($configHash.($type + 'PropList')Selection)) {
            #     $configHash.($type + 'PropList')Selection = @()
            #     Remove-Variable skipAdd -ErrorAction SilentlyContinue
            # }

        }

        if (Test-Path $savedConfig) {
               
            foreach ($type in @('User', 'Comp')) {  
            
                if (!($configHash.($type + 'PropListSelection'))) {
                    $configHash.($type + 'PropListSelection') = @()
                }

                for ($i = 1; $i -le $configHash.boxMax; $i++) {
            
                    if ($i -le $configHash.($type + 'boxCount')) {
                    
                        Remove-Variable Selected -ErrorAction SilentlyContinue
      
                        $selected = $configHash.($type + 'PropList') | Where-Object { $_.Field -eq $i }
       
                        if ($null -ne $selected.FieldName) { 
                
                            $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Header').Content = $selected.FieldName
                            $configHash.($type + 'PropListSelection') += $selected.PropName

                            if ($selected.ActionName -notlike "*editable*") {
                                $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'TextBox').Tag = 'NoEdit'
                                $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'EditClip').Visibility = "Collapsed"
                                $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1').Visibility = "Collapsed"                    
                            }

                            else {
                                # define edit buttons
                                $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1').Add_Click( {
                                        param([Parameter(Mandatory)][Object]$sender)
                       
                                        if ($sender.Name -like "*ubox*") {
                                            $id = $sender.Name -replace "ubox|box.*", "" 
                                            $type = "User"
                                        }

                                        else {
                                            $id = $sender.Name -replace "cbox|box.*", "" 
                                            $type = "Comp"
                                        }

                                        if ($configHash.adPropertyMap) {
                                            Remove-Variable ChangedValue, propType, ldapValue, user -ErrorAction SilentlyContinue
                                            $changedValue = $syncHash.($type[0] + 'box' + $id + 'resources').($type[0] + 'box' + $id + 'TextBox').Text
                                            $propType = $configHash.($type + 'PropList')[$id - 1].PropType
                            
                                            if ($changedValue -as $propType) {    
                                                $changedValue = $changedValue -as $propType
                                
                                                try {
                                            
                                                    if ($type -eq 'User') {
                                                        $editObject = ($queryHash[$syncHash.tabControl.SelectedItem.Name]).SamAccountName
                                                        $ldapValue = $configHash.adPropertyMap[(($configHash.($type + 'PropList')[$id - 1].PropName))]
                                                        Set-ADUser -Identity $editObject -Replace @{$ldapValue = $changedValue }
                                                        
                                                    
                                                    }

                                                    else {
                                                        $editObject = ($queryHash[$syncHash.tabControl.SelectedItem.Name]).Name
                                                        $ldapValue = $configHash.adPropertyMap[(($configHash.($type + 'PropList')[$id - 1].PropName))]
                                                        Set-ADComputer -Identity $editObject -Replace @{$ldapValue = $changedValue }
                                                        
                                                    }

                                                    $queryHash.($editObject).($configHash.($type + 'PropList')[$id - 1].PropName) = $changedValue
                                                    $syncHash.SnackMsg.MessageQueue.Enqueue("[EDIT]: SUCCESS on [$($($editObject).toLower())]")
                                                }
                                    
                                                catch {
                                                    $syncHash.SnackMsg.MessageQueue.Enqueue("[EDIT]: FAILED on [$($($editObject).toLower())]")
                                                }
                                            }

                                            else {
                                                if ([string]::IsNullOrEmpty($changedValue)) {
                                                    $ldapValue = $configHash.adPropertyMap[(($configHash.($type + 'PropList')[$id - 1].PropName))]
                                                    
                                                    try {
                                                        if ($type -eq 'User') {
                                                            $editObject = ($queryHash[$syncHash.tabControl.SelectedItem.Name]).SamAccountName
                                                            Set-ADUser -Identity $editObject -Clear $ldapValue
                                                        }
                                                        else {
                                                            $editObject = ($queryHash[$syncHash.tabControl.SelectedItem.Name]).Name
                                                            Set-ADComputer -Identity $editObject -Clear $ldapValue
                                                        }

                                                        $syncHash.SnackMsg.MessageQueue.Enqueue("[CLEAR]: SUCCESS on [$($($editObject).toLower())]")
                                                    }

                                                    catch {
                                                        $syncHash.SnackMsg.MessageQueue.Enqueue("[CLEAR]: FAIL on [$($($editObject).toLower())]")
                                                    }
                                                }

                                                else {     
                                                    $syncHash.SnackMsg.MessageQueue.Enqueue("[EDIT]: FAILED on [$($($editObject).toLower())] - expected type  [$($($propType).toUpper())]")
                                                }
                                            }
                                        }
                        
                                        else {
                                            $syncHash.SnackMsg.MessageQueue.Enqueue("[EDIT]: FAILED on [$($($editObject).toLower())] - ADEntity .dll missing; cannot edit")
                                        }
                                    })

                            }

                            if ($selected.ActionName -notmatch "Actionable") {
                                $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1Action1').Visibility = "Collapsed"  
                                $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1Action2').Visibility = "Collapsed"  
                            }

                            else {
                                for ($b = 1; $b -le 2; $b++) {
 
                                    # Event assigment for panels in user expander
                                    if ($configHash.($type + 'PropList')[$i - 1].('validAction' + $b)) {
                                        $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1Action' + $b).Content = $configHash.($type + 'PropList')[$i - 1].('actionCmd' + $b + 'Icon')
                                        $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1Action' + $b).ToolTip = $configHash.($type + 'PropList')[$i - 1].('actionCmd' + $b + 'ToolTip')
                         
                                        if (($configHash.($type + 'PropList')[$i - 1]).('actionCmd' + $b + 'Multi') -eq $true) {                       
                                            $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + ('Box1Action' + $b)).Add_Click( {
                            
                                                    param([Parameter(Mandatory)][Object]$sender)
                                    
                                                    if ($sender.Name -like "*ubox*") {
                                                        $id = $sender.Name -replace "ubox|Box.*" 
                                                        $type = "User"
                                                    }

                                                    else {
                                                        $id = $sender.Name -replace "cbox|Box.*" 
                                                        $type = "Comp"
                                                    }

                                                    $b = $sender.Name -replace ".*action"
                                                    $rsCmd = @{
                                                        id           = $id
                                                        type         = $type
                                                        propList     = $configHash.($type + 'PropList')[$id - 1]
                                                        actionObject = if ($type -eq 'User') { ($queryHash[$syncHash.tabControl.SelectedItem.Name]).SamAccountName }
                                                                       else { ($queryHash[$syncHash.tabControl.SelectedItem.Name]).Name }
                                                        boxResources = $syncHash.($type[0] + 'box' + $id + 'resources')
                                                        window       = $syncHash.Window
                                                        snackMsg     = $syncHash.SnackMsg
                                                    }
    
                               
                                                    Start-RSJob -Name threadedAction -ArgumentList $rsCmd, $queryHash, $b -ScriptBlock {
                                                        Param($rsCmd, $queryHash, $b)
                                                        
                                                        
                                                        Start-Sleep -Milliseconds 500                   
                                                        $actionNameString = "[$($($rsCmd.propList.('actionCmd' + $b + 'ToolTip')).toUpper())]"
                                                        $actionObjectString = "[$($($rsCmd.actionObject).toLower())]"
                                                        $cmd = $rsCmd.propList.('actionCmd' + $b)
                                                        $type = $rsCmd.Type
                                                        $propName = $rsCmd.propList.PropName
                                            
                                                        New-Variable -Name $type -Value $rsCmd.actionObject
                                                        $prop = $queryHash.(Get-Variable -Name $type -ValueOnly).($propName)

                                                        try {
                                                            Invoke-Expression -Command $cmd
                                                            $rsCmd.Window.Dispatcher.Invoke([Action] { $rsCmd.snackMsg.MessageQueue.Enqueue("$actionNameString SUCCESS on $actionObjectString") })
              
                                                            if ($rsCmd.propList.('actionCmd' + $b + 'Refresh')) {
                                                                if ($PropName -ne 'Non-Ad Property') {
                                                                    if ($type -eq 'User') {                               
                                                                        $result = (Get-ADUser -Identity $user -Properties $propName).$propName
                                                                        $queryHash.($user).($propName) = $result

                                                                        
                                                                        #here##
                                                                    }
                                                                    else {
                                                                        $result = (Get-AdComputerName -Identity $comp -Properties $propName).$propName
                                                                        $queryHash.($comp).($propName) = $result
                                                                    }
                                                                }
                                                        
                                                                if ($queryHash.(Get-Variable -Name $type -ValueOnly).ActiveItem -eq $true) {  
                                                                    if ($rsCmd.propList.ValidCmd) {
                                                                        $tranCmd = $rsCmd.propList.translationCmd
                                                                        $value = Invoke-Expression -Command $tranCmd
                                                       
                                                                        if ($resultColor) {
                                                                            $rsCmd.Window.Dispatcher.Invoke([Action] { $rsCmd.boxResources.($type[0] + 'box' + $rsCmd.id + "TextBox").Foreground = $resultColor })
                                                                        }

                                                                        $updatedValue = $value

                                                                    }

                                                                    else {
                                                                        $updatedValue = $result
                                                                    }

                                                                    $rsCmd.Window.Dispatcher.Invoke([Action] { $rsCmd.boxResources.($type[0] + 'box' + $rsCmd.id + "TextBox").Text = $updatedValue })
                                        
                                                                }

                                                            }
                                                        }
                          
                                                        catch {
                                                            $rsCmd.Window.Dispatcher.Invoke([Action] { $rsCmd.snackMsg.MessageQueue.Enqueue("$actionNameString FAILURE on $actionObjectString") })
                                                        }
                                                    }
                                                })                         
                                        }
                          
                                        else {
                                            $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + ('Box1Action' + $b)).Add_Click( {
                                                    param([Parameter(Mandatory)][Object]$sender)

                                    
                                                    if ($sender.Name -like "*ubox*") {
                                                        $id = $sender.Name -replace "ubox|Box.*" 
                                                        $type = "User"
                                                        $actionObject = ($queryHash[$syncHash.tabControl.SelectedItem.Name]).SamAccountName
                                                    }

                                                    else {
                                                        $id = $sender.Name -replace "cbox|Box.*" 
                                                        $type = "Comp"
                                                        $actionObject = ($queryHash[$syncHash.tabControl.SelectedItem.Name]).Name
                                                    }
                                   
                                                    $b = $sender.Name -replace ".*action"

                                                    Remove-Variable -Name $type -ErrorAction SilentlyContinue
                                                    New-Variable -Name $type -Value $actionObject
                                                    $propName = $configHash.($type + 'PropList')[$id - 1].PropName
                                                    $prop = $queryHash.(Get-Variable -Name $type -ValueOnly).($propName)
                                                    $actionName = $configHash.($type + 'PropList')[$id - 1].('actionCmd' + $b + 'ToolTip')

                                                    try {
                                                        Invoke-Expression -Command ($configHash.($type + 'PropList')[$id - 1].('actionCmd' + $b))
                                                        $syncHash.SnackMsg.MessageQueue.Enqueue("[$($($actionName).toUpper())]: SUCCESS on [$($(Get-Variable -Name $type -ValueOnly).toLower())]")

                                                        if ($configHash.($type + 'PropList')[$id - 1].('actionCmd' + $b + 'Refresh')) {
                                                            if (($configHash.($type + 'PropList')[$id - 1].PropName) -ne 'Non-Ad Property') {
                                                    
                                                                if ($type -eq 'User') {                               
                                                                    $result = (Get-ADUser -Identity $user -Properties ($configHash.($type + 'PropList')[$id - 1].PropName)).($configHash.($type + 'PropList')[$id - 1].PropName)
                                                                    $queryHash.($user).($propName) = $result
                                                                }
                                                                else {
                                                                    $result = (Get-ADComputer -Identity $comp -Properties ($configHash.($type + 'PropList')[$id - 1].PropName)).($configHash.($type + 'PropList')[$id - 1].PropName)
                                                                    $queryHash.($comp).($propName) = $result
                                                                }

                                                                $queryHash.(Get-Variable -Name $type -ValueOnly).($propName) = $result

                                                            }
                                                   
                                                    
                                                            if (($configHash.($type + 'PropList') | Where-Object { $_.Field -eq $id }).ValidCmd -and ($configHash.($type + 'PropList') | Where-Object { $_.Field -eq $id }).transCmdsEnabled) {
                                                                Remove-Variable resultColor -ErrorAction SilentlyContinue
                                                                $value = Invoke-Expression -Command (($configHash.($type + 'PropList') | Where-Object { $_.Field -eq $id }).TranslationCmd)
                                        
                                                                if ($resultColor) {
                                                                    $syncHash.($type[0] + 'box' + $id + "resources").($type[0] + 'box' + $id + "TextBox").Foreground = $resultColor
                                                                }

                                                                $updatedValue = $value

                                                            }

                                                            else {
                                                                $updatedValue = $result
                                                            }

                                                            $syncHash.($type[0] + 'box' + $id + "resources").($type[0] + 'box' + $id + "TextBox").Text = $updatedValue
                                                        }
                        
                                                    }
                          
                                                    catch {
                                                        $syncHash.SnackMsg.MessageQueue.Enqueue("[$($($actionName).toUpper())] FAILURE on [$($(Get-Variable -Name $type -ValueOnly).toLower())]")
                                                    }
                                                                
                                                })
                                        }


                                    }
                    
                                    else {
                                        $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + ('Box1Action' + $b)).Visibility = "Collapsed"  
                                    }
                                }
                                #  if ($configHash.($type + 'PropList')[$i - 1].validAction2) {
                                #      $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1Action2').Content = $configHash.($type + 'PropList')[$i - 1].actionCmd2Icon 
                                #  }

                                #  else {
                                #      $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1Action2').Visibility = "Collapsed" 
                                #  }
                            }
            
                        }

                        else {
                            $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Header').Visibility = "Collapsed"
                            $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Border').Visibility = "Collapsed"
                        }
            
                    
            
            
            
                    }
                }
        
                if ($configHash.($type + 'boxCount') -le 7) {
                    $syncHash.Window.Dispatcher.Invoke([Action] {
                            $syncHash.($type + 'detailGrid').Columns = 1
                        })
                }

                elseif ($configHash.($type + 'boxCount') -gt 7 -and $configHash.($type + 'boxCount') -le 14) {
                    $syncHash.Window.Dispatcher.Invoke([Action] {
                            $syncHash.($type + 'detailGrid').Columns = 2
                        })
                }

                else {
                    $syncHash.Window.Dispatcher.Invoke([Action] {
                            $syncHash.($type + 'detailGrid').Columns = 3
                        })
                }

               

                

            }
            # populate compgrid connection button icons           
            foreach ($buttonType in @('rbut', 'rcbut')) {
                $syncHash.($buttonType + 1).Source = ([Convert]::FromBase64String($configHash.rtConfig.MSTSC.Icon))
                $syncHash.($buttonType + 1).Parent.Tag = "MSTSC"
                $syncHash.($buttonType + 1).Parent.ToolTip = $configHash.rtConfig.MSTSC.DisplayName
                $syncHash.($buttonType + 1).Parent.Add_Click( {
                        param([Parameter(Mandatory)][Object]$sender) 
                        if ($sender.Name -match "rbut") {
                            if ($syncHash.userCompFocusHostToggle.IsChecked) {
                                $comp = $syncHash.UserCompGrid.SelectedItem.HostName
                            }
                            else {
                                $comp = $syncHash.UserCompGrid.SelectedItem.ClientName
                            }

                            $user = $syncHash.UserCompGrid.SelectedItem.UserName
                
                        }

                        else {
                
                            if ($syncHash.compUserFocusUserToggle.IsChecked) {
                                $comp = $synchash.tabControl.SelectedItem.Name
                            }
                
                            else {
                                $comp = $syncHash.CompUserGrid.SelectedItem.ClientName
                            }

                            $user = $syncHash.CompUserGrid.SelectedItem.UserName

                        }
            
                        mstsc /v $comp /admin
                    })

                $syncHash.($buttonType + 2).Source = ([Convert]::FromBase64String($configHash.rtConfig.MSRA.Icon))
                $syncHash.($buttonType + 2).Parent.Tag = "MSRA"
                $syncHash.($buttonType + 2).Parent.ToolTip = $configHash.rtConfig.MSRA.DisplayName
                $syncHash.($buttonType + 2).Parent.Add_Click( { 
                        param([Parameter(Mandatory)][Object]$sender)
                        if ($sender.Name -match "rbut") {
                            if ($syncHash.userCompFocusHostToggle.IsChecked) {
                                $comp = $syncHash.UserCompGrid.SelectedItem.HostName
                            }
                            else {
                                $comp = $syncHash.UserCompGrid.SelectedItem.ClientName
                            }

                            $user = $syncHash.UserCompGrid.SelectedItem.UserName
                
                        }

                        else {
                
                            if ($syncHash.compUserFocusUserToggle.IsChecked) {
                                $comp = $synchash.tabControl.SelectedItem.Name
                            }
                
                            else {
                                $comp = $syncHash.CompUserGrid.SelectedItem.ClientName
                            }

                            $user = $syncHash.CompUserGrid.SelectedItem.UserName

                        }
                        msra /expert /offera $comp 
                    })    


                foreach ($rtID in $configHash.rtConfig.Keys.Where{$_ -like "RT*"}) {
                                 

                    $syncHash.customRT.$rtID.$buttonType = New-Object System.Windows.Controls.Button -Property @{
                        Padding = "0"
                        Tag     = $rtID
                        ToolTip = $configHash.rtConfig.$rtID.DisplayName
                        Style   = $syncHash.Window.FindResource('itemButton')
                        Name    = $buttonType + ([int]($rtID -replace 'rt') + 2)
                    }
            
                    $syncHash.customRT.$rtID.($buttonType + 'img') = New-Object System.Windows.Controls.Image -Property @{
                        Width  = 15
                        Height = 15
                        Source = ([Convert]::FromBase64String($configHash.rtConfig.$rtID.Icon))
                    }

                    $syncHash.customRT.$rtID.$buttonType.AddChild($syncHash.customRT.$rtID.($buttonType + 'img'))

                    $syncHash.customRT.$rtID.($buttonType + 'img').Parent.Add_Click( { 
                            param([Parameter(Mandatory)][Object]$sender) 
                            $id = 'rt' + ([int]($sender.Name -replace ".*but") - 2)              

                            if ($sender.Name -match "rbut") {
                
                                if ($syncHash.userCompFocusHostToggle.IsChecked) {
                                    $comp = $syncHash.UserCompGrid.SelectedItem.HostName
                                }
                                else {
                                    $comp = $syncHash.UserCompGrid.SelectedItem.ClientName
                                }

                                $user = $syncHash.UserCompGrid.SelectedItem.UserName
                
                            }

                            else {
                
                                if ($syncHash.compUserFocusUserToggle.IsChecked) {
                                    $comp = $synchash.tabControl.SelectedItem.Name
                                }
                
                                else {
                                    $comp = $syncHash.CompUserGrid.SelectedItem.ClientName
                                }

                                $user = $syncHash.CompUserGrid.SelectedItem.UserName

                            }

                            Invoke-Expression -Command  ($configHash.rtConfig.$id.cmd)
                        })

                    $syncHash.($buttonType + 'Grid').AddChild($syncHash.customRT.$rtID.$buttonType)
                }

                #here#
                foreach ($contextBut in $configHash.contextConfig.Where{$_.ValidAction -eq $true}) {
                   
                    if ($null -eq $syncHash.customContext.('cxt' + $contextBut.IDNum)) {
                        $syncHash.customContext.('cxt' + $contextBut.IDNum) = @{}
                    }
                               
                    $syncHash.customContext.('cxt' + $contextBut.IDNum).($buttonType + 'context' + $contextBut.IDNum) = New-Object System.Windows.Controls.Button -Property @{
                        Padding = "0"
                        Tag     = $buttonType + 'context' + $contextBut.IDNum
                        ToolTip = $configHash.contextConfig[$contextBut.IDnum - 1].ActionName
                        Style   = $syncHash.Window.FindResource('itemButton')
                        Name    = $buttonType + 'context' + $contextBut.IDNum
                        FontFamily = "Segoe MDL2 Assets"
                        Content = $configHash.contextConfig[$contextBut.IDnum - 1].actionCmdIcon
                    }
            

                    $syncHash.customContext.('cxt' + $contextBut.IDNum).($buttonType + 'context' + $contextBut.IDNum).Add_Click( { 
                        param([Parameter(Mandatory)][Object]$sender) 
                        $id = [int]($sender.Name -replace ".*context")
                        if ($sender.Parent.Name -eq 'rbutContextGrid') {
                            
                            $user = $configHash.currentTabItem
                            $sessionID = $syncHash.UserCompGrid.SelectedItem.SessionID

                            if ($syncHash.userCompFocusHostToggle.IsChecked) {
                                $comp = $syncHash.UserCompGrid.SelectedItem.HostName
                            }
                            else {
                                $comp = $syncHash.UserCompGrid.SelectedItem.ClientName
                            }
                        }
                        
                        else {

                            $user = $syncHash.CompUserGrid.SelectedItem.UserName
                            $sessionID = $syncHash.CompUserGrid.SelectedItem.SessionID

                            if ($syncHash.compUserFocusUserToggle.IsChecked) {
                                $comp = $configHash.currentTabItem
                            }
                
                            else {
                                $comp = $syncHash.CompUserGrid.SelectedItem.ClientName
                            }     
                        }
                        
                        


                        if ($configHash.contextConfig[$id - 1].actionCmdMulti) {

                            $rsCmd = @{
                                comp           = $comp
                                user           = $user
                                buttonSettings = $configHash.contextConfig[$id - 1]
                                sessionID      = $sessionID
                            }

                            Start-RSJob -ArgumentList $rsCmd, $syncHash -ScriptBlock {
                                Param($rsCmd, $syncHash)
                            
                                try {
                                    Invoke-Expression -Command $rsCmd.buttonSettings.actionCmd
                                    $syncHash.Window.Dispatcher.Invoke([Action]{
                                        $syncHash.SnackMsg.MessageQueue.Enqueue("[$($rsCmd.buttonSettings.ActionName.ToUpper())]: SUCCESS on $($rsCmd.user.ToLower()) on $($rsCmd.comp.ToLower())")
                                    })
                                }
                                catch {
                                    $syncHash.Window.Dispatcher.Invoke([Action]{
                                        $syncHash.SnackMsg.MessageQueue.Enqueue("[$($rsCmd.buttonSettings.ActionName.ToUpper())]: FAIL on $($rsCmd.user.ToLower()) on $($rsCmd.comp.ToLower())")  
                                    })
                                }
                            }
                        }

                        else {

                            try {
                                Invoke-Expression -Command $configHash.contextConfig[$id - 1].actionCmd
                                $syncHash.SnackMsg.MessageQueue.Enqueue("[$($configHash.contextConfig[$id - 1].ActionName.ToUpper())]: SUCCESS on $($user.ToLower()) on $($comp.ToLower())")
                            }
                            
                            catch {
                               $syncHash.SnackMsg.MessageQueue.Enqueue("[$($configHash.contextConfig[$id - 1].ActionName.ToUpper())]: FAIL on $($user.ToLower()) on $($comp.ToLower())")
                            }
                            

                        }

                    })
                
                    $syncHash.($buttonType + 'ContextGrid').AddChild($syncHash.customContext.('cxt' + $contextBut.IDNum).($buttonType + 'context' + $contextBut.IDNum))


                   
                }




            }

            $customField = 0
            $configHash.userLogMapping | Where-Object { $_.FieldSel -eq 'Custom' -and $_.Ignore -eq $false } | ForEach-Object {
                # populate custom dock items
                $customField++
                $syncHash.Window.Dispatcher.Invoke([Action] {

                        $syncHash.('customPropDock' + $customField) = New-Object System.Windows.Controls.StackPanel
                        $syncHash.('customPropLabel' + $customField) = New-Object System.Windows.Controls.Label
                        $syncHash.('customPropText' + $customField) = New-Object System.Windows.Controls.Textbox
                        $syncHash.('customPropLabel' + $customField).Content = $_.Header
                        $syncHash.('customPropDock' + $customField).VerticalAlignment = "Top"
                        $syncHash.('customPropDock' + $customField).Margin = "0,-10,0,0"
                

                        $syncHash.('customPropDock' + $customField).AddChild(($syncHash.('customPropLabel' + $customField)))
                        $syncHash.('customPropDock' + $customField).AddChild(($syncHash.('customPropText' + $customField)))
     

                        $syncHash.('customPropLabel' + $customField).FontSize = "10"
                        $syncHash.('customPropLabel' + $customField).Foreground = $syncHash.Window.FindResource('MahApps.Brushes.SystemControlBackgroundBaseMediumLow')
                        $syncHash.('customPropText' + $customField).Style = $syncHash.Window.FindResource('compItemBox') 
                        $syncHash.userLogExtraPropGrid.AddChild(($syncHash.('customPropDock' + $customField)))

                        # Create and set a binding on the textbox object
                        $Binding = New-Object System.Windows.Data.Binding
                        $Binding.UpdateSourceTrigger = "PropertyChanged"
                        $binding.Source = $syncHash.userCompGrid
                        $binding.Path = "SelectedItem.$($_.Header)"
                        $Binding.Mode = [System.Windows.Data.BindingMode]::TwoWay
     

                        [void][System.Windows.Data.BindingOperations]::SetBinding(($syncHash.('customPropText' + $customField)), [System.Windows.Controls.TextBox]::TextProperty, $Binding)
                    })
            }

            $customField = 0        
            $configHash.compLogMapping | Where-Object { $_.FieldSel -eq 'Custom' -and $_.Ignore -eq $false } | ForEach-Object {
                # populate custom dock items
                $customField++
                $syncHash.Window.Dispatcher.Invoke([Action] {

                        $syncHash.('customCompPropDock' + $customField) = New-Object System.Windows.Controls.StackPanel
                        $syncHash.('customCompPropLabel' + $customField) = New-Object System.Windows.Controls.Label
                        $syncHash.('customCompPropText' + $customField) = New-Object System.Windows.Controls.Textbox
                        $syncHash.('customCompPropLabel' + $customField).Content = $_.Header
                        $syncHash.('customCompPropDock' + $customField).VerticalAlignment = "Top"
                        $syncHash.('customCompPropDock' + $customField).Margin = "0,-10,0,0"

                        $syncHash.('customCompPropDock' + $customField).AddChild(($syncHash.('customCompPropLabel' + $customField)))
                        $syncHash.('customCompPropDock' + $customField).AddChild(($syncHash.('customCompPropText' + $customField)))
     

                        $syncHash.('customCompPropLabel' + $customField).FontSize = "10"
                        $syncHash.('customCompPropLabel' + $customField).Foreground = $syncHash.Window.FindResource('MahApps.Brushes.SystemControlBackgroundBaseMediumLow')
                        $syncHash.('customCompPropText' + $customField).Style = $syncHash.Window.FindResource('compItemBox') 
                        $syncHash.compLogExtraPropGrid.AddChild(($syncHash.('customCompPropDock' + $customField)))

                        # Create and set a binding on the textbox object
                        $Binding = New-Object System.Windows.Data.Binding
                        $Binding.UpdateSourceTrigger = "PropertyChanged"
                        $binding.Source = $syncHash.compUserGrid
                        $binding.Path = "SelectedItem.$($_.Header)"
                        $Binding.Mode = [System.Windows.Data.BindingMode]::TwoWay
     

                        [void][System.Windows.Data.BindingOperations]::SetBinding(($syncHash.('customCompPropText' + $customField)), [System.Windows.Controls.TextBox]::TextProperty, $Binding)
                    })
        
            }
        

      

        }

        else { 
            $syncHash.queryTab.IsEnabled = $false
            $syncHash.toolTab.IsEnabled = $false
            $syncHash.newTab.IsEnabled = $false
        }

    })
                
$syncHash.tabControl.add_SelectionChanged( {



    if ($configHash.itemRefreshing -eq $true) {
        $configHash.itemRefreshing = $false 
        $syncHash.expanderDisplay.Content = $null
        $syncHash.expanderTypeDisplay.Content = $null
        $syncHash.compExpanderTypeDisplay.Content = $null
        $syncHash.userExpander.IsExpanded = $false
        $syncHash.compExpander.IsExpanded = $false
        $syncHash.settingToolParent.Visibility = "Collapsed"
        $syncHash.userToolControlPanel.Visibility = "Collapsed"
        $syncHash.compToolControlPanel.Visibility = "Collapsed"
        $syncHash.tabControl.IsEnabled = $false
    }

    else {

        Get-RSJob -State Completed | Remove-RSJob
        $syncHash.tabControl.IsEnabled = $true
        $currentTabItem = $syncHash.tabControl.SelectedItem.Name
        $configHash.currentTabItem = $currentTabItem
    
        $queryHash.Keys | ForEach-Object { $queryHash[$_].ActiveItem = $false }
        $queryHash[$currentTabItem].ActiveItem = $true
        

    
        Start-RSJob -Name "VisualChange" -ThreadOptions UseNewThread -ArgumentList  $syncHash, $queryHash, $configHash, $currentTabItem -ScriptBlock {
            Param($syncHash, $queryHash, $configHash, $currentTabItem)
                $syncHash.Window.Dispatcher.Invoke([Action]{
                    $syncHash.expanderDisplay.Content = $null
                    $syncHash.expanderTypeDisplay.Content = $null
                    $syncHash.compExpanderTypeDisplay.Content = $null
                    $syncHash.userExpander.IsExpanded = $false
                    $syncHash.compExpander.IsExpanded = $false
                    $syncHash.settingToolParent.Visibility = "Collapsed"
                    $syncHash.userToolControlPanel.Visibility = "Collapsed"
                    $syncHash.compToolControlPanel.Visibility = "Collapsed"
                    $syncHash.compExpanderProgressBar.Visibility = "Visible"
                    $syncHash.expanderProgressBar.Visibility = "Visible"
  
                    if ($queryHash.($syncHash.tabControl.SelectedItem.Name).ObjectClass -eq 'Computer') {
                        if (($syncHash.compToolControlPanel.Children | Measure-Object).Count -eq 0) {
                            $syncHash.settingTools.Visibility = "Collapsed"
                        }
                    }
                    else {
                        if (($syncHash.userToolControlPanel.Children | Measure-Object).Count -eq 0) {
                            $syncHash.settingTools.Visibility = "Collapsed"
                        }
                    }
                })
               
  
        
       
    
            if ($null -ne $currentTabItem) {

                Start-RSJob -Name displayUpdate -ArgumentList $syncHash, $queryHash, $configHash, $currentTabItem -ScriptBlock {        
                    Param($syncHash, $queryHash, $configHash, $currentTabItem)                     
                    Start-Sleep -Milliseconds 500
                    $type = $queryHash[$currentTabItem].ObjectClass -replace 'Computer', 'Comp'

                    New-Variable -Name $type -Value $configHash.currentTabItem

                    for ($i = 1; $i -le $configHash.boxMax; $i++) {
        
                        if ($i -le $configHash.($type + 'boxCount')) {

                            Remove-Variable -Name resultColor -ErrorAction SilentlyContinue                 
            
                            if (($configHash.($type + 'propList') | Where-Object { $_.Field -eq $i }).ValidCmd -eq $true -and ($configHash.($type + 'propList') | Where-Object { $_.Field -eq $i }).transCmdsEnabled) {
                                $result = ($queryHash[$currentTabItem]).(($configHash.($type + 'propList') | Where-Object { $_.Field -eq $i }).PropName)
                                $value = Invoke-Expression -Command (($configHash.($type + 'propList') | Where-Object { $_.Field -eq $i }).TranslationCmd)
            
                                if ($resultColor) {
                                    $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.($type[0] + 'box' + $i + "resources").($type[0] + 'box' + $i + "TextBox").Foreground = $resultColor })
                
                                }
                            }
       
                            elseif (($configHash.($type + 'propList') | Where-Object { $_.Field -eq $i }).PropName) {
                                $value = ($queryHash[$currentTabItem]).(($configHash.($type + 'propList') | Where-Object { $_.Field -eq $i }).PropName)
                            }
       
                            if ($null -notlike ($configHash.($type + 'propList') | Where-Object { $_.Field -eq $i }).PropName) {

                                if ($null -notlike $value) {
                                    $syncHash.Window.Dispatcher.invoke([action] { $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Textbox').Text = $value })
                                }           
                            }
        
                            else {
                                break
                            }
        
                        }
                    }
    
                    if ($queryHash[$currentTabItem].ObjectClass -eq 'User') {  
                        $syncHash.Window.Dispatcher.invoke([action] {            
                                $syncHash.expanderTypeDisplay.Content = "USER   "
                                $syncHash.compExpanderTypeDisplay.Content = "COMPUTERS   "
                                $syncHash.compDetailMainPanel.Visibility = "Collapsed"
                                $syncHash.userDetailMainPanel.Visibility = "Visible"
                                $syncHash.expanderDisplay.Content = "$($queryHash[$currentTabItem].Name)"
                                $syncHash.userExpander.IsExpanded = $true
                                $syncHash.userCompGrid.Visibility = "Visible"
                                $syncHash.expanderProgressBar.Visibility = "Hidden"

                           
                                if (($syncHash.userToolControlPanel.Children | Measure-Object).Count -gt 0) {
                                     $syncHash.settingTools.Visibility = "Visible"
                                     $syncHash.settingToolParent.Visibility = "Visible"
                                     $syncHash.userToolControlPanel.Visibility = "Visible"
                                }
            
                            }) 

                    }  
    
                    elseif ($queryHash[$currentTabItem].ObjectClass -eq 'Computer') {  
                        $syncHash.Window.Dispatcher.invoke([action] {            
                                $syncHash.expanderTypeDisplay.Content = "COMPUTER   "
                                $syncHash.userDetailMainPanel.Visibility = "Collapsed"
                                $syncHash.compDetailMainPanel.Visibility = "Visible"
                                $syncHash.compExpanderTypeDisplay.Content = "USERS   "
                                $syncHash.expanderDisplay.Content = "$($queryHash[$currentTabItem].Name)"
                                $syncHash.userExpander.IsExpanded = $true
                                $synchash.userGrid.Visibility = "Visible" 
                                $syncHash.expanderProgressBar.Visibility = "Hidden"

                                if (($syncHash.compToolControlPanel.Children | Measure-Object).Count -gt 0) {
                                     $syncHash.settingTools.Visibility = "Visible"
                                     $syncHash.settingToolParent.Visibility = "Visible"
                                     $syncHash.compToolControlPanel.Visibility = "Visible"
                                }

                            })    
                    }      
                }    
  

                Start-RSJob -Name displayLogUpdate -ArgumentList $syncHash, $queryHash, $configHash, $currentTabItem  -ScriptBlock {        
                    Param($syncHash, $queryHash, $configHash, $currentTabItem)        
    
                    $type = $queryHash[$currentTabItem].ObjectClass -replace 'Computer', 'Comp'
               
                    do { } until ($queryHash[$currentTabItem].logsSearched -eq $true)

                    if ($null -ne ($queryHash[$currentTabItem]).LoginLogListView) {
                        if ($type -eq 'User') {
                            $syncHash.userCompGrid.Dispatcher.Invoke([Action] { $syncHash.userCompGrid.ItemsSource = ($queryHash[$currentTabItem]).LoginLogListView })
                            $syncHash.Window.Dispatcher.invoke([action] {
                                    $syncHash.userCompNotFound.Visibility = "Hidden"
                                    $syncHash.userCompDockPanel.Visibility = "Visible"
                                    $syncHash.compUserDockPanel.Visibility = "Collapsed"
                                    $syncHash.compExpander.IsExpanded = $true
                                    $syncHash.compExpanderProgressBar.Visibility = "Hidden"
                                    #$syncHash.userCompGrid.UpdateLayout()
                                })
                        }
                        else {
                            $syncHash.userCompGrid.Dispatcher.Invoke([Action] { $syncHash.compUserGrid.ItemsSource = ($queryHash[$currentTabItem]).LoginLogListView })
                            $syncHash.Window.Dispatcher.invoke([action] {
                                    $syncHash.userCompNotFound.Visibility = "Hidden"
                                    $syncHash.userCompDockPanel.Visibility = "Collapsed"
                                    $syncHash.compUserDockPanel.Visibility = "Visible"
                                    $syncHash.compExpander.IsExpanded = $true
                                    $syncHash.compExpanderProgressBar.Visibility = "Hidden"
                                    #$syncHash.compUserGrid.UpdateLayout()
                                })
                        }            
                    }

                    else  {
                        $syncHash.Window.Dispatcher.invoke([action] {
                                $syncHash.userCompNotFound.Visibility = "Visible"
                                $syncHash.compUserDockPanel.Visibility = "Collapsed"
                                $syncHash.userCompDockPanel.Visibility = "Collapsed"
                                $syncHash.compExpander.IsExpanded = $true
                                $syncHash.compExpanderProgressBar.Visibility = "Hidden"

                            })
                    }
                }
                #update valuebox
            }


            }

    }

    })
        
$syncHash.tabControl.add_tabItemClosingEvent( {


        $queryHash.Remove($configHash.currentTabItem)

        if ($syncHash.tabControl.Items.Count -le 1) {
            $syncHash.userExpander.IsExpanded = $false
            $syncHash.compExpander.IsExpanded = $false
            $syncHash.compExpanderProgressBar.Visibility = "Visible"
            $syncHash.expanderProgressBar.Visibility = "Visible"
            $synchash.userGrid.Visibility = "Collapsed"
            $syncHash.userCompGrid.Visibility = "Collapsed"
            
        }
       
    })

$syncHash.userCompGrid.Add_SelectionChanged( {

        if ($null -like $syncHash.userCompGrid.SelectedItem) {
            $syncHash.userCompControlPanel.IsEnabled = $false
        }

        else {
            $syncHash.userCompControlPanel.IsEnabled = $true
        }

        if ([string]::IsNullOrEmpty($syncHash.userCompGrid.SelectedItem.ClientName)) {          
            $syncHash.userCompFocusClientToggle.Visibility = "Hidden"
            $syncHash.userLogClientPropGrid.Visibility = "Hidden"

        }

        else {       
            $syncHash.userCompFocusClientToggle.Visibility = "Visible"
            $syncHash.userLogClientPropGrid.Visibility = "Visible"
        }

        if ($syncHash.userCompFocusClientToggle.IsChecked) {
            $syncHash.userCompFocusHostToggle.IsChecked = $true
        }

        else {

            foreach ($button in $syncHash.Keys.Where({$_ -like "*rbutbut*"})) {
                if (($syncHash.userCompGrid.SelectedItem.Type -in $configHash.rtConfig.($syncHash[$button].Tag).Types) -and
                   (!($configHash.rtConfig.($syncHash[$button].Tag).RequireOnline -and $syncHash.userCompGrid.SelectedItem.Connectivity -eq $false)) -and 
                   (!($configHash.rtConfig.($syncHash[$button].Tag).RequireUser -and $syncHash.userCompGrid.SelectedItem.userOnline -eq $false))) {                
                
                        $syncHash[$button].IsEnabled = $true
            
                }
                else {
                    $syncHash[$button].IsEnabled = $false
                }            
            }

            foreach ($button in $synchash.customRT.Keys) {
                if (($syncHash.userCompGrid.SelectedItem.Type -in $configHash.rtConfig.$button.Types) -and
                   (!($configHash.rtConfig.$button.RequireOnline -and $syncHash.userCompGrid.SelectedItem.Connectivity -eq $false)) -and 
                   (!($configHash.rtConfig.$button.RequireUser -and $syncHash.userCompGrid.SelectedItem.userOnline -eq $false))) {
                    $synchash.customRT.$button.rbut.IsEnabled = $true
                }
                else {
                    $synchash.customRT.$button.rbut.IsEnabled = $false
                }            
            }

            foreach ($button in $syncHash.customContext.Keys) {
                if (($syncHash.userCompGrid.SelectedItem.Type -in $configHash.contextConfig[($button -replace 'cxt') - 1].Types) -and
                   (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireOnline -and $syncHash.userCompGrid.SelectedItem.Connectivity -eq $false)) -and 
                   (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireUser -and $syncHash.userCompGrid.SelectedItem.userOnline -eq $false))) {
                    $syncHash.customContext.$button.('rbutcontext' +  ($button -replace 'cxt')).IsEnabled = $true
                }
                else {
                    $syncHash.customContext.$button.('rbutcontext' +  ($button -replace 'cxt')).IsEnabled = $false
                }            
            }

        }
    })

$syncHash.compUserGrid.Add_SelectionChanged( {

        if ($null -like $syncHash.compUserGrid.SelectedItem) {
            $syncHash.compUserControlPanel.IsEnabled = $false
        }

        else {
            $syncHash.compUserControlPanel.IsEnabled = $true
        }

        if ([string]::IsNullOrEmpty($syncHash.compUserGrid.SelectedItem.ClientName)) {
            $syncHash.compUserFocusClientToggle.Visibility = "Hidden"
            $syncHash.compLogClientPropGrid.Visibility = "Hidden"

        }

        else {
        
            $syncHash.compUserFocusClientToggle.Visibility = "Visible"
            $syncHash.compLogClientPropGrid.Visibility = "Visible"
        }

        
        if ($syncHash.compUserFocusClientToggle.IsChecked) {
            $syncHash.compUserFocusUserToggle.IsChecked = $true
        }

        else {

            foreach ($button in $syncHash.Keys.Where({$_ -like "*rcbutbut*"})) {
                if (($syncHash.compUserGrid.SelectedItem.Type -in $configHash.rtConfig.($syncHash[$button].Tag).Types) -and
                   (!($configHash.rtConfig.($syncHash[$button].Tag).RequireOnline -and $syncHash.compUserGrid.SelectedItem.Connectivity -eq $false)) -and 
                   (!($configHash.rtConfig.($syncHash[$button].Tag).RequireUser -and $syncHash.compUserGrid.SelectedItem.userOnline -eq $false))) {                
                
                    $syncHash[$button].IsEnabled = $true
            
                }
                else {
                    $syncHash[$button].IsEnabled = $false
                }
            }

            foreach ($button in $synchash.customRT.Keys) {
                if (($syncHash.compUserGrid.SelectedItem.Type -in $configHash.rtConfig.$button.Types) -and
                (!($configHash.rtConfig.$button.RequireOnline -and $syncHash.compUserGrid.SelectedItem.Connectivity -eq $false)) -and 
                (!($configHash.rtConfig.$button.RequireUser -and $syncHash.compUserGrid.SelectedItem.userOnline -eq $false))) {
                    $synchash.customRT.$button.rcbut.IsEnabled = $true
                }
                else {
                    $synchash.customRT.$button.rcbut.IsEnabled = $false
                }            
            }
       
            foreach ($button in $syncHash.customContext.Keys) {
                if (($syncHash.compUserGrid.SelectedItem.Type -in $configHash.contextConfig[($button -replace 'cxt') - 1].Types) -and
                   (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireOnline -and $syncHash.compUserGrid.SelectedItem.Connectivity -eq $false)) -and 
                   (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireUser -and $syncHash.compUserGrid.SelectedItem.userOnline -eq $false))) {
                    $syncHash.customContext.$button.('rcbutcontext' +  ($button -replace 'cxt')).IsEnabled = $true
                }
                else {
                    $syncHash.customContext.$button.('rcbutcontext' +  ($button -replace 'cxt')).IsEnabled = $false
                }            
            }
        }
        

    })

$syncHash.userCompFocusHostToggle.Add_Checked( {

        $syncHash.userCompFocusClientToggle.IsChecked = $false    
        
        foreach ($button in $syncHash.Keys.Where({$_ -like "*rbutbut*"})) {
            if (($syncHash.userCompGrid.SelectedItem.Type -in $configHash.rtConfig.($syncHash[$button].Tag).Types) -and
               (!($configHash.rtConfig.($syncHash[$button].Tag).RequireOnline -and $syncHash.userCompGrid.SelectedItem.Connectivity -eq $false)) -and 
               (!($configHash.rtConfig.($syncHash[$button].Tag).RequireUser -and $syncHash.userCompGrid.SelectedItem.userOnline -eq $false))) {                
                
                    $syncHash[$button].IsEnabled = $true
            
            }
            else {
                $syncHash[$button].IsEnabled = $false
            }            
        }

        foreach ($button in $synchash.customRT.Keys) {
            if (($syncHash.userCompGrid.SelectedItem.Type -in $configHash.rtConfig.$button.Types) -and
               (!($configHash.rtConfig.$button.RequireOnline -and $syncHash.userCompGrid.SelectedItem.Connectivity -eq $false)) -and 
               (!($configHash.rtConfig.$button.RequireUser -and $syncHash.userCompGrid.SelectedItem.userOnline -eq $false))) {
                $synchash.customRT.$button.rbut.IsEnabled = $true
            }
            else {
                $synchash.customRT.$button.rbut.IsEnabled = $false
            }            
        }

        foreach ($button in $syncHash.customContext.Keys) {
                if (($syncHash.userCompGrid.SelectedItem.Type -in $configHash.contextConfig[($button -replace 'cxt') - 1].Types) -and
                   (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireOnline -and $syncHash.userCompGrid.SelectedItem.Connectivity -eq $false)) -and 
                   (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireUser -and $syncHash.userCompGrid.SelectedItem.userOnline -eq $false))) {
                    $syncHash.customContext.$button.('rbutcontext' +  ($button -replace 'cxt')).IsEnabled = $true
                }
                else {
                    $syncHash.customContext.$button.('rbutcontext' +  ($button -replace 'cxt')).IsEnabled = $false
                }            
            }
        
          
    })

$syncHash.userCompFocusHostToggle.Add_Unchecked( {
        if ([string]::IsNullOrEmpty($syncHash.userCompGrid.SelectedItem.ClientName)) {
            $syncHash.userCompFocusHostToggle.IsChecked = $true
        }
        else {
            $syncHash.userCompFocusClientToggle.IsChecked = $true     
        } 
    })

$syncHash.userCompFocusClientToggle.Add_Checked( {
        $syncHash.userCompFocusHostToggle.IsChecked = $false   
        
        foreach ($button in $syncHash.Keys.Where({$_ -like "*rbutbut*"})) {
            if (($syncHash.userCompGrid.SelectedItem.ClientType -in $configHash.rtConfig.($syncHash[$button].Tag).Types) -and
               (!($configHash.rtConfig.($syncHash[$button].Tag).RequireOnline -and $syncHash.userCompGrid.SelectedItem.ClientOnline -eq $false)) -and 
               (!($configHash.rtConfig.($syncHash[$button].Tag).RequireUser -eq $true))) {                
                
                    $syncHash[$button].IsEnabled = $true
            
            }
            else {
                $syncHash[$button].IsEnabled = $false
            }            
        }

         foreach ($button in $synchash.customRT.Keys) {
            if (($syncHash.userCompGrid.SelectedItem.ClientType -in $configHash.rtConfig.$button.Types) -and
               (!($configHash.rtConfig.$button.RequireOnline -and $syncHash.userCompGrid.SelectedItem.ClientOnline -eq $false)) -and 
               (!($configHash.rtConfig.$button.RequireUser -ne $true))) {
                $synchash.customRT.$button.rbut.IsEnabled = $true
            }
            else {
                $synchash.customRT.$button.rbut.IsEnabled = $false
            }            
        }

         foreach ($button in $syncHash.customContext.Keys) {
                if (($syncHash.userCompGrid.SelectedItem.ClientType -in $configHash.contextConfig[($button -replace 'cxt') - 1].Types) -and
                   (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireOnline -and $syncHash.userCompGrid.SelectedItem.ClientOnline -eq $false)) -and 
                   (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireUser  -ne $true))) {
                    $syncHash.customContext.$button.('rbutcontext' +  ($button -replace 'cxt')).IsEnabled = $true
                }
                else {
                    $syncHash.customContext.$button.('rbutcontext' +  ($button -replace 'cxt')).IsEnabled = $false
                }            
            }
           
    })

$syncHash.userCompFocusClientToggle.Add_Unchecked( {
        $syncHash.userCompFocusHostToggle.IsChecked = $true      
    })

$syncHash.compUserFocusUserToggle.Add_Checked( {

        $syncHash.compUserFocusClientToggle.IsChecked = $false 
        
        foreach ($button in $syncHash.Keys.Where({$_ -like "*rcbutbut*"})) {
            if (($syncHash.compUserGrid.SelectedItem.Type -in $configHash.rtConfig.($syncHash[$button].Tag).Types) -and
               (!($configHash.rtConfig.($syncHash[$button].Tag).RequireOnline -and $syncHash.compUserGrid.SelectedItem.Connectivity -eq $false)) -and 
               (!($configHash.rtConfig.($syncHash[$button].Tag).RequireUser -and $syncHash.compUserGrid.SelectedItem.userOnline -eq $false))) {                
                
                $syncHash[$button].IsEnabled = $true
            
            }
            else {
                $syncHash[$button].IsEnabled = $false
            }
        }

        foreach ($button in $synchash.customRT.Keys) {
            if (($syncHash.compUserGrid.SelectedItem.Type -in $configHash.rtConfig.$button.Types) -and
            (!($configHash.rtConfig.$button.RequireOnline -and $syncHash.compUserGrid.SelectedItem.Connectivity -eq $false)) -and 
            (!($configHash.rtConfig.$button.RequireUser -and $syncHash.compUserGrid.SelectedItem.userOnline -eq $false))) {
                $synchash.customRT.$button.rcbut.IsEnabled = $true
            }
            else {
                $synchash.customRT.$button.rcbut.IsEnabled = $false
            }            
        }

        foreach ($button in $syncHash.customContext.Keys) {
            if (($syncHash.compUserGrid.SelectedItem.Type -in $configHash.contextConfig[($button -replace 'cxt') - 1].Types) -and
                (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireOnline -and $syncHash.compUserGrid.SelectedItem.Connectivity -eq $false)) -and 
                (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireUser -and $syncHash.compUserGrid.SelectedItem.userOnline -eq $false))) {
                $syncHash.customContext.$button.('rcbutcontext' +  ($button -replace 'cxt')).IsEnabled = $true
            }
            else {
                $syncHash.customContext.$button.('rcbutcontext' +  ($button -replace 'cxt')).IsEnabled = $false
            }            
        }
    })

$syncHash.compUserFocusUserToggle.Add_Unchecked( {
        if ([string]::IsNullOrEmpty($syncHash.compUserGrid.SelectedItem.ClientName)) {
            $syncHash.compUserFocusUserToggle.IsChecked = $true
        }
        else {
            $syncHash.compUserFocusClientToggle.IsChecked = $true     
        } 
    })

$syncHash.compUserFocusClientToggle.Add_Checked( {
        $syncHash.compUserFocusUserToggle.IsChecked = $false  
        
        foreach ($button in $syncHash.Keys.Where({$_ -like "*rcbutbut*"})) {
            if (($syncHash.compUserGrid.SelectedItem.ClientType -in $configHash.rtConfig.($syncHash[$button].Tag).Types) -and
               (!($configHash.rtConfig.($syncHash[$button].Tag).RequireOnline -and $syncHash.compUserGrid.SelectedItem.ClientOnline -eq $false)) -and 
               (!($configHash.rtConfig.($syncHash[$button].Tag).RequireUser -ne $true))) {                
                
                $syncHash[$button].IsEnabled = $true
            
            }
            else {
                $syncHash[$button].IsEnabled = $false
            }
        }

        foreach ($button in $synchash.customRT.Keys) {
            if (($syncHash.compUserGrid.SelectedItem.ClientType -in $configHash.rtConfig.$button.Types) -and
            (!($configHash.rtConfig.$button.RequireOnline -and $syncHash.compUserGrid.SelectedItem.ClientOnline -eq $false)) -and 
            (!($configHash.rtConfig.$button.RequireUser -ne $true))) {
                $synchash.customRT.$button.rcbut.IsEnabled = $true
            }
            else {
                $synchash.customRT.$button.rcbut.IsEnabled = $false
            }            
        }

        foreach ($button in $syncHash.customContext.Keys) {
            if (($syncHash.compUserGrid.SelectedItem.ClientType -in $configHash.contextConfig[($button -replace 'cxt') - 1].Types) -and
                (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireOnline -and $syncHash.compUserGrid.SelectedItem.ClientOnline -eq $false)) -and 
                (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireUser -ne $true))) {
                $syncHash.customContext.$button.('rcbutcontext' +  ($button -replace 'cxt')).IsEnabled = $true
            }
            else {
                $syncHash.customContext.$button.('rcbutcontext' +  ($button -replace 'cxt')).IsEnabled = $false
            }            
        }
            
    })

$syncHash.compUserFocusClientToggle.Add_Unchecked( {
        $syncHash.compUserFocusUserToggle.IsChecked = $true      
    })



$syncHash.SearchBox.add_KeyDown( {

   
        if ($_.Key -eq "Enter" -or $_.Key -eq 'Escape') {
            if ($null -like $syncHash.SearchBox.Text -and $_.Key -ne 'Escape') {
                $syncHash.SnackMsg.MessageQueue.Enqueue("Empty!")
            }
           
            elseif ($syncHash.SearchBox.Text.Length -gt 3 -or $_.Key -eq 'Escape') {

                $rsCmd = [PSObject]@{
                    key        = $($_.Key)
                    searchTag  = $syncHash.SearchBox.Tag
                    searchText = $syncHash.SearchBox.Text
                    queue      = $syncHash.snackMsg.MessageQueue
                }

                
                Start-RSJob -Name Search -ArgumentList $queryHash, $configHash, $match, $syncHash, $rsCmd -ThreadOptions UseNewThread  -FunctionsToImport Test-OnlineFast, Resolve-Location, Get-RDSession -ScriptBlock {
                param($queryHash, $configHash, $match, $syncHash, $rsCmd) 
               
                $rscmd > C:\wtf.txt
                if ($rsCmd.key -eq 'Escape') {
                    $match = (Get-ADObject -Filter "(SamAccountName -eq '$($rsCmd.searchTag)'  -and ObjectClass -eq 'User') -or 
                        (Name -eq '$($rsCmd.searchTag)' -and ObjectClass -eq 'Computer')" -Properties SamAccountName) 
                    
                }
                else {           
                    $match = (Get-ADObject -Filter "(SamAccountName -like '*$($rsCmd.searchText)*' -and ObjectClass -eq 'User') -or 
                        (Name -like '*$($rsCmd.searchText)*' -and ObjectClass -eq 'Computer')" -Properties SamAccountName)
                }

                if (($match | Measure-Object).Count -eq 1) {
                    $syncHash.Window.Dispatcher.Invoke([Action]{                 
                        $syncHash.compExpander.IsExpanded = $false
                        $syncHash.compExpanderProgressBar.Visibility = "Visible"                    
                        $syncHash.userExpander.IsExpanded = $false 
                        $syncHash.expanderProgressBar.Visibility = "Visible"
                    })
                        
                    if ($match.ObjectClass -eq 'User') {
                        $match = (Get-ADUser -Identity $match.SamAccountName -Properties @($configHash.UserPropList.PropName.Where( { $_ -ne 'Non-AD Property' }) | Sort-Object -Unique))                 
                    }

                    elseif ($match.ObjectClass -eq 'Computer') {                       
                        $match = (Get-ADComputer -Identity $match.SamAccountName -Properties @($configHash.CompPropList.PropName.Where( { $_ -ne 'Non-AD Property' }) | Sort-Object -Unique))

                    }
                    if ($match.SamAccountName -notin $syncHash.tabControl.Items.Name -and $match.Name -notin $syncHash.tabControl.Items.Name) {
                                           
                        if ($match.ObjectClass -eq 'User') {
            
                            $queryHash.($match.SamAccountName) = @{}
                            $match.PSObject.Properties | ForEach-Object { 
                                $queryHash.($match.SamAccountName)[$_.Name] = $_.Value }

                             $addItem = ($match | Select-Object @{Label = 'Name'; Expression = { $_.SamAccountName } })
                            $syncHash.Window.Dispatcher.Invoke([Action]{$syncHash.tabControl.ItemsSource.Add($addItem)})
                            $syncHash.Window.Dispatcher.Invoke([Action]{$syncHash.tabControl.SelectedIndex = $syncHash.tabControl.Items.Count - 1})


                            Start-RSJob -Name userLogPull -ArgumentList $queryHash, $configHash, $match, $syncHash -ThreadOptions UseCurrentThread -FunctionsToImport Test-OnlineFast, Resolve-Location, Get-RDSession -ScriptBlock {
                                param($queryHash, $configHash, $match, $syncHash) 
                            
                               

                                $syncHash.userCompGrid.Dispatcher.Invoke([Action] { $syncHash.userCompGrid.ItemsSource = $null })

                                if ($configHash.UserLogPath) {
                                    $queryHash.$($match.SamAccountName).LoginLog = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
                                }
                                  
                                Start-Sleep -Milliseconds 1250
                                  

                                if (Test-Path (Join-Path -Path $confighash.UserLogPath -ChildPath "$($match.SamAccountName).txt")) {
                                        
                                            

                                    $queryHash.$($match.SamAccountName).LoginLogRaw = Get-Content (Join-Path -Path $confighash.UserLogPath -ChildPath "$($match.SamAccountName).txt") | Select-Object -Last 100 | 
                                        ConvertFrom-Csv -Header $configHash.userLogMapping.Header |
                                            Select-Object *, @{Label = 'DateTime'; Expression = { $_.DateRaw -as [datetime] } } -ExcludeProperty DateRaw |
                                                Where-Object { $_.DateTime -gt (Get-Date).AddDays(-60) } |
                                                    Sort-Object DateTime -Descending 
                                    
                                    if ($queryHash.$($match.SamAccountName).LoginLogRaw) {
                                        $loginCounts = $queryHash.$($match.SamAccountName).LoginLogRaw | Group-Object -Property ComputerName | Select-Object Name, Count
                                        $queryHash.$($match.SamAccountName).LoginLogListView = [System.Windows.Data.ListCollectionView]($queryHash.$($match.SamAccountName).LoginLog)  
                                        $queryHash.$($match.SamAccountName).LoginLogListView.GroupDescriptions.Add((New-Object System.Windows.Data.PropertyGroupDescription "compLogon"))
                                        $syncHash.userCompGrid.Dispatcher.Invoke([Action] { $syncHash.userCompGrid.ItemsSource = $queryHash.$($match.SamAccountName).LoginLogListView })
                                        
                                       
                                        $queryHash.$($match.SamAccountName).LoginLogRaw | Sort-Object -Unique -Property ComputerName | Sort-Object DateTime -Descending | ForEach-Object {
                                            
                                            Remove-Variable sessionInfo, clientLocation, hostLocation -ErrorAction SilentlyContinue
                                                
                                            $rawLogEntry = $_
                                            $comp = $_.ComputerName
                                            $ruleCount = ($configHash.nameMapList | Measure-Object).Count
                                            $queryHash.$($match.SamAccountName)
                                            
                                            $hostConnectivity = Test-OnlineFast -ComputerName $_.ComputerName
                                            $clientOnline = Test-OnlineFast -ComputerName $_.ClientName
                                            
                                            if ($hostConnectivity.Online) {
                                                $sessionInfo = Get-RDSession -ComputerName $_.ComputerName -UserName $match.SamAccountName -ErrorAction SilentlyContinue
                                                $hostLocation = (Resolve-Location -computerName $_.ComputerName -IPList $configHash.netMapList -ErrorAction SilentlyContinue).Location
                                            }

                                            if ($clientOnline.Online) {
                                                $clientLocation = (Resolve-Location -computerName $_.ClientName -IPList $configHash.netMapList -ErrorAction SilentlyContinue).Location
                                            }
                                            
                                            
                                            $queryHash.$($match.SamAccountName).LoginLog.Add((
                                                    New-Object PSCustomObject -Property @{
                          
                                                        logonTime      = Get-Date($_.DateTime) -Format MM/dd/yyyy
                                                        HostName       = $_.ComputerName
                                                        LoginDC        = $_.LoginDC
                                                        UserName       = $match.SamAccountName
                                                        Connectivity   = ($hostConnectivity.Online).toString()
                                                        CompType       = "VM"
                                                        IPAddress      = $hostConnectivity.IPV4Address
                                                        userOnline     = if ($sessionInfo) {
                                                                            $true
                                                                        }
                                                                        else {
                                                                            $false
                                                                         }
                                                        sessionID = if ($sessionInfo) {
                                                                        $sessionInfo.sessionID
                                                                    }
                                                                    else {
                                                                        $null
                                                                    }
                                                        IdleTime = if ($sessionInfo) {
                                                                            if ("{0:dd\:hh\:mm}" -f $($sessionInfo.IdleTime) -eq '00:00:00') {
                                                                                "Active"
                                                                            }
                                                                            else {
                                                                             "{0:dd\:hh\:mm}" -f $($sessionInfo.IdleTime)
                                                                            }   
                                                                        }
                                                                        else {$null}
                                                        ClientName     = $_.ClientName 
                                                        ClientLocation = $clientLocation
                                                        compLogon      = if (($queryHash.$($match.SamAccountName).LoginLog | Measure-Object).Count -eq 0) {
                                                            "Last"
                                                        }
                                                        else {
                                                            "Past"
                                                        }
                                                        loginCount     = ($loginCounts | Where-Object { $_.Name -eq $comp }).Count
                                                        DeviceLocation = $hostLocation
                                                        Type           = for ($r = ($ruleCount - 1); $r -ge 0; $r--) {
    
                                                            if ($r -eq 0) {
                                                                "Computer"
                                                            }

                                                            else {

                                                                if ($configHash.nameMapList[$r].Condition) {
                                                                    try {
                                                                        if (Invoke-Expression $configHash.nameMapList[$r].Condition) {
                                                                            $configHash.nameMapList[$r].Name
                                                                            break
                                                                        }
                                                                    }

                                                                    catch {}
            
                                                                   
                                                                }
                                                            }
                                                        }
                                                        ClientOnline = ($clientOnline.Online).toString()
                                                        ClientIPAddress = $clientOnline.IPV4Address
                                                        ClientType           = for ($r = ($ruleCount - 1); $r -ge 0; $r--) {
                                                                
                                                            $comp = $_.ClientName

                                                            if ($r -eq 0) {
                                                                "Computer"
                                                            }

                                                            else {

                                                                if ($configHash.nameMapList[$r].Condition) {
                                                                    try {
                                                                        if (Invoke-Expression $configHash.nameMapList[$r].Condition) {
                                                                            $configHash.nameMapList[$r].Name
                                                                            break
                                                                        }
                                                                    }

                                                                    catch {}
            
                                                                   
                                                                }
                                                            }
                                                        }

                                                    }))
                                               
                                            if ($configHash.userLogMapping.FieldSel -contains 'Custom') {
                                                foreach ($customHeader in ($configHash.userLogMapping | Where-Object { $_.FieldSel -eq 'Custom' })) {
                                                    foreach ($item in ($queryHash.$($match.SamAccountName).LoginLog)) {
                                                        $item | Add-Member -Force -MemberType NoteProperty -Name $customHeader.Header -Value $rawLogEntry.($customHeader.Header)
                                                    }
                                                }
                                            }

                                        
                                           $syncHash.Window.Dispatcher.Invoke([Action]{$queryHash.$($match.SamAccountName).LoginLogListView.Refresh()})
                                           # $syncHash.userCompGrid.Dispatcher.InvokeAsync([Action] { $syncHash.userCompGrid.ItemsSource.Refresh() })

                                            if (($syncHash.userCompGrid.Items | Measure-Object).Count -eq 1) {
                                            
                                                $syncHash.Window.Dispatcher.Invoke([Action] { 
                                                    $syncHash.UserCompGrid.SelectedItem = $syncHash.UserCompGrid.Items[0]
                                                })

                                                $queryHash[$match.SamAccountName].logsSearched = $true
                                            }

                                        }                                  
                                            
                                    }
                                    else {
                                        $queryHash[$match.SamAccountName].logsSearched = $true
                                    }
                                }
                                    
                                else {
                                    $queryHash[$match.SamAccountName].logsSearched = $true
                                }
                            }      
                        }
                    
                        elseif ($match.ObjectClass -eq 'Computer') {
                                             
                            $queryHash.($match.Name) = @{}
                            $match.PSObject.Properties | ForEach-Object { 
                                $queryHash.($match.Name)[$_.Name] = $_.Value }

                            if ($configHash.pcLogPath) {
                                $queryHash.$($match.Name).LoginLog = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
                            }

                            $addItem = ($match | Select-Object @{Label = 'Name'; Expression = { $_.Name } })
                            $syncHash.Window.Dispatcher.Invoke([Action]{$syncHash.tabControl.ItemsSource.Add($addItem)})  
                            $syncHash.Window.Dispatcher.Invoke([Action]{$syncHash.tabControl.SelectedIndex = $syncHash.tabControl.Items.Count - 1})

                            Start-RSJob -Name compLogPull -ArgumentList $queryHash, $configHash, $match, $syncHash -FunctionsToImport Resolve-Location, Get-RDSession, Test-OnlineFast -ScriptBlock {
                                param($queryHash, $configHash, $match, $syncHash) 
                            
                                $syncHash.userCompGrid.Dispatcher.Invoke([Action] { $syncHash.compUserGrid.ItemsSource = $null })
                                  
                                Start-Sleep -Milliseconds 1250                               

                                if (Test-Path (Join-Path -Path $configHash.pcLogPath -ChildPath "$($match.Name).txt")) {
                                        
                                            
                                    $queryHash.$($match.Name).LoginLogRaw = Get-Content (Join-Path -Path $confighash.pcLogPath -ChildPath "$($match.Name).txt") | Select-Object -Last 100 | 
                                        ConvertFrom-Csv -Header $configHash.compLogMapping.Header |
                                            Select-Object *, @{Label = 'DateTime'; Expression = { $_.DateRaw -as [datetime] } } -ExcludeProperty DateRaw |
                                                Where-Object { $_.DateTime -gt (Get-Date).AddDays(-60) } |
                                                    Sort-Object DateTime -Descending 
                                        
                                    if ($queryHash.$($match.Name).LoginLogRaw) {



                                        $loginCounts = $queryHash.$($match.Name).LoginLogRaw | Group-Object -Property User | select Name, Count
                                        $queryHash.$($match.Name).LoginLogListView = [System.Windows.Data.ListCollectionView]($queryHash.$($match.Name).LoginLog)  
                                        $queryHash.$($match.Name).LoginLogListView.GroupDescriptions.Add((New-Object System.Windows.Data.PropertyGroupDescription "compLogon"))
                                        $syncHash.compUserGrid.Dispatcher.Invoke([Action] { $syncHash.compUserGrid.ItemsSource = $queryHash.$($match.Name).LoginLogListView })
                                        $ruleCount = ($configHash.nameMapList | Measure-Object).Count

                                       
                                            

                                        $compType = for ($r = ($ruleCount - 1); $r -ge 0; $r--) {
                                                            $comp = $match.Name
                                                            if ($r -eq 0) {
                                                                "Computer"
                                                            }

                                                            else {

                                                                if ($configHash.nameMapList[$r].Condition) {
                                                                    try {
                                                                        if (Invoke-Expression $configHash.nameMapList[$r].Condition) {
                                                                            $configHash.nameMapList[$r].Name
                                                                            break
                                                                        }
                                                                    }

                                                                    catch {}
            
                                                                   
                                                                }
                                                            }
                                                        }


                                        if ((Test-OnlineFast $match.Name).Online) {
                                            $sessionInfo = Get-RDSession -ComputerName $match.Name -ErrorAction SilentlyContinue
                                        }

                                        $queryHash.$($match.Name).LoginLogRaw | Sort-Object -Unique -Property User | Sort-Object DateTime -Descending | ForEach-Object {
                                                                                                         
                                            $rawLogEntry = $_
                                            $tempCN = $_.User                                           
                                            $clientOnline = Test-OnlineFast -ComputerName $_.ClientName
                                            $userSession = $sessionInfo.Where{$_.UserName -eq $tempCN}                                           
                                            
                                            Remove-Variable clientLocation -ErrorAction SilentlyContinue

                                            if ($clientOnline.Online) {

                                                $clientLocation = (Resolve-Location -ComputerName $_.ClientName -IPList $configHash.netMapList).Location
                                            }
                                            


                                            $queryHash.$($match.Name).LoginLog.Add((
                                                    New-Object PSCustomObject -Property @{
                                                        logonTime  = Get-Date($_.DateTime) -Format MM/dd/yyyy
                                                        UserName   = $_.User
                                                        LoginDC    = $_.LoginDC
                                                        Name       = (Get-ADUser -Identity $_.User).Name
                                                        userOnline     = if ($userSession) {
                                                                            $true
                                                                        }
                                                                        else {
                                                                            $false
                                                                         }
                                                        sessionID = if ($userSession) {
                                                                        $userSession.SessionId
                                                                    }
                                                                    else {
                                                                        $null
                                                                    }
                                                        IdleTime = if ($userSession) {
                                                                            if ("{0:dd\:hh\:mm}" -f $($userSession.IdleTime) -eq '00:00:00') {
                                                                                "Active"
                                                                            }
                                                                            else {
                                                                             "{0:dd\:hh\:mm}" -f $($userSession.IdleTime)
                                                                            }   
                                                                        }
                                                                        else {$null}
                                                        ClientName = $_.ClientName 
                                                        Type = $compType
                                                        compLogon  = if (($queryHash.$($match.Name).LoginLog | Measure-Object).Count -eq 0) {
                                                                        "Last"
                                                                     }
                                                                     else {
                                                                        "Past"
                                                                     }
                                                        loginCount = ($loginCounts | Where-Object { $_.Name -eq $tempCN }).Count
                                                        ClientOnline = ($clientOnline.Online).toString()
                                                        ClientIPAddress = $clientOnline.IPV4Address
                                                        ClientLocation = $clientLocation
                                                        
                                                        ClientType = for ($r = ($ruleCount - 1); $r -ge 0; $r--) {
                                                                
                                                            $comp = $_.ClientName

                                                            if ($r -eq 0) {
                                                                "Computer"
                                                            }

                                                            else {

                                                                if ($configHash.nameMapList[$r].Condition) {
                                                                    try {
                                                                        if (Invoke-Expression $configHash.nameMapList[$r].Condition) {
                                                                            $configHash.nameMapList[$r].Name
                                                                            break
                                                                        }
                                                                    }

                                                                    catch {}
            
                                                                   
                                                                }
                                                            }
                                                        }
                                                    }))
                                        
                                            if ($configHash.compLogMapping.FieldSel -contains 'Custom') {
                                                foreach ($customHeader in ($configHash.compLogMapping | Where-Object { $_.FieldSel -eq 'Custom' })) {
                                                    foreach ($item in ($queryHash.$($match.Name).LoginLog)) {
                                                        $item | Add-Member -Force -MemberType NoteProperty -Name $customHeader.Header -Value $rawLogEntry.($customHeader.Header)
                                                    }
                                                }
                                            }
                                            $syncHash.Window.Dispatcher.Invoke([Action]{$queryHash.$($match.Name).LoginLogListView.Refresh()})
                                            #$syncHash.compUserGrid.Dispatcher.Invoke([Action] { $syncHash.compUserGrid.ItemsSource.Refresh() })

                                            if (($syncHash.compUserGrid.Items | Measure-Object).Count -eq 1) {
                                            
                                                $syncHash.Window.Dispatcher.Invoke([Action] { 
                                                    $syncHash.compUserGrid.SelectedItem = $syncHash.compUserGrid.Items[0]
                                                })

                                                $queryHash[$match.Name].logsSearched = $true
                                            }

                                        }                                  
                                            
                                    }
                                            
                                    else {
                                        $queryHash[$match.Name].logsSearched = $true
                                    }
                                
                                }
                                    
                                else {
                                   $queryHash[$match.Name].logsSearched = $true       
                                }
                            }
                        }

                         
                    }

                    else {
                        if ($match.ObjectClass -eq 'User') {
                            $match.PSObject.Properties | ForEach-Object { $queryHash.($match.SamAccountName)[$_.Name] = $_.Value }
                            $itemIndex = [Array]::IndexOf($syncHash.tabControl.Items.Name,$($match.SamAccountName))     
                            $syncHash.Window.Dispatcher.Invoke([Action]{ $syncHash.tabControl.SelectedIndex = $itemIndex  })
                        }
                        else {
                            $match.PSObject.Properties | ForEach-Object { $queryHash.($match.Name)[$_.Name] = $_.Value }
                            $itemIndex = [Array]::IndexOf($syncHash.tabControl.Items.Name,$($match.Name))     
                            $syncHash.Window.Dispatcher.Invoke([Action]{ $syncHash.tabControl.SelectedIndex = $itemIndex  })
                        }   
                    }

                    $syncHash.Window.Dispatcher.Invoke([Action]{$syncHash.userGrid.Visibility = "Visible"})

                    

                }

                elseif (($match | Measure-Object).Count -gt 1) {
                    $rsCmd.queue.Enqueue("Too many matches!")
                    $syncHash.Window.Dispatcher.Invoke([Action]{
                        $syncHash.resultsSidePane.IsOpen = $true
                        $syncHash.resultsSidePaneGrid.ItemsSource = $match
                    })
                }

                else {
                     $rsCmd.queue.Enqueue("No match!")
                }

                }
            }

            else {
                $syncHash.SnackMsg.MessageQueue.Enqueue("Query must be at least 3 characters long!")
            }
        }
   
    })

$syncHash.itemToolDialogConfirmButton.Add_Click({

    Start-RSJob -Name ItemTool -ArgumentList $syncHash.snackMsg.MessageQueue, $syncHash.itemToolDialogConfirmButton.Tag, $configHash, $queryHash -ScriptBlock {
    Param($queue, $toolID, $configHash, $queryHash)

        $item = ($configHash.currentTabItem).toLower()
        $toolName = ($configHash.objectToolConfig[$toolID - 1].toolActionToolTip).ToUpper()

        try {                     
            Invoke-Expression $configHash.objectToolConfig[$toolID - 1].toolAction
            $queue.Enqueue("[$toolName]: Success on [$item] - tool action complete")
        }
        catch {
            $queue.Enqueue("[$toolName]: Fail on [$item] - tool action incomplete")
        }
    }

    $syncHash.itemToolDialog.IsOpen = $false
})

$syncHash.itemToolDialogConfirmCancel.Add_Click({

    $syncHash.itemToolDialog.IsOpen = $false
   
})

$syncHash.itemToolDialog.Add_ClosingFinished({
    $syncHash.itemToolDialogConfirm.Visibility = "Collapsed"
    $syncHash.itemToolListSelect.Visibility = "Collapsed"
    $syncHash.itemToolListSelectListBox.ItemsSource = $null
    $syncHash.itemToolADSelectedItem.Content = $null
    $syncHash.itemToolGridADSelectedItem.Content = $null
    $syncHash.itemToolImageBorder.Visibility = "Collapsed"
    $syncHash.itemToolGridSelect.Visibility = "Collapsed"

})

[System.Windows.RoutedEventHandler]$eventonNetDataGrid = {
    $button = $_.OriginalSource
    $configHash.button = $button

    if ($button.Name -match "settingNetClearItem") { 
        $configHash.netMapList.Remove(($configHash.netMapList | Where-Object { $_.ID -eq ($syncHash.settingNetDataGrid.SelectedItem.ID) }))
    }

    elseif ($button.Name -match "settingnetIP") {
        try {
            [IPAddress]$syncHash.settingNetDataGrid.SelectedItem.Network
            $syncHash.settingNetDataGrid.SelectedItem.validNetwork = $true
            $button.Foreground = "White"
            $button.Tooltip = $null
            $button.TextDecorations = $null
        }
        catch {
            $syncHash.settingNetDataGrid.SelectedItem.validNetwork = $false
            $button.Foreground = "Red"
            $button.Tooltip = "Invalid IP"
            $button.TextDecorations = "Underline"
        }
    }

    elseif ($button.Name -match "settingnetMask") {
        try {
                       
            if ([int]$syncHash.settingNetDataGrid.SelectedItem.Mask -gt 0 -and [int]$syncHash.settingNetDataGrid.SelectedItem.Mask -le 32) {
                $syncHash.settingNetDataGrid.SelectedItem.validMask = $true
                $button.Foreground = "White"
                $button.Tooltip = $null
                $button.TextDecorations = $null
            }
            else {
                1 / 0
            }

        }
        catch {
            $syncHash.settingNetDataGrid.SelectedItem.validMask = $false
            $button.Foreground = "Red"
            $button.Tooltip = "Invalid mask"
            $button.TextDecorations = "Underline"
        }
    }
}

[System.Windows.RoutedEventHandler]$eventonNameDataGrid = {
    $button = $_.OriginalSource
    $configHash.button = $button

    if ($button.Name -match "settingNameClearItem") { 
        $id = $syncHash.settingNameDataGrid.SelectedItem.ID 
        $configHash.NameMapList.Remove(($configHash.NameMapList | Where-Object { $_.ID -eq ($syncHash.settingNameDataGrid.SelectedItem.ID) }))
        ($configHash.NameMapList | Where-Object { $_.Id -eq ($syncHash.settingNameDataGrid.SelectedItem.ID) })


        
       
        $configHash.NameMapList | Where-Object { $_.Id -gt $id } | ForEach-Object { $_.Id = $_.Id - 1 }
       
        $configHash.nameMapListView.Refresh()
        $syncHash.settingNameDataGrid.Items.Refresh()
       

    }

    elseif ($button.Name -match "settingNameUpItem") {
        $ID = $syncHash.settingNameDataGrid.SelectedItem.Id


        if ($ID + 1 -eq ($configHash.NameMapList.Id | Sort-Object -Descending | Select-Object -First 1)) {
            ($configHash.NameMapList | Where-Object { $_ -eq $syncHash.settingNameDataGrid.SelectedItem }).TopPos = $true
        }

        ($configHash.NameMapList | Where-Object { $_.Id -eq $ID + 1 }).TopPos = $false
        ($configHash.NameMapList | Where-Object { $_.Id -eq $ID + 1 }).Id = $ID


       
        ($configHash.NameMapList | Where-Object { $_ -eq $syncHash.settingNameDataGrid.SelectedItem }).ID = $ID + 1
        $configHash.nameMapListView.Refresh()
        $syncHash.settingNameDataGrid.Items.Refresh()
    }

    elseif ($button.Name -match "settingNameDownItem") {
        $ID = $syncHash.settingNameDataGrid.SelectedItem.Id
    
        if ($ID -eq ($configHash.NameMapList.Id | Sort-Object -Descending | Select-Object -First 1)) {
            ($configHash.NameMapList | Where-Object { $_.Id -eq $ID - 1 }).TopPos = $true
        }

        ($configHash.NameMapList | Where-Object { $_.Id -eq $ID - 1 }).Id = $ID

        ($configHash.NameMapList | Where-Object { $_ -eq $syncHash.settingNameDataGrid.SelectedItem }).TopPos = $false
        ($configHash.NameMapList | Where-Object { $_ -eq $syncHash.settingNameDataGrid.SelectedItem }).ID = $ID - 1
        $configHash.nameMapListView.Refresh()
        $syncHash.settingNameDataGrid.Items.Refresh()
    }

    elseif ($button.Name -match "settingConditionBox") {

        $syncHash.settingNameDialog.IsOpen = $true
          

    }


}

$syncHash.settingNameDialog.Add_DialogClosing( {
        $synchash.settingNameDataGrid.Items.Refresh()
        $configHash.nameMapListView.Refresh()
    })

[System.Windows.RoutedEventHandler]$EventonDataGrid = {

    # GET THE NAME OF EVENT SOURCE
    $button = $_.OriginalSource
    # THIS RETURN THE ROW DATA AVAILABLE
    # resultObj scope is the whole script
    if ($syncHash.settingUserPropGrid.Visibility -eq 'Visible') {
        $type = 'User'
    }

    else {
        $type = 'Comp'
    }

    if ($button.Name -match 'settingActionComboBox' -or $button.Name -match "settingEdit") {

        $switchVal = (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1]).ActionName)
        switch -wildcard ($switchVal) {
            
            { $switchVal -notmatch "raw" } { 
                  
                $syncHash.settingFlyoutTranslate.Visibility = "Visible"
                (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).transCmdsEnabled = $true   

            }

            { $switchVal -notmatch "actionable" } { 
                $syncHash.settingFlyoutTranslate.IsExpanded = $true
                $syncHash.settingFlyoutAction1.Visibility = "Collapsed"
                $syncHash.settingFlyoutAction2.Visibility = "Collapsed"
                (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).actionCmdsEnabled = $false
                    
                if ($null -like (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).ActionCmd1) {
                    (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).ActionCmd1 = "Do-Something -$type $('$' + $type)..."   
                } 
                                   
                if ($null -like (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).ActionCmd2) {
                    (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).ActionCmd2 = "Do-Something -$type $('$' + $type)..." 
                }    
            }

            { $switchVal -match "raw" } { 
                $syncHash.settingFlyoutTranslate.Visibility = "Collapsed"                     
                (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).transCmdsEnabled = $false                   
                if ($null -like (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).TranslationCmd) {
                    (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).TranslationCmd = 'if ($result -eq $false)...'
                }
            }

            { $switchVal -match "actionable" } {                  
                $syncHash.settingFlyoutAction1.Visibility = "Visible"
                $syncHash.settingFlyoutAction2.Visibility = "Visible"
                if ($syncHash.settingFlyoutTranslate.Visibility -eq "Collapsed") {
                    $syncHash.settingFlyoutAction1.IsExpanded = $true
                }
            }
        }

        if ($button.Name -match "settingEdit") {
 
            $syncHash.settingResultBox.Text = ($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1]).Result
            $syncHash.settingBox1Icon.ItemsSource = $configHash.buttonGlyphs
            $syncHash.settingBox2Icon.ItemsSource = $configHash.buttonGlyphs
            $syncHash.settingUserPropDefFlyout.Height = $syncHash.settingChildHeight.ActualHeight
           
            $syncHash.settingUserPropDefFlyout.IsOpen = $true
            $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.settingUserPropDefFlyout.Background.Color).ToString()
            $syncHash.settingChildWindow.Title = "Define Button ($type)"
            $syncHash.settingChildWindow.ShowCloseButton = $false

        }

    }

    elseif ($button.Name -match ("setting" + "$type" + "ComboBox") -and ($syncHash.('setting' + $type + 'PropGrid').SelectedItem)) {
           
        Remove-Variable set -ErrorAction SilentlyContinue
        $set = (($configHash.($type + 'PropPullList') | Where-Object { $_.Name -eq ($syncHash.('setting' + $type + 'PropGrid').SelectedItem).PropName }).TypeNameOfValue -replace ".*(?=\.).", "").toString()
        $syncHash.settingTypeBox.Text = $set
        ($configHash.($type + 'PropList') | Where-Object { $_.Field -eq (($syncHash.('setting' + $type + 'PropGrid').SelectedItem).Field) }).PropType = $set
    }

    elseif ($button.Name -match "settingClearItem") {  


        $num = $syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field 
        $configHash.($type + 'PropList').RemoveAt($syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1)
        $configHash.($type + 'PropList') | Where-Object { $_.Field -gt $num } | ForEach-Object { $_.Field = $_.Field - 1 }
       

        $syncHash.('setting' + $type + 'PropGrid').Items.Refresh()
              
    }

    if (($syncHash.('setting' + $type + 'PropGrid').SelectedItem).PropName -eq 'Non-Ad Property') {
        $syncHash.settingFlyoutResultHeader.Content = "Retrieval Command"
    }

   
   

    else {
        $syncHash.settingFlyoutResultHeader.Content = "Result Presentation"
    }
}

[System.Windows.RoutedEventHandler]$EventonContextGrid = {

    # GET THE NAME OF EVENT SOURCE
    $button = $_.OriginalSource
    # THIS RETURN THE ROW DATA AVAILABLE
    # resultObj scope is the whole script
   

    if ($button.Name -match 'settingContextEdit') {


    $syncHash.settingContextResultBox.Text = $configHash.contextConfig[$syncHash.settingContextPropGrid.SelectedItem.IDNum - 1].Result
    $syncHash.settingContextIcon.ItemsSource = $configHash.buttonGlyphs
    $syncHash.settingContextDefFlyout.Height = $syncHash.settingChildHeight.ActualHeight      
    $syncHash.settingContextDefFlyout.IsOpen = $true
    $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.settingContextDefFlyout.Background.Color).ToString()
    $syncHash.settingChildWindow.Title = "Define Context Button"
    $syncHash.settingChildWindow.ShowCloseButton = $false

    }

  

    elseif ($button.Name -match "settingContextClearItem") {  

        $num = $syncHash.settingContextPropGrid.SelectedItem.IDNum 
        $configHash.contextConfig.RemoveAt($syncHash.settingContextPropGrid.SelectedItem.IDNum - 1)
        $configHash.contextConfig | Where-Object { $_.IDNum -gt $num } | ForEach-Object { $_.IDNum = $_.IDNum - 1 }       
        $syncHash.settingContextPropGrid.Items.Refresh()
              
    }
}

[System.Windows.RoutedEventHandler]$EventonObjectToolGrid = {

    # GET THE NAME OF EVENT SOURCE
    $button = $_.OriginalSource
    # THIS RETURN THE ROW DATA AVAILABLE
    # resultObj scope is the whole script
   

    if ($button.Name -match 'settingObjectToolsEdit') {

    $syncHash.settingObjectToolDefFlyout.Height = $syncHash.settingChildHeight.ActualHeight      
    $syncHash.settingObjectToolDefFlyout.IsOpen = $true
    $syncHash.settingObjectToolIcon.ItemsSource = $configHash.buttonGlyphs
    $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.settingObjectToolDefFlyout.Background.Color).ToString()
    $syncHash.settingChildWindow.Title = "Define Tool"
    $syncHash.settingChildWindow.ShowCloseButton = $false

    }

    elseif ($button.Name -match "settingObjectToolsClearItem") {  

        $num = $syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID 
        $configHash.objectToolConfig.RemoveAt($syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID  - 1)
        $configHash.objectToolConfig | Where-Object { $_.ToolID -gt $num } | ForEach-Object { $_.ToolID = $_.ToolID - 1 }       
        $syncHash.settingObjectToolsPropGrid.Items.Refresh()
              
    }
}

$syncHash.settingFlyoutExit.Add_Click( {


        if ($syncHash.settingUserPropGrid.Visibility -eq 'Visible') {
            $type = 'User'
        }

        else {
            $type = 'Computer'
        }

        $syncHash.settingUserPropDefFlyout.IsOpen = $false
        $syncHash.settingChildWindow.ShowCloseButton = $true
        $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.Window.BorderBrush.Color).ToString()
        $syncHash.settingChildWindow.Title = "$type Property Mappings"
        $syncHash.settingResultBox.Foreground = "White"

    })



$syncHash.settingExecute.Add_Click( {

        if ($syncHash.settingUserPropGrid.Visibility -eq 'Visible') {
            $type = 'User'
            $user = $syncHash.('setting' + $type + 'PropGrid').SelectedItem.querySubject
        }

        else {
            $type = 'Comp'
            $comp = $syncHash.('setting' + $type + 'PropGrid').SelectedItem.querySubject
        }

        Remove-Variable -Name resultColor -ErrorAction SilentlyContinue          
   
        # Empty collection for errors
        $scriptBlockErrors = @()

        # Define input script
        [void][System.Management.Automation.Language.Parser]::ParseInput(($syncHash.('setting' + $type + 'PropGrid').SelectedItem.translationCmd), [ref]$null, [ref]$scriptBlockErrors)

        if ($scriptBlockErrors) {
                
            $syncHash.settingResultBox.Text = "Invalid scriptblock - $($scriptBlockErrors.Count) errors"
            $configHash.UserPropList[($syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1)].ValidCmd = $false

            $syncHash.settingResultError.ToolTip = $null
        
            for ($err = 0; $err -lt $scriptBlockErrors.Count; $err++) {

                if ($err -eq $scriptBlockErrors.Count - 1) {
                    $syncHash.settingResultError.ToolTip += [string] $scriptBlockErrors[$err].Message
                 
                }
                else {
                    $syncHash.settingResultError.ToolTip += "$([string] $scriptBlockErrors[$err].Message) `n" 
                }
            }
      
            $syncHash.settingResultError.Visibility = 'Visible'    
       
        }

        else {      

            if ($syncHash.('setting' + $type + 'PropGrid').SelectedItem.PropName -ne 'Non-AD Property') {
                if ($type -eq 'User') {
                    $result = (Get-ADUser -Identity $syncHash.settingUserPropGrid.SelectedItem.querySubject -Properties $syncHash.settingUserPropGrid.SelectedItem.PropName -ErrorAction Continue).($syncHash.settingUserPropGrid.SelectedItem.PropName) 
                }
                else {
                    $result = (Get-ADComputer -Identity $syncHash.settingCompPropGrid.SelectedItem.querySubject -Properties $syncHash.settingCompPropGrid.SelectedItem.PropName -ErrorAction Continue).($syncHash.settingCompPropGrid.SelectedItem.PropName) 
                }               
            }
        
            $return = Invoke-Expression -Command $syncHash.('setting' + $type + 'PropGrid').SelectedItem.translationCmd -ErrorAction Continue -ErrorVariable errorVar
       
            if ($errorVar) {
                $syncHash.settingResultError.ToolTip = [string]$errorVar
                $syncHash.settingResultError.Visibility = 'Visible'             
            }

            else {
                $syncHash.settingResultError.Visibility = 'Hidden'
            }

            if (($return | Measure-Object -Line).Lines -gt 2) {
                $syncHash.settingResultBox.Text = ("Invalid: Return too long ($($($return | Measure-Object -Line).Lines) lines)")
            }

            else {
                if ($null -eq $return) {
                    if ($errorVar) {
                        $syncHash.settingResultBox.Text = ("Invalid: Returned error code")
                    }
                    
                    else {
                        $syncHash.settingResultBox.Text = ("Valid (return was empty)")
                    }
                }
                else {
                    $syncHash.settingResultBox.Text = ("Valid: " + [string]$return)
                }
                $return = [string]$return
                $configHash.($type + 'PropList')[($syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1)].result = $return
                $configHash.($type + 'PropList')[($syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1)].validCmd = $true
               
            }


            if ($resultColor) {
    
                try {
                    $syncHash.settingResultBox.Foreground = $resultColor
                }

                catch { }
            }
        }

    })


$syncHash.settingAction2HidablePanel.Add_isEnabledChanged( {

        if ($syncHash.settingUserPropGrid.Visibility -eq 'Visible') {
            $type = 'User'
        }

        else {
            $type = 'Comp'
        }

        $configHash.($type + 'PropList')[($syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1)].actionCmd2Enabled = $syncHash.settingAction2Enable.IsChecked
   
    })


for ($i = 1; $i -le 2; $i++) {

    $syncHash.('settingBox' + $i + 'Execute').Add_Click( {

            param([Parameter(Mandatory)][Object]$sender)
            
            $id = $sender.Name -replace "settingBox|Execute" 
            $type = $sender.DataContext.ItemType
            #  try {
            #      $user = (Get-AdUser -Identity $syncHash.settingUserPropGrid.SelectedItem.querySubject).SamAccountName
        
            #      if ($null -eq $user) {
            #           $configHash.UserPropList[($syncHash.settingUserPropGrid.SelectedItem.Field-1)].Result = "User could not be found"
            #           break
            #        }


            # Empty collection for errors
            $scriptBlockErrors = @()

            # Define input script
            [void][System.Management.Automation.Language.Parser]::ParseInput(($syncHash.('setting' + $type + 'PropGrid').SelectedItem.('actionCmd' + $id)), [ref]$null, [ref]$scriptBlockErrors)

            if ($scriptBlockErrors) {
                
                $configHash.($type + 'PropList')[(('setting' + $type + 'PropGrid').SelectedItem.Field - 1)].('actionCmd' + $id + 'result') = "Invalid scriptblock - $($scriptBlockErrors.Count) errors"
                $configHash.($type + 'PropList')[(('setting' + $type + 'PropGrid').SelectedItem.Field - 1)].('ValidAction' + $id) = $false
            }

            else {

                $configHash.($type + 'PropList')[($syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1)].('actionCmd' + $id + 'result') = "Scriptblock validated"
                $configHash.($type + 'PropList')[($syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1)].('ValidAction' + $id) = $true
            
            }
          
            #   }
            
            #   catch { 
     
            #      $configHash.UserPropList[($syncHash.settingUserPropGrid.SelectedItem.Field-1)].('actionCmd' + $id + 'result') = "Invalid principal for search."
            #       $configHash.UserPropList[($syncHash.settingUserPropGrid.SelectedItem.Field-1)].('ValidAction' + $id) = $false 
            #   }

            $syncHash.('settingBox' + $id + 'ResultBox').Text = $configHash.($type + 'PropList')[($syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1)].('actionCmd' + $id + 'result')           
        }        
    )
}

$syncHash.settingContextExecute.Add_Click( {
    # Empty collection for errors
    $scriptBlockErrors = @()

    # Define input script
    [void][System.Management.Automation.Language.Parser]::ParseInput(($syncHash.settingContextPropGrid.SelectedItem.actionCmd), [ref]$null, [ref]$scriptBlockErrors)

    if ($scriptBlockErrors) {
                
        $syncHash.settingContextResultBox.Text = "Invalid scriptblock - $($scriptBlockErrors.Count) errors"
        $configHash.contextConfig[($syncHash.settingContextPropGrid.SelectedItem.IDNum - 1)].ValidAction = $false
    }

    else {
        
                
        $syncHash.settingContextResultBox.Text = "Scriptblock validated"
        $configHash.contextConfig[($syncHash.settingContextPropGrid.SelectedItem.IDNum - 1)].ValidAction = $true
       
    }
})

$syncHash.settingObjectToolExecute.Add_Click({
    # Empty collection for errors
    $scriptBlockErrors = @()

    # Define input script
    [void][System.Management.Automation.Language.Parser]::ParseInput(($syncHash.settingObjectToolsPropGrid.SelectedItem.toolAction), [ref]$null, [ref]$scriptBlockErrors)

    if ($scriptBlockErrors) {
                
        $syncHash.settingObjectToolResultBox.Text = "Invalid scriptblock - $($scriptBlockErrors.Count) errors"
        $configHash.objectToolConfig[($syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID - 1)].toolActionValid = $false
    }

    else {
        
                
        $syncHash.settingObjectToolResultBox.Text = "Scriptblock validated"
        $configHash.objectToolConfig[($syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID - 1)].toolActionValid = $true
       
    }


})

[System.Windows.RoutedEventHandler]$addItemClick = {
 
    if ($syncHash.settingUserPropGrid.Visibility -eq 'Visible') {
        $type = 'User'
    }

    else {
        $type = 'Comp'
    }

    $i = ($configHash.($type + 'PropList').Field | Sort-Object -Descending | Select-Object -First 1) + 1

    $configHash.($type + 'PropList').Add([PSCustomObject]@{

            Field             = $i
            FieldName         = $null
            ItemType          = $type
            PropName          = $null
            propList          = $configHash.($type + 'PropPullListNames')
            translationCmd    = 'if ($result -eq $false)...'
            actionCmd1        = 'Do-Something -User $user'
            actionCmd1ToolTip = 'Action name'                 
            actionCmd1Icon    = $null
            actionCmd1Refresh = $false
            actionCmd1Multi   = $false 
            ValidCmd          = $false
            ValidAction1      = $false
            ValidAction2      = $false
            actionCmd2        = 'Do-Something -User $user'
            actionCmd2ToolTip = 'Action name'                  
            actionCmd2Icon    = $null 
            actionCmd2Refresh = $false  
            actionCmd2Multi   = $false                    
            querySubject      = $env:USERNAME                                       
            result            = '(null)'
            actionCmdsEnabled = $true
            transCmdsEnabled  = $true 
            actionCmd1result  = '(null)'
            actionCmd2result  = '(null)'
            actionCmd2Enabled = $false 
            PropType          = $null
            actionList        = @('ReadOnly', 'ReadOnly-Raw', 'Editable', 'Editable-Raw', 'Actionable', 'Actionable-Raw', 'Editable-Actionable', 'Editable-Actionable-Raw')
                      
            ActionName        = 'null'
              
        })
      
    $configHash.('box' + $type + 'Count') = ($configHash.($type + 'PropList') | Measure-Object).Count
}

[System.Windows.RoutedEventHandler]$addObjectToolItemClick = {
 

    $i = ($configHash.objectToolConfig.ToolID | Sort-Object -Descending | Select-Object -First 1) + 1

    $configHash.objectToolConfig.Add([PSCustomObject]@{

            ToolID                = $i
            ToolName              = $null
            toolTypeList          = @("Execute","Select","Grid","List")
            toolType              = 'null'
            objectList            = @("Comp","User","Both")
            objectType            = $null
            toolAction            = 'Do-Something -UserName $user'
            toolActionValid       = $false
            toolActionConfirm     = $true
            toolActionToolTip     = $null
            toolActionIcon        = $null
            toolActionSelectAD    = $false
            toolFetchCmd          = 'Get-Something'
            toolActionMultiSelect = $false
            toolDescription       = "Sort item description"
            toolTargetFetchCmd    = 'Get-Something -Identity $target'
              
        })
      
    $configHash.objectToolCount = ($configHash.objectToolConfig | Measure-Object).Count
}

[System.Windows.RoutedEventHandler]$addResultsItemClick = {
 

      $syncHash.SearchBox.Tag = ($syncHash.resultsSidePaneGrid.SelectedItem | Select-Object -ExpandProperty SamAccountName) -replace '\$' 
      $syncHash.SearchBox.Focus()
      $wshell = New-Object -ComObject wscript.shell;
      $wshell.SendKeys('{ESCAPE}')
      $syncHash.resultsSidePane.IsOpen = $false

      
   

}

[System.Windows.RoutedEventHandler]$addContextItemClick = {
 
    $i = ($configHash.contextConfig.IDnum | Sort-Object -Descending | Select-Object -First 1) + 1

    $configHash.contextConfig.Add(([PSCustomObject]@{

            IDnum           = $i
            ActionName      = "Name"
            RequireOnline   = $false
            RequireUser     = $false
            Types           = 'a'
            actionCmd       = 'Do-Something -User $user'              
            actionCmdIcon   = $null
            actionCmdMulti  = $false         
            ValidAction     = $false
           
              
        }))
      
    $configHash.boxContextCount = ($configHash.contextConfig | Measure-Object).Count

}

$syncHash.settingUserPropGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $EventonDataGrid)
$syncHash.settingContextPropGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $EventonContextGrid)

$syncHash.settingUserPropGrid.AddHandler([System.Windows.Controls.Combobox]::SelectionChangedEvent, $EventonDataGrid)
$syncHash.settingCompPropGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $EventonDataGrid)
$syncHash.settingObjectToolsPropGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $EventonObjectToolGrid)
$syncHash.settingCompPropGrid.AddHandler([System.Windows.Controls.Combobox]::SelectionChangedEvent, $EventonDataGrid)
$syncHash.settingNetDataGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $eventonNetDataGrid)
$syncHash.settingNetDataGrid.AddHandler([System.Windows.Controls.TextBox]::LostKeyboardFocusEvent, $eventonNetDataGrid)
$syncHash.settingCompAddItemClick.AddHandler([System.Windows.Controls.Button]::ClickEvent, $addItemClick)
$syncHash.settingUserAddItemClick.AddHandler([System.Windows.Controls.Button]::ClickEvent, $addItemClick)
$syncHash.resultsSidePaneGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $addResultsItemClick)


$syncHash.settingObjectToolsAddItemClick.AddHandler([System.Windows.Controls.Button]::ClickEvent, $addObjectToolItemClick)
$syncHash.settingContextAddItemClick.AddHandler([System.Windows.Controls.Button]::ClickEvent, $addContextItemClick)
$syncHash.settingNameDataGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $eventonNameDataGrid)
$syncHash.settingNameDataGrid.AddHandler([System.Windows.Controls.TextBox]::PreviewMouseLeftButtonUpEvent, $eventonNameDataGrid)

$syncHash.settingNameDialogClose.Add_Click( {
        $syncHash.settingNameDialog.IsOpen = $false
        $configHash.nameMapListView.Refresh()

    })

$syncHash.settingRtAddClick.Add_Click( {

        if (!($syncHash.customRt)) {
            $syncHash.customRt = @{}
        }

        if (!$configHash.rtConfig) {
            $configHash.rtConfig = @{}
        }

        $rtID = "rt" + [string]([int](((($configHash.rtConfig.Keys | Where-Object {$_ -like "RT*"}) -replace 'rt') | Sort-Object -Descending | Select-Object -First 1)) + 1) 
       
        $syncHash.customRt.$rtID = @{
            parentDock      = New-Object System.Windows.Controls.DockPanel -Property @{
                Margin              = "0,0,15,0" 
                HorizontalAlignment = "Stretch" 
            }
            childStack      = New-Object System.Windows.Controls.StackPanel
            InfoHeader      = New-Object System.Windows.Controls.Label -Property @{
                Content = "Custom Tool $($rtID -replace 'rt')"
                Style   = $syncHash.Window.FindResource('rtHeader')
            }
            InfoSubheader   = New-Object System.Windows.Controls.TextBlock -Property @{
                Text  = "Custom remote tool $($rtID -replace 'rt')"
                Style = $syncHash.Window.FindResource('rtSubHeader')
            }
            AlertGlyph      = New-Object System.Windows.Controls.Label -Property @{
                Style = $syncHash.Window.FindResource('rtLabel')
            }
            ConfigureButton = New-Object System.Windows.Controls.Button -Property  @{
                Style = $syncHash.Window.FindResource('rtClick')
            }
            DelButton = New-Object System.Windows.Controls.Button -Property  @{
                Style = $syncHash.Window.FindResource('rtClickDel')
            }
        }

        $syncHash.customRt.$rtID.ConfigureButton.Name = $rtID
        $syncHash.customRt.$rtID.DelButton.Name = $rtID + 'del'
        $syncHash.customRt.$rtID.parentDock.AddChild($syncHash.customRt.$rtID.childStack)
        $syncHash.customRt.$rtID.parentDock.AddChild($syncHash.customRt.$rtID.AlertGlyph)
        $syncHash.customRt.$rtID.parentDock.AddChild($syncHash.customRt.$rtID.ConfigureButton)
        $syncHash.customRt.$rtID.parentDock.AddChild($syncHash.customRt.$rtID.DelButton)
        $syncHash.customRt.$rtID.childStack.AddChild($syncHash.customRt.$rtID.InfoHeader)
        $syncHash.customRt.$rtID.childStack.AddChild($syncHash.customRt.$rtID.InfoSubheader)
  
        $syncHash.settingRTPanel.AddChild($syncHash.customRt.$rtID.parentDock)

        $configHash.rtConfig.$rtID = [PSCustomObject]@{
            Name  = "Custom Tool $($rtID -replace 'rt')"
            Path  = $null
            Icon  = $null
            Cmd   = " "
            Types = @()
            RequireOnline = $true
            RequireUser = $false
            DisplayName = 'Tool'
        }
 
        $syncHash.customRt.$rtID.DelButton.Add_Click( {
            param([Parameter(Mandatory)][Object]$sender)
            $rtID = $sender.Name -replace 'del'
            $syncHash.customRt.$rtID.parentDock.Visibility = "Collapsed"
            $syncHash.customRt.$rtID.Clear()
            $configHash.rtConfig.Remove($rtID)
        })    

        $syncHash.customRt.$rtID.ConfigureButton.Add_Click( {
                param([Parameter(Mandatory)][Object]$sender)
                $rtID = $sender.Name
                $syncHash.settingRtExeSelect.Visibility = "Visible"
                $syncHash.settingRtPathSelect.Visibility = "Visible"
                $syncHash.rtSettingRequiresOnline.Visibility = "Visible"
                $syncHash.rtSettingRequiresUser.Visibility = "Visible"
                $syncHash.rtDock.DataContext = $configHash.rtConfig.$rtID
                $syncHash.settingRemoteFlyout.isOpen = $true
            })

    })

$syncHash.SearchBoxButton.Add_Click({

    if ($syncHash.SearchBox.Text.Length -eq 0) {

    $searchVal = (Select-ADObject -Type UsersComputers).FetchedAttributes -replace '$'
    
    if ($searchVal) {
        
        $syncHash.SearchBox.Tag = $searchVal
        $syncHash.SearchBox.Focus()
        $wshell = New-Object -ComObject wscript.shell;
        $wshell.SendKeys('{ESCAPE}')
       
    }

    }
    else {
        $syncHash.SearchBox.Clear()
    }

})
$syncHash.userQueryItem.Add_Click({

  $searchVal = $syncHash.userCompGrid.SelectedItem.HostName
  if ($searchVal) {
        
        $syncHash.SearchBox.Tag = $searchVal
        $syncHash.SearchBox.Focus()
        $wshell = New-Object -ComObject wscript.shell;
        $wshell.SendKeys('{ESCAPE}')
        
    }

    

})

$syncHash.compQueryItem.Add_Click({

  $searchVal = $syncHash.compUserGrid.SelectedItem.UserName
  if ($searchVal) {
        
        $syncHash.SearchBox.Tag = $searchVal
        $syncHash.SearchBox.Focus()
        $wshell = New-Object -ComObject wscript.shell;
        $wshell.SendKeys('{ESCAPE}')
      

    }

    

})

$syncHash.itemRefresh.Add_Click({

 $configHash.itemRefreshing = $true
 $searchVal = $configHash.currentTabItem
  if ($searchVal) {
        $syncHash.tabControl.ItemsSource.RemoveAt($syncHash.tabControl.SelectedIndex)
        $syncHash.SearchBox.Tag = $searchVal
        $syncHash.SearchBox.Focus()
        $wshell = New-Object -ComObject wscript.shell;
        $wshell.SendKeys('{ESCAPE}')
      

    }

})

$syncHash.sidePaneExit.Add_Click({
    $syncHash.resultsSidePane.IsOpen = $false
    

})
$syncHash.resultsSidePane.Add_ClosingFinished({

    $syncHash.resultsSidePaneGrid.ItemsSource = $null
})



 
 

$syncHash.Window.ShowDialog() | Out-Null

