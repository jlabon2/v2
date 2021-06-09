# V2 Configurable HD ToolKit #
##############################
$ver = 0.91
$ConfirmPreference = 'None'

Add-Type -AssemblyName 'System.Windows.Forms'

if (!($global:baseConfigPath = (Split-Path $PSCommandPath))) {$global:baseConfigPath = 'C:\V2'}
$modList = @((Join-Path $baseConfigPath -ChildPath '\func\func.psm1'), (Join-Path $baseConfigPath -ChildPath '\internal\internal.psm1'))
Import-Module $modList -Force -DisableNameChecking

# Set window visibility (if not in ISE, set key values as global variables referenced elsewhere
Set-WindowVisibility 
Set-GlobalVars -BasePath $baseConfigPath

# Generate hash tables used throughout tool
New-HashTables

# Generate info pane text
Set-InfoPaneHash

# Import from JSON and add to hash table
Set-Config -ConfigPath $savedConfig -Type Import -ConfigHash $configHash

# Process loaded data or creates initial item templates for various config items and datagrids
@('userPropList', 'compPropList', 'contextConfig', 'objectToolConfig', 'nameMapList', 'netMapList', 'varListConfig', 'searchBaseConfig', 'queryDefConfig', 'modConfig') | Set-InitialValues -ConfigHash $configHash -PullDefaults
@('userLogMapping', 'compLogMapping', 'settingHeaderConfig', 'SACats') | Set-InitialValues -ConfigHash $configHash

$configHash.configVer = Set-DefaultVersionInfo -ConfigHash $configHash 

# Matches config'd user/comp logins with default headers, creates new headers from custom values
$defaultList = @('User', 'DateRaw', 'LoginDc', 'ClientName', 'ComputerName', 'Ignore', 'Custom')
@('userLogMapping', 'compLogMapping') | Set-LoggingStructure -DefaultList $defaultList -ConfigHash $configHash
Set-ActionLog -ConfigHash $configHash

$configHash.modConfig.modPath | ForEach-Object -Process { $modList += $_ }
$configHash.modList = $modList | Where-Object -FilterScript { [string]::IsNullOrWhiteSpace($_) -eq $false }

# Add default values if they are missing
@('MSRA', 'MSTSC') | Set-RTDefaults -ConfigHash $configHash

# Load required DLLs
foreach ($dll in ((Get-ChildItem -Path (Join-Path $baseConfigPath lib) -Filter *.dll).FullName)) { $null = [System.Reflection.Assembly]::LoadFrom($dll) }

# Read xaml and load wpf controls into synchash (named synchash)
Set-WPFControls -TargetHash $syncHash -XAMLPath $xamlPath

Get-Glyphs -ConfigHash $configHash -GlyphList $glyphList
$syncHash.settingLogo.Source = Join-Path $baseConfigPath trident.png

# builds custom WPF controls from whatever was defined and saved in ConfigHash
Add-CustomItemBoxControls -SyncHash $syncHash -ConfigHash $configHash
Add-CustomToolControls -SyncHash $syncHash -ConfigHash $configHash
Add-CustomRTControls -SyncHash $syncHash -ConfigHash $configHash

$syncHash.externalToolList.ItemsSource = Set-ExternalTools -ConfigHash $configHash -BaseConfigPath $baseConfigPath

Set-ConfiguredDomainName -SyncHash $syncHash -DomainName $configHash.configuredDomain
Set-Version -Version "v$ver" -CID $configHash.configVer.Ver -SyncHash $syncHash
$sysCheckHash.missingCheckFail = $false

#region Item Tool Events
#region Item Tools - Grid
$syncHash.ItemToolGridADSelectionButton.Add_Click({ Start-CustomItemSelection -SyncHash $syncHash -ConfigHash $configHash -Control Grid })
  
$syncHash.itemToolGridItemsGrid.Add_SelectionChanged( {
    if ($syncHash.itemToolGridItemsGrid.SelectedItem.Image) { $syncHash.itemToolImageSource.Source = [byte[]]($syncHash.itemToolGridItemsGrid.SelectedItem.Image) }
})

$syncHash.itemToolCustomConfirm.Add_Click({
    $configHash.customDialogClosed = $true
    if ($syncHash.itemToolCustomContent.Tag  -eq 'Choice') { $configHash.customInput = $syncHash.itemToolCustomContentChoice.SelectedValue }
    else { $configHash.customInput = $syncHash.itemToolCustomContent.Text }
})

$syncHash.itemToolCustomContent.add_KeyDown({ 
    if ($_.Key -eq 'Enter') {
        $configHash.customDialogClosed = $true
        $configHash.customInput = $syncHash.itemToolCustomContent.Text
    } 
})

$syncHash.itemToolCustomContent.Add_TextChanged({
    if (![string]::IsNullOrEmpty($syncHash.itemToolCustomContent.Text)) {
        if ($syncHash.itemToolCustomContent.Tag -eq 'int') {
            try {
                [int]$syncHash.itemToolCustomContent.Text 
                 $syncHash.itemToolCustomConfirm.IsEnabled = $true  
            }
            catch { $syncHash.itemToolCustomConfirm.IsEnabled = $false }
        }
        else {
            try {
             [string]$syncHash.itemToolCustomContent.Text
              $syncHash.itemToolCustomConfirm.IsEnabled = $true  
            }
            catch {  $syncHash.itemToolCustomConfirm.IsEnabled = $true }
        }
    }
    else { $syncHash.itemToolCustomConfirm.IsEnabled = $false }
})

$syncHash.itemToolGridItemsGrid.Add_AutoGeneratingColumn( {

    if ($syncHash.itemToolGridItemsGrid.Columns.Count -eq 1) {
        $syncHash.itemToolGridItemsGrid.Width = 160
    }

    $_.Column.CanUserSort = $true
    $_.Column.Width = 150
    
    if ($syncHash.itemToolGridItemsGrid.Width -lt 800) { $syncHash.itemToolGridItemsGrid.Width  = $syncHash.itemToolGridItemsGrid.Width + 160 }

    if ($_.Column.Header -eq 'Image') {
        $_.Cancel = $true 
        $syncHash.itemToolImageBorder.Visibility = 'Visible'
    }
   


})

$syncHash.itemToolGridSearchBox.Add_TextChanged( {
    $syncHash.itemToolGridItemsGrid.ItemsSource.Filter = $null
    $syncHash.itemToolGridItemsGrid.ItemsSource.Filter = { param ($item) $item -match $syncHash.itemToolGridSearchBox.Text }
    if (!$syncHash.itemToolGridItemsGrid.HasItems) {  $syncHash.itemToolGridItemsEmptyText.Visibility = 'Visible' }
    else  {  $syncHash.itemToolGridItemsEmptyText.Visibility = 'Hidden'}
 
})

$syncHash.itemToolGridSelectConfirmCancel.Add_Click( { $syncHash.itemToolDialog.IsOpen = $false })

$syncHash.itemToolGridSelectAllButton.Add_Click( {    
        if ($syncHash.itemToolGridItemsGrid.Items.Count -eq $syncHash.itemToolGridItemsGrid.SelectedItems.Count) { $syncHash.itemToolGridItemsGrid.UnselectAll() } 
        else { $syncHash.itemToolGridItemsGrid.SelectAll() }      
    })

#endregion

#region Item Tools - List
$syncHash.ItemToolADSelectionButton.Add_Click({ Start-CustomItemSelection -SyncHash $syncHash -ConfigHash $configHash -Control ListBox })



$syncHash.itemToolListSearchBox.Add_TextChanged( {
        $syncHash.itemToolListSelectListBox.ItemsSource.Filter = $null
        $syncHash.itemToolListSelectListBox.ItemsSource.Filter = { param ($item) $item -match $syncHash.itemToolListSearchBox.Text }
        if (!$syncHash.itemToolListSelectListBox.HasItems) {  $syncHash.itemToolListItemsEmptyText.Visibility = 'Visible' }
        else  {  $syncHash.itemToolListItemsEmptyText.Visibility = 'Hidden'}
    })

$syncHash.itemToolListSelectConfirmButton.Add_Click( {
        $itemList = $syncHash.itemToolListSelectListBox.SelectedItems.Name
        Start-ItemToolAction -ConfigHash $configHash -SyncHash $syncHash -Control List -ItemList $itemList
    })

$syncHash.itemToolGridSelectConfirmButton.Add_Click( {
        $itemList = $syncHash.itemToolGridItemsGrid.SelectedItems
        Start-ItemToolAction -ConfigHash $configHash -SyncHash $syncHash -Control Grid -ItemList $itemList
    })

$syncHash.itemToolListSelectConfirmCancel.Add_Click( { $syncHash.itemToolDialog.IsOpen = $false })

$syncHash.itemToolListSelectAllButton.Add_Click( {
        if ($syncHash.itemToolListSelectListBox.Items.Count -eq $syncHash.itemToolListSelectListBox.SelectedItems.Count) { $syncHash.itemToolListSelectListBox.UnselectAll() } 
        else { $syncHash.itemToolListSelectListBox.SelectAll() }      
    })
#endregion
#endregion

  
#region ChildWindow opening events
$syncHash.settingRemoteClick.add_Click( { Set-ChildWindow -Panel settingRTContent -Title 'Configure Remote Connection Clients' -SyncHash $syncHash })

$syncHash.settingNetworkClick.add_Click( { Set-ChildWindow -Panel settingNetContent -Title 'Networking Mappings' -SyncHash $syncHash -Height 275 })

$syncHash.settingNamingClick.add_Click( { Set-ChildWindow -Panel settingNameContent -Title 'Computer Categorization' -SyncHash $syncHash -Height 275 })

$syncHash.settingVarClick.add_Click( { Set-ChildWindow -Panel settingVarContent -Title 'Resources and Variables' -SyncHash $syncHash -Height 275 -Width 530 })

$syncHash.settingGeneralClick.add_Click( { Set-ChildWindow -Panel settingGeneralContent -Title 'General Settings' -SyncHash $syncHash -Height 370 -Width 530 })

$syncHash.settingUserPropClick.add_Click( { Set-ChildWindow -Panel settingUserPropContent -Title 'User Property Mappings' -SyncHash $syncHash -Height 365 -Width 655 })

$syncHash.settingCompPropClick.add_Click( { Set-ChildWindow -Panel settingCompPropContent -Title 'Computer Property Mappings' -SyncHash $syncHash -Height 365 -Width 655 })

$syncHash.settingObjectToolsClick.add_Click( { Set-ChildWindow -Panel settingItemToolsContent -Title 'Tools' -SyncHash $syncHash -Height 365 -Width 655 })

$syncHash.settingContextClick.add_Click( { Set-ChildWindow -Panel settingContextPropContent -Title 'Contextual Actions Mappings'-SyncHash $syncHash -Height 365 -Width 400 })

$syncHash.settingLoggingClick.add_Click( {
        if (!($configHash.compLogPath) -or !(Test-Path -Path $configHash.compLogPath)) { $syncHash.compLogPopupButton.IsEnabled = $false }
        if (!($configHash.userLogPath) -or !(Test-Path -Path $configHash.userLogPath)) { $syncHash.userLogPopupButton.IsEnabled = $false }

        Set-ChildWindow -Panel settingLoggingContent -Title 'Login Log Paths' -SyncHash $syncHash -Height 300 -Width 480
    })

#endregion

#region settingload / systems check
$syncHash.settingLogo.add_Loaded( {
        $syncHash.Window.Activate()
        $configRoot = Split-Path $savedConfig 
        $sysCheckHash.sysChecks = New-Object -TypeName System.Collections.ObjectModel.ObservableCollection[Object]
    
        $sysCheckHash.sysChecks.Add([PSCustomObject]@{
                'ADModule'                   = 'False'
                'RSModule'                   = 'False'
                'ADMember'                   = 'False'
                'ADDCConnectivity'           = 'False'
                'IsInAdmin'                  = 'False'
                'IsDomainAdmin'              = 'False'
                'IsDomainAdminOrDelegated'   = 'False'
                'IsDelegated'                = 'False'
                'IsReport'                   = 'False'
                'Admin'                      = 'False'
                'Modules'                    = 'False'
                'ADDS'                       = 'False'
                'DelegatedGroupName'         =  if (Test-Path (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-delegatedGroup.json")) { (Get-Content (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-delegatedGroup.json") | ConvertFrom-Json).Name }
                                                else {$null}
                'ReportGroupName'            =  if (Test-Path (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-reportGroup.json")) { (Get-Content (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-reportGroup.json") | ConvertFrom-Json).Name }
                                                else {$null}
            })

        $modList = Get-Module -ListAvailable ActiveDirectory, PoshRSJob | Sort-Object -Unique

        if ($modList.Name -contains 'ActiveDirectory') { $sysCheckHash.sysChecks[0].ADModule = 'True' }

        if ($modList.Name -contains 'PoshRSJob') {
            $sysCheckHash.sysChecks[0].RSModule = 'True'

            $rsArgs = @{
                Name            = 'init'
                ArgumentList    = @($syncHash, $sysCheckHash, $configHash, $savedConfig, $syncHash.settingADRegenText)
                ModulesToImport = $configHash.modList
            }

            Start-RSJob @rsArgs -ScriptBlock {        
                Param($syncHash, $sysCheckHash, $configHash, $savedConfig, $adLabel)
             
                Start-BasicADCheck -SysCheckHash $sysCheckHash -configHash $configHash
                
                Start-AdminCheck -SysCheckHash $sysCheckHash -ConfiguredDomain $configHash.ConfiguredDomain
                # Check individual checks; mark parent categories as true is children are true       
                switch ($sysCheckHash.sysChecks) {
                    { $_.ADModule -eq $true -and $_.RSModule -eq $true } { $sysCheckHash.sysChecks[0].Modules = 'True' }
                    { $_.ADMember -eq $true -and $_.ADDCConnectivity -eq $true } { $sysCheckHash.sysChecks[0].ADDS = 'True' }
                    { $_.IsInAdmin -eq $true -and $_.IsDomainAdminOrDelegated -eq $true } { $sysCheckHash.sysChecks[0].Admin = 'True' }
                }

                @('settingADMemberLabel', 'settingADDCLabel', 'settingModADLabel', 'settingModRSLabel', 'settingDomainAdminLabel', 
                    'settingLocalAdminLabel', 'settingPermLabel', 'settingADLabel', 'settingModLabel', 'settingDelegatedPanel', 'settingDelegatedGroupSelection', 'settingReportGroupSelection') | 
                    Set-RSDataContext -SyncHash $syncHash -DataContext $sysCheckHash.sysChecks

                  
               
                $sysCheckHash.checkComplete = $true

                Start-Sleep -Seconds 1

                if (($sysCheckHash.sysChecks[0].ADDS -eq $false -or 
                    $sysCheckHash.sysChecks[0].Modules -eq $false -or 
                    $sysCheckHash.sysChecks[0].Admin -eq $false) -and
                    $sysCheckHash.sysChecks[0].IsReport -eq $false) { Suspend-FailedItems -SyncHash $syncHash -CheckedItems SysCheck}
                
                elseif ($sysCheckHash.sysChecks[0].IsReport -eq $true) {
                    Set-ReportView -SyncHash $syncHash -ConfigHash $configHash
                }

                else { 
                    Set-DefaultDC -ConfigHash $configHash -Domain $configHash.configuredDomain
                    Start-PropBoxPopulate -configHash $configHash -Window $syncHash.Window -AdLabel $adLabel -SavedConfig $savedConfig
                    Set-ADGenericQueryNames -ConfigHash $configHash               
                    Set-QueryPropertyList -SyncHash $syncHash -ConfigHash $configHash

                    if ($syncHash.settingStatusChildBoard.Visibility -eq 'Collapsed') {
                        $syncHash.Window.Dispatcher.invoke([action] { 
                                $syncHash.settingStatusChildBoard.Visibility = 'Visible'
                                $syncHash.settingConfigPanel.Visibility = 'Visible'
                            })
                    }              
                }  
                
                Show-WPFWindow -SyncHash $syncHash -ConfigHash $configHash
            }
        }
        else {                                   
            Suspend-FailedItems -SyncHash $syncHash -CheckedItems RSCheck
            Show-WPFWindow -SyncHash $syncHash
        }
    })

$syncHash.settingDelegatedGroupPick.Add_Click({
    $configRoot = Split-Path $savedConfig 
    $groupSel = Select-ADObject -Type Groups
    
    if ($groupSel -ne 'Cancel') {
        $sysCheckHash.sysChecks[0].DelegatedGroupName = $groupSel.Path.Split('/')[2] + '\' + $groupSel.Name
        $syncHash.settingDelegatedGroupSelection.Text = $sysCheckHash.sysChecks[0].DelegatedGroupName
        $sysCheckHash.sysChecks[0].DelegatedGroupName | Select-Object @{Label = "Name"; Expression = {$_}} | 
            ConvertTo-Json | Out-File (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-delegatedGroup.json")
    }
})

$syncHash.settingReportGroupPick.Add_Click({
    $configRoot = Split-Path $savedConfig 
    $groupSel = Select-ADObject -Type Groups

    if ($groupSel -ne 'Cancel') {
    $sysCheckHash.sysChecks[0].ReportGroupName = $groupSel.Path.Split('/')[2] + '\' + $groupSel.Name
    $syncHash.settingReportGroupSelection.Text = $sysCheckHash.sysChecks[0].ReportGroupName
    $sysCheckHash.sysChecks[0].ReportGroupName | Select-Object @{Label = "Name"; Expression = {$_}} | 
        ConvertTo-Json | Out-File (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-reportGroup.json")
    }
})
#endregion 

#region RemoteTools events

$syncHash.settingRtRDPClick.Add_Click( { Set-StaticRTContent -SyncHash $syncHash -ConfigHash $configHash -Tool MSTSC })

$syncHash.settingRtMSRAClick.Add_Click( { Set-StaticRTContent -SyncHash $syncHash -ConfigHash $configHash -Tool MSRA })

$syncHash.settingRemoteFlyout.Add_OpeningFinished({ 
    $syncHash.settingChildWindow.ShowCloseButton = $false
    Get-RTFlyoutContent -ConfigHash $configHash -SyncHash $syncHash 
})

$syncHash.settingRemoteFlyoutExit.Add_Click( {
        $syncHash.settingChildWindow.ShowCloseButton = $true
        Reset-ChildWindow -SyncHash $syncHash -Title 'Configure Remote Connection Clients' -SkipContentPaneReset -SkipResize
        Set-SelectedRTTypes -SyncHash $syncHash -ConfigHash $configHash
        Set-CurrentPane -SyncHash $syncHash -Panel 'settingRTContent'
    })


$syncHash.settingRtExeSelect.Add_Click( { Get-RTExePath -SyncHash $syncHash -ConfigHash $configHash })

$syncHash.settingRtAddClick.Add_Click( {
        if (!($syncHash.customRt)) { $syncHash.customRt = @{} }
        if (!$configHash.rtConfig) { $configHash.rtConfig = @{} }

        $rtID = 'rt' + [string]([int](((($configHash.rtConfig.Keys |
            Where-Object -FilterScript { $_ -like 'RT*' }) -replace 'rt') |
                Sort-Object -Descending |
                        Select-Object -First 1)) + 1) 
       
        New-CustomRTConfigControls -SyncHash $syncHash -ConfigHash $configHash -RTID $rtID -NewTool
    })

#endregion

#region NetworkMap events

$syncHash.settingNetMapClick.Add_Click( {

        if ($syncHash.settingNetImportClick.ItemsSource -eq $null) {
            $syncHash.settingNetImportClick.ItemsSource = @('Current device address','ADDS subnets','DHCP scopes')
            $syncHash.settingNetImportClick.SelectedIndex = 0
        }

        $syncHash.settingNetDataGrid.Visibility = 'Visible'
        $syncHash.settingNetDataGrid.ItemsSource = $configHash.netMapList 
        Set-ChildWindow -SyncHash $syncHash -Title 'Network Mappings' -HideCloseButton -Background Flyout
        $syncHash.settingNetFlyout.IsOpen = $true
    })

$syncHash.settingNetAddClick.Add_Click( { Set-NetworkMapItem -SyncHash $syncHash -ConfigHash $configHash })

$syncHash.settingNetImportClick.Add_Click( { Set-NetworkMapItem -SyncHash $syncHash -ConfigHash $configHash -Import })

$syncHash.settingNetFlyoutExit.Add_Click( {
        $syncHash.settingNetFlyout.IsOpen = $false
        Set-ChildWindow -SyncHash $syncHash -Background Standard -Title 'Networking Mappings'        
       
        [Array]$configHash.netMapList | ForEach-Object -Process {
            if ($_.ValidMask -ne $true -and $_.ValidNetwork -ne $true) { [Array]$configHash.netMapList.RemoveAt([Array]::IndexOf($configHash.netMapList.ID, $_.ID)) }
        }

        $configHash.Remove('NetMapListView')
    })


#endregion



#region UserLog events
$syncHash.userLogPopupButton.Add_Click( { Set-LogMapGrid -SyncHash $syncHash -ConfigHash $configHash -Type User })

$syncHash.compLogPopupButton.Add_Click( { Set-LogMapGrid -SyncHash $syncHash -ConfigHash $configHash -Type Comp })
#endregion

#region Naming events
$syncHash.settingNameMapClick.Add_Click( {
        if ($configHash.nameMapList.GetType().Name -ne 'ListCollectionView') {
            $configHash.nameMapListView = [System.Windows.Data.ListCollectionView]$configHash.nameMapList 
            $configHash.nameMapListView.IsLiveSorting = $true
          #  $configHash.nameMapListView.SortDescriptions.Add((New-Object -TypeName System.ComponentModel.SortDescription -Property @{ PropertyName = 'ID' }))        
            $syncHash.settingNameDataGrid.ItemsSource = $configHash.nameMapListView         
            $configHash.nameMapListView.Refresh()
        }

        $configHash.nameMapListView.LiveSortingProperties.Add('Id')   
        $configHash.nameMapListView.SortDescriptions.Add((New-Object -TypeName System.ComponentModel.SortDescription -ArgumentList ('Id', 'Descending')))
        $syncHash.settingNameDataGrid.Visibility = 'Visible'
        
        Set-ChildWindow -SyncHash $syncHash -Title 'Computer Categorization Rules' -HideCloseButton -Background Flyout

        $syncHash.settingNameFlyout.IsOpen = $true
    })

$syncHash.settingNameAddClick.Add_Click( {
        if (($configHash.nameMapList | Measure-Object).Count -gt 1) {
            ($configHash.nameMapList |
                    Sort-Object -Property ID -Descending |
                        Select-Object -First 1).TopPos = $false
        }

        $configHash.nameMapList.Add([PSCustomObject]@{
                Id        = ($configHash.nameMapList.ID |
                        Sort-Object -Descending |
                            Select-Object -First 1) + 1
                Name      = $null
                Condition = $null
                topPos    = $true
            })    

        $syncHash.settingNameDataGrid.Items.Refresh()
    })

$syncHash.settingCommandGridAddClick.Add_Click( {
        $configHash.objectToolConfig[([Array]::IndexOf($configHash.objectToolConfig, $syncHash.settingObjectToolsPropGrid.SelectedItem))].toolCommandGridConfig.Add([PSCustomObject]@{
                ToolID         = ($syncHash.settingObjectToolsPropGrid.SelectedItem.toolCommandGridConfig.toolID |
                        Sort-Object -Descending |
                            Select-Object -First 1) + 1
                SetName        = $null
                queryCmd       = 'Check-Something -UserName $user'
                actionCmd      = 'Do-Something -UserName $user'
                actionCmdValid = 'null'
                queryCmdValid  = 'null'
            })
 
        $syncHash.settingCommandGridDataGrid.Items.Refresh()
        $syncHash.SettingToolFlyoutScroller.ScrollToBottom()
    })



$syncHash.settingGeneralAddClick.Add_Click( {
        if ($syncHash.settingGeneralAddClick.Tag -eq 'OU') {
            $configHash.searchBaseConfig.Add([PSCustomObject]@{
                    OUNum          = ($configHash.searchBaseConfig.OUNum |
                            Sort-Object -Descending |
                                Select-Object -First 1) + 1
                    OU             = $null
                    QueryScopeList = @('OneLevel', 'Subtree')
                    QueryScope     = 'Subtree'
                })    

            $syncHash.settingOUDataGrid.Items.Refresh()
        }
    
        else {
            $configHash.queryDefConfig.Add([PSCustomObject]@{
                    ID               = ($configHash.queryDefConfig.ID |
                            Sort-Object -Descending |
                                Select-Object -First 1) + 1
                    Name             = $null
                    QueryDefTypeList = $configHash.adPropertyMap.Keys | Sort-Object
                    QueryDefType     = $null
                })    

            $syncHash.settingQueryDefDataGrid.Items.Refresh()
        }
    })


$syncHash.settingVarAddClick.Add_Click( {
        if ($syncHash.settingVarAddClick.Tag -eq 'Var') {
            $configHash.varListConfig.Add([PSCustomObject]@{
                    VarNum              = ($configHash.varListConfig.VarNum |
                            Sort-Object -Descending |
                                Select-Object -First 1) + 1
                    VarName             = $null
                    VarCmd              = $null
                    UpdateFrequencyList = @('All Queries', 'User Queries', 'Comp Queries', 'Daily', 'Hourly', 'Every 15 mins', 'Program Start')
                    UpdateFrequency     = 'Program Start'
                    VarDesc             = $null
                })    

            $syncHash.settingVarDataGrid.Items.Refresh()
        }
    
        else {
            $configHash.modConfig.Add([PSCustomObject]@{
                    ModNum  = ($configHash.modConfig.ModNum |
                            Sort-Object -Descending |
                                Select-Object -First 1) + 1
                    ModName = $null
                    ModPath = $null
                })    

            $syncHash.settingModDataGrid.Items.Refresh()
        }
    })


$syncHash.settingVarDialogClose.Add_Click( {
        $syncHash.settingVarDialog.IsOpen = $false      
        $syncHash.settingVarDataGrid.Items.Refresh()
    })

$syncHash.settingCommandGridDialogClose.Add_Click( {
        $syncHash.settingCommandGridDialog.IsOpen = $false    
        $syncHash.settingCommandGridDataGrid.Items.Refresh()
    })

#endregion

#region ContextTools

#endregion

#region VarConfig
$syncHash.settingVarMapClick.Add_Click( { 
        Set-ChildWindow -SyncHash $syncHash -Title 'Variable Definitions' -Panel settingVarDataGrid -HideCloseButton -Background Flyout
        $syncHash.settingVarFlyout.IsOpen = $true
    })

$syncHash.settingModMapClick.Add_Click( { 
        Set-ChildWindow -SyncHash $syncHash -Title 'Module Definitions' -Panel settingModDataGrid -HideCloseButton -Background Flyout
        $syncHash.settingVarFlyout.IsOpen = $true
    })



$syncHash.settingOUMapClick.Add_Click( { 
        Set-ChildWindow -SyncHash $syncHash -Title 'Search Base Definitions' -Panel settingOUDataGrid -HideCloseButton -Background Flyout
        $syncHash.settingGeneralFlyout.IsOpen = $true
    })



$syncHash.settingQueryMapClick.Add_Click( { 
        Set-ChildWindow -SyncHash $syncHash -Title 'Query Definitions' -Panel settingQueryDefDataGrid -HideCloseButton -Background Flyout
        $syncHash.settingGeneralFlyout.IsOpen = $true
    })

$syncHash.settingMiscClick.Add_Click( {
        Set-ChildWindow -SyncHash $syncHash -Title 'Misc. Settings' -Panel settingMiscGrid -HideCloseButton -Background Flyout
    
        if (!$configHash.actionlogPath) { $configHash.actionlogPath = 'C:\Logs' }
        $syncHash.settingLogPath.Text = $configHash.actionlogPath 

        if ($configHash.MinHeight -and $configHash.MinHeight -is [int]) { $SyncHash.MinWindowHeight.Text = $configHash.MinHeight }
        else { $SyncHash.MinWindowHeight.Text = 700 }

        if ($configHash.MinWidth -and $configHash.MinWidth -is [int]) { $SyncHash.MinWindowWidth.Text = $configHash.MinWidth }
        else {  $SyncHash.MinWindowWidth.Text = 1000 }

        if (!$configHash.settingHeaderConfig) { 
            $configHash.settingHeaderConfig = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
            $configHash.settingHeaderConfig.Add([PSCustomObject]@{
                headerAdd              = $null
                headerColor            = '#FFFFFF'
                headerUser             = $false
            })
        }

        $syncHash.settingHeaderDef.DataContext = $configHash.settingHeaderConfig

        $syncHash.settingGeneralFlyout.IsOpen = $true
    })

#endregion
$syncHash.settingSearchDaySpan.Add_ValueChanged({$ConfigHash.searchDays = $syncHash.settingSearchDaySpan.Value })


#region FlyOutExits
$syncHash.settingLoggingCompFlyoutExit.Add_Click( { Reset-ChildWindow -SyncHash $syncHash -Title 'Login Log Paths' -SkipContentPaneReset -SkipResize })

$syncHash.settingResourceFlyoutExit.Add_Click( {
    Reset-ChildWindow -SyncHash $syncHash -Title 'Resources and Variables' -SkipContentPaneReset -SkipResize
    Set-CurrentPane -SyncHash $syncHash -Panel 'settingVarContent'
})

$syncHash.settingLoggingUserFlyoutExit.Add_Click( { Reset-ChildWindow -SyncHash $syncHash -Title 'Login Log Paths' -SkipContentPaneReset -SkipResize })

$syncHash.settingObjectToolFlyoutExit.Add_Click( { 
    Reset-ChildWindow -SyncHash $syncHash -Title 'Object Tools Mappings' -SkipContentPaneReset -SkipResize -SkipDataGridReset
    Set-CurrentPane -SyncHash $syncHash -Panel 'settingItemToolsContent' 
})

$syncHash.settingNameFlyoutExit.Add_Click( { Reset-ChildWindow -SyncHash $syncHash -Title 'Computer Categorization' -SkipContentPaneReset -SkipResize })

$syncHash.settingGeneralFlyoutExit.Add_Click( {
        if ($syncHash.settingGeneralAddClick.Tag -eq 'null' -and (Test-Path -Path $syncHash.settingLogPath.Text)) { $configHash.actionlogPath = $syncHash.settingLogPath.Text }

        try { $configHash.MinHeight = [int]$SyncHash.MinWindowHeight.Text }
        catch { $configHash.MinHeight = 700 }

        try { $configHash.MinWidth = [int]$SyncHash.MinWindowWidth.Text }
        catch { $configHash.MinWidth = 1000 }
            
        Reset-ChildWindow -SyncHash $syncHash -Title 'General Settings' -SkipContentPaneReset -SkipResize    
        Set-CurrentPane -SyncHash $syncHash -Panel 'settingGeneralContent'
    })

$syncHash.settingContextDefFlyout.Add_OpeningFinished( {
        $syncHash.settingContextListTypes.ItemsSource = $configHash.nameMapList
    
        $syncHash.settingContextListTypes.Items |
            Where-Object -FilterScript { $_.Name -in $configHash.contextConfig[$syncHash.settingContextPropGrid.SelectedItem.IDNUM - 1].Types } |
                ForEach-Object -Process { $syncHash.settingContextListTypes.SelectedItems.Add(($_)) }
    })

$syncHash.settingContextFlyoutExit.Add_Click( {
        $configHash.contextConfig[$syncHash.settingContextPropGrid.SelectedItem.IDNUM - 1].Types = @()
    
        $syncHash.settingContextListTypes.SelectedItems | 
            ForEach-Object -Process { $configHash.contextConfig[$syncHash.settingContextPropGrid.SelectedItem.IDNUM - 1].Types += $_.Name } 

        Reset-ChildWindow -SyncHash $syncHash -Title 'Contextual Action Buttons' -SkipContentPaneReset -SkipResize -SkipDataGridReset
        Set-CurrentPane -SyncHash $syncHash -Panel 'settingContextPropContent'

        $syncHash.settingRemoteListTypes.SelectedItems.Clear()
    })

$syncHash.settingFlyoutExit.Add_Click( {
    if ($syncHash.settingUserPropContent.Visibility -eq 'Visible') {
        $type = 'User'
        Set-CurrentPane -SyncHash $syncHash -Panel 'settingUserPropContent'
    }
    else { 
        $type = 'Computer' 
        Set-CurrentPane -SyncHash $syncHash -Panel 'settingCompPropContent'
    }

    Reset-ChildWindow -SyncHash $syncHash -Title "$type Property Mappings" -SkipContentPaneReset -SkipResize -SkipDataGridReset
})


#endregion

$syncHash.settingActionPathClick.Add_Click( { $syncHash.settingLogPath.Text = New-FolderSelection -Title "Select tool's logging directory" })
        
$syncHash.settingLoggingPcPathClick.Add_Click( { Set-LoggingDirectory -SyncHash $syncHash -ConfigHash $configHash -Type Comp })

$syncHash.settingLoggingUserPathClick.Add_Click( { Set-LoggingDirectory -SyncHash $syncHash -ConfigHash $configHash -Type User })
    
#endregion

$syncHash.settingChildWindow.add_ClosingFinished( { Reset-ChildWindow -SyncHash $syncHash })

#region Category button events
$syncHash.settingModClick.add_Click( { Set-ChildWindow -SyncHash $syncHash -Panel settingModContent -Title 'Required PS Modules' })

$syncHash.settingADClick.add_Click( { Set-ChildWindow -SyncHash $syncHash -Panel settingADContent -Title 'ADDS' })

$syncHash.settingPermClick.add_Click( { Set-ChildWindow -SyncHash $syncHash -Panel settingAdminContent -Title 'Admin Permissions' -Height 230 }) 

#endregion Category button events
#region Action pane exit events

$syncHash.settingCloseClick.add_Click( { $syncHash.Window.Close() })
    
$syncHash.settingConfigCancelClick.add_Click( { $syncHash.Window.Close() })

$syncHash.settingImportClick.add_Click( { Import-Config -SyncHash $syncHash -ConfigMap $configMap })

$syncHash.headerConfigUpdate.Add_Click({
    Import-Config -SyncHash $syncHash -ConfigMap $configMap -ConfigSelection (Join-Path -Path $configHash.configVer.configPublishPath -ChildPath 'config.json')
})

$synchash.importConfirmButton.add_Click({ Start-Import -ImportItems $importItems -SelectedItems $syncHash.importListBox.SelectedItems -ConfigMap $configMap -ConfigHash $configHash -savedConfig $savedConfig -BaseConfigPath $baseConfigPath -Monitor $syncHash.importMonitor.IsChecked})

$syncHash.importSelectAllButton.add_Click({  
    if ($syncHash.importListBox.Items.Count -eq $syncHash.importListBox.SelectedItems.Count) { $syncHash.importListBox.UnselectAll() } 
    else { $syncHash.importListBox.SelectAll() }
})   

$syncHash.importCancel.add_Click({ $syncHash.importDialog.IsOpen = $false })

$syncHash.settingConfigClick.add_Click({ 
    
   $syncHash.saveChangeLogEdit.Text =  $configHash.configVer.changeLog
    
    if ( $configHash.configVer.configPublishPath -and (Test-Path  $configHash.configVer.configPublishPath)) {
        $syncHash.savePublishPath.Text =  $configHash.configVer.configPublishPath
    }

    $syncHash.SaveDialog.IsOpen = $true })
      

$syncHash.saveConfirmClick.add_Click({
    
   $configHash.configVer.changeLog = $syncHash.saveChangeLogEdit.Text
    
    if ($syncHash.saveReset.IsChecked) {
        $configHash.configVer.Ver = 1
        $configHash.configVer.ID  = ([guid]::NewGuid()).Guid
    }

    if ($syncHash.saveIncrement.IsChecked -and !($syncHash.saveReset.IsChecked)) { $configHash.configVer.Ver = $configHash.configVer.Ver + 1 }


    if ($syncHash.SavePublish.IsChecked -and (Test-Path $syncHash.savePublishPath.Text)) {
        $configHash.configVer.configPublishPath = $syncHash.savePublishPath.Text 
        Export-VersionConfig -ConfigPath (Join-Path -Path $configHash.configVer.configPublishPath -ChildPath 'configVer.json') -ConfigHash $configHash
        Set-Config -ConfigPath  (Join-Path -Path $configHash.configVer.configPublishPath -ChildPath 'config.json') -Type Export -ConfigHash $configHash
          
    }
      
    Set-Config -ConfigPath $savedConfig -Type Export -ConfigHash $configHash

    Start-Sleep -Seconds 1
    $syncHash.Window.Close()
    Start-Process -WindowStyle Hidden -FilePath "$PSHOME\powershell.exe" -ArgumentList " -ExecutionPolicy Bypass -NonInteractive -File $($PSCommandPath)"
    exit
})

$syncHash.saveCancelClick.add_Click({ $syncHash.SaveDialog.IsOpen = $false })

$syncHash.savePublishPathClick.add_Click({
    $folderPath = New-FolderSelection
    if (![string]::IsNullOrEmpty($folderPath) -and (Test-Path $folderPath)) {
        $syncHash.savePublishPath.Text = $folderPath
    }
})

#endregion

###### MAIN WINDOW
$syncHash.TabMenu.add_SelectionChanged( {
    # need to change this back to XAML - placeholder
    $syncHash.TabMenu.Items | ForEach-Object -Process { $_.Background = '#FF444444' }
    $syncHash.tabMenu.SelectedItem.Background = '#576573'
    $syncHash.TabMenu.Items.Header | ForEach-Object -Process { $_.Foreground = 'Gray' }
    $syncHash.tabMenu.SelectedItem.Header.Foreground = 'AliceBlue'

    $syncHash.historySideDataGrid.ItemsSource = $configHash.actionLog
    $syncHash.historySideDataGrid.Items.Refresh()

    if ($syncHash.tabMenu.SelectedItem.Tag -eq 'Query') { $syncHash.SearchBox.Focus() }

    if ($syncHash.tabMenu.SelectedItem.Tag -eq 'Console') { $syncHash.consoleControl.StartProcess(("$PSHOME\powershell.exe")) }

    elseif ($syncHash.consoleControl.IsProcessRunning) { $syncHash.consoleControl.StopProcess() }

    $syncHash.historySidePane.IsOpen = $false

})

$syncHash.splashLoad.Add_IsVisibleChanged({
    # set ADDS PS default params
    $global:PSDefaultParameterValues = @{"*-AD*:Server"=$configHash.defaultDC;"Choose-ADOrganizationalUnit:Domain"=$configHash.configuredDomain}
 
})

$syncHash.searchBoxHelp.add_Click( {
        $syncHash.childHelp.isOpen = $true  
    })

$syncHash.HistoryToggle.Add_MouseLeftButtonUp( {
        if ($syncHash.historySidePane.IsOpen) { $syncHash.historySidePane.IsOpen = $false }
        else { $syncHash.historySidePane.IsOpen = $true }
        $syncHash.historySideDataGrid.Items.Refresh()
    })

$syncHash.historyButton.Add_Click( {
        if ($syncHash.historySidePane.IsOpen) { $syncHash.historySidePane.IsOpen = $false }
        else { $syncHash.historySidePane.IsOpen = $true }
        $syncHash.historySideDataGrid.Items.Refresh()
    })

$syncHash.settingSelDomainClear.Add_Click({ Start-DomainChangeDialog -SyncHash $syncHash -ConfigHash $configHash -ConfigPath $savedConfig })

$syncHash.tabControl.ItemsSource = New-Object -TypeName System.Collections.ObjectModel.ObservableCollection[Object]

$syncHash.tabMenu.add_Loaded( {

        if ($sysCheckHash.sysChecks.RSModule -eq $true) {
            
            $rsArgs = @{
                Name            = 'menuLoad'
                ArgumentList    = @($syncHash, $savedConfig, $sysCheckHash)
                ModulesToImport = $configHash.modList
            }

            Start-RSJob @rsArgs -ScriptBlock {        
                Param($syncHash, $savedConfig, $sysCheckHash)

                do {} until ($sysCheckHash.checkComplete)
                
                if ($sysCheckHash.sysChecks.ADDS -eq $false -or $sysCheckHash.sysChecks.Modules -eq $false -and $sysCheckHash.sysChecks[0].IsReport -eq $false -and $sysCheckHash.sysChecks.Admin -eq $false) { Suspend-FailedItems -SyncHash $syncHash -CheckedItems SysCheck }             

                elseif (!(Test-Path $savedConfig)) { Suspend-FailedItems -SyncHash $syncHash -CheckedItems Config }

                elseif ($sysCheckHash.sysChecks[0].IsReport -eq $true -and $sysCheckHash.sysChecks.Admin -eq $false) {
                    Set-ReportView -SyncHash $syncHash -ConfigHash $configHash
                }
        
                else { $syncHash.Window.Dispatcher.invoke([action] { $syncHash.tabMenu.SelectedIndex = 0 }) }
            }
       
            New-VarUpdater -ConfigHash $configHash
            Start-VarUpdater -ConfigHash $configHash -VarHash $varHash -QueryHash $queryHash -SyncHash $syncHash
            Set-WPFHeader -ConfigHash $configHash -SyncHash $syncHash
        }

        if (Test-Path $savedConfig) {              
            foreach ($type in @('User', 'Comp')) {             
                if (!($configHash.($type + 'PropListSelection'))) { $configHash.($type + 'PropListSelection') = @() }

                for ($i = 1; $i -le $configHash.boxMax; $i++) {
                    if ($i -le $configHash.($type + 'boxCount')) {
                        Remove-Variable -Name Selected -ErrorAction SilentlyContinue
                        $selected = $configHash.($type + 'PropList') | Where-Object -FilterScript { $_.Field -eq $i }
       
                        if ($null -ne $selected.FieldName) { 
                            $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Header').Content = $selected.FieldName
                            $configHash.($type + 'PropListSelection') += $selected.PropName

                            switch ($selected.ActionName) {

                                { $selected.ActionName -notmatch 'Editable' } {
                                    $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'TextBox').Tag = 'NoEdit'
                                    $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'EditClip').Visibility = 'Collapsed'
                                    $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1').Visibility = 'Collapsed'                                                   
                                }

                                { $selected.ActionName -notmatch 'Actionable' } {
                                    $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1Action1').Visibility = 'Collapsed'  
                                    $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1Action2').Visibility = 'Collapsed'  
                                }

                                { $selected.ActionName -match 'Editable' } {
                                    $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1').Add_Click( {
                                            param([Parameter(Mandatory)][Object]$sender)                       
                                            if ($sender.Name -like '*ubox*') {
                                                $id = $sender.Name -replace 'ubox|box.*'
                                                $type = 'User'
                                            }
                                            else {
                                                $id = $sender.Name -replace 'cbox|box.*'
                                                $type = 'Comp'
                                            }

                                            if ($configHash.adPropertyMap) {
                                                Remove-Variable -Name ChangedValue, propType, ldapValue, user -ErrorAction SilentlyContinue
                                                $changedValue = $syncHash.($type[0] + 'box' + $id + 'resources').($type[0] + 'box' + $id + 'TextBox').Text
                                                $propType = $configHash.($type + 'PropList')[$id - 1].PropType
                            
                                                if ($changedValue -as $propType) {    
                                                    $changedValue = $changedValue -as $propType
                                                    

                                                    try {
                                                        if ($type -eq 'User') {
                                                            $editObject = ($queryHash[$syncHash.tabControl.SelectedItem.Name]).SamAccountName
                                                            $oldValue = $queryHash.($editObject).($configHash.($type + 'PropList')[$id - 1].PropName)

                                                            if ($configHash.adPropertyMap[(($configHash.($type + 'PropList')[$id - 1].PropName))]) { $ldapValue = $configHash.adPropertyMap[(($configHash.($type + 'PropList')[$id - 1].PropName))] }
                                                            else { $ldapValue = $configHash.($type + 'PropList')[$id - 1].PropName }

                                                            Set-ADUser -Identity $editObject -Replace @{
                                                                $ldapValue = $changedValue
                                                            }
                                                        }

                                                        else {
                                                            $editObject = ($queryHash[$syncHash.tabControl.SelectedItem.Name]).Name
                                                            $oldValue = $queryHash.($editObject).($configHash.($type + 'PropList')[$id - 1].PropName)

                                                            if ($configHash.adPropertyMap[(($configHash.($type + 'PropList')[$id - 1].PropName))]) { $ldapValue = $configHash.adPropertyMap[(($configHash.($type + 'PropList')[$id - 1].PropName))] }
                                                            else { $ldapValue = $configHash.($type + 'PropList')[$id - 1].PropName }
                                                            
                                                            Set-ADComputer -Identity $editObject -Replace @{
                                                                $ldapValue = $changedValue
                                                            } 
                                                        }

                                                        $queryHash.($editObject).($configHash.($type + 'PropList')[$id - 1].PropName) = $changedValue
                                                        Write-SnackMsg -Queue $syncHash.SnackMsg.MessageQueue -ToolName 'EDIT' -SubjectName $editObject -Status Success                                                     
                                                        Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName Edit -SubjectName $editObject -SubjectType $type -OldValue $oldValue -NewValue $changedValue -ArrayList $configHash.actionLog 
                                                    }
                                    
                                                    catch { 
                                                        
                                                        if ($null -eq $changedValue) { $changedValue = "[N/A]" }
                                                        Write-SnackMsg -Queue $syncHash.SnackMsg.MessageQueue -ToolName 'EDIT' -SubjectName $editObject -Status Fail  
                                                        Write-LogMessage -Path $configHash.actionlogPath -Message Fail -ActionName Edit -SubjectName $editObject -SubjectType $type -OldValue $oldValue -NewValue $changedValue -ArrayList  $configHash.actionLog -Error $_
                                                    }
                                                }

                                                else {
                                                    if ([string]::IsNullOrEmpty($changedValue)) {

                                                        if ($configHash.adPropertyMap[(($configHash.($type + 'PropList')[$id - 1].PropName))]) { $ldapValue = $configHash.adPropertyMap[(($configHash.($type + 'PropList')[$id - 1].PropName))] }
                                                        else { $ldapValue = $configHash.($type + 'PropList')[$id - 1].PropName }
                                                    
                                                        try {
                                                            if ($type -eq 'User') {
                                                                $editObject = ($queryHash[$syncHash.tabControl.SelectedItem.Name]).SamAccountName
                                                                $oldValue = $queryHash.($editObject).($configHash.($type + 'PropList')[$id - 1].PropName)
                                                                Set-ADUser -Identity $editObject -Clear $ldapValue
                                                            }
                                                            else {
                                                                $editObject = ($queryHash[$syncHash.tabControl.SelectedItem.Name]).Name
                                                                $oldValue = $queryHash.($editObject).($configHash.($type + 'PropList')[$id - 1].PropName)
                                                                Set-ADComputer -Identity $editObject -Clear $ldapValue
                                                            }
                                                            
                                                            Write-SnackMsg -Queue $syncHash.SnackMsg.MessageQueue -ToolName 'CLEAR' -SubjectName $editObject -Status Success
                                                            Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName Clear -SubjectName $editObject -SubjectType $type -OldValue $oldValue -ArrayList $configHash.actionLog 
                                                        }

                                                        catch { 
                                                            Write-LogMessage -Path $configHash.actionlogPath -Message Fail -ActionName Clear -SubjectName $editObject -SubjectType $type -OldValue $oldValue -ArrayList $configHash.actionLog -Error $_
                                                            Write-SnackMsg -Queue $syncHash.SnackMsg.MessageQueue -ToolName 'CLEAR' -SubjectName $editObject -Status Fail
                                                        }
                                                    }

                                                    else {     
                                                        Write-LogMessage -Path $configHash.actionlogPath -Message Fail -ActionName Edit -SubjectName $editObject -SubjectType $type -OldValue $oldValue -ArrayList $configHash.actionLog -Error $_
                                                        $syncHash.SnackMsg.MessageQueue.Enqueue("[EDIT]: FAILED on [$($($editObject).toLower())] - expected type  [$($($propType).toUpper())]")
                                                    }
                                                }
                                            }
                        
                                            else { 
                                                Write-LogMessage -Path $configHash.actionlogPath -Message Fail -ActionName Edit -SubjectName $editObject -SubjectType $type -ArrayList $configHash.actionLog -Error $_
                                                $syncHash.SnackMsg.MessageQueue.Enqueue("[EDIT]: FAILED on [$($($editObject).toLower())] - ADEntity .dll missing; cannot edit")
                                            }
                                        })
                                }

                                { $selected.ActionName -match 'Actionable' } {
                                    for ($b = 1; $b -le 2; $b++) {
                                        # Event assigment for panels in user expander
                                        if ($configHash.($type + 'PropList')[$i - 1].('validAction' + $b) -and !($b -eq 2 -and ($configHash.($type + 'PropList')[$i - 1].actionCmd2Enabled -eq $false))) {

                                  

                                            $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1Action' + $b).Content = $configHash.($type + 'PropList')[$i - 1].('actionCmd' + $b + 'Icon')
                                            $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1Action' + $b).ToolTip = $configHash.($type + 'PropList')[$i - 1].('actionCmd' + $b + 'ToolTip')
                                            
                                            if (($configHash.($type + 'PropList')[$i - 1]).('actionCmd' + $b + 'Multi') -eq $true) {                       
                                                $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + ('Box1Action' + $b)).Add_Click( {
                                                        param([Parameter(Mandatory)][Object]$sender)
                                    
                                                        if ($sender.Name -like '*ubox*') {
                                                            $id = $sender.Name -replace 'ubox|Box.*' 
                                                            $type = 'User'
                                                        }

                                                        else {
                                                            $id = $sender.Name -replace 'cbox|Box.*' 
                                                            $type = 'Comp'
                                                        }

                                                        $b = $sender.Name -replace '.*action'

                                                        $rsCmd = @{
                                                            id           = $id
                                                            type         = $type
                                                            propList     = $configHash.($type + 'PropList')[$id - 1]
                                                            actionObject = if ($type -eq 'User') { ($queryHash[$syncHash.tabControl.SelectedItem.Name]).SamAccountName }
                                                            else { ($queryHash[$syncHash.tabControl.SelectedItem.Name]).Name }
                                                            boxResources = $syncHash.($type[0] + 'box' + $id + 'resources')
                                                            window       = $syncHash.Window
                                                            snackMsg     = $syncHash.SnackMsg.MessageQueue
                                                        }
    
                                                          
                                                        $rsArgs = @{
                                                            Name            = 'threadedAction'
                                                            ArgumentList    = @($rsCmd, $queryHash, $b, $configHash, $syncHash.adHocConfirmWindow, $syncHash.Window, $syncHash.adHocConfirmText, $varHash)
                                                            ModulesToImport = $configHash.modList
                                                        }

                                                        Start-RSJob @rsArgs -ScriptBlock {
                                                            Param($rsCmd, $queryHash, $b, $configHash, $confirmWindow, $window, $textBlock, $varHash)
                                                        
                                                            Start-Sleep -Milliseconds 500
                                                           
                                                            Set-CustomVariables -VarHash $varHash -ConfigHash $configHash

                                                            $actionName = $rsCmd.propList.('actionCmd' + $b + 'ToolTip')
                                                            $actionObject = $rsCmd.actionObject

                                                            $cmd = $rsCmd.propList.('actionCmd' + $b)
                                                            $type = $rsCmd.Type
                                                            $propName = $rsCmd.propList.PropName
                                                            $fieldName = $rsCmd.propList.FieldName
                                            
                                                            New-Variable -Name $type -Value $rsCmd.actionObject
                                                
                                                            if ($propName -ne 'Non-Ad Property') { $prop = $queryHash.$actionObject.$propName }
                                                            else { $prop = $queryHash.$actionObject.$fieldName } 

                                                            try {
                                                                ([scriptblock]::Create($cmd)).Invoke()                                                                
                                                                Write-SnackMsg -Queue $rsCmd.SnackMsg -ToolName $actionName -Status Success -SubjectName $actionObject                           
                                                                Write-LogMessage -syncHashWindow $rsCmd.Window -Path $configHash.actionlogPath -Message Succeed -ActionName $actionName -SubjectName $actionObject -SubjectType $type -ArrayList $configHash.actionLog 

                                                                if ($rsCmd.propList.('actionCmd' + $b + 'Refresh')) {
                                                                    if ($propName -ne 'Non-Ad Property') {
                                                                        if ($type -eq 'User') {                               
                                                                            $result = (Get-ADUser -Identity $user -Properties $propName).$propName
                                                                            $queryHash.($user).($propName) = $result                                                                       
                                                                        }
                                                                        else {
                                                                            $result = (Get-AdComputer -Identity $comp -Properties $propName).$propName
                                                                            $queryHash.($comp).($propName) = $result
                                                                        }
                                                                    }
    


                                                                    if ($queryHash.(Get-Variable -Name $type -ValueOnly).ActiveItem -eq $true) {  
                                                                        if ($rsCmd.propList.ValidCmd) {
                                                                            $tranCmd = $rsCmd.propList.translationCmd
                                                                            $value = Invoke-Expression $tranCmd
                                                                           
                                                                            if ($resultColor) { $rsCmd.Window.Dispatcher.Invoke([Action] { $rsCmd.boxResources.($type[0] + 'box' + $rsCmd.id + 'TextBox').Foreground = $resultColor }) }

                                                                            $updatedValue = $value
                                                                        }

                                                                        else { $updatedValue = $result }

                                                                        $rsCmd.Window.Dispatcher.Invoke([Action] { 
                                                                        
                                                                        if (($rsCmd.propList.actionCmd1CanOff) -and ($updatedValue -like ($rsCmd.propList.actionCmd1offStr))) {
                                                                            $rsCmd.boxResources.($type[0] + 'box' + $rscmd.id + 'Box1Action1').Tag = '$null'
                                                                        }
                                                                        else { $rsCmd.boxResources.($type[0] + 'box' + $rscmd.id + 'Box1Action1').Tag = '' }

                                                                        if (($rsCmd.propList.actionCmd2CanOff) -and ($updatedValue -like ($rsCmd.propList.actionCmd2offStr))) {
                                                                            $rsCmd.boxResources.($type[0] + 'box' + $rscmd.id + 'Box1Action1').Tag = '$null'
                                                                        }
                                                                        else { $rsCmd.boxResources.($type[0] + 'box' + $rscmd.id + 'Box1Action2').Tag = '' }
                                                                            

                                                                        $rsCmd.boxResources.($type[0] + 'box' + $rsCmd.id + 'TextBox').Text = $updatedValue })                                        
                                                                    }
                                                                }
                                                            }
                          
                                                            catch {
                                                                Write-SnackMsg -Queue $rsCmd.SnackMsg -ToolName $actionName -Status Fail -SubjectName $actionObject
                                                                Write-LogMessage -Path $configHash.actionlogPath -Message Fail -ActionName $actionName -SubjectName $actionObject -SubjectType $type -ArrayList $configHash.actionLog -Error $_
                                                            }
                                                        }
                                                    })                         
                                            }
                          
                                            else {
                                                $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + ('Box1Action' + $b)).Add_Click( {
                                                        param([Parameter(Mandatory)][Object]$sender)

                                    
                                                        if ($sender.Name -like '*ubox*') {
                                                            $id = $sender.Name -replace 'ubox|Box.*' 
                                                            $type = 'User'
                                                            $actionObject = ($queryHash[$syncHash.tabControl.SelectedItem.Name]).SamAccountName
                                                        }

                                                        else {
                                                            $id = $sender.Name -replace 'cbox|Box.*' 
                                                            $type = 'Comp'
                                                            $actionObject = ($queryHash[$syncHash.tabControl.SelectedItem.Name]).Name
                                                        }
                                   
                                                        $b = $sender.Name -replace '.*action'

                                                        Set-CustomVariables -VarHash $varHash -ConfigHash $configHash

                                                        Remove-Variable -Name $type -ErrorAction SilentlyContinue
                                                        New-Variable -Name $type -Value $actionObject
                                                        $propName = $configHash.($type + 'PropList')[$id - 1].PropName
                                                        $prop = $queryHash.(Get-Variable -Name $type -ValueOnly).($propName)
                                                        $actionName = $configHash.($type + 'PropList')[$id - 1].('actionCmd' + $b + 'ToolTip')

                                                        try {
                                                            $cmd = $configHash.($type + 'PropList')[$id - 1].('actionCmd' + $b)
                                                            ([scriptblock]::Create($cmd)).Invoke()
                                                            Write-SnackMsg -Queue ($syncHash.SnackMsg.MessageQueue) -ToolName $actionName -Status Success -SubjectName $actionObject
                                                            Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName $actionName -SubjectName $actionObject -SubjectType $type -ArrayList $configHash.actionLog 

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
                                                   
                                                    
                                                                if (($configHash.($type + 'PropList') | Where-Object -FilterScript { $_.Field -eq $id }).ValidCmd -and ($configHash.($type + 'PropList') | Where-Object -FilterScript { $_.Field -eq $id }).transCmdsEnabled) {
                                                                    Remove-Variable -Name resultColor -ErrorAction SilentlyContinue
                                                                    $value = Invoke-Expression ($configHash.($type + 'PropList') | Where-Object { $_.Field -eq $id }).TranslationCmd
                                                              
                                        
                                                                    if ($resultColor) { $syncHash.($type[0] + 'box' + $id + 'resources').($type[0] + 'box' + $id + 'TextBox').Foreground = $resultColor }

                                                                    $updatedValue = $value
                                                                }

                                                                else { $updatedValue = $result }

                                                                $syncHash.($type[0] + 'box' + $id + 'resources').($type[0] + 'box' + $id + 'TextBox').Text = $updatedValue

                                                                (1..2) | ForEach-Object {
                                                                    if (($configHash.($type + 'PropList')[$id - 1].('actionCmd' + $_ + 'CanOff')) -and ($updatedValue -like ($configHash.($type + 'PropList')[$id - 1].('actionCmd' + $_ + 'offStr')))) {
                                                                        ($syncHash.($type[0] + 'box' + $id + 'resources')).($type[0] + 'box' + $id + 'Box1Action' + $_).Tag = '$null'
                                                                    }
                                                                    else { ($syncHash.($type[0] + 'box' + $id + 'resources')).($type[0] + 'box' + $id + 'Box1Action' + $_).Tag = '' }
                                                                }                                                                                
                                                            }                       
                                                        }
                          
                                                        catch {
                                                            Write-SnackMsg -Queue ($syncHash.SnackMsg.MessageQueue) -ToolName $actionName -Status Fail -SubjectName $actionObject
                                                            Write-LogMessage -Path $configHash.actionlogPath -Message Fail -ActionName $actionName -SubjectName $actionObject -SubjectType $type -ArrayList $configHash.actionLog -Error $_
                                                        }                                                               
                                                    })
                                            }
                                        }
                    
                                        else { $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + ('Box1Action' + $b)).Visibility = 'Collapsed' }
                                    }
                                }
                            }
                        }

                        else {
                            $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Header').Visibility = 'Collapsed'
                            $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Border').Visibility = 'Collapsed'
                        }          
                    }
                }
        
                if ($configHash.($type + 'boxCount') -ge 7) { $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.($type + 'detailGrid').Columns = 2 }) }

               
            
                else { $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.($type + 'detailGrid').Columns = 1 }) }
            }
            # populate compgrid connection button icons           
            foreach ($buttonType in @('rbut', 'rcbut')) {
                $syncHash.($buttonType + 1).Source = ([Convert]::FromBase64String($configHash.rtConfig.MSTSC.Icon))
                $syncHash.($buttonType + 1).Parent.Tag = 'MSTSC'
                $syncHash.($buttonType + 1).Parent.ToolTip = $configHash.rtConfig.MSTSC.DisplayName
                $syncHash.($buttonType + 1).Parent.Add_Click( {
                        param([Parameter(Mandatory)][Object]$sender) 
                        if ($sender.Name -match 'rbut') {
                            if ($syncHash.userCompFocusHostToggle.IsChecked) { $comp = $syncHash.UserCompGrid.SelectedItem.HostName }
                            else { $comp = $syncHash.UserCompGrid.SelectedItem.ClientName }
                        }
                        else {
                            if ($syncHash.compUserFocusUserToggle.IsChecked) { $comp = $syncHash.tabControl.SelectedItem.Name }               
                            else { $comp = $syncHash.CompUserGrid.SelectedItem.ClientName }
                        }

                        Set-CustomVariables -VarHash $varHash -ConfigHash $configHash

                        try { 
                            ([scriptblock]::Create($configHash.rtConfig.MSTSC.cmd)).Invoke() 
                            Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName 'Remote Tool (RDP)' -SubjectName $comp -SubjectType 'Computer' -ArrayList $configHash.actionLog 
                        }

                        catch { 
                            Write-SnackMsg -Queue $syncHash.SnackMsg.MessageQueue -ToolName 'Remote Tool (RDP)' -SubjectName $comp -Status Fail                                                     
                            Write-LogMessage -Path $configHash.actionlogPath -Message Fail -ActionName 'Remote Tool (RDP)' -SubjectName $comp -SubjectType 'Computer' -ArrayList $configHash.actionLog -Error $_
                        }
                    })

                $syncHash.($buttonType + 2).Source = ([Convert]::FromBase64String($configHash.rtConfig.MSRA.Icon))
                $syncHash.($buttonType + 2).Parent.Tag = 'MSRA'
                $syncHash.($buttonType + 2).Parent.ToolTip = $configHash.rtConfig.MSRA.DisplayName
                $syncHash.($buttonType + 2).Parent.Add_Click( { 
                        param([Parameter(Mandatory)][Object]$sender)
                        if ($sender.Name -match 'rbut') {
                            if ($syncHash.userCompFocusHostToggle.IsChecked) { $comp = $syncHash.UserCompGrid.SelectedItem.HostName }
                            else { $comp = $syncHash.UserCompGrid.SelectedItem.ClientName } 
                        }

                        else {
                            if ($syncHash.compUserFocusUserToggle.IsChecked) { $comp = $syncHash.tabControl.SelectedItem.Name }             
                            else { $comp = $syncHash.CompUserGrid.SelectedItem.ClientName } 
                        }
                        
                        Set-CustomVariables -VarHash $varHash -ConfigHash $configHash
                        
                        try { 
                            ([scriptblock]::Create($configHash.rtConfig.MSRA.cmd)).Invoke() 
                            Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName 'Remote Tool (MSRA)' -SubjectName $comp -SubjectType 'Computer' -ArrayList $configHash.actionLog 
                        }

                        catch { 
                            Write-SnackMsg -Queue $syncHash.SnackMsg.MessageQueue -ToolName 'Remote Tool (MSRA)' -SubjectName $comp -Status Fail                                                     
                            Write-LogMessage -Path $configHash.actionlogPath -Message Fail -ActionName 'Remote Tool (MSRA)' -SubjectName $comp -SubjectType 'Computer' -ArrayList $configHash.actionLog -Error $_
                        }
                        
                    })    


                foreach ($rtID in $configHash.rtConfig.Keys.Where{ $_ -like 'RT*' }) {
                    $syncHash.customRT.$rtID.$buttonType = New-Object -TypeName System.Windows.Controls.Button -Property @{
                        Padding = '0'
                        Tag     = $rtID
                        ToolTip = $configHash.rtConfig.$rtID.DisplayName
                        Style   = $syncHash.Window.FindResource('itemButton')
                        Name    = $buttonType + ([int]($rtID -replace 'rt') + 2)
                    }
            
                    $syncHash.customRT.$rtID.($buttonType + 'img') = New-Object -TypeName System.Windows.Controls.Image -Property @{
                        Width  = 15
                        Height = 15
                        Source = ([Convert]::FromBase64String($configHash.rtConfig.$rtID.Icon))
                    }

                    $syncHash.customRT.$rtID.$buttonType.AddChild($syncHash.customRT.$rtID.($buttonType + 'img'))

                    $syncHash.customRT.$rtID.($buttonType + 'img').Parent.Add_Click( { 
                            param([Parameter(Mandatory)][Object]$sender) 
                            $id = 'rt' + ([int]($sender.Name -replace '.*but') - 2)  
                            $toolName = $configHash.rtConfig.$id.DisplayName            

                            if ($sender.Name -match 'rbut') {
                                if ($syncHash.userCompFocusHostToggle.IsChecked) { $comp = $syncHash.UserCompGrid.SelectedItem.HostName }
                                else { $comp = $syncHash.UserCompGrid.SelectedItem.ClientName }

                                $exe = $configHash.rtConfig.$id.Path
                                $user = $syncHash.UserCompGrid.SelectedItem.UserName
                                $sessionID = $syncHash.UserCompGrid.SelectedItem.SessionID
                            }

                            else {
                                if ($syncHash.compUserFocusUserToggle.IsChecked) { $comp = $syncHash.tabControl.SelectedItem.Name }
                
                                else { $comp = $syncHash.CompUserGrid.SelectedItem.ClientName }

                                $exe = $configHash.rtConfig.$id.Path
                                $user = $syncHash.CompUserGrid.SelectedItem.UserName
                                $sessionID = $syncHash.CompUserGrid.SelectedItem.SessionID
                            }

                            Set-CustomVariables -VarHash $varHash -ConfigHash $configHash
                            
                              
                            try { 
                                ([scriptblock]::Create($configHash.rtConfig.$id.cmd)).Invoke()
                                Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName "Remote Tool ($($toolName))" -SubjectName $comp -ContextSubjectName $user -SubjectType 'Computer' -ArrayList $configHash.actionLog 
                            }

                            catch { 
                                Write-SnackMsg -Queue $syncHash.SnackMsg.MessageQueue -ToolName "Remote Tool ($($toolName))" -SubjectName $comp -Status Fail                                                     
                                Write-LogMessage -Path $configHash.actionlogPath -Message Fail -ActionName "Remote Tool ($($toolName))" -SubjectName $comp -ContextSubjectName $user -SubjectType 'Computer' -ArrayList $configHash.actionLog  -Error $_
                            }
                   
                        })

                    $syncHash.($buttonType + 'Grid').AddChild($syncHash.customRT.$rtID.$buttonType)
                }

                #here#
                foreach ($contextBut in $configHash.contextConfig.Where{ $_.ValidAction -eq $true }) {
                    if ($null -eq $syncHash.customContext.('cxt' + $contextBut.IDNum)) { $syncHash.customContext.('cxt' + $contextBut.IDNum) = @{} }
                               
                    $syncHash.customContext.('cxt' + $contextBut.IDNum).($buttonType + 'context' + $contextBut.IDNum) = New-Object -TypeName System.Windows.Controls.Button -Property @{
                        Padding    = '0'
                        Tag        = $buttonType + 'context' + $contextBut.IDNum
                        ToolTip    = $configHash.contextConfig[$contextBut.IDnum - 1].ActionName
                        Style      = $syncHash.Window.FindResource('itemButton')
                        Name       = $buttonType + 'context' + $contextBut.IDNum
                        FontFamily = 'Segoe MDL2 Assets'
                        Content    = $configHash.contextConfig[$contextBut.IDnum - 1].actionCmdIcon
                    }
            

                    $syncHash.customContext.('cxt' + $contextBut.IDNum).($buttonType + 'context' + $contextBut.IDNum).Add_Click( { 
                            param([Parameter(Mandatory)][Object]$sender) 
                            $id = [int]($sender.Name -replace '.*context')
                            if ($sender.Parent.Name -eq 'rbutContextGrid') {
                                $user = $configHash.currentTabItem
                                $sessionID = $syncHash.UserCompGrid.SelectedItem.SessionID

                                if ($syncHash.userCompFocusHostToggle.IsChecked) { $comp = $syncHash.UserCompGrid.SelectedItem.HostName }
                                else { $comp = $syncHash.UserCompGrid.SelectedItem.ClientName }
                            }
                        
                            else {
                                $user = $syncHash.CompUserGrid.SelectedItem.UserName
                                $sessionID = $syncHash.CompUserGrid.SelectedItem.SessionID

                                if ($syncHash.compUserFocusUserToggle.IsChecked) { $comp = $configHash.currentTabItem }
                
                                else { $comp = $syncHash.CompUserGrid.SelectedItem.ClientName }     
                            }
                        
                            if ($configHash.contextConfig[$id - 1].actionCmdMulti) {
                                $rsCmd = @{
                                    comp           = $comp
                                    user           = $user
                                    buttonSettings = $configHash.contextConfig[$id - 1]
                                    sessionID      = $sessionID
                                    snackBar       = $syncHash.SnackMsg.MessageQueue
                                }
   
                                $rsArgs = @{
                                    Name            = 'buttonThreadedAction'
                                    ArgumentList    = @($rsCmd, $syncHash, $configHash, $syncHash.adHocConfirmWindow, $syncHash.Window, $syncHash.adHocConfirmText, $varHash)
                                    ModulesToImport = $configHash.modList
                                }

                                Start-RSJob @rsArgs -ScriptBlock {
                                    Param($rsCmd, $syncHash, $configHash, $confirmWindow, $window, $textBlock, $varHash)
                            
                                    $user = $rsCmd.user
                                    $comp = $rsCmd.comp
                                    $sessionID = $rsCmd.SessionID

                                    $combinedString = "$($user.ToLower()) on $($comp.toUpper())"
                                    $toolName = $rsCmd.buttonSettings.ActionName
                                    Set-CustomVariables -VarHash $varHash -ConfigHash $configHash

                                    try {
                                        ([scriptblock]::Create($rsCmd.buttonSettings.actionCmd)).Invoke()
                                        Write-SnackMsg -Queue $rsCmd.Snackbar -ToolName $toolName -Status Success -SubjectName $combinedString                            
                                        Write-LogMessage -syncHashWindow $syncHash.Window -Path $configHash.actionlogPath -Message Succeed -ActionName $toolname -SubjectName $user -ContextSubjectName $comp -SubjectType 'Context' -ArrayList $configHash.actionLog
                                    }
                            
                                    catch {
                                        Write-SnackMsg -Queue $rsCmd.Snackbar -ToolName $toolName -Status Fail -SubjectName $combinedString
                                        Write-LogMessage -syncHashWindow $syncHash.Window -Path $configHash.actionlogPath -Message Fail -ActionName $toolname -SubjectName $user -ContextSubjectName $comp -SubjectType 'Context' -ArrayList $configHash.actionLog -Error $_ 
                                    }
                                }
                            }

                            else {
                                $combinedString = "$($user.ToLower()) on $($comp.ToUpper())"
                                $toolName = $configHash.contextConfig[$id - 1].ActionName
                                Set-CustomVariables -VarHash $varHash -ConfigHash $configHash

                                try {
                                    ([scriptblock]::Create($configHash.contextConfig[$id - 1].actionCmd)).Invoke()
                                    Write-SnackMsg -Queue $syncHash.SnackMsg.MessageQueue -ToolName $toolName -Status Success -SubjectName $combinedString
                                    Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName $toolName -SubjectName $user -ContextSubjectName $comp -SubjectType 'Context' -ArrayList $configHash.actionLog 
                                }
                            
                                catch {
                                    Write-SnackMsg -Queue $syncHash.SnackMsg.MessageQueue -ToolName $toolName -Status Fail -SubjectName $combinedString
                                    Write-LogMessage -Path $configHash.actionlogPath -Message Fail -ActionName $toolName -SubjectName $user -ContextSubjectName $comp -SubjectType 'Context' -ArrayList $configHash.actionLog -Error $_
                                }
                            }
                        })
                
                    $syncHash.($buttonType + 'ContextGrid').AddChild($syncHash.customContext.('cxt' + $contextBut.IDNum).($buttonType + 'context' + $contextBut.IDNum))
                }
            }
      
            # add fields in panel under historical view for custom fields (for both user/comp logs)
            $configHash.userLogMapping |
                Where-Object -FilterScript { $_.FieldSel -eq 'Custom' -and $_.Ignore -eq $false } |
                   New-CustomLogHeader -SyncHash $syncHash -ConfigHash $configHash -Type User   
          
            $configHash.compLogMapping |
                Where-Object -FilterScript { $_.FieldSel -eq 'Custom' -and $_.Ignore -eq $false } |
                      New-CustomLogHeader -SyncHash $syncHash -ConfigHash $configHash -Type Comp

            # resize the panel under historical view to accomodate custom fields and extra context buttons
            Set-ItemControlPanelSize -SyncHash $syncHash -ConfigHash $configHash
        }

        else { 
            $syncHash.queryTab.IsEnabled = $false
            $syncHash.toolTab.IsEnabled = $false
        }
    })
                
$syncHash.tabControl.add_SelectionChanged( {

        if ($configHash.itemRefreshing -eq $true) {
            $configHash.itemRefreshing = $false 
          
            Set-ItemExpanders -SyncHash $syncHash -ConfigHash $configHash -IsActive Disable -ClearContent

            $syncHash.settingToolParent.Visibility = 'Collapsed'
            $syncHash.userToolControlPanel.Visibility = 'Collapsed'
            $syncHash.compToolControlPanel.Visibility = 'Collapsed'

            $syncHash.tabControl.IsEnabled = $false
        }

        else {
            Get-RSJob -State Completed | Remove-RSJob
            $syncHash.tabControl.IsEnabled = $true
            $syncHash.userCompOutdated.Tag = $null
              
            $SyncHash.userCompGrid.ItemsSource = $null           
            $SyncHash.compUserGrid.ItemsSource = $null 
   
            $configHash.currentTabItem = $syncHash.tabControl.SelectedItem.Name
            $queryHash.Keys | ForEach-Object -Process { $queryHash.$_.ActiveItem = $false }
    
            $rsArgs = @{
                Name            = 'VisualChange'
                ArgumentList    = @($syncHash, $queryHash, $configHash, $varHash)
                ModulesToImport = $configHash.modList
            }

            Start-RSJob @rsArgs -ScriptBlock {
                Param($syncHash, $queryHash, $configHash, $varHash)
            
                Set-ItemExpanders -SyncHash $syncHash -ConfigHash $configHash -IsActive Disable -ClearContent
         
                if ($null -ne $configHash.currentTabItem) {
                
                    $syncHash.Window.Dispatcher.Invoke([Action] {
                        if ($queryHash.($configHash.currentTabItem).ObjectClass -eq 'Computer') {
                            Set-QueryVarsToUpdate -ConfigHash $ConfigHash -Type Comp                          
                            if (($syncHash.compToolControlPanel.Children | Measure-Object).Count -eq 0) { $syncHash.settingTools.Visibility = 'Collapsed' }
                        }
                        else {
                            Set-QueryVarsToUpdate -ConfigHash $ConfigHash -Type User
                            if (($syncHash.userToolControlPanel.Children | Measure-Object).Count -eq 0) { $syncHash.settingTools.Visibility = 'Collapsed' }
                        }
                    })

                    $queryHash.($configHash.currentTabItem).ActiveItem = $true

                    $rsArgs = @{
                        Name            = 'displayUpdate'
                        ArgumentList    = @($syncHash, $queryHash, $configHash, $varHash)
                        ModulesToImport = $configHash.modList
                    }

                    Start-RSJob @rsArgs -ScriptBlock {        
                        Param($syncHash, $queryHash, $configHash, $varHash)                     
                        #Start-Sleep -Milliseconds 500
               
                        $currentTabItem = $configHash.currentTabItem

                        $type = $queryHash[$currentTabItem].ObjectClass -replace 'Computer', 'Comp'
                        $statusTable = [hashtable]::Synchronized(@{ })
                        for ($i = 1; $i -le $configHash.boxMax; $i++) {
                          
                            if ($i -le $configHash.($type + 'boxCount')) {                                                        
                                $rsArgs = @{
                                    Name            = 'displayUpdateSub'
                                    Batch           = 'displayUpdateSub'
                                    Throttle        = ($configHash.($type + 'boxCount'))
                                    ArgumentList    = @($syncHash, $queryHash, $configHash, $currentTabItem, $i, $varHash, ($configHash.($type + 'boxCount')), $statusTable)
                                    ModulesToImport = $configHash.modList
                                }

                                Start-RSJob @rsArgs -ScriptBlock {        
                                    Param($syncHash, $queryHash, $configHash, $currentTabItem, $i, $varHash, $boxCount, $statusTable)   
                                    $type = $queryHash[$currentTabItem].ObjectClass -replace 'Computer', 'Comp'

                                    Set-CustomVariables -VarHash $varHash -ConfigHash $configHash
                                    New-Variable -Name $type -Value $configHash.currentTabItem
                                    # Remove-Variable -Name resultColor -ErrorAction SilentlyContinue                 

                                    $updateHash = @{'Text' = $null }
                                    $selectedBox = ($configHash.($type + 'propList') | Where-Object { $_.Field -eq $i })        

                                    if ($selectedBox.ValidCmd -eq $true -and $selectedBox.transCmdsEnabled) {
                                        $result = ($queryHash[$currentTabItem]).($selectedBox.PropName)
                                        $updateHash.Text = Invoke-Expression $selectedBox.TranslationCmd
                                        
                                        if ($resultColor) { $updateHash.Foreground = $resultColor }             
                                    }

                                  
   
                                    elseif ($selectedBox.PropName) { $updateHash.Text = ($queryHash[$currentTabItem]).($selectedBox.PropName) }

                                    if ($selectedBox.PropName -eq 'Non-Ad Property') { ($queryHash[$currentTabItem]) | Add-Member -Force -MemberType NoteProperty -Name ($selectedBox.FieldName) -Value $value }
                                      
                                    if ($selectedBox.actionCmd1CanOff -and $updateHash.Text -like $selectedBox.actionCmd1OffStr) {
                                        $updateHash.Disable1 = $true
                                    }

                                    if ($selectedBox.actionCmd2CanOff -and $updateHash.Text -like $selectedBox.actionCmd2OffStr) {
                                        $updateHash.Disable2 = $true
                                    }

                                     $statusTable.([string]$i) = $updateHash

                                    
                              
                                }
                            }
                        }

                        do { Start-Sleep -Seconds 1 } while ($statusTable.Keys.Count -lt ($configHash.($type + 'boxCount')))
                   
                        $syncHash.Window.Dispatcher.Invoke([action] {
                         foreach ($key in $statusTable.Keys) {
                            $syncHash.($type[0] + 'box' + $key + 'resources').($type[0] + 'box' + $key + 'Textbox').Text = $statusTable.$key.Text
                            
                          
                         
                            if ($statusTable.$key.Disable1) { ($syncHash.($type[0] + 'box' + $key + 'resources')).($type[0] + 'box' + $key + 'Box1Action1').Tag  = '$null' }
                            else { ($syncHash.($type[0] + 'box' + $key + 'resources')).($type[0] + 'box' + $key + 'Box1Action1').Tag  = '' }

                            if ($statusTable.$key.Disable2) { ($syncHash.($type[0] + 'box' + $key + 'resources')).($type[0] + 'box' + $key + 'Box1Action2').Tag  = '$null' }
                            else { ($syncHash.($type[0] + 'box' + $key + 'resources')).($type[0] + 'box' + $key + 'Box1Action2').Tag  = '' }

                            if ($statusTable.$key.Foreground) { $syncHash.($type[0] + 'box' + $key + 'resources').($type[0] + 'box' +$key + 'Textbox').Foreground = $statusTable.$key.Foreground }        
                         }
                        })
                          

                        if ($queryHash[$currentTabItem].ObjectClass -eq 'User') {  
                            $syncHash.Window.Dispatcher.invoke([action] {       
                                    $syncHash.expanderTypeDisplay.Content = 'USER DETAILS  '
                                    $syncHash.compExpanderTypeDisplay.Content = 'COMPUTER HISTORY  '
                                    $syncHash.compDetailMainPanel.Visibility = 'Collapsed'
                                    $syncHash.userDetailMainPanel.Visibility = 'Visible'
                                    $syncHash.expanderDisplay.Content = "$($queryHash[$currentTabItem].Name)"
                                    $syncHash.userExpander.IsExpanded = $true
                                    
                                    $syncHash.expanderProgressBar.Visibility = 'Hidden'

                       
                                    if (($syncHash.userToolControlPanel.Children | Measure-Object).Count -gt 0) {
                                        $syncHash.settingTools.Visibility = 'Visible'
                                        $syncHash.settingToolParent.Visibility = 'Visible'
                                        $syncHash.userToolControlPanel.Visibility = 'Visible'
                                    }
                                })
                        }  

                        elseif ($queryHash[$currentTabItem].ObjectClass -eq 'Computer') {  
                            $syncHash.Window.Dispatcher.invoke([action] {               
                                    $syncHash.expanderTypeDisplay.Content = 'COMPUTER DETAILS  '
                                    $syncHash.userDetailMainPanel.Visibility = 'Collapsed'
                                    $syncHash.compDetailMainPanel.Visibility = 'Visible'
                                    $syncHash.compExpanderTypeDisplay.Content = 'USER HISTORY  '
                                    $syncHash.expanderDisplay.Content = "$($queryHash[$currentTabItem].Name)"
                                    $syncHash.userExpander.IsExpanded = $true
                                    $syncHash.expanderProgressBar.Visibility = 'Hidden'

                                    if (($syncHash.compToolControlPanel.Children | Measure-Object).Count -gt 0) {
                                        $syncHash.settingTools.Visibility = 'Visible'
                                        $syncHash.settingToolParent.Visibility = 'Visible'
                                        $syncHash.compToolControlPanel.Visibility = 'Visible'
                                    }
                                })    
                        }      
                    }   

                    $rsArgs = @{
                        Name            = 'displayLogUpdate'
                        ArgumentList    = @($syncHash, $queryHash, $configHash)
                        ModulesToImport = $configHash.modList
                    }

                    Start-RSJob @rsArgs  -ScriptBlock {        
                        Param($syncHash, $queryHash, $configHash)        

                        $currentTabItem = $configHash.currentTabItem
                        $type = $queryHash[$currentTabItem].ObjectClass -replace 'Computer', 'Comp'
           
                        do { Start-Sleep -Seconds 1 } until ($queryHash[$currentTabItem].logsSearched -eq $true)

                        if ($null -ne ($queryHash[$currentTabItem]).LoginLogListView) {
                            if ($type -eq 'User') {
                                $syncHash.userCompGrid.Dispatcher.Invoke([Action] { $syncHash.userCompGrid.ItemsSource = ($queryHash[$currentTabItem]).LoginLogListView })
                                $syncHash.Window.Dispatcher.invoke([action] {
                                        $syncHash.compExpander.IsExpanded = $true
                                        $syncHash.compExpanderProgressBar.Visibility = 'Hidden'
                                    })
                            }
                            else {
                                $syncHash.userCompGrid.Dispatcher.Invoke([Action] { $syncHash.compUserGrid.ItemsSource = ($queryHash[$currentTabItem]).LoginLogListView })
                                $syncHash.Window.Dispatcher.invoke([action] {
                                        $syncHash.compExpander.IsExpanded = $true
                                        $syncHash.compExpanderProgressBar.Visibility = 'Hidden'
                                    })
                            }            
                        }

                        else {
                            $syncHash.Window.Dispatcher.invoke([action] {
                                    $syncHash.compExpander.IsExpanded = $true
                                    $syncHash.compExpanderProgressBar.Visibility = 'Hidden'

                                    if ($queryHash[$currentTabItem].HasLogs) { $syncHash.userCompOutdated.Tag = 'HasLogs' }
                                })
                        }
                    }
                    #update valuebox
                }
            }
        }
    })
        
$syncHash.settingObjectToolDefFlyout.Add_ClosingFinished({
    if ($null -ne $syncHash.settingObjectStandaloneCat.Text -and $syncHash.settingObjectStandaloneCat.Text -notin $configHash.SACats) {
        $configHash.SACats.Add($syncHash.settingObjectStandaloneCat.Text)
    }
})

$syncHash.itemToolGridExport.Add_Click({

    
    $configHash.gridExportList = [System.Collections.ArrayList]@()
    $syncHash.itemToolGridItemsGrid.Items | ForEach-Object { $configHash.gridExportList.Add($_) | Out-Null}


    $rsArgs = @{
            Name            = 'ExportGrid'           
            ModulesToImport = $configHash.modList 
            ArgumentList    =  $configHash, $syncHash.itemToolDialog.Title
        }

    $syncHash.snackMsg.MessageQueue.Enqueue("Grid contents exporting...")

    Start-RSJob @rsArgs -ScriptBlock {
        Param ($configHash, $tool) 

        New-HTML -TitleText $title -ShowHTML {
            New-HTMLText -Text $title -FontSize 16 -Fontweight Bolder -Alignment center
            New-HTMLTable -DataTable $configHash.gridExportList -DateTimeSortingFormat 'M/D/YYYY HH:mm' -DefaultSortColumn Date -DefaultSortOrder Descending -SearchBuilder {
                New-HTMLTableStyle -FontFamily 'Segoe UI' -FontWeight 500
            }
       }
    }
})

$syncHash.externalToolList.Add_SelectionChanged({

    if ($syncHash.externalToolList.SelectedItem) {
        & $syncHash.externalToolList.SelectedItem.Exe
        $syncHash.externalToolList.SelectedItem = $null
    }
})

$syncHash.externalToolListPopup.Add_PreviewMouseLeftButtonUp({
$syncHash.externalToolListPopup.IsPopupOpen = $false
})

$syncHash.tabControl.add_tabItemClosingEvent( {

      
        $queryHash.Remove($configHash.currentTabItem)

        if ($syncHash.tabControl.Items.Count -le 1) { Set-ItemExpanders -SyncHash $syncHash -ConfigHash $configHash -IsActive Disable -ClearContent }

        
    })

$syncHash.userCompGrid.Add_SelectionChanged( { Set-GridButtons -SyncHash $syncHash -ConfigHash $configHash -Type User })

$syncHash.compUserGrid.Add_SelectionChanged( { Set-GridButtons -SyncHash $syncHash -ConfigHash $configHash -Type Comp })

$syncHash.userCompFocusHostToggle.Add_Checked( {
        $syncHash.userCompFocusClientToggle.IsChecked = $false    
        
        foreach ($button in $syncHash.Keys.Where( { $_ -like '*rbutbut*' })) {
            if (($syncHash.userCompGrid.SelectedItem.Type -in $configHash.rtConfig.($syncHash[$button].Tag).Types) -and
                (!($configHash.rtConfig.($syncHash[$button].Tag).RequireOnline -and $syncHash.userCompGrid.SelectedItem.Connectivity -eq $false)) -and 
                (!($configHash.rtConfig.($syncHash[$button].Tag).RequireUser -and $syncHash.userCompGrid.SelectedItem.userOnline -eq $false))) { $syncHash[$button].IsEnabled = $true }
            else { $syncHash[$button].IsEnabled = $false }            
        }

        foreach ($button in $syncHash.customRT.Keys) {
            if (($syncHash.userCompGrid.SelectedItem.Type -in $configHash.rtConfig.$button.Types) -and
                (!($configHash.rtConfig.$button.RequireOnline -and $syncHash.userCompGrid.SelectedItem.Connectivity -eq $false)) -and 
                (!($configHash.rtConfig.$button.RequireUser -and $syncHash.userCompGrid.SelectedItem.userOnline -eq $false))) { $syncHash.customRT.$button.rbut.IsEnabled = $true }
            else { $syncHash.customRT.$button.rbut.IsEnabled = $false }            
        }

        foreach ($button in $syncHash.customContext.Keys) {
            if (($syncHash.userCompGrid.SelectedItem.Type -in $configHash.contextConfig[($button -replace 'cxt') - 1].Types) -and
                (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireOnline -and $syncHash.userCompGrid.SelectedItem.Connectivity -eq $false)) -and 
                (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireUser -and $syncHash.userCompGrid.SelectedItem.userOnline -eq $false))) { $syncHash.customContext.$button.('rbutcontext' + ($button -replace 'cxt')).IsEnabled = $true }
            else { $syncHash.customContext.$button.('rbutcontext' + ($button -replace 'cxt')).IsEnabled = $false }            
        }
    })

$syncHash.userCompFocusHostToggle.Add_Unchecked( {
        if ([string]::IsNullOrEmpty($syncHash.userCompGrid.SelectedItem.ClientName)) { $syncHash.userCompFocusHostToggle.IsChecked = $true }
        else { $syncHash.userCompFocusClientToggle.IsChecked = $true } 
    })

$syncHash.userCompFocusClientToggle.Add_Checked( {
        $syncHash.userCompFocusHostToggle.IsChecked = $false   
        Set-ClientGridButtons -SyncHash $syncHash -ConfigHash $configHash -Type User
    })

$syncHash.searchSettingsPopUp.Add_Closed( {
        $configHash.queryProps = [System.Collections.ArrayList]@()
        Get-LDAPSearchNames -ConfigHash $configHash -SyncHash $syncHash | ForEach-Object -Process { $null = $configHash.queryProps.Add($_) }
    })

$syncHash.userCompFocusClientToggle.Add_Unchecked( { $syncHash.userCompFocusHostToggle.IsChecked = $true })

$syncHash.compUserFocusUserToggle.Add_Checked( {
        $syncHash.compUserFocusClientToggle.IsChecked = $false 
        Set-GridButtons -SyncHash $syncHash -ConfigHash $configHash -Type Comp -SkipSelectionChange
    })

$syncHash.compUserFocusUserToggle.Add_Unchecked( {
        if ([string]::IsNullOrEmpty($syncHash.compUserGrid.SelectedItem.ClientName)) { $syncHash.compUserFocusUserToggle.IsChecked = $true }
        else { $syncHash.compUserFocusClientToggle.IsChecked = $true } 
    })

$syncHash.compUserFocusClientToggle.Add_Checked( {
        $syncHash.compUserFocusUserToggle.IsChecked = $false  
        Set-ClientGridButtons -ConfigHash $configHash -SyncHash $syncHash -Type Comp
    })

$syncHash.compUserFocusClientToggle.Add_Unchecked( { $syncHash.compUserFocusUserToggle.IsChecked = $true })



$syncHash.SearchBox.add_KeyDown( {
        if ($_.Key -eq 'Enter' -or $_.Key -eq 'Escape') {
            if ($null -like $syncHash.SearchBox.Text -and $_.Key -ne 'Escape') { $syncHash.SnackMsg.MessageQueue.Enqueue('Searchbox is empty - cannot query') }
            elseif ($configHash.IsSearching) { $syncHash.SnackMsg.MessageQueue.Enqueue('Currently querying... Please wait') }
            elseif ($syncHash.SearchBox.Text.Length -ge 3 -or $_.Key -eq 'Escape') { 
                if (!($configHash.queryProps)) { 
                    $configHash.queryProps = [System.Collections.ArrayList]@()
                    Get-LDAPSearchNames -ConfigHash $configHash -SyncHash $syncHash | ForEach-Object -Process { $null = $configHash.queryProps.Add($_) } 
                }
                Start-ObjectSearch -SyncHash $syncHash -ConfigHash $configHash -QueryHash $queryHash -Key $_.Key
            }

            else { $syncHash.SnackMsg.MessageQueue.Enqueue('Query must be at least 3 characters long!') }
        } 
    })

$syncHash.itemToolDialogConfirmButton.Add_Click( {

        $rsArgs = @{
            Name            = 'ItemTool'
            ArgumentList    = @($syncHash.snackMsg.MessageQueue, $syncHash.itemToolDialogConfirmButton.Tag, $configHash, $queryHash, $syncHash.Window, $syncHash.adHocConfirmWindow, $syncHash.adHocConfirmText, $varHash)
            ModulesToImport = $configHash.modList
        }

        Start-RSJob @rsArgs -ScriptBlock {
            Param($queue, $toolID, $configHash, $queryHash, $window, $confirmWindow, $textBlock, $varHash)

            $toolName = ($configHash.objectToolConfig[$toolID - 1].toolName).ToUpper()
            Set-CustomVariables -VarHash $varHash -ConfigHash $configHash
            if ($configHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {Remove-Variable ActiveObject, ActiveObjectType, ActiveObjectData -ErrorAction SilentlyContinue}

            try {              
                 ([scriptblock]::Create($configHash.objectToolConfig[$toolID - 1].toolAction)).Invoke()     
                          
                if ($configHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -Status Success
                    Write-LogMessage -syncHashWindow $window -Path $configHash.actionlogPath -Message Succeed -ActionName $toolName -SubjectType 'Standalone' -ArrayList $configHash.actionLog
                }

                else {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -Status Success -SubjectName $activeObject
                    Write-LogMessage -syncHashWindow $window -Path $configHash.actionlogPath -Message Succeed -SubjectName $activeObject -SubjectType $activeObjectType -ActionName $toolName  -ArrayList $configHash.actionLog 
                }
            }
            catch {
                if ($configHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -Status Fail
                    Write-LogMessage -syncHashWindow $window -Path $configHash.actionlogPath -Message Fail -ActionName $toolName -ArrayList $configHash.actionLog -Error $_
                }
                else {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -Status Fail -SubjectName $activeObject
                    Write-LogMessage -syncHashWindow $window -Path $configHash.actionlogPath -Message Fail -SubjectName $activeObject -SubjectType $activeObjectType -ActionName $toolName -ArrayList $configHash.actionLog -Error $_
                }
            }
        }

        $syncHash.itemToolDialog.IsOpen = $false
    })

$syncHash.itemToolDialogConfirmCancel.Add_Click( { $syncHash.itemToolDialog.IsOpen = $false })

$syncHash.itemToolDialog.Add_ClosingFinished( {
        $syncHash.itemToolDialogConfirm.Visibility = 'Collapsed'
        $syncHash.itemToolListSelect.Visibility = 'Collapsed'
        $syncHash.itemToolListSelectListBox.ItemsSource = $null
        $syncHash.itemToolADSelectedItem.Content = $null
        $syncHash.itemToolListSearchBox.Text = $null
        $syncHash.itemToolGridSearchBox.Text = $null
        $syncHash.itemToolGridADSelectedItem.Content = $null
        $syncHash.itemToolGridExport.Visibility = 'Collapsed'
        $syncHash.itemToolImageBorder.Visibility = 'Collapsed'
        $syncHash.itemToolGridSelect.Visibility = 'Collapsed'
        $syncHash.itemToolCommandGridPanel.Visibility = 'Collapsed'
        $syncHash.itemToolCustomDialog.Visibility = 'Collapsed'
        $syncHash.itemToolGridSelectAllButton.Visibility = 'Collapsed'
        $syncHash.itemToolListSelectAllButton.Visibility = 'Collapsed'
        $syncHash.itemToolGridItemsEmptyText.Visibility = 'Hidden'
        $syncHash.itemToolListItemsEmptyText.Visibility = 'Hidden'
    })

$syncHash.compExpanderOpenLog.Add_Click({  Invoke-Item $queryHash[$configHash.currentTabItem].LoginLogPath })

$syncHash.userCompOutdated.Add_Click({  Invoke-Item $queryHash[$configHash.currentTabItem].LoginLogPath })

$syncHAsh.settingResetADPropMap.Add_Click({ 
    Remove-SavedPropertyLists -SavedConfig $savedConfig 
    $syncHAsh.settingResetADPropMap.IsEnabled = $false
})
#region ActionLogReview

$syncHash.toolsSingleLogExport.Add_Click( { 
        $syncHash.toolsExportDialog.isOpen = $true
        $syncHash.toolsExportDialog.Tag = 'Single'
    })

$syncHash.toolsAllLogExport.Add_Click( { 
        $syncHash.toolsExportDialog.isOpen = $true
        $syncHash.toolsExportDialog.Tag = 'All'
    })

$syncHash.toolsExportConfirmButton.Add_Click( {

        if ($ConfigHash.actionlogPath -and (Test-Path $ConfigHash.actionlogPath)) {

            if ($syncHash.toolsExportDialog.Tag -eq 'Single') { New-LogHTMLExport -Scope User -ConfigHash $configHash -TimeFrame $syncHash.toolsExportDate.SelectedValue.Content }
            else { New-LogHTMLExport -Scope All -ConfigHash $configHash -TimeFrame $syncHash.toolsExportDate.SelectedValue.Content }

            $syncHash.SnackMsg.MessageQueue.Enqueue('Log export starting...') 
            
        }

        else { $syncHash.SnackMsg.MessageQueue.Enqueue('Log path is invalid or unset; cannot export') }

        $syncHash.toolsExportDialog.isOpen = $false

    })

$syncHash.toolsExportConfirmCancel.Add_Click( { $syncHash.toolsExportDialog.isOpen = $false })

$syncHash.toolsAllLogView.Add_Click( { 

    if ($ConfigHash.actionlogPath -and (Test-Path $ConfigHash.actionlogPath)) {
        $syncHash.toolsLogProgress.Tag = 'Unloaded'
        $syncHash.toolsLogEmpty.Tag = 'Unloaded'
        Initialize-LogGrid -Scope All -ConfigHash $configHash -SyncHash $syncHash
    }

    else { $syncHash.SnackMsg.MessageQueue.Enqueue('Log path is invalid or unset; cannot view logs') }

    })

$syncHash.toolsSingleLogView.Add_Click( { 
        
        if ($ConfigHash.actionlogPath -and (Test-Path $ConfigHash.actionlogPath)) {
            $syncHash.toolsLogProgress.Tag = 'Unloaded'
            $syncHash.toolsLogEmpty.Tag = 'Unloaded'
            Initialize-LogGrid -Scope User -ConfigHash $configHash -SyncHash $syncHash 
        }

        else { $syncHash.SnackMsg.MessageQueue.Enqueue('Log path is invalid or unset; cannot view logs') }

    })

$syncHash.toolsLogStartDate.Add_CalendarClosed( { Set-FilteredLogs -SyncHash $syncHash -LogView $configHash.logCollectionView })

$syncHash.toolsLogEndDate.Add_CalendarClosed( { Set-FilteredLogs -SyncHash $syncHash -LogView $configHash.logCollectionView })

$syncHash.toolsLogSearchBox.add_KeyDown( { if ($_.Key -eq 'Enter') { Set-FilteredLogs -SyncHash $syncHash -LogView $configHash.logCollectionView } })

$syncHash.adHocConfirm.Add_Click({
    $configHash.confirmCode = 'continue'
    $syncHash.adHocConfirmWindow.IsOpen = $false
})

$syncHash.adHocConfirmCancel.Add_Click({
    $configHash.confirmCode = 'cancel'
    $syncHash.adHocConfirmWindow.IsOpen = $false
})

$syncHash.toolsLogDialogClose.Add_Click( { 
    $syncHash.toolsLogDataGrid.Items.Clear()
    $syncHash.toolsLogDialog.IsOpen = $false
})
#endregion

$syncHash.itemToolSearchBoxClear.Add_Click({ $syncHash.itemToolListSearchBox.Text = $null })
$syncHash.toolsLogSearchBoxClear.Add_Click({ $syncHash.toolsLogSearchBox.Text = $null })
$syncHash.itemToolGridSearchBoxClear.Add_Click({ $syncHash.itemToolGridSearchBox.Text = $null })

$syncHash.settingNameDialog.Add_DialogClosing( {
        $syncHash.settingNameDataGrid.Items.Refresh()
        $configHash.nameMapListView.Refresh()
    })

$syncHash.itemToolCommandGridClose.Add_Click( { $syncHash.itemToolDialog.IsOpen = $false }) 

$syncHash.toolsCommandGridExecuteAll.Add_Click( {
        $rsCmd = @{
            cmdGridItems       = $syncHash.itemToolCommandGridDataGrid.Items
            cmdGridParentIndex = $syncHash.itemToolCommandGridDataGrid.Items[0].ParentToolIndex
            item               = ($configHash.currentTabItem).toLower()
        }

        $rsArgs = @{
            Name            = 'CommandGridAllRun'
            ArgumentList    = @($syncHash.snackMsg.MessageQueue, $configHash, $queryHash, $syncHash, $rsCmd, $syncHash.adHocConfirmWindow, $syncHash.Window, $syncHash.adHocConfirmText, $varHash)
            ModulesToImport = $configHash.modList
        }

        Start-RSJob @rsArgs -ScriptBlock {
            Param($queue, $configHash, $queryHash, $syncHash, $rsCmd, $confirmWindow, $window, $textBlock, $varhash)

            Set-CustomVariables -VarHash $varHash -ConfigHash $configHash
            if ($configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].objectType -eq 'Standalone') {Remove-Variable ActiveObject, ActiveObjectType, ActiveObjectData -ErrorAction SilentlyContinue}
            $toolName = $configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].ToolName

            try {
                foreach ($gridItem in ($rsCmd.cmdGridItems | Where-Object { $_.Result -ne 'True' })) {
                    $actionCmd = ($configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].toolCommandGridConfig |
                            Where-Object { $_.actionCmdValid -eq 'True' -and $_.queryCmdValid -eq 'True' })[$gridItem.Index].actionCmd
                
                    $queryCmd = ($configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].toolCommandGridConfig |
                            Where-Object { $_.actionCmdValid -eq 'True' -and $_.queryCmdValid -eq 'True' })[$gridItem.Index].queryCmd
                
                    $result = $gridItem.result

                    ([scriptblock]::Create($actionCmd)).Invoke() 

                    $result = (Invoke-Expression $queryCmd).ToString()

                    $syncHash.Window.Dispatcher.Invoke([action] { $syncHash.itemToolCommandGridDataGrid.Items[$gridItem.Index].Result = $result })
                }

                $syncHash.Window.Dispatcher.Invoke([action] { $syncHash.itemToolCommandGridDataGrid.Items.Refresh() })

                if (($syncHash.itemToolCommandGridDataGrid.Items | Where-Object {$_.Result -notmatch 'True'} | Measure-Object).Count -eq 0){ $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.toolsCommandGridExecuteAll.Tag = 'False' }) }
                  
                if ($configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].objectType -eq 'Standalone') {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -Status Success
                    Write-LogMessage -syncHashWindow $syncHash.window -Path $configHash.actionlogPath -Message Succeed -ActionName $toolName -SubjectType 'Standalone' -ArrayList $configHash.actionLog
                }
                
                else {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -SubtoolName $rsCmd.cmdGridItemName -Status Success -SubjectName $activeObject
                    Write-LogMessage -syncHashWindow $syncHash.window -Path $configHash.actionlogPath -Message Succeed -SubjectName $activeObject -SubjectType $activeObjectType -ActionName $toolName -ArrayList $configHash.actionLog 
                }
            }    
         

            catch {
                if ($configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].objectType -eq 'Standalone') {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -Status Fail
                    Write-LogMessage -syncHashWindow $syncHash.window -Path $configHash.actionlogPath -SubjectType 'Standalone' -Message Fail -ActionName $toolName -ArrayList $configHash.actionLog -Error $_
                }
                else {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -Status Fail -SubjectName $activeObject
                    Write-LogMessage -syncHashWindow $syncHash.window -Path $configHash.actionlogPath -Message Fail -SubjectName $activeObject -SubjectType $activeObjectType -ActionName $toolName -ArrayList $configHash.actionLog -Error $_
                }
            }
        }
    })

$syncHash.settingExecute.Add_Click({
        if ($syncHash.settingUserPropGrid.Visibility -eq 'Visible') { $type = 'User' }

        else { $type = 'Comp' }
       
        $scriptTestArgs = @{
            ScriptBlock  =  $syncHash.('setting' + $type + 'PropGrid').SelectedItem.translationCmd
            ErrorControl = 'settingResultBox'
            SyncHash     = $syncHash
            ItemSet      = $configHash.($type + 'PropList')[($syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1)]
            StatusName   = 'ValidCmd'
        }

        Test-UserScriptBlock @scriptTestArgs

    })

$syncHash.settingAction2HidablePanel.Add_isEnabledChanged( {
        if ($syncHash.settingUserPropGrid.Visibility -eq 'Visible') { $type = 'User' }

        else { $type = 'Comp' }

        $configHash.($type + 'PropList')[($syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1)].actionCmd2Enabled = $syncHash.settingAction2Enable.IsChecked
    })


for ($i = 1; $i -le 2; $i++) {
    $syncHash.('settingBox' + $i + 'Execute').Add_Click( {
            param([Parameter(Mandatory)][Object]$sender)
            
            $id = $sender.Name -replace 'settingBox|Execute' 
            $type = $sender.DataContext.ItemType
          
            $scriptTestArgs = @{
                ScriptBlock  =  $syncHash.('setting' + $type + 'PropGrid').SelectedItem.('actionCmd' + $id)
                ErrorControl = ('settingBox' + $id + 'ResultBox')
                SyncHash     = $syncHash
                ItemSet      = $configHash.($type + 'PropList')[($syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1)]
                StatusName   = ('ValidAction' + $id)
            }

            Test-UserScriptBlock @scriptTestArgs

        }        
    )
}

$syncHash.settingContextExecute.Add_Click( {

        $scriptTestArgs = @{
            ScriptBlock  =  $syncHash.settingContextPropGrid.SelectedItem.actionCmd
            ErrorControl = 'settingContextResultBox'
            SyncHash     = $syncHash
            ItemSet      = $configHash.contextConfig[($syncHash.settingContextPropGrid.SelectedItem.IDNum - 1)]
            StatusName   = 'ValidAction'
        }

        Test-UserScriptBlock @scriptTestArgs

    })


$syncHash.settingTranslationScriptBlockBox.Add_TextChanged({ 
    if ($syncHash.settingTranslationScriptBlockBox.IsFocused -and $syncHash.settingResultBox.Tag -ne 'Unchecked') {

        if ($syncHash.settingUserPropGrid.Visibility -eq 'Visible') { $type = 'User' }

        else { $type = 'Comp' }
       
        $resetArgs = @{
            SyncHash      = $syncHash
            ResultBoxName = 'settingResultBox'
            ItemSet       = $configHash.($type + 'PropList')[($syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1)]
            StatusName    = 'ValidCmd'
        } 
      
        Reset-ScriptBlockValidityStatus @resetArgs

    }

})

$syncHash.settingAction1ScriptBlockBox.Add_TextChanged({ 
    if ($syncHash.settingAction1ScriptBlockBox.IsFocused -and $syncHash.settingBox1ResultBox.Tag -ne 'Unchecked') {
           
        
        if ($syncHash.settingUserPropGrid.Visibility -eq 'Visible') { $type = 'User' }

        else { $type = 'Comp' }

        $resetArgs = @{
            SyncHash      = $syncHash
            ResultBoxName = 'settingBox1ResultBox'
            ItemSet       = $configHash.($type + 'PropList')[($syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1)]
            StatusName    = 'ValidAction1'
        } 
      
        Reset-ScriptBlockValidityStatus @resetArgs

    }

})

$syncHash.settingAction2ScriptBlockBox.Add_TextChanged({ 
    if ($syncHash.settingAction2ScriptBlockBox.IsFocused -and $syncHash.settingBox2ResultBox.Tag -ne 'Unchecked') {
           
        
        if ($syncHash.settingUserPropGrid.Visibility -eq 'Visible') { $type = 'User' }

        else { $type = 'Comp' }

        $resetArgs = @{
            SyncHash      = $syncHash
            ResultBoxName = 'settingBox2ResultBox'
            ItemSet       = $configHash.($type + 'PropList')[($syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1)]
            StatusName    = 'ValidAction2'
        } 
      
        Reset-ScriptBlockValidityStatus @resetArgs

    }

})

$syncHash.settingContextScriptBlockBox.Add_TextChanged({
     if ($syncHash.settingContextScriptBlockBox.IsFocused -and $syncHash.settingContextResultBox.Tag -ne 'Unchecked') {
         $resetArgs = @{
                SyncHash      = $syncHash
                ResultBoxName = 'settingContextResultBox'
                ItemSet       = $configHash.contextConfig[($syncHash.settingContextPropGrid.SelectedItem.IDNum - 1)]
                StatusName    = 'ValidAction'
            } 
      
            Reset-ScriptBlockValidityStatus @resetArgs
    }
})

$syncHash.settingObjectToolExecute.Add_Click( {

           $scriptTestArgs = @{
            ScriptBlock  = $syncHash.settingObjectToolsPropGrid.SelectedItem.toolAction
            ErrorControl = 'settingObjectToolResultBox'
            SyncHash     = $syncHash
            ItemSet      = $configHash.objectToolConfig[($syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID - 1)]
            StatusName   = 'toolActionValid'
        }

        Test-UserScriptBlock @scriptTestArgs
    
    })

$syncHash.settingObjectToolScriptBox.Add_TextChanged({
     if ($syncHash.settingObjectToolScriptBox.IsFocused -and $syncHash.settingObjectToolResultBox.Tag -ne 'Unchecked') {
         $resetArgs = @{
                SyncHash      = $syncHash
                ResultBoxName = 'settingObjectToolResultBox'
                 ItemSet      = $configHash.objectToolConfig[($syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID - 1)]
                StatusName    = 'toolActionValid'
            } 
      
            Reset-ScriptBlockValidityStatus @resetArgs
    }
})


$syncHash.settingObjectToolSelectionExecute.Add_Click( {

           $scriptTestArgs = @{
            ScriptBlock  = $syncHash.settingObjectToolsPropGrid.SelectedItem.toolFetchCmd
            ErrorControl = 'settingObjectToolSelectionResultBox'
            SyncHash     = $syncHash
            ItemSet      = $configHash.objectToolConfig[($syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID - 1)]
            StatusName   = 'toolSelectValid'
        }

        Test-UserScriptBlock @scriptTestArgs
    
    })

$syncHash.settingObjectToolSelectionScriptBox.Add_TextChanged({
     if ($syncHash.settingObjectToolSelectionScriptBox.IsFocused -and $syncHash.settingObjectToolSelectionResultBox.Tag -ne 'Unchecked') {
         $resetArgs = @{
                SyncHash      = $syncHash
                ResultBoxName = 'settingObjectToolSelectionResultBox'
                ItemSet       = $configHash.objectToolConfig[($syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID - 1)]
                StatusName    = 'toolSelectValid'
            } 
      
            Reset-ScriptBlockValidityStatus @resetArgs
    }
})


$syncHash.settingObjectToolExtraExecute.Add_Click( {

           $scriptTestArgs = @{
            ScriptBlock  = $syncHash.settingObjectToolsPropGrid.SelectedItem.toolTargetFetchCmd
            ErrorControl = 'settingObjectToolExtraResultBox'
            SyncHash     = $syncHash
            ItemSet      = $configHash.objectToolConfig[($syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID - 1)]
            StatusName   = 'toolExtraValid'
        }

        Test-UserScriptBlock @scriptTestArgs
    
    })
$syncHash.settingObjectToolExtraScriptBox.Add_TextChanged({
     if ($syncHash.settingObjectToolExtraScriptBox.IsFocused -and $syncHash.settingObjectToolExtraResultBox.Tag -ne 'Unchecked') {
         $resetArgs = @{
                SyncHash      = $syncHash
                ResultBoxName = 'settingObjectToolExtraResultBox'
                ItemSet       = $configHash.objectToolConfig[($syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID - 1)]
                StatusName    = 'toolExtraValid'
            } 
      
            Reset-ScriptBlockValidityStatus @resetArgs
    }
})


$syncHash.settingNameDialogClose.Add_Click( {
        $syncHash.settingNameDialog.IsOpen = $false
        $configHash.nameMapListView.Refresh()
    })

$syncHash.settingInfoDialogClose.Add_Click({ 
    $syncHash.settingInfoDialog.IsOpen = $false 

    Start-RSJob -ArgumentList ($syncHash.settingInfoDialogScroller, $syncHash.Window) -ScriptBlock {
    param ($scroller, $window)
        Start-Sleep -Seconds 1 
        $window.Dispatcher.Invoke([action]{$scroller.ScrollToTop()})
    }
})

$syncHash.settingInfoDialogOpen.Add_Click( {
        if ($syncHash.settingInfoDialog.IsOpen) { $syncHash.settingInfoDialog.IsOpen = $false }
        else { $syncHash.settingInfoDialog.IsOpen = $true }
        Set-InfoPaneContent -SyncHash $syncHash -SettingInfoHash $settingInfoHash -ConfigHash $configHash
    })

$syncHash.SearchBoxButton.Add_Click( {
        if ($syncHash.SearchBox.Text.Length -eq 0) {
            $searchVal = (Select-ADObject -Type UsersComputers).FetchedAttributes -replace '$'
    
            if ($searchVal) {
                $syncHash.SearchBox.Tag = $searchVal
                $syncHash.SearchBox.Focus()
                $wshell = New-Object -ComObject wscript.shell
                $wshell.SendKeys('{ESCAPE}')
            }
        }
        else { $syncHash.SearchBox.Clear() }
    })
$syncHash.userQueryItem.Add_Click( {
        $searchVal = $syncHash.userCompGrid.SelectedItem.HostName
        if ($searchVal) {
            $syncHash.SearchBox.Tag = $searchVal
            $syncHash.SearchBox.Focus()
            $wshell = New-Object -ComObject wscript.shell
            $wshell.SendKeys('{ESCAPE}')
        }
    })

$syncHash.compQueryItem.Add_Click( {
        $searchVal = $syncHash.compUserGrid.SelectedItem.UserName
        if ($searchVal) {
            $syncHash.SearchBox.Tag = $searchVal
            $syncHash.SearchBox.Focus()
            $wshell = New-Object -ComObject wscript.shell
            $wshell.SendKeys('{ESCAPE}')
        }
    })

$syncHash.itemRefresh.Add_Click( {
        $configHash.itemRefreshing = $true
        $searchVal = $configHash.currentTabItem
        $configHash.currentTabItem = $null
        if ($searchVal) {
            $syncHash.tabControl.ItemsSource.RemoveAt($syncHash.tabControl.SelectedIndex)
            $syncHash.SearchBox.Tag = $searchVal
            $syncHash.SearchBox.Focus()
            $wshell = New-Object -ComObject wscript.shell
            $wshell.SendKeys('{ESCAPE}')
        }
    })

# Gets event handlers and adds them to their respective controls
Invoke-Expression -Command (Get-Content (Join-Path $baseConfigPath 'internal\eventhandlers.ps1') -Raw)

$syncHash.Window.ShowDialog() | Out-Null
$syncHash.Window.Close()
$configHash.IsClosed = $true