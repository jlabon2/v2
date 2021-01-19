

# TODO
#
#- help dialogs : textbox databound to hashtable, text triggered by title contents

#- general options item: searchBase, other search patterns, logging options
#- static variables / source item:
#- Color switches on glyph buttons???
# - categorzation horzontal scrollbar on dialog textbox


# credmgr package?
# Tool pane?
#
# https://github.com/JimmyCushnie/Font-Character-Lister






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
#Copy-Item -Path  C:\TempData\v3\MainWindow.xaml -Destination C:\TempData\MainWindow.xaml -Force
Copy-Item -Path "\\labtop\TempData\v3\v3\MainWindow.xaml"  -Destination C:\TempData\MainWindow.xaml -Force
$xamlPath = "C:\TempData\MainWindow.xaml"
$helpXAMLPath = "C:\TempData\HelpContent.xaml"
$glyphList = 'C:\TempData\segoeGlyphs.txt'

######

$savedConfig = "C:\TempData\config.json"

Import-Module C:\TempData\func\func.psm1 -DisableNameChecking
Remove-Module internal
Import-Module C:\TempData\internal\internal.psm1

# generated hash tables used throughout tool
New-HashTables

# Import from JSON and add to hash table
Set-Config -ConfigPath $savedConfig -Type Import -ConfigHash $configHash

# process loaded data or creates initial item templates for various config datagrids
@('userPropList','compPropList','contextConfig','objectToolConfig','nameMapList', 'netMapList', 'varListConfig', 'searchBaseConfig', 'queryDefConfig') | Set-InitialValues -ConfigHash $configHash -PullDefaults
@('userLogMapping','compLogMapping') | Set-InitialValues -ConfigHash $configHash

# matches config'd user/comp logins with default headers, creates new headers from custom values
$defaultList = @('User', 'DateRaw', 'LoginDc', 'ClientName', 'ComputerName', 'Ignore', 'Custom')
@('userLogMapping','compLogMapping') | Set-LoggingStructure -DefaultList $DefaultList -ConfigHash $configHash
Set-ActionLog -ConfigHash $configHash

# Add default values if they are missing
@('MSRA','MSTSC') | Set-RTDefaults -ConfigHash $configHash

# loaded required DLLs
foreach ($dll in ((Get-ChildItem C:\TempData\asm\ -Filter *.dll).FullName)) {
    [System.Reflection.Assembly]::LoadFrom($dll) | Out-Null   
}

# read xaml and load wpf controls into synchash (named synchash)
Set-WPFControls -TargetHash $syncHash -XAMLPath $xamlPath

Get-Glyphs -ConfigHash $configHash -GlyphList $glyphList
$syncHash.settingLogo.Source = "C:\tempdata\trident.png"
$syncHash.logoBackdrop.Source = "C:\tempdata\trident.png"
#Set-WPFControls -TargetHash $helpHash -XAMLPath $helpXAMLPath

# builds custom WPF controls from whatever was defined and saved in ConfigHash
Add-CustomItemBoxControls -SyncHash $syncHash -ConfigHash $configHash
Add-CustomToolControls -SyncHash $syncHash -ConfigHash $configHash
Add-CustomRTControls -SyncHash $syncHash -ConfigHash $configHash
$sysCheckHash.missingCheckFail = $false


#region Item Tool Events
#region Item Tools - Grid
$syncHash.ItemToolGridADSelectionButton.Add_Click({ Set-ADItemBox -ConfigHash $configHash -SyncHash $syncHash -Control Grid })

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

$synchash.itemToolGridSearchBox.Add_TextChanged({
    $syncHash.itemToolGridItemsGrid.ItemsSource.Filter = $null
    $syncHash.itemToolGridItemsGrid.ItemsSource.Filter = {param ($item) $item -match $syncHash.itemToolGridSearchBox.Text}
})

$syncHash.itemToolGridSelectConfirmCancel.Add_Click({ $syncHash.itemToolDialog.IsOpen = $false })

$syncHash.itemToolGridSelectAllButton.Add_Click({    
    if ($syncHash.itemToolGridItemsGrid.Items.Count -eq  $syncHash.itemToolGridItemsGrid.SelectedItems.Count) {
       $syncHash.itemToolGridItemsGrid.UnselectAll()
    } 
    else {
        $syncHash.itemToolGridItemsGrid.SelectAll()
    }      
})

#endregion

#region Item Tools - List
$syncHash.ItemToolADSelectionButton.Add_Click({ Set-ADItemBox -ConfigHash $configHash -SyncHash $syncHash -Control ListBox })

$syncHash.itemToolListSearchBox.Add_TextChanged({
    $syncHash.itemToolListSelectListBox.ItemsSource.Filter = $null
    $syncHash.itemToolListSelectListBox.ItemsSource.Filter = {param ($item) $item -match $syncHash.itemToolListSearchBox.Text}
})

$syncHash.itemToolListSelectConfirmButton.Add_Click({
    $itemList = $syncHash.itemToolListSelectListBox.SelectedItems.Name
    Start-ItemToolAction -ConfigHash $configHash -SyncHash $syncHash -Control List -ItemList $ItemList
})

$syncHash.itemToolGridSelectConfirmButton.Add_Click({
    $itemList = $syncHash.itemToolGridItemsGrid.SelectedItems
    Start-ItemToolAction -ConfigHash $configHash -SyncHash $syncHash -Control Grid -ItemList $ItemList
})

$syncHash.itemToolListSelectConfirmCancel.Add_Click({ $syncHash.itemToolDialog.IsOpen = $false })

$syncHash.itemToolListSelectAllButton.Add_Click({
    if ($syncHash.itemToolListSelectListBox.Items.Count -eq  $syncHash.itemToolListSelectListBox.SelectedItems.Count) {
       $syncHash.itemToolListSelectListBox.UnselectAll()
    } 
    else {
        $syncHash.itemToolListSelectListBox.SelectAll()
    }      
})
#endregion
#endregion

  
  

#region ChildWindow opening events

$syncHash.settingRemoteClick.add_Click({
    Set-ChildWindow -Panel settingRTContent -Title "Configure Remote Connection Clients" -SyncHash $syncHash
})

$syncHash.settingNetworkClick.add_Click({
    Set-ChildWindow -Panel settingNetContent -Title "Networking Mappings" -SyncHash $syncHash -Height 275
})

$syncHash.settingNamingClick.add_Click({ Set-ChildWindow -Panel settingNameContent -Title "Device Categorization" -SyncHash $syncHash -Height 275 })

$synchash.settingVarClick.add_Click({ Set-ChildWindow -Panel settingVarContent -Title "Resources and Variables" -SyncHash $syncHash -Height 275 -Width 530 })

$synchash.settingGeneralClick.add_Click({ Set-ChildWindow -Panel settingGeneralContent -Title "General Settings" -SyncHash $syncHash -Height 275 -Width 530 })

$syncHash.settingUserPropClick.add_Click({ Set-ChildWindow -Panel settingUserPropContent -Title "User Property Mappings" -SyncHash $syncHash -Height 365 -Width 600 })

$syncHash.settingCompPropClick.add_Click({ Set-ChildWindow -Panel settingCompPropContent -Title "Computer Property Mappings" -SyncHash $syncHash -Height 365 -Width 600 })

$syncHash.settingObjectToolsClick.add_Click({ Set-ChildWindow -Panel settingItemToolsContent -Title "Object Tools Mappings" -SyncHash $syncHash -Height 365 -Width 600 })

$syncHash.settingContextClick.add_Click({ Set-ChildWindow -Panel settingContextPropContent -Title "Contextual Actions Mappings"-SyncHash $syncHash -Height 365 -Width 390 })

$syncHash.settingLoggingClick.add_Click( {
    if (!($configHash.compLogPath) -or !(Test-Path $configHash.compLogPath)) {$syncHash.compLogPopupButton.IsEnabled = $false}
    if (!($configHash.userLogPath) -or !(Test-Path $configHash.userLogPath)) {$syncHash.userLogPopupButton.IsEnabled = $false}

    Set-ChildWindow -Panel settingLoggingContent -Title "Login Log Paths" -SyncHash $syncHash -Height 300 -Width 480
})

#endregion

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

        Start-RSJob -Name init -ArgumentList $syncHash, $psexec, $sysCheckHash, $configHash, $savedConfig -ModulesToImport ActiveDirectory, C:\TempData\internal\internal.psm1 -ScriptBlock {        
            Param($syncHash, $psExec, $sysCheckHash, $configHash, $savedConfig)
             
            Start-BasicADCheck -SysCheckHash $sysCheckHash -ConfigHash $configHash
                
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

            Start-Sleep -Seconds 2

            if ($sysCheckHash.sysChecks[0].ADDS -eq $false -or 
                $sysCheckHash.sysChecks[0].Modules -eq $false -or 
                $sysCheckHash.sysChecks[0].Admin -eq $false) {

                Suspend-FailedItems -SyncHash $syncHash -CheckedItems SysCheck
            }
                
            else { 
                Start-PropBoxPopulate -ConfigHash $configHash 
                Set-ADGenericQueryNames -ConfigHash $configHash               
                Set-QueryPropertyList -SyncHash $syncHash -ConfigHash $configHash
                
            }  
                
            Show-WPFWindow -SyncHash $syncHash     
        }
    }
    else {                                   
        Suspend-FailedItems -SyncHash $syncHash -CheckedItems RSCheck
        Show-WPFWindow -SyncHash $syncHash
    }
})

#endregion 



#region RemoteTools events

$syncHash.settingRtRDPClick.Add_Click( { Set-StaticRTContent -SyncHash $syncHash -ConfigHash $configHash -Tool MSTSC  })

$syncHash.settingRtMSRAClick.Add_Click({ Set-StaticRTContent -SyncHash $syncHash -ConfigHash $configHash -Tool MSRA })

$syncHash.settingRemoteFlyout.Add_OpeningFinished({ Get-RTFlyoutContent -ConfigHash $configHash -SyncHash $syncHash })

$syncHash.settingRemoteFlyoutExit.Add_Click( {
    Reset-ChildWindow -SyncHash $syncHash -Title "Configure Remote Connection Clients" -SkipContentPaneReset -SkipResize
    Set-SelectedRTTypes -SyncHash $syncHash -ConfigHash $configHash
})


$syncHash.settingRtExeSelect.Add_Click( {
    Get-RTExePath -SyncHash $syncHash -ConfigHash $configHash    
})

$syncHash.settingRtAddClick.Add_Click({

    if (!($syncHash.customRt)) { $syncHash.customRt = @{} }
    if (!$configHash.rtConfig) { $configHash.rtConfig = @{} }

    $rtID = "rt" + [string]([int](((($configHash.rtConfig.Keys | Where-Object {$_ -like "RT*"}) -replace 'rt') |
         Sort-Object -Descending | Select-Object -First 1)) + 1) 
       
    New-CustomRTConfigControls -SyncHash $syncHash -ConfigHash $configHash -RTID $rtID -NewTool

})

#endregion

#region NetworkMap events

$syncHash.settingNetMapClick.Add_Click({
    $syncHash.settingNetDataGrid.Visibility = "Visible"
    $syncHash.settingNetDataGrid.ItemsSource = $configHash.netMapList 
    Set-ChildWindow -SyncHash $syncHash -Title "Network Mappings" -HideCloseButton -Background Flyout
    $syncHash.settingNetFlyout.IsOpen = $true
})

$syncHash.settingNetAddClick.Add_Click({ Set-NetworkMapItem -SyncHash $syncHash -ConfigHash $configHash })

$syncHash.settingNetImportClick.Add_Click({ Set-NetworkMapItem -SyncHash $syncHash -ConfigHash $configHash -Import })

$syncHash.settingNetFlyoutExit.Add_Click( {
    $syncHash.settingNetFlyout.IsOpen = $false
    Set-ChildWindow -SyncHash $syncHash -Background Standard -Title "Networking Mappings"        
       
    [Array]$configHash.netMapList | ForEach-Object {
        if ($_.ValidMask -ne $true -and $_.ValidNetwork -ne $true) {
            [Array]$configHash.netMapList.RemoveAt([Array]::IndexOf($configHash.netMapList.ID,$_.ID))
        }
    }

    $configHash.Remove('NetMapListView')

})


#endregion


#region UserLog events
$syncHash.userLogPopupButton.Add_Click({ Set-LogMapGrid -SyncHash $syncHash -ConfigHash $configHash -Type User })

$syncHash.compLogPopupButton.Add_Click({ Set-LogMapGrid -SyncHash $syncHash -ConfigHash $configHash -Type Comp })
#endregion

#region Naming events
$syncHash.settingNameMapClick.Add_Click( {
    if ($configHash.nameMapList.GetType().Name -ne 'ListCollectionView') {
        $configHash.nameMapListView = [System.Windows.Data.ListCollectionView]$configHash.nameMapList 
        $configHash.nameMapListView.IsLiveSorting = $true
       $configHash.nameMapListView.SortDescriptions.Add((New-Object System.Windows.Data.PropertySortDescription "Date"))
        
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

$syncHash.settingCommandGridAddClick.Add_Click({
     $configHash.objectToolConfig[([Array]::IndexOf($configHash.objectToolConfig, $syncHash.settingObjectToolsPropGrid.SelectedItem))].toolCommandGridConfig.Add([PSCustomObject]@{
       ToolID      =  ($syncHash.settingObjectToolsPropGrid.SelectedItem.toolCommandGridConfig.toolID | Sort-Object -Descending | Select-Object -First 1) + 1
       SetName     =  $null
       queryCmd    =  'Check-Something -UserName $user'
       actionCmd   = 'Do-Something -UserName $user'
	   actionCmdValid = "null"
	   queryCmdValid  = "null"
    })
    
    $syncHash.settingCommandGridDataGrid.Items.Refresh()    

})

$syncHash.settingVarAddClick.Add_Click( {
    
    $configHash.varListConfig.Add([PSCustomObject]@{
        VarNum              = ($configHash.varListConfig.VarNum | Sort-Object -Descending | Select-Object -First 1) + 1
        VarName             = $null
        VarCmd              = $null
        UpdateFrequencyList = @('All Queries','User Queries','Comp Queries','Daily','Hourly','Every 15 mins','Program Start')
        UpdateFrequency     = 'Program Start'
    })    

    $syncHash.settingVarDataGrid.Items.Refresh()                   

})

$syncHash.settingGeneralAddClick.Add_Click( {
    
    if ($syncHash.settingGeneralAddClick.Tag -eq 'OU') {
        $configHash.searchBaseConfig.Add([PSCustomObject]@{
            OUNum          = ($configHash.searchBaseConfig.OUNum | Sort-Object -Descending | Select-Object -First 1) + 1
            OU             = $null
            QueryScopeList  = @('Base','OneLevel','Subtree')
            QueryScope      = 'Subtree'
        })    

        $syncHash.settingOUDataGrid.Items.Refresh()
    }
    
    else {
   
        $configHash.queryDefConfig.Add([PSCustomObject]@{
            ID                = ($configHash.queryDefConfig.ID | Sort-Object -Descending | Select-Object -First 1) + 1
            Name              = $null
            QueryDefTypeList  = $configHash.QueryADValues.Key
            QueryDefType      = $null
        })    

        $syncHash.settingQueryDefDataGrid.Items.Refresh()
    }                   

})

$syncHash.settingVarDialogClose.Add_Click({
    $syncHash.settingVarDialog.IsOpen = $false 
    $synchash.settingVarDataGrid.Items.Refresh()
})

$syncHash.settingCommandGridDialogClose.Add_Click({
    $syncHash.settingCommandGridDialog.IsOpen = $false
    $syncHash.settingCommandGridDataGrid.Items.Refresh()
})

#endregion

#region ContextTools

#endregion

#region VarConfig
$syncHash.settingVarMapClick.Add_Click({ 
    Set-ChildWindow -SyncHash $syncHash -Title "Variable Definitions" -Panel settingVarContent -HideCloseButton -Background Flyout
    $syncHash.settingVarFlyout.IsOpen = $true
})

$syncHash.settingOUMapClick.Add_Click({ 
    Set-ChildWindow -SyncHash $syncHash -Title "Search Base Definitions" -Panel settingOUDataGrid -HideCloseButton -Background Flyout
    $syncHash.settingGeneralFlyout.IsOpen = $true
})


$syncHash.settingQueryMapClick.Add_Click({ 
    Set-ChildWindow -SyncHash $syncHash -Title "Query Definitions" -Panel settingQueryDefDataGrid -HideCloseButton -Background Flyout
    $syncHash.settingGeneralFlyout.IsOpen = $true
})

$syncHash.settingMiscClick.Add_Click({
    Set-ChildWindow -SyncHash $syncHash -Title "Misc. Settings" -Panel settingMiscGrid -HideCloseButton -Background Flyout
    
    if (!$configHash.actionlogPath) {$configHash.actionlogPath = "C:\Logs" }
    $syncHash.settingLogPath.Text = $configHash.actionlogPath 
    $syncHash.settingGeneralFlyout.IsOpen = $true
})

#endregion

#region FlyOutExits
$syncHash.settingLoggingCompFlyoutExit.Add_Click({ Reset-ChildWindow -SyncHash $syncHash -Title "Login Log Paths" -SkipContentPaneReset -SkipResize })

$syncHash.settingVarFlyoutExit.Add_Click({ Reset-ChildWindow -SyncHash $syncHash -Title "Resources and Variables" -SkipContentPaneReset -SkipResize })

$syncHash.settingLoggingUserFlyoutExit.Add_Click({ Reset-ChildWindow -SyncHash $syncHash -Title "Login Log Paths" -SkipContentPaneReset -SkipResize })

$syncHash.settingObjectToolFlyoutExit.Add_Click({ Reset-ChildWindow -SyncHash $syncHash -Title "Object Tools Mappings" -SkipContentPaneReset -SkipResize -SkipDataGridReset })

$syncHash.settingNameFlyoutExit.Add_Click({ Reset-ChildWindow -SyncHash $syncHash -Title "Computer Categorization" -SkipContentPaneReset -SkipResize })

$syncHash.settingGeneralFlyoutExit.Add_Click({
    if ($syncHash.settingGeneralAddClick.Tag -eq 'null' -and (Test-Path $syncHash.settingLogPath.Text)) { $configHash.actionlogPath = $syncHash.settingLogPath.Text }
    Reset-ChildWindow -SyncHash $syncHash -Title "General Settings" -SkipContentPaneReset -SkipResize    
})

$syncHash.settingContextDefFlyout.Add_OpeningFinished({
    $syncHash.settingContextListTypes.ItemsSource = $configHash.nameMapList
    
    $syncHash.settingContextListTypes.Items | Where-Object {$_.Name -in $configHash.contextConfig[$syncHash.settingContextPropGrid.SelectedItem.IDNUM - 1].Types} |
        ForEach-Object { $syncHash.settingContextListTypes.SelectedItems.Add(($_)) }
        
})

$syncHash.settingContextFlyoutExit.Add_Click({
    $configHash.contextConfig[$syncHash.settingContextPropGrid.SelectedItem.IDNUM - 1].Types = @()
    
    $syncHash.settingContextListTypes.SelectedItems | 
        ForEach-Object { $configHash.contextConfig[$syncHash.settingContextPropGrid.SelectedItem.IDNUM - 1].Types += $_.Name } 

    Reset-ChildWindow -SyncHash $syncHash -Title "Contextual Action Buttons" -SkipContentPaneReset -SkipResize -SkipDataGridReset
    $syncHash.settingRemoteListTypes.SelectedItems.Clear()   

})

$syncHash.settingFlyoutExit.Add_Click( {
    if ($syncHash.settingUserPropContent.Visibility -eq 'Visible') { $type = 'User' }
    else { $type = 'Computer' }

    Reset-ChildWindow -SyncHash $syncHash -Title "$type Property Mappings" -SkipContentPaneReset -SkipResize -SkipDataGridReset

})


#endregion

$syncHash.settingActionPathClick.Add_Click({ $syncHash.settingLogPath.Text = New-FolderSelection -Title "Select tool's logging directory" })
        
$syncHash.settingLoggingPcPathClick.Add_Click({ Set-LoggingDirectory -SyncHash $syncHash -ConfigHash $configHash -Type Comp })

$syncHash.settingLoggingUserPathClick.Add_Click({ Set-LoggingDirectory -SyncHash $syncHash -ConfigHash $configHash -Type User })
    
#endregion

$syncHash.settingChildWindow.add_ClosingFinished({ Reset-ChildWindow -SyncHash $syncHash })

#region Category button events
$syncHash.settingModClick.add_Click({ Set-ChildWindow -SyncHash $syncHash -Panel settingModContent -Title "Required PS Modules"  })

$syncHash.settingADClick.add_Click({ Set-ChildWindow -SyncHash $syncHash -Panel settingADContent -Title "ADDS" })

$syncHash.settingPermClick.add_Click({ Set-ChildWindow -SyncHash $syncHash -Panel settingAdminContent -Title "Admin Permissions" }) 

#endregion Category button events
#region Action pane exit events

$syncHash.settingCloseClick.add_Click({ $syncHash.Window.Close() })
    
$syncHash.settingConfigCancelClick.add_Click({ $syncHash.Window.Close()})

$syncHash.settingImportClick.add_Click({ Import-Config -SyncHash $syncHash })

$syncHash.settingConfigClick.add_Click( {
    Set-Config -ConfigPath $savedConfig -Type Export -ConfigHash $configHash
    
    Start-Sleep -Seconds 1
    $syncHash.Window.Close()
    Start-Process -WindowStyle Hidden -FilePath "$PSHOME\powershell.exe" -ArgumentList " -ExecutionPolicy Bypass -NonInteractive -File $($PSCommandPath)"
    exit
})

#endregion

###### MAIN WINDOW
$syncHash.TabMenu.add_SelectionChanged( {
       $syncHash.TabMenu.Items | ForEach-Object { $_.Background = "#FF444444" }
       $syncHash.tabMenu.SelectedItem.Background = "#576573"
       $syncHash.TabMenu.Items.Header | ForEach-Object { $_.Foreground = "Gray" }
       $syncHash.tabMenu.SelectedItem.Header.Foreground = "AliceBlue"

       $syncHash.historySideDataGrid.ItemsSource = $configHash.actionLog
       $syncHash.historySideDataGrid.Items.Refresh()

        if ($syncHash.tabMenu.SelectedItem.Tag -eq 'Query') {
            $syncHash.SearchBox.Focus()
        }

        if ($syncHash.tabMenu.SelectedItem.Tag -eq 'Console') {         
           $syncHash.consoleControl.StartProcess(("$PSHOME\powershell.exe"))            
        }

        elseif ($syncHash.consoleControl.IsProcessRunning) {
            $syncHash.consoleControl.StopProcess()
        }


    })

$syncHash.searchBoxHelp.add_MouseLeftButtonDown({$syncHash.searchBoxHelp.Background = "#576573" })

$syncHash.searchBoxHelp.add_MouseLeftButtonUp({
    $syncHash.childHelp.isOpen = $true  
    $syncHash.searchBoxHelp.Background = "Transparent"
})

$syncHash.HistoryToggle.Add_MouseLeftButtonUp({
    if ($syncHash.historySidePane.IsOpen) { $syncHash.historySidePane.IsOpen = $false }
    else { $syncHash.historySidePane.IsOpen = $true}
    $syncHash.historySideDataGrid.Items.Refresh()
})


$syncHash.tabControl.ItemsSource = New-Object System.Collections.ObjectModel.ObservableCollection[Object]

$syncHash.tabMenu.add_Loaded( {
    if ($sysCheckHash.sysChecks.RSModule -eq $true) {
        Start-RSJob -Name menuLoad -ArgumentList $syncHash, $savedConfig, $sysCheckHash -ModulesToImport C:\TempData\internal\internal.psm1 -ScriptBlock {        
            Param($syncHash, $savedConfig, $sysCheckHash)

            do {} until ($sysCheckHash.checkComplete)
           
            if ($sysCheckHash.sysChecks.Admin -eq $false -or $sysCheckHash.sysChecks.ADDS -eq $false  -or $sysCheckHash.sysChecks.Modules -eq $false ) {
                Suspend-FailedItems -SyncHash $syncHash -CheckedItems SysCheck  
            }

            elseif (!(Test-Path $savedConfig)) {
                Suspend-FailedItems -SyncHash $syncHash -CheckedItems Config
            }
        
            else {  
                $syncHash.Window.Dispatcher.invoke([action] {$syncHash.tabMenu.SelectedIndex = 0 })                  
            }
        }
       
        New-VarUpdater -ConfigHash $configHash
        Start-VarUpdater -ConfigHash $configHash -VarHash $varHash
    
    }

    if (Test-Path $savedConfig) {              
        foreach ($type in @('User', 'Comp')) {             
            if (!($configHash.($type + 'PropListSelection'))) { $configHash.($type + 'PropListSelection') = @()}

            for ($i = 1; $i -le $configHash.boxMax; $i++) {
            
                if ($i -le $configHash.($type + 'boxCount')) {
                    Remove-Variable Selected -ErrorAction SilentlyContinue
                    $selected = $configHash.($type + 'PropList') | Where-Object { $_.Field -eq $i }
       
                    if ($null -ne $selected.FieldName) { 
                        $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Header').Content = $selected.FieldName
                        $configHash.($type + 'PropListSelection') += $selected.PropName

                        switch ($selected.ActionName) {

                            {$selected.ActionName -notmatch "Editable"} {
                                $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'TextBox').Tag = 'NoEdit'
                                $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'EditClip').Visibility = "Collapsed"
                                $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1').Visibility = "Collapsed"                                                   
                            }

                            {$selected.ActionName -notmatch "Actionable"} {
                                $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1Action1').Visibility = "Collapsed"  
                                $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1Action2').Visibility = "Collapsed"  
                            }

                            {$selected.ActionName -match "Editable"} {
                                
                                $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Box1').Add_Click({
                                param([Parameter(Mandatory)][Object]$sender)                       
                                    if ($sender.Name -like "*ubox*") {
                                        $id = $sender.Name -replace "ubox|box.*"
                                        $type = "User"
                                    }
                                    else {
                                        $id = $sender.Name -replace "cbox|box.*"
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
                                                Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName Edit -SubjectName $editObject -SubjectType $type -ArrayList $configHash.actionLog 
                                            }
                                    
                                            catch { 
                                                 Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName Edit -SubjectName $editObject -SubjectType $type -ArrayList $configHash.actionLog -Error $_
                                                $syncHash.SnackMsg.MessageQueue.Enqueue("[EDIT]: FAILED on [$($($editObject).toLower())]") }
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
                                                    Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName Clear -SubjectName $editObject -SubjectType $type -ArrayList $configHash.actionLog 
                                                }

                                                catch { 
                                                    Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName Clear -SubjectName $editObject -SubjectType $type -ArrayList $configHash.actionLog -Error $_
                                                    $syncHash.SnackMsg.MessageQueue.Enqueue("[CLEAR]: FAIL on [$($($editObject).toLower())]") 
                                                }
                                            }

                                            else {     
                                                Write-LogMessage -Path $configHash.actionlogPath -Message Fail -ActionName Edit -SubjectName $editObject -SubjectType $type -ArrayList $configHash.actionLog -Error $_
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

                            {$selected.ActionName -match "Actionable"} {
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
    
                               
                                            Start-RSJob -Name threadedAction -ArgumentList $rsCmd, $queryHash, $b, $configHash -ModulesToImport  C:\TempData\internal\internal.psm1 -ScriptBlock {
                                                Param($rsCmd, $queryHash, $b, $configHash)
                                                        
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
                                                    Write-LogMessage -syncHashWindow $rsCmd.Window -Path $configHash.actionlogPath -Message Succeed -ActionName $actionNameString -SubjectName $actionObjectString -SubjectType $type -ArrayList $configHash.actionLog 

                                                    if ($rsCmd.propList.('actionCmd' + $b + 'Refresh')) {
                                                        if ($PropName -ne 'Non-Ad Property') {
                                                            if ($type -eq 'User') {                               
                                                                $result = (Get-ADUser -Identity $user -Properties $propName).$propName
                                                                $queryHash.($user).($propName) = $result                                                                       
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

                                                            else {  $updatedValue = $result }

                                                            $rsCmd.Window.Dispatcher.Invoke([Action] { $rsCmd.boxResources.($type[0] + 'box' + $rsCmd.id + "TextBox").Text = $updatedValue })                                        
                                                        }
                                                    }
                                                }
                          
                                                catch {
                                                    $rsCmd.Window.Dispatcher.Invoke([Action] { $rsCmd.snackMsg.MessageQueue.Enqueue("$actionNameString FAILURE on $actionObjectString") })
                                                     Write-LogMessage -Path $configHash.actionlogPath -Message Fail -ActionName $actionNameString -SubjectName $actionObjectString -SubjectType $type -ArrayList $configHash.actionLog -Error $_
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
                                                        Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName $actionName -SubjectName (Get-Variable -Name $type -ValueOnly) -SubjectType $type -ArrayList $configHash.actionLog 

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

                                                            else { $updatedValue = $result }

                                                            $syncHash.($type[0] + 'box' + $id + "resources").($type[0] + 'box' + $id + "TextBox").Text = $updatedValue
                                                        }                       
                                                    }
                          
                                                    catch {
                                                        $syncHash.SnackMsg.MessageQueue.Enqueue("[$($($actionName).toUpper())] FAILURE on [$($(Get-Variable -Name $type -ValueOnly).toLower())]")
                                                        Write-LogMessage -Path $configHash.actionlogPath -Message Fail -ActionName $actionName -SubjectName (Get-Variable -Name $type -ValueOnly) -SubjectType $type -ArrayList $configHash.actionLog -Error $_
                                                    }                                                               
                                                })
                                        }
                                    }
                    
                                    else {
                                        $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + ('Box1Action' + $b)).Visibility = "Collapsed"  
                                    }
                                }
                            }
                        }
                    }

                    else {
                        $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Header').Visibility = "Collapsed"
                        $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Border').Visibility = "Collapsed"
                    }          
                }
            }
        
            if ($configHash.($type + 'boxCount') -le 7) { $syncHash.Window.Dispatcher.Invoke([Action]{$syncHash.($type + 'detailGrid').Columns = 1 })}

            elseif ($configHash.($type + 'boxCount') -gt 7 -and $configHash.($type + 'boxCount') -le 14) { $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.($type + 'detailGrid').Columns = 2})}
            
            else { $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.($type + 'detailGrid').Columns = 3 }) }

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

                        Start-RSJob -ArgumentList $rsCmd, $syncHash, $configHash -ModulesToImport C:\TempData\internal\internal.psm1 -ScriptBlock {
                            Param($rsCmd, $syncHash, $configHash)
                            
                            $user = $rsCmd.user
                            $comp = $rsCmd.comp

                            try {
                                Invoke-Expression -Command $rsCmd.buttonSettings.actionCmd
                                $syncHash.Window.Dispatcher.Invoke([Action]{
                                    $syncHash.SnackMsg.MessageQueue.Enqueue("[$($rsCmd.buttonSettings.ActionName.ToUpper())]: SUCCESS on $($rsCmd.user.ToLower()) on $($rsCmd.comp.ToLower())")
                                    Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName ($rsCmd.buttonSettings.ActionName.ToUpper()) -SubjectName $rsCmd.User -SubjectType Context -ArrayList $configHash.actionLog 
                                })
                            }
                            catch {
                                $syncHash.Window.Dispatcher.Invoke([Action]{
                                    $syncHash.SnackMsg.MessageQueue.Enqueue("[$($rsCmd.buttonSettings.ActionName.ToUpper())]: FAIL on $($rsCmd.user.ToLower()) on $($rsCmd.comp.ToLower())")  
                                    Write-LogMessage -Path $configHash.actionlogPath -Message Fail -ActionName ($rsCmd.buttonSettings.ActionName.ToUpper()) -SubjectName $rsCmd.User -SubjectType Context -ArrayList $configHash.actionLog -Error $_
                                })
                            }
                        }
                    }

                    else {

                        try {
                            Invoke-Expression -Command $configHash.contextConfig[$id - 1].actionCmd
                            $syncHash.SnackMsg.MessageQueue.Enqueue("[$($configHash.contextConfig[$id - 1].ActionName.ToUpper())]: SUCCESS on $($user.ToLower()) on $($comp.ToLower())")
                            Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName ($configHash.contextConfig[$id - 1].ActionName.ToUpper()) -SubjectName $user -SubjectType Context -ArrayList $configHash.actionLog 
                        }
                            
                        catch {
                            $syncHash.SnackMsg.MessageQueue.Enqueue("[$($configHash.contextConfig[$id - 1].ActionName.ToUpper())]: FAIL on $($user.ToLower()) on $($comp.ToLower())")
                            Write-LogMessage -Path $configHash.actionlogPath -Message Fail -ActionName ($configHash.contextConfig[$id - 1].ActionName.ToUpper()) -SubjectName $user -SubjectType Context -ArrayList $configHash.actionLog -Error $_
                        }
                            

                    }

                })
                
                $syncHash.($buttonType + 'ContextGrid').AddChild($syncHash.customContext.('cxt' + $contextBut.IDNum).($buttonType + 'context' + $contextBut.IDNum))

            }
        }

        $customField = 0
        $sizeToExpand = ([math]::Ceiling(($configHash.userLogMapping | Where-Object {$_.FieldSel -ne 'Ignore' }).Count/5) - 1) * 55 + $syncHash.userCompControlRow.Height.Value
        $syncHash.Window.Dispatcher.Invoke([Action]{$syncHash.userCompControlRow.Height = $sizeToExpand})
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
        $sizeToExpand = ([math]::Ceiling(($configHash.compLogMapping | Where-Object {$_.FieldSel -ne 'Ignore' }).Count/5) - 1) * 55 + $syncHash.compUserControlRow.Height.Value
        $syncHash.Window.Dispatcher.Invoke([Action]{$syncHash.compUserControlRow.Height = $sizeToExpand}) 
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
        
        $configHash.currentTabItem = $syncHash.tabControl.SelectedItem.Name
        $queryHash.Keys | ForEach-Object { $queryHash.$_.ActiveItem = $false }
        
        
        
         $currentTabItem = $syncHash.tabControl.SelectedItem.Name
    
        Start-RSJob -Name "VisualChange" -ArgumentList  $syncHash, $queryHash, $configHash, $currentTabItem -ScriptBlock {
            Param($syncHash, $queryHash, $configHash, $currentTabItem)
                
                $queryHash.$currentTabItem.ActiveItem = $true

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
        Set-ItemExpanders -SyncHash $syncHash -ConfigHash $configHash -IsActive Disable
    }
       
})

$syncHash.userCompGrid.Add_SelectionChanged({ Set-GridButtons -SyncHash $syncHash -ConfigHash $configHash -Type User })

$syncHash.compUserGrid.Add_SelectionChanged({ Set-GridButtons -SyncHash $syncHash -ConfigHash $configHash -Type Comp })

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

$syncHash.userCompFocusClientToggle.Add_Checked({
    $syncHash.userCompFocusHostToggle.IsChecked = $false   
    Set-ClientGridButtons -SyncHash $syncHash -ConfigHash $configHash -Type User
})

$syncHash.searchSettingsPopUp.Add_Closed({
    $configHash.queryProps = [System.Collections.ArrayList]@()
    Get-LDAPSearchNames -ConfigHash $configHash -SyncHash $syncHash | ForEach-Object {$configHash.queryProps.Add($_) | Out-Null}
})

$syncHash.userCompFocusClientToggle.Add_Unchecked({ $syncHash.userCompFocusHostToggle.IsChecked = $true })

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

$syncHash.compUserFocusClientToggle.Add_Unchecked( {
        $syncHash.compUserFocusUserToggle.IsChecked = $true      
    })



$syncHash.SearchBox.add_KeyDown( {
 

    if ($_.Key -eq "Enter" -or $_.Key -eq 'Escape') {
        if ($null -like $syncHash.SearchBox.Text -and $_.Key -ne 'Escape') { $syncHash.SnackMsg.MessageQueue.Enqueue("Empty!") }
           
        elseif ($syncHash.SearchBox.Text.Length -ge 3 -or $_.Key -eq 'Escape') { 
            if (!($configHash.queryProps)) { 
                $configHash.queryProps = [System.Collections.ArrayList]@()
                Get-LDAPSearchNames -ConfigHash $configHash -SyncHash $syncHash | ForEach-Object {$configHash.queryProps.Add($_) | Out-Null} 
            }
            Start-ObjectSearch -SyncHash $syncHash -ConfigHash $configHash -QueryHash $queryHash -Key $_.Key }

        else { $syncHash.SnackMsg.MessageQueue.Enqueue("Query must be at least 3 characters long!") }
    } 
})

$syncHash.itemToolDialogConfirmButton.Add_Click({

    Start-RSJob -Name ItemTool -ArgumentList $syncHash.snackMsg.MessageQueue, $syncHash.itemToolDialogConfirmButton.Tag, $configHash, $queryHash, $syncHash.Window -ModulesToImport C:\TempData\internal\internal.psm1  -ScriptBlock {
    Param($queue, $toolID, $configHash, $queryHash, $window)

        $item = ($configHash.currentTabItem).toLower()
        $toolName = ($configHash.objectToolConfig[$toolID - 1].toolActionToolTip).ToUpper()

        try {                     
            Invoke-Expression $configHash.objectToolConfig[$toolID - 1].toolAction
            if ($configHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {
                $queue.Enqueue("[$toolName]: Success - Standalone tool complete")
                Write-LogMessage -syncHashWindow $window -Path $configHash.actionlogPath -Message Succeed -ActionName $toolName -SubjectType 'Standalone' -ArrayList $configHash.actionLog
            }

            else {
                $queue.Enqueue("[$toolName]: Success on [$item] - tool complete")
                Write-LogMessage -syncHashWindow $window -Path $configHash.actionlogPath -Message Succeed -SubjectName $item -ActionName $toolName -ArrayList $configHash.actionLog 
            }
                                           
        }
        catch {
            if ($configHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {
                $queue.Enqueue("[$toolName]: Fail - Standalone tool incomplete") 
                Write-LogMessage -syncHashWindow $Window -Path $configHash.actionlogPath -Message Fail -ActionName $toolName -ArrayList $configHash.actionLog -Error $_
            }
            else {
                $queue.Enqueue("[$toolName]: Fail on [$item] - tool incomplete")
                Write-LogMessage -syncHashWindow $Window -Path $configHash.actionlogPath -Message Fail -SubjectName $item -ActionName $toolName -SubjectType 'Standalone' -ArrayList $configHash.actionLog -Error $_
            }
        }
    }

    $syncHash.itemToolDialog.IsOpen = $false
})

$syncHash.itemToolDialogConfirmCancel.Add_Click({ $syncHash.itemToolDialog.IsOpen = $false })

$syncHash.itemToolDialog.Add_ClosingFinished({
    $syncHash.itemToolDialogConfirm.Visibility = "Collapsed"
    $syncHash.itemToolListSelect.Visibility = "Collapsed"
    $syncHash.itemToolListSelectListBox.ItemsSource = $null
    $syncHash.itemToolADSelectedItem.Content = $null
    $syncHash.itemToolGridADSelectedItem.Content = $null
    $syncHash.itemToolImageBorder.Visibility = "Collapsed"
    $syncHash.itemToolGridSelect.Visibility = "Collapsed"
    $syncHash.itemToolCommandGridPanel.Visibility = "Collapsed"

})


#region ActionLogReview

$syncHash.toolsSingleLogExport.Add_Click({ 
    $syncHash.toolsExportDialog.isOpen = $true
    $syncHash.toolsExportDialog.Tag = 'Single'
})

$syncHash.toolsAllLogExport.Add_Click({ 
    $syncHash.toolsExportDialog.isOpen = $true
    $syncHash.toolsExportDialog.Tag = 'All'
})

$syncHash.toolsExportConfirmButton.Add_Click({
    if ($syncHash.toolsExportDialog.Tag -eq 'Single') { New-LogHTMLExport -Scope User -ConfigHash $configHash -TimeFrame $syncHash.toolsExportDate.SelectedValue.Content }
    else { New-LogHTMLExport -Scope All -ConfigHash $configHash -TimeFrame $syncHash.toolsExportDate.SelectedValue.Content }

    $syncHash.SnackMsg.MessageQueue.Enqueue("Log export started") 
    $syncHash.toolsExportDialog.isOpen = $false
})

$syncHash.toolsExportConfirmCancel.Add_Click({ $syncHash.toolsExportDialog.isOpen = $false })

$syncHash.toolsAllLogView.Add_Click({ 
    $syncHash.toolsLogProgress.Tag = "Unloaded"
    $syncHash.toolsLogEmpty.Tag = "Unloaded"
    Initialize-LogGrid -Scope All -ConfigHash $configHash -SyncHash $syncHash
})

$syncHash.toolsSingleLogView.Add_Click({ 
    $syncHash.toolsLogProgress.Tag = "Unloaded"
    $syncHash.toolsLogEmpty.Tag = "Unloaded"
    Initialize-LogGrid -Scope User -ConfigHash $configHash -SyncHash $syncHash 
})

$syncHash.toolsLogStartDate.Add_CalendarClosed({ Set-FilteredLogs -SyncHash $syncHash -LogView $configHash.logCollectionView })

$syncHash.toolsLogEndDate.Add_CalendarClosed({ Set-FilteredLogs -SyncHash $syncHash -LogView $configHash.logCollectionView })

$syncHash.toolsLogSearchBox.add_KeyDown( { if ($_.Key -eq "Enter") { Set-FilteredLogs -SyncHash $syncHash -LogView $configHash.logCollectionView } })





$syncHAsh.toolsLogDialogClose.Add_Click({
    $syncHash.toolsLogDialog.IsOpen = $false
})
#endregion






[System.Windows.RoutedEventHandler]$eventonNetDataGrid = {
    $button = $_.OriginalSource
   

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

[System.Windows.RoutedEventHandler]$eventonVarDataGrid = {
    $button = $_.OriginalSource

    if ($button.Name -match "settingVarClearItem") { 
        $id = $syncHash.settingVarDataGrid.SelectedItem.VarNum 
        $configHash.varListConfig.RemoveAt($syncHash.settingVarDataGrid.SelectedItem.VarNum - 1)
        $configHash.varListConfig | Where-Object { $_.VarNum -gt $id } | ForEach-Object { $_.VarNum = $_.VarNum - 1 }
        $syncHash.settingVarDataGrid.Items.Refresh()
       

    }

  

    elseif ($button.Name -match "settingVarBox") {

        $syncHash.settingVarDialog.IsOpen = $true
          

    }


}

[System.Windows.RoutedEventHandler]$eventonCommandGridDataGrid = {
    $button = $_.OriginalSource

    if ($button.Name -match "settingCommandGridClearItem") { 
        $id = $syncHash.settingCommandGridDataGrid.SelectedItem.ToolID
        $configHash.objectToolConfig[([Array]::IndexOf($configHash.objectToolConfig, $syncHash.settingObjectToolsPropGrid.SelectedItem))].toolCommandGridConfig.RemoveAt($syncHash.settingCommandGridDataGrid.SelectedItem.ToolID - 1)

        $configHash.objectToolConfig[([Array]::IndexOf($configHash.objectToolConfig, $syncHash.settingObjectToolsPropGrid.SelectedItem))].toolCommandGridConfig |
            Where-Object { $_.ToolID -gt $id } | ForEach-Object { $_.ToolID = $_.ToolID - 1 }
        
        $syncHash.settingCommandGridDataGrid.Items.Refresh()
    }

    
    elseif ($button.Name -match "settingCommandGridQueryBox") {
        $syncHash.settingCommandGridPopupText.Tag = "Query"
        $syncHash.settingCommandGridDialog.IsOpen = $true
    
    }

     elseif ($button.Name -match "settingCommandGridActionBox") {
        $syncHash.settingCommandGridPopupText.Tag = "Action"
        $syncHash.settingCommandGridDialog.IsOpen = $true
    
    }

    elseif ($button.Name -match 'settingCommandGridExecute') {
     $id = $syncHash.settingCommandGridDataGrid.SelectedItem.ToolID
        foreach ($cmd in @('queryCmd','actionCmd')) {
            
            $scriptBlockErrors = @()
            [void][System.Management.Automation.Language.Parser]::ParseInput(($configHash.objectToolConfig[([Array]::IndexOf($configHash.objectToolConfig, $syncHash.settingObjectToolsPropGrid.SelectedItem))].toolCommandGridConfig[$id - 1].$cmd), [ref]$null, [ref]$scriptBlockErrors)

            if ($scriptBlockErrors) {
                $configHash.objectToolConfig[([Array]::IndexOf($configHash.objectToolConfig, $syncHash.settingObjectToolsPropGrid.SelectedItem))].toolCommandGridConfig[$id - 1].($cmd + 'Valid') = 'False'
            }

            else {
                    $configHash.objectToolConfig[([Array]::IndexOf($configHash.objectToolConfig, $syncHash.settingObjectToolsPropGrid.SelectedItem))].toolCommandGridConfig[$id - 1].($cmd + 'Valid') = 'True'
            }
        }   
        
        $syncHash.settingCommandGridDataGrid.Items.Refresh()   
    }
}

[System.Windows.RoutedEventHandler]$eventonOUDataGrid = {
    $button = $_.OriginalSource

    if ($button.Name -match "settingOUClearItem") { 
        $id = $syncHash.settingOUDataGrid.SelectedItem.OUNum 
        $configHash.searchBaseConfig.RemoveAt($syncHash.settingOUDataGrid.SelectedItem.OUNum - 1)
        $configHash.searchBaseConfig | Where-Object { $_.OUNum -gt $id } | ForEach-Object { $_.OUNum = $_.OUNum - 1 }
        $syncHash.settingOUDataGrid.Items.Refresh()
       
       

    }

  

    elseif ($button.Name -match "settingOUSelect") {

         $syncHash.settingOUDataGrid.SelectedItem.OU = (Choose-ADOrganizationalUnit -HideNewOUFeature).DistinguishedName
          $syncHash.settingOUDataGrid.Items.Refresh()

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


[System.Windows.RoutedEventHandler]$eventonHistoryDataGrid = {


    $button = $_.OriginalSource
   

    if ($button.Name -match 'resultsErrorItem') { 
        $syncHash.historySideDataGrid.SelectedItem.Error | Set-Clipboard 
        $syncHash.SnackMsg.MessageQueue.Enqueue("Error copied") 
        }

    if  ($button.Name -match 'resultsQueryItem')  {
        $searchVal = $syncHash.historySideDataGrid.SelectedItem.SubjectName       
        $syncHash.SearchBox.Tag = $searchVal
        $syncHash.historySidePane.IsOpen = $false
        $syncHash.SearchBox.Focus()
        $wshell = New-Object -ComObject wscript.shell;
        $wshell.SendKeys('{ESCAPE}')
       
    }       
}


[System.Windows.RoutedEventHandler]$EventonContextGrid = {


    $button = $_.OriginalSource
   

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
[System.Windows.RoutedEventHandler]$EventonObjectToolCommandGridGrid = {
 $button = $_.OriginalSource

if ($button.Name -match 'toolsCommandGridExecute') {
    Write-Host "TEST!"
 }
 }    


[System.Windows.RoutedEventHandler]$EventonObjectToolGrid = {

    # GET THE NAME OF EVENT SOURCE
    $button = $_.OriginalSource
    # THIS RETURN THE ROW DATA AVAILABLE
    # resultObj scope is the whole script
   

    if ($button.Name -match 'settingObjectToolsEdit') {

        if ($syncHash.settingObjectToolsPropGrid.SelectedItem.toolType -eq 'CommandGrid') {
            $syncHash.settingCommandGridDataGrid.ItemsSource = $syncHash.settingObjectToolsPropGrid.SelectedItem.toolCommandGridConfig
        }
       
        $syncHash.settingObjectToolDefFlyout.Height = $syncHash.settingChildHeight.ActualHeight      
        $syncHash.settingObjectToolDefFlyout.IsOpen = $true
        $syncHash.settingObjectToolIcon.ItemsSource = $configHash.buttonGlyphs
        $syncHash.settingCommandGridToolIcon.ItemsSource = $configHash.buttonGlyphs
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
            toolTypeList          = @("Execute","Select","Grid","CommandGrid")
            toolCommandGridConfig = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
            toolType              = 'null'
            objectList            = @("Comp","User","Both","Standalone")
            objectType            = $null
            toolAction            = 'Do-Something -UserName $user'
            toolActionValid       = $false
            toolActionConfirm     = $true
            toolActionToolTip     = $null
            toolActionIcon        = $null
            toolActionSelectAD    = $false
            toolFetchCmd          = 'Get-Something'
            toolActionMultiSelect = $false
            toolDescription       = "Generic tool description"
            toolTargetFetchCmd    = 'Get-Something -Identity $target'
              
        })
    
    $temp = Get-InitialValues -GroupName 'toolCommandGridConfig'
    $temp | ForEach-Object { $configHash.objectToolConfig[$i -1].toolCommandGridConfig.Add($_) }

    $configHash.objectToolCount = ($configHash.objectToolConfig | Measure-Object).Count
}

[System.Windows.RoutedEventHandler]$addResultsItemClick = {
 

      $syncHash.SearchBox.Tag = ($syncHash.resultsSidePaneGrid.SelectedItem | Select-Object -ExpandProperty SamAccountName) -replace '\$' 
      $syncHash.SearchBox.Focus()
      
      $wshell = New-Object -ComObject wscript.shell;
      $wshell.SendKeys('{ESCAPE}')
      $syncHash.resultsSidePane.IsOpen = $false
     # $syncHash.resultsSidePaneGrid.ItemsSource = $null

      
   

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
$syncHash.settingVarDataGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $eventonVarDataGrid)
$syncHash.settingVarDataGrid.AddHandler([System.Windows.Controls.TextBox]::PreviewMouseLeftButtonUpEvent, $eventonVarDataGrid)
$syncHash.settingOUDataGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $eventonOUDataGrid)
$syncHash.historySideDataGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $eventonHistoryDataGrid)
$syncHash.toolsLogDataGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $eventonHistoryDataGrid)
$syncHash.itemToolCommandGridDataGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $EventonObjectToolCommandGridGrid)
$syncHash.settingCommandGridDataGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $eventonCommandGridDataGrid)
$syncHash.settingCommandGridDataGrid.AddHandler([System.Windows.Controls.TextBox]::PreviewMouseLeftButtonUpEvent, $eventonCommandGridDataGrid)






$syncHash.settingNameDialogClose.Add_Click( {
        $syncHash.settingNameDialog.IsOpen = $false
        $configHash.nameMapListView.Refresh()
})

$syncHash.settingInfoDialogClose.Add_Click({
    $syncHash.settingInfoDialog.IsOpen = $false

})

$syncHash.settingInfoDialogOpen.Add_Click({
    if ($syncHash.settingInfoDialog.IsOpen) { $syncHash.settingInfoDialog.IsOpen = $false }
    else { $syncHash.settingInfoDialog.IsOpen = $true }
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


 
 

$syncHash.Window.ShowDialog() | Out-Null
$configHash.IsClosed = $true

