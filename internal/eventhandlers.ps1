[System.Windows.RoutedEventHandler]$global:addItemClick = {
    if ($syncHash.settingUserPropGrid.Visibility -eq 'Visible') { $type = 'User' }

    else { $type = 'Comp' }

    $i = ($configHash.($type + 'PropList').Field |
            Sort-Object -Descending |
                Select-Object -First 1) + 1

    $configHash.($type + 'PropList').Add([PSCustomObject]@{
            Field                  = $i
            FieldName              = $null
            ItemType               = $type
            PropName               = $null
            propList               = $configHash.($type + 'PropPullListNames')
            translationCmd         = 'if ($result -eq $false)...'
            actionCmd1             = 'Do-Something -User $user'
            actionCmd1ToolTip      = 'Action name'
            actionCmd1Icon         = $null
            actionCmd1Refresh      = $false
            actionCmd1Multi        = $false
            ValidCmd               = $null
            ValidAction1           = $null
            ValidAction2           = $null
            actionCmd1CanOff       = $false
            actionCmd1OffStr       = ''
            actionCmd2             = 'Do-Something -User $user'
            actionCmd2ToolTip      = 'Action name'
            actionCmd2Icon         = $null
            actionCmd2Refresh      = $false
            actionCmd2Multi        = $false
            actionCmdsEnabled      = $true
            transCmdsEnabled       = $true
            actionCmd2Enabled      = $false
            actionCmd2CanOff       = $false
            actionCmd2OffStr       = ''
            PropType               = $null
            actionList             = @('ReadOnly', 'ReadOnly-Raw', 'Editable', 'Editable-Raw', 'Actionable', 'Actionable-Raw', 'Editable-Actionable', 'Editable-Actionable-Raw')
            ActionName             = 'null'
        })   
      
    $configHash.('box' + $type + 'Count') = ($configHash.($type + 'PropList') | Measure-Object).Count
}

[System.Windows.RoutedEventHandler]$global:addObjectToolItemClick = {
    $i = ($configHash.objectToolConfig.ToolID |
            Sort-Object -Descending |
                Select-Object -First 1) + 1

    $configHash.objectToolConfig.Add([PSCustomObject]@{
            ToolID                 = $i
            ToolName               = $null
            toolTypeList           = @('Execute', 'Select', 'Grid', 'CommandGrid')
            toolCommandGridConfig  = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
            toolType               = 'null'
            objectList             = @('Comp', 'User', 'Both', 'Standalone')
            objectType             = $null
            toolAction             = 'Do-Something -UserName $user'
            toolActionValid        = $null
            toolSelectValid        = $null
            toolExtraValid         = $null
            toolActionConfirm      = $true
            toolActionExportable   = $false
            toolStandAloneCat      = $null
            toolActionToolTip      = $null
            toolActionIcon         = $null
            toolActionSelectCustom = $false
            toolActionCustomOptions= @('AD Object (any)','AD User', 'AD Computer','AD Group','OU','String','Integer','Choice','File','Directory')
            toolActionCustomSelect = $null
            toolActionSelectChoice = $null
            toolFetchCmd           = 'Get-Something'
            toolActionMultiSelect  = $false
            toolDescription        = 'Generic tool description'
            toolTargetFetchCmd     = 'Get-Something -Identity $target'
        })
    
    $temp = Get-InitialValues -GroupName 'toolCommandGridConfig'
    $temp | ForEach-Object { $configHash.objectToolConfig[$i - 1].toolCommandGridConfig.Add($_) }
    $configHash.objectToolCount = ($configHash.objectToolConfig | Measure-Object).Count
}

[System.Windows.RoutedEventHandler]$global:addResultsItemClick = {
    $syncHash.SearchBox.Tag = ($syncHash.resultsSidePaneGrid.SelectedItem | Select-Object -ExpandProperty SamAccountName) -replace '\$' 
    $syncHash.SearchBox.Focus()
      
    $wshell = New-Object -ComObject wscript.shell
    $wshell.SendKeys('{ESCAPE}')
    $syncHash.resultsSidePane.IsOpen = $false
    # $syncHash.resultsSidePaneGrid.ItemsSource = $null
}

[System.Windows.RoutedEventHandler]$global:addContextItemClick = {
    $i = ($configHash.contextConfig.IDnum |
            Sort-Object -Descending |
                Select-Object -First 1) + 1

    $configHash.contextConfig.Add(([PSCustomObject]@{
                IDnum          = $i
                ActionName     = 'Name'
                RequireOnline  = $false
                RequireUser    = $false
                Types          = 'a'
                actionCmd      = 'Do-Something -User $user'
                actionCmdIcon  = $null
                actionCmdMulti = $false
                ValidAction    = $null
            }))
      
    $configHash.boxContextCount = ($configHash.contextConfig | Measure-Object).Count
}

[System.Windows.RoutedEventHandler]$global:EventonObjectToolCommandGridGrid = {
    $button = $_.OriginalSource

    if ($button.Name -match 'toolsCommandGridExecute') {
        $rsCmd = @{
            cmdGridItemIndex   = $syncHash.itemToolCommandGridDataGrid.SelectedItem.Index
            cmdGridParentIndex = $syncHash.itemToolCommandGridDataGrid.SelectedItem.ParentToolIndex
            result             = $syncHash.itemToolCommandGridDataGrid.SelectedItem.Result
            cmdGridItemName    = $syncHash.itemToolCommandGridDataGrid.SelectedItem.ItemName
        }

        $rsArgs = @{
            Name            = 'CommandGridRun'
            ArgumentList    = @($syncHash.snackMsg.MessageQueue, $configHash, $queryHash, $syncHash, $rsCmd, $syncHash.adHocConfirmWindow, $syncHash.Window, $syncHash.adHocConfirmText, $varHash)
            ModulesToImport = $configHash.modList
        }

        Start-RSJob @rsArgs -ScriptBlock {
            Param($queue, $configHash, $queryHash, $syncHash, $rsCmd, $confirmWindow, $window, $textBlock, $varHash)
            
            Set-CustomVariables -VarHash $varHash


            $actionCmd = $configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].toolCommandGridConfig[$rsCmd.cmdGridItemIndex].actionCmd
            $queryCmd = $configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].toolCommandGridConfig[$rsCmd.cmdGridItemIndex].queryCmd
            $toolName = $configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].toolName

            try {
                $result = $rsCmd.Result                     
                ([scriptblock]::Create($actionCmd)).Invoke() 
                $result = (Invoke-Expression $queryCmd).ToString()
                
                                
                $syncHash.Window.Dispatcher.Invoke([action] {
                        $syncHash.itemToolCommandGridDataGrid.SelectedItem.Result = $result
                        $syncHash.itemToolCommandGridDataGrid.Items.Refresh()               
                    })
               
                if ($configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].objectType -eq 'Standalone') {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -SubtoolName $rsCmd.cmdGridItemName -Status Success
                    Write-LogMessage -syncHashWindow $syncHash.window -Path $configHash.actionlogPath -Message Succeed -ActionName ($($($toolName) + '-' + $($rsCmd.cmdGridItemName))) -SubjectType 'Standalone' -ArrayList $configHash.actionLog
                
                }
                
                else {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -SubtoolName $rsCmd.cmdGridItemName -Status Success -SubjectName $activeObject
                    Write-LogMessage -syncHashWindow $syncHash.window -Path $configHash.actionlogPath -Message Succeed -SubjectName $activeObject -SubjectType $activeObjectType -ActionName ($($($toolName) + '-' + $($rsCmd.cmdGridItemName))) -ArrayList $configHash.actionLog 
                }
                
                if (($syncHash.itemToolCommandGridDataGrid.Items | Where-Object {$_.Result -notmatch 'True'} | Measure-Object).Count -eq 0) { $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.toolsCommandGridExecuteAll.Tag = 'False' }) }                                          
            }

            catch {
                if ($configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].objectType -eq 'Standalone') {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -SubtoolName $rsCmd.cmdGridItemName -Status Fail
                    Write-LogMessage -syncHashWindow $syncHash.window -Path $configHash.actionlogPath -Message Fail -ActionName ($($($toolName) + '-' + $($rsCmd.cmdGridItemName))) -SubjectName 'Standalone' -ArrayList $configHash.actionLog -Error $_
              
                }
                else {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -SubtoolName $rsCmd.cmdGridItemName -Status Fail -SubjectName $activeObject
                    Write-LogMessage -syncHashWindow $syncHash.window -Path $configHash.actionlogPath -Message Fail -SubjectName $activeObject -SubjectType $activeObjectType -ActionName ($($($toolName) + '-' + $($rsCmd.cmdGridItemName))) -ArrayList $configHash.actionLog -Error $_
                }
            }
        }
    }
}    

[System.Windows.RoutedEventHandler]$global:EventonObjectToolGrid = {
    # GET THE NAME OF EVENT SOURCE
    $button = $_.OriginalSource
    # THIS RETURN THE ROW DATA AVAILABLE
    # resultObj scope is the whole script
   
        $itemSource = $configHash.objectToolConfig[$syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID - 1]

        Update-ScriptBlockValidityStatus -syncHash $syncHash -itemSet $itemSource -statusName 'toolActionValid' -ResultBoxName 'settingObjectToolResultBox'
        Update-ScriptBlockValidityStatus -SyncHash $syncHash -ItemSet $itemSource -StatusName 'toolSelectValid' -ResultBoxName 'settingObjectToolSelectionResultBox'
        Update-ScriptBlockValidityStatus -SyncHash $syncHash -ItemSet $itemSource -StatusName 'toolExtraValid' -ResultBoxName 'settingObjectToolExtraResultBox'

    if ($button.Name -match 'settingObjectToolsEdit') {
        if ($syncHash.settingObjectToolsPropGrid.SelectedItem.toolType -eq 'CommandGrid') {
            $syncHash.settingCommandGridDataGrid.Visibility = 'Visible'
            $syncHash.settingCommandGridDataGrid.ItemsSource = $syncHash.settingObjectToolsPropGrid.SelectedItem.toolCommandGridConfig
        }
       
        $syncHash.settingObjectToolDefFlyout.Height = $syncHash.settingChildHeight.ActualHeight      
        $syncHash.settingObjectToolDefFlyout.IsOpen = $true
        $syncHash.settingObjectToolIcon.ItemsSource = $configHash.buttonGlyphs
        $syncHash.settingObjectStandaloneCat.ItemsSource = $configHash.SACats
        $syncHash.settingCommandGridToolIcon.ItemsSource = $configHash.buttonGlyphs
        $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.settingObjectToolDefFlyout.Background.Color).ToString()
        $syncHash.settingChildWindow.Title = 'Define Tool'
        $syncHash.settingChildWindow.ShowCloseButton = $false

        switch ($syncHash.settingObjectToolsPropGrid.SelectedItem.toolType) {
            'CommandGrid' {  Set-CurrentPane -SyncHash $syncHash -Panel 'settingObjectToolCommandGrid'; break }
            'Select' {  Set-CurrentPane -SyncHash $syncHash -Panel 'settingObjectToolSelect' }
            'Execute' {  Set-CurrentPane -SyncHash $syncHash -Panel 'settingObjectToolExecute' }
            'Grid' {  Set-CurrentPane -SyncHash $syncHash -Panel 'settingObjectToolGrid' }
            
        }
            

       # Set-CurrentPane -SyncHash $syncHash -Panel 'settingObjectToolDefFlyout'
    }

    elseif ($button.Name -match 'settingObjectToolsClearItem') {
        $num = $syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID 
        $configHash.objectToolConfig.RemoveAt($syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID - 1)
        $configHash.objectToolConfig |
            Where-Object { $_.ToolID -gt $num } |
                ForEach-Object { $_.ToolID = $_.ToolID - 1 }       
        $syncHash.settingObjectToolsPropGrid.Items.Refresh()
    }

    elseif ($button.Name -match 'settingUpItem') {
        Move-DataGridItem -SyncHash $syncHash -ConfigHash $configHash -CollectionName objectToolConfig -GridName settingObjectToolsPropGrid -IDName 'ToolID' -Direction Up   
    }

    elseif ($button.Name -match 'settingDownItem') {
        Move-DataGridItem -SyncHash $syncHash -ConfigHash $configHash -CollectionName objectToolConfig -GridName settingObjectToolsPropGrid -IDName 'ToolID' -Direction Down   
    }
}

[System.Windows.RoutedEventHandler]$global:EventonContextGrid = {
    $button = $_.OriginalSource

    $itemSource = $configHash.contextConfig[$syncHash.settingContextPropGrid.SelectedItem.IDNum - 1]
    Update-ScriptBlockValidityStatus -syncHash $syncHash -itemSet $itemSource -statusName 'ValidAction' -ResultBoxName 'settingContextResultBox'


    if ($button.Name -match 'settingContextEdit') {
        $syncHash.settingContextIcon.ItemsSource = $configHash.buttonGlyphs
        $syncHash.settingContextDefFlyout.Height = $syncHash.settingChildHeight.ActualHeight      
        $syncHash.settingContextDefFlyout.IsOpen = $true
        $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.settingContextDefFlyout.Background.Color).ToString()
        $syncHash.settingChildWindow.Title = 'Define Context Button'
        $syncHash.settingChildWindow.ShowCloseButton = $false
        Set-CurrentPane -SyncHash $syncHash -Panel 'settingContextDefFlyout'
    }

 
    elseif ($button.Name -match 'settingContextClearItem') {
        $num = $syncHash.settingContextPropGrid.SelectedItem.IDNum 
        $configHash.contextConfig.RemoveAt($syncHash.settingContextPropGrid.SelectedItem.IDNum - 1)
        $configHash.contextConfig |
            Where-Object -FilterScript { $_.IDNum -gt $num } |
                ForEach-Object -Process { $_.IDNum = $_.IDNum - 1 }       
        $syncHash.settingContextPropGrid.Items.Refresh()
    }

    elseif ($button.Name -match 'settingUpItem') {
        Move-DataGridItem -SyncHash $syncHash -ConfigHash $configHash -CollectionName contextConfig -GridName settingContextPropGrid -IDName 'IDNum' -Direction Up   
    }

    elseif ($button.Name -match 'settingDownItem') {
        Move-DataGridItem -SyncHash $syncHash -ConfigHash $configHash -CollectionName contextConfig -GridName settingContextPropGrid -IDName 'IDNum' -Direction Down   
    }
}

[System.Windows.RoutedEventHandler]$global:EventonPropertyDataGrid = {
    # GET THE NAME OF EVENT SOURCE
    $button = $_.OriginalSource
    # THIS RETURN THE ROW DATA AVAILABLE
    # resultObj scope is the whole script
    if ($syncHash.settingUserPropGrid.Visibility -eq 'Visible') { $type = 'User' }

    else { $type = 'Comp' }

    if ($button.Name -match 'settingEdit') {

         $itemSource = $configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1]

        Update-ScriptBlockValidityStatus -syncHash $syncHash -itemSet $itemSource -statusName 'ValidCmd' -ResultBoxName 'settingResultBox'
        Update-ScriptBlockValidityStatus -SyncHash $syncHash -ItemSet $itemSource -StatusName 'ValidAction1' -ResultBoxName 'settingBox1ResultBox'
        Update-ScriptBlockValidityStatus -SyncHash $syncHash -ItemSet $itemSource -StatusName 'ValidAction1' -ResultBoxName 'settingBox2ResultBox'

       
        $switchVal = $itemSource.ActionName

        switch -wildcard ($switchVal) {
            
            { $switchVal -notmatch 'raw' } {
                $syncHash.settingFlyoutTranslate.Visibility = 'Visible'
                (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).transCmdsEnabled = $true
            }

            { $switchVal -notmatch 'actionable' } {
                $syncHash.settingFlyoutTranslate.IsExpanded = $true
                $syncHash.settingFlyoutAction1.Visibility = 'Collapsed'
                $syncHash.settingFlyoutAction2.Visibility = 'Collapsed'
                (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).actionCmdsEnabled = $false
                    
                if ($null -like (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).ActionCmd1) { (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).ActionCmd1 = "Do-Something -$type $('$' + $type)..." } 
                                   
                if ($null -like (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).ActionCmd2) { (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).ActionCmd2 = "Do-Something -$type $('$' + $type)..." }    
            }

            { $switchVal -match 'raw' } { 
                $syncHash.settingFlyoutTranslate.Visibility = 'Collapsed'                     
                (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).transCmdsEnabled = $false                   
                if ($null -like (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).TranslationCmd) { (($configHash.($type + 'PropList')[$syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1])).TranslationCmd = 'if ($result -eq $false)...' }
            }

            { $switchVal -match 'actionable' } {
                $syncHash.settingFlyoutAction1.Visibility = 'Visible'
                $syncHash.settingFlyoutAction2.Visibility = 'Visible'
                if ($syncHash.settingFlyoutTranslate.Visibility -eq 'Collapsed') { $syncHash.settingFlyoutAction1.IsExpanded = $true }
            }

            
        }

        if ($type -eq 'User') { Set-CurrentPane -SyncHash $syncHash -Panel 'settingPropUserDefine' }
        else { Set-CurrentPane -SyncHash $syncHash -Panel 'settingPropCompDefine' }
            
        
        $syncHash.settingBox1Icon.ItemsSource = $configHash.buttonGlyphs
        $syncHash.settingBox2Icon.ItemsSource = $configHash.buttonGlyphs
        $syncHash.settingUserPropDefFlyout.Height = $syncHash.settingChildHeight.ActualHeight
           
        $syncHash.settingUserPropDefFlyout.IsOpen = $true
        $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.settingUserPropDefFlyout.Background.Color).ToString()
        $syncHash.settingChildWindow.Title = "Define Button ($type)"
        $syncHash.settingChildWindow.ShowCloseButton = $false
    }

    elseif ($button.Name -match ('setting' + "$type" + 'ComboBox') -and ($syncHash.('setting' + $type + 'PropGrid').SelectedItem)) {
        Remove-Variable -Name set -ErrorAction SilentlyContinue
        $set = (($configHash.($type + 'PropPullList') | Where-Object -FilterScript { $_.Name -eq ($syncHash.('setting' + $type + 'PropGrid').SelectedItem).PropName }).TypeNameOfValue -replace '.*(?=\.).', '').toString()
        $syncHash.settingTypeBox.Text = $set
        ($configHash.($type + 'PropList') | Where-Object -FilterScript { $_.Field -eq (($syncHash.('setting' + $type + 'PropGrid').SelectedItem).Field) }).PropType = $set
    }

    elseif ($button.Name -match 'settingClearItem') {
        $num = $syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field 
        $configHash.($type + 'PropList').RemoveAt($syncHash.('setting' + $type + 'PropGrid').SelectedItem.Field - 1)
        $configHash.($type + 'PropList') |
            Where-Object -FilterScript { $_.Field -gt $num } |
                ForEach-Object -Process { $_.Field = $_.Field - 1 }
       

        $syncHash.('setting' + $type + 'PropGrid').Items.Refresh()
    }

    elseif ($button.Name -match 'settingUpItem') {
        Move-DataGridItem -SyncHash $syncHash -ConfigHash $configHash -CollectionName ($type + 'PropList') -GridName ('setting' + $type + 'PropGrid') -IDName 'Field' -Direction Up   
    }

    elseif ($button.Name -match 'settingDownItem') {
        Move-DataGridItem -SyncHash $syncHash -ConfigHash $configHash -CollectionName ($type + 'PropList') -GridName ('setting' + $type + 'PropGrid') -IDName 'Field' -Direction Down   
    }

    if (($syncHash.('setting' + $type + 'PropGrid').SelectedItem).PropName -eq 'Non-Ad Property') { $syncHash.settingFlyoutResultHeader.Content = 'Retrieval Command' }

   
   

    else { $syncHash.settingFlyoutResultHeader.Content = 'Result Presentation' }
}

[System.Windows.RoutedEventHandler]$global:eventonHistoryDataGrid = {
    $button = $_.OriginalSource
   

    if ($button.Name -match 'resultsErrorItem') { 
        $syncHash.historySideDataGrid.SelectedItem.Error | Set-Clipboard 
        $syncHash.SnackMsg.MessageQueue.Enqueue('Error copied') 
    }

    if ($button.Name -match 'resultsQueryItem') {
        
        if ($syncHash.tabMenu.SelectedIndex -ne 0) { $syncHash.tabMenu.SelectedIndex = 0 }

        $searchVal = $syncHash.historySideDataGrid.SelectedItem.SubjectName       

        if ($searchVal -in $syncHash.tabControl.Items.Name) {
            if ($searchVal -ne $syncHash.tabControl.SelectedItem.Name) {
               $caseCorrectedSearchVal = $syncHash.tabControl.Items.Name | Where-Object {$_ -eq $searchVal}
               $syncHash.tabControl.SelectedIndex = [Array]::IndexOf($syncHash.tabControl.Items.Name, $caseCorrectedSearchVal)
            }
        }
        
        else {  
            $syncHash.SearchBox.Tag = $searchVal
            $syncHash.SearchBox.Focus()
            $wshell = New-Object -ComObject wscript.shell
            $wshell.SendKeys('{ESCAPE}')
        }
    
    $syncHash.historySidePane.IsOpen = $false
    
    }       
}

[System.Windows.RoutedEventHandler]$global:eventonToolsLogDataGrid = {
    $button = $_.OriginalSource
   

    if ($button.Name -match 'resultsErrorItem') { 
        $syncHash.toolsLogDataGrid.SelectedItem.Error | Set-Clipboard 
        $syncHash.SnackMsg.MessageQueue.Enqueue('Error copied') 
    }

    if ($button.Name -match 'resultsQueryItem') {
        $syncHash.tabMenu.SelectedIndex = 0  
        $searchVal = $syncHash.toolsLogDataGrid.SelectedItem.SubjectName           
        $syncHash.SearchBox.Text = $searchVal

    }       
}

[System.Windows.RoutedEventHandler]$global:eventonCommandGridDataGrid = {
    $button = $_.OriginalSource

    if ($button.Name -match 'settingCommandGridClearItem') { 
        $id = $syncHash.settingCommandGridDataGrid.SelectedItem.ToolID
        $configHash.objectToolConfig[([Array]::IndexOf($configHash.objectToolConfig, $syncHash.settingObjectToolsPropGrid.SelectedItem))].toolCommandGridConfig.RemoveAt($syncHash.settingCommandGridDataGrid.SelectedItem.ToolID - 1)

        $configHash.objectToolConfig[([Array]::IndexOf($configHash.objectToolConfig, $syncHash.settingObjectToolsPropGrid.SelectedItem))].toolCommandGridConfig |
            Where-Object -FilterScript { $_.ToolID -gt $id } |
                ForEach-Object -Process { $_.ToolID = $_.ToolID - 1 }
        
        $syncHash.settingCommandGridDataGrid.Items.Refresh()
    }

    
    elseif ($button.Name -match 'settingCommandGridQueryBox') {
        $syncHash.settingCommandGridPopupText.Tag = 'Query'
        $syncHash.settingCommandGridDialog.IsOpen = $true
    }

    elseif ($button.Name -match 'settingCommandGridActionBox') {
        $syncHash.settingCommandGridPopupText.Tag = 'Action'
        $syncHash.settingCommandGridDialog.IsOpen = $true
    }

    elseif ($button.Name -match 'settingCommandGridExecute') {
        $id = $syncHash.settingCommandGridDataGrid.SelectedItem.ToolID
        foreach ($cmd in @('queryCmd', 'actionCmd')) {
            $scriptBlockErrors = @()
            [void][System.Management.Automation.Language.Parser]::ParseInput(($configHash.objectToolConfig[([Array]::IndexOf($configHash.objectToolConfig, $syncHash.settingObjectToolsPropGrid.SelectedItem))].toolCommandGridConfig[$id - 1].$cmd), [ref]$null, [ref]$scriptBlockErrors)

            if ($scriptBlockErrors) { $configHash.objectToolConfig[([Array]::IndexOf($configHash.objectToolConfig, $syncHash.settingObjectToolsPropGrid.SelectedItem))].toolCommandGridConfig[$id - 1].($cmd + 'Valid') = 'False' }

            else { $configHash.objectToolConfig[([Array]::IndexOf($configHash.objectToolConfig, $syncHash.settingObjectToolsPropGrid.SelectedItem))].toolCommandGridConfig[$id - 1].($cmd + 'Valid') = 'True' }
        }   
        
        $syncHash.settingCommandGridDataGrid.Items.Refresh()   
    }
}

[System.Windows.RoutedEventHandler]$global:eventonOUDataGrid = {
    $button = $_.OriginalSource

    if ($button.Name -match 'settingOUClearItem') { 
        $id = $syncHash.settingOUDataGrid.SelectedItem.OUNum 
        $configHash.searchBaseConfig.RemoveAt($syncHash.settingOUDataGrid.SelectedItem.OUNum - 1)
        $configHash.searchBaseConfig |
            Where-Object -FilterScript { $_.OUNum -gt $id } |
                ForEach-Object -Process { $_.OUNum = $_.OUNum - 1 }
        $syncHash.settingOUDataGrid.Items.Refresh()
    }

  

    elseif ($button.Name -match 'settingOUSelect') {
        $syncHash.settingOUDataGrid.SelectedItem.OU = (Choose-ADOrganizationalUnit -HideNewOUFeature).DistinguishedName
        $syncHash.settingOUDataGrid.Items.Refresh()
    }
}

[System.Windows.RoutedEventHandler]$global:eventonNetDataGrid = {
    $button = $_.OriginalSource
   

    if ($button.Name -match 'settingNetClearItem') {
        $id = $syncHash.settingNetDataGrid.SelectedItem.ID  
        $configHash.netMapList.Remove(($configHash.netMapList | Where-Object -FilterScript { $_.ID -eq ($syncHash.settingNetDataGrid.SelectedItem.ID) }))
        
        $configHash.netMapList |
            Where-Object -FilterScript { $_.Id -gt $id } |
                ForEach-Object -Process { $_.Id = $_.Id - 1 }       
        
        $syncHash.settingNetDataGrid.Items.Refresh()
    }

    elseif ($button.Name -match 'settingnetIP') {
        try {
            [IPAddress]$syncHash.settingNetDataGrid.SelectedItem.Network
            $syncHash.settingNetDataGrid.SelectedItem.validNetwork = $true
            $button.Foreground = 'White'
            $button.Tooltip = $null
            $button.TextDecorations = $null
        }
        catch {
            $syncHash.settingNetDataGrid.SelectedItem.validNetwork = $false
            $button.Foreground = 'Red'
            $button.Tooltip = 'Invalid IP'
            $button.TextDecorations = 'Underline'
        }
    }

    elseif ($button.Name -match 'settingnetMask') {
        try {
            if ([int]$syncHash.settingNetDataGrid.SelectedItem.Mask -gt 0 -and [int]$syncHash.settingNetDataGrid.SelectedItem.Mask -le 32) {
                $syncHash.settingNetDataGrid.SelectedItem.validMask = $true
                $button.Foreground = 'White'
                $button.Tooltip = $null
                $button.TextDecorations = $null
            }
            else { 1 / 0 }
        }
        catch {
            $syncHash.settingNetDataGrid.SelectedItem.validMask = $false
            $button.Foreground = 'Red'
            $button.Tooltip = 'Invalid mask'
            $button.TextDecorations = 'Underline'
        }
    }
}

[System.Windows.RoutedEventHandler]$global:eventonNameDataGrid = {
    $button = $_.OriginalSource

    if ($button.Name -match 'settingNameClearItem') { 
        $id = $syncHash.settingNameDataGrid.SelectedItem.ID 
        $configHash.NameMapList.Remove(($configHash.NameMapList | Where-Object -FilterScript { $_.ID -eq ($syncHash.settingNameDataGrid.SelectedItem.ID) }))
      #  ($configHash.NameMapList | Where-Object -FilterScript { $_.Id -eq ($syncHash.settingNameDataGrid.SelectedItem.ID) })

        $configHash.NameMapList |
            Where-Object -FilterScript { $_.Id -gt $id } |
                ForEach-Object -Process { $_.Id = $_.Id - 1 }
       
        $configHash.nameMapListView.Refresh()
        $syncHash.settingNameDataGrid.Items.Refresh()
    }

    elseif ($button.Name -match 'settingNameUpItem') {
        $id = $syncHash.settingNameDataGrid.SelectedItem.Id


        if ($id + 1 -eq ($configHash.NameMapList.Id |
                    Sort-Object -Descending |
                        Select-Object -First 1)) { ($configHash.NameMapList | Where-Object -FilterScript { $_ -eq $syncHash.settingNameDataGrid.SelectedItem }).TopPos = $true }

        ($configHash.NameMapList | Where-Object -FilterScript { $_.Id -eq $id + 1 }).TopPos = $false
        ($configHash.NameMapList | Where-Object -FilterScript { $_.Id -eq $id + 1 }).Id = $id


       
        ($configHash.NameMapList | Where-Object -FilterScript { $_ -eq $syncHash.settingNameDataGrid.SelectedItem }).ID = $id + 1
        $configHash.nameMapListView.Refresh()
        $syncHash.settingNameDataGrid.Items.Refresh()
    }

    elseif ($button.Name -match 'settingNameDownItem') {
        $id = $syncHash.settingNameDataGrid.SelectedItem.Id
    
        if ($id -eq ($configHash.NameMapList.Id |
                    Sort-Object -Descending |
                        Select-Object -First 1)) { ($configHash.NameMapList | Where-Object -FilterScript { $_.Id -eq $id - 1 }).TopPos = $true }

        ($configHash.NameMapList | Where-Object -FilterScript { $_.Id -eq $id - 1 }).Id = $id

        ($configHash.NameMapList | Where-Object -FilterScript { $_ -eq $syncHash.settingNameDataGrid.SelectedItem }).TopPos = $false
        ($configHash.NameMapList | Where-Object -FilterScript { $_ -eq $syncHash.settingNameDataGrid.SelectedItem }).ID = $id - 1
        $configHash.nameMapListView.Refresh()
        $syncHash.settingNameDataGrid.Items.Refresh()
    }

    elseif ($button.Name -match 'settingConditionBox') { $syncHash.settingNameDialog.IsOpen = $true }
}

[System.Windows.RoutedEventHandler]$global:eventonVarDataGrid = {
    $button = $_.OriginalSource

    if ($button.Name -match 'settingVarClearItem') { 
        $id = $syncHash.settingVarDataGrid.SelectedItem.VarNum 
        $configHash.varListConfig.RemoveAt($syncHash.settingVarDataGrid.SelectedItem.VarNum - 1)
        $configHash.varListConfig |
            Where-Object -FilterScript { $_.VarNum -gt $id } |
                ForEach-Object -Process { $_.VarNum = $_.VarNum - 1 }
        $syncHash.settingVarDataGrid.Items.Refresh()
    }

  

    elseif ($button.Name -match 'settingVarBox') { $syncHash.settingVarDialog.IsOpen = $true }
}

[System.Windows.RoutedEventHandler]$global:eventonModDataGrid = {
    $button = $_.OriginalSource

    if ($button.Name -match 'settingModClearItem') { 
        $id = $syncHash.settingModDataGrid.SelectedItem.modNum 
        $configHash.modConfig.RemoveAt($syncHash.settingModDataGrid.SelectedItem.modNum - 1)
        $configHash.modConfig |
            Where-Object -FilterScript { $_.modNum -gt $id } |
                ForEach-Object -Process { $_.modNum = $_.modNum - 1 }
        $syncHash.settingModDataGrid.Items.Refresh()
    }

    elseif ($button.Name -match 'settingModPathSelect') {
        $modPath = Get-PSModulePath

        $syncHash.settingModDataGrid.SelectedItem.ModPath = $modPath.fileName
        
        if ([string]::IsNullOrEmpty($syncHash.settingModDataGrid.SelectedItem.ModName)) { $syncHash.settingModDataGrid.SelectedItem.ModName = [io.path]::GetFileNameWithoutExtension($modPath.SafeFileName) }

        $syncHash.settingModDataGrid.Items.Refresh()
    }
}

[System.Windows.RoutedEventHandler]$global:eventonQueryDefDataGrid = {
    $button = $_.OriginalSource

    if ($button.Name -match 'settingQueryDefClearItem') { 
        $id = $syncHash.settingQueryDefDataGrid.SelectedItem.ID 
        $configHash.queryDefConfig.RemoveAt($syncHash.settingQueryDefDataGrid.SelectedItem.ID - 1)
        $configHash.queryDefConfig |
            Where-Object -Filter { $_.ID -gt $id } |
                ForEach-Object -Process { $_.ID = $_.ID - 1 }
        $syncHash.settingQueryDefDataGrid.Items.Refresh()
    }
}

$syncHash.sidePaneExit.Add_Click( { $syncHash.resultsSidePane.IsOpen = $false })
$syncHash.settingUserPropGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $EventonPropertyDataGrid)
$syncHash.settingContextPropGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $EventonContextGrid)
$syncHash.settingUserPropGrid.AddHandler([System.Windows.Controls.Combobox]::SelectionChangedEvent, $EventonPropertyDataGrid)
$syncHash.settingCompPropGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $EventonPropertyDataGrid)
$syncHash.settingObjectToolsPropGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $EventonObjectToolGrid)
$syncHash.settingCompPropGrid.AddHandler([System.Windows.Controls.Combobox]::SelectionChangedEvent, $EventonPropertyDataGrid)
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
$syncHash.toolsLogDataGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $eventonToolsLogDataGrid)
$syncHash.itemToolCommandGridDataGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $EventonObjectToolCommandGridGrid)
$syncHash.settingCommandGridDataGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $eventonCommandGridDataGrid)
$syncHash.settingCommandGridDataGrid.AddHandler([System.Windows.Controls.TextBox]::PreviewMouseLeftButtonUpEvent, $eventonCommandGridDataGrid)
$syncHash.settingModDataGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $eventonModDataGrid)
$syncHash.settingQueryDefDataGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $eventonQueryDefDataGrid)
