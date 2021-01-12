# saves the config import json to synchash or export synchash to json

#region initial setup

# generated hash tables used throughout tool
function New-HashTables {

    # Stores values logging missing or errored items during init
    $global:sysCheckHash = [hashtable]::Synchronized(@{ })

    # Stores config values imported JSON, during config, or both
    $global:configHash = [hashtable]::Synchronized(@{ })

    # Stores WPF controls
    $global:syncHash = [hashtable]::Synchronized(@{ })

     # Stores config'd vars
    $global:varHash = [hashtable]::Synchronized(@{ })

    # Stores data related to queried objects
    $global:queryHash = [hashtable]::Synchronized(@{ })

}

function Get-Glyphs {
    param (
        $ConfigHash,
        $GlyphList)

    $glyphs = Get-Content $glyphList
    $configHash.buttonGlyphs = [System.Collections.ArrayList]@()
    $glyphs | ForEach-Object { $configHash.buttonGlyphs.Add($_) | Out-Null }
}

function Set-WPFControls {
    param (
        [Parameter(Mandatory)]$XAMLPath,
        [Parameter(Mandatory)][Hashtable]$TargetHash
    ) 

    $inputXML = Get-Content -Path $xamlPath
    
    $inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace ' x:Class="v3.Window1"' -replace ' x:Class="V3.Build.MainWindow"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = $inputXML

    $xmlReader = (New-Object System.Xml.XmlNodeReader $xaml)
    
    try { $TargetHash.Window = [Windows.Markup.XamlReader]::Load($xmlReader) }
    catch { Write-Warning "Unable to parse XML, with error: $($Error[0])" }

    ## Load each named control into PS hashtable
    foreach ($controlName in ($xaml.SelectNodes("//*[@Name]").Name)) {
        $TargetHash.$controlName = $TargetHash.Window.FindName($controlName) 
    }

    $syncHash.windowContent.Visibility = "Hidden"
    $syncHash.Window.Height = 500
    $syncHash.Window.ResizeMode = "NoResize"
    $syncHash.Window.ShowTitleBar = $false
    $syncHash.Window.ShowCloseButton = $false
    $syncHash.Window.Width = 500
    $syncHash.splashLoad.Visibility = "Visible"     

}

function Show-WPFWindow {
    param($SyncHash) 
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

function Set-Config {
    Param ( 
        [CmdletBinding()]
        [Parameter(Mandatory = $true)]
        [String]
        $ConfigPath,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Export', 'Import')]
        [string[]]
        $Type,

        [Parameter(Mandatory = $true)]
        [Hashtable]
        $ConfigHash
    )

    switch ($type) {
        'Import' {

            if (Test-Path $configPath) {

                if ((Get-ChildItem -LiteralPath $configPath).Length -eq 0 -and (Get-ChildItem -LiteralPath $($savedConfig + '.bak')).Length -gt 0) {
                    Copy-Item -LiteralPath $configPath -Destination $($configPath + '.bak')
                }

                (Get-Content $configPath | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $configHash[$_.Name] = $_.Value }

            }
        }
        'Export' {
            @('User','Comp') | ForEach-Object {
                $configHash.($_ + 'PropList') | ForEach-Object {$_.PropList = $null}
                $configHash.($_ + 'PropListSelection') = $null
                $configHash.($_ + 'PropPullListNames') = $null
                $configHash.($_ + 'PropPullList') = $null
            }

            $configHash.buttonGlyphs = $null
            $configHash.adPropertyMap = $null
            $configHash.queryProps = $null
            $configHash.actionLog = $null

            $configHash | ConvertTo-Json -Depth 8 | Out-File $($configPath + '.bak') -Force

            if ((Get-ChildItem -LiteralPath $($configPath + '.bak')).Length -gt 0) {
                Copy-Item -LiteralPath $($configPath + '.bak') -Destination $configPath 
            }
        }
    }
}

function Import-Config {
    param ($SyncHash)

    $configSelection = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        InitialDirectory = [Environment]::GetFolderPath('MyComputer')
        Filter           = "config|config.json"
        Title            = "Select config.json"
    }
    
    $configSelection.ShowDialog() | Out-Null

    if (![string]::IsNullOrEmpty($configSelection.fileName)) {
        Copy-Item -Path $configSelection.fileName -Destination $PSScriptRoot -Force
        $syncHash.Window.Close()
        Start-Process -WindowStyle Minimized -FilePath "$PSHOME\powershell.exe" -ArgumentList " -ExecutionPolicy Bypass -NonInteractive -File $($PSCommandPath)"
        exit
       
    }
}

function Suspend-FailedItems {
    param (
        $SyncHash,
        [Parameter(Mandatory = $true)][ValidateSet('Config', 'SysCheck', 'RSCheck')][string[]]$CheckedItems
    )

    switch ($CheckedItems) {

        "RSCheck" {
            $syncHash.splashLoad.Visibility = "Collapsed"
            $syncHash.windowContent.Visibility = "Visible"
            $syncHash.Window.MinWidth = "1000"
            $syncHash.Window.MinHeight = "700"
            $syncHash.Window.ResizeMode = "CanResizeWithGrip"
            $syncHash.Window.ShowTitleBar = $true
            $syncHash.Window.ShowCloseButton = $true   
            $syncHash.windowContent.Visibility = "Visible"
            $syncHash.settingADClick.IsEnabled = $false
            $syncHash.settingPermClick.IsEnabled = $false
            $syncHash.settingFailPanel.Visibility = "Visible"  
            $syncHash.settingConfigSeperator.Visibility = 'Hidden'
            $syncHash.settingsConfigItems.Visibility = 'Hidden' 
            $syncHash.settingStatusChildBoard.Visibility = 'Visible'
            $syncHash.settingModADLabel.Visibility = "Collapsed"
            $syncHash.settingADLabel.Visibility = "Collapsed"
            $syncHash.settingPermLabel.Visibility = "Collapsed"
             
            break                       
        }

        "SysCheck" {     

            $syncHash.Window.Dispatcher.invoke([action] {                        
                    $syncHash.settingStatusChildBoard.Visibility = 'Visible'
                    $syncHash.settingFailPanel.Visibility = 'Visible'
                    $syncHash.settingConfigSeperator.Visibility = 'Hidden'
                    $syncHash.settingsConfigItems.Visibility = 'Hidden'
                })

            break
        }  
        
        "Config" {

            $syncHash.Window.Dispatcher.invoke([action] {            
                    $syncHash.settingStatusChildBoard.Visibility = 'Visible'
                    $syncHash.settingConfigPanel.Visibility = 'Visible'
                    $syncHash.settingConfigMissing.Visibility = 'Visible'
                })          
        } 
    }

    $syncHash.tabMenu.Items | ForEach-Object { $syncHash.Window.Dispatcher.invoke([action] { $_.IsEnabled = $false }) }
    
    $syncHash.Window.Dispatcher.invoke([action] { 
            $syncHash.tabMenu.Items[3].IsEnabled = $true
            $syncHash.tabMenu.SelectedIndex = 3 
        })   
} 



# gets initial values saved in json and loads into PCO
function Get-InitialValues {
    Param ( 
        [parameter(Mandatory = $true)]$GroupName
    )

    $basePath = Join-Path -Path $PSScriptRoot -ChildPath base

    if (Test-Path (Join-Path -Path $basePath -ChildPath ($($groupName) + '.json'))) {
        $initialConfig = (Get-Content (Join-Path -Path $basePath -ChildPath ($($groupName) + '.json'))) | ConvertFrom-Json
    }

    return $initialConfig

}

# process loaded data or creates initial item templates for various config datagrids
function Set-InitialValues {
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory, ValueFromPipeline)]$Type,
        [Parameter(Mandatory)][Hashtable]$ConfigHash,
        [switch]$PullDefaults     
    )

    Begin { 
        Set-ADGenericQueryNames -ConfigHash $ConfigHash
    }

    Process {
       
        # check if values already exist
        if ($configHash.$type) { $tempList = $configHash.$type }
        
        # pull from base templates if not
        elseif ($PullDefaults) { $tempList = Get-InitialValues -GroupName $type }

        else { $tempList = $null }
        

        # create observable collection and add values
        $configHash.$type = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
        $tempList | ForEach-Object { $configHash.$type.Add($_) }

    }

    End {
        # get the current max values for each of the main property boxes
        $configHash.UserboxCount = ($configHash.userPropList | Measure-Object).Count
        $configHash.CompboxCount = ($configHash.compPropList | Measure-Object).Count
        $configHash.boxMax = ($configHash.UserboxCount, $configHash.compPropList | Measure-Object -Maximum).Maximum

    }
}

# matches config'd user/comp logins with default headers, creates new headers
# will append with number if defined values are duplicates
function Set-LoggingStructure {
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory, ValueFromPipeline)]$Type,
        [Parameter(Mandatory)][Hashtable]$ConfigHash,
        [Array]$DefaultList      
    )

    Process {

        $addList = @()

        if ($configHash.$type) {

            $duplicate = 0
            $configHash.$type | Add-Member -MemberType NoteProperty -Name 'Header' -Value $null -ErrorAction SilentlyContinue

            $configHash.$type | ForEach-Object {
                $_.FieldSelList = $DefaultList
        
                if ($_.FieldSel -notin $addList) {
                    if ($_.FieldSel -eq 'Custom') { $_.Header = $_.CustomFieldName; $addList += $_.CustomFieldName }
                    else { $_.Header = $_.FieldSel; $addList += $_.FieldSel }
                }

                else {
                    $_.Header = ($_.FieldSel + $duplicate)
                    $duplicate++
                }
            }
        }
    }
}

function Set-QueryPropertyList {
   Param ($SyncHash, $ConfigHash)

    if (($configHash.queryDefConfig.Name | Measure-Object).Count -eq 1) { 
        $syncHash.Window.Dispatcher.Invoke([Action]{ $syncHash.searchPropSelection.ItemsSource = @($configHash.queryDefConfig.Name) })
    }
    elseif (($configHash.queryDefConfig.Name | Measure-Object).Count -gt 1) { 
        $syncHash.Window.Dispatcher.Invoke([Action]{ $syncHash.searchPropSelection.ItemsSource = $configHash.queryDefConfig.Name }) 
    }

     $syncHash.searchPropSelection.Items | ForEach-Object {$syncHash.Window.Dispatcher.Invoke([Action]{$syncHash.searchPropSelection.SelectedItems.Add(($_))})}
}
   
function Set-RTDefaults {
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory, ValueFromPipeline)]$Type,
        [Parameter(Mandatory)][Hashtable]$ConfigHash
    )

    Begin {

        if ([string]::IsNullOrEmpty($configHash.rtConfig)) {
            $configHash.rtConfig = @{ }
        }
        
        else {
            $rtTemp = $configHash.rtConfig
            $configHash.rtConfig = @{ }
            $rtTemp.PSObject.Properties | ForEach-Object { $configHash.rtConfig[$_.Name] = $_.Value }
        }
    }

    Process {

        if ([string]::IsNullOrEmpty($configHash.rtConfig.$Type)) {
            $configHash.rtConfig.$type = Get-InitialValues -GroupName $type
            if ($configHash.rtConfig.$type.Path) {
                $configHash.rtConfig.$type.Icon = Get-Icon -Path $configHash.rtConfig.$type.Path -ToBase64
            }
        }
    }
}

function Add-CustomItemBoxControls {
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory)][Hashtable]$SyncHash,
        [Parameter(Mandatory)][Hashtable]$ConfigHash
    )

    $syncHash.customContext = @{ }

    foreach ($type in @('ubox', 'cbox')) {
    
        switch ($type) {
            'ubox' { $typeName = 'user' }
            'cbox' { $typeName = 'comp' }
        }

        for ($i = 1; $i -le $configHash.boxMax; $i++) {

            if ($i -le $configHash.($typeName + 'boxCount')) {
            
                $syncHash.(($type + $i + 'resources')) = @{

                    ($type + $i + 'Border')      = New-Object System.Windows.Controls.Border -Property @{
                        Style = $syncHash.Window.FindResource('itemBorder')
                        Name  = ($type + $i + 'Border')
                    }

                    ($type + $i + 'Grid')        = New-Object System.Windows.Controls.Grid
                    ($type + $i + 'DockPanel')   = New-Object System.Windows.Controls.DockPanel
                    ($type + $i + 'StackPanel')  = New-Object System.Windows.Controls.StackPanel -Property @{Style = $syncHash.Window.FindResource('itemStackPanel') }

                    ($type + $i + 'Header')      = New-Object System.Windows.Controls.Label -Property @{
                        Style = $syncHash.Window.FindResource('itemBoxHeader')
                        Name  = ($type + $i + 'Header')           
                    }

                    ($type + $i + 'EditClip')    = New-Object System.Windows.Controls.Label -Property @{
                        Style = $syncHash.Window.FindResource('itemEditClip')
                        Name  = ($type + $i + 'EditClip')
                    }

                    ($type + $i + 'ViewBox')     = New-Object System.Windows.Controls.ViewBox -Property @{Style = $syncHash.Window.FindResource('itemViewBox') }

                    ($type + $i + 'TextBox')     = New-Object System.Windows.Controls.TextBox -Property @{
                        Style = $syncHash.Window.FindResource('itemBox')
                        Name  = ($type + $i + 'TextBox')
                    }

                    ($type + $i + 'Box1Action1') = New-Object System.Windows.Controls.Button -Property @{
                        Style = $syncHash.Window.FindResource('itemButton')
                        Name  = ($type + $i + 'Box1Action1')
                    }

                    ($type + $i + 'Box1Action2') = New-Object System.Windows.Controls.Button -Property @{
                        Style = $syncHash.Window.FindResource('itemButton')
                        Name  = ($type + $i + 'Box1Action2')
                    }
            
                    ($type + $i + 'Box1')        = New-Object System.Windows.Controls.Button -Property @{
                        Style = $syncHash.Window.FindResource('itemEditButton')
                        Name  = ($type + $i + 'Box1')
                    }

                }
        
                # Add col def objects, then add to outside grid
                $colDef1 = New-Object System.Windows.Controls.ColumnDefinition
                $colDef1.Width = "*"

                $colDef2 = New-Object System.Windows.Controls.ColumnDefinition
                $colDef2.Width = "Auto"
                $colDef2.MaxWidth = "90"   

                $syncHash.(($type + $i + 'resources')).($type + $i + 'Grid').ColumnDefinitions.Add($colDef1)
                $syncHash.(($type + $i + 'resources')).($type + $i + 'Grid').ColumnDefinitions.Add($colDef2)

                # add child controls

                $syncHash.(($type + $i + 'resources')).($type + $i + 'Border').AddChild($syncHash.(($type + $i + 'resources')).($type + $i + 'Grid'))
                $syncHash.(($type + $i + 'resources')).($type + $i + 'Grid').AddChild($syncHash.(($type + $i + 'resources')).($type + $i + 'DockPanel'))
                $syncHash.(($type + $i + 'resources')).($type + $i + 'DockPanel').AddChild($syncHash.(($type + $i + 'resources')).($type + $i + 'Header'))
                $syncHash.(($type + $i + 'resources')).($type + $i + 'DockPanel').AddChild($syncHash.(($type + $i + 'resources')).($type + $i + 'EditClip'))
                $syncHash.(($type + $i + 'resources')).($type + $i + 'Grid').AddChild($syncHash.(($type + $i + 'resources')).($type + $i + 'ViewBox'))
                $syncHash.(($type + $i + 'resources')).($type + $i + 'ViewBox').AddChild($syncHash.(($type + $i + 'resources')).($type + $i + 'TextBox'))
                $syncHash.(($type + $i + 'resources')).($type + $i + 'Grid').AddChild($syncHash.(($type + $i + 'resources')).($type + $i + 'StackPanel'))
                $syncHash.(($type + $i + 'resources')).($type + $i + 'StackPanel').AddChild($syncHash.(($type + $i + 'resources')).($type + $i + 'Box1'))
                $syncHash.(($type + $i + 'resources')).($type + $i + 'StackPanel').AddChild($syncHash.(($type + $i + 'resources')).($type + $i + 'Box1Action1'))
                $syncHash.(($type + $i + 'resources')).($type + $i + 'StackPanel').AddChild($syncHash.(($type + $i + 'resources')).($type + $i + 'Box1Action2'))
    
                # add it top item to uniform grid

                if ($type -eq 'ubox') {
                    $syncHash.userDetailGrid.AddChild($syncHash.(($type + $i + 'resources')).($type + $i + 'Border'))
                }
                else {
                    $syncHash.compDetailGrid.AddChild($syncHash.(($type + $i + 'resources')).($type + $i + 'Border'))
                }
            }
        }
    }
}

function Add-CustomToolControls {
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory)][Hashtable]$SyncHash,
        [Parameter(Mandatory)][Hashtable]$ConfigHash
    )

    $syncHash.objectTools = @{ }
    #create custom tool buttons for queried objects
    foreach ($tool in $configHash.objectToolConfig) {
        if ($tool.toolActionValid) {
            switch ($tool.objectType) {
                { $_ -match "Both|Comp" } { 
        
                    $syncHash.objectTools.('ctool' + $tool.ToolID) = @{
                        ToolButton = New-Object System.Windows.Controls.Button -Property  @{
                            Style   = $syncHash.Window.FindResource('itemToolButton')
                            Name    = ('ctool' + $tool.ToolID)
                            Content = $tool.toolActionIcon
                            ToolTip = $tool.toolActionToolTip
                        }
                    }

                    $syncHash.compToolControlPanel.AddChild($syncHash.objectTools.('ctool' + $tool.ToolID).ToolButton)            

                }
                { $_ -match "Both|User" } { 
        
                    $syncHash.objectTools.('utool' + $tool.ToolID) = @{
                        ToolButton = New-Object System.Windows.Controls.Button -Property  @{
                            Style   = $syncHash.Window.FindResource('itemToolButton')
                            Name    = ('utool' + $tool.ToolID)
                            Content = $tool.toolActionIcon
                            ToolTip = $tool.toolActionToolTip
                        }
                    }

                    $syncHash.userToolControlPanel.AddChild($syncHash.objectTools.('utool' + $tool.ToolID).ToolButton)
           
                }
                { $_ -eq "Standalone" } { 
        
                $syncHash.objectTools.('tool' + $tool.ToolID) = @{
                        ToolButton = New-Object System.Windows.Controls.Button -Property  @{
                            Style   = $syncHash.Window.FindResource('standAloneButton')
                            Name    = ('tool' + $tool.ToolID)
                            Content = $tool.toolActionIcon
                            ToolTip = $tool.toolActionToolTip
                        }
                    }

                    $syncHash.standaloneControlPanel.AddChild($syncHash.objectTools.('tool' + $tool.ToolID).ToolButton)
           
                }           
            }

            foreach ($toolButton in (($syncHash.objectTools.Keys).Where{ ($_ -replace ".*tool") -eq $tool.ToolID })) {

                $syncHash.objectTools.$toolButton.ToolButton.Add_Click( {
                        param([Parameter(Mandatory)][Object]$sender)
                        $toolID = $sender.Name -replace ".*tool"
                    
                        switch ($configHash.objectToolConfig[$toolId - 1].toolType) {

                            'Execute' {

                                $syncHash.itemToolDialog.Title = "Confirm"

                                if ($configHash.objectToolConfig[$toolId - 1].toolActionConfirm) {
                       
                                    $syncHash.itemToolDialogConfirmActionName.Text = $configHash.objectToolConfig[$toolId - 1].ToolName
                                    $syncHash.itemToolDialogConfirmObjectName.Text = $configHash.currentTabItem
                                    $syncHash.itemToolDialogConfirm.Visibility = 'Visible'
                                    $syncHash.itemToolDialogConfirmButton.Tag = $toolID
                                    $syncHash.itemToolDialog.IsOpen = $true                      
                                }

                                else {
                                    Start-RSJob -Name ItemTool -ArgumentList $syncHash.snackMsg.MessageQueue, $toolID, $configHash, $queryHash  -ScriptBlock {
                                        Param($queue, $toolID, $configHash, $queryHash)

                                        $item = ($configHash.currentTabItem).toLower()
                                        $toolName = ($configHash.objectToolConfig[$toolID - 1].toolActionToolTip).ToUpper()

                                        try {                     
                                            Invoke-Expression $configHash.objectToolConfig[$toolID - 1].toolAction
                                            if ($configHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {
                                               $queue.Enqueue("[$toolName]: Success - Standalone tool complete")
                                               Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName $toolName -ArrayList $configHash.actionLog 
                                            }

                                            else {
                                                $queue.Enqueue("[$toolName]: Success on [$item] - tool complete")
                                                 Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -SubjectName $item -ActionName $toolName -ArrayList $configHash.actionLog 
                                            }
                                           
                                        }
                                        catch {
                                            if ($configHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {
                                               $queue.Enqueue("[$toolName]: Fail - Standalone tool incomplete") 
                                                Write-LogMessage -Path $configHash.actionlogPath -Message Fail -ActionName $toolName -ArrayList $configHash.actionLog -Error $_
                                            }
                                            else {
                                                $queue.Enqueue("[$toolName]: Fail on [$item] - tool incomplete")
                                                Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -SubjectName $item -ActionName $toolName -ArrayList $configHash.actionLog -Error $_
                                            }
                                        }
                                    }
                                }
                                break
                            }
                            'Select' {
                        
                                $syncHash.itemToolDialog.Title = $configHash.objectToolConfig[$toolId - 1].toolName
                                $syncHash.itemToolListSelectText.Text = $configHash.objectToolConfig[$toolId - 1].toolDescription
                                $syncHash.itemToolListSelect.Visibility = "Visible"
                                $syncHash.itemToolListSelectConfirmButton.Tag = $toolID 
                                $syncHash.itemToolListSelectListBox.ItemsSource = $null
                                $syncHash.itemToolDialog.IsOpen = $true 
                        

                                if ($configHash.objectToolConfig[$toolId - 1].toolActionSelectAD -eq $false) {
                                
                                    $syncHash.ItemToolADSelectionPanel.Visibility = "Collapsed"
                            
                                    Start-RSJob -Name PopulateListboxNoAD -ArgumentList $configHash, $syncHash, $toolID -ScriptBlock {
                                        param($configHash, $syncHash, $toolID)
                            
                                                      
                                        $syncHash.Window.Dispatcher.Invoke([Action] {                                    
                                                $syncHash.itemTooListBoxProgress.Visibility = "Visible"
                                            })
                    
                                        $list = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
                                        Invoke-Expression ($configHash.objectToolConfig[$toolId - 1].toolFetchCmd) | ForEach-Object { $list.Add([PSCUstomObject]@{'Name' = $_ }) }
                                 
                                        $syncHash.Window.Dispatcher.Invoke([Action] {
                                                $syncHash.itemToolListSelectListBox.ItemsSource = [System.Windows.Data.ListCollectionView]$list
                                                $syncHash.itemTooListBoxProgress.Visibility = "Collapsed"
                                            })
                                    } 
                                               
                                }

                                else {
                                    $syncHash.ItemToolADSelectionPanel.Visibility = "Visible"
                               
                                    $syncHash.itemToolListSelect.Visibility = 'Visible'
                                    $syncHash.itemToolListSelectListBox.ItemsSource = $null

                                }  
                            
                                if ($configHash.objectToolConfig[$toolId - 1].toolActionMultiSelect) {
                                    $syncHash.itemToolListSelectListBox.SelectionMode = "Multiple"
                                }

                                else {
                                    $syncHash.itemToolListSelectListBox.SelectionMode = "Single"
                                }  

                     
  

                            }
                            'Grid' {
                        
                                $syncHash.itemToolDialog.Title = $configHash.objectToolConfig[$toolId - 1].toolName
                                $syncHash.itemToolGridItemsGrid.ItemsSource = $null
                                $syncHash.itemToolGridSelectConfirmButton.Tag = $toolID 
                                $syncHash.itemToolGridSelectText.Text = $configHash.objectToolConfig[$toolId - 1].toolDescription
                                $syncHash.itemToolGridSelect.Visibility = "Visible"
                                $syncHash.itemToolDialog.IsOpen = $true  

                                if ($configHash.objectToolConfig[$toolId - 1].toolActionSelectAD -eq $false) {
                                
                                    $syncHash.itemToolGridADSelectionPanel.Visibility = "Collapsed"
                            
                                    Start-RSJob -Name PopulateGridbox -FunctionsToImport Get-Icon -ArgumentList $configHash, $syncHash, $toolID -ScriptBlock {
                                        param($configHash, $syncHash, $toolID)

                            
                                                      
                                        $syncHash.Window.Dispatcher.Invoke([Action] {                                  
                                                $syncHash.itemToolGridProgress.Visibility = "Visible"
                                            })

                                
                                        $list = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
                                        Invoke-Expression ($configHash.objectToolConfig[$toolId - 1].toolFetchCmd) | ForEach-Object { $list.Add($_) }
                                 
                                        $syncHash.Window.Dispatcher.Invoke([Action] {
                                                $syncHash.itemToolGridItemsGrid.ItemsSource = [System.Windows.Data.ListCollectionView]$list
                                                $syncHash.itemToolGridProgress.Visibility = "Collapsed"
                                            })
                                    } 
                                               
                                }

                                else {
                                    $syncHash.ItemToolADSelectionPanel.Visibility = "Visible"
                               
                                    $syncHash.itemToolGridSelect.Visibility = 'Visible'
                                    $syncHash.itemToolGridItemsGrid.ItemsSource = $null

                                }  
                            
                                if ($configHash.objectToolConfig[$toolId - 1].toolActionMultiSelect) {
                                    $syncHash.itemToolGridItemsGrid.SelectionMode = "Extended"
                                }

                                else {
                                    $syncHash.itemToolGridItemsGrid.SelectionMode = "Single"
                                }  

                            }

                        }
                    })           
            }
        }
    }
}

function Start-BasicADCheck {
    param ($SysCheckHash, $configHash) 

    if ((Get-WmiObject -Class Win32_ComputerSystem).PartofDomain) {               
        $sysCheckHash.sysChecks[0].ADMember = 'True'
                    
        if ($sysCheckHash.sysChecks[0].ADModule -eq $true) {                                                 
            $selectedDC = Get-ADDomainController -Discover -Service ADWS -ErrorAction SilentlyContinue 

            if (Test-Connection -Count 1 -Quiet -ComputerName $selectedDC.HostName) {             
                $sysCheckHash.sysChecks[0].ADDCConnectivity = 'True'
                            
                try {
                    $adEntity = [Microsoft.ActiveDirectory.Management.ADEntity].Assembly
                    $adFields = $adEntity.GetType('Microsoft.ActiveDirectory.Management.Commands.LdapAttributes').GetFields('Static,NonPublic') | Where-Object { $_.IsLiteral }
                    $configHash.adPropertyMap = @{ }
                            
                    $adFields | ForEach-Object {                                       
                        $configHash.adPropertyMap[$_.Name] = $_.GetRawConstantValue()
                    }
                } 

                catch { $configHash.Remove('adPropertyMap') }

            }                                                  
        }  
    }
}

function Start-AdminCheck {
    param ($SysCheckHash) 
    if ((New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).isInRole([Security.Principal.WindowsBUiltInRole]::Administrator)) {
        $sysCheckHash.sysChecks[0].IsInAdmin = 'True'
    }

    if ((New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).isInRole("Domain Admins")) {
        $sysCheckHash.sysChecks[0].IsDomainAdmin = 'True'
    }
}

function Set-RSDataContext {
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory, ValueFromPipeline)]$ControlName,
        [Parameter(Mandatory)][Hashtable]$SyncHash,
        [Parameter(Mandatory)]$DataContext   
    )

    Process {
        $syncHash.Window.Dispatcher.invoke([action] { $syncHash.$ControlName.DataContext = $DataContext })
    }
}

function Get-AdObjectPropertyList {
    param ($configHash) 

    $configHash.rawADValues = [System.Collections.ArrayList]@()
    $configHash.QueryADValues = @{}         
      
    (Get-ADObject -Filter "(SamAccountName -eq '$($env:USERNAME)') -or (SamAccountName -eq '$($env:ComputerName)$')" -Properties *) | 
        ForEach-Object { $_.PSObject.Properties | 
            ForEach-Object {$configHash.rawADValues.Add($_.Name) | Out-Null}}

    foreach ($ldapValue in ($configHash.rawADValues | Sort-Object -Unique)) {
        $configHash.QueryADValues.(($configHash.adPropertyMap.GetEnumerator() | Where-Object { $_.Value -eq $ldapValue}).Key) = $ldapValue
    }

    $configHash.QueryADValues = ($configHash.QueryADValues.GetEnumerator().Where({$_.Key}))
}

function Get-PropertyLists {
    param ($ConfigHash) 
 
    foreach ($type in ("user", "comp")) {
        
        if ($type -eq 'user') {
            $configHash.userPropPullList = (Get-ADUser -Identity $env:USERNAME -Properties *).PSObject.Properties | Select-Object Name, TypeNameofValue
        }
        else {
            $configHash.compPropPullList = (Get-ADComputer -Identity $env:COMPUTERNAME -Properties *).PSObject.Properties | Select-Object Name, TypeNameofValue
        }
        
        Get-AdObjectPropertyList -ConfigHash $configHash        
                
        $configHash.($type + 'PropPullListNames') = [System.Collections.ArrayList]@()
        $configHash.($type + 'PropPullListNames').Add("Non-AD Property") 
        $configHash.($type + 'PropPullList').Name | ForEach-Object { $configHash.($type + 'PropPullListNames').Add($_) }
    }
}

function Set-ADGenericQueryNames {
    param($ConfigHash) 

    foreach ($id in ($configHash.queryDefConfig.ID)) {
            $configHash.queryDefConfig[$id - 1].QueryDefTypeList = $configHash.QueryADValues.Key
        }

}



function Start-PropBoxPopulate {
    param ($configHash)

    Get-PropertyLists -ConfigHash $configHash

    foreach ($type in @('User', 'Comp')) {   

        $tempList = $configHash.($type + 'PropList')
        $configHash.($type + 'PropList') = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
 
        for ($i = 1; $i -le $configHash.boxMax; $i++) {

            if ($i -le $configHash.($type + 'boxCount')) {
                $configHash.($type + 'PropList').Add([PSCustomObject]@{
                        Field             = $i
                                    
                        FieldName         = ( $tempList | Where-Object { $_.Field -eq $i }).FieldName
                                    
                        ItemType          = $type
                                    
                        PropName          = ( $tempList | Where-Object { $_.Field -eq $i }).PropName
                                    
                        propList          = $configHash.($type + 'PropPullListNames')
                                    
                        translationCmd    = if (($tempList | Where-Object { $_.Field -eq $i }).ValidCmd -eq $true) { ($tempList | Where-Object { $_.Field -eq $i }).translationCmd } 
                        else { 'if ($result -eq $false)...' }
                                    
                        actionCmd1        = if (($tempList | Where-Object { $_.Field -eq $i }).ValidAction1 -eq $true) { ($tempList | Where-Object { $_.Field -eq $i }).actionCmd1 } 
                        else { "Do-Something -$type $('$' + $type)..." }
                                    
                        actionCmd1ToolTip = if (($tempList | Where-Object { $_.Field -eq $i }).ValidAction1 -eq $true) { ($tempList | Where-Object { $_.Field -eq $i }).actionCmd1ToolTip } 
                        else { 'Action name' }
                  
                        actionCmd1Icon    = if (($tempList | Where-Object { $_.Field -eq $i }).ValidAction1 -eq $true) { ($tempList | Where-Object { $_.Field -eq $i }).actionCmd1Icon } 
                        else { $null }
                                    
                        actionCmd1Refresh = if (($tempList | Where-Object { $_.Field -eq $i }).actionCmd1Refresh) { $true } 
                        else { $false }  
                                    
                        actionCmd1Multi   = if (($tempList | Where-Object { $_.Field -eq $i }).actionCmd1Multi) { $true } 
                        else { $false }  
                                    
                        ValidCmd          = if (($tempList | Where-Object { $_.Field -eq $i }).ValidCmd) { $true } 
                        else { $false }
                                    
                        ValidAction1      = if (($tempList | Where-Object { $_.Field -eq $i }).ValidAction1) { $true } 
                        else { $false } 
                                    
                        ValidAction2      = if (($tempList | Where-Object { $_.Field -eq $i }).ValidAction2) { $true } 
                        else { $false } 
                                    
                        actionCmd2        = if (($tempList | Where-Object { $_.Field -eq $i }).ValidAction2 -eq $true) { ($tempList | Where-Object { $_.Field -eq $i }).actionCmd2 } 
                        else { "Do-Something -$type $('$' + $type)..." }
                                    
                        actionCmd2ToolTip = if (($tempList | Where-Object { $_.Field -eq $i }).ValidAction2 -eq $true) { ($tempList | Where-Object { $_.Field -eq $i }).actionCmd2ToolTip } 
                        else { 'Action name' }
                  
                        actionCmd2Icon    = if (($tempList | Where-Object { $_.Field -eq $i }).ValidAction1 -eq $true) { ($tempList | Where-Object { $_.Field -eq $i }).actionCmd2Icon } 
                        else { $null }
                                    
                        actionCmd2Refresh = if (($tempList | Where-Object { $_.Field -eq $i }).actionCmd2Refresh) { $true } 
                        else { $false }  
                                    
                        actionCmd2Multi   = if (($tempList | Where-Object { $_.Field -eq $i }).actionCmd2Multi) { $true } 
                        else { $false }                       
                                    
                        querySubject      = if ($type -eq 'user') { $env:USERNAME }
                        else { $env:COMPUTERNAME }
                                       
                        result            = '(null)'
                                    
                        actionCmdsEnabled = if (($tempList | Where-Object { $_.Field -eq $i }).actionCmdsEnabled -eq $false) { $false }
                        else { $true }
                                    
                        transCmdsEnabled  = if (($tempList | Where-Object { $_.Field -eq $i }).transCmdsEnabled -eq $false) { $false }
                        else { $true }
                                    
                        actionCmd1result  = '(null)'
                                    
                        actionCmd2result  = '(null)'
                                    
                        actionCmd2Enabled = if (($tempList | Where-Object { $_.Field -eq $i }).ValidAction2 -eq $false) { $false }
                        elseif (($tempList | Where-Object { $_.Field -eq $i }).actionCmd2Enabled) { $true } 
                        else { $false }  
                                    
                        PropType          = (($configHash.($type + 'PropPullList') | Where-Object { $_.Name -eq (($tempList | Where-Object { $_.Field -eq $i }).PropName) }).TypeNameOfValue -replace ".*(?=\.).", "")
                                    
                        actionList        = @('ReadOnly', 'ReadOnly-Raw', 'Editable', 'Editable-Raw', 'Actionable', 'Actionable-Raw', 'Editable-Actionable', 'Editable-Actionable-Raw')
                      
                                    
                        ActionName        = if (($tempList | Where-Object { $_.Field -eq $i }).ActionName) { ($tempList | Where-Object { $_.Field -eq $i }).ActionName } 
                        else { 'null' }
              
                    })
            }        
        }                        
    }
}

function Set-QueryVarsToUpdate {
    param ($configHash,[Parameter(Mandatory)][ValidateSet('User', 'Comp')]$Type)

    
    switch ($type) {
        'User' { if ($configHash.varListConfig.UpdateFrequency -match "User Queries|All Queries")  {$configHash.varData.UpdateUser = $true } }

        'Comp' { if ($configHash.varListConfig.UpdateFrequency -match "Comp Queries|All Queries")  {$configHash.varData.UpdateComp = $true } }
    }
}

function Set-DurationVarsToUpdate {
    param ($configHash, $startTime)

    $currentTime = Get-Date

    switch ($startTime) {
        {($configHash.varListConfig.UpdateFrequency -contains 'Daily') -and ($startTime.AddDays($configHash.VarData.UpdateDayCount) -le $currentTime)} {
            $configHash.varData.UpdateDay = $true
            $configHash.VarData.UpdateDayCount = $configHash.VarData.UpdateDayCount + 1
        }

        {($configHash.varListConfig.UpdateFrequency -contains 'Hourly') -and ($startTime.AddHours($configHash.VarData.UpdateHourCount) -le $currentTime)} {
            $configHash.varData.UpdateHour = $true
            $configHash.VarData.UpdateHourCount = $configHash.VarData.UpdateHourCount + 1
        }
        
        {($configHash.varListConfig.UpdateFrequency -contains 'Every 15 mins') -and ($startTime.AddMinutes($configHash.varData.UpdateMinCount) -le $currentTime)} {
            $configHash.varData.UpdateMinute = $true
            $configHash.varData.UpdateMinCount = $configHash.varData.UpdateMinCount + 15
        }
    
    }
}

function New-VarUpdater {
    param ($configHash)
       
    $configHash.varData = @{
        UpdateDayCount = 1   
        UpdateMinCount = 15
        UpdateHourCount = 1
    }  
}

function Start-VarUpdater {
    [CmdletBinding()]
    param ($configHash, $varHash)

    Start-RSJob -Name VarUpdater -ArgumentList $configHash, $varHash -ModulesToImport C:\TempData\internal\internal.psm1 {
        param($configHash, $varHash)

        $startTime = Get-Date

        do {
            
            Set-DurationVarsToUpdate -ConfigHash $configHash -StartTime $startTime

            if ($configHash.varData.ContainsValue($true)) {
                foreach ($varInfo in ($configHash.varData.Keys)) {
                    if ($configHash.varData.$varInfo -eq $true) {
                        switch ($varInfo) {
                            'UpdateMinute'   {
                                 $configHash.varListConfig | Where-Object {$_.UpdateFrequency -eq "Every 15 mins"} | ForEach-Object {
                                    $varHash.($_.VarName) = Invoke-Expression $_.VarCmd }
                                    $configHash.varData.$varInfo = $false
                                }
                            'UpdateHour' {
                                $configHash.varListConfig | Where-Object {$_.UpdateFrequency -eq "Hourly"} | ForEach-Object {
                                    $varHash.($_.VarName) = Invoke-Expression $_.VarCmd }
                                    $configHash.varData.$varInfo = $false
                                }                                       
                            'UpdateDay'  {
                                $configHash.varListConfig | Where-Object {$_.UpdateFrequency -eq "Daily"} | ForEach-Object {
                                    $varHash.($_.VarName) = Invoke-Expression $_.VarCmd }
                                    $configHash.varData.$varInfo = $false
                                    }
                            'UpdateUser' { 
                                $configHash.varListConfig | Where-Object {$_.UpdateFrequency -match "User Queries|All Queries"} | ForEach-Object {
                                    $varHash.($_.VarName) = Invoke-Expression $_.VarCmd }
                                    $configHash.varData.$varInfo = $false
                                }
                            'UpdateComp' {
                                $configHash.varListConfig | Where-Object {$_.UpdateFrequency -match "Comp Queries|All Queries"} | ForEach-Object {
                                    $varHash.($_.VarName) = Invoke-Expression $_.VarCmd }
                                    $configHash.varData.$varInfo = $false
                                }
                            }
                        }
                    }
                }
    
        } until ($configHash.IsClosed -eq $true) 
    }
}

#endregion

#region setting child window

Function Set-ChildWindow {
    param (
        $Panel,
        $Title,
        $Height,
        $Width,
        [Parameter(Mandatory)][Hashtable]$SyncHash,
        [switch]$HideCloseButton,
        [ValidateSet('Standard', 'Flyout')][string[]]$Background
    )

    if ($Panel) {
        $syncHash.$Panel.Visibility = "Visible"
        $syncHash.settingChildWindow.IsOpen = $true

        switch ($Panel) {

            'settingUserPropContent' { 
                if ($configHash.UserPropList -ne $null) { $syncHash.settingUserPropGrid.ItemsSource = $configHash.UserPropList }
                $syncHash.settingPropContent.Visibility = "Visible"
                $syncHash.settingUserPropGrid.Visibility = "Visible"
            }

            'settingCompPropContent' {
                if ($configHash.CompPropList -ne $null) { $syncHash.settingCompPropGrid.ItemsSource = $configHash.CompPropList }
                $syncHash.settingPropContent.Visibility = "Visible"
                $syncHash.settingCompPropGrid.Visibility = "Visible"
            }

            'settingItemToolsContent' {
                if ($configHash.objectToolConfig -ne $null) { $syncHash.settingObjectToolsPropGrid.ItemsSource = $configHash.objectToolConfig }
                $syncHash.settingObjectToolsPropGrid.Visibility = "Visible"
            }
        
            'settingContextPropContent' {
                if ($configHash.contextConfig -ne $null) { $syncHash.settingContextPropGrid.ItemsSource = $configHash.contextConfig }
                $syncHash.settingContextGrid.Visibility = "Visible"
                $syncHash.settingContextPropGrid.Visibility = "Visible"
            }

            'settingVarContent' {            
                if ($configHash.varListConfig -ne $null) { $syncHash.settingVarDataGrid.ItemsSource = $configHash.varListConfig }
                $syncHash.settingVarDataGrid.Visibility = "Visible"

            }

            'settingOUDataGrid' {
                if ($configHash.searchbaseConfig -ne $null) { $syncHash.settingOUDataGrid.ItemsSource = $configHash.searchbaseConfig }
                $syncHash.settingOUDataGrid.Visibility = "Visible" 
                $syncHash.settingGeneralAddClick.Tag = "OU"           
            }    
            
            'settingQueryDefDataGrid' { 
                if ($configHash.queryDefConfig -ne $null) { $syncHash.settingQueryDefDataGrid.ItemsSource = $configHash.queryDefConfig }
                $syncHash.settingQueryDefDataGrid.Visibility = "Visible" 
                $syncHash.settingGeneralAddClick.Tag = "Query"   
            } 
            
            'settingMiscGrid' { 
            $syncHash.settingGeneralAddClick.Tag = 'null'
            $syncHash.settingMiscGrid.Visibility = "Visible" 
                  
            } 
        }
    }

    if ($Title) { $syncHash.settingChildWindow.Title = $Title }
    if ($Height) { $syncHash.settingChildHeight.Height = $Height }
    if ($Width) { $syncHash.settingChildHeight.Width = $Width }

    switch ($Background) {       
        'Standard' { $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.Window.BorderBrush.Color).ToString(); break }
        'Flyout' { $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.settingNameFlyout.Background.Color).ToString() }
    }

    if ($HideCloseButton) { $SyncHash.settingChildWindow.ShowCloseButton = $false }
    else { $SyncHash.settingChildWindow.ShowCloseButton = $true }

}

function Reset-ChildWindow {
    param (
        [Parameter(Mandatory)][Hashtable]$SyncHash,
        [switch]$SkipResize,
        [switch]$SkipContentPaneReset,
        [switch]$SkipDataGridReset,
        [switch]$SkipFlyoutClose,
        [string]$Title
    )

    if (!($SkipContentPaneReset)) {
        foreach ($contentPane in ($syncHash.Keys.Where( { $_ -like "setting*Content" }))) {
            $syncHash.$contentPane.Visibility = "Hidden"
        }
    }

    if (!($SkipDataGridReset)) {
        foreach ($dataGrid in ($syncHash.Keys.Where( { $_ -like "setting*Grid" }))) {
            $syncHash.$dataGrid.Visibility = "Collapsed"
        }
    }

    if (!($SkipFlyoutClose)) {
        foreach ($flyOut in ($syncHash.Keys.Where( { $_ -like "setting*Flyout" }))) {
            $syncHash.$flyOut.IsOpen = $false
        }
    }

    if ($Title) { Set-ChildWindow -SyncHash $syncHash -Title $title }

    if (!($SkipResize)) { Set-ChildWindow -SyncHash $syncHash -Width 400 -Height 215 }
    
    Set-ChildWindow -SyncHash $syncHash -Background Standard

    if ($syncHash.settingStatusChildBoard.Visibility -eq 'Collapsed') {
        $syncHash.settingStatusChildBoard.Visibility = 'Visible'
        $syncHash.settingConfigPanel.Visibility = 'Visible'
    }

}

#endregion

#region itemToolFunctions

function Set-ADItemBox {
    param(
        [Parameter(Mandatory)][hashtable]$ConfigHash, 
        [Parameter(Mandatory)][hashtable]$SyncHash,
        [Parameter(Mandatory)][ValidateSet('ListBox', 'Grid')][string]$Control) 


    Start-RSJob -Name ('populate' + $Control) -ArgumentList $configHash, $syncHash, $syncHash.itemToolListSelectConfirmButton.Tag, $Control -FunctionsToImport Select-ADObject -ScriptBlock {
        param($configHash, $syncHash, $toolID, $Control)


        $syncHash.Window.Dispatcher.Invoke([Action] {
                if ($Control -eq 'ListBox') { $syncHash.itemToolListSelectListBox.ItemsSource = $null }
                else { $syncHash.itemToolGridItemsGrid.ItemsSource = $null }
            })

        $selectedObject = (Select-ADObject -Type All -MultiSelect $false).FetchedAttributes -replace '$'
     
        $syncHash.Window.Dispatcher.Invoke([Action] {
                if ($Control -eq 'ListBox') { 
                    $syncHash.itemToolADSelectedItem.Content = $selectedObject
                    $syncHash.itemTooListBoxProgress.Visibility = "Visible"
                }
                else {
                    $syncHash.itemToolGridADSelectedItem.Content = $selectedObject
                    $syncHash.itemToolGridProgress.Visibility = "Visible"
                }
            })
     
        $list = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
        
        if ($Control -eq 'ListBox') {
            Invoke-Expression ($configHash.objectToolConfig[$toolId - 1].toolFetchCmd) | ForEach-Object { $list.Add([PSCUstomObject]@{'Name' = $_ }) }
        }
        else {
            Invoke-Expression ($configHash.objectToolConfig[$toolId - 1].toolFetchCmd) | ForEach-Object { $list.Add($_) }
        }

        $syncHash.Window.Dispatcher.Invoke([Action] {
                if ($Control -eq 'ListBox') {
                    $syncHash.itemToolListSelectListBox.ItemsSource = [System.Windows.Data.ListCollectionView]$list
                    $syncHash.itemTooListBoxProgress.Visibility = "Collapsed"
                }
                else { 
                    $syncHash.itemToolGridItemsGrid.ItemsSource = [System.Windows.Data.ListCollectionView]$list 
                    $syncHash.itemToolGridProgress.Visibility = "Collapsed"
                }
            })
    }      
}

function Start-ItemToolAction {
    param(
        [Parameter(Mandatory)][hashtable]$ConfigHash, 
        [Parameter(Mandatory)][hashtable]$SyncHash,
        [Parameter(Mandatory)][ValidateSet('List', 'Grid')][string]$Control,
        $ItemList)

    Start-RSJob -ArgumentList $configHash, $itemList, $syncHash.snackMsg.MessageQueue, $syncHash.('itemTool' + $Control + 'SelectConfirmButton').Tag -ScriptBlock {
        param($configHash, $itemList, $queue, $toolID) 

        $toolName = $configHash.objectToolConfig[$toolID - 1].toolActionToolTip
        $target = $configHash.currentTabItem

        try {

            Invoke-Expression -Command $configHash.objectToolConfig[$toolID - 1].toolTargetFetchCmd

            foreach ($selectedItem in $itemList) {
                Invoke-Expression -Command $configHash.objectToolConfig[$toolID - 1].toolAction
            }

            $queue.Enqueue("[$toolName]: SUCCESS: tool ran on $target")
             Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -SubjectName $Target -ActionName $toolName -ArrayList $configHash.actionLog 

        }
        
        catch {
            $queue.Enqueue("[$toolName]: FAIL: tool incomplete on $target") 
            Write-LogMessage -Path $configHash.actionlogPath -Message Fail -SubjectName $Target -ActionName $toolName -ArrayList $configHash.actionLog -Error $_
        }
    }

    $syncHash.itemToolDialog.IsOpen = $false
}



#endregion

#region RemoteTools

function Get-RTFlyoutContent {
    param (
        $SyncHash,
        $ConfigHash
    )

    $syncHash.settingRemoteListTypes.ItemsSource = $configHash.nameMapList

    switch ($syncHash.settingRALabel.Content) {
            
        'MSTSC' {
            if ($configHash.rtConfig.MSTSC.Icon) {
                $syncHash.settingRTIcon.Source = ([Convert]::FromBase64String($configHash.rtConfig.MSTSC.Icon))
            }

            $syncHash.settingRemoteListTypes.Items | Where-Object { $_.Name -in $configHash.rtConfig.MSTSC.Types } | ForEach-Object { 
                $syncHash.settingRemoteListTypes.SelectedItems.Add(($_))
            }
                
            break
        }

        'MSRA' {
            if ($configHash.rtConfig.MSRA.Icon) {
                $syncHash.settingRTIcon.Source = ([Convert]::FromBase64String($configHash.rtConfig.MSRA.Icon))
            }
            $syncHash.settingRemoteListTypes.Items | Where-Object { $_.Name -in $configHash.rtConfig.MSRA.Types } | ForEach-Object { 
                $syncHash.settingRemoteListTypes.SelectedItems.Add(($_))
            }
            break
        }

        Default {
            $rtID = 'rt' + [string]($syncHash.settingRALabel.Content -replace ".[A-Z]* ")

            if ($configHash.rtConfig.$rtID.Icon) {
                $syncHash.settingRTIcon.Source = ([Convert]::FromBase64String($configHash.rtConfig.$rtID.Icon))
            }
               
            $syncHash.settingRemoteListTypes.Items | Where-Object { $_.Name -in $configHash.rtConfig.$rtID.Types } | ForEach-Object { 
                $syncHash.settingRemoteListTypes.SelectedItems.Add(($_))
            }
        }
    }
}

function Set-SelectedRTTypes {
    param (
        $SyncHash,
        $ConfigHash
    )

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

}

function Set-StaticRTContent {
    param (
        $SyncHash,
        $ConfigHash,
        [ValidateSet('MSRA', 'MSTSC')][string]$Tool
    )

    $syncHash.settingRtExeSelect.Visibility = "Hidden"
    $syncHash.settingRtPathSelect.Visibility = "Hidden"
    $syncHash.rtSettingRequiresOnline.Visibility = "Hidden"
    $syncHash.rtSettingRequiresUser.Visibility = "Hidden"
    $syncHash.rtDock.DataContext = $configHash.rtConfig.$tool
    $syncHash.settingRemoteFlyout.isOpen = $true
    Set-ChildWindow -SyncHash $syncHash -Title "Remote Tool Options ($tool)" -Background Flyout
}

function Get-RTExePath {
    param (
        $SyncHash,
        $ConfigHash
    )

    $rtID = 'rt' + [string]($syncHash.settingRALabel.Content -replace ".[A-Z]* ")

    $customSelection = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        initialDirectory = [Environment]::GetFolderPath('ProgramFilesx86')
        Title            = "Select Custom RT Executable"
    }

    $customSelection.ShowDialog() | Out-Null

    if (![string]::IsNullOrEmpty($customSelection.fileName)) {

        if (Test-Path $customSelection.fileName) {
            $configHash.rtConfig.$rtID.Path = $customSelection.fileName
            $syncHash.settingRtPathSelect.Text = $customSelection.fileName

            if ($customSelection.fileName -like "\\*") {
                Copy-Item $customSelection.fileName -Destination C:\tmp.exe
                $configHash.rtConfig.$rtID.Icon = Get-Icon -Path C:\tmp.exe -ToBase64
                Remove-Item c:\tmp.exe
            }

            else { $configHash.rtConfig.$rtID.Icon = Get-Icon -Path $customSelection.fileName -ToBase64 }

            $syncHash.settingRTIcon.Source = ([Convert]::FromBase64String(($configHash.rtConfig.$rtID.Icon)))
        }
    }
}

function New-CustomRTConfigControls {
    param (
        $ConfigHash,
        $SyncHash,
        $RTID,
        [switch]$NewTool
    )

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
            Text  = if ($NewTool) { "Custom remote tool $($rtID -replace 'rt')" }
            else { $configHash.rtConfig.$rtID.DisplayName }
            Style = $syncHash.Window.FindResource('rtSubHeader')
        }
            
        ConfigureButton = New-Object System.Windows.Controls.Button -Property  @{
            Style = $syncHash.Window.FindResource('rtClick')
        }
            
        DelButton       = New-Object System.Windows.Controls.Button -Property  @{
            Style = $syncHash.Window.FindResource('rtClickDel')
        }
    }

    $syncHash.customRt.$rtID.ConfigureButton.Name = $rtID
    $syncHash.customRt.$rtID.DelButton.Name = $rtID + 'del'
    
    $syncHash.customRt.$rtID.parentDock.AddChild($syncHash.customRt.$rtID.childStack)
    $syncHash.customRt.$rtID.parentDock.AddChild($syncHash.customRt.$rtID.ConfigureButton)
    $syncHash.customRt.$rtID.parentDock.AddChild($syncHash.customRt.$rtID.DelButton)
    $syncHash.customRt.$rtID.childStack.AddChild($syncHash.customRt.$rtID.InfoHeader)
    $syncHash.customRt.$rtID.childStack.AddChild($syncHash.customRt.$rtID.InfoSubheader)
  
    $syncHash.settingRTPanel.AddChild($syncHash.customRt.$rtID.parentDock)

    if ($NewTool) {
        $configHash.rtConfig.$rtID = [PSCustomObject]@{
            Name          = "Custom Tool $($rtID -replace 'rt')"
            Path          = $null
            Icon          = $null
            Cmd           = " "
            Types         = @()
            RequireOnline = $true
            RequireUser   = $false
            DisplayName   = 'Tool'
        }
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
            Set-ChildWindow -SyncHash $syncHash -Title "Remote Tool Options ($rtID)" -Background Flyout
            $syncHash.settingRemoteFlyout.isOpen = $true
        })
}

function Add-CustomRTControls {
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory)][Hashtable]$SyncHash,
        [Parameter(Mandatory)][Hashtable]$ConfigHash
    )


    $syncHash.customRT = @{ }
    foreach ($rtID in $configHash.rtConfig.Keys.Where{ $_ -like "RT*" }) {
        New-CustomRTConfigControls -ConfigHash $configHash -SyncHash $syncHash -RTID $RtID   
    }
}


#endregion

#region NetworkingMapping

function Set-NetworkMapItem {
    param (
        $SyncHash,
        $ConfigHash,
        [switch]$Import
    )

    if ($Import) {

        $configHash.netMapList = (New-Object System.Collections.ObjectModel.ObservableCollection[Object])
        $subnets = Get-ADReplicationSubnet -Filter * -Properties * | Select-Object Name, Site, Location, Description

        if (!($subnets)) {
            $localAddress = ((Get-NetIPInterface -AddressFamily IPv4 | Get-NetIPAddress | Where-Object { $_.PrefixOrigin -ne 'WellKnown' }))
            
            foreach ($address in $localAddress) {
                $ip = [ipaddress]$address.IPAddress
                $subNet = [ipaddress]([ipaddress]([math]::pow(2, 32) - 1 -bxor [math]::pow(2, (32 - $($address.PrefixLength))) - 1))
                $netid = [ipaddress]($ip.address -band $subnet.address)
          
                $configHash.netMapList.Add([PSCustomObject]@{
                        Id           = ($configHash.netMapList | Measure-Object).Count + 1
                        Network      = $netid.IPAddressToString
                        ValidNetwork = $true
                        Mask         = $address.PrefixLength
                        ValidMask    = $true
                        Location     = "Default"
                    })                                                               
            }    
        }

        else {

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
        }

        $syncHash.settingNetDataGrid.ItemsSource = $configHash.netMapList
    
    }

    else {
        $configHash.netMapList.Add([PSCustomObject]@{
                ID           = ($configHash.netMapList.ID | Sort-Object -Descending | Select-Object -First 1) + 1
                Network      = $null
                ValidNetwork = $false
                Mask         = $null
                ValidMask    = $false
                Location     = "New"
            })    
    }
}


#endregion

#region UserLog

function Set-LogMapGrid { 
    param (
        $SyncHash,
        $ConfigHash,
        [parameter(Mandatory)][ValidateSet('Comp', 'User')]$Type)

    # Get last 10 entries of newest logged item; select latest of these entries not ending 
    # in a comma (which would indicate that var was empty on login)
    $testLog = Get-Content ((Get-ChildItem -Path $confighash.($type + 'LogPath') |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName) | Select-Object -Last 11 |
    Where-Object { $_.Trim() -ne '' -and $_ -notmatch ",.[(\W)]$" -and $_ -notmatch ",$" } | 
    Select-Object -Last 1

    # If empty, all latest entries have a seperate field without a value, so we'll just grab the last non-empty line
    if (!$testLog) { 
        $testLog = Get-Content ((Get-ChildItem -Path $confighash.($type + 'LogPath') | 
                Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName) |
        Where-Object { $_.Trim() -ne '' } | Select-Object -Last 1
    }

    $fieldCount = ($testLog.ToCharArray() | Where-Object { $_ -eq ',' } | Measure-Object).Count + 1

    $header = @()
    for ($i = 1; $i -le $fieldCount; $i++) { $header += $i }

    $csv = $testLog | ConvertFrom-Csv -Header $header
    
    if (!($ConfigHash.($Type + 'LogMapping'))) {
        
        $configHash.($type + 'LogMapping') = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
            
        for ($i = 1; $i -le $fieldCount; $i++) {
            $configHash.($type + 'LogMapping').Add([PSCustomObject]@{
                    ID              = $i
                    Field           = $csv.$i
                    FieldSelList    = @('User', 'DateRaw', 'LoginDc', 'ClientName', 'ComputerName', 'Ignore', 'Custom')
                    FieldSel        = $null
                    CustomFieldName = $null
                    Ignore          = $false
                })                                                                
        }
    }

    else {
        $currentMapping = $configHash.($type + 'LogMapping')
        $configHash.($type + 'LogMapping') = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
        
        for ($i = 1; $i -le $fieldCount; $i++) {
            $configHash.($type + 'LogMapping').Add([PSCustomObject]@{
                    ID              = $i
                    Field           = $csv.$i
                    FieldSelList    = @('User', 'DateRaw', 'LoginDc', 'ClientName', 'ComputerName', 'Ignore', 'Custom')
                    FieldSel        = $currentMapping[$i - 1].FieldSel
                    CustomFieldName = $currentMapping[$i - 1].CustomFieldName
                    Ignore          = $false
                })                                                 
        }
    }
      
    $syncHash.($type + 'LogListView').ItemsSource = $configHash.($type + 'LogMapping')
    Set-ChildWindow -SyncHash $syncHash -Title "Map $type Login Logs" -HideCloseButton -Background Flyout
    $syncHash.('settingLogging' + $type + 'Flyout').IsOpen = $true
}

function Set-LoggingDirectory {
    param (
        $SyncHash,
        $ConfigHash,
        [parameter(Mandatory)][ValidateSet('Comp', 'User')]$Type)

    $selectedDirectory = New-FolderSelection -Title "Select client logging directory"

    if (![string]::IsNullOrEmpty($selectedDirectory) -and (Test-Path $selectedDirectory)) {
        $configHash.($type + 'LogPath') = $selectedDirectory              
        $syncHash.($type + 'LogPopupButton').IsEnabled = $true
    }

    else {
        $syncHash.($type + 'LogPopupButton').IsEnabled = $false
    }
}



function Get-LDAPSearchNames {
    param (
    $ConfigHash,
    $SyncHash) 

    (($configHash.queryDefConfig | Where-Object {$_.Name -in $syncHash.searchPropSelection.SelectedItems}).QueryDefType |
        ForEach-Object {$configHash.QueryADValues[([Array]::IndexOf($configHash.QueryADValues.Key, $_))]}).Value

}


#endregion

#region Querying
function Set-ItemExpanders {
    param($SyncHash, $ConfigHash,
        [ValidateSet('Enable', 'Disable')]$IsActive)
    
    if ($IsActive -eq 'Disable') { 
        $syncHash.Window.Dispatcher.Invoke([Action] {                 
                $syncHash.compExpander.IsExpanded = $false
                $syncHash.compExpanderProgressBar.Visibility = "Visible"                    
                $syncHash.userExpander.IsExpanded = $false 
                $syncHash.expanderProgressBar.Visibility = "Visible"
            })
    }
    else {
        $syncHash.Window.Dispatcher.Invoke([Action] {                 
                $syncHash.compExpander.IsExpanded = $true
                $syncHash.compExpanderProgressBar.Visibility = "Collapsed"                    
                $syncHash.userExpander.IsExpanded = $true 
                $syncHash.expanderProgressBar.Visibility = "Collapsed"
            })
    }
}

function Start-ObjectSearch {
    param ($SyncHash, $ConfigHash, $QueryHash, $Key)  
    
    $rsCmd = [PSObject]@{
        key        = $key
        searchTag  = $syncHash.SearchBox.Tag
        searchText = $syncHash.SearchBox.Text
        queue      = $syncHash.snackMsg.MessageQueue
    }

    $rsJob = @{
        Name            = 'Search'
        ArgumentList    = $queryHash, $configHash, $syncHash, $rsCmd
        ModulesToImport = @('C:\TempData\internal\internal.psm1', 'C:\TempData\func\func.psm1')
    }


    Start-RSJob @rsJob -ScriptBlock {
        param($queryHash, $configHash, $syncHash, $rsCmd)        
               
        if ($rsCmd.key -eq 'Escape') {
            $match = (Get-ADObject -Filter "(SamAccountName -eq '$($rsCmd.searchTag)'  -and ObjectClass -eq 'User') -or 
                (Name -eq '$($rsCmd.searchTag)' -and ObjectClass -eq 'Computer')" -Properties SamAccountName) 
        }
        
        else {
            if (!($configHash.searchBaseConfig.OU)) {         
                $match = Get-ADObject -Filter (Get-FilterString -PropertyList $configHash.queryProps -SyncHash $syncHash -Query $rsCmd.searchText) -Properties SamAccountName, Name
            }
            else {
                $match = [System.Collections.ArrayList]@()
                $filter = Get-FilterString -PropertyList $configHash.queryProps -SyncHash $syncHash -Query $rsCmd.searchText
                foreach ($searchBase in ($configHash.searchBaseConfig | Where-Object {$null -ne $_.OU})) {
                    $result = (Get-ADObject -Filter $filter -SearchBase $searchBase.OU -SearchScope $searchBase.QueryScope -Properties SamAccountName, Name) | Where-Object { $_.ObjectClass -match "user|computer"}
                    if ($result) { $result | ForEach-Object {$match.Add($_) | Out-Null} }
                }
            }
        }

        if (($match | Measure-Object).Count -eq 1) { 
                
            Set-ItemExpanders -SyncHash $syncHash -ConfigHash $configHash -IsActive Disable
                        
            if ($match.ObjectClass -eq 'User') {
                $match = (Get-ADUser -Identity $match.SamAccountName -Properties @($configHash.UserPropList.PropName.Where( { $_ -ne 'Non-AD Property' }) | Sort-Object -Unique))                 
                Set-QueryVarsToUpdate -ConfigHash $configHash -Type User
                 Write-LogMessage -Path $configHash.actionlogPath -Message Query -SubjectName $match.SamAccountName -ActionName "Query" -SubjectType "User" -ArrayList $configHash.actionLog
            }

            elseif ($match.ObjectClass -eq 'Computer') {                       
                $match = (Get-ADComputer -Identity $match.SamAccountName -Properties @($configHash.CompPropList.PropName.Where( { $_ -ne 'Non-AD Property' }) | Sort-Object -Unique))
                Set-QueryVarsToUpdate -ConfigHash $configHash -Type Comp
                Write-LogMessage -Path $configHash.actionlogPath -Message Query -SubjectName $match.Name -ActionName "Query" -SubjectType "Computer" -ArrayList $configHash.actionLog
            }
                
            if ($match.SamAccountName -notin $syncHash.tabControl.Items.Name -and $match.Name -notin $syncHash.tabControl.Items.Name) {
                                           
                if ($match.ObjectClass -eq 'User') { Find-ObjectLogs -SyncHash $syncHash -QueryHash $queryHash -ConfigHash $configHash -Type User -Match $match }
                    
                elseif ($match.ObjectClass -eq 'Computer') { Find-ObjectLogs -SyncHash $syncHash -QueryHash $queryHash -ConfigHash $configHash -Type Comp -Match $match }
                                                        
            }

            else {
                if ($match.ObjectClass -eq 'User') {
                    $match.PSObject.Properties | ForEach-Object { $queryHash.($match.SamAccountName)[$_.Name] = $_.Value }
                    $itemIndex = [Array]::IndexOf($syncHash.tabControl.Items.Name, $($match.SamAccountName))     
                    $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.tabControl.SelectedIndex = $itemIndex })
                   
                }
                else {
                    $match.PSObject.Properties | ForEach-Object { $queryHash.($match.Name)[$_.Name] = $_.Value }
                    $itemIndex = [Array]::IndexOf($syncHash.tabControl.Items.Name, $($match.Name))     
                    $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.tabControl.SelectedIndex = $itemIndex })
                    
                }   
            }                    
        }

    

        elseif (($match | Measure-Object).Count -gt 1) {
            $rsCmd.queue.Enqueue("Too many matches!")
            $syncHash.Window.Dispatcher.Invoke([Action] {
                    $syncHash.resultsSidePane.IsOpen = $true
                    $syncHash.resultsSidePaneGrid.ItemsSource = $match | Select-Object Name, SamAccountName, ObjectClass
                })
        }

        else {
            $rsCmd.queue.Enqueue("No match!")
        }
    }
}

function Find-ObjectLogs {
    param (
        $SyncHash, $QueryHash, $ConfigHash, $Match,
        [ValidateSet('User', 'Comp')]$Type)

    if ($type -eq 'User') { $idProp = 'SamAccountName' }
    else { $idProp = 'Name' }

    $queryHash.($match.$idProp) = @{ }
    
    $match.PSObject.Properties | ForEach-Object { $queryHash.($match.$idProp)[$_.Name] = $_.Value }

    if ($configHash.($type + 'LogPath')) { $queryHash.$($match.$idProp).LoginLog = New-Object System.Collections.ObjectModel.ObservableCollection[Object] }

    $addItem = ($match | Select-Object @{Label = 'Name'; Expression = { $_.$idProp } })
    $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.tabControl.ItemsSource.Add($addItem) })  
    $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.tabControl.SelectedIndex = $syncHash.tabControl.Items.Count - 1 })

    $rsLogPull = @{
        Name            = $type + 'LogPull'
        ArgumentList    = $queryHash, $configHash, $match, $syncHash
        ModulesToImport = @('C:\TempData\internal\internal.psm1', 'C:\TempData\func\func.psm1')
    }

    if ($type -eq 'User') {
        Start-RSJob @rsLogPull -ScriptBlock {
            param($queryHash, $configHash, $match, $syncHash) 
                            
            $syncHash.userCompGrid.Dispatcher.Invoke([Action] { $syncHash.userCompGrid.ItemsSource = $null })
            Start-Sleep -Milliseconds 1250
        
            if ($confighash.UserLogPath -and (Test-Path (Join-Path -Path $confighash.UserLogPath -ChildPath "$($match.SamAccountName).txt"))) {
                                        
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
                                        
                                       
                    foreach ($log in ($queryHash.$($match.SamAccountName).LoginLogRaw | Sort-Object -Unique -Property ComputerName | Sort-Object DateTime -Descending)) {                    
                        Remove-Variable sessionInfo, clientLocation, hostLocation -ErrorAction SilentlyContinue
                                              
                        $ruleCount = ($configHash.nameMapList | Measure-Object).Count
                      
                                            
                        $hostConnectivity = Test-OnlineFast -ComputerName $log.ComputerName
                                                                              
                        if ($hostConnectivity.Online) {  $sessionInfo = Get-RDSession -ComputerName $log.ComputerName -UserName $match.SamAccountName -ErrorAction SilentlyContinue}
                        if ($hostConnectivity.IPV4Address) { $hostLocation = (Resolve-Location -computerName $log.ComputerName -IPList $configHash.netMapList -ErrorAction SilentlyContinue).Location}
                        
                        if ($log.ClientName) { $clientOnline = Test-OnlineFast -ComputerName $log.ClientName }      
                        if ($clientOnline.Online) {
                            $clientLocation = (Resolve-Location -computerName $log.ClientName -IPList $configHash.netMapList -ErrorAction SilentlyContinue).Location
                        }
                                            
                                            
                        $queryHash.$($match.SamAccountName).LoginLog.Add(( New-Object PSCustomObject -Property @{
                          
                                    logonTime         = Get-Date($log.DateTime) -Format MM/dd/yyyy
                            
                                    HostName          = $log.ComputerName
                            
                                    LoginDC           = $log.LoginDC -replace '\\'
                            
                                    UserName          = $match.SamAccountName
                            
                                    Connectivity      = ($hostConnectivity.Online).toString()
                            
                                    IPAddress         = $hostConnectivity.IPV4Address
                            
                                    userOnline        = if ($sessionInfo) { $true }
                                                        else { $false }
                            
                                    sessionID         = if ($sessionInfo) { $sessionInfo.sessionID }
                                                        else { $null }
                            
                                    IdleTime          = if ($sessionInfo) {
                                                            if ("{0:dd\:hh\:mm}" -f $($sessionInfo.IdleTime) -eq '00:00:00') { "Active" }
                                                            else { "{0:dd\:hh\:mm}" -f $($sessionInfo.IdleTime) }   
                                                            }
                                                        else { $null }

                                    ClientName        = $log.ClientName 
                            
                                    ClientLocation    = $clientLocation
                            
                                    compLogon         = if (($queryHash.$($match.SamAccountName).LoginLog | Measure-Object).Count -eq 0) { "Last" }
                                                        else { "Past" }
                            
                                    loginCount        = ($loginCounts | Where-Object { $_.Name -eq $log.ComputerName }).Count
                            
                                    DeviceLocation    = $hostLocation
                            
                                    Type              = for ($r = ($ruleCount - 1); $r -ge 0; $r--) {
                                                        if ($r -le 0) { "Computer" }
                                                        else {
                                                            if (($configHash.nameMapList | Sort-Object -Property ID -Descending)[$r].Condition) {
                                                                try {
                                                                    if (Invoke-Expression $configHash.nameMapList[$r].Condition) {
                                                                        $configHash.nameMapList[$r].Name
                                                                        break
                                                                    }
                                                                }
                                                                catch { }
                                                            }
                                                        }
                                                    }                          
                                    ClientOnline      = if ($clientOnline) { ($clientOnline.Online).toString() };
                            
                                    ClientIPAddress   = if ($clientOnline) { $clientOnline.IPV4Address };
                            
                                    ClientType        = for ($r = ($ruleCount - 1); $r -ge 0; $r--) {
                                                        $comp = $_.ClientName
                                                        if ($r -eq 0) { "Computer" }
                                                        else {
                                                            if ($configHash.nameMapList[$r].Condition) {
                                                            try {
                                                                if (Invoke-Expression $configHash.nameMapList[$r].Condition) {
                                                                    $configHash.nameMapList[$r].Name
                                                                    break
                                                                }
                                                            }
                                                            catch { }
                                                        }
                                                    }
                                                }
                        }))
                                                        
                        if ($configHash.userLogMapping.FieldSel -contains 'Custom') {
                            foreach ($customHeader in ($configHash.userLogMapping | Where-Object { $_.FieldSel -eq 'Custom' })) {
                                foreach ($item in ($queryHash.$($match.SamAccountName).LoginLog)) {
                                    $item | Add-Member -Force -MemberType NoteProperty -Name $customHeader.Header -Value $log.($customHeader.Header)
                                }
                            }
                        }

                        $syncHash.Window.Dispatcher.Invoke([Action] { $queryHash.$($match.SamAccountName).LoginLogListView.Refresh() })
                            

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

    else {
        Start-RSJob @rsLogPull -ScriptBlock {
            param($queryHash, $configHash, $match, $syncHash) 
            $syncHash.userCompGrid.Dispatcher.Invoke([Action] { $syncHash.compUserGrid.ItemsSource = $null })
            Start-Sleep -Milliseconds 1250                               

            if ($configHash.compLogPath -and (Test-Path (Join-Path -Path $configHash.compLogPath -ChildPath "$($match.Name).txt"))) {
    
                $queryHash.$($match.Name).LoginLogRaw = Get-Content (Join-Path -Path $confighash.compLogPath -ChildPath "$($match.Name).txt") | 
                    Select-Object -Last 100 | ConvertFrom-Csv -Header $configHash.compLogMapping.Header |
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
                        if ($r -eq 0) { "Computer" }
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

                    $compPing = Test-OnlineFast $match.Name

                    if ($compPing.Online) {  $sessionInfo = Get-RDSession -ComputerName $match.Name -ErrorAction SilentlyContinue }
                    if ($compPing.IPV4Address) {  $hostLocation = (Resolve-Location -computerName $match.Name -IPList $configHash.netMapList -ErrorAction SilentlyContinue).Location }

                    foreach ($log in ($queryHash.$($match.Name).LoginLogRaw | Sort-Object -Unique -Property User | Sort-Object DateTime -Descending)) {
                        Remove-Variable clientLocation -ErrorAction SilentlyContinue

                        if ($log.ClientName) {$clientOnline = Test-OnlineFast -ComputerName $log.ClientName}
    
                        $userSession = $sessionInfo | Where-Object {$_.UserName -eq $log.User}                                                                                                              
        
                        if ($clientOnline.IPV4Address) {
                            $clientLocation = (Resolve-Location -computerName $log.ClientName -IPList $configHash.netMapList -ErrorAction SilentlyContinue).Location
                        }

                        $queryHash.$($match.Name).LoginLog.Add(( New-Object PSCustomObject -Property @{
        
                            logonTime     = Get-Date($log.DateTime) -Format MM/dd/yyyy
        
                            UserName      = ($log.User).ToLower()
        
                            LoginDC       = $log.LoginDC -replace '\\'
        
                            Name          = (Get-ADUser -Identity $log.User).Name
        
                            loginCount    = ($loginCounts | Where-Object { $_.Name -eq $log.User }).Count
        
                            userOnline    = if ($userSession) { $true }
                                            else { $false }
        
                            sessionID     = if ($userSession) { $userSession.SessionId }
                                            else { $null }
        
                            IdleTime      = if ($userSession) {
                                                if ("{0:dd\:hh\:mm}" -f $($userSession.IdleTime) -eq '00:00:00') {  "Active" }
                                                else { "{0:dd\:hh\:mm}" -f $($userSession.IdleTime) }   
                                            }
                                            else {$null}

                            ClientName     = $log.ClientName 

                            Connectivity   = ($compPing.Online).toString()


                            ClientOnline   = if ($clientOnline) { ($clientOnline.Online).toString() };
                            
                            ClientIPAddress = if ($clientOnline.IPV4Address) { $clientOnline.IPV4Address };
        
                            ClientLocation =  $clientLocation

                            DeviceLocation =  $hostLocation

                            Type           = $compType
            
                            CompLogon      = if (($queryHash.$($match.Name).LoginLog | Measure-Object).Count -eq 0) {  "Last" }
                                             else { "Past" }
                                                      
                                                        
                            ClientType     = for ($r = ($ruleCount - 1); $r -ge 0; $r--) {
                                                $comp = $log.ClientName
                                                if ($comp) {
                                                    if ($r -eq 0) { "Computer" }
                                                    else {
                                                        if (($configHash.nameMapList | Sort-Object -Property ID -Descending)[$r].Condition) {
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
                                                }
                        }))
                                        
                        if ($configHash.compLogMapping.FieldSel -contains 'Custom') {
                            foreach ($customHeader in ($configHash.compLogMapping | Where-Object { $_.FieldSel -eq 'Custom' })) {
                                foreach ($item in ($queryHash.$($match.Name).LoginLog)) {
                                    $item | Add-Member -Force -MemberType NoteProperty -Name $customHeader.Header -Value $log.($customHeader.Header)
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

function Set-GridButtons {
    param (
    $SyncHash,
    $ConfigHash,
    [parameter(Mandatory)][ValidateSet('Comp', 'User')]$Type,
    [switch]$SkipSelectionChange)

    if ($type -eq 'User') { 
        $itemName = "userComp" 
        $butCode = 'r'
    }
    else {
        $itemName = "compUser" 
        $butCode = 'rc'
    }
        
    if (!$SkipSelectionChange) {  
        
        if ($null -like $syncHash.($itemName + 'Grid').SelectedItem) { $syncHash.($itemName + 'ControlPanel').IsEnabled = $false }
        else { $syncHash.($itemName + 'ControlPanel').IsEnabled = $true }

        if ([string]::IsNullOrEmpty($syncHash.($itemName + 'Grid').SelectedItem.ClientName)) {          
            $syncHash.($itemName + 'FocusClientToggle').Visibility = "Hidden"
            $syncHash.($type + 'LogClientPropGrid').Visibility = "Hidden"
        }

        else {       
            $syncHash.($itemName + 'FocusClientToggle').Visibility = "Visible"
            $syncHash.($type + 'LogClientPropGrid').Visibility = "Visible"
        }
    }

    if ($type -eq 'User' -and ($syncHash.userCompFocusClientToggle.IsChecked) -and !($SkipSelectionChange)) { $syncHash.userCompFocusHostToggle.IsChecked = $true }
    elseif ($type -eq 'Comp' -and ($syncHash.compUserFocusClientToggle.IsChecked) -and !($SkipSelectionChange)) { $syncHash.compUserFocusUserToggle.IsChecked = $true }
    else {
        foreach ($button in $syncHash.Keys.Where({$_ -like "*butbut*"})) {
            if (($syncHash.($itemName + 'Grid').SelectedItem.Type -in $configHash.rtConfig.($syncHash[$button].Tag).Types) -and
                (!($configHash.rtConfig.($syncHash[$button].Tag).RequireOnline -and $syncHash.($itemName + 'Grid').SelectedItem.Connectivity -eq $false)) -and 
                (!($configHash.rtConfig.($syncHash[$button].Tag).RequireUser -and $syncHash.($itemName + 'Grid').SelectedItem.userOnline -eq $false))) {                
                
                    $syncHash[$button].IsEnabled = $true
            
            }
            else {
                $syncHash[$button].IsEnabled = $false
            }            
        }

        foreach ($button in $synchash.customRT.Keys) {
            if (($syncHash.($itemName + 'Grid').SelectedItem.Type -in $configHash.rtConfig.$button.Types) -and
                (!($configHash.rtConfig.$button.RequireOnline -and $syncHash.($itemName + 'Grid').SelectedItem.Connectivity -eq $false)) -and 
                (!($configHash.rtConfig.$button.RequireUser -and $syncHash.($itemName + 'Grid').SelectedItem.userOnline -eq $false))) {
                $synchash.customRT.$button.($butCode + 'but').IsEnabled = $true
            }
            else {
                $synchash.customRT.$button.($butCode + 'but').IsEnabled = $false
            }            
        }

        foreach ($button in $syncHash.customContext.Keys) {
            if (($syncHash.($itemName + 'Grid').SelectedItem.Type -in $configHash.contextConfig[($button -replace 'cxt') - 1].Types) -and
                (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireOnline -and $syncHash.($itemName + 'Grid').SelectedItem.Connectivity -eq $false)) -and 
                (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireUser -and $syncHash.($itemName + 'Grid').SelectedItem.userOnline -eq $false))) {
                $syncHash.customContext.$button.(($butCode + 'but') + 'context' +  ($button -replace 'cxt')).IsEnabled = $true
            }
            else {
                $syncHash.customContext.$button.(($butCode + 'but') + 'context' +  ($button -replace 'cxt')).IsEnabled = $false
            }            
        }
    } 
}

function Set-ClientGridButtons {
    param (
    $SyncHash,
    $ConfigHash,
    [parameter(Mandatory)][ValidateSet('Comp', 'User')]$Type)

    if ($type -eq 'User') { $itemName = "userComp" }
    else { $itemName = "compUser" }

    foreach ($button in $syncHash.Keys.Where({$_ -like "*rbutbut*"})) {
        if (($syncHash.($itemName + 'Grid').SelectedItem.ClientType -in $configHash.rtConfig.($syncHash[$button].Tag).Types) -and
            (!($configHash.rtConfig.($syncHash[$button].Tag).RequireOnline -and $syncHash.($itemName + 'Grid').SelectedItem.ClientOnline -eq $false)) -and 
            (!($configHash.rtConfig.($syncHash[$button].Tag).RequireUser -eq $true))) {                
                
                $syncHash[$button].IsEnabled = $true
            
        }
        else {
            $syncHash[$button].IsEnabled = $false
        }            
    }

    foreach ($button in $synchash.customRT.Keys) {
        if (($syncHash.($itemName + 'Grid').SelectedItem.ClientType -in $configHash.rtConfig.$button.Types) -and
            (!($configHash.rtConfig.$button.RequireOnline -and $syncHash.($itemName + 'Grid').SelectedItem.ClientOnline -eq $false)) -and 
            (!($configHash.rtConfig.$button.RequireUser -ne $true))) {
            $synchash.customRT.$button.rbut.IsEnabled = $true
        }
        else {
            $synchash.customRT.$button.rbut.IsEnabled = $false
        }            
}

    foreach ($button in $syncHash.customContext.Keys) {
        if (($syncHash.($itemName + 'Grid').SelectedItem.ClientType -in $configHash.contextConfig[($button -replace 'cxt') - 1].Types) -and
            (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireOnline -and $syncHash.($itemName + 'Grid').SelectedItem.ClientOnline -eq $false)) -and 
            (!($configHash.contextConfig[($button -replace 'cxt') - 1].RequireUser  -ne $true))) {
            $syncHash.customContext.$button.('rbutcontext' +  ($button -replace 'cxt')).IsEnabled = $true
        }
        else {
            $syncHash.customContext.$button.('rbutcontext' +  ($button -replace 'cxt')).IsEnabled = $false
        }            
    }
}

function Set-ActionLog {
param ($ConfigHash)
    $configHash.actionLog = New-Object System.Collections.ObjectModel.ObservableCollection[Object] 
    #$configHash.actionLog.Add(([PSCustomObject]@{
    #    Message     = $null
  #      SubjectName = $null
   #     ActionName  = $null
  ##      SubjectType = $null
   #     Admin       = $null
   #     Date        = $null
    #    Time        = $null
      #  DateFull    = $null
   # }))
}
        

function Write-LogMessage {
    param (
        $Path,
        [ValidateSet('Fail', 'Succeed', 'Query')]$Message,
        $ActionName,
        $SubjectName,
        $SubjectType,
        $ArrayList,
        $Error)

    $logMsg = ([PSCustomObject]@{
        ActionName  = $actionName
        Message     = $message
        SubjectName = $SubjectName
        SubjectType = $SubjectType
        Date        = (Get-Date -format d)
        Time        = (Get-Date -format t)
        DateFull    = Get-Date
        Admin       = $env:USERNAME
        Error       = if ($Error) {$error}
                      else {'null'}
    }) 

    $ArrayList.Add($logMsg) | Out-Null 
    if ($Path -and (Test-Path $Path)) { 
        if (!(Test-Path (Join-Path $Path -ChildPath "$($env:USERNAME)"))) {New-Item -ItemType Directory -Path (Join-Path $Path -ChildPath "$($env:USERNAME)") | Out-Null}
        ($logMsg | ConvertTo-CSV -NoTypeInformation)[1] | Out-File -Append -FilePath (Join-Path $path -ChildPath "$($env:USERNAME)\$(Get-Date -format MM.dd.yyyy).log") -Force}

}

function Get-FilterString {
    param ($PropertyList, $Query, $SyncHash)

    if ($syncHash.searchExactToggle.IsChecked) {$compareOp = '-eq'}
    else {$compareOp = '-like'}

    for ($i = 0; $i -lt $PropertyList.Count; $i++) {
        
        $searchString = $searchString + "$($PropertyList[$i]) $compareOp `"*$query*`""

        if ($i -lt ($PropertyList.Count -1)) {
            $searchString = $searchString + ' -or '
        }
    }

    $searchString
}

#endregion

#endwindow