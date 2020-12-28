# saves the config import json to synchash or export synchash to json

#region initial setup

# generated hash tables used throughout tool
function New-HashTables {

    # Stores values logging missing or errored items during init
    $global:sysCheckHash = [hashtable]::Synchronized(@{})

    # Stores config values imported JSON, during config, or both
    $global:configHash = [hashtable]::Synchronized(@{})

    # Stores WPF controls
    $global:syncHash = [hashtable]::Synchronized(@{})

    # Stores data related to queried objects
    $global:queryHash = [hashtable]::Synchronized(@{})

}

function Set-WPFControls {
    param (
        [Parameter(Mandatory)]$XAMLPath,
        [Parameter(Mandatory)][Hashtable]$SyncHash
    ) 

    $inputXML = Get-Content -Path $xamlPath
    
    $inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace ' x:Class="V3.Build.MainWindow"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = $inputXML

    $xmlReader = (New-Object System.Xml.XmlNodeReader $xaml)
    
    try { $SyncHash.Window = [Windows.Markup.XamlReader]::Load($xmlReader) }
    catch { Write-Warning "Unable to parse XML, with error: $($Error[0])" }

    ## Load each named control into PS hashtable
    foreach ($controlName in ($xaml.SelectNodes("//*[@Name]").Name)) {
        $syncHash.$controlName = $syncHash.Window.FindName($controlName) 
    }

    
    $syncHash.windowContent.Visibility = "Hidden"
    $syncHash.Window.Height = 500
    $syncHash.Window.ResizeMode = "NoResize"
    $syncHash.Window.ShowTitleBar = $false
    $syncHash.Window.ShowCloseButton = $false
    $syncHash.Window.Width = 500
    $syncHash.splashLoad.Visibility = "Visible" 
    
}

function Set-Config {
    Param ( 
        [CmdletBinding()]
        [Parameter(Mandatory=$true)]
        [String]
        $ConfigPath,

        [Parameter(Mandatory=$true)]
        [ValidateSet('Export','Import')]
        [string[]]
        $Type,

        [Parameter(Mandatory=$true)]
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
            $configHash | ConvertTo-Json -Depth 8 | Out-File $($configPath + '.bak') -Force

            if ((Get-ChildItem -LiteralPath $($configPath + '.bak')).Length -gt 0) {
                Copy-Item -LiteralPath $($configPath + '.bak') -Destination $configPath 
            }
        }
    }
}

function Suspend-FailedItems {
param (
    $SyncHash,
    [Parameter(Mandatory=$true)][ValidateSet('Config','SysCheck','RSCheck')][string[]]$CheckedItems
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
            $syncHash.settingToolClick.IsEnabled = $false
            $syncHash.settingADClick.IsEnabled = $false
            $syncHash.settingPermClick.IsEnabled = $false
            $syncHash.settingFailPanel.Visibility = "Visible"  
            $syncHash.settingConfigSeperator.Visibility = 'Hidden'
            $syncHash.settingsConfigItems.Visibility = 'Hidden' 
            $syncHash.settingStatusChildBoard.Visibility = 'Visible'
             
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

    $syncHash.tabMenu.Items | ForEach-Object {$syncHash.Window.Dispatcher.invoke([action]{$_.IsEnabled = $false})}
    
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
        [Parameter(Mandatory,ValueFromPipeline)]$Type,
        [Parameter(Mandatory)][Hashtable]$ConfigHash,
        [switch]$PullDefaults     
    )

    Process {
       
        # check if values already exist
            if ($configHash.$type) { $tempList = $configHash.$type }
            # pull from base templates if not
            elseif ($PullDefaults) { $tempList = Get-InitialValues -GroupName $type }
        

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
        [Parameter(Mandatory,ValueFromPipeline)]$Type,
        [Parameter(Mandatory)][Hashtable]$ConfigHash,
        [Array]$DefaultList      
    )

    Process {

        if ($configHash.$type) {

            $duplicate = 0
            $configHash.$type | Add-Member -MemberType NoteProperty -Name 'Header' -Value $null -ErrorAction SilentlyContinue

            $configHash.$type | ForEach-Object {
                $_.FieldSelList = $BuiltInList
        
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
   
function Set-RTDefaults {
  [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory,ValueFromPipeline)]$Type,
        [Parameter(Mandatory)][Hashtable]$ConfigHash
    )

    Begin {

        if ([string]::IsNullOrEmpty($configHash.rtConfig)) {
            $configHash.rtConfig = @{}
        }
        
        else {
            $rtTemp = $configHash.rtConfig
            $configHash.rtConfig = @{}
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

function Add-CustomRTControls {
 [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory)][Hashtable]$SyncHash,
        [Parameter(Mandatory)][Hashtable]$ConfigHash
    )


    $syncHash.customRT = @{}
    foreach ($rtID in $configHash.rtConfig.Keys.Where{$_ -like "RT*"}) {
       
        $syncHash.customRt.$rtID = @{
        
            parentDock              = New-Object System.Windows.Controls.DockPanel -Property @{
                Margin              = "0,0,15,0" 
                HorizontalAlignment = "Stretch" 
            }
        
            childStack              = New-Object System.Windows.Controls.StackPanel
        
            InfoHeader              = New-Object System.Windows.Controls.Label -Property @{
                Content             = "Custom Tool $($rtID -replace 'rt')"
                Style               = $syncHash.Window.FindResource('rtHeader')
            }
        
            InfoSubheader           = New-Object System.Windows.Controls.TextBlock -Property @{
                Text                = $configHash.rtConfig.rt1.DisplayName
                Style               = $syncHash.Window.FindResource('rtSubHeader')
            }
        
            AlertGlyph              = New-Object System.Windows.Controls.Label -Property @{
                Style               = $syncHash.Window.FindResource('rtLabel')
            }
        
            ConfigureButton         = New-Object System.Windows.Controls.Button -Property  @{
                Style               = $syncHash.Window.FindResource('rtClick')
            }
        
            DelButton               = New-Object System.Windows.Controls.Button -Property  @{
                Style               = $syncHash.Window.FindResource('rtClickDel')
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

        $syncHash.customRt.$rtID.DelButton.Add_Click({
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
    }
}

function Add-CustomItemBoxControls {
 [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory)][Hashtable]$SyncHash,
        [Parameter(Mandatory)][Hashtable]$ConfigHash
    )

    $syncHash.customContext = @{}

    foreach ($type in @('ubox', 'cbox')) {
    
        switch ($type) {
            'ubox' {$typeName = 'user'}
            'cbox' {$typeName = 'comp'}
        }

        for ($i = 1; $i -le $configHash.boxMax; $i++) {

            if ($i -le $configHash.($typeName + 'boxCount')) {
            
                $syncHash.(($type + $i + 'resources')) = @{

                    ($type + $i + 'Border')      = New-Object System.Windows.Controls.Border -Property @{
                        Style                    = $syncHash.Window.FindResource('itemBorder')
                        Name                     = ($type + $i + 'Border')
                    }

                    ($type + $i + 'Grid')        = New-Object System.Windows.Controls.Grid
                    ($type + $i + 'DockPanel')   = New-Object System.Windows.Controls.DockPanel
                    ($type + $i + 'StackPanel')  = New-Object System.Windows.Controls.StackPanel -Property @{Style = $syncHash.Window.FindResource('itemStackPanel') }

                    ($type + $i + 'Header')      = New-Object System.Windows.Controls.Label -Property @{
                        Style                    = $syncHash.Window.FindResource('itemBoxHeader')
                        Name                     = ($type + $i + 'Header')           
                    }

                    ($type + $i + 'EditClip')    = New-Object System.Windows.Controls.Label -Property @{
                        Style                    = $syncHash.Window.FindResource('itemEditClip')
                        Name                     = ($type + $i + 'EditClip')
                    }

                    ($type + $i + 'ViewBox')     = New-Object System.Windows.Controls.ViewBox -Property @{Style = $syncHash.Window.FindResource('itemViewBox') }

                    ($type + $i + 'TextBox')     = New-Object System.Windows.Controls.TextBox -Property @{
                        Style                    = $syncHash.Window.FindResource('itemBox')
                        Name                     = ($type + $i + 'TextBox')
                    }

                    ($type + $i + 'Box1Action1') = New-Object System.Windows.Controls.Button -Property @{
                        Style                    = $syncHash.Window.FindResource('itemButton')
                        Name                     = ($type + $i + 'Box1Action1')
                    }

                    ($type + $i + 'Box1Action2') = New-Object System.Windows.Controls.Button -Property @{
                        Style                    = $syncHash.Window.FindResource('itemButton')
                        Name                     = ($type + $i + 'Box1Action2')
                    }
            
                    ($type + $i + 'Box1')        = New-Object System.Windows.Controls.Button -Property @{
                        Style                    = $syncHash.Window.FindResource('itemEditButton')
                        Name                     = ($type + $i + 'Box1')
                    }

                }
        
                # Add col def objects, then add to outside grid
                $colDef1                         = New-Object System.Windows.Controls.ColumnDefinition
                $colDef1.Width                   = "*"

                $colDef2                         = New-Object System.Windows.Controls.ColumnDefinition
                $colDef2.Width                   = "Auto"
                $colDef2.MaxWidth                = "90"   

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

    $syncHash.objectTools = @{}
    #create custom tool buttons for queried objects
    foreach ($tool in $configHash.objectToolConfig) {
        if ($tool.toolActionValid) {
            switch ($tool.objectType) {
                {$_ -match "Both|Comp"} { 
        
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
                {$_ -match "Both|User"} { 
        
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
            }

            foreach ($toolButton in (($syncHash.objectTools.Keys).Where{($_ -replace ".tool") -eq $tool.ToolID})) {

                $syncHash.objectTools.$toolButton.ToolButton.Add_Click({
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
                                Start-RSJob -Name ItemTool -ArgumentList $syncHash.snackMsg.MessageQueue, $toolID, $configHash, $queryHash -ScriptBlock {
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
                            
                                                      
                                    $syncHash.Window.Dispatcher.Invoke([Action]{                                    
                                        $syncHash.itemTooListBoxProgress.Visibility = "Visible"
                                    })
                    
                                    $list = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
                                    Invoke-Expression ($configHash.objectToolConfig[$toolId - 1].toolFetchCmd) | ForEach-Object {$list.Add([PSCUstomObject]@{'Name' = $_})}
                                 
                                    $syncHash.Window.Dispatcher.Invoke([Action]{
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

                            
                                                      
                                    $syncHash.Window.Dispatcher.Invoke([Action]{                                  
                                        $syncHash.itemToolGridProgress.Visibility = "Visible"
                                    })

                                
                                    $list = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
                                    Invoke-Expression ($configHash.objectToolConfig[$toolId - 1].toolFetchCmd) | ForEach-Object {$list.Add($_)}
                                 
                                    $syncHash.Window.Dispatcher.Invoke([Action]{
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
param ($SysCheckHash) 

    if ((Get-WmiObject -Class Win32_ComputerSystem).PartofDomain) {               
        $sysCheckHash.sysChecks[0].ADMember = 'True'
                    
        if (($sysCheckHash.sysChecks[0].ADModule -eq $true) -and (Test-Connection -Quiet -Count 1 -ComputerName ($env:logonServer -replace '\\', ''))) {                                                 
            $selectedDC = Get-ADDomainController -Discover -Service ADWS -ErrorAction SilentlyContinue 

            if (Test-Connection -Count 1 -Quiet -ComputerName $selectedDC.HostName) {             
                $sysCheckHash.sysChecks[0].ADDCConnectivity = 'True'
                            
                try {
                    $adEntity = [Microsoft.ActiveDirectory.Management.ADEntity].Assembly
                    $adFields = $adEntity.GetType('Microsoft.ActiveDirectory.Management.Commands.LdapAttributes').GetFields('Static,NonPublic') | Where-Object { $_.IsLiteral }
                    $configHash.adPropertyMap = @{}
                            
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
        [Parameter(Mandatory,ValueFromPipeline)]$ControlName,
        [Parameter(Mandatory)][Hashtable]$SyncHash,
        [Parameter(Mandatory)]$DataContext   
    )

    Process {
        $syncHash.Window.Dispatcher.invoke([action]{$syncHash.$ControlName.DataContext = $DataContext})
    }
}

function Get-PropertyLists {
param ($ConfigHash) 
 
    foreach ($type in ("user","comp")) {
        
        if ($type -eq 'user') {
            $configHash.userPropPullList = (Get-ADUser -Identity $env:USERNAME -Properties *).PSObject.Properties | Select-Object Name, TypeNameofValue
        }
        else {
            $configHash.compPropPullList = (Get-ADComputer -Identity $env:COMPUTERNAME -Properties *).PSObject.Properties | Select-Object Name, TypeNameofValue
        }        
                
        $configHash.($type + 'PropPullListNames') = [System.Collections.ArrayList]@()
        $configHash.($type + 'PropPullListNames').Add("Non-AD Property") 
        $configHash.($type + 'PropPullList').Name | ForEach-Object {$configHash.($type + 'PropPullListNames').Add($_)}
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
        [ValidateSet('Standard','Flyout')][string[]]$Background
    )

    if ($Panel) {
        $syncHash.$Panel.Visibility = "Visible"
        $syncHash.settingChildWindow.IsOpen = $true
    }

    if ($Title)  {$syncHash.settingChildWindow.Title = $Title}
    if ($Height) {$syncHash.settingChildHeight.Height = $Height}
    if ($Width)  {$syncHash.settingChildHeight.Width = $Width}

    switch ($Background) {       
        'Standard' {$syncHash.settingChildWindow.TitleBarBackground = ($syncHash.Window.BorderBrush.Color).ToString(); break}
        'Flyout'  {$syncHash.settingChildWindow.TitleBarBackground = ($syncHash.settingNameFlyout.Background.Color).ToString()}
    }

    if ($HideCloseButton) {$SyncHash.settingChildWindow.ShowCloseButton = $false}
    else {$SyncHash.settingChildWindow.ShowCloseButton = $true}

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
        foreach ($contentPane in ($syncHash.Keys.Where({$_ -like "setting*Content"}))) {
            $syncHash.$contentPane.Visibility = "Hidden"
        }
    }

    if (!($SkipDataGridReset)) {
        foreach ($dataGrid in ($syncHash.Keys.Where({$_ -like "setting*Grid"}))) {
            $syncHash.$dataGrid.Visibility = "Hidden"
        }
    }

    if (!($SkipFlyoutClose)) {
        foreach ($flyOut in ($syncHash.Keys.Where({$_ -like "setting*Flyout"}))) {
            $syncHash.$flyOut.IsOpen = $false
        }
    }

    $syncHash.userPropGrid.HorizontalAlignment = "Left"
    $syncHash.compPropGrid.HorizontalAlignment = "Left"

    $syncHash.settingUserAddItemClick.Visibility = "Hidden"
    $syncHash.settingCompAddItemClick.Visibility = "Hidden"

    if ($Title) { Set-ChildWindow -SyncHash $syncHash -Title $title }

    if (!($SkipResize)) { Set-ChildWindow -SyncHash $syncHash -Width 400 -Height 215 }
    
    Set-ChildWindow -SyncHash $syncHash -Background Standard

    if ($syncHash.settingStatusChildBoard.Visibility -eq 'Collapsed') {
        $syncHash.settingStatusChildBoard.Visibility = 'Visible'
        $syncHash.settingConfigPanel.Visibility = 'Visible'
    }

}



#endwindow