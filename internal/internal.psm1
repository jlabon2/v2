# saves the config import json to synchash or export synchash to json

#region initial setup

# generated hash tables used throughout tool
function New-HashTables {
    # Stores values log0ging missing or errored items during init
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
function Set-CurrentPane {
    param ($SyncHash, $Panel)
    $SyncHash.infoPaneContent.Tag = $Panel
}


function Get-RelatedClass {
    param( [string]$ClassName )
  
    $Classes = @($ClassName)
  
    $SubClass = Get-ADObject -SearchBase "$((Get-ADRootDSE).SchemaNamingContext)" -Filter { lDAPDisplayName -eq $ClassName } -Properties subClassOf | Select-Object -ExpandProperty subClassOf
    if ( $SubClass -and $SubClass -ne $ClassName ) { $Classes += Get-RelatedClass $SubClass }
  
    $auxiliaryClasses = Get-ADObject -SearchBase "$((Get-ADRootDSE).SchemaNamingContext)" -Filter { lDAPDisplayName -eq $ClassName } -Properties auxiliaryClass | Select-Object -ExpandProperty auxiliaryClass
    foreach ( $auxiliaryClass in $auxiliaryClasses ) { $Classes += Get-RelatedClass $auxiliaryClass }

    $systemAuxiliaryClasses = Get-ADObject -SearchBase "$((Get-ADRootDSE).SchemaNamingContext)" -Filter { lDAPDisplayName -eq $ClassName } -Properties systemAuxiliaryClass | Select-Object -ExpandProperty systemAuxiliaryClass
    foreach ( $systemAuxiliaryClass in $systemAuxiliaryClasses ) { $Classes += Get-RelatedClass $systemAuxiliaryClass }
    Return $Classes  
}

function Get-AllUserAttributes {
    $ADUser = Get-ADUser -ResultSetSize 1 -Filter * -Properties objectClass
    $AllClasses = ( Get-RelatedClass $ADUser.ObjectClass | Sort-Object -Unique )

    $AllAttributes = @()
    Foreach ( $Class in $AllClasses ) {
        $attributeTypes = 'MayContain', 'MustContain', 'systemMayContain', 'systemMustContain'
        $ClassInfo = Get-ADObject -SearchBase "$((Get-ADRootDSE).SchemaNamingContext)" -Filter { lDAPDisplayName -eq $Class } -Properties $attributeTypes 
        ForEach ($attribute in $attributeTypes) { $AllAttributes += $ClassInfo.$attribute }
    }
    $AllAttributes | Sort-Object -Unique
}

function Get-AllCompAttributes {
    $ADUser = Get-ADComputer -ResultSetSize 1 -Filter * -Properties objectClass
    $AllClasses = ( Get-RelatedClass $ADUser.ObjectClass | Sort-Object -Unique )

    $AllAttributes = @()
    Foreach ( $Class in $AllClasses ) {
        $attributeTypes = 'MayContain', 'MustContain', 'systemMayContain', 'systemMustContain'
        $ClassInfo = Get-ADObject -SearchBase "$((Get-ADRootDSE).SchemaNamingContext)" -Filter { lDAPDisplayName -eq $Class } -Properties $attributeTypes 
        ForEach ($attribute in $attributeTypes) { $AllAttributes += $ClassInfo.$attribute }
    }
    $AllAttributes | Sort-Object -Unique
}

function Create-AttributeList {
    param ([Parameter(Mandatory)][ValidateSet('User', 'Computer')]$Type, $ConfigHash)

    $masterList = @()
    if ($type -eq 'User')  { 
        $masterList += Get-AllUserAttributes
        $masterList += (Get-AdUser -Identity $env:USERNAME -Properties *).PsObject.Properties.Name
    }
    else { 
        $masterList += Get-AllCompAttributes 
        $masterList += (Get-ADComputer -Identity $env:COMPUTERNAME -Properties *).PsObject.Properties.Name    
    }

    $attributeList = [System.Collections.ArrayList]@()

    foreach ($object in ($masterList | Sort-Object -Unique )) {
        if ($type -eq 'User') { 
           #convert to friendly name from ldap name - if possible
           if  ($configHash.adPropertyMapInverse[$object]) { $object = $configHash.adPropertyMapInverse[$object]}
                $attributeList.Add(((Get-ADUser -ResultSetSize 1 -filter * -Properties $object -ErrorAction SilentlyContinue | Select-Object $object).PSObject.Properties | Select-Object -Property Name, TypeNameofValue)) | Out-Null}
        else  { 
           if  ($configHash.adPropertyMapInverse[$object]) { $object = $configHash.adPropertyMapInverse[$object]}
            $attributeList.Add(((Get-ADComputer -ResultSetSize 1 -filter * -Properties $object -ErrorAction SilentlyContinue | Select-Object $object).PSObject.Properties | Select-Object -Property Name, TypeNameofValue)) | Out-Null}
    }

    $attributeList
}

function Reset-ScriptBlockValidityStatus {
    param ($SyncHash, $ResultBoxName, $ItemSet, $StatusName)

    $syncHash.($ResultBoxName + 'ErrorStatus').Tag = $null
    $syncHash.$ResultBoxName.Tag = 'Unchecked'
    $itemSet.$StatusName = $null
}


function Update-ScriptBlockValidityStatus {
    param ($syncHash, $ResultBoxName, $itemSet, $statusName)
    
    $syncHash.($ResultBoxName + 'ErrorStatus').Tag = $null
  
    switch ($itemSet.$StatusName) {
        $null {  $syncHash.$ResultBoxName.Tag = 'Unchecked'; break }
        $true {  $syncHash.$ResultBoxName.Tag = 'Pass'; break }
        $false {  $syncHash.$ResultBoxName.Tag = 'Fail'; break }
        
    }
}


function Test-UserScriptBlock {
    param ($errorControl, $scriptBlock, $syncHash, $StatusName, $itemSet)
    
    $scriptDef = @{
        ScriptDefinition = $scriptBlock
        Settings         = @('CodeFormatting', 'CmdletDesign', 'ScriptFunctions', 'ScriptingStyle', 'ScriptSecurity')
        ExcludeRule      = 'PSUseDeclaredVarsMoreThanAssignments'
    }
  
    $errorControlBase = $syncHash.$errorControl
    $errorControlTool = $syncHash.($errorControl + 'ErrorStatus')
    
    $results = Invoke-ScriptAnalyzer @scriptDef     
    $toolTipBox = New-Object -TypeName System.Windows.Controls.TextBlock
        
    # parse errors - invalid block
    if ($results.Severity -contains 'ParseError') {
        $errorControlTool.Tag = 'Fail'
        $errorControlBase.Tag = 'Fail'        
        
        $itemSet.$StatusName = $false
       # $errorControlBase.Text = "Scriptblock is invalid ($((($results | Where-Object {$_.Severity -eq 'ParseError'}) | Measure-Object).Count) major errors)"              
    }
     
    # non-terminating errors - valid block   
    elseif ($results) {
        $errorControlTool.Tag = 'PassWithIssues'
        $errorControlBase.Tag = 'PassWithIssues'
        
        $itemSet.$StatusName = $true
    #    $errorControlBase.Text = "Scriptblock is valid but has warnings ($(($results | Measure-Object).Count) errors)"       
    }
    
    # no issues - valid block    
    else { 
        $errorControlTool.Tag = 'Pass'
        $errorControlBase.Tag = 'Pass'

        $itemSet.$StatusName = $true       
     #   $errorControlBase.Text = "Scriptblock is valid and has no warnings"
    }
           
    # build issue list
    if ($results) {
        $toolTipBox.AddChild((New-Object -TypeName System.Windows.Documents.Run -Property @{
                    Text  = "Scriptblock invalid ($($results.Count) errors)"
                    Style = $syncHash.Window.FindResource('scriptBlockToolTipSeverity')              
                }))
    
        $toolTipBox.AddChild((New-Object -TypeName System.Windows.Documents.LineBreak))
          
    
        foreach ($result in $results) {
            
            $listArray = [System.Collections.ArrayList]@()
                
            $listArray.Add((New-Object -TypeName System.Windows.Documents.LineBreak)) | Out-Null
            
            $listArray.Add((New-Object -TypeName System.Windows.Documents.Run -Property @{
                        Text  = $result.Severity
                        Style = $syncHash.Window.FindResource('scriptBlockToolTipSeverity')              
                    })) | Out-Null
    
            $listArray.Add(( New-Object -TypeName System.Windows.Documents.Run -Property @{
                        Text  = " - Line $($result.Line)"  
                        Style = $syncHash.Window.FindResource('scriptBlockToolTipBase')                    
                    })) | Out-Null
            $listArray.Add((New-Object -TypeName System.Windows.Documents.LineBreak))
    
            $listArray.Add(( New-Object -TypeName System.Windows.Documents.Run -Property @{
                        Text  = "$($result.Message)"        
                        Style = $syncHash.Window.FindResource('scriptBlockToolTipBase')           
                    })) | Out-Null
    
            $listArray.Add((New-Object -TypeName System.Windows.Documents.LineBreak)) | Out-Null
              
            $listArray | ForEach-Object { $toolTipBox.AddChild($_) }
        }
    
        $errorControlTool.ToolTip.AddChild($toolTipBox)
    }

    else {
        $toolTipBox.AddChild((New-Object -TypeName System.Windows.Documents.Run -Property @{
                    Text  = "No issues detected on scriptblock"
                    Style = $syncHash.Window.FindResource('scriptBlockToolTipSeverity')              
                }))
    }
    
    $errorControlTool.ToolTip = $toolTipBox
    
}


function Set-InfoPaneContent {
    param ($SyncHash , $SettingInfoHash, $ConfigHash) 
    
    $SyncHash.infoPaneContent.DataContext = $settingInfoHash.($SyncHash.infoPaneContent.Tag)
    $typeArray = @()

        $syncHash.Keys | Where-Object {$_ -like "infoPane*Panel"} | ForEach-Object {$syncHash.$_.Visibility = 'Collapsed'}

    if ( $settingInfoHash.($SyncHash.infoPaneContent.Tag).Vars) {            
        $typeArray += 'vars'
        $SyncHash.infoPaneVarsPanel.Visibility = 'Visible'

        if ($configHash.varListConfig.VarCmd) {
            $typeArray += 'customvars'
            $SyncHash.infoPaneCustomVarsPanel.Visibility = 'Visible'
        }

    }

    if ($settingInfoHash.($SyncHash.infoPaneContent.Tag).Tips) {            
        $typeArray += 'tips'
        $SyncHash.infoPaneTipsPanel.Visibility = 'Visible'
    }

    if ($settingInfoHash.($SyncHash.infoPaneContent.Tag).Types) { 
        $typeArray += 'types'
        $SyncHash.infoPaneTypesPanel.Visibility = 'Visible'
    }

        
 
        

    foreach ($type in $typeArray) {
        $SyncHash.('infoPane' + $type + 'List').Blocks.Clear()
        $flowPara = New-Object -TypeName System.Windows.Documents.Paragraph

        if ($type -ne 'customvars') {

            foreach ($item in (($settingInfoHash.($SyncHash.infoPaneContent.Tag).$type).Keys)) {

                if ($item -ne 'customvars') {
                    $bulletHeader = New-Object -TypeName System.Windows.Documents.Run -Property @{
                        Text  = $item
                        Style = $SyncHash.Window.FindResource('varNameRun')
                    }

                    $bulletContent = New-Object -TypeName System.Windows.Documents.Run -Property @{
                        Text  = " - $(($settingInfoHash.($SyncHash.infoPaneContent.Tag).$type)[$item])"
                        Style = $SyncHash.Window.FindResource('varDescRun')
                    }
                }

                $bulletEnd = New-Object -TypeName System.Windows.Documents.LineBreak
                    
                $flowPara.AddChild($bulletHeader)
                $flowPara.AddChild($bulletContent)
                $flowPara.AddChild($bulletEnd)
            }

        }

        else {
            
            foreach ($item in ($configHash.varListConfig | Where-Object {$null -ne $_.VarName -and $_.null -ne $_.VarCmd -and $null -ne $_.VarDesc})) {
                $bulletHeader = New-Object -TypeName System.Windows.Documents.Run -Property @{
                    Text  = $item.VarName
                    Style = $SyncHash.Window.FindResource('varNameRun')
                }

                $bulletContent = New-Object -TypeName System.Windows.Documents.Run -Property @{
                    Text  = " - $($item.VarDesc)"
                    Style = $SyncHash.Window.FindResource('varDescRun')
                }
            

                $bulletEnd = New-Object -TypeName System.Windows.Documents.LineBreak
                    
                $flowPara.AddChild($bulletHeader)
                $flowPara.AddChild($bulletContent)
                $flowPara.AddChild($bulletEnd)

            }
        }


        $SyncHash.('infoPane' + $type + 'List').AddChild($flowPara)
    }
}

function Set-CustomVariables {
    param ($VarHash) 

    foreach ($var in $varHash.Keys) { 
        Set-Variable -Name $var -Value $varHash.$var -Scope global
    }

}

function Get-Glyphs {
    param (
        $ConfigHash,
        $GlyphList)

    $glyphs = Get-Content $GlyphList
    $ConfigHash.buttonGlyphs = [System.Collections.ArrayList]@()
    $glyphs | ForEach-Object -Process { $null = $ConfigHash.buttonGlyphs.Add($_) }
}

function Set-WPFControls {
    param (
        [Parameter(Mandatory)]$XAMLPath,
        [Parameter(Mandatory)][Hashtable]$TargetHash
    ) 

    $inputXML = Get-Content -Path $XAMLPath
    
    $inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace ' x:Class="v3.Window1"' -replace ' x:Class="V3.Build.MainWindow"', '' -replace 'x:N', 'N' -replace '^<Win.*', '<Window'
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = $inputXML

    $xmlReader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $XAML)
    
    try { $TargetHash.Window = [Windows.Markup.XamlReader]::Load($xmlReader) }
    catch { Write-Warning -Message "Unable to parse XML, with error: $($Error[0])" }

    ## Load each named control into PS hashtable
    foreach ($controlName in ($XAML.SelectNodes('//*[@Name]').Name)) { $TargetHash.$controlName = $TargetHash.Window.FindName($controlName) }

    $SyncHash.windowContent.Visibility = 'Hidden'
    $SyncHash.Window.Height = 500
    $SyncHash.Window.ResizeMode = 'NoResize'
    $SyncHash.Window.ShowTitleBar = $false
    $SyncHash.Window.ShowCloseButton = $false
    $SyncHash.Window.Width = 500
    $SyncHash.splashLoad.Visibility = 'Visible'
}

function Show-WPFWindow {
    param($SyncHash) 
    $SyncHash.Window.Dispatcher.invoke([action] {                       
            $SyncHash.windowContent.Visibility = 'Visible'
            $SyncHash.Window.MinWidth = '1000'
            $SyncHash.Window.MinHeight = '700'
            $SyncHash.Window.ResizeMode = 'CanResizeWithGrip'
            $SyncHash.Window.ShowTitleBar = $true
            $SyncHash.Window.ShowCloseButton = $true                   
            $SyncHash.splashLoad.Visibility = 'Collapsed' 
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
        $type,

        [Parameter(Mandatory = $true)]
        [Hashtable]
        $ConfigHash
    )

    switch ($type) {
        'Import' {
            if (Test-Path $ConfigPath) {
                if ((Get-ChildItem -LiteralPath $ConfigPath).Length -eq 0 -and (Get-ChildItem -LiteralPath $($savedConfig + '.bak')).Length -gt 0) { Copy-Item -LiteralPath $ConfigPath -Destination $($ConfigPath + '.bak') }

                (Get-Content $ConfigPath | ConvertFrom-Json).PSObject.Properties | ForEach-Object -Process { $ConfigHash[$_.Name] = $_.Value }
            }
        }
        'Export' {
            @('User', 'Comp') | ForEach-Object -Process {
                $ConfigHash.($_ + 'PropList') | ForEach-Object { $_.PropList = $null }
                $ConfigHash.($_ + 'PropListSelection') = $null
                $ConfigHash.($_ + 'PropPullListNames') = $null
                $ConfigHash.($_ + 'PropPullList') = $null
            }

            $ConfigHash.queryDefConfig.ID | ForEach-Object { $ConfigHash.queryDefConfig[$_ - 1].QueryDefTypeList = $null }

            $ConfigHash.buttonGlyphs = $null
            $ConfigHash.adPropertyMap = $null
            $ConfigHash.queryProps = $null
            $ConfigHash.actionLog = $null
            $ConfigHash.modList = $null
            $configHash.adPropertyMapInverse = $null
            $configHash.QueryADValues = $null
            $configHash.rawADValues = $null
            $configHash.currentTabItem = $null

            $ConfigHash |
                ConvertTo-Json -Depth 8 |
                    Out-File -FilePath $($ConfigPath + '.bak') -Force

            if ((Get-ChildItem -LiteralPath $($ConfigPath + '.bak')).Length -gt 0) { Copy-Item -LiteralPath $($ConfigPath + '.bak') -Destination $ConfigPath }
        }
    }
}

function Import-Config {
    param ($SyncHash)

    $configSelection = New-Object -TypeName System.Windows.Forms.OpenFileDialog -Property @{
        InitialDirectory = [Environment]::GetFolderPath('MyComputer')
        Filter           = 'config|config.json'
        Title            = 'Select config.json'
    }
    
    $null = $configSelection.ShowDialog()

    if (![string]::IsNullOrEmpty($configSelection.fileName)) {
        Copy-Item -Path $configSelection.fileName -Destination $PSScriptRoot -Force
        $SyncHash.Window.Close()
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

        'RSCheck' {
            $SyncHash.splashLoad.Visibility = 'Collapsed'
            $SyncHash.windowContent.Visibility = 'Visible'
            $SyncHash.Window.MinWidth = '1000'
            $SyncHash.Window.MinHeight = '700'
            $SyncHash.Window.ResizeMode = 'CanResizeWithGrip'
            $SyncHash.Window.ShowTitleBar = $true
            $SyncHash.Window.ShowCloseButton = $true   
            $SyncHash.windowContent.Visibility = 'Visible'
            $SyncHash.settingADClick.IsEnabled = $false
            $SyncHash.settingPermClick.IsEnabled = $false
            $SyncHash.settingFailPanel.Visibility = 'Visible'  
            $SyncHash.settingConfigSeperator.Visibility = 'Hidden'
            $SyncHash.settingsConfigItems.Visibility = 'Hidden' 
            $SyncHash.settingStatusChildBoard.Visibility = 'Visible'
            $SyncHash.settingModADLabel.Visibility = 'Collapsed'
            $SyncHash.settingADLabel.Visibility = 'Collapsed'
            $SyncHash.settingPermLabel.Visibility = 'Collapsed'
             
            break                       
        }

        'SysCheck' {
            $SyncHash.Window.Dispatcher.invoke([action] {                        
                    $SyncHash.settingStatusChildBoard.Visibility = 'Visible'
                    $SyncHash.settingFailPanel.Visibility = 'Visible'
                    $SyncHash.settingConfigSeperator.Visibility = 'Hidden'
                    $SyncHash.settingsConfigItems.Visibility = 'Hidden'
                })

            break
        }  
        
        'Config' {
            $SyncHash.Window.Dispatcher.invoke([action] {            
                    $SyncHash.settingStatusChildBoard.Visibility = 'Visible'
                    $SyncHash.settingConfigPanel.Visibility = 'Visible'
                    $SyncHash.settingConfigMissing.Visibility = 'Visible'
                })          
        } 
    }

    $SyncHash.tabMenu.Items | ForEach-Object -Process { $SyncHash.Window.Dispatcher.invoke([action] { $_.IsEnabled = $false }) }
    
    $SyncHash.Window.Dispatcher.invoke([action] { 
            $SyncHash.tabMenu.Items[3].IsEnabled = $true
            $SyncHash.tabMenu.SelectedIndex = 3
        })
} 



# gets initial values saved in json and loads into PCO
function Get-InitialValues {
    Param ( 
        [parameter(Mandatory = $true)]$GroupName
    )

    $basePath = Join-Path -Path $PSScriptRoot -ChildPath base

    if (Test-Path (Join-Path -Path $basePath -ChildPath ($($GroupName) + '.json'))) { $initialConfig = (Get-Content (Join-Path -Path $basePath -ChildPath ($($GroupName) + '.json'))) | ConvertFrom-Json }

    return $initialConfig
}

# process loaded data or creates initial item templates for various config datagrids
function Set-InitialValues {
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory, ValueFromPipeline)]$type,
        [Parameter(Mandatory)][Hashtable]$ConfigHash,
        [switch]$PullDefaults     
    )

    Begin { 
        Set-ADGenericQueryNames -ConfigHash $ConfigHash
    }

    Process {
       
        # check if values already exist
        if ($ConfigHash.$type) { $tempList = $ConfigHash.$type }
        
        # pull from base templates if not
        elseif ($PullDefaults) { $tempList = Get-InitialValues -GroupName $type }

        else { $tempList = $null }
        

        # create observable collection and add values
        $ConfigHash.$type = New-Object -TypeName System.Collections.ObjectModel.ObservableCollection[Object]
        if ($type -eq 'objectToolConfig') {
            foreach ($item in $tempList) {
                $ConfigHash.$type.Add($item) 
                $temp = $ConfigHash.$type[([Array]::IndexOf($ConfigHash.$type, $item))].toolCommandGridConfig
                $ConfigHash.$type[([Array]::IndexOf($ConfigHash.$type, $item))].toolCommandGridConfig = New-Object -TypeName System.Collections.ObjectModel.ObservableCollection[Object]
                
                if ($temp) { $temp | ForEach-Object -Process { $ConfigHash.$type[([Array]::IndexOf($ConfigHash.$type, $item))].toolCommandGridConfig.Add($_) } }
                else {
                    $temp = Get-InitialValues -GroupName 'toolCommandGridConfig'
                    $temp | ForEach-Object -Process { $ConfigHash.$type[([Array]::IndexOf($ConfigHash.$type, $item))].toolCommandGridConfig.Add($_) }
                }
            }
        }

        else { $tempList | ForEach-Object -Process { $ConfigHash.$type.Add($_) } }


     

    }

    End {
        # get the current max values for each of the main property boxes
        $ConfigHash.UserboxCount = ($ConfigHash.userPropList | Measure-Object).Count
        $ConfigHash.CompboxCount = ($ConfigHash.compPropList | Measure-Object).Count
        $ConfigHash.boxMax = ($ConfigHash.UserboxCount, $ConfigHash.compPropList | Measure-Object -Maximum).Maximum

        if (!$configHash.searchDays) { $ConfigHash.searchDays = 60 }

    }
}

# matches config'd user/comp logins with default headers, creates new headers
# will append with number if defined values are duplicates
function Set-LoggingStructure {
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory, ValueFromPipeline)]$type,
        [Parameter(Mandatory)][Hashtable]$ConfigHash,
        [Array]$DefaultList      
    )

    Process {

        $addList = @()

        if ($ConfigHash.$type) {
            $duplicate = 0
            $ConfigHash.$type | Add-Member -MemberType NoteProperty -Name 'Header' -Value $null -ErrorAction SilentlyContinue

            $ConfigHash.$type | ForEach-Object -Process {
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

    if (($ConfigHash.queryDefConfig.Name | Measure-Object).Count -eq 1) { $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.searchPropSelection.ItemsSource = @($ConfigHash.queryDefConfig.Name) }) }
    elseif (($ConfigHash.queryDefConfig.Name | Measure-Object).Count -gt 1) { $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.searchPropSelection.ItemsSource = $ConfigHash.queryDefConfig.Name }) }

    $SyncHash.searchPropSelection.Items | ForEach-Object -Process { $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.searchPropSelection.SelectedItems.Add(($_)) }) }
}
   
function Set-RTDefaults {
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory, ValueFromPipeline)]$type,
        [Parameter(Mandatory)][Hashtable]$ConfigHash
    )

    Begin {

        if ([string]::IsNullOrEmpty($ConfigHash.rtConfig)) { $ConfigHash.rtConfig = @{ } }
        
        else {
            $rtTemp = $ConfigHash.rtConfig
            $ConfigHash.rtConfig = @{ }
            $rtTemp.PSObject.Properties | ForEach-Object -Process { $ConfigHash.rtConfig[$_.Name] = $_.Value }
            foreach ($key in $ConfigHash.rtConfig.Keys) {
                if (!$ConfigHash.rtConfig.$key.RequireOnline) {$ConfigHash.rtConfig.$key.RequireUser = $false }
            }
        }
    }

    Process {

        if ([string]::IsNullOrEmpty($ConfigHash.rtConfig.$type)) {
            $ConfigHash.rtConfig.$type = Get-InitialValues -GroupName $type
            if ($ConfigHash.rtConfig.$type.Path) { $ConfigHash.rtConfig.$type.Icon = Get-Icon -Path $ConfigHash.rtConfig.$type.Path -ToBase64 }
        }
    }
}

function Add-CustomItemBoxControls {
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory)][Hashtable]$SyncHash,
        [Parameter(Mandatory)][Hashtable]$ConfigHash
    )

    $SyncHash.customContext = @{ }

    foreach ($type in @('ubox', 'cbox')) {
        switch ($type) {
            'ubox' { $typeName = 'user' }
            'cbox' { $typeName = 'comp' }
        }

        for ($i = 1; $i -le $ConfigHash.boxMax; $i++) {
            if ($i -le $ConfigHash.($typeName + 'boxCount')) {
                $SyncHash.(($type + $i + 'resources')) = @{

                    ($type + $i + 'Border')      = New-Object -TypeName System.Windows.Controls.Border -Property @{
                        Style = $SyncHash.Window.FindResource('itemBorder')
                        Name  = ($type + $i + 'Border')
                    }

                    ($type + $i + 'Grid')        = New-Object -TypeName System.Windows.Controls.Grid
                    ($type + $i + 'DockPanel')   = New-Object -TypeName System.Windows.Controls.DockPanel
                    ($type + $i + 'StackPanel')  = New-Object -TypeName System.Windows.Controls.StackPanel -Property @{Style = $SyncHash.Window.FindResource('itemStackPanel') }

                    ($type + $i + 'Header')      = New-Object -TypeName System.Windows.Controls.Label -Property @{
                        Style = $SyncHash.Window.FindResource('itemBoxHeader')
                        Name  = ($type + $i + 'Header')           
                    }

                    ($type + $i + 'EditClip')    = New-Object -TypeName System.Windows.Controls.Label -Property @{
                        Style = $SyncHash.Window.FindResource('itemEditClip')
                        Name  = ($type + $i + 'EditClip')
                    }

                    ($type + $i + 'ViewBox')     = New-Object -TypeName System.Windows.Controls.ViewBox -Property @{Style = $SyncHash.Window.FindResource('itemViewBox') }

                    ($type + $i + 'TextBox')     = New-Object -TypeName System.Windows.Controls.TextBox -Property @{
                        Style = $SyncHash.Window.FindResource('itemBox')
                        Name  = ($type + $i + 'TextBox')
                    }

                    ($type + $i + 'Box1Action1') = New-Object -TypeName System.Windows.Controls.Button -Property @{
                        Style = $SyncHash.Window.FindResource('itemButton')
                        Name  = ($type + $i + 'Box1Action1')
                    }

                    ($type + $i + 'Box1Action2') = New-Object -TypeName System.Windows.Controls.Button -Property @{
                        Style = $SyncHash.Window.FindResource('itemButton')
                        Name  = ($type + $i + 'Box1Action2')
                    }
            
                    ($type + $i + 'Box1')        = New-Object -TypeName System.Windows.Controls.Button -Property @{
                        Style = $SyncHash.Window.FindResource('itemEditButton')
                        Name  = ($type + $i + 'Box1')
                    }

                }
        
                # Add col def objects, then add to outside grid
                $colDef1 = New-Object -TypeName System.Windows.Controls.ColumnDefinition
                $colDef1.Width = '*'

                $colDef2 = New-Object -TypeName System.Windows.Controls.ColumnDefinition
                $colDef2.Width = '75'   

                $SyncHash.(($type + $i + 'resources')).($type + $i + 'Grid').ColumnDefinitions.Add($colDef1)
                $SyncHash.(($type + $i + 'resources')).($type + $i + 'Grid').ColumnDefinitions.Add($colDef2)

                # add child controls

                $SyncHash.(($type + $i + 'resources')).($type + $i + 'Border').AddChild($SyncHash.(($type + $i + 'resources')).($type + $i + 'Grid'))
                $SyncHash.(($type + $i + 'resources')).($type + $i + 'Grid').AddChild($SyncHash.(($type + $i + 'resources')).($type + $i + 'DockPanel'))
                $SyncHash.(($type + $i + 'resources')).($type + $i + 'DockPanel').AddChild($SyncHash.(($type + $i + 'resources')).($type + $i + 'Header'))
                $SyncHash.(($type + $i + 'resources')).($type + $i + 'DockPanel').AddChild($SyncHash.(($type + $i + 'resources')).($type + $i + 'EditClip'))
                $SyncHash.(($type + $i + 'resources')).($type + $i + 'Grid').AddChild($SyncHash.(($type + $i + 'resources')).($type + $i + 'ViewBox'))
                $SyncHash.(($type + $i + 'resources')).($type + $i + 'ViewBox').AddChild($SyncHash.(($type + $i + 'resources')).($type + $i + 'TextBox'))
                $SyncHash.(($type + $i + 'resources')).($type + $i + 'Grid').AddChild($SyncHash.(($type + $i + 'resources')).($type + $i + 'StackPanel'))
                $SyncHash.(($type + $i + 'resources')).($type + $i + 'StackPanel').AddChild($SyncHash.(($type + $i + 'resources')).($type + $i + 'Box1'))
                $SyncHash.(($type + $i + 'resources')).($type + $i + 'StackPanel').AddChild($SyncHash.(($type + $i + 'resources')).($type + $i + 'Box1Action1'))
                $SyncHash.(($type + $i + 'resources')).($type + $i + 'StackPanel').AddChild($SyncHash.(($type + $i + 'resources')).($type + $i + 'Box1Action2'))
    
                # add it top item to uniform grid

                if ($type -eq 'ubox') { $SyncHash.userDetailGrid.AddChild($SyncHash.(($type + $i + 'resources')).($type + $i + 'Border')) }
                else { $SyncHash.compDetailGrid.AddChild($SyncHash.(($type + $i + 'resources')).($type + $i + 'Border')) }
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

    $SyncHash.objectTools = @{ }
    #create custom tool buttons for queried objects
    foreach ($tool in $ConfigHash.objectToolConfig) {
        if (($tool.toolActionValid -or $tool.ToolType -eq 'CommandGrid') -and !($tool.ToolType -match "Select|Grid" -and ($tool.toolSelectValid -ne $true -or $tool.toolExtraValid -ne $true))) {
            switch ($tool.objectType) {

                { $_ -match 'Both|Comp' } {
                    $SyncHash.objectTools.('ctool' + $tool.ToolID + 'buttonContent') = New-Object -TypeName System.Windows.Controls.StackPanel -Property  @{
                        Name                = $SyncHash.objectTools.('ctool' + $tool.ToolID + 'buttonContent') 
                        HorizontalAlignment = 'Center'
                        VerticalAlignment   = 'Center'
                    }

                    $SyncHash.objectTools.('ctool' + $tool.ToolID + 'buttonGlyph') = New-Object -TypeName System.Windows.Controls.Label -Property  @{
                        Content             = $tool.toolActionIcon
                        FontFamily          = 'Segoe MDL2 Assets'
                        HorizontalAlignment = 'Center'
                        FontSize            = '28'

                    }
                
                    $SyncHash.objectTools.('ctool' + $tool.ToolID + 'Label1') = New-Object -TypeName System.Windows.Controls.Label -Property  @{
                        FontFamily          = 'Segoe UI'
                        FontSize            = '9.5'
                        Margin              = '0,-8,0,0'
                        HorizontalAlignment = 'Center'
                        Content             = $tool.ToolName

                    }
              
                            
                
                    $SyncHash.objectTools.('ctool' + $tool.ToolID + 'buttonContent').AddChild($SyncHash.objectTools.('ctool' + $tool.ToolID + 'buttonGlyph'))
                    $SyncHash.objectTools.('ctool' + $tool.ToolID + 'buttonContent').AddChild($SyncHash.objectTools.('ctool' + $tool.ToolID + 'Label1'))


                    $SyncHash.objectTools.('ctool' + $tool.ToolID) = @{
                        ToolButton = New-Object -TypeName System.Windows.Controls.Button -Property  @{
                            Style   = $SyncHash.Window.FindResource('itemToolButton')
                            Name    = ('ctool' + $tool.ToolID)
                            Content = $SyncHash.objectTools.('ctool' + $tool.ToolID + 'buttonContent')
                            ToolTip = $tool.toolActionToolTip
                        }
                    }

                    $SyncHash.compToolControlPanel.AddChild($SyncHash.objectTools.('ctool' + $tool.ToolID).ToolButton)
                }
                { $_ -match 'Both|User' } {
                    $SyncHash.objectTools.('utool' + $tool.ToolID + 'buttonContent') = New-Object -TypeName System.Windows.Controls.StackPanel -Property  @{
                        Name                = $SyncHash.objectTools.('utool' + $tool.ToolID + 'buttonContent') 
                        HorizontalAlignment = 'Center'
                        VerticalAlignment   = 'Center'
                    }

                    $SyncHash.objectTools.('utool' + $tool.ToolID + 'buttonGlyph') = New-Object -TypeName System.Windows.Controls.Label -Property  @{
                        Content             = $tool.toolActionIcon
                        FontFamily          = 'Segoe MDL2 Assets'
                        HorizontalAlignment = 'Center'
                        FontSize            = '28'

                    }
                
                    $SyncHash.objectTools.('utool' + $tool.ToolID + 'Label1') = New-Object -TypeName System.Windows.Controls.Label -Property  @{
                        FontFamily          = 'Segoe UI'
                        FontSize            = '9.5'
                        Margin              = '0,-8,0,0'
                        HorizontalAlignment = 'Center'
                        Content             = $tool.ToolName

                    }
                
             
                            
                
                    $SyncHash.objectTools.('utool' + $tool.ToolID + 'buttonContent').AddChild($SyncHash.objectTools.('utool' + $tool.ToolID + 'buttonGlyph'))
                    $SyncHash.objectTools.('utool' + $tool.ToolID + 'buttonContent').AddChild($SyncHash.objectTools.('utool' + $tool.ToolID + 'Label1'))
        
                    $SyncHash.objectTools.('utool' + $tool.ToolID) = @{
                        ToolButton = New-Object -TypeName System.Windows.Controls.Button -Property  @{
                            Style   = $SyncHash.Window.FindResource('itemToolButton')
                            Name    = ('utool' + $tool.ToolID)
                            Content = $SyncHash.objectTools.('utool' + $tool.ToolID + 'buttonContent')
                            ToolTip = $tool.toolActionToolTip
                        }
                    }

                    $SyncHash.userToolControlPanel.AddChild($SyncHash.objectTools.('utool' + $tool.ToolID).ToolButton)
                }
                { $_ -eq 'Standalone' } {
                    $SyncHash.objectTools.('tool' + $tool.ToolID + 'buttonContent') = New-Object -TypeName System.Windows.Controls.StackPanel -Property  @{
                        Name                = $SyncHash.objectTools.('tool' + $tool.ToolID + 'buttonContent') 
                        HorizontalAlignment = 'Center'
                        VerticalAlignment   = 'Center'
                    }

                    $SyncHash.objectTools.('tool' + $tool.ToolID + 'buttonGlyph') = New-Object -TypeName System.Windows.Controls.Label -Property  @{
                        Content             = $tool.toolActionIcon
                        FontFamily          = 'Segoe MDL2 Assets'
                        HorizontalAlignment = 'Center'
                        FontSize            = '28'

                    }
                
                    $SyncHash.objectTools.('tool' + $tool.ToolID + 'Label1') = New-Object -TypeName System.Windows.Controls.Label -Property  @{
                        FontFamily          = 'Segoe UI'
                        FontSize            = '12'
                        HorizontalAlignment = 'Center'
                        Content             = $tool.ToolName

                    }
                
                    $SyncHash.objectTools.('tool' + $tool.ToolID + 'Label2') = New-Object -TypeName System.Windows.Controls.Label -Property  @{
                        FontFamily          = 'Segoe UI Light'
                        FontSize            = '10'
                        HorizontalAlignment = 'Center'
                        Margin              = '0,-10,0,0'
                        Content             = $tool.toolActionToolTip
                    }
                            
                
                    $SyncHash.objectTools.('tool' + $tool.ToolID + 'buttonContent').AddChild($SyncHash.objectTools.('tool' + $tool.ToolID + 'buttonGlyph'))
                    $SyncHash.objectTools.('tool' + $tool.ToolID + 'buttonContent').AddChild($SyncHash.objectTools.('tool' + $tool.ToolID + 'Label1'))
                    $SyncHash.objectTools.('tool' + $tool.ToolID + 'buttonContent').AddChild($SyncHash.objectTools.('tool' + $tool.ToolID + 'Label2'))

                    $SyncHash.objectTools.('tool' + $tool.ToolID) = @{
                        ToolButton = New-Object -TypeName System.Windows.Controls.Button -Property  @{
                            Style   = $SyncHash.Window.FindResource('standAloneButton')
                            Name    = ('tool' + $tool.ToolID)
                            Content = $SyncHash.objectTools.('tool' + $tool.ToolID + 'buttonContent')
                            ToolTip = $tool.toolActionToolTip
                                  
                        }
                    }

                    $SyncHash.standaloneControlPanel.AddChild($SyncHash.objectTools.('tool' + $tool.ToolID).ToolButton)
                }           
            }

            foreach ($toolButton in (($SyncHash.objectTools.Keys).Where{ ($_ -replace '.*tool') -eq $tool.ToolID })) {
                $SyncHash.objectTools.$toolButton.ToolButton.Add_Click( {
                        param([Parameter(Mandatory)][Object]$sender)
                        $toolID = $sender.Name -replace '.*tool'

                        switch ($ConfigHash.objectToolConfig[$toolID - 1].toolType) {

                            'Execute' {
                                $SyncHash.itemToolDialog.Title = 'Confirm'

                                if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionConfirm) {
                                    $SyncHash.itemToolDialogConfirmActionName.Text = $ConfigHash.objectToolConfig[$toolID - 1].ToolName
                                    $SyncHash.itemToolDialogConfirmObjectName.Text = $ConfigHash.currentTabItem
                                    $SyncHash.itemToolDialogConfirm.Visibility = 'Visible'
                                    $SyncHash.itemToolDialogConfirmButton.Tag = $toolID
                                    $SyncHash.itemToolDialog.IsOpen = $true                      
                                }

                                else {

                                    $rsArgs = @{
                                        Name            = 'ItemTool'
                                        ArgumentList    = @($SyncHash.snackMsg.MessageQueue, $toolID, $ConfigHash, $queryHash, $SyncHash.Window, $varHash)
                                        ModulesToImport = $configHash.modList
                                    }

                                    Start-RSJob @rsArgs -ScriptBlock {
                                        Param($queue, $toolID, $ConfigHash, $queryHash, $window, $varHash)

                                        $item = ($ConfigHash.currentTabItem).toLower()
                                        $targetType = $queryHash[$item].ObjectClass -replace 'c', 'C' -replace 'u', 'U'
                                        $toolName = ($ConfigHash.objectToolConfig[$toolID - 1].toolName).ToUpper()

                                        Set-CustomVariables -VarHash $varHash

                                        try {                     
                                            Invoke-Expression $ConfigHash.objectToolConfig[$toolID - 1].toolAction
                                            if ($ConfigHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {
                                                Write-SnackMsg -Queue $queue -ToolName $toolName -Status Success 
                                                Write-LogMessage -syncHashWindow $window -Path $ConfigHash.actionlogPath -Message Succeed -ActionName $toolName -ArrayList $ConfigHash.actionLog 
                                            }

                                            else {
                                                Write-SnackMsg -Queue $queue -ToolName $toolName -SubjectName $target -Status Success 
                                                Write-LogMessage -syncHashWindow $window -Path $ConfigHash.actionlogPath -Message Succeed -SubjectName $item -SubjectType $targetType -ActionName $toolName -ArrayList $ConfigHash.actionLog 
                                            }
                                        }
                                        catch {
                                            if ($ConfigHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {
                                                Write-SnackMsg -Queue $queue -ToolName $toolName -Status Fail 
                                                Write-LogMessage -syncHashWindow $window -Path $ConfigHash.actionlogPath -Message Fail -ActionName $toolName -ArrayList $ConfigHash.actionLog -Error $_
                                            }
                                            else {
                                                Write-SnackMsg -Queue $queue -ToolName $toolName -SubjectName $target -Status Fail 
                                                Write-LogMessage -syncHashWindow $window -Path $ConfigHash.actionlogPath -Message Succeed -SubjectName $item -ActionName $toolName -SubjectType $targetType -ArrayList $ConfigHash.actionLog -Error $_
                                            }
                                        }
                                    }
                                }
                                break
                            }
                            'Select' {
                                $SyncHash.itemToolDialog.Title = $ConfigHash.objectToolConfig[$toolID - 1].toolName
                                $SyncHash.itemToolListSelectText.Text = $ConfigHash.objectToolConfig[$toolID - 1].toolDescription
                                $SyncHash.itemToolListSelect.Visibility = 'Visible'
                                $SyncHash.itemToolListSelectConfirmButton.Tag = $toolID 
                                $SyncHash.itemToolListSelectListBox.ItemsSource = $null
                                $SyncHash.itemToolDialog.IsOpen = $true 
                        

                            if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionSelectAD -eq $false -and
                                $ConfigHash.objectToolConfig[$toolID - 1].toolActionSelectOU -eq $false -and
                                $ConfigHash.objectToolConfig[$toolID - 1].toolActionSelectCustom -eq $false) {
                                    $SyncHash.ItemToolADSelectionPanel.Visibility = 'Collapsed'

                                    $rsVars = @{
                                        target     = $ConfigHash.currentTabItem
                                        targetType = $queryHash[$ConfigHash.currentTabItem].ObjectClass
                                    }

                                    $rsArgs = @{
                                        Name            = 'PopulateListboxNoAD'
                                        ArgumentList    = @($ConfigHash, $SyncHash, $toolID, $rsVars, $syncHash.adHocConfirmWindow, $syncHash.Window, $varHash)
                                        ModulesToImport = $configHash.modList
                                    }

                                    Start-RSJob @rsArgs -ScriptBlock {
                                        param($ConfigHash, $SyncHash, $toolID, $rsVars, $confirmWindow, $window, $varHash)
                            
                                        $target = $rsVars.target
                                        $targetType = $rsVars.targetType
                                                      
                                        Set-CustomVariables -VarHash $varHash

                                        $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.itemTooListBoxProgress.Visibility = 'Visible' })
                    
                                        $list = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
                                        Invoke-Expression ($ConfigHash.objectToolConfig[$toolID - 1].toolFetchCmd) | ForEach-Object { $list.Add([PSCUstomObject]@{'Name' = $_ }) }
                                 
                                        $SyncHash.Window.Dispatcher.Invoke([Action] {
                                            $SyncHash.itemToolListSelectListBox.ItemsSource = [System.Windows.Data.ListCollectionView]$list
                                            $SyncHash.itemTooListBoxProgress.Visibility = 'Collapsed'
                                            if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionMultiSelect) {
                                                $SyncHash.itemToolListSelectListBox.SelectionMode = 'Multiple'
                                                $SyncHash.itemToolListSelectAllButton.Visibility = 'Visible'
                                            }

                                            else {
                                                $SyncHash.itemToolListSelectListBox.SelectionMode = 'Single'
                                                $SyncHash.itemToolListSelectAllButton.Visibility = 'Collapsed'
                                            }
                                        })
                                    }
                                }

                                else {
                                  
                                    if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionSelectAD) {
                                        $syncHash.itemToolADSelectionButton.Tag = 'AD' 
                                        $syncHash.itemToolItemText.Content = "AD Object Selection:" 
                                    }
                                    elseif ($ConfigHash.objectToolConfig[$toolID - 1].toolActionSelectOU) {
                                        $syncHash.itemToolADSelectionButton.Tag = 'OU' 
                                        $syncHash.itemToolItemText.Content = "OU Selection:"
                                    }
                                    elseif ($ConfigHash.objectToolConfig[$toolID - 1].toolActionSelectCustom) {
                                        $syncHash.itemToolADSelectionButton.Tag = 'Custom' 
                                        $syncHash.itemToolItemText.Content = "Custom Item Selection:"
                                    }

                                    $SyncHash.ItemToolADSelectionPanel.Visibility = 'Visible'                            
                                    $SyncHash.itemToolListSelect.Visibility = 'Visible'
                                    $SyncHash.itemToolListSelectListBox.ItemsSource = $null
                                }                          
                          
                            }
                            'Grid' {
                                $SyncHash.itemToolDialog.Title = $ConfigHash.objectToolConfig[$toolID - 1].toolName
                                $SyncHash.itemToolGridItemsGrid.ItemsSource = $null
                                $SyncHash.itemToolGridSelectConfirmButton.Tag = $toolID 
                                $SyncHash.itemToolGridSelectText.Text = $ConfigHash.objectToolConfig[$toolID - 1].toolDescription
                                $SyncHash.itemToolGridSelect.Visibility = 'Visible'
                                $SyncHash.itemToolDialog.IsOpen = $true  

                                $rsVars = @{
                                    target     = $ConfigHash.currentTabItem
                                    targetType = $queryHash[$ConfigHash.currentTabItem].ObjectClass
                                }

                                if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionSelectAD -eq $false -and $ConfigHash.objectToolConfig[$toolID - 1].toolActionSelectOU -eq $false -and $ConfigHash.objectToolConfig[$toolID - 1].toolActionSelectCustom -eq $false) {

                                    $SyncHash.itemToolGridADSelectionPanel.Visibility = 'Collapsed'
                            
                                    $rsArgs = @{
                                        Name            = 'PopulateGridbox'
                                        ArgumentList    = @($ConfigHash, $SyncHash, $toolID, $rsVars, $syncHash.adHocConfirmWindow, $syncHash.Window, $varHash)
                                        ModulesToImport = $configHash.modList
                                    }

                                    Start-RSJob @rsArgs -ScriptBlock {
                                        param($ConfigHash, $SyncHash, $toolID, $rsVars, $confirmWindow, $window, $varHash)

                                        $target = $rsVars.target
                                        $targetType = $rsVars.targetType
                                        
                                        Set-CustomVariables -VarHash $varHash
                                                      
                                        $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.itemToolGridProgress.Visibility = 'Visible' })
                                
                                        $list = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
                                        Invoke-Expression ($ConfigHash.objectToolConfig[$toolID - 1].toolFetchCmd) | ForEach-Object { $list.Add($_) }
                                 
                                        $SyncHash.Window.Dispatcher.Invoke([Action] {
                                                $SyncHash.itemToolGridItemsGrid.ItemsSource = [System.Windows.Data.ListCollectionView]$list
                                                $SyncHash.itemToolGridProgress.Visibility = 'Collapsed'
                                                $SyncHash.itemToolGridItemsGrid.Items.Refresh()
                                                if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionMultiSelect) {
                                                    $SyncHash.itemToolGridItemsGrid.SelectionMode = 'Extended'
                                                    $SyncHash.itemToolGridSelectAllButton.Visibility = 'Visible'
                                                }

                                                else {
                                                    $SyncHash.itemToolGridItemsGrid.SelectionMode = 'Single'
                                                    $SyncHash.itemToolGridSelectAllButton.Visibility = 'Collapsed'
                                                }
                                            })
                                    }
                                }

                                else {
                                    if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionSelectAD) {
                                        $SyncHash.itemToolGridADSelectionButton.Tag = 'AD' 
                                        $syncHash.itemToolGridItemText.Content = "AD Object Selection:" 
                                    }
                                    elseif ($ConfigHash.objectToolConfig[$toolID - 1].toolActionSelectOU) {  
                                        $SyncHash.itemToolGridADSelectionButton.Tag = 'OU' 
                                        $syncHash.itemToolGridItemText.Content = "OU Selection:"
                                    }
                                    elseif ($ConfigHash.objectToolConfig[$toolID - 1].toolActionSelectCustom) {
                                        $SyncHash.itemToolGridADSelectionButton.Tag = 'Custom' 
                                        $syncHash.itemToolGridItemText.Content = "Custom Item Selection:" 
                                    }
                                                                            
                                    $SyncHash.ItemToolADSelectionPanel.Visibility = 'Visible'                           
                                    $SyncHash.itemToolGridSelect.Visibility = 'Visible'
                                    $SyncHash.itemToolGridItemsGrid.ItemsSource = $null
                                }  
                            
                               
                            }
                            'CommandGrid' {
                                $SyncHash.itemToolDialog.Title = $ConfigHash.objectToolConfig[$toolID - 1].toolName
                                $SyncHash.itemCommandGridText.Text = $ConfigHash.objectToolConfig[$toolID - 1].toolDescription
                                $SyncHash.itemToolCommandGridPanel.Visibility = 'Visible'
                                $SyncHash.itemToolCommandGridDataGrid.ItemsSource = $null
                                $SyncHash.toolsCommandGridExecuteAllPanel.Visibility = 'Collapsed'                          

                                $SyncHash.itemToolDialog.IsOpen = $true
                                
                                $rsArgs = @{
                                    Name            = 'PopulateCommandGridbox'
                                    ArgumentList    = @($ConfigHash, $SyncHash, $toolID, $varHash)
                                    ModulesToImport = $configHash.modList
                                }

                                Start-RSJob @rsArgs -ScriptBlock {
                                    param($ConfigHash, $SyncHash, $toolID, $varHash)
                                  
                                    Set-CustomVariables -VarHash $varHash

                                    $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.itemCommandGridProgress.Visibility = 'Visible' })
                                    $source = $ConfigHash.objectToolConfig[$toolID - 1].toolCommandGridConfig
                                    $list = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
                                    $index = 0

                                    foreach ($item in ($source | Where-Object { $_.ActionCmdValid -eq 'True' -and $_.queryCmdValid -eq 'True' })) {
                                        $list.Add(([PSCustomObject]@{
                                                    Index           = [int]$item.ToolID - 1
                                                    ItemName        = $item.SetName 
                                                    ParentToolIndex = $toolID - 1
                                                    Result          = (Invoke-Expression $item.queryCmd).toString()
                                                }))

                                        $index++
                                    }
                                        
                                    if (($list.Result -notmatch 'True' | Measure-Object).Count -gt 0) { $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.toolsCommandGridExecuteAll.Tag = 'True' }) }
                                        
                                    Start-Sleep -Seconds 1 

                                    $SyncHash.Window.Dispatcher.Invoke([Action] {
                                            $SyncHash.itemToolCommandGridDataGrid.ItemsSource = [System.Windows.Data.ListCollectionView]$list
                                            $SyncHash.itemCommandGridProgress.Visibility = 'Collapsed'
                                            $SyncHash.toolsCommandGridExecuteAllPanel.Visibility = 'Visible'
                                        })
                                }
                            }

                        }
                    })           
            }
        }
    }
}

function Start-BasicADCheck {
    param ($SysCheckHash, $ConfigHash) 

    if ((Get-WmiObject -Class Win32_ComputerSystem).PartofDomain) {               
        $SysCheckHash.sysChecks[0].ADMember = 'True'
                    
        if ($SysCheckHash.sysChecks[0].ADModule -eq $true) {                                                 
            $selectedDC = Get-ADDomainController -Discover -Service ADWS -ErrorAction SilentlyContinue 

            if (Test-Connection -Count 1 -Quiet -ComputerName $selectedDC.HostName) {             
                $SysCheckHash.sysChecks[0].ADDCConnectivity = 'True'
                            
                try {
                    $adEntity = [Microsoft.ActiveDirectory.Management.ADEntity].Assembly
                    $adFields = $adEntity.GetType('Microsoft.ActiveDirectory.Management.Commands.LdapAttributes').GetFields('Static,NonPublic') | Where-Object -FilterScript { $_.IsLiteral }
                    $ConfigHash.adPropertyMap = @{ }     
                    $adFields | ForEach-Object -Process { $ConfigHash.adPropertyMap[$_.Name] = $_.GetRawConstantValue() }
                    
                    # create inverse hash

                    $configHash.adPropertyMapInverse = @{}
                    foreach ($key in $configHash.adPropertyMap.Keys) { $configHash.adPropertyMapInverse[$configHash.adPropertyMap.$key] = $key }

                } 

                catch { $ConfigHash.Remove('adPropertyMap') }
            }                                                  
        }  
    }
}

function Start-AdminCheck {
    param ($SysCheckHash) 
    if ((New-Object -TypeName Security.Principal.WindowsPrincipal -ArgumentList ([Security.Principal.WindowsIdentity]::GetCurrent())).isInRole([Security.Principal.WindowsBUiltInRole]::Administrator)) { $SysCheckHash.sysChecks[0].IsInAdmin = 'True' }

    if ((New-Object -TypeName Security.Principal.WindowsPrincipal -ArgumentList ([Security.Principal.WindowsIdentity]::GetCurrent())).isInRole('Domain Admins')) { 
        $SysCheckHash.sysChecks[0].IsDomainAdmin            = 'True' 
        $SysCheckHash.sysChecks[0].IsDomainAdminOrDelegated = 'True' 
    }

    else {
        if (![string]::IsNullOrEmpty($SysCheckHash.sysChecks[0].DelegatedGroupName)) {
            if ((New-Object -TypeName Security.Principal.WindowsPrincipal -ArgumentList ([Security.Principal.WindowsIdentity]::GetCurrent())).isInRole($SysCheckHash.sysChecks[0].DelegatedGroupName)) {
                $SysCheckHash.sysChecks[0].IsDomainAdminOrDelegated = 'True' 
                $SysCheckHash.sysChecks[0].IsDelegated = 'True' 
            }
        }

    }
}

function Set-RSDataContext {
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory, ValueFromPipeline)]$controlName,
        [Parameter(Mandatory)][Hashtable]$SyncHash,
        [Parameter(Mandatory)]$DataContext   
    )

    Process {
        $SyncHash.Window.Dispatcher.invoke([action] { $SyncHash.$controlName.DataContext = $DataContext })
    }
}

function Get-AdObjectPropertyList {
    param ($ConfigHash) 

    $ConfigHash.rawADValues = [System.Collections.ArrayList]@()
    $ConfigHash.QueryADValues = @{}         
      
    (Get-ADObject -Filter "(SamAccountName -eq '$($env:USERNAME)') -or (SamAccountName -eq '$($env:ComputerName)$')" -Properties *) | 
        ForEach-Object -Process {
            $_.PSObject.Properties | 
                ForEach-Object { $ConfigHash.rawADValues.Add($_.Name) | Out-Null }
            }

    foreach ($ldapValue in ($ConfigHash.rawADValues | Sort-Object -Unique)) { $ConfigHash.QueryADValues.(($ConfigHash.adPropertyMap.GetEnumerator() | Where-Object -FilterScript { $_.Value -eq $ldapValue }).Key) = $ldapValue }

    $ConfigHash.QueryADValues = ($ConfigHash.QueryADValues.GetEnumerator().Where( { $_.Key }))
}

function Remove-SavedPropertyLists {
    param ($savedConfig)
    $configRoot = Split-Path $savedConfig 
    if (Test-Path (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-user.json")) { Remove-Item  (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-user.json")}
    if (Test-Path (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-comp.json")) { Remove-Item  (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-comp.json")}
}

function Get-PropertyLists {
    param ($ConfigHash, $savedConfig, $syncHash, $window, $adLabel) 
 
    $configRoot = Split-Path $savedConfig 

    foreach ($type in ('user', 'comp')) {
        if ($type -eq 'user') {  
            if (Test-Path (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-$type.json")) {
                $ConfigHash.userPropPullList = Get-Content (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-$type.json") | ConvertFrom-Json
            }
            else { 
                 $window.Dispatcher.Invoke([Action]{$adLabel.Visibility = "Visible"})
                $ConfigHash.userPropPullList = Create-AttributeList -Type User -ConfigHash $configHash | Sort-Object -Unique Name
                $ConfigHash.userPropPullList | ConvertTo-Json | Out-File (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-$type.json")
            }
        }

        else { 
            if (Test-Path (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-$type.json")) {
                $ConfigHash.compPropPullList = Get-Content (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-$type.json") | ConvertFrom-Json
            }
            else { 
                $window.Dispatcher.Invoke([Action]{$adLabel.Visibility = "Visible"})
                $ConfigHash.compPropPullList = Create-AttributeList -Type Computer -ConfigHash $configHash | Sort-Object -Unique Name
                $ConfigHash.compPropPullList | ConvertTo-Json | Out-File (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-$type.json")
            }
        }
        
        Get-AdObjectPropertyList -ConfigHash $ConfigHash        
                
        $ConfigHash.($type + 'PropPullListNames') = [System.Collections.ArrayList]@()
        $ConfigHash.($type + 'PropPullListNames').Add('Non-AD Property') 
        $ConfigHash.($type + 'PropPullList').Name | ForEach-Object -Process { $ConfigHash.($type + 'PropPullListNames').Add($_) }
    }
}

function Set-ADGenericQueryNames {
    param($ConfigHash) 

    foreach ($id in ($ConfigHash.queryDefConfig.ID)) { $ConfigHash.queryDefConfig[$id - 1].QueryDefTypeList = $ConfigHash.QueryADValues.Key }
}



function Start-PropBoxPopulate {
    param ($ConfigHash, $savedConfig, $window, $adLabel)

    Get-PropertyLists -ConfigHash $ConfigHash -Window $window -ADLabel $adLabel -SavedConfig $savedConfig

    foreach ($type in @('User', 'Comp')) {
        $tempList = $ConfigHash.($type + 'PropList')
        $ConfigHash.($type + 'PropList') = New-Object -TypeName System.Collections.ObjectModel.ObservableCollection[Object]
 
        for ($i = 1; $i -le $ConfigHash.boxMax; $i++) {
            if ($i -le $ConfigHash.($type + 'boxCount')) {
                $ConfigHash.($type + 'PropList').Add([PSCustomObject]@{
                        Field             = $i
                                    
                        FieldName         = ( $tempList | Where-Object -FilterScript { $_.Field -eq $i }).FieldName
                                    
                        ItemType          = $type
                                    
                        PropName          = ( $tempList | Where-Object -FilterScript { $_.Field -eq $i }).PropName
                                    
                        propList          = $ConfigHash.($type + 'PropPullListNames')
                                    
                        translationCmd    = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).ValidCmd -eq $true) { ($tempList | Where-Object -FilterScript { $_.Field -eq $i }).translationCmd } 
                        else { 'if ($result -eq $false)...' }
                                    
                        actionCmd1        = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).ValidAction1 -eq $true) { ($tempList | Where-Object -FilterScript { $_.Field -eq $i }).actionCmd1 } 
                        else { "Do-Something -$type $('$' + $type)..." }
                                    
                        actionCmd1ToolTip = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).ValidAction1 -eq $true) { ($tempList | Where-Object -FilterScript { $_.Field -eq $i }).actionCmd1ToolTip } 
                        else { 'Action name' }
                  
                        actionCmd1Icon    = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).ValidAction1 -eq $true) { ($tempList | Where-Object -FilterScript { $_.Field -eq $i }).actionCmd1Icon } 
                        else { $null }
                                    
                        actionCmd1Refresh = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).actionCmd1Refresh) { $true } 
                        else { $false }  
                                    
                        actionCmd1Multi   = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).actionCmd1Multi) { $true } 
                        else { $false }  
                                    
                        ValidCmd          = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).ValidCmd) { $true } 
                        elseif ((($tempList | Where-Object -FilterScript { $_.Field -eq $i }).ValidCmd) -eq $null) { $null }
                        else { $false }
                                    
                        ValidAction1      = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).ValidAction1) { $true } 
                        elseif ((($tempList | Where-Object -FilterScript { $_.Field -eq $i }).ValidAction1) -eq $null) { $null }
                        else { $false }
                                    
                        ValidAction2      = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).ValidAction2) { $true } 
                        elseif ((($tempList | Where-Object -FilterScript { $_.Field -eq $i }).ValidAction2) -eq $null) { $null }
                        else { $false }
                                    
                        actionCmd2        = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).ValidAction2 -eq $true) { ($tempList | Where-Object -FilterScript { $_.Field -eq $i }).actionCmd2 } 
                        else { "Do-Something -$type $('$' + $type)..." }
                                    
                        actionCmd2ToolTip = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).ValidAction2 -eq $true) { ($tempList | Where-Object -FilterScript { $_.Field -eq $i }).actionCmd2ToolTip } 
                        else { 'Action name' }
                  
                        actionCmd2Icon    = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).ValidAction1 -eq $true) { ($tempList | Where-Object -FilterScript { $_.Field -eq $i }).actionCmd2Icon } 
                        else { $null }
                                    
                        actionCmd2Refresh = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).actionCmd2Refresh) { $true } 
                        else { $false }  
                                    
                        actionCmd2Multi   = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).actionCmd2Multi) { $true } 
                        else { $false }                       
                                    
                                                          
                        actionCmdsEnabled = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).actionCmdsEnabled -eq $false) { $false }
                        else { $true }
                                    
                        transCmdsEnabled  = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).transCmdsEnabled -eq $false) { $false }
                        else { $true }
                                    
                                                          
                        actionCmd2Enabled = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).ValidAction2 -eq $false) { $false }
                        elseif (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).actionCmd2Enabled) { $true } 
                        else { $false }  
                                    
                        PropType          = (($ConfigHash.($type + 'PropPullList') | Where-Object -FilterScript { $_.Name -eq (($tempList | Where-Object { $_.Field -eq $i }).PropName) }).TypeNameOfValue -replace '.*(?=\.).', '')
                                    
                        actionList        = @('ReadOnly', 'ReadOnly-Raw', 'Editable', 'Editable-Raw', 'Actionable', 'Actionable-Raw', 'Editable-Actionable', 'Editable-Actionable-Raw')
                      
                                    
                        ActionName        = if (($tempList | Where-Object -FilterScript { $_.Field -eq $i }).ActionName) { ($tempList | Where-Object -FilterScript { $_.Field -eq $i }).ActionName } 
                        else { 'null' }
              
                    })
            }        
        }                        
    }
}

function Set-QueryVarsToUpdate {
    param ($ConfigHash, [Parameter(Mandatory)][ValidateSet('User', 'Comp')]$type)

    
    switch ($type) {
        'User' { if ($ConfigHash.varListConfig.UpdateFrequency -match 'User Queries|All Queries') { $ConfigHash.varData.UpdateUser = $true } }

        'Comp' { if ($ConfigHash.varListConfig.UpdateFrequency -match 'Comp Queries|All Queries') { $ConfigHash.varData.UpdateComp = $true } }
    }
}

function Set-DurationVarsToUpdate {
    param ($ConfigHash, $startTime)

    $currentTime = Get-Date

    switch ($startTime) {
        { ($ConfigHash.varListConfig.UpdateFrequency -contains 'Daily') -and ($startTime.AddDays($ConfigHash.VarData.UpdateDayCount) -le $currentTime) } {
            $ConfigHash.varData.UpdateDay = $true
            $ConfigHash.VarData.UpdateDayCount = $ConfigHash.VarData.UpdateDayCount + 1
        }

        { ($ConfigHash.varListConfig.UpdateFrequency -contains 'Hourly') -and ($startTime.AddHours($ConfigHash.VarData.UpdateHourCount) -le $currentTime) } {
            $ConfigHash.varData.UpdateHour = $true
            $ConfigHash.VarData.UpdateHourCount = $ConfigHash.VarData.UpdateHourCount + 1
        }
        
        { ($ConfigHash.varListConfig.UpdateFrequency -contains 'Every 15 mins') -and ($startTime.AddMinutes($ConfigHash.varData.UpdateMinCount) -le $currentTime) } {
            $ConfigHash.varData.UpdateMinute = $true
            $ConfigHash.varData.UpdateMinCount = $ConfigHash.varData.UpdateMinCount + 15
        }
    
    }
}

function New-VarUpdater {
    param ($ConfigHash)
       
    $ConfigHash.varData = @{
        UpdateDayCount  = 1   
        UpdateMinCount  = 15
        UpdateHourCount = 1
    }  
}

function Start-VarUpdater {
    [CmdletBinding()]
    param ($ConfigHash, $varHash)
    
    $rsArgs = @{
        Name            = 'VarUpdater'
        ArgumentList    = @($ConfigHash, $varHash )
        ModulesToImport = $configHash.modList
    }

    Start-RSJob @rsArgs -ScriptBlock {
        param($ConfigHash, $varHash)

        $startTime = Get-Date
        $first = $true
        do {
            Set-DurationVarsToUpdate -ConfigHash $ConfigHash -StartTime $startTime

            if ($first -eq $true) {
                $ConfigHash.varData.UpdateMinute = $true
                $ConfigHash.varData.UpdateHour = $true
                $ConfigHash.varData.UpdateDay = $true
                $first = $false
            }

            if ($ConfigHash.varData.ContainsValue($true)) {
                foreach ($varInfo in ($ConfigHash.varData.Keys)) {
                    if ($ConfigHash.varData.$varInfo -eq $true) {
                        switch ($varInfo) {
                            'UpdateMinute' {
                                $ConfigHash.varListConfig |
                                    Where-Object { $_.UpdateFrequency -eq 'Every 15 mins' } |
                                        ForEach-Object { $varHash.($_.VarName) = Invoke-Expression $_.VarCmd }
                                $ConfigHash.varData.$varInfo = $false
                            }
                            'UpdateHour' {
                                $ConfigHash.varListConfig |
                                    Where-Object { $_.UpdateFrequency -eq 'Hourly' } |
                                        ForEach-Object { $varHash.($_.VarName) = Invoke-Expression $_.VarCmd }
                                $ConfigHash.varData.$varInfo = $false
                            }                                       
                            'UpdateDay' {
                                $ConfigHash.varListConfig |
                                    Where-Object { $_.UpdateFrequency -eq 'Daily' } |
                                        ForEach-Object { $varHash.($_.VarName) = Invoke-Expression $_.VarCmd }
                                $ConfigHash.varData.$varInfo = $false
                            }
                            'UpdateUser' { 
                                $ConfigHash.varListConfig |
                                    Where-Object { $_.UpdateFrequency -match 'User Queries|All Queries' } |
                                        ForEach-Object { $varHash.($_.VarName) = Invoke-Expression $_.VarCmd }
                                $ConfigHash.varData.$varInfo = $false
                            }
                            'UpdateComp' {
                                $ConfigHash.varListConfig |
                                    Where-Object { $_.UpdateFrequency -match 'Comp Queries|All Queries' } |
                                        ForEach-Object { $varHash.($_.VarName) = Invoke-Expression $_.VarCmd }
                                $ConfigHash.varData.$varInfo = $false
                            }
                        }
                    }
                }
            }
        }
        until ($ConfigHash.IsClosed -eq $true) 
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

    $syncHash.settingInfoDialogOpen.Visibility = 'Visible'

    if ($Panel) {
        $SyncHash.$Panel.Visibility = 'Visible'
        $SyncHash.settingChildWindow.IsOpen = $true
        $SyncHash.infoPaneContent.Tag = $Panel

        switch ($Panel) {
            'settingModContent' { $syncHash.settingInfoDialogOpen.Visibility = 'Hidden' }

            'settingADContent' { $syncHash.settingInfoDialogOpen.Visibility = 'Hidden' }

            'settingUserPropContent' { 
                if ($ConfigHash.UserPropList -ne $null) { $SyncHash.settingUserPropGrid.ItemsSource = $ConfigHash.UserPropList }
                $SyncHash.settingPropContent.Visibility = 'Visible'
                $SyncHash.settingUserPropGrid.Visibility = 'Visible'
            }

            'settingCompPropContent' {
                if ($ConfigHash.CompPropList -ne $null) { $SyncHash.settingCompPropGrid.ItemsSource = $ConfigHash.CompPropList }
                $SyncHash.settingPropContent.Visibility = 'Visible'
                $SyncHash.settingCompPropGrid.Visibility = 'Visible'
            }

            'settingItemToolsContent' {
                if ($ConfigHash.objectToolConfig -ne $null) { $SyncHash.settingObjectToolsPropGrid.ItemsSource = $ConfigHash.objectToolConfig }
                $SyncHash.settingObjectToolsPropGrid.Visibility = 'Visible'
            }
        
            'settingContextPropContent' {
                if ($ConfigHash.contextConfig -ne $null) { $SyncHash.settingContextPropGrid.ItemsSource = $ConfigHash.contextConfig }
                $SyncHash.settingContextGrid.Visibility = 'Visible'
                $SyncHash.settingContextPropGrid.Visibility = 'Visible'
            }

            'settingVarContent' {            
                if ($ConfigHash.varListConfig -ne $null) { $SyncHash.settingVarDataGrid.ItemsSource = $ConfigHash.varListConfig }
                $SyncHash.settingVarDataGrid.Visibility = 'Visible'
                $SyncHash.settingModDataGrid.Visibility = 'Collapsed'
                $SyncHash.settingVarAddClick.Tag = 'Var'
            }

            'settingModDataGrid' {
                if ($ConfigHash.modConfig -ne $null) { $SyncHash.settingModDataGrid.ItemsSource = $ConfigHash.modConfig }
                $SyncHash.settingModDataGrid.Visibility = 'Visible'
                $SyncHash.settingVarDataGrid.Visibility = 'Collapsed'
                $SyncHash.settingVarAddClick.Tag = 'Mod'
            }

            'settingOUDataGrid' {
                if ($ConfigHash.searchbaseConfig -ne $null) { $SyncHash.settingOUDataGrid.ItemsSource = $ConfigHash.searchbaseConfig }
                $SyncHash.settingOUDataGrid.Visibility = 'Visible' 
                $SyncHash.settingGeneralAddClick.Tag = 'OU'           
            }    
            
            'settingQueryDefDataGrid' { 
                if ($ConfigHash.queryDefConfig -ne $null) { $SyncHash.settingQueryDefDataGrid.ItemsSource = $ConfigHash.queryDefConfig }
                $SyncHash.settingQueryDefDataGrid.Visibility = 'Visible' 
                $SyncHash.settingGeneralAddClick.Tag = 'Query'   
            } 
            
            'settingMiscGrid' { 
                $SyncHash.settingGeneralAddClick.Tag = 'null'
                $SyncHash.settingMiscGrid.Visibility = 'Visible'
            } 

            'settingMiscGrid' {
             
               $syncHash.settingSearchDaySpan.Value = $ConfigHash.searchDays
               

            }
        }
    }

    if ($Title) { $SyncHash.settingChildWindow.Title = $Title }
    if ($Height) { $SyncHash.settingChildHeight.Height = $Height }
    if ($Width) { $SyncHash.settingChildHeight.Width = $Width }

    switch ($Background) {       
        'Standard' { $SyncHash.settingChildWindow.TitleBarBackground = ($SyncHash.Window.BorderBrush.Color).ToString(); break }
        'Flyout' { $SyncHash.settingChildWindow.TitleBarBackground = ($SyncHash.settingNameFlyout.Background.Color).ToString() }
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
        foreach ($contentPane in ($SyncHash.Keys.Where( { $_ -like 'setting*Content' }))) { $SyncHash.$contentPane.Visibility = 'Hidden' }
    }

    if (!($SkipDataGridReset)) {
        foreach ($dataGrid in ($SyncHash.Keys.Where( { $_ -like 'setting*Grid' }))) { $SyncHash.$dataGrid.Visibility = 'Collapsed' }
    }

    if (!($SkipFlyoutClose)) {
        foreach ($flyOut in ($SyncHash.Keys.Where( { $_ -like 'setting*Flyout' }))) { $SyncHash.$flyOut.IsOpen = $false }
    }

    if ($Title) { Set-ChildWindow -SyncHash $SyncHash -Title $Title }

    if (!($SkipResize)) { Set-ChildWindow -SyncHash $SyncHash -Width 400 -Height 215 }
    
    Set-ChildWindow -SyncHash $SyncHash -Background Standard
}

#endregion

#region itemToolFunctions

function Set-ADItemBox {
    param(
        [Parameter(Mandatory)][hashtable]$ConfigHash, 
        [Parameter(Mandatory)][hashtable]$SyncHash,
        [Parameter(Mandatory)][ValidateSet('ListBox', 'Grid')][string]$Control) 

    if ($Control -eq 'Grid') { $toolID = $SyncHash.itemToolGridSelectConfirmButton.Tag }
    else { $toolID = $SyncHash.itemToolListSelectConfirmButton.Tag  }

    $rsArgs = @{
        Name            = ('populate' + $Control)
        ArgumentList    = @($ConfigHash, $SyncHash, $toolID, $Control, $varHash)
        ModulesToImport = $configHash.modList
    }

    Start-RSJob @rsArgs -ScriptBlock {
        param($ConfigHash, $SyncHash, $toolID, $Control, $varHash)

        Set-CustomVariables -VarHash $varHash

        $SyncHash.Window.Dispatcher.Invoke([Action] {
                if ($Control -eq 'ListBox') { $SyncHash.itemToolListSelectListBox.ItemsSource = $null }
                else { $SyncHash.itemToolGridItemsGrid.ItemsSource = $null }
            })

        $selectedObject = (Select-ADObject -Type All -MultiSelect $false).FetchedAttributes -replace '$'
     
        $SyncHash.Window.Dispatcher.Invoke([Action] {
                if ($Control -eq 'ListBox') { 
                    $SyncHash.itemToolADSelectedItem.Content = $selectedObject
                    $SyncHash.itemTooListBoxProgress.Visibility = 'Visible'
                }
                else {
                    $SyncHash.itemToolGridADSelectedItem.Content = $selectedObject
                    $SyncHash.itemToolGridProgress.Visibility = 'Visible'
                }
            })
     
        $list = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
        
        if ($Control -eq 'ListBox') { Invoke-Expression ($ConfigHash.objectToolConfig[$toolID - 1].toolFetchCmd) | ForEach-Object { $list.Add([PSCustomObject]@{'Name' = $_ }) } }
        else { Invoke-Expression ($ConfigHash.objectToolConfig[$toolID - 1].toolFetchCmd) | ForEach-Object { $list.Add($_) } }

        $SyncHash.Window.Dispatcher.Invoke([Action] {
                if ($Control -eq 'ListBox') {
                    $SyncHash.itemToolListSelectListBox.ItemsSource = [System.Windows.Data.ListCollectionView]$list
                    $SyncHash.itemTooListBoxProgress.Visibility = 'Collapsed'

                    if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionMultiSelect) {
                        $SyncHash.itemToolListSelectListBox.SelectionMode = 'Multiple'
                        $SyncHash.itemToolListSelectAllButton.Visibility = 'Visible'
                    }

                    else {
                        $SyncHash.itemToolListSelectListBox.SelectionMode = 'Single'
                        $SyncHash.itemToolListSelectAllButton.Visibility = 'Collapsed'
                    }

                }
                else { 
                    $SyncHash.itemToolGridItemsGrid.ItemsSource = [System.Windows.Data.ListCollectionView]$list 
                    $SyncHash.itemToolGridProgress.Visibility = 'Collapsed'

                       if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionMultiSelect) {
                        $SyncHash.itemToolGridItemsGrid.SelectionMode = 'Extended'
                        $SyncHash.itemToolGridSelectAllButton.Visibility = 'Visible'
                    }

                    else {
                        $SyncHash.itemToolGridItemsGrid.SelectionMode = 'Single'
                        $SyncHash.itemToolGridSelectAllButton.Visibility = 'Collapsed'
                    }


                }
            })
    }      
}

function Set-OUItemBox {
    param(
        [Parameter(Mandatory)][hashtable]$ConfigHash, 
        [Parameter(Mandatory)][hashtable]$SyncHash,
        [Parameter(Mandatory)][ValidateSet('ListBox', 'Grid')][string]$Control) 

    if ($Control -eq 'Grid') {  $toolID = $SyncHash.itemToolGridSelectConfirmButton.Tag  }
    else { $toolID = $SyncHash.itemToolListSelectConfirmButton.Tag }

    $rsArgs = @{
        Name            = ('populate' + $Control)
        ArgumentList    = @($ConfigHash, $SyncHash, $toolID, $Control, $varHash)
        ModulesToImport = $configHash.modList
    }

    Start-RSJob @rsArgs -ScriptBlock {
        param($ConfigHash, $SyncHash, $toolID, $Control, $varHash)

        Set-CustomVariables -VarHash $varHash

        $SyncHash.Window.Dispatcher.Invoke([Action] {
                if ($Control -eq 'ListBox') { $SyncHash.itemToolListSelectListBox.ItemsSource = $null  }
                else { $SyncHash.itemToolGridItemsGrid.ItemsSource = $null  }
            })

        $selectedObject = (Choose-ADOrganizationalUnit -HideNewOUFeature $true).DistinguishedName
     
        $SyncHash.Window.Dispatcher.Invoke([Action] {
                if ($Control -eq 'ListBox') { 
                    $SyncHash.itemToolADSelectedItem.Content = $selectedObject
                    $SyncHash.itemTooListBoxProgress.Visibility = 'Visible'
                }
                else {
                    $SyncHash.itemToolGridADSelectedItem.Content = $selectedObject
                    $SyncHash.itemToolGridProgress.Visibility = 'Visible'
                }
            })
     
        $list = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
        
        if ($Control -eq 'ListBox') { Invoke-Expression ($ConfigHash.objectToolConfig[$toolID - 1].toolFetchCmd) | ForEach-Object { $list.Add([PSCustomObject]@{'Name' = $_ }) } }
        else { Invoke-Expression ($ConfigHash.objectToolConfig[$toolID - 1].toolFetchCmd) | ForEach-Object { $list.Add($_) } }

        $SyncHash.Window.Dispatcher.Invoke([Action] {
                if ($Control -eq 'ListBox') {
                    $SyncHash.itemToolListSelectListBox.ItemsSource = [System.Windows.Data.ListCollectionView]$list
                    $SyncHash.itemTooListBoxProgress.Visibility = 'Collapsed'

                    if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionMultiSelect) {
                        $SyncHash.itemToolListSelectListBox.SelectionMode = 'Multiple'
                        $SyncHash.itemToolListSelectAllButton.Visibility = 'Visible'
                    }

                    else {
                        $SyncHash.itemToolListSelectListBox.SelectionMode = 'Single'
                        $SyncHash.itemToolListSelectAllButton.Visibility = 'Collapsed'
                    }

                }
                else { 
                    $SyncHash.itemToolGridItemsGrid.ItemsSource = [System.Windows.Data.ListCollectionView]$list 
                    $SyncHash.itemToolGridProgress.Visibility = 'Collapsed'

                       if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionMultiSelect) {
                        $SyncHash.itemToolGridItemsGrid.SelectionMode = 'Extended'
                        $SyncHash.itemToolGridSelectAllButton.Visibility = 'Visible'
                    }

                    else {
                        $SyncHash.itemToolGridItemsGrid.SelectionMode = 'Single'
                        $SyncHash.itemToolGridSelectAllButton.Visibility = 'Collapsed'
                    }


                }
            })
    }      
}

function Set-CustomItemBox {
    param(
        [Parameter(Mandatory)][hashtable]$ConfigHash, 
        [Parameter(Mandatory)][hashtable]$SyncHash,
        [Parameter(Mandatory)][ValidateSet('ListBox', 'Grid')][string]$Control) 

    if ($Control -eq 'Grid') { $toolID = $SyncHash.itemToolGridSelectConfirmButton.Tag }
    else { $toolID = $SyncHash.itemToolListSelectConfirmButton.Tag }

    $rsArgs = @{
        Name            = ('populate' + $Control)
        ArgumentList    = @($ConfigHash, $SyncHash, $toolID, $Control, $varHash)
        ModulesToImport = $configHash.modList
    }

    Start-RSJob @rsArgs -ScriptBlock {
        param($ConfigHash, $SyncHash, $toolID, $Control, $varHash)

        Set-CustomVariables -VarHash $varHash

        $configHash.customDialogClosed = $false
        $configHash.customInput = $null

        $SyncHash.Window.Dispatcher.Invoke([Action] {
                if ($Control -eq 'ListBox') { $SyncHash.itemToolListSelectListBox.ItemsSource = $null }
                else { $SyncHash.itemToolGridItemsGrid.ItemsSource = $null }
                $syncHash.itemToolCustomContent.Text = $null
                $syncHash.itemToolCustomDialog.Visibility = 'Visible'
                $syncHash.itemToolCustomDialog.IsOpen = $true
            })

        
        do { } until ($configHash.customDialogClosed -eq $true)
        
        $SyncHash.Window.Dispatcher.Invoke([Action] {$syncHash.itemToolCustomDialog.IsOpen = $false })
        Start-Sleep -Seconds 1
        $selectedObject = $configHash.customInput

        $SyncHash.Window.Dispatcher.Invoke([Action] {
                if ($Control -eq 'ListBox') { 
                    $SyncHash.itemToolADSelectedItem.Content = $selectedObject
                    $SyncHash.itemTooListBoxProgress.Visibility = 'Visible'
                }
                else {
                    $SyncHash.itemToolGridADSelectedItem.Content = $selectedObject
                    $SyncHash.itemToolGridProgress.Visibility = 'Visible'
                }
            })
     
        $list = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
        
        if ($Control -eq 'ListBox') { Invoke-Expression ($ConfigHash.objectToolConfig[$toolID - 1].toolFetchCmd) | ForEach-Object { $list.Add([PSCustomObject]@{'Name' = $_ }) } }
        else { Invoke-Expression ($ConfigHash.objectToolConfig[$toolID - 1].toolFetchCmd) | ForEach-Object { $list.Add($_) } }

        $SyncHash.Window.Dispatcher.Invoke([Action] {
                if ($Control -eq 'ListBox') {
                    $SyncHash.itemToolListSelectListBox.ItemsSource = [System.Windows.Data.ListCollectionView]$list
                    $SyncHash.itemTooListBoxProgress.Visibility = 'Collapsed'

                    if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionMultiSelect) {
                        $SyncHash.itemToolListSelectListBox.SelectionMode = 'Multiple'
                        $SyncHash.itemToolListSelectAllButton.Visibility = 'Visible'
                    }

                    else {
                        $SyncHash.itemToolListSelectListBox.SelectionMode = 'Single'
                        $SyncHash.itemToolListSelectAllButton.Visibility = 'Collapsed'
                    }

                }
                else { 
                    $SyncHash.itemToolGridItemsGrid.ItemsSource = [System.Windows.Data.ListCollectionView]$list 
                    $SyncHash.itemToolGridProgress.Visibility = 'Collapsed'
                    $syncHash.itemToolCustomDialog.Visibility = 'Collapsed'

                       if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionMultiSelect) {
                        $SyncHash.itemToolGridItemsGrid.SelectionMode = 'Extended'
                        $SyncHash.itemToolGridSelectAllButton.Visibility = 'Visible'
                    }

                    else {
                        $SyncHash.itemToolGridItemsGrid.SelectionMode = 'Single'
                        $SyncHash.itemToolGridSelectAllButton.Visibility = 'Collapsed'
                    }


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

    $rsArgs = @{
        Name            = 'ItemToolAction'
        ArgumentList    = @($ConfigHash, $ItemList, $SyncHash.snackMsg.MessageQueue, $SyncHash.('itemTool' + $Control + 'SelectConfirmButton').Tag, $SyncHash.Window, $queryHash, $varHash)
        ModulesToImport = $ConfigHash.modList
    }

    Start-RSJob @rsArgs -ScriptBlock {
        param($ConfigHash, $ItemList, $queue, $toolID, $window, $queryHash, $varHash) 

        $toolName = $ConfigHash.objectToolConfig[$toolID - 1].toolName
        $target = $ConfigHash.currentTabItem
        $targetType = $queryHash[$target].ObjectClass -replace 'u', 'U' -replace 'c', 'C'

        Set-CustomVariables -VarHash $varHash

        try {
            Invoke-Expression -Command $ConfigHash.objectToolConfig[$toolID - 1].toolTargetFetchCmd

            foreach ($selectedItem in $ItemList) { Invoke-Expression -Command $ConfigHash.objectToolConfig[$toolID - 1].toolAction }
            
            if ($ConfigHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {
                Write-SnackMsg -Queue $queue -ToolName $toolName -Status Success 
                Write-LogMessage -syncHashWindow $window -Path $ConfigHash.actionlogPath -Message Succeed -ActionName $toolName -SubjectType 'Standalone' -ArrayList $ConfigHash.actionLog
            }
            
            else {
                Write-SnackMsg -Queue $queue -ToolName $toolName -SubjectName $target -Status Success 
                Write-LogMessage -syncHashWindow $window -Path $ConfigHash.actionlogPath -Message Succeed -SubjectName $target -SubjectType $targetType -ActionName $toolName -ArrayList $ConfigHash.actionLog 
            }
        }
        
        catch {
            if ($ConfigHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {
                Write-SnackMsg -Queue $queue -ToolName $toolName -Status Fail 
                Write-LogMessage -syncHashWindow $window -Path $ConfigHash.actionlogPath -Message Fail -ActionName $toolName -ArrayList $ConfigHash.actionLog -Error $_
            }
            else {
                Write-SnackMsg -Queue $queue -ToolName $toolName -SubjectName $target -Status Fail 
                Write-LogMessage -syncHashWindow $window -Path $ConfigHash.actionlogPath -Message Fail -SubjectName $target -SubjectType $targetType  -ActionName $toolName -ArrayList $ConfigHash.actionLog -Error $_
            }
        }
    }

    $SyncHash.itemToolDialog.IsOpen = $false
}

function New-DialogWait {
    param($confirmWindow, $window, $configHash)

    $window.Dispatcher.Invoke([action]{$confirmWindow.IsOpen = $true})
    $configHash.confirmCode = 'wait'
                
    do {} until ($configHash.confirmCode  -match 'continue|cancel')

    if ($configHash.confirmCode -eq 'cancel') {
        exit
    }
}

#endregion

#region RemoteTools

function Get-RTFlyoutContent {
    param (
        $SyncHash,
        $ConfigHash
    )
    
    Set-CurrentPane -SyncHash $syncHash -Panel "rtConfigFlyout"

    $SyncHash.settingRemoteListTypes.ItemsSource = $ConfigHash.nameMapList

    switch ($SyncHash.settingRALabel.Content) {
            
        'MSTSC' {
            if ($ConfigHash.rtConfig.MSTSC.Icon) { $SyncHash.settingRTIcon.Source = ([Convert]::FromBase64String($ConfigHash.rtConfig.MSTSC.Icon)) }

            $SyncHash.settingRemoteListTypes.Items |
                Where-Object { $_.Name -in $ConfigHash.rtConfig.MSTSC.Types } |
                    ForEach-Object { $SyncHash.settingRemoteListTypes.SelectedItems.Add(($_)) }
                
            break
        }

        'MSRA' {
            if ($ConfigHash.rtConfig.MSRA.Icon) { $SyncHash.settingRTIcon.Source = ([Convert]::FromBase64String($ConfigHash.rtConfig.MSRA.Icon)) }
            $SyncHash.settingRemoteListTypes.Items |
                Where-Object { $_.Name -in $ConfigHash.rtConfig.MSRA.Types } |
                    ForEach-Object { $SyncHash.settingRemoteListTypes.SelectedItems.Add(($_)) }
            break
        }

        Default {
            $rtID = 'rt' + [string]($SyncHash.settingRALabel.Content -replace '.[A-Z]* ')

            if ($ConfigHash.rtConfig.$rtID.Icon) { $SyncHash.settingRTIcon.Source = ([Convert]::FromBase64String($ConfigHash.rtConfig.$rtID.Icon)) }
               
            $SyncHash.settingRemoteListTypes.Items |
                Where-Object { $_.Name -in $ConfigHash.rtConfig.$rtID.Types } |
                    ForEach-Object { $SyncHash.settingRemoteListTypes.SelectedItems.Add(($_)) }
        }
    }
}

function Set-SelectedRTTypes {
    param (
        $SyncHash,
        $ConfigHash
    )

    $SyncHash.settingRTIcon.Source = $null

    switch ($SyncHash.settingRALabel.Content) {
        'MSTSC' { 
            $ConfigHash.rtConfig.MSTSC.Types = @()
            $SyncHash.settingRemoteListTypes.SelectedItems | ForEach-Object { $ConfigHash.rtConfig.MSTSC.Types += $_.Name }
            break 
        }
        'MSRA' {
            $ConfigHash.rtConfig.MSRA.Types = @()
            $SyncHash.settingRemoteListTypes.SelectedItems | ForEach-Object { $ConfigHash.rtConfig.MSRA.Types += $_.Name }
            break 
        }
        Default { 
            $rtID = 'rt' + [string]($SyncHash.settingRALabel.Content -replace '.[A-Z]* ')
            $ConfigHash.rtConfig.$rtID.Types = @()
            $SyncHash.settingRemoteListTypes.SelectedItems | ForEach-Object { $ConfigHash.rtConfig.$rtID.Types += $_.Name }
            break
        }
    }
    
    $SyncHash.settingRemoteListTypes.SelectedItems.Clear()
}

function Set-StaticRTContent {
    param (
        $SyncHash,
        $ConfigHash,
        [ValidateSet('MSRA', 'MSTSC')][string]$tool
    )

    $syncHash.settingChildWindow.ShowCloseButton = $false
    Set-CurrentPane -SyncHash $syncHash -Panel "rtConfigFlyout"

    $SyncHash.settingRtExeSelect.Visibility = 'Hidden'
    $SyncHash.settingRtPathSelect.Visibility = 'Hidden'
    $SyncHash.rtSettingRequiresOnline.Visibility = 'Hidden'
    $SyncHash.rtSettingRequiresUser.Visibility = 'Hidden'
    $SyncHash.rtDock.DataContext = $ConfigHash.rtConfig.$tool
    $SyncHash.settingRemoteFlyout.isOpen = $true
    Set-ChildWindow -SyncHash $SyncHash -Title "Remote Tool Options ($tool)" -Background Flyout
    $syncHash.settingChildWindow.ShowCloseButton = $false
}

function Get-RTExePath {
    param (
        $SyncHash,
        $ConfigHash
    )

    $rtID = 'rt' + [string]($SyncHash.settingRALabel.Content -replace '.[A-Z]* ')

    $customSelection = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        initialDirectory = [Environment]::GetFolderPath('ProgramFilesx86')
        Title            = 'Select Custom RT Executable'
    }

    $null = $customSelection.ShowDialog()

    if (![string]::IsNullOrEmpty($customSelection.fileName)) {
        if (Test-Path $customSelection.fileName) {
            $ConfigHash.rtConfig.$rtID.Path = $customSelection.fileName
            $SyncHash.settingRtPathSelect.Text = $customSelection.fileName

            if ($customSelection.fileName -like '\\*') {
                Copy-Item $customSelection.fileName -Destination C:\tmp.exe
                $ConfigHash.rtConfig.$rtID.Icon = Get-Icon -Path C:\tmp.exe -ToBase64
                Remove-Item c:\tmp.exe
            }

            else { $ConfigHash.rtConfig.$rtID.Icon = Get-Icon -Path $customSelection.fileName -ToBase64 }

            $SyncHash.settingRTIcon.Source = ([Convert]::FromBase64String(($ConfigHash.rtConfig.$rtID.Icon)))
        }
    }
}

function New-CustomRTConfigControls {
    param (
        $ConfigHash,
        $SyncHash,
        $rtID,
        [switch]$NewTool
    )

    $SyncHash.customRt.$rtID = @{
        parentDock      = New-Object System.Windows.Controls.DockPanel -Property @{
            Margin              = '0,0,25,0' 
            HorizontalAlignment = 'Stretch' 
        }

        childStack      = New-Object System.Windows.Controls.StackPanel
            
        InfoHeader      = New-Object System.Windows.Controls.Label -Property @{
            Content = "Custom Tool $($rtID -replace 'rt')"
            Style   = $SyncHash.Window.FindResource('rtHeader')
        }
            
        InfoSubheader   = New-Object System.Windows.Controls.TextBlock -Property @{
            Text  = if ($NewTool) { "Custom remote tool $($rtID -replace 'rt')" }
            else { $ConfigHash.rtConfig.$rtID.DisplayName }
            Style = $SyncHash.Window.FindResource('rtSubHeader')
        }
            
        ConfigureButton = New-Object System.Windows.Controls.Button -Property  @{
            Style = $SyncHash.Window.FindResource('rtClick')
        }
            
        DelButton       = New-Object System.Windows.Controls.Button -Property  @{
            Style = $SyncHash.Window.FindResource('rtClickDel')
        }
    }

    $SyncHash.customRt.$rtID.ConfigureButton.Name = $rtID
    $SyncHash.customRt.$rtID.DelButton.Name = $rtID + 'del'
    
    $SyncHash.customRt.$rtID.parentDock.AddChild($SyncHash.customRt.$rtID.childStack)
    $SyncHash.customRt.$rtID.parentDock.AddChild($SyncHash.customRt.$rtID.ConfigureButton)
    $SyncHash.customRt.$rtID.parentDock.AddChild($SyncHash.customRt.$rtID.DelButton)
    $SyncHash.customRt.$rtID.childStack.AddChild($SyncHash.customRt.$rtID.InfoHeader)
    $SyncHash.customRt.$rtID.childStack.AddChild($SyncHash.customRt.$rtID.InfoSubheader)
  
    $SyncHash.settingRTPanel.AddChild($SyncHash.customRt.$rtID.parentDock)

    if ($NewTool) {
        $ConfigHash.rtConfig.$rtID = [PSCustomObject]@{
            Name          = "Custom Tool $($rtID -replace 'rt')"
            Path          = $null
            Icon          = $null
            Cmd           = ' '
            Types         = @()
            RequireOnline = $true
            RequireUser   = $false
            DisplayName   = 'Tool'
        }
    }
 
    $SyncHash.customRt.$rtID.DelButton.Add_Click( {
            param([Parameter(Mandatory)][Object]$sender)
            $rtID = $sender.Name -replace 'del'
            $SyncHash.customRt.$rtID.parentDock.Visibility = 'Collapsed'
            $SyncHash.customRt.$rtID.Clear()
            $ConfigHash.rtConfig.Remove($rtID)
        })    

    $SyncHash.customRt.$rtID.ConfigureButton.Add_Click( {
            param([Parameter(Mandatory)][Object]$sender)
            $rtID = $sender.Name           
            $SyncHash.settingRtExeSelect.Visibility = 'Visible'
            $SyncHash.settingRtPathSelect.Visibility = 'Visible'
            $SyncHash.rtSettingRequiresOnline.Visibility = 'Visible'
            $SyncHash.rtSettingRequiresUser.Visibility = 'Visible'
            $SyncHash.rtDock.DataContext = $ConfigHash.rtConfig.$rtID
            Set-ChildWindow -SyncHash $SyncHash -Title "Remote Tool Options ($rtID)" -Background Flyout
            $SyncHash.settingRemoteFlyout.isOpen = $true
            $syncHash.settingChildWindow.ShowCloseButton = $false
        })
}

function Add-CustomRTControls {
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory)][Hashtable]$SyncHash,
        [Parameter(Mandatory)][Hashtable]$ConfigHash
    )


    $SyncHash.customRT = @{ }
    foreach ($rtID in $ConfigHash.rtConfig.Keys.Where{ $_ -like 'RT*' }) { New-CustomRTConfigControls -ConfigHash $ConfigHash -SyncHash $SyncHash -RTID $rtID }
}


#endregion

#region NetworkingMapping
function Convert-IpAddressToMaskLength([string] $dottedIpAddressString)
{
  $result = 0; 
  # ensure we have a valid IP address
  [IPAddress] $ip = $dottedIpAddressString;
  $octets = $ip.IPAddressToString.Split('.');
  foreach($octet in $octets)
  {
    while(0 -ne $octet) 
    {
      $octet = ($octet -shl 1) -band [byte]::MaxValue
      $result++; 
    }
  }
  return $result;
}
function Set-NetworkMapItem {
    param (
        $SyncHash,
        $ConfigHash,
        [switch]$Import
    )

    if ($Import) {

        if ($SyncHash.settingNetImportClick.SelectedIndex -eq 0) {
            $localAddress = ((Get-NetIPInterface -AddressFamily IPv4 |
                        Get-NetIPAddress |
                            Where-Object { $_.PrefixOrigin -ne 'WellKnown' }))
            
            foreach ($address in $localAddress) {
                $ip = [ipaddress]$address.IPAddress
                $subNet = [ipaddress]([ipaddress]([math]::pow(2, 32) - 1 -bxor [math]::pow(2, (32 - $($address.PrefixLength))) - 1))
                $netid = [ipaddress]($ip.address -band $subNet.address)
          
                $ConfigHash.netMapList.Add([PSCustomObject]@{
                        Id           = ($ConfigHash.netMapList | Measure-Object).Count + 1
                        Network      = $netid.IPAddressToString
                        ValidNetwork = $true
                        Mask         = $address.PrefixLength
                        ValidMask    = $true
                        Location     = 'Default'
                    })                                                               
            }    
        }

        elseif ($SyncHash.settingNetImportClick.SelectedIndex -eq 1) {

            $subnets = Get-ADReplicationSubnet -Filter * -Properties * | Select-Object Name, Site, Location, Description

            for ($i = 1; $i -le (($subnets | Measure-Object).Count); $i++) {
                $ConfigHash.netMapList.Add([PSCustomObject]@{
                        Id           = ($ConfigHash.netMapList | Measure-Object).Count + 1
                        Network      = ($subnets[$i - 1].Name -replace '//*.*', '')
                        ValidNetwork = $true
                        Mask         = ($subnets[$i - 1].Name -replace '.*/', '')
                        ValidMask    = $true
                        Location     = if ($subnets[$i - 1].Location -ne $null) { $subnets[$i - 1].Location }
                        elseif ($subnets[$i - 1].Description -ne $null) { $subnets[$i - 1].Description }
                    
                    })                                                               
            }
        }


        elseif ($SyncHash.settingNetImportClick.SelectedIndex -eq 2) {

            $scopes =  Get-DhcpServerInDC | ForEach-Object {Get-DhcpServerv4Scope -ComputerName $_.DNSName} |
                            Select-Object ScopeID, SubnetMask, Name

            foreach ($scope in $scopes) {
                $ConfigHash.netMapList.Add([PSCustomObject]@{
                        Id           = ($ConfigHash.netMapList | Measure-Object).Count + 1
                        Network      = $scope.ScopeID
                        ValidNetwork = $true
                        Mask         = Convert-IpAddressToMaskLength -dottedIpAddressString $scope.SubnetMask
                        ValidMask    = $true
                        Location     = $scope.Name
                    })                                                               
            }     
        }

        $SyncHash.settingNetDataGrid.ItemsSource = $ConfigHash.netMapList
    }

    else {
        $ConfigHash.netMapList.Add([PSCustomObject]@{
                ID           = ($ConfigHash.netMapList.ID |
                        Sort-Object -Descending |
                            Select-Object -First 1) + 1
                Network      = $null
                ValidNetwork = $false
                Mask         = $null
                ValidMask    = $false
                Location     = 'New'
            })    
    }
}


#endregion

#region UserLog

function Set-LogMapGrid { 
    param (
        $SyncHash,
        $ConfigHash,
        [parameter(Mandatory)][ValidateSet('Comp', 'User')]$type)

    # Get last 10 entries of newest logged item; select latest of these entries not ending 
    # in a comma (which would indicate that var was empty on login)
    $testLog = Get-Content ((Get-ChildItem -Path $ConfigHash.($type + 'LogPath') |
                Sort-Object LastWriteTime -Descending |
                    Select-Object -First 1).FullName) |
                Select-Object -Last 11 |
                    Where-Object { $_.Trim() -ne '' -and $_ -notmatch ",.[(\W)]$" -and $_ -notmatch ",$" } | 
                        Select-Object -Last 1

    # If empty, all latest entries have a seperate field without a value, so we'll just grab the last non-empty line
    if (!$testLog) { 
        $testLog = Get-Content ((Get-ChildItem -Path $ConfigHash.($type + 'LogPath') | 
                    Sort-Object LastWriteTime -Descending |
                        Select-Object -First 1).FullName) |
                    Where-Object { $_.Trim() -ne '' } |
                        Select-Object -Last 1
    }

    $fieldCount = ($testLog.ToCharArray() |
            Where-Object { $_ -eq ',' } |
                Measure-Object).Count + 1

    $header = @()
    for ($i = 1; $i -le $fieldCount; $i++) { $header += $i }

    $csv = $testLog | ConvertFrom-Csv -Header $header
    
    if (!($ConfigHash.($type + 'LogMapping'))) {
        $ConfigHash.($type + 'LogMapping') = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
            
        for ($i = 1; $i -le $fieldCount; $i++) {
            $ConfigHash.($type + 'LogMapping').Add([PSCustomObject]@{
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
        $currentMapping = $ConfigHash.($type + 'LogMapping')
        $ConfigHash.($type + 'LogMapping') = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
        
        for ($i = 1; $i -le $fieldCount; $i++) {
            $ConfigHash.($type + 'LogMapping').Add([PSCustomObject]@{
                    ID              = $i
                    Field           = $csv.$i
                    FieldSelList    = @('User', 'DateRaw', 'LoginDc', 'ClientName', 'ComputerName', 'Ignore', 'Custom')
                    FieldSel        = $currentMapping[$i - 1].FieldSel
                    CustomFieldName = $currentMapping[$i - 1].CustomFieldName
                    Ignore          = $false
                })                                                 
        }
    }
      
    $SyncHash.($type + 'LogListView').ItemsSource = $ConfigHash.($type + 'LogMapping')
    Set-ChildWindow -SyncHash $SyncHash -Title "Map $type Login Logs" -HideCloseButton -Background Flyout
    $SyncHash.('settingLogging' + $type + 'Flyout').IsOpen = $true
}

function Set-LoggingDirectory {
    param (
        $SyncHash,
        $ConfigHash,
        [parameter(Mandatory)][ValidateSet('Comp', 'User')]$type)

    $selectedDirectory = New-FolderSelection -Title 'Select client logging directory'

    if (![string]::IsNullOrEmpty($selectedDirectory) -and (Test-Path $selectedDirectory)) {
        $ConfigHash.($type + 'LogPath') = $selectedDirectory              
        $SyncHash.($type + 'LogPopupButton').IsEnabled = $true
    }

    else { $SyncHash.($type + 'LogPopupButton').IsEnabled = $false }
}



function Get-LDAPSearchNames {
    param (
        $ConfigHash,
        $SyncHash) 

    (($ConfigHash.queryDefConfig | Where-Object { $_.Name -in $SyncHash.searchPropSelection.SelectedItems }).QueryDefType |
            ForEach-Object { $ConfigHash.QueryADValues[([Array]::IndexOf($ConfigHash.QueryADValues.Key, $_))] }).Value
}


#endregion

#region Querying
function Set-ItemExpanders {
    param($SyncHash, $ConfigHash,
        [ValidateSet('Enable', 'Disable')]$IsActive)
    
    if ($IsActive -eq 'Disable') { 
        $SyncHash.Window.Dispatcher.Invoke([Action] {                 
                $SyncHash.compExpander.IsExpanded = $false
                $SyncHash.compExpanderProgressBar.Visibility = 'Visible'                    
                $SyncHash.userExpander.IsExpanded = $false 
                $SyncHash.expanderProgressBar.Visibility = 'Visible'
            })
    }
    else {
        $SyncHash.Window.Dispatcher.Invoke([Action] {                 
                $SyncHash.compExpander.IsExpanded = $true
                $SyncHash.compExpanderProgressBar.Visibility = 'Collapsed'                    
                $SyncHash.userExpander.IsExpanded = $true 
                $SyncHash.expanderProgressBar.Visibility = 'Collapsed'
            })
    }
}

function Start-ObjectSearch {
    param ($SyncHash, $ConfigHash, $queryHash, $Key)  
    
    $rsCmd = [PSObject]@{
        key        = $Key
        searchTag  = $SyncHash.SearchBox.Tag
        searchText = $SyncHash.SearchBox.Text
        queue      = $SyncHash.snackMsg.MessageQueue
        exact      = $SyncHash.searchExactToggle.IsChecked
    }

    $rsArgs = @{
        Name            = 'Search'
        ArgumentList    = $queryHash, $ConfigHash, $SyncHash, $rsCmd
        ModulesToImport = $ConfigHash.modList
    }


    Start-RSJob @rsArgs -ScriptBlock {
        param($queryHash, $ConfigHash, $SyncHash, $rsCmd)        
               
        if ($rsCmd.key -eq 'Escape') {
            $match = (Get-ADObject -Filter "(SamAccountName -eq '$($rsCmd.searchTag)' -and ObjectClass -eq 'User') -or 
			(Name -eq '$($rsCmd.searchTag)' -and ObjectClass -eq 'Computer')" -Properties SamAccountName) 
        }
        
        else {
            if (!($ConfigHash.searchBaseConfig.OU)) { $match = Get-ADObject -Filter (Get-FilterString -PropertyList $ConfigHash.queryProps -SyncHash $SyncHash -Query $rsCmd.searchText -Exact $rsCmd.Exact) -Properties SamAccountName, Name }
            else {
                $match = [System.Collections.ArrayList]@()
                $filter = Get-FilterString -PropertyList $ConfigHash.queryProps -SyncHash $SyncHash -Query $rsCmd.searchText -Exact $rsCmd.Exact
                foreach ($searchBase in ($ConfigHash.searchBaseConfig | Where-Object { $null -ne $_.OU })) {
                    $result = (Get-ADObject -Filter $filter -SearchBase $searchBase.OU -SearchScope $searchBase.QueryScope -Properties SamAccountName, Name) | Where-Object { $_.ObjectClass -match 'user|computer' }
                    if ($result) { $result | ForEach-Object { $match.Add($_) | Out-Null } }
                }
            }
        }

        if (($match | Measure-Object).Count -eq 1) {
            Set-ItemExpanders -SyncHash $SyncHash -ConfigHash $ConfigHash -IsActive Disable
                        
            if ($match.ObjectClass -eq 'User') {
                $match = (Get-ADUser -Identity $match.SamAccountName -Properties @($ConfigHash.UserPropList.PropName.Where( { $_ -ne 'Non-AD Property' }) | Sort-Object -Unique))                 
                Set-QueryVarsToUpdate -ConfigHash $ConfigHash -Type User
                Write-LogMessage -Path $ConfigHash.actionlogPath -Message Query -SubjectName $match.SamAccountName -ActionName 'Query' -SubjectType 'User' -ArrayList $ConfigHash.actionLog
            }

            elseif ($match.ObjectClass -eq 'Computer') {                       
                $match = (Get-ADComputer -Identity $match.SamAccountName -Properties @($ConfigHash.CompPropList.PropName.Where( { $_ -ne 'Non-AD Property' }) | Sort-Object -Unique))
                Set-QueryVarsToUpdate -ConfigHash $ConfigHash -Type Comp
                Write-LogMessage -Path $ConfigHash.actionlogPath -Message Query -SubjectName $match.Name -ActionName 'Query' -SubjectType 'Computer' -ArrayList $ConfigHash.actionLog
            }
                
            if ($match.SamAccountName -notin $SyncHash.tabControl.Items.Name -and $match.Name -notin $SyncHash.tabControl.Items.Name) {
                if ($match.ObjectClass -eq 'User') { Find-ObjectLogs -SyncHash $SyncHash -QueryHash $queryHash -ConfigHash $ConfigHash -Type User -Match $match }
                    
                elseif ($match.ObjectClass -eq 'Computer') { Find-ObjectLogs -SyncHash $SyncHash -QueryHash $queryHash -ConfigHash $ConfigHash -Type Comp -Match $match }
            }

            else {
                if ($match.ObjectClass -eq 'User') {
                    $match.PSObject.Properties | ForEach-Object { $queryHash.($match.SamAccountName)[$_.Name] = $_.Value }
                    $itemIndex = [Array]::IndexOf($SyncHash.tabControl.Items.Name, $($match.SamAccountName))     
                    $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.tabControl.SelectedIndex = $itemIndex })
                }
                else {
                    $match.PSObject.Properties | ForEach-Object { $queryHash.($match.Name)[$_.Name] = $_.Value }
                    $itemIndex = [Array]::IndexOf($SyncHash.tabControl.Items.Name, $($match.Name))     
                    $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.tabControl.SelectedIndex = $itemIndex })
                }   
            }                    
        }

    

        elseif (($match | Measure-Object).Count -gt 1) {
            $rsCmd.queue.Enqueue('Too many matches!')
            $SyncHash.Window.Dispatcher.Invoke([Action] {
                    $SyncHash.resultsSidePane.IsOpen = $true
                    $SyncHash.resultsSidePaneGrid.ItemsSource = $match | Select-Object Name, SamAccountName, ObjectClass
                })
        }

        else { $rsCmd.queue.Enqueue('No match!') }
    }
}

function Find-ObjectLogs {
    param (
        $SyncHash, $queryHash, $ConfigHash, $match,
        [ValidateSet('User', 'Comp')]$type)

    if ($type -eq 'User') { $idProp = 'SamAccountName' }
    else { $idProp = 'Name' }

    $queryHash.($match.$idProp) = @{ }
    
    $match.PSObject.Properties | ForEach-Object { $queryHash.($match.$idProp)[$_.Name] = $_.Value }

    if ($ConfigHash.($type + 'LogPath')) { $queryHash.$($match.$idProp).LoginLog = New-Object System.Collections.ObjectModel.ObservableCollection[Object] }

    $addItem = ($match | Select-Object @{Label = 'Name'; Expression = { $_.$idProp } })
    $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.tabControl.ItemsSource.Add($addItem) })  
    $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.tabControl.SelectedIndex = $SyncHash.tabControl.Items.Count - 1 })

    $rsArgs = @{
        Name            = $type + 'LogPull'
        ArgumentList    = $queryHash, $ConfigHash, $match, $SyncHash
        ModulesToImport = $ConfigHash.modList
    }

    if ($type -eq 'User') {
        Start-RSJob @rsArgs -ScriptBlock {
            param($queryHash, $ConfigHash, $match, $SyncHash) 
                            
            $SyncHash.userCompGrid.Dispatcher.Invoke([Action] { $SyncHash.userCompGrid.ItemsSource = $null })
            Start-Sleep -Milliseconds 1250
        
            if ($ConfigHash.UserLogPath -and (Test-Path (Join-Path -Path $ConfigHash.UserLogPath -ChildPath "$($match.SamAccountName).txt"))) {
                $queryHash.$($match.SamAccountName).LoginLogPath = (Join-Path -Path $ConfigHash.UserLogPath -ChildPath "$($match.SamAccountName).txt")
                $queryHash.$($match.SamAccountName).LoginLogRaw = Get-Content (Join-Path -Path $ConfigHash.UserLogPath -ChildPath "$($match.SamAccountName).txt") |
                    Select-Object -Last ($ConfigHash.searchDays * 2.5) | 
                        ConvertFrom-Csv -Header $ConfigHash.userLogMapping.Header |
                            Select-Object *, @{Label = 'DateTime'; Expression = { $_.DateRaw -as [datetime] } } -ExcludeProperty DateRaw |
                                Where-Object { $_.DateTime -gt (Get-Date).AddDays(-($ConfigHash.searchDays)) } |
                                    Sort-Object DateTime
                                    
                if ($queryHash.$($match.SamAccountName).LoginLogRaw) {
                    $loginCounts = $queryHash.$($match.SamAccountName).LoginLogRaw |
                        Group-Object -Property ComputerName |
                            Select-Object Name, Count
                    $queryHash.$($match.SamAccountName).LoginLogListView = [System.Windows.Data.ListCollectionView]($queryHash.$($match.SamAccountName).LoginLog)  
                    $queryHash.$($match.SamAccountName).LoginLogListView.GroupDescriptions.Add((New-Object System.Windows.Data.PropertyGroupDescription -ArgumentList 'compLogon'))
                    $SyncHash.userCompGrid.Dispatcher.Invoke([Action] { $SyncHash.userCompGrid.ItemsSource = $queryHash.$($match.SamAccountName).LoginLogListView })
                                        
                                       
                    foreach ($log in ($queryHash.$($match.SamAccountName).LoginLogRaw |
                                Group-Object ComputerName | ForEach-Object {$_ |
                                     Select-Object -ExpandProperty group |
                                         Sort-Object DateTime -Descending |
                                             Select-Object -First 1} ) | Sort-Object DateTime -Descending) {               
                                                  
                        Remove-Variable sessionInfo, clientLocation, hostLocation -ErrorAction SilentlyContinue
                                              
                        $ruleCount = ($ConfigHash.nameMapList | Measure-Object).Count
                      
                                            
                        $hostConnectivity = Test-OnlineFast -ComputerName $log.ComputerName
                                                                              
                        if ($hostConnectivity.Online) { $sessionInfo = Get-RDSession -ComputerName $log.ComputerName -UserName $match.SamAccountName -ErrorAction SilentlyContinue }
                        if ($hostConnectivity.IPV4Address) { $hostLocation = Resolve-Location -computerName $log.ComputerName -IPList $ConfigHash.netMapList -ErrorAction SilentlyContinue }
                        
                        if ($log.ClientName) { $clientOnline = Test-OnlineFast -ComputerName $log.ClientName }      
                        if ($clientOnline.Online) { $clientLocation = Resolve-Location -computerName $log.ClientName -IPList $ConfigHash.netMapList -ErrorAction SilentlyContinue }
                                            
                                            
                        $queryHash.$($match.SamAccountName).LoginLog.Add(( New-Object PSCustomObject -Property @{
                          
                                    logonTime       = Get-Date($log.DateTime) -Format MM/dd/yyyy
                            
                                    HostName        = $log.ComputerName
                            
                                    LoginDC         = $log.LoginDC -replace '\\'
                            
                                    UserName        = $match.SamAccountName
                            
                                    Connectivity    = ($hostConnectivity.Online).toString()
                            
                                    IPAddress       = $hostConnectivity.IPV4Address
                            
                                    userOnline      = if ($sessionInfo) { $true }
                                    else { $false }
                            
                                    sessionID       = if ($sessionInfo) { $sessionInfo.sessionID }
                                    else { $null }
                            
                                    IdleTime        = if ($sessionInfo) {
                                        if ('{0:dd\:hh\:mm}' -f $($sessionInfo.IdleTime) -eq '00:00:00') { 'Active' }
                                        else { '{0:dd\:hh\:mm}' -f $($sessionInfo.IdleTime) }   
                                    }
                                    else { $null }

                                    ClientName      = $log.ClientName 
                            
                                    ClientLocation  = $clientLocation
                            
                                    compLogon       = if (($queryHash.$($match.SamAccountName).LoginLog | Measure-Object).Count -eq 0) { 'Last' }
                                    else { 'Past' }
                            
                                    loginCount      = ($loginCounts | Where-Object { $_.Name -eq $log.ComputerName }).Count
                            
                                    DeviceLocation  = $hostLocation
                            
                                    Type            = for ($r = ($ruleCount - 1); $r -ge 0; $r--) {
                                        $comp = $log.ComputerName
                                        if ($r -le 0) { 'Computer' }
                                        else {
                                            if (($ConfigHash.nameMapList | Sort-Object -Property ID -Descending)[$r].Condition) {
                                                try {
                                                    if (Invoke-Expression $ConfigHash.nameMapList[$r].Condition) {
                                                        $ConfigHash.nameMapList[$r].Name
                                                        break
                                                    }
                                                }
                                                catch { }
                                            }
                                        }
                                    }                          
                                    ClientOnline    = if ($clientOnline) { ($clientOnline.Online).toString() };
                            
                                    ClientIPAddress = if ($clientOnline) { $clientOnline.IPV4Address };
                            
                                    ClientType      = for ($r = ($ruleCount - 1); $r -ge 0; $r--) {
                                        $comp = $log.ClientName
                                        if ($r -eq 0) { 'Computer' }
                                        else {                                                          
                                            if ($ConfigHash.nameMapList[$r].Condition) {
                                                try {
                                                    if (Invoke-Expression $ConfigHash.nameMapList[$r].Condition) {
                                                        $ConfigHash.nameMapList[$r].Name
                                                        break
                                                    }
                                                }
                                                catch { }
                                            }
                                        }
                                    }
                                }))
                                                        
                        if ($ConfigHash.userLogMapping.FieldSel -contains 'Custom') {
                            foreach ($customHeader in ($ConfigHash.userLogMapping | Where-Object { $_.FieldSel -eq 'Custom' })) {
                                foreach ($item in ($queryHash.$($match.SamAccountName).LoginLog)) { $item | Add-Member -Force -MemberType NoteProperty -Name $customHeader.Header -Value $log.($customHeader.Header) }
                            }
                        }

                        if (!$refreshTimer -or ($refreshTimer.Elapsed.TotalSeconds -ge 5)) {
                            $SyncHash.Window.Dispatcher.Invoke([Action] { $queryHash.$($match.SamAccountName).LoginLogListView.Refresh() })
                            $refreshTimer =  [system.diagnostics.stopwatch]::StartNew()
                        }
                            

                        if (($SyncHash.userCompGrid.Items | Measure-Object).Count -eq 1) {
                            $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.UserCompGrid.SelectedItem = $SyncHash.UserCompGrid.Items[0] })

                            $queryHash[$match.SamAccountName].logsSearched = $true
                        }
                    } 
                    
                    $SyncHash.Window.Dispatcher.Invoke([Action] { $queryHash.$($match.SamAccountName).LoginLogListView.Refresh() })
                                                                                                 
                }
                else { $queryHash[$match.SamAccountName].logsSearched = $true }
            }                                    
            else { $queryHash[$match.SamAccountName].logsSearched = $true }
        }      
    }

    else {
        Start-RSJob @rsArgs -ScriptBlock {
            param($queryHash, $ConfigHash, $match, $SyncHash) 
            $SyncHash.userCompGrid.Dispatcher.Invoke([Action] { $SyncHash.compUserGrid.ItemsSource = $null })
            Start-Sleep -Milliseconds 1250                               

            if ($ConfigHash.compLogPath -and (Test-Path (Join-Path -Path $ConfigHash.compLogPath -ChildPath "$($match.Name).txt"))) {
                $queryHash.$($match.Name).LoginLogPath = (Join-Path -Path $ConfigHash.compLogPath -ChildPath "$($match.Name).txt")
                $queryHash.$($match.Name).LoginLogRaw = Get-Content (Join-Path -Path $ConfigHash.compLogPath -ChildPath "$($match.Name).txt") | 
                    Select-Object -Last ($ConfigHash.searchDays * 2.5) |
                        ConvertFrom-Csv -Header $ConfigHash.compLogMapping.Header |
                            Select-Object *, @{Label = 'DateTime'; Expression = { $_.DateRaw -as [datetime] } } -ExcludeProperty DateRaw |
                                Where-Object { $_.DateTime -gt (Get-Date).AddDays(-($ConfigHash.searchDays)) } |
                                    Sort-Object DateTime -Descending 
                                        
                if ($queryHash.$($match.Name).LoginLogRaw) {
                    $loginCounts = $queryHash.$($match.Name).LoginLogRaw |
                        Group-Object -Property User |
                            Select-Object Name, Count
                    $queryHash.$($match.Name).LoginLogListView = [System.Windows.Data.ListCollectionView]($queryHash.$($match.Name).LoginLog)  
                    $queryHash.$($match.Name).LoginLogListView.GroupDescriptions.Add((New-Object System.Windows.Data.PropertyGroupDescription -ArgumentList 'compLogon'))
                    $SyncHash.compUserGrid.Dispatcher.Invoke([Action] { $SyncHash.compUserGrid.ItemsSource = $queryHash.$($match.Name).LoginLogListView })
                    $ruleCount = ($ConfigHash.nameMapList | Measure-Object).Count
                   

                    $compType = for ($r = ($ruleCount - 1); $r -ge 0; $r--) {
                        $comp = $match.Name
                        if ($r -eq 0) { 'Computer' }
                        else {
                            if ($ConfigHash.nameMapList[$r].Condition) {
                                try {
                                    if (Invoke-Expression $ConfigHash.nameMapList[$r].Condition) {
                                        $ConfigHash.nameMapList[$r].Name
                                        break
                                    }
                                }
                                catch {}                                     
                            }
                        }
                    }

                    $compPing = Test-OnlineFast $match.Name

                    if ($compPing.Online) { $sessionInfo = Get-RDSession -ComputerName $match.Name -ErrorAction SilentlyContinue }
                    if ($compPing.IPV4Address) { $hostLocation = Resolve-Location -computerName $match.Name -IPList $ConfigHash.netMapList -ErrorAction SilentlyContinue }

                    foreach ($log in ($queryHash.$($match.Name).LoginLogRaw |
                                Group-Object User | ForEach-Object {$_ |
                                     Select-Object -ExpandProperty group |
                                         Sort-Object DateTime -Descending |
                                             Select-Object -First 1} ) | Sort-Object DateTime -Descending) {
                        Remove-Variable clientLocation -ErrorAction SilentlyContinue

                        if ($log.ClientName) { $clientOnline = Test-OnlineFast -ComputerName $log.ClientName }
    
                        $userSession = $sessionInfo | Where-Object { $_.UserName -eq $log.User }                                                                                                              
        
                        if ($clientOnline.IPV4Address) { $clientLocation = Resolve-Location -computerName $log.ClientName -IPList $ConfigHash.netMapList -ErrorAction SilentlyContinue}

                        $queryHash.$($match.Name).LoginLog.Add(( New-Object PSCustomObject -Property @{
        
                                    logonTime       = Get-Date($log.DateTime) -Format MM/dd/yyyy
        
                                    UserName        = ($log.User).ToLower()
        
                                    LoginDC         = $log.LoginDC -replace '\\'
        
                                    Name            = (Get-ADUser -Identity $log.User).Name
        
                                    loginCount      = ($loginCounts | Where-Object { $_.Name -eq $log.User }).Count
        
                                    userOnline      = if ($userSession) { $true }
                                    else { $false }
        
                                    sessionID       = if ($userSession) { $userSession.SessionId }
                                    else { $null }
        
                                    IdleTime        = if ($userSession) {
                                        if ('{0:dd\:hh\:mm}' -f $($userSession.IdleTime) -eq '00:00:00') { 'Active' }
                                        else { '{0:dd\:hh\:mm}' -f $($userSession.IdleTime) }   
                                    }
                                    else { $null }

                                    ClientName      = $log.ClientName 

                                    Connectivity    = ($compPing.Online).toString()


                                    ClientOnline    = if ($clientOnline) { ($clientOnline.Online).toString() };
                            
                                    ClientIPAddress = if ($clientOnline.IPV4Address) { $clientOnline.IPV4Address };
        
                                    ClientLocation  = $clientLocation

                                    DeviceLocation  = $hostLocation

                                    Type            = $compType
            
                                    CompLogon       = if (($queryHash.$($match.Name).LoginLog | Measure-Object).Count -eq 0) { 'Last' }
                                    else { 'Past' }
                                                      
                                                        
                                    ClientType      = for ($r = ($ruleCount - 1); $r -ge 0; $r--) {
                                        $comp = $log.ClientName
                                        if ($comp) {
                                            if ($r -eq 0) { 'Computer' }
                                            else {
                                                if (($ConfigHash.nameMapList | Sort-Object -Property ID -Descending)[$r].Condition) {
                                                    try {
                                                        if (Invoke-Expression $ConfigHash.nameMapList[$r].Condition) {
                                                            $ConfigHash.nameMapList[$r].Name
                                                            break
                                                        }
                                                    }

                                                    catch {}
                                                }
                                            }
                                        }
                                    }
                                }))
                                        
                        if ($ConfigHash.compLogMapping.FieldSel -contains 'Custom') {
                            foreach ($customHeader in ($ConfigHash.compLogMapping | Where-Object { $_.FieldSel -eq 'Custom' })) {
                                foreach ($item in ($queryHash.$($match.Name).LoginLog)) { $item | Add-Member -Force -MemberType NoteProperty -Name $customHeader.Header -Value $log.($customHeader.Header) }
                            }
                        }

                        

                        if (!$refreshTimer -or ($refreshTimer.Elapsed.TotalSeconds -ge 5)) {
                            $SyncHash.Window.Dispatcher.Invoke([Action] { $queryHash.$($match.Name).LoginLogListView.Refresh() })
                            $refreshTimer =  [system.diagnostics.stopwatch]::StartNew()
                        }
                        

                        if (($SyncHash.compUserGrid.Items | Measure-Object).Count -eq 1) {
                            $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.compUserGrid.SelectedItem = $SyncHash.compUserGrid.Items[0] })

                            $queryHash[$match.Name].logsSearched = $true
                        }
                    }

                    $SyncHash.Window.Dispatcher.Invoke([Action] { $queryHash.$($match.Name).LoginLogListView.Refresh() })
                }
                                            
                else { $queryHash[$match.Name].logsSearched = $true }
            }
                                    
            else { $queryHash[$match.Name].logsSearched = $true }
        }        
    }
}

function Set-GridButtons {
    param (
        $SyncHash,
        $ConfigHash,
        [parameter(Mandatory)][ValidateSet('Comp', 'User')]$type,
        [switch]$SkipSelectionChange)

    if ($type -eq 'User') { 
        $itemName = 'userComp' 
        $butCode = 'r'
    }
    else {
        $itemName = 'compUser' 
        $butCode = 'rc'
    }
        
    if (!$SkipSelectionChange) {
        if ($null -like $SyncHash.($itemName + 'Grid').SelectedItem) { $SyncHash.($itemName + 'ControlPanel').IsEnabled = $false }
        else { $SyncHash.($itemName + 'ControlPanel').IsEnabled = $true }

        if ([string]::IsNullOrEmpty($SyncHash.($itemName + 'Grid').SelectedItem.ClientName)) {          
            $SyncHash.($itemName + 'FocusClientToggle').Visibility = 'Hidden'
            $SyncHash.($type + 'LogClientPropGrid').Visibility = 'Hidden'
        }

        else {       
            $SyncHash.($itemName + 'FocusClientToggle').Visibility = 'Visible'
            $SyncHash.($type + 'LogClientPropGrid').Visibility = 'Visible'
        }
    }

    if ($type -eq 'User' -and ($SyncHash.userCompFocusClientToggle.IsChecked) -and !($SkipSelectionChange)) { $SyncHash.userCompFocusHostToggle.IsChecked = $true }
    elseif ($type -eq 'Comp' -and ($SyncHash.compUserFocusClientToggle.IsChecked) -and !($SkipSelectionChange)) { $SyncHash.compUserFocusUserToggle.IsChecked = $true }
    else {
        foreach ($button in $SyncHash.Keys.Where( { $_ -like '*butbut*' })) {
            if (($SyncHash.($itemName + 'Grid').SelectedItem.Type -in $ConfigHash.rtConfig.($SyncHash[$button].Tag).Types) -and
                (!($ConfigHash.rtConfig.($SyncHash[$button].Tag).RequireOnline -and $SyncHash.($itemName + 'Grid').SelectedItem.Connectivity -eq $false)) -and 
                (!($ConfigHash.rtConfig.($SyncHash[$button].Tag).RequireUser -and $SyncHash.($itemName + 'Grid').SelectedItem.userOnline -eq $false))) { $SyncHash[$button].IsEnabled = $true }
            else { $SyncHash[$button].IsEnabled = $false }            
        }

        foreach ($button in $SyncHash.customRT.Keys) {
            if (($SyncHash.($itemName + 'Grid').SelectedItem.Type -in $ConfigHash.rtConfig.$button.Types) -and
                (!($ConfigHash.rtConfig.$button.RequireOnline -and $SyncHash.($itemName + 'Grid').SelectedItem.Connectivity -eq $false)) -and 
                (!($ConfigHash.rtConfig.$button.RequireUser -and $SyncHash.($itemName + 'Grid').SelectedItem.userOnline -eq $false))) { $SyncHash.customRT.$button.($butCode + 'but').IsEnabled = $true }
            else { $SyncHash.customRT.$button.($butCode + 'but').IsEnabled = $false }            
        }

        foreach ($button in $SyncHash.customContext.Keys) {
            if (($SyncHash.($itemName + 'Grid').SelectedItem.Type -in $ConfigHash.contextConfig[($button -replace 'cxt') - 1].Types) -and
                (!($ConfigHash.contextConfig[($button -replace 'cxt') - 1].RequireOnline -and $SyncHash.($itemName + 'Grid').SelectedItem.Connectivity -eq $false)) -and 
                (!($ConfigHash.contextConfig[($button -replace 'cxt') - 1].RequireUser -and $SyncHash.($itemName + 'Grid').SelectedItem.userOnline -eq $false))) { $SyncHash.customContext.$button.(($butCode + 'but') + 'context' + ($button -replace 'cxt')).IsEnabled = $true }
            else { $SyncHash.customContext.$button.(($butCode + 'but') + 'context' + ($button -replace 'cxt')).IsEnabled = $false }            
        }
    } 
}

function Set-ClientGridButtons {
    param (
        $SyncHash,
        $ConfigHash,
        [parameter(Mandatory)][ValidateSet('Comp', 'User')]$type)

    if ($type -eq 'User') { $itemName = 'userComp' }
    else { $itemName = 'compUser' }

    foreach ($button in ($SyncHash | Where-Object { $_ -like '*rbutbut*' })) {
        if (($SyncHash.($itemName + 'Grid').SelectedItem.ClientType -in $ConfigHash.rtConfig.($SyncHash[$button].Tag).Types) -and
            (!($ConfigHash.rtConfig.($SyncHash[$button].Tag).RequireOnline -and $SyncHash.($itemName + 'Grid').SelectedItem.ClientOnline -eq $false)) -and 
            (!($ConfigHash.rtConfig.($SyncHash[$button].Tag).RequireUser -eq $true))) { $SyncHash[$button].IsEnabled = $true }
        else { $SyncHash[$button].IsEnabled = $false }            
    }

    foreach ($button in $SyncHash.customRT.Keys) {
        if (($SyncHash.($itemName + 'Grid').SelectedItem.ClientType -in $ConfigHash.rtConfig.$button.Types) -and
            (!($ConfigHash.rtConfig.$button.RequireOnline -and $SyncHash.($itemName + 'Grid').SelectedItem.ClientOnline -eq $false)) -and 
            ($ConfigHash.rtConfig.$button.RequireUser -ne $true)) { $SyncHash.customRT.$button.rbut.IsEnabled = $true }
        else { $SyncHash.customRT.$button.rbut.IsEnabled = $false }            
    }

    foreach ($button in $SyncHash.customContext.Keys) {
        if (($SyncHash.($itemName + 'Grid').SelectedItem.ClientType -in $ConfigHash.contextConfig[($button -replace 'cxt') - 1].Types) -and
            (!($ConfigHash.contextConfig[($button -replace 'cxt') - 1].RequireOnline -and $SyncHash.($itemName + 'Grid').SelectedItem.ClientOnline -eq $false)) -and 
            ($ConfigHash.contextConfig[($button -replace 'cxt') - 1].RequireUser -ne $true)) { $SyncHash.customContext.$button.('rbutcontext' + ($button -replace 'cxt')).IsEnabled = $true }
        else { $SyncHash.customContext.$button.('rbutcontext' + ($button -replace 'cxt')).IsEnabled = $false }            
    }
}

function Set-ActionLog {
    param ($ConfigHash)
    $ConfigHash.actionLog = New-Object System.Collections.ObjectModel.ObservableCollection[Object] 
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


function Write-SnackMsg {
    param ( 
        [Parameter(Mandatory)]$queue, 
        [Parameter(Mandatory)]$toolName, 
        [ValidateSet('Fail', 'Success')]$Status,
        $SubtoolName, 
        $SubjectName)

    if ($SubtoolName) { $toolName = $toolName + ':' + $SubtoolName }
    
    if ($Status -eq 'Fail') { $subStatus = 'incomplete' }
    else { $subStatus = 'complete' }

    if ($SubjectName) { $queue.Enqueue("[$($toolName)]: $Status on $($SubjectName.toLower()) - action $subStatus") }
    else { $queue.Enqueue("[$($toolName)]: $Status - standalone tool $subStatus") }
}
        

function Write-LogMessage {
    param (
        $Path,
        [ValidateSet('Fail', 'Succeed', 'Query')]$Message,
        [Parameter(Mandatory)]$ActionName,
        $SubjectName,
        $ContextSubjectName,
        $SubjectType,
        [Parameter(Mandatory)]$ArrayList,
        $Error,
        $OldValue,
        $NewValue,
        $syncHashWindow)
   
    $textInfo = (Get-Culture).TextInfo
 
    $logMsg = ([PSCustomObject]@{
            ActionName  = $textInfo.ToTitleCase(($ActionName.ToLower() -replace '[][]')) 
            Message     = $textInfo.ToTitleCase($Message.ToLower() -replace '[][]')
            SubjectName = if ($SubjectName) { ($SubjectName -replace '[][]').ToLower()}
                          else { "[N/A]" }
            CxtSubName  = if ($ContextSubjectName) { ($ContextSubjectName -replace '[][]').ToLower()}
                          else { "[N/A]" }
            SubjectType = if ($SubjectType) { $textInfo.ToTitleCase(($SubjectType.ToLower() -replace '[][]' -replace "Comp$",'Computer')) }
                          else { "[N/A]" }
            Date        = (Get-Date -Format d)
            Time        = (Get-Date -Format t)
            DateFull    = Get-Date
            Admin       = ($env:USERNAME).ToLower()
            Error       = if ($Error) { $Error }
                          else { '[none]' }
            ogValue     = if ( $OldValue ) {  $OldValue }
                          else { '[N/A]' }
            newValue    = if ( $newValue ) { $newValue }
                          else { '[N/A]' }

        }) 

    if ($syncHashWindow) { $syncHashWindow.Dispatcher.Invoke([Action] { $ArrayList.Add($logMsg) }) }
    else { $null = $ArrayList.Add($logMsg) }

    if ($Path -and (Test-Path $Path)) { 
        if (!(Test-Path (Join-Path $Path -ChildPath "$($env:USERNAME)"))) { $null = New-Item -ItemType Directory -Path (Join-Path $Path -ChildPath "$($env:USERNAME)") }
        ($logMsg | ConvertTo-Csv -NoTypeInformation)[1] | Out-File -Append -FilePath (Join-Path $Path -ChildPath "$($env:USERNAME)\$(Get-Date -Format MM.dd.yyyy).log") -Force
    }
}

function Get-FilterString {
    param ($PropertyList, $Query, $SyncHash, [bool]$Exact)

    if ($Exact) { $compareOp = '-eq' }
    else { $compareOp = '-like' }

    for ($i = 0; $i -lt $PropertyList.Count; $i++) {
        $searchString = $searchString + "$($PropertyList[$i]) $compareOp `"*$Query*`""

        if ($i -lt ($PropertyList.Count - 1)) { $searchString = $searchString + ' -or ' }
    }

    if ($Exact) { $searchString -replace '\*' }
    else { $searchString }
}


function Get-LogItems {
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, HelpMessage = 'Data to process')]
        $InputObject,
        [Parameter(Mandatory)]$Collection
    )

    begin { $tempArray = [System.Collections.ArrayList]@() }

    process {
    
        (Get-Content $InputObject.FullName |
                ConvertFrom-Csv -Header ActionName, Message, SubjectName, ContextSubject, SubjectType, Date, Time, DateFull, Admin, Error, OldValue, NewValue) |
                Select-Object ActionName, Message, SubjectName, SubjectType, ContextSubject, Admin, Error, OldValue, NewValue, @{Label = 'Date'; Expression = { Get-Date(Get-Date($_.DateFull)) -Format 'M/d/yyyy hh:mm:ss tt' } } |
                    ForEach-Object { $tempArray.Add($_) }
        
        
    }

    end {
        $tempArray |
            Sort-Object -Property Date -Descending |
                ForEach-Object { $Collection.Add($_) } 
    }
}

function Set-FilteredLogs {
    param ($SyncHash, $LogView)

    $searchText = $SyncHash.toolsLogSearchBox.Text
    $endDateTime = (Get-Date($SyncHash.toolsLogEndDate.Text)).AddHours(23.9999)
    $startDateTime = [datetime]$SyncHash.toolsLogStartDate.Text
    # $window.Dispatcher.Invoke([Action]{$LogView.Filter = $null})

    $LogView.Filter = $null

    if ([string]::IsNullOrWhiteSpace($searchText)) { $LogView.Filter = { param ($item)  (([datetime]$item.Date -ge $startDateTime) -and ([datetime]$item.Date -le $endDateTime)) } }

    else {
        $LogView.Filter = {
            param ($item) ((([datetime]$item.Date -ge $startDateTime) -and ([datetime]$item.Date -le $endDateTime)) -and
                ((($item.ActionName -like "*$searchText*") -or ($item.SubjectName -like "*$searchText*") -or ($item.Admin -like "*$searchText*") )))
        }
    }
}

function Initialize-LogGrid {
    param([ValidateSet('All', 'User')]$Scope, $ConfigHash, $SyncHash)

    $SyncHash.toolsLogStartDate.Text = Get-Date ((Get-Date).AddDays(-7))  -Format MM/d/yyyy
    $SyncHash.toolsLogStartDate.DisplayDateStart = (Get-Date).AddDays(-7)
    $SyncHash.toolsLogStartDate.DisplayDateEnd = (Get-Date)
    $SyncHash.toolsLogEndDate.DisplayDateStart = (Get-Date).AddDays(-7)
    $SyncHash.toolsLogEndDate.DisplayDateEnd = (Get-Date)


    $rsCmd = @{
        currentUser = $env:USERNAME
        window      = $SyncHash.Window
        Scope       = $Scope
        Unloaded    = $SyncHash.toolsLogProgress.Tag
    }

    $SyncHash.toolsLogDataGrid.ItemsSource = $null
    

    $ConfigHash.logCollection = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
    $ConfigHash.logCollectionView = [System.Windows.Data.ListCollectionView]$ConfigHash.logCollection 
    $SyncHash.toolsLogDataGrid.ItemsSource = $ConfigHash.logCollectionView
    $SyncHash.toolsLogDataGrid.ItemsSource.Refresh()

    $rsArgs = @{
        Name            = 'GridInit'
        ArgumentList    = @($rsCmd, $ConfigHash, $SyncHash)
        ModulesToImport = $configHash.modList
    }

    Start-RSJob @rsArgs -ScriptBlock {
        param ($rsCmd, $ConfigHash, $SyncHash)

        Start-Sleep -Seconds 2
        
        if ($rsCmd.scope -eq 'User') { $logList = Get-ChildItem (Join-Path $ConfigHash.actionlogPath -ChildPath $rsCmd.CurrentUser) }
        else { $logList = Get-ChildItem $ConfigHash.actionlogPath -Recurse }
        
        $logList |
            Where-Object { ([datetime]($_.Name -replace '.log')) -ge ((Get-Date).AddDays(-7)) } |
                Get-LogItems -Collection $ConfigHash.logCollection 
       
        Start-Sleep -Seconds 1
        $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.toolsLogDataGrid.ItemsSource.Refresh() })
        
        $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.toolsLogProgress.Tag = 'Loaded' })
        $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.toolsLogEmpty.Tag = 'Loaded' })
    }
     
    $SyncHash.toolsLogDialog.IsOpen = $true
}

function New-LogHTMLExport {
    param([ValidateSet('All', 'User')]$Scope, $ConfigHash, $TimeFrame) 

    process {

        if ($TimeFrame) {
            switch ($TimeFrame) {

                'All' { break }
                '1 Year' { $span = (Get-Date).AddYears(-1) }
                '6 Month' { $span = (Get-Date).AddMonths(-6) }
                '3 Month' { $span = (Get-Date).AddMonths(-3) }
                '1 Month' { $span = (Get-Date).AddMonths(-1) }
                '2 Week' { $span = (Get-Date).AddDays(-14) }
                '1 Week' { $span = (Get-Date).AddDays(-7) }
                'Last Day' { $span = (Get-Date).AddDays(-1) }
                { '*' } { $span = Get-Date((Get-Date $span -Format M/d/yyyy)) }

            }
        }

        $rsArgs = @{
            Name            = 'LogWriteHtml'
            ArgumentList    = $ConfigHash, $Scope, $env:USERNAME, $TimeFrame, $span
            ModulesToImport = $ConfigHash.modList
        }   
        
        Start-RSJob @rsArgs -ScriptBlock {
            param ($ConfigHash, $Scope, $currentUser, $TimeFrame, $span)

            $logList = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
    
            if ($Scope -eq 'User') { 
                $logs = Get-ChildItem (Join-Path $ConfigHash.actionlogPath -ChildPath $env:USERNAME)
                $Title = "Log Data for $currentUser ($TimeFrame)"
            }
            else {
                $logs = Get-ChildItem $ConfigHash.actionlogPath -Recurse -File 
                $Title = "Log Data for All Admins ($TimeFrame)"
            }

            if ($span) {
                $logs |
                    Where-Object { ([datetime]($_.Name -replace '.log')) -ge $span } |
                        Get-LogItems -Collection $logList 
            }
            else { $logs | Get-LogItems -Collection $logList }


            $actionTotal = ($logList |
                    Where-Object { $_.ActionName -ne 'query' } |
                        Measure-Object).Count

            New-HTML -Name 'Logged Actions' -Temporary -Show {
                New-HTMLContent  -HeaderText $Title { New-HTMLTable -DataTable $logList -DefaultSortColumn 'Date' -DateTimeSortingFormat 'M/d/yyyy hh:mm:ss tt' -DefaultSortOrder Descending -Style display }

                New-HTMLContent -HeaderText 'Metrics' {
                    New-HTMLPanel {
                        New-HTMLChart -Gradient -Title 'Actions' -TitleAlignment center {
                            New-ChartToolbar -Download               
                            $logList |
                                Where-Object { $_.ActionName -ne 'query' } |
                                    Group-Object -Property Actionname |
                                        ForEach-Object { New-ChartPie -Name $_.Name -Value $_.Count }
                        }
                    }

                    New-HTMLPanel {
                        New-HTMLChart -Gradient -Title 'Top Users' -TitleAlignment center {
                            New-ChartLegend -Name 'Events'
                            $logList |
                                Where-Object { $_.SubjectType -eq 'user' -and $_.SubjectName -notlike $null } |
                                    Group-Object -Property SubjectName |
                                        Sort-Object -Property Count -Descending |
                                            Select-Object -First 10 |
                                                ForEach-Object { New-ChartBar -Name $_.Name -Value $_.Count }
                        }
                    }

                    if ($Scope -eq 'All') {
                        New-HTMLPanel {
                            New-HTMLChart -Gradient -Title 'Admins' -TitleAlignment center {
                                New-ChartToolbar -Download               
                                $logList |
                                    Where-Object { $_.ActionName -ne 'query' } |
                                        Group-Object -Property Admin |
                                            ForEach-Object { New-ChartPie -Name $_.Name -Value $_.Count }
                            }
                        }
                    }                     
                }       
            }
        }
    }
}


#endregion

#endwindow