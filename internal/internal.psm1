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

    $global:exToolHash = [hashtable]::Synchronized(@{})
}
function Set-Version {
    param ($Version, $SyncHash, $CID) 
    if ($Version) {
        $syncHash.toolVer.Text = $Version
    }
    if ($CID) {
        $syncHash.configVer.Text = "CID:" + ([string]$CID).PadLeft(4,'0')
    }

}

function Set-CurrentPane {
    param ($SyncHash, $Panel)
    $SyncHash.infoPaneContent.Tag = $Panel
}

function Set-WindowVisibility {
param($BasePath)
    if ($host.name -eq 'ConsoleHost') {
        $SW_HIDE, $SW_SHOW = 0, 5
        $TypeDef = '[DllImport("User32.dll")]public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'
        Add-Type -MemberDefinition $TypeDef -Namespace Win32 -Name Functions
        $hWnd = (Get-Process -Id $PID).MainWindowHandle
        $Null = [Win32.Functions]::ShowWindow($hWnd,$SW_HIDE)
        
    }
}

function Set-GlobalVars {
    param ($BasePath)
    $global:xamlPath     = Join-Path $BasePath WindowContent.xaml
    $global:glyphList    = Join-Path $BasePath \internal\base\segoeGlyphs.txt
    $global:savedConfig  =  Join-Path $BasePath config.json
    $global:ConfigMap    = Set-ConfigMap
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

function Move-DataGridItem {
    param($SyncHash, $ConfigHash, $GridName, $CollectionName, $IDName, [Parameter(Mandatory)][ValidateSet('Up', 'Down')]$Direction)

        $ID = $syncHash.$gridName.SelectedItem.$IDName
        $idRange = $configHash.$collectionName.$IDName | Measure-Object -Maximum -Minimum

        if ($Direction -eq 'Up' -and $ID -ne $idRange.Minimum) {
            ($configHash.$CollectionName | Where-Object -FilterScript { $_.$IDName -eq $ID}).$IDName = 'x'
            ($configHash.$CollectionName | Where-Object -FilterScript { $_.$IDName -eq ($ID - 1)}).$IDName= $ID
            ($configHash.$CollectionName | Where-Object -FilterScript { $_.$IDName -eq 'x'}).$IDName = $ID - 1
            $refresh = $true
        }

        elseif ($Direction -eq 'Down' -and $ID -ne $idRange.Maximum) {
            ($configHash.$CollectionName | Where-Object -FilterScript { $_.$IDName -eq $ID}).$IDName = 'x'
            ($configHash.$CollectionName | Where-Object -FilterScript { $_.$IDName -eq ($ID + 1)}).$IDName = $ID
            ($configHash.$CollectionName | Where-Object -FilterScript { $_.$IDName -eq 'x'}).$IDName = $ID + 1
            $refresh = $true
        }

        if ($refresh) {
            $syncHash.$GridName.Items.SortDescriptions.Add((New-Object -TypeName System.ComponentModel.SortDescription -ArgumentList ($IDName, 'Ascending')))
            $syncHash.$GridName.Items.Refresh()
        }
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
    param ($SyncHash, $SettingInfoHash, $ConfigHash) 
    
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
        $flowPara = New-Object -TypeName System.Windows.Documents.Paragraph -Property @{
                    Margin  = 0
                    Padding = 0
        }

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
    param($SyncHash, $ConfigHash) 
    $SyncHash.Window.Dispatcher.invoke([action] {                       
            $SyncHash.windowContent.Visibility = 'Visible'

            if ($configHash.MinWidth -and $configHash.MinWidth -is [int]) { $SyncHash.Window.MinWidth = $configHash.MinWidth }
            else {  $SyncHash.Window.MinWidth = '1000' }

            if ($configHash.MinHeight -and $configHash.MinHeight -is [int]) { $SyncHash.Window.MinHeight = $configHash.MinHeight }
            else { $SyncHash.Window.MinHeight = '700' }

            $SyncHash.Window.ResizeMode = 'CanResizeWithGrip'
            $SyncHash.Window.ShowTitleBar = $true
            $SyncHash.Window.ShowCloseButton = $true                   
            $SyncHash.splashLoad.Visibility = 'Collapsed' 
        })           
}

function Set-WPFHeader {
    param ($ConfigHash, $SyncHash)   
    if ($configHash.settingHeaderConfig) {
        if ($configHash.settingHeaderConfig.headerUser) {$syncHash.headerUser.Text = "$(($env:USERDNSDOMAIN).ToLower())\$($env:USERNAME)"}
        $syncHash.headerControl.DataContext = $configHash.settingHeaderConfig
    }
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
            if (Test-Path $ConfigPath) {
                if ((Get-ChildItem -LiteralPath $ConfigPath).Length -eq 0 -and (Get-ChildItem -LiteralPath $($savedConfig + '.bak')).Length -gt 0) { Copy-Item -LiteralPath $ConfigPath -Destination $($ConfigPath + '.bak') }

                (Get-Content $ConfigPath | ConvertFrom-Json).PSObject.Properties | ForEach-Object -Process { $ConfigHash[$_.Name] = $_.Value }
            }
        }
        'Export' {

        #null all session based data - no need to store
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
            $configHash.logCollection = $null
            $configHash.gridExportList = $null
            $configHash.logCollectionView = $null
            $configHash.IsSearching = $null
            $configHash.newVersionStubInfo = $null

          if ($null -notlike $configHash.settingHeaderConfig) { $configHash.settingHeaderConfig[0].headerColor = $configHash.settingHeaderConfig[0].headerColor.ToString() }

            $ConfigHash |
                ConvertTo-Json -Depth 8 |
                    Out-File -FilePath $($ConfigPath + '.bak') -Force

            if ((Get-ChildItem -LiteralPath $($ConfigPath + '.bak')).Length -gt 0) { Copy-Item -LiteralPath $($ConfigPath + '.bak') -Destination $ConfigPath }
        }
    }
}

function Export-VersionConfig {
    Param ( 
        [CmdletBinding()]
        [Parameter(Mandatory = $true)]
        [String]
        $ConfigPath,

        [Parameter(Mandatory = $true)]
        [Hashtable]
        $ConfigHash
    )

        $ConfigHash.configVer |
        ConvertTo-Json -Depth 8 |
            Out-File -FilePath $($ConfigPath + '.bak') -Force

    if ((Get-ChildItem -LiteralPath $($ConfigPath + '.bak')).Length -gt 0) { Copy-Item -LiteralPath $($ConfigPath + '.bak') -Destination $ConfigPath }
        
}

# Inevitably, when this is done, it will be dumped into a JSON for storage
function Set-InfoPaneHash {
    $global:settingInfoHash = @{
        'settingCompPropContent' = [PSCustomObject]@{
            'Body'  = @' 
    This table allows you to map Active Directory properties to display in the details pane after querying a computer.  These properties can be made editable or have buttons performing related actions attached to them, or can be simply read only. Additional items can be added with the '+' button.

    Each item can be given a friendly name in the 'FIELD NAME'  box - this will be displayed as the header. The corresponding drop-down box will determine the Active Directory property queried. The 'TYPE' drop down will determine the type of actions that can be performed with the item. These can be defined for each field with its respective define (DEF) button. The types are explained below.

    The 'Non-AD Property' selection is an actionable-only field that allows its content to be populated using a custom source.
'@
        'Types' = [ordered]@{
            'ReadOnly'   = 'The field only shows the content as pulled from Active Directory.'
            'Editable'   = 'The field value can be updated or cleared and then saved to Active Directory.'
            'Actionable' = 'The field will allow up to two definable buttons to perform custom actions.'
            'Raw'        = 'Any raw field will display the value directly as pulled from Active Directory. Otherwise, the presentation of the content can be defined.'
        }
    } 
    'settingUserPropContent' = [PSCustomObject]@{
        'Body'  = @' 
This table allows you to map Active Directory properties to display in the details pane after querying a user. These properties can be made editable or have buttons performing related actions attached to them, or can be simply read only. Additional items can be added with the '+' button.

Each item can be given a friendly name in the 'FIELD NAME'  box - this will be displayed as the header. The corresponding drop-down box will determine the Active Directory property queried. The 'TYPE' drop down will determine the type of actions that can be performed with the item. These can be defined for each field with its respective define (DEF) button. The types are explained below.

The 'Non-AD Property' selection is an actionable-only field that allows its content to be populated using a custom source.
'@
        'Types' = [ordered]@{
            'ReadOnly'   = 'The field only shows the content as pulled from Active Directory.'
            'Editable'   = 'The field value can be updated or cleared and then saved to Active Directory.'
            'Actionable' = 'The field will allow up to two definable buttons to perform custom actions.'
            'Raw'        = 'Any raw field will display the value directly as pulled from Active Directory. Otherwise, the presentation of the content can be defined.'
        }
    } 
    'settingPropUserDefine'  = [PSCustomObject]@{
        'Body' = @'
These fields define how the selected property will function in regards to the button type selected in the previous table. 

The 'Result Presentation' scriptblock is present in any non-raw type. This will allow the property returned to be presented as an alternative value (e.g. a TRUE or FALSE value can be passed through an if statement and alternate text can be returned and displayed).

The 'Attached Actions’ sections correspond to buttons that will attach to the returned value when queried. Their respective scriptblock will run when the button is pressed. Along each attached action, an icon can be selected for use with the button. The 'Refresh Prop' option will requery Active Directory after the action completes and update the value in the display. The 'New Thread' option will run the action in a new thread and is generally recommended, though this may not necessarily be faster overall for quicker actions. The ‘Disable if result like:’ option, when enabled, will disable the button if the value, as presented, matches the value in the ‘string match’ textbox (using basic -like wildcard options).

All script blocks must be validated using the 'Execute' button, which analyzes the box for fatal syntaxical errors and other warnings.

The variables below can be referenced or manipulated within the script blocks.
'@
        'Vars' = [ordered]@{          
            '$result [Result Presentation]'      = 'The actual value of the returned Active Directory property.'
            '$resultColor [Result Presentation]' = 'Setting the resultColor will determine the color of the returned value. This uses all valid .NET brush names and HEX values.'
            '$user [Actionable Items]'           = 'The value of the current, queried user. Only populated on action buttons for user properties.'
            '$actionObject [Actionable Items] '  = 'General variable containing hte name of queried object (i.e. the user or computer)'
            '$propName [Actionable Items]'       = 'The name of the property attached to the field. Not applicable to non-AD queries.'
            '$prop [Actionable Items]'           = 'The value of the queried property attached to the field.'
            '$activeObject' = 'The current, selected username of the item on the query tab (i.e. the queried item the tool is run on). Not applicable to standalone tools.'
            '$activeObjectData' = 'A collection of the Active Directory properties attached to the Active Object - they can be referenced with dot notation (i.e. $activeObjectData.''PropertyName''). Not applicable to standalone tools.'
            '$activeObjectType' = 'The current $activeObject  item type (user/computer). Not applicable to standalone tools.'
        }
    }
    'settingPropCompDefine'  = [PSCustomObject]@{
        'Body' = @'
The options below define how the selected item will function, as chosen by the TYPE in the previous table. 

The 'Result Presentation' scriptblock is used in any non-raw type. This will allow the property returned to be presented as an alternative value (e.g. a TRUE or FALSE value can be passed through an if statement and alternate text can be returned and displayed).

The 'Attached Actions’ sections correspond to buttons that will attach to the returned value in the details pane after an item is queried. Their respective scriptblock, defined here, will run when the button is pressed. Along each attached action, an icon can be selected for use with the button. The 'Refresh Prop' option will requery Active Directory after the action completes and update the value in the display. The 'New Thread' option will run the action in a new thread and is generally recommended, though this may not necessarily be faster overall for quicker actions. The ‘Disable if result like:’ option, when enabled, will disable the button if the value, as presented, matches the value in the ‘string match’ textbox (using basic -like wildcard options).

All script blocks must be validated using the 'Execute' button, which analyzes the box for fatal syntaxical errors and other warnings.

The variables below can be referenced or manipulated within the script blocks.
'@
        'Vars' = [ordered]@{          
            '$result [Result Presentation]'      = 'The actual value of the returned Active Directory property.'
            '$resultColor [Result Presentation]' = 'Setting the resultColor will determine the color of the returned value. This uses all valid .NET brush names and HEX values.'          
            '$comp [Actionable Items]'           = 'The value of the current, queried computer. Only populated on action buttons for computer properties.'
            '$actionObject [Actionable Items] '  = 'General variable containing hte name of queried object (i.e. the user or computer)'
            '$propName [Actionable Items]'       = 'The name of the property attached to the field. Not applicable to non-AD queries.'
            '$prop [Actionable Items]'           = 'The value of the queried property attached to the field.'
            '$activeObject' = 'The current, selected username of the item on the query tab (i.e. the queried item the tool is run on). Not applicable to standalone tools.'
            '$activeObjectData' = 'A collection of the Active Directory properties attached to the Active Object - they can be referenced with dot notation (i.e. $activeObjectData.''PropertyName''). Not applicable to standalone tools.'
            '$activeObjectType' = 'The current $activeObject  item type (user/computer). Not applicable to standalone tools.'
        }
    }
    'settingItemToolsContent' =  [PSCustomObject]@{
        'Body'  = @' 
Tools are advanced actions that can be used on queried users or computers or as standalone tools independent of either. Configuring and adding new tools can be done in the preceding table. For each entry, the TOOL NAME field is the name used on the tool label, and is used when the results of the tool is logged. The TOOL TYPE defines the template and presentation of tools. There are several different tool types - these are defined below and are described in detail within the information pane when defining the respective tool. The OBJECT TYPE refers to where the tool will be used - either on users, computers, both, or as a standalone tool in the tool tab. The DEF button opens the definition pane for the respective tool.
'@
    'Types' = [ordered]@{
	    ‘Execute’ = ‘Execute tools are the most basic tool types. When accessed, they simply execute the defined scriptblock.’
	    ‘Select’ = ‘Selection tools query data from a defined source and allow the returned data, populated as a list, to be selected and then used to manipulate the current, queried user or computer (or the listed items themselves) through a defined action. These are designed for single property lists.’
	    ‘Grid’ = 'Grid tools are similar to select tools - they allow querying data but allow items with multiple properties. These need not be tied to any action - the grid contents can be exported into HTML reports.'
	    'CommandGrid' = 'The command grid tool allows defining a series of queries. Each defined query has an associated action - if the query returns anything but the boolean ''true'' the action will be eligible to run. This is useful for running a set of actions for similar processes (e.g. diagnostics, user outboarding, etc).'
    }
}
    'settingContextPropContent' =  [PSCustomObject]@{
        'Body'  = @' 
Contextual tools are actions accessible through the historical view. It uses the currently queried object and the selection made within the historical view to perform tasks defined within this configuration section. 
'@
      
}
    'settingVarContent' =  [PSCustomObject]@{
        'Body'  = @'  
Resources and variables are items that capture defined values that update at defined intervals. These allow custom data to be accessed through out scriptblocks found throughout the tool.

Modules refer to PowerShell modules to be added to allow access to scriptblocks executing in new threads. If not added, their functions cannot be accessed in the thread.
'@
}
    'settingModDataGrid' =  [PSCustomObject]@{
        'Body'  = @' 
Items below will allow defining PowerShell modules to be accessed within scriptblocks that use new threads.

The NAME is a 'friendly name' for the module - it is only for descriptive purposes used within this list. The PATH is the local or network path to the module's .psm1 file. This can be selected by using the button in the respective row.
'@
}
    'settingOUDataGrid' =  [PSCustomObject]@{
        'Body'  = @' 
A list of the Active Directory OUs and containers for querying - these will be the locations searched during queries for users and computers.

Select the '+' button to add an additional OU definition. Select the button it its respective row to choose the OU. Select the search scope from the drop down - the search scope types are listed below.
'@
        'Types' = [ordered]@{
            'Subtree' = 'The default. This scope will search an OU and all OUs within it recursively.'
            'OneLevel' ='This will search an OU for objects only one level deep. Nested OUs are ignored.'}
}
    'settingQueryDefDataGrid' =  [PSCustomObject]@{
        'Body'  = @'
When querying for users or computers, each of the items added as query definitions will act as the properties searched to match against the provided search term. The NAME field is the 'friendly name' that will appear in the search settings (which can be toggled prior to querying). The Active Directory Property field is the actual Active Directory property queried.
'@
}
    'settingRTContent' =  [PSCustomObject]@{
        'Body'  = @' 
Remote Connection Clients reference tools used in establishing remote connections into computers, user sessions, or other systems. The tools defined in this section will appear as buttons in a computer or user's historical view. Based on the settings defined for each item, the respective remote tool's button with initiate a connection using that tool.
 
Microsoft’s RDP and Remote Assistant clients are added by default. Additional third party clients can be added by using the ‘+’ button or removed using the 'trash' button. Clients can be configured by selecting the tool’s respective gear button.
'@
}
    'settingNetContent' =  [PSCustomObject]@{
        'Body'  = @' 
Network mapping allows the ability to define your environment’s IP space by location to better assess where a particular computer is located. When querying users, login logs (if defined) are analyzed. During this process, the IP address of each computer from each log entry is evaluated to determine if the address is part of any of the networks defined in this section. If it matches a network, the historical view will display the location defined for that network.

There are several mechanisms available to import existing network information described in the Types section below. Networks can also be manually added and defined. The NETWORK field corresponds to the IP address of the network, while the MASK is its subnet mask. Invalid IP addresses or an invalid subnet will flag after input and the entry will not be not saved. The LOCATION field is the value returned when an evaluated IP address is found to be within that network’s IP space.
'@
   'Types' = [ordered]@{          
        'Import from current computer''s NIC' = 'This will import the network information from the NICs found on the current computer.'
        'Import from defined ADDS subnets'    = 'This will import the IP and mask defined in the domain'’s ADDS replication subnets. It will use the location property set on that subnet in Active Directory to populate the LOCATION field. If this is undefined, it will use its Active Directory Site’'s location property. If this is also undefined, it will use the value in the description property for the subnet.'
        'Import from defined DHCP scopes'     = 'This will import the IP and mask defined in the domain''s DHCP servers'’ scopes. It will use the scope’’s name as the LOCATION.’
    
        

    }
}
    'settingNameContent' =  [PSCustomObject]@{
        'Body'  = @' 
Computer categorization allows defining distinct sets of computer types. Contextual actions and remote tools rely on these categories to restrict or allow access.

When querying users, login logs (if defined) are analyzed. During this process, the conditions defined in this section are evaluated against the given computer from each log entry to determine its category. The rules are evaluated in descending order based on the RULE number. Once a rule's CONDITION is returned as TRUE, the evaluation is stopped and the object is categorized by the value in the NAME field for the corresponding condition. If ‘clientname’ is defined in the logs, this is also evaluated. When querying computers, the computer itself is analyzed, but each entry, if ‘clientname’ is defined, will be evaluated similarly. 

When creating a condition, selecting the respective rule’s condition field will expand the scriptblock to allow entry. The block should be written to return a boolean value and should the variable list below.

Since entries are evaluated in reverse order, the last rule - after all others have been evaluated - will categorize the evaluated computer in the generic ‘computer’ category. Custom rules can be repositioned to change the order they are evaluated in.
'@
        'Vars' = [ordered]@{
            '$comp'           = 'The computer name of the evalauted system.'
            '$clientLocation' = 'If the computer has an A record resolvable through DNS, this value will populate with the LOCATION field of the matched network defined in the ''NETWORK MAPPINGS'' configuration section'          
        }
}
    'settingGeneralContent' =  [PSCustomObject]@{
        'Body'  = @' 
These are general options related to querying, logging, and general tool settings.
'@
 'Types' = [ordered]@{	
        ‘SearchBase OUs'   = 'Defines the Active Directory OUs/containers to search when queries are made.'
        'Query Properties' = 'Defines query properties evaluated in Active Directory to match a given search term.'
        'Misc. Settings'   = 'Miscellaneous settings related to the tool.'
        } 
}
   'settingMiscGrid' =  [PSCustomObject]@{
        'Body'  = @' 
Miscellaneous settings.

The minimum window size sets the smallest size the tool will shrink down to when resized. This is helpful if a lot of content is configured and requires a larger window to show without relying on scroll bars.

The logging path defines where this tool's actions are logged. Ideally, this should be the same network location for all administrators using this tool to allow for reporting to show all actions.

Login log view depth refers to how far back (in days) login logs will be searched, analyzed, and displayed on the historical view. Larger depth will result in longer overall querying time, but this number may need to be adjusted to best fit your enviornment's usage.

Active Directory mappings are an index of the entire list of the Active Directory properties and their object types that are generating on first load. These are used in the property mappings for both users and computers. If the Active Directory schema is updated, these should be refreshed using the button in this section.

Header content allows you to select whether the domain and username of the current user shows on the header of the application. Additionally, you can add a custom label and select the color of its font.
'@
}
 'rtConfigFlyout' =  [PSCustomObject]@{
        'Body'  = @' 
The settings below define the resources for the remote tool and the systems and conditions required to recognize the tool as applicable for a given system or user.

The text in DISPLAY NAME field will appear as the tool tip for the remote tool button within the historical view after querying a user or computer.

The APPLICABLE SYSTEMS lists the computers categorized in the ‘Naming Convention’ setting section. The selected systems in this list will be eligible to use this tool. For inapplicable systems, the tool’s button will be inactive on the historical view.

The gear button will allow you to choose the path to the executable for the remote tool. The path will populate the PATH field. The icon of the file will also be captured and used on the tool’s button in the historical view. 

TARGET MUST BE ONLINE and TARGET USER MUST BE ONLINE refer to the connection and session states of the computer and user, respectively. If both or either is selected and their condition is not met,  the remote tool’s button will be inactive.

The COMMAND field holds the command executed after selecting the tool’s button. This command will reference the executable and requires usage of its command line switches. Use the variables listed below to define the connection command.
'@
    'Vars' = [ordered]@{
            '$comp'      = 'The computer name of the targeted system.'
            '$exe'       = 'The path to the remote tool''s executable.' 
            '$user'      = 'The username of the targeted user, if needed.'
            '$sessionID' = 'The session ID of the targeted user, if needed.'
                     
        }
}
'settingLoggingContent' =  [PSCustomObject]@{
        'Body'  = @' 
Select directories that store the login logs for both clients and users. Ideally, this will be generated upon user login using login scripts. The output of these should be either .csv, .txt, or .log files and must be delimited by commas. They should at least contain the username of the user, the date, and the computer name. After choosing the respective directory, the structure of the logs can be defined to map each attribute.

After selecting the directory with the gear button, the edit button will enable - from here, you can map each value to its respective type. The given values are pulled from the newest log entry in the previously defined logging path. 

The values from the newest log entry act as a reference and are shown in the FIELD section. Using the five predefined properties, use the given values in the FIELD sections to align each to its correct type in the PROPERTY section to later be used in analyzing the logs. These properties include the username (User), the date (DateRaw), login server (LoginDC), computer (ComputerName), and name of the connecting client (ClientName).  For unneeded properties, the IGNORE selection will skip the property. For custom properties outside the predefined five, the CUSTOM selection will allow you to assign a custom friendly name to display this value in the historical view when this item type is queried.
'@
       
}
'settingContextDefFlyout' =  [PSCustomObject]@{
        'Body'  = @' 
These fields define how and when the selected contextual action is peformed.

The selection(s) made within the TARGET TYPES list will determine what type of computers will allow access to this action's button. Inelligble computer types will disable the button. The values in this list are referenced from the computer types defined in the NAMING CONVENTION configuration section. 

Within the BUTTON SETTINGS section, the scriptblock will determine the actions performed when the button is selected. The variables listed below can be used within this scriptblock to reference the current selected items.

As with other scriptblocks, this tool's scriptblock must be validated before it can be used. Selecting the EXECUTE button will evaluate the block for any warnings, issues, or parsing errors - the results will be displayed in the SCRIPTBLOCK VALIDITY text box. After validating, the tooltip of the status icon found within this box will list any warnings or errors. 

The TARGET USER MUST BE LOGGED IN option will only allow this action's button to be accessible if the targeted user has a current session on the target machine. The TARGET MUST BE ONLINE option will only allow this action's button to be accessible if the targeted computer is online. The NEW THREAD option will run the button’s action in a new thread - this is suggested for longer or more complicated tasks (however, it may not be required if the action is simple or quick to complete). The ICON will appear as the button’s icon within the historical view.
'@
        'Vars' = [ordered]@{
            '$comp'      = 'The computer name of the targeted system. If the queried object is a computer, this variable refers to that item. If it is not, it refers to the selected computer in the queried user''s historical view'
            '$user'      = 'The username of the targeted user.  If the queried object is a user, this variable refers to that item. If it is not, it refers to the selected user in the queried computer''s historical view'
            '$sessionID' = 'The session ID of the targeted user, if the selected item in the historical view has an active session.'
            '$activeObject' = 'The current, selected username of the item on the query tab (i.e. the queried item the tool is run on). Not applicable to standalone tools.'
            '$activeObjectData' = 'A collection of the Active Directory properties attached to the Active Object - they can be referenced with dot notation (i.e. $activeObjectData.''PropertyName''). Not applicable to standalone tools.'
            '$activeObjectType' = 'The current $activeObject  item type (user/computer). Not applicable to standalone tools.'
                     
        }

        'Tips' = [ordered]@{
            'Confirmation Window' = @'
Adding the following line at any point within the scriptblock will launch a confirmation window before continuing the block:
    New-DialogWait -ConfirmWindow $confirmWindow -Window $window -configHash $configHash
         -TextBlock $syncHash.adHocConfirmText -Text 'Custom text here'
'@         
        }
}
'settingObjectToolCommandGrid' =  [PSCustomObject]@{
        'Body'  = @' 
To be completed.. settingObjectToolCommandGrid
'@
        'Types' = @{}
}
'settingObjectToolSelect' =  [PSCustomObject]@{
        'Body'  = @' 
Selection tools query data from a defined source and allow the returned data, populated into a list, to be selected and then used to manipulate the current, queried user or computer (or the listed items themselves) through a defined action. These are designed for single property arrays.

There are several sections to configure for a select tool definition. Firstly, a short description of what the tool does can be added in the DESCRIPTION block. This will appear in the dialog after the tool is opened and allows the opportunity to give a short explanation to the administrator of what the tool queries and what actions it will execute. 

Next, the selection query must be defined. This can be done through a variety of ways. A reference object can be defined (see the TYPES list below for a full listing and explanation of each) by selecting the PROMPT REFERENCE OBJECT option, and selecting the desired object type in the combo box. This will allow an administrator, when using the tool, to be prompted to select a reference object of that type. This reference object can then be used in the SELECTION scriptblock to query and return a set of data. Or, conversely, no reference object can be set for selection and the scriptblock defined without relying on one as a data source. The output will be used in populating the select list. The output of the scriptblock should have only one property, and that property must be expanded (i.e. by using Select-Object’s –ExpandProperty option to remove the data’s property name header).

A secondary query can be optionally defined in the TARGET ITEM ADDITIONAL DATA scriptblock. The target item refers to the currently active user or computer. Using this block, variables can be set and later used in the execution scriptblock.

Then, an action itself must be defined in the SCRIPTBLOCK block. When executed, this will script for each item selected in the list. 

Lastly, miscellaneous settings can be configured. The name of the tool – to be used in the tool’s respective action button as the tool tip – can be added. This can be more descriptive than the name previously set in the tool table that is used as the button’s label. The icon can also be set – which is also displayed on the button. Finally, the MULTISELECT option allows selecting more than one item in the list.
'@
        'Vars' = [ordered]@{
        '$activeObject' = 'The current, selected username of the item on the query tab (i.e. the queried item the tool is run on). Not applicable to standalone tools.'
        '$activeObjectData' = 'A collection of the Active Directory properties attached to the Active Object - they can be referenced with dot notation (i.e. $activeObjectData.''PropertyName''). Not applicable to standalone tools.'
        '$activeObjectType' = 'The current $activeObject  item type (user/computer). Not applicable to standalone tools.'
        '$inputObject' = 'The value returned from the reference object selection. Only applicable in the SELECTION sciptblock.'
        '$selectedItem' = 'While executing, each selected item will be iterated through and given this variable. Only appliable to the SCRIPTBLOCK block for execution.'
        }

        'Types' = [ordered]@{	
        ‘AD User’ = ‘Prompts for the selection of an Active Directory user.’
        ‘AD Computer' = ‘Prompts for the selection of an Active Directory computer.’
        ‘AD Group' = ‘Prompts for the selection of an Active Directory group.’
        ‘AD Object (any)’ = ‘Prompts for the selection of any of the above Active Directory objects.’
        ‘OU’ = ‘Prompts for the selection of an organization unit or container.’
        ‘String’ = ‘Prompts for the input of a string.’
        ‘Integer’= ‘Prompts for the input of a number.’
        ‘Choice' = ‘Prompts for the selection from a defined list. This list can be defined in the nearby textbox after selecting CHOICE as the reference object. Entries must be separated by a comma.’
        ‘File’ = ‘Prompts for the selection of a file.’
        ‘Directory’ = ‘Prompts for the selection of a directory.’
        } 



}
'settingObjectToolExecute' =  [PSCustomObject]@{
        'Body'  = @' 
To be completed.. settingObjectToolExecute
'@
        'Types' = @{}
}
'settingObjectToolGrid' =  [PSCustomObject]@{
        'Body'  = @' 
Grid tools query data from a defined source and allow the returned data, populated into a grid supporting multiple properties, to be selected and then used to manipulate the current, queried user or computer (or the listed items themselves) through a defined action. These are designed for multiple property collections.

There are several sections to configure for a grid tool definition. Firstly, a short description of what the tool does can be added in the DESCRIPTION block. This will appear in the dialog after the tool is opened and allows the opportunity to give a short explanation to the administrator of what the tool queries and what actions it will execute. 

Next, the selection query must be defined. This can be done through a variety of ways. A reference object can be defined (see the TYPES list below for a full listing and explanation of each) by selecting the PROMPT REFERENCE OBJECT option, and selecting the desired object type in the combo box. This will allow an administrator, when using the tool, to be prompted to select a reference object of that type. This reference object can then be used in the SELECTION scriptblock to query and return a set of data. Or, conversely, no reference object can be set for selection and the scriptblock defined without relying on one as a data source. The output will be used in populating the select list. The output of the scriptblock should have only one property, and that property must be expanded (i.e. by using Select-Object’s –ExpandProperty option to remove the data’s property name header).

A secondary query can be optionally defined in the TARGET ITEM ADDITIONAL DATA scriptblock. The target item refers to the currently active user or computer. Using this block, variables can be set and later used in the execution scriptblock.

Then, an action itself must be defined in the SCRIPTBLOCK block. When executed, this will script for each item selected in the list. 

Lastly, miscellaneous settings can be configured. The name of the tool – to be used in the tool’s respective action button as the tool tip – can be added. This can be more descriptive than the name previously set in the tool table that is used as the button’s label. The icon can also be set – which is also displayed on the button. Finally, the MULTISELECT option allows selecting more than one item in the list.
'@
        'Vars' = [ordered]@{
        '$activeObject' = 'The current, selected username of the item on the query tab (i.e. the queried item the tool is run on). Not applicable to standalone tools.'
        '$activeObjectData' = 'A collection of the Active Directory properties attached to the Active Object - they can be referenced with dot notation (i.e. $activeObjectData.''PropertyName''). Not applicable to standalone tools.'
        '$activeObjectType' = 'The current $activeObject  item type (user/computer). Not applicable to standalone tools.'
        '$inputObject' = 'The value returned from the reference object selection. Only applicable in the SELECTION sciptblock.'
        '$selectedItem' = 'While executing, each selected item will be iterated through and given this variable. Only appliable to the SCRIPTBLOCK block for execution.'
        }

        'Types' = [ordered]@{	
        ‘AD User’ = ‘Prompts for the selection of an Active Directory user.’
        ‘AD Computer' = ‘Prompts for the selection of an Active Directory computer.’
        ‘AD Group' = ‘Prompts for the selection of an Active Directory group.’
        ‘AD Object (any)’ = ‘Prompts for the selection of any of the above Active Directory objects.’
        ‘OU’ = ‘Prompts for the selection of an organization unit or container.’
        ‘String’ = ‘Prompts for the input of a string.’
        ‘Integer’= ‘Prompts for the input of a number.’
        ‘Choice' = ‘Prompts for the selection from a defined list. This list can be defined in the nearby textbox after selecting CHOICE as the reference object. Entries must be separated by a comma.’
        ‘File’ = ‘Prompts for the selection of a file.’
        ‘Directory’ = ‘Prompts for the selection of a directory.’
        } 



}
'settingVarDataGrid' =  [PSCustomObject]@{
        'Body'  = @' 
Items defined below will populate variables according to the set frequency and allow access from within scriptblocks defined elsewhere.

The NAME field will be the name of the variable. These can and will overwrite the default variables referenceable from within scriptblocks if the same names are used. The UPDATE FREQUENCY dictates how often these variables update. They are described in the list below. The DESCRIPTION field should explain the variable - it appears alongside the name in all information panels that list variables. The DEF field is the scriptblock that executes whenever the variable updates.
'@
        'Types'  = [ordered]@{
            'All Queries'      = 'The value will update whenever any query is made.'
            'User Queries'     = 'The value will update whenever any user query is made.'
            'Computer Queries' = 'The value will update whenever any user query is made.'
            'Daily'            = 'The value will update when the program starts, and then every 24 hours afterwards.'
            'Hourly'           = 'The value will update when the program starts, and then every hour afterwards.'
            'Every 15 minutes' = 'The value will update when the program starts, and then every 15 minutes afterwards.'
            'Program Start'    = 'The value will update only upon program start.'
        }
}
'settingAdminContent' = [PSCustomObject]@{
        'Body'  = @' 
Generally, the tool will check that it is running in an administrative context and that the launching account is in the Domain Administrator role. While the account MUST be a local administrator for the tool to function, a non-domain administrator security group can be defined to accomodate groups given delegated rights. 

If the launching user opens the tool and is not a domain administrator, this section will allow the option to select another security group to check for membership. If the launching account is a member, the tool will load normally. However, the selected group should have delegated rights to the OUs defined in the query section. Additionally, the actions themselves should be constructed within the scope of the delegated group’s limited rights.
'@
}
}
}

function Set-ConfigMap {

     [ordered]@{
        'Versioning Info'         = 'configVer'
        'Network Mapping'         = 'netMapList'
        'Computer Categorization' = 'nameMapListView'
        'Remote Tools'            = 'rtConfig'
        'User Logs Path'          = 'userLogPath'
        'User Log Mapping'        = 'userLogMapping'
        'Comp Logs Path'          = 'compLogPath'
        'Comp Log Mapping '       = 'compLogMapping'
        'Item Tools'              = 'objectToolConfig'
        'Importable Modules'      = 'modConfig'
        'User AD Properties'      = 'userPropList'
        'Computer AD Properties'  = 'compPropList'
        'Query Definitions'       = 'queryDefConfig'
        'Queriable OUs'           = 'searchBaseConfig'
        'Variable Definitions'    = 'varListConfig'
        'Contextual Actions'      = 'contextConfig' 
        'Tool Logging Path'       = 'actionLogPath'
        'Tool Categories'         = 'SACats'
    }
}


function Import-Config {
    param ($SyncHash, $configMap, $configSelection)

    if (!$configSelection) {     
        $configSelection = New-Object -TypeName System.Windows.Forms.OpenFileDialog -Property @{
            InitialDirectory = [Environment]::GetFolderPath('MyComputer')
            Filter           = 'config|config.json'
            Title            = 'Select config.json'
        }

        $syncHash.importChangeLog.Text = $null
        $null = $configSelection.ShowDialog()
        $configPath = $configSelection.fileName
        
        if ($configPath -and (Test-Path $configPath)) {
            $global:importItems = Get-Content $configPath | ConvertFrom-Json    
            $syncHash.importListBox.ItemsSource = $configMap.Keys
            $syncHash.importListBox.SelectAll()
            $syncHash.importDialog.IsOpen = $true
        }
    }
    

    else {
        $syncHash.importChangeLog.Text = $configHash.newVersionStubInfo.changeLog

        $configPath = $configSelection
        $global:importItems = Get-Content $configPath | ConvertFrom-Json    
        $syncHash.importListBox.ItemsSource = $configMap.Keys
        $syncHash.importListBox.SelectAll()
        $syncHash.importDialog.IsOpen = $true
    }           
}

#
function Start-Import {
    param ($ConfigMap, $ImportItems, $ConfigHash, $SavedConfig, $SelectedItems, $baseConfigPath, [bool]$Monitor)

    foreach ($selectedItem in $selectedItems) {  $configHash.($ConfigMap.$selectedItem) = $ImportItems.($ConfigMap.$selectedItem) 
    }
        if (!$Monitor) {
            $ConfigHash.configVer = $null
        }
        else {
            $configHash.configVer = $ImportItems.configVer
        }
     
        Set-Config -ConfigPath $savedConfig -Type Export -ConfigHash $configHash

        $SyncHash.Window.Close()
        Start-Process -WindowStyle Minimized -FilePath "$PSHOME\powershell.exe" -ArgumentList " -ExecutionPolicy Bypass -NonInteractive -File $(Join-Path $baseConfigPath -ChildPath 'v2.ps1')"
        exit
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
            $SyncHash.tabMenu.Items[2].IsEnabled = $true
            $SyncHash.tabMenu.SelectedIndex = 2
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

function Set-DefaultVersionInfo {
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory)][Hashtable]$ConfigHash
    )

    $output = [PSCustomObject]@{
    Ver               = if ($configHash.configVer.Ver) { $configHash.configVer.Ver }
                        else { 1 }
    ID                = if ($configHash.configVer.ID) { $configHash.configVer.ID }
                        else { ([guid]::NewGuid()).Guid }
    configPublishPath = if (![string]::IsNullOrWhiteSpace($configHash.configVer.configPublishPath) -and (Test-Path $configHash.configVer.configPublishPath -ErrorAction SilentlyContinue)) { $configHash.configVer.configPublishPath }
                        else { $null }
    changeLog         = if (![string]::IsNullOrWhiteSpace($configHash.configVer.changeLog)) { $configHash.configVer.changeLog }
                        else { "(No changes recordered)" }

    }

    $output
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

                    ($type + $i + 'HeaderViewBox')     = New-Object -TypeName System.Windows.Controls.ViewBox -Property @{Style = $SyncHash.Window.FindResource('itemHeaderViewBox') }

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
                        Style = $SyncHash.Window.FindResource('actionItemButton')
                        Name  = ($type + $i + 'Box1Action1')

                    }
                   
                    ($type + $i + 'Box1Action2') = New-Object -TypeName System.Windows.Controls.Button -Property @{
                        Style = $SyncHash.Window.FindResource('actionItemButton')
                        Name  = ($type + $i + 'Box1Action2')
                        Tag   = $binding
                    }
            
                    ($type + $i + 'Box1')        = New-Object -TypeName System.Windows.Controls.Button -Property @{
                        Style = $SyncHash.Window.FindResource('itemEditButton')
                        Name  = ($type + $i + 'Box1')
                    }

                }

        #        $binding = New-Object System.Windows.Data.Binding -Property @{
        #            'ElementName'         = ($type + $i + 'TextBox')
        #            'Path'                = 'Text'
        #            'UpdateSourceTrigger' = 'PropertyChanged'   
        #        } 

        #        [void][System.Windows.Data.BindingOperations]::SetBinding($SyncHash.(($type + $i + 'resources')).($type + $i + 'Box1Action1'),[System.Windows.Controls.Button]::TagProperty, $binding)
        #        [void][System.Windows.Data.BindingOperations]::SetBinding($SyncHash.(($type + $i + 'resources')).($type + $i + 'Box1Action2'),[System.Windows.Controls.Button]::TagProperty, $binding)
        
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
                $SyncHash.(($type + $i + 'resources')).($type + $i + 'HeaderViewBox').AddChild($SyncHash.(($type + $i + 'resources')).($type + $i + 'Header'))
                  $SyncHash.(($type + $i + 'resources')).($type + $i + 'DockPanel').AddChild($SyncHash.(($type + $i + 'resources')).($type + $i + 'HeaderViewBox'))
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
        if (($tool.toolActionValid -or $tool.ToolType -eq 'CommandGrid' -or 
            ($tool.ToolType -eq 'Grid' -and $tool.toolSelectValid -and $tool.toolAction -eq '$null' -and $tool.toolTargetFetchCmd -eq '$null')) -and 
            !(($tool.ToolType -match "Select|Grid" -and  $tool.ToolType -ne 'CommandGrid') -and ($tool.toolSelectValid -ne $true -or $tool.toolExtraValid -ne $true))) {
           
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
                        FontSize            = '8'
                        HorizontalAlignment = 'Center'
                        Content             = $tool.ToolName
                    }
              
                    $SyncHash.objectTools.('ctool' + $tool.ToolID + 'LabelVB') = New-Object -TypeName System.Windows.Controls.ViewBox -Property @{
                        Style = $SyncHash.Window.FindResource('itemViewBox') 
                        Margin              = '0,-8,0,0'
                        HorizontalAlignment = 'Center'
                    }
                    
                    $SyncHash.objectTools.('ctool' + $tool.ToolID + 'buttonContent').AddChild($SyncHash.objectTools.('ctool' + $tool.ToolID + 'buttonGlyph'))
                    $SyncHash.objectTools.('ctool' + $tool.ToolID + 'LabelVB').AddChild($SyncHash.objectTools.('ctool' + $tool.ToolID + 'Label1'))
                    $SyncHash.objectTools.('ctool' + $tool.ToolID + 'buttonContent').AddChild($SyncHash.objectTools.('ctool' + $tool.ToolID + 'LabelVB'))
                    
                    


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
                        FontSize            = '8'
                        HorizontalAlignment = 'Center'
                        Content             = $tool.ToolName
                    }
              
                    $SyncHash.objectTools.('utool' + $tool.ToolID + 'LabelVB') = New-Object -TypeName System.Windows.Controls.ViewBox -Property @{
                        Style = $SyncHash.Window.FindResource('itemViewBox') 
                        Margin              = '0,-8,0,0'
                        HorizontalAlignment = 'Center'
                    }
                
                    

                    $SyncHash.objectTools.('utool' + $tool.ToolID + 'buttonContent').AddChild($SyncHash.objectTools.('utool' + $tool.ToolID + 'buttonGlyph'))
                    $SyncHash.objectTools.('utool' + $tool.ToolID + 'LabelVB').AddChild($SyncHash.objectTools.('utool' + $tool.ToolID + 'Label1'))
                    $SyncHash.objectTools.('utool' + $tool.ToolID + 'buttonContent').AddChild($SyncHash.objectTools.('utool' + $tool.ToolID + 'LabelVB'))

        
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
                
                 
                            
                    $SyncHash.objectTools.('tool' + $tool.ToolID + 'LabelVB1') = New-Object -TypeName System.Windows.Controls.ViewBox -Property @{
                        Style = $SyncHash.Window.FindResource('itemViewBox') 
                        HorizontalAlignment = 'Center'
                        StretchDirection = 'DownOnly'
                    }


                    $SyncHash.objectTools.('tool' + $tool.ToolID + 'buttonContent').AddChild($SyncHash.objectTools.('tool' + $tool.ToolID + 'buttonGlyph'))
                    $SyncHash.objectTools.('tool' + $tool.ToolID + 'LabelVB1').AddChild($SyncHash.objectTools.('tool' + $tool.ToolID + 'Label1'))
                    $SyncHash.objectTools.('tool' + $tool.ToolID + 'buttonContent').AddChild($SyncHash.objectTools.('tool' + $tool.ToolID + 'LabelVB1'))
      

                    $SyncHash.objectTools.('tool' + $tool.ToolID) = @{
                        ToolButton = New-Object -TypeName System.Windows.Controls.Button -Property  @{
                            Style   = $SyncHash.Window.FindResource('standAloneButton')
                            Name    = ('tool' + $tool.ToolID)
                            Content = $SyncHash.objectTools.('tool' + $tool.ToolID + 'buttonContent')
                            ToolTip = $tool.toolActionToolTip
                                  
                        }
                    }

                    if ($tool.toolStandAloneCat -and $tool.toolStandAloneCat -notin $syncHash.SADocks.Keys) {
                        if (!$syncHash.SADocks) {
                            $SyncHash.SADocks = @{}
                        }
                        
                        $syncHash.SADocks.($tool.toolStandAloneCat) = @{
                            Dock      = New-Object -TypeName System.Windows.Controls.DockPanel -Property @{
                                Style   = $SyncHash.Window.FindResource('toolsSADock')
                            }
                            Label     = New-Object -TypeName System.Windows.Controls.Label -Property @{
                                Style   = $SyncHash.Window.FindResource('toolsSALabel')
                                Content = $tool.toolStandAloneCat
                            }
                         #   Scroller  = New-Object -TypeName System.Windows.Controls.ScrollViewer -Property @{
                         #       Style   = $SyncHash.Window.FindResource('toolsSAScroll')
                         #   }
                            WrapPanel = New-Object -TypeName System.Windows.Controls.WrapPanel -Property @{
                                Margin  = '25,0,25,0'
                            }
                        }

                        $syncHash.toolParent.AddChild($syncHash.SADocks.($tool.toolStandAloneCat).Dock)
                        $syncHash.SADocks.($tool.toolStandAloneCat).Dock.AddChild($syncHash.SADocks.($tool.toolStandAloneCat).Label)
                       # $syncHash.SADocks.($tool.toolStandAloneCat).Dock.AddChild($syncHash.SADocks.($tool.toolStandAloneCat).Scroller)
                       #$syncHash.SADocks.($tool.toolStandAloneCat).Scroller.AddChild($syncHash.SADocks.($tool.toolStandAloneCat).WrapPanel)
                         $syncHash.SADocks.($tool.toolStandAloneCat).Dock.AddChild($syncHash.SADocks.($tool.toolStandAloneCat).WrapPanel)
                        $syncHash.SADocks.($tool.toolStandAloneCat).WrapPanel.AddChild($SyncHash.objectTools.('tool' + $tool.ToolID).ToolButton)

                    }

                    elseif ($tool.toolStandAloneCat) {
                        $syncHash.SADocks.($tool.toolStandAloneCat).WrapPanel.AddChild($SyncHash.objectTools.('tool' + $tool.ToolID).ToolButton)
                    }
                    else {
                        $SyncHash.standaloneControlPanel.AddChild($SyncHash.objectTools.('tool' + $tool.ToolID).ToolButton)
                    }
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
                                    if ($configHash.objectToolConfig[$toolID - 1].ObjectType -eq 'Standalone') {
                                        $SyncHash.itemToolDialogConfirmObjectName.Text = 'Standalone'
                                    }
                                    $SyncHash.itemToolDialogConfirm.Visibility = 'Visible'
                                    $SyncHash.itemToolDialogConfirmButton.Tag = $toolID
                                    $SyncHash.itemToolDialog.IsOpen = $true                      
                                }

                                else {

                                    $rsArgs = @{
                                        Name            = 'ItemTool'
                                        ArgumentList    = @($toolID, $ConfigHash, $queryHash, $SyncHash.Window, $varHash, $syncHash.adHocConfirmWindow, $syncHash.adHocConfirmText, $SyncHash.snackMsg.MessageQueue)
                                        ModulesToImport = $configHash.modList
                                    }

                                    Start-RSJob @rsArgs -ScriptBlock {
                                        Param($toolID, $ConfigHash, $queryHash, $window, $varHash, $confirmWindow, $textBlock, $queue)

                                        $toolName = ($ConfigHash.objectToolConfig[$toolID - 1].toolName).ToUpper()                                     
                                        Set-CustomVariables -VarHash $varHash
                                        if ($configHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {Remove-Variable ActiveObject, ActiveObjectType, ActiveObjectData -ErrorAction SilentlyContinue}

                                        try {             
                                            ([scriptblock]::Create($ConfigHash.objectToolConfig[$toolID - 1].toolAction)).Invoke()         

                                            if ($ConfigHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {
                                                Write-SnackMsg -Queue $queue -ToolName $toolName -Status Success 
                                                Write-LogMessage -syncHashWindow $window -Path $ConfigHash.actionlogPath -Message Succeed -ActionName $toolName -ArrayList $ConfigHash.actionLog 
                                            }

                                            else {
                                                Write-SnackMsg -Queue $queue -ToolName $toolName -SubjectName $activeObject -Status Success 
                                                Write-LogMessage -syncHashWindow $window -Path $ConfigHash.actionlogPath -Message Succeed -SubjectName $activeObject -SubjectType $activeObjectType -ActionName $toolName -ArrayList $ConfigHash.actionLog 
                                            }
                                        }
                                        catch {
                                            if ($ConfigHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {
                                                Write-SnackMsg -Queue $queue -ToolName $toolName -Status Fail 
                                                Write-LogMessage -syncHashWindow $window -Path $ConfigHash.actionlogPath -Message Fail -ActionName $toolName -ArrayList $ConfigHash.actionLog -Error $_
                                            }
                                            else {
                                                Write-SnackMsg -Queue $queue -ToolName $toolName -SubjectName $activeObject -Status Fail 
                                                Write-LogMessage -syncHashWindow $window -Path $ConfigHash.actionlogPath -Message Succeed -SubjectName $activeObject -ActionName $activeObjectType -SubjectType $targetType -ArrayList $ConfigHash.actionLog -Error $_
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
                        

                            if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionSelectCustom -eq $false) {
                                    $SyncHash.ItemToolADSelectionPanel.Visibility = 'Collapsed'

                                    $rsVars = @{
                                        target     = $ConfigHash.currentTabItem
                                        targetType = $queryHash[$ConfigHash.currentTabItem].ObjectClass
                                    }

                                    $rsArgs = @{
                                        Name            = 'PopulateListboxNoAD'
                                        ArgumentList    = @($ConfigHash, $SyncHash, $toolID, $rsVars, $syncHash.adHocConfirmWindow, $syncHash.Window, $syncHash.adHocConfirmText, $varHash)
                                        ModulesToImport = $configHash.modList
                                    }

                                    Start-RSJob @rsArgs -ScriptBlock {
                                        param($ConfigHash, $SyncHash, $toolID, $rsVars, $confirmWindow, $window, $textBlock, $varHash)
                            
                                        $target = $rsVars.target
                                        $targetType = $rsVars.targetType
                                                      
                                        Set-CustomVariables -VarHash $varHash

                                        $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.itemTooListBoxProgress.Visibility = 'Visible' })
                    
                                        $list = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
                                        ([scriptblock]::Create($ConfigHash.objectToolConfig[$toolID - 1].toolFetchCmd)).Invoke() | ForEach-Object { $list.Add([PSCUstomObject]@{'Name' = $_ }) }
                                 
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
                                  
                                    if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionSelectCustom) {
                                        $syncHash.itemToolADSelectionButton.Tag = $ConfigHash.objectToolConfig[$toolID - 1].toolActionCustomSelect
                                        $SyncHash.ItemToolADSelectionPanel.Visibility = 'Visible'                            
                                        $SyncHash.itemToolListSelect.Visibility = 'Visible'
                                        $SyncHash.itemToolListSelectListBox.ItemsSource = $null
                                    }
                                }                          
                          
                            }
                            'Grid' {
                                $SyncHash.itemToolDialog.Title = $ConfigHash.objectToolConfig[$toolID - 1].toolName
                                $SyncHash.itemToolGridItemsGrid.ItemsSource = $null
                                $SyncHash.itemToolGridSelectConfirmButton.Tag = $toolID 
                                $SyncHash.itemToolGridSelectText.Text = $ConfigHash.objectToolConfig[$toolID - 1].toolDescription
                                $SyncHash.itemToolGridSelect.Visibility = 'Visible'
                                $SyncHash.itemToolDialog.IsOpen = $true  

                        

                                if ($ConfigHash.objectToolConfig[$toolID - 1].toolAction -eq '$null' -and $ConfigHash.objectToolConfig[$toolID - 1].toolTargetFetchCmd -eq '$null') {
                                    $syncHash.itemToolGridSelectConfirmButton.Visibility = "Collapsed" 
                                    $syncHash.itemToolGridSelectConfirmCancel.Content = 'Close'
                                }

                                else {
                                    $syncHash.itemToolGridSelectConfirmButton.Visibility = "Visible" 
                                    $syncHash.itemToolGridSelectConfirmCancel.Content = 'Cancel'
                                }

                                if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionSelectCustom -eq $false) {

                                    $SyncHash.itemToolGridADSelectionPanel.Visibility = 'Collapsed'
                            
                                    if ($ConfigHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {

                                    $rsVars = @{
                                            target     = '[N/A]'
                                            targetType = 'Standalone'
                                        }

                                    }

                                    else {

                                        $rsVars = @{
                                            target     = $ConfigHash.currentTabItem
                                            targetType = $queryHash[$ConfigHash.currentTabItem].ObjectClass
                                        }

                                    }

                                    $rsArgs = @{
                                        Name            = 'PopulateGridbox'
                                        ArgumentList    = @($ConfigHash, $SyncHash, $toolID, $rsVars, $syncHash.adHocConfirmWindow, $syncHash.Window, $syncHash.adHocConfirmText, $varHash)
                                        ModulesToImport = $configHash.modList
                                    }

                                    Start-RSJob @rsArgs -ScriptBlock {
                                        param($ConfigHash, $SyncHash, $toolID, $rsVars, $confirmWindow, $window, $textBlock, $varHash)

                                        $target = $rsVars.target
                                        $targetType = $rsVars.targetType
                                        
                                        Set-CustomVariables -VarHash $varHash
                                                      
                                        $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.itemToolGridProgress.Visibility = 'Visible' })
                                
                                        $list = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
                                        ([scriptblock]::Create($ConfigHash.objectToolConfig[$toolID - 1].toolFetchCmd)).Invoke() | ForEach-Object { $list.Add($_) }
                                                                      
                                        $SyncHash.Window.Dispatcher.Invoke([Action] {
                                                $SyncHash.itemToolGridItemsGrid.ItemsSource = [System.Windows.Data.ListCollectionView]$list
                                                $SyncHash.itemToolGridItemsGrid.ItemsSource.IsLiveSorting = $true
                                                $SyncHash.itemToolGridProgress.Visibility = 'Collapsed'
                                                $SyncHash.itemToolGridItemsGrid.Items.Refresh()

                                                if ($SyncHash.itemToolGridItemsGrid.HasItems) {

                                                    if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionMultiSelect -and
                                                        ($null -ne $ConfigHash.objectToolConfig[$toolID - 1].toolAction)) {

                                                        $SyncHash.itemToolGridItemsGrid.SelectionMode = 'Extended'
                                                        $SyncHash.itemToolGridSelectAllButton.Visibility = 'Visible'
                                                    }

                                                    else {
                                                        $SyncHash.itemToolGridItemsGrid.SelectionMode = 'Single'
                                                        $SyncHash.itemToolGridSelectAllButton.Visibility = 'Collapsed'
                                                    }

                                                    if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionExportable) { $syncHash.itemToolGridExport.Visibility = 'Visible'}
                                                    
                                                    else {$syncHash.itemToolGridExport.Visibility = 'Collapsed'}

                                                     $syncHash.itemToolGridItemsEmptyText.Visibility = 'Hidden'

                                                }

                                                else { $syncHash.itemToolGridItemsEmptyText.Visibility = 'Visible' }

                                            


                                            })
                                    }
                                }

                                else {
                                    if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionSelectCustom) {
                                        $syncHash.itemToolGridADSelectionButton.Tag = $ConfigHash.objectToolConfig[$toolID - 1].toolActionCustomSelect                                                                       
                                        $SyncHash.itemToolGridADSelectionPanel.Visibility = 'Visible'                           
                                        $SyncHash.itemToolGridSelect.Visibility = 'Visible'
                                        $SyncHash.itemToolGridItemsGrid.ItemsSource = $null
                                    }
                                }  
                            
                               
                            }
                            'CommandGrid' {
                                $SyncHash.itemToolDialog.Title = $ConfigHash.objectToolConfig[$toolID - 1].toolName
                                $SyncHash.itemCommandGridText.Text = $ConfigHash.objectToolConfig[$toolID - 1].toolDescription
                                $SyncHash.itemToolCommandGridPanel.Visibility = 'Visible'
                                $SyncHash.itemToolCommandGridDataGrid.ItemsSource = $null
                                $SyncHash.toolsCommandGridExecuteAllPanel.Visibility = 'Collapsed'        
                                $SyncHash.toolsCommandGridExecuteAll.Tag = 'False'                  

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
                                                    Result          = (Invoke-Expression $item.queryCmd).ToString()
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

function New-CustomLogHeader {
    param
    (
        [Object]
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Data to process")]
        $InputObject,
        $SyncHash,
        $ConfigHash,
        [Parameter(Mandatory = $true)][ValidateSet('User', 'Comp')]$Type
    )

    begin {
        $customField = 0

        if ($type -eq 'User') {

            $grid = 'UserCompGrid'
            $panel = 'UserLogExtraPropGrid'
        }

        else {
            $grid = 'CompUserGrid'
            $panel = 'CompLogExtraPropGrid'
        }          
    }

    process {
        
        # populate custom dock items
        $customField++
        $syncHash.Window.Dispatcher.Invoke([Action] {
                $syncHash.('customPropDock' + $customField) = New-Object System.Windows.Controls.StackPanel
                $syncHash.('customPropLabel' + $customField) = New-Object System.Windows.Controls.Label
                $syncHash.('customPropText' + $customField) = New-Object System.Windows.Controls.Textbox
                $syncHash.('customPropLabel' + $customField).Content = $InputObject.Header
                $syncHash.('customPropDock' + $customField).VerticalAlignment = 'Top'
                $syncHash.('customPropDock' + $customField).Margin = '0,-10,0,0'
                        
        
                $syncHash.('customPropDock' + $customField).AddChild(($syncHash.('customPropLabel' + $customField)))
                $syncHash.('customPropDock' + $customField).AddChild(($syncHash.('customPropText' + $customField)))
             
        
                $syncHash.('customPropLabel' + $customField).FontSize = '10'
                $syncHash.('customPropLabel' + $customField).Foreground = $syncHash.Window.FindResource('MahApps.Brushes.SystemControlBackgroundBaseMediumLow')
                $syncHash.('customPropText' + $customField).Style = $syncHash.Window.FindResource('compItemBox') 
                $syncHash.$panel.AddChild(($syncHash.('customPropDock' + $customField)))
        
                # Create and set a binding on the textbox object
                $Binding = New-Object System.Windows.Data.Binding
                $Binding.UpdateSourceTrigger = 'PropertyChanged'
                $Binding.Source = $syncHash.$grid
                $Binding.Path = "SelectedItem.$($InputObject.Header)"
                $Binding.Mode = [System.Windows.Data.BindingMode]::TwoWay
             
        
                [void][System.Windows.Data.BindingOperations]::SetBinding(($syncHash.('customPropText' + $customField)), [System.Windows.Controls.TextBox]::TextProperty, $Binding)
            })
                            
    }
}

function Set-ItemControlPanelSize {
    param
    (
        $SyncHash,
        $ConfigHash
    )

    @('User', 'Comp') | ForEach-Object {

        if ($_ -eq 'User') { $rowName = 'userCompControlRow' }
        else { $rowName = 'compUserControlRow' }

        $sizeToExpandFromCustomLogEntries = ([math]::Ceiling(($configHash.($_ + 'LogMapping') | Where-Object -FilterScript { $_.FieldSel -ne 'Ignore' }).Count / 5) - 1) * 57.5 
        $sizeToExpandFromExtraContextButtons = ([math]::Ceiling(($configHash.contextConfig | Where-Object { $_.ValidAction -eq $true }).Count / 8) - 1) * 30
        
        $newSize = $sizeToExpandFromCustomLogEntries + $sizeToExpandFromExtraContextButtons + $syncHash.$rowName.Height.Value
        
        $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.$rowName.Height = $newSize })
    
    }
}



function Start-BasicADCheck {
    param ($SysCheckHash, $ConfigHash) 

    if ((Get-WmiObject -Class Win32_ComputerSystem).PartofDomain) {               
        $SysCheckHash.sysChecks[0].ADMember = 'True'
                    
        if ($SysCheckHash.sysChecks[0].ADModule -eq $true) {                                                 
            $selectedDC = Get-ADDomainController -Discover -Service ADWS -ErrorAction SilentlyContinue 
            
            if (Test-Connection -Count 1 -Quiet -ComputerName $selectedDC.HostName) { 
                $global:PSDefaultParameterValues.Add("*-AD*:Server",$selectedDC.HostName[0])            
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
        
        #Get-AdObjectPropertyList -ConfigHash $ConfigHash        
                
        $ConfigHash.($type + 'PropPullListNames') = [System.Collections.ArrayList]@()
        $ConfigHash.($type + 'PropPullListNames').Add('Non-AD Property') 
        $ConfigHash.($type + 'PropPullList').Name | ForEach-Object -Process { $ConfigHash.($type + 'PropPullListNames').Add($_) }
    }
}

function Set-ADGenericQueryNames {
    param($ConfigHash) 

    foreach ($id in ($ConfigHash.queryDefConfig.ID)) { $ConfigHash.queryDefConfig[$id - 1].QueryDefTypeList = $ConfigHash.adPropertyMap.Keys | Sort-Object }
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
                        
                        actionCmd1CanOff  = if ($value = ($tempList | Where-Object -FilterScript { $_.Field -eq $i }).actionCmd1CanOff) { $value } 
                        else { $false }       

                        actionCmd1OffStr  = if ($value = ($tempList | Where-Object -FilterScript { $_.Field -eq $i }).actionCmd1OffStr) { $value } 
                        else { $null }                  
                                    
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
                         
                        actionCmd2CanOff  = if ($value = ($tempList | Where-Object -FilterScript { $_.Field -eq $i }).actionCmd2CanOff) { $value } 
                        else { $false }       

                        actionCmd2OffStr  = if ($value = ($tempList | Where-Object -FilterScript { $_.Field -eq $i }).actionCmd2OffStr) { $value } 
                        else { $null }     
                                    
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
        'User' { $ConfigHash.varData.UpdateUser = $true } 

        'Comp' { $ConfigHash.varData.UpdateComp = $true }
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

function Find-UpdatedConfig {
    param ($configHash, $syncHash)
    Start-Sleep -Seconds 3
    if (![string]::IsNullOrEmpty($configHash.configVer.configPublishPath)) {
        $configVerPath = Join-Path -Path $configHash.configVer.configPublishPath -ChildPath 'configVer.json'
        if (Test-Path  $configVerPath) {
            $configVerInfo = Get-Content $configVerPath | ConvertFrom-Json
            $configHash.newVersionStubInfo = $configVerInfo
            if ($configVerInfo.Ver -gt $configHash.configVer.ver -and $configVerInfo.ID -eq $configHash.ConfigVer.ID) {
                $syncHash.Window.Dispatcher.Invoke([Action]{$syncHash.headerConfigUpdate.Visibility = "Visible" })
            }
        }
    }   
}

function Start-VarUpdater {
    [CmdletBinding()]
    param ($ConfigHash, $varHash, $queryHash, $syncHash)
    
    $rsArgs = @{
        Name            = 'VarUpdater'
        ArgumentList    = @($ConfigHash, $varHash, $queryHash, $syncHash )
        ModulesToImport = $configHash.modList
    }
    
    $configHash.IsClosed    = $false
    $configHash.IsSearching = $false

    Start-RSJob @rsArgs -ScriptBlock {
        param($ConfigHash, $varHash, $queryHash, $syncHash)

        $startTime = Get-Date
        $first = $true
        do {
            Set-DurationVarsToUpdate -ConfigHash $ConfigHash -StartTime $startTime

            if ($first -eq $true) {
                $ConfigHash.varData.UpdateMinute = $true
                $ConfigHash.varData.UpdateHour = $true
                $ConfigHash.varData.UpdateDay = $true

                $ConfigHash.varListConfig | Where-Object { $_.UpdateFrequency -eq 'Program Start' } |
                    ForEach-Object { $varHash.($_.VarName) = ([scriptblock]::Create($_.VarCmd)).Invoke() }

                $first = $false
            }

            if ($ConfigHash.varData.ContainsValue($true)) {
                $keyList = [array]($ConfigHash.varData.Keys)
                foreach ($varInfo in $keyList) {
                    if ($ConfigHash.varData.$varInfo -eq $true) {
                        switch ($varInfo) {
                            'UpdateMinute' {
                                $ConfigHash.varListConfig |
                                    Where-Object { $_.UpdateFrequency -eq 'Every 15 mins' } |
                                        ForEach-Object { $varHash.($_.VarName) = ([scriptblock]::Create($_.VarCmd)).Invoke() }
                                $ConfigHash.varData.$varInfo = $false
                                break
                            }
                            'UpdateHour' {
                                $ConfigHash.varListConfig |
                                    Where-Object { $_.UpdateFrequency -eq 'Hourly' } |
                                        ForEach-Object { $varHash.($_.VarName) = ([scriptblock]::Create($_.VarCmd)).Invoke() }
                                $ConfigHash.varData.$varInfo = $false
                                Find-UpdatedConfig -ConfigHash $configHash -SyncHash $syncHash
                                break
                            }                                       
                            'UpdateDay' {
                                $ConfigHash.varListConfig |
                                    Where-Object { $_.UpdateFrequency -eq 'Daily' } |
                                        ForEach-Object { $varHash.($_.VarName) = ([scriptblock]::Create($_.VarCmd)).Invoke() }
                                $ConfigHash.varData.$varInfo = $false
                                break
                            }
                            'UpdateUser' { 
                                $ConfigHash.varListConfig |
                                    Where-Object { $_.UpdateFrequency -match 'User Queries|All Queries' } |
                                        ForEach-Object { $varHash.($_.VarName) = ([scriptblock]::Create($_.VarCmd)).Invoke() }
                            }
                            'UpdateComp' {
                                $ConfigHash.varListConfig |
                                    Where-Object { $_.UpdateFrequency -match 'Comp Queries|All Queries' } |
                                        ForEach-Object { $varHash.($_.VarName) = ([scriptblock]::Create($_.VarCmd)).Invoke() }
                            }
                            {$_ -match "UpdateComp|UpdateUser"} { 
                                $varHash.ActiveObject     = $configHash.currentTabItem 
                                $varHash.ActiveObjectType = $queryHash[$configHash.currentTabItem].ObjectClass -replace 'u', 'U' -replace 'c', 'C'
                                $varHash.ActiveObjectData = $queryHash[$configHash.currentTabItem]                             
                                $ConfigHash.varData.$varInfo = $false 
                            }
                        }
                    }
                }
            }

            Start-Sleep -Milliseconds 500

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

        $rsArgs = @{
            ArgumentList = @($syncHash, $syncHash[($SyncHash.Keys.Where({ $_ -like 'setting*FlyoutScroller'}))])
            Name = 'FlyoutScrollerReset'
        }

        Start-RSJob @rsArgs -ScriptBlock {
            param ($syncHash, $scrollerList)
            Start-Sleep -Milliseconds 500
            $syncHash.Window.Dispatcher.Invoke([action]{           
            foreach ($scroller in $scrollerlist) { $scroller.ScrollToTop() }
            })
       }
    }

    if ($Title) { Set-ChildWindow -SyncHash $SyncHash -Title $Title }

    if (!($SkipResize)) { Set-ChildWindow -SyncHash $SyncHash -Width 400 -Height 215 }
    
    Set-ChildWindow -SyncHash $SyncHash -Background Standard
}

#endregion

#region itemToolFunctions

function Suspend-ToolControls {
    param (
        $ConfigHash,
        $SyncHash,
        $ToolID,
        [Parameter(Mandatory)][ValidateSet('ListBox', 'Grid')][string]$Control
    ) 
 
    if ($Control -eq 'Grid') {
 
        if ($SyncHash.itemToolGridItemsGrid.HasItems) {

            if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionMultiSelect -and ($null -ne $ConfigHash.objectToolConfig[$toolID - 1].toolAction)) {
                $SyncHash.itemToolGridItemsGrid.SelectionMode = 'Extended'
                $SyncHash.itemToolGridSelectAllButton.Visibility = 'Visible'
            }

            else {
                $SyncHash.itemToolGridItemsGrid.SelectionMode = 'Single'
                $SyncHash.itemToolGridSelectAllButton.Visibility = 'Collapsed'
            }

            if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionExportable) { $syncHash.itemToolGridExport.Visibility = 'Visible'}
                                                    
            else {$syncHash.itemToolGridExport.Visibility = 'Collapsed'}

            $syncHash.itemToolGridItemsEmptyText.Visibility = 'Hidden'

        }

        else { $syncHash.itemToolGridItemsEmptyText.Visibility = 'Visible' }

        $SyncHash.itemToolGridProgress.Visibility = 'Collapsed'

    }

    else {
        
        if ($SyncHash.itemToolListSelectListBox.HasItems) {

            if ($ConfigHash.objectToolConfig[$toolID - 1].toolActionMultiSelect -and ($null -ne $ConfigHash.objectToolConfig[$toolID - 1].toolAction)) {
                $SyncHash.itemToolListSelectListBox.SelectionMode = 'Multiple'
                $SyncHash.itemToolListSelectAllButton.Visibility = 'Visible'
            }

            else {
                $SyncHash.itemToolListSelectListBox.SelectionMode = 'Single'
                $SyncHash.itemToolListSelectAllButton.Visibility = 'Collapsed'
            }

            $syncHash.itemToolListItemsEmptyText.Visibility = 'Hidden'

        }

        else { $syncHash.itemToolListItemsEmptyText.Visibility = 'Visible' }

        $SyncHash.itemTooListBoxProgress.Visibility = 'Collapsed'

    }
}

function Start-CustomItemSelection {
    param (
            $SyncHash, 
            $ConfigHash, 
            [Parameter(Mandatory)][ValidateSet('ListBox', 'Grid')][string]$Control
        )
    
    if ($control -eq 'ListBox') { $switchValue = $syncHash.ItemToolADSelectionButton.Tag}
    else { $switchValue = $syncHash.ItemToolGridADSelectionButton.Tag }


    switch ($switchValue) {
        'AD Object (any)' { Get-CustomItem -ConfigHash $configHash -SyncHash $syncHash -Control $Control -Type AD -Scope All}
        'AD User' { Get-CustomItem -ConfigHash $configHash -SyncHash $syncHash -Control $Control -Type AD -Scope Users }
        'AD Computer' { Get-CustomItem -ConfigHash $configHash -SyncHash $syncHash -Control $Control -Type AD -Scope Computers}
        'AD Group' { Get-CustomItem -ConfigHash $configHash -SyncHash $syncHash -Control $Control -Type AD -Scope Groups }
        'OU' { Get-CustomItem -ConfigHash $configHash -SyncHash $syncHash -Control $Control -Type OU }
        'String' { Get-CustomItemBox -ConfigHash $configHash -SyncHash $syncHash -Control $Control -Type String}
        'Integer' { Get-CustomItemBox -ConfigHash $configHash -SyncHash $syncHash -Control $Control -Type Int }
        'Choice' { Get-CustomItemBox -ConfigHash $configHash -SyncHash $syncHash -Control $Control -Type Choice }
        'File' { Get-CustomItem -ConfigHash $configHash -SyncHash $syncHash -Control $Control -Type File }
        'Directory' { Get-CustomItem -ConfigHash $configHash -SyncHash $syncHash -Control $Control -Type Folder }
    }
}



function Set-SelectedItem {
     param(
        [Parameter(Mandatory)][hashtable]$ConfigHash, 
        [Parameter(Mandatory)][hashtable]$SyncHash,
        [Parameter(Mandatory)][ValidateSet('ListBox', 'Grid')][string]$Control,
        [Parameter(Mandatory)][string]$InputObject,
        [Parameter(Mandatory)]$ToolID
        )

    $SyncHash.Window.Dispatcher.Invoke([Action] {
            if ($Control -eq 'ListBox') { 
                $SyncHash.itemToolADSelectedItem.Content = $inputObject
                $SyncHash.itemTooListBoxProgress.Visibility = 'Visible'
            }
            else {
                $SyncHash.itemToolGridADSelectedItem.Content = $inputObject
                $SyncHash.itemToolGridProgress.Visibility = 'Visible'
            }
        })
     
    $list = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
        
    if ($Control -eq 'ListBox') { ([scriptblock]::Create($ConfigHash.objectToolConfig[$toolID - 1].toolFetchCmd)).Invoke() | ForEach-Object { $list.Add([PSCustomObject]@{'Name' = $_ }) } }
    else { ([scriptblock]::Create($ConfigHash.objectToolConfig[$toolID - 1].toolFetchCmd)).Invoke() | ForEach-Object { $list.Add($_) } }

    $SyncHash.Window.Dispatcher.Invoke([Action] {
        if ($Control -eq 'ListBox') {
            if (($list.Name | Measure-Object).Count -ge 1) { $SyncHash.itemToolListSelectListBox.ItemsSource = [System.Windows.Data.ListCollectionView]$list }
                $SyncHash.itemTooListBoxProgress.Visibility = 'Collapsed'
            

            Suspend-ToolControls -ConfigHash $configHash -SyncHash $syncHash -ToolID $toolID -Control ListBox

        }
        else { 
            if (($list | Measure-Object).Count -ge 1) { $SyncHash.itemToolGridItemsGrid.ItemsSource = [System.Windows.Data.ListCollectionView]$list }   
            $SyncHash.itemToolGridProgress.Visibility = 'Collapsed' 
            Suspend-ToolControls -ConfigHash $configHash -SyncHash $syncHash -ToolID $toolID -Control Grid

        }
    })

    
}

function  Set-ExternalTools {
param ($configHash, $baseConfigPath)

    $extList = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
    $toolDir = Join-Path -Path $baseConfigPath -ChildPath 'Tools'

    foreach ($remoteTool in $configHash.rtConfig.Keys) {
        $extList.Add([PSCustomObject]@{
            Name = $configHash.rtConfig.$remoteTool.DisplayName
            Icon = [Convert]::FromBase64String($ConfigHash.rtConfig.$remoteTool.Icon)
            Exe  = $configHash.rtConfig.$remoteTool.Path
        })
    }

    if (Test-Path -Path $toolDir) {
        $toolList = Get-ChildItem $toolDir -File

        foreach ($tool in $toolList) {
            $extList.Add([PSCustomObject]@{
                Name = [IO.Path]::GetFileNameWithoutExtension($tool.FullName)
                Icon = [Convert]::FromBase64String((Get-Icon -Path $tool.FullName -ToBase64))
                Exe  = $tool.FullName
            })
        }
    }

    $extList
}

function Get-CustomItem {
    param(
        [Parameter(Mandatory)][hashtable]$ConfigHash, 
        [Parameter(Mandatory)][hashtable]$SyncHash,
        [Parameter(Mandatory)][ValidateSet('ListBox','Grid')][string]$Control,
        [Parameter(Mandatory)][ValidateSet('AD','OU','File','Folder')][string]$Type,
        [ValidateSet('All', 'Users','Computers','Groups')][string]$Scope
    ) 

    if ($Control -eq 'Grid') {  $toolID = $SyncHash.itemToolGridSelectConfirmButton.Tag  }
    else { $toolID = $SyncHash.itemToolListSelectConfirmButton.Tag }

    $syncHash.itemToolGridExport.Visibility = 'Collapsed'
    $syncHash.itemToolListSelectAllButton.Visibility = 'Collapsed'
    $syncHash.itemToolGridSelectAllButton.Visibility = 'Collapsed'


    $rsArgs = @{
        Name            = ('populate' + $Control)
        ArgumentList    = @($ConfigHash, $SyncHash, $toolID, $Control, $varHash, $Type, $scope)
        ModulesToImport = $configHash.modList
    }

    Start-RSJob @rsArgs -ScriptBlock {
        param($ConfigHash, $SyncHash, $toolID, $Control, $varHash, $Type, $Scope)

        Set-CustomVariables -VarHash $varHash
        
       
        $SyncHash.Window.Dispatcher.Invoke([Action] {
                if ($Control -eq 'ListBox') { $SyncHash.itemToolListSelectListBox.ItemsSource = $null  }
                else { $SyncHash.itemToolGridItemsGrid.ItemsSource = $null  }
            })

        $inputObject = switch ($Type) {
            'AD'     { (Select-ADObject -Type $Scope).FetchedAttributes -replace '$'}
            'OU'     { (Choose-ADOrganizationalUnit -HideNewOUFeature $true).DistinguishedName }
            'File'   { New-FileDialog -InitialDirectory 'C:\' -Header "Select file" } 
            'Folder' { New-FolderSelection }

        }
     
        if ($inputObject) { Set-SelectedItem -ConfigHash $configHash -SyncHash $syncHash -Control $control -InputObject $inputObject -ToolID $toolID }
    }      
}

function Get-CustomItemBox {
    param(
        [Parameter(Mandatory)][hashtable]$ConfigHash, 
        [Parameter(Mandatory)][hashtable]$SyncHash,
        [Parameter(Mandatory)][ValidateSet('ListBox', 'Grid')][string]$Control,
        [Parameter(Mandatory)][ValidateSet('String', 'Int', 'Choice')][string]$Type) 

    if ($Control -eq 'Grid') { $toolID = $SyncHash.itemToolGridSelectConfirmButton.Tag }
    else { $toolID = $SyncHash.itemToolListSelectConfirmButton.Tag }

    $syncHash.itemToolGridExport.Visibility = 'Collapsed'
    $syncHash.itemToolListSelectAllButton.Visibility = 'Collapsed'
    $syncHash.itemToolGridSelectAllButton.Visibility = 'Collapsed'
    $syncHash.itemToolCustomContent.Visibility = 'Collapsed'
    $syncHash.itemToolCustomContentChoice.Visibility = 'Collapsed'

    if ($type -eq 'int') { 
        $syncHash.customInputLabel.Content = "Paste or insert number required for query"
        $syncHash.itemToolCustomContent.Tag = 'int' 
        $syncHash.itemToolCustomContent.Visibility = 'Visible'
    }
    elseif ($type -eq 'String')  { 
        $syncHash.customInputLabel.Content = "Paste or insert text required for query"
        $syncHash.itemToolCustomContent.Tag = 'string' 
        $syncHash.itemToolCustomContent.Visibility = 'Visible'
    }
    else {
        $syncHash.customInputLabel.Content = "Select option from list"
        $syncHash.itemToolCustomContent.Tag = 'choice' 
        $syncHash.itemToolCustomContentChoice.Visibility = 'Visible'
    }

    $rsArgs = @{
        Name            = ('populate' + $Control)
        ArgumentList    = @($ConfigHash, $SyncHash, $toolID, $Control, $varHash, $type)
        ModulesToImport = $configHash.modList
    }



    Start-RSJob @rsArgs -ScriptBlock {
        param($ConfigHash, $SyncHash, $toolID, $Control, $varHash, $type)

        Set-CustomVariables -VarHash $varHash
       
       
        $configHash.customDialogClosed = $false
        $configHash.customInput = $null

        $SyncHash.Window.Dispatcher.Invoke([Action] {
            if ($Control -eq 'ListBox') { $SyncHash.itemToolListSelectListBox.ItemsSource = $null }
            else { $SyncHash.itemToolGridItemsGrid.ItemsSource = $null }
        })

        if ($type -ne 'choice') {
            $SyncHash.Window.Dispatcher.Invoke([Action] {
                $syncHash.itemToolCustomContent.Text = $null
                $syncHash.itemToolCustomDialog.Visibility = 'Visible'
                $syncHash.itemToolCustomDialog.IsOpen = $true
            })
        }

        else {
            $choiceList = $ConfigHash.objectToolConfig[$toolID - 1].toolActionSelectChoice -replace ', ', ',' -split ','
           
             $SyncHash.Window.Dispatcher.Invoke([Action] {
                $syncHash.itemToolCustomContentChoice.ItemsSource = $choiceList
                $syncHash.itemToolCustomContentChoice.SelectedIndex = 0
                $syncHash.itemToolCustomDialog.Visibility = 'Visible'
                $syncHash.itemToolCustomConfirm.IsEnabled = $true
                $syncHash.itemToolCustomDialog.IsOpen = $true
            })
        }

        do { } until ($configHash.customDialogClosed -eq $true)
        
        $SyncHash.Window.Dispatcher.Invoke([Action] {$syncHash.itemToolCustomDialog.IsOpen = $false })
        Start-Sleep -Seconds 1
        $inputObject = $configHash.customInput

        if ($inputObject) { Set-SelectedItem -ConfigHash $configHash -SyncHash $syncHash -Control $control -InputObject $inputObject -ToolID $toolID }
      
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
        ArgumentList    = @($ConfigHash, $ItemList, $SyncHash.snackMsg.MessageQueue, $SyncHash.('itemTool' + $Control + 'SelectConfirmButton').Tag, $SyncHash.Window, $varHash)
        ModulesToImport = $ConfigHash.modList
    }

    Start-RSJob @rsArgs -ScriptBlock {
        param($ConfigHash, $ItemList, $queue, $toolID, $window, $varHash) 

        $toolName = $ConfigHash.objectToolConfig[$toolID - 1].toolName
        Set-CustomVariables -VarHash $varHash

        try {
            
             Invoke-Expression $ConfigHash.objectToolConfig[$toolID - 1].toolTargetFetchCmd
           
             
            foreach ($selectedItem in $ItemList) { ([scriptblock]::Create($ConfigHash.objectToolConfig[$toolID - 1].toolAction)).Invoke() }
            
            if ($ConfigHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {
                Write-SnackMsg -Queue $queue -ToolName $toolName -Status Success 
                Write-LogMessage -syncHashWindow $window -Path $ConfigHash.actionlogPath -Message Succeed -ActionName $toolName -SubjectType 'Standalone' -ArrayList $ConfigHash.actionLog
            }
            
            else {
                Write-SnackMsg -Queue $queue -ToolName $toolName -SubjectName $activeObject -Status Success 
                Write-LogMessage -syncHashWindow $window -Path $ConfigHash.actionlogPath -Message Succeed -SubjectName $activeObject -SubjectType $activeObjectType -ActionName $toolName -ArrayList $ConfigHash.actionLog 
            }
        }
        
        catch {
            if ($ConfigHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {
                Write-SnackMsg -Queue $queue -ToolName $toolName -Status Fail 
                Write-LogMessage -syncHashWindow $window -Path $ConfigHash.actionlogPath -Message Fail -ActionName $toolName -ArrayList $ConfigHash.actionLog -Error $_
            }
            else {
                Write-SnackMsg -Queue $queue -ToolName $toolName -SubjectName $activeObject -Status Fail 
                Write-LogMessage -syncHashWindow $window -Path $ConfigHash.actionlogPath -Message Fail -SubjectName $activeObject -SubjectType $activeObjectType  -ActionName $toolName -ArrayList $ConfigHash.actionLog -Error $_
            }
        }
    }

    $SyncHash.itemToolDialog.IsOpen = $false
}

function New-DialogWait {
    param($confirmWindow, $window, $configHash, $textBlock, $text)

    $window.Dispatcher.Invoke([action]{
        
        if ($text -and $textBlock) { $textBlock.Text = $text }
        else {  $textBlock.Text = 'Continue with action?' }
        
        $confirmWindow.IsOpen = $true
        
    })

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

    $SyncHash.settingRemoteListTypes.ItemsSource = ([array]($ConfigHash.nameMapList | Select-Object Name))

    switch ($SyncHash.settingRALabel.Text) {
            
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
            $rtID = 'rt' + [string]($SyncHash.settingRALabel.Text -replace '.[A-Z]* ')

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

    switch ($SyncHash.settingRALabel.Text) {
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
            $rtID = 'rt' + [string]($SyncHash.settingRALabel.Text -replace '.[A-Z]* ')
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

    $rtID = 'rt' + [string]($SyncHash.settingRALabel.Text -replace '.[A-Z]* ')

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
    foreach ($rtID in ($ConfigHash.rtConfig.Keys | Where-Object { $_ -like "RT*" } | Sort-Object)) { New-CustomRTConfigControls -ConfigHash $ConfigHash -SyncHash $SyncHash -RTID $rtID }
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

function Get-NetIPLocationList {
    Begin {
        $ipArray = [System.Collections.ArrayList]@()
        $adSites = Get-ADReplicationSite -Filter *
    }

    Process {
        foreach ($adSite in $adSites) {
            (Get-ADReplicationSubnet -Filter "Site -eq ""$($adSite.DistinguishedName)""" -Properties Name, Location, Description) | ForEach-Object {
                
                $netAddress = $_.Name -split '/'

                    $ipArray.Add([PSCustomObject]@{
                    
                        Location = if ($_.Location) { $_.Location }
                                   elseif ($_.Description) { ($_.Description) } 
                                   elseif ($adSite.Description) { $adSite.Description }
                                   else { 'Unknown' }

                        Network = $netAddress[0]

                        Mask    = $netAddress[1]

                    }) | Out-Null

            }
        }
    }

    End { $ipArray }
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

            $subnets = Get-NetIPLocationList

            for ($i = 1; $i -le (($subnets | Measure-Object).Count); $i++) {
                $ConfigHash.netMapList.Add([PSCustomObject]@{
                        Id           = ($ConfigHash.netMapList | Measure-Object).Count + 1
                        Network      = ($subnets[$i - 1].Network)
                        ValidNetwork = $true
                        Mask         = $subnets[$i - 1].Mask
                        ValidMask    = $true
                        Location     = if ($subnets[$i - 1].Location -ne $null) { $subnets[$i - 1].Location }
                                        else { "Unknown" }
                    
                    })                                                               
            }
        }


        elseif ($SyncHash.settingNetImportClick.SelectedIndex -eq 2) {

            $scopes =  Get-DhcpServerInDC | ForEach-Object {Get-DhcpServerv4Scope -ComputerName $_.DNSName} |
                            Select-Object ScopeID, SubnetMask, Name

            foreach ($scope in $scopes) {
                $ConfigHash.netMapList.Add([PSCustomObject]@{
                        Id           = ($ConfigHash.netMapList | Measure-Object).Count + 1
                        Network      = ($scope.ScopeID).IPAddresstoString
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
    $testLog = Get-Content ((Get-ChildItem -Path (Join-Path $ConfigHash.($type + 'LogPath') -ChildPath *) -Include *.txt, *.log, *.csv |
                Sort-Object LastWriteTime -Descending |
                    Select-Object -First 1).FullName) |
                Select-Object -Last 11 |
                    Where-Object { $_.Trim() -ne '' -and $_ -notmatch ",.[(\W)]$" -and $_ -notmatch ",$" } | 
                        Select-Object -Last 1

    # If empty, all latest entries have a seperate field without a value, so we'll just grab the last non-empty line
    if (!$testLog) { 
        $testLog = Get-Content ((Get-ChildItem -Path (Join-Path $ConfigHash.($type + 'LogPath') -ChildPath *) -Include *.txt, *.log, *.csv |
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

    if (!$selectedDirectory -and ($ConfigHash.($type + 'LogPath') -and (Test-Path $ConfigHash.($type + 'LogPath')))) {$selectedDirectory = $ConfigHash.($type + 'LogPath')}

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

    ($ConfigHash.queryDefConfig | Where-Object { $_.Name -in $SyncHash.searchPropSelection.SelectedItems }).QueryDefType |
            ForEach-Object { $ConfigHash.adPropertyMap[$_] }
}


#endregion

#region Querying
function Set-ItemExpanders {
    param($SyncHash, $ConfigHash,
        [ValidateSet('Enable', 'Disable')]$IsActive,
        [ValidateSet('All', 'CompOnly','UserOnly')]$Selection,
        [switch]$ClearContent)
    
    if ($IsActive -eq 'Disable') { 
        $SyncHash.Window.Dispatcher.Invoke([Action] {                 
                $SyncHash.compExpander.IsExpanded = $false
                $SyncHash.compExpanderProgressBar.Visibility = 'Visible'                    
                $SyncHash.userExpander.IsExpanded = $false 
                $SyncHash.expanderProgressBar.Visibility = 'Visible'
                $syncHash.userToolControlPanel.Visibility = 'Collapsed'
                $syncHash.compToolControlPanel.Visibility = 'Collapsed'
                $syncHash.settingToolParent.Visibility = "Collapsed"
                   
                if ($ClearContent) {
                    $syncHash.expanderDisplay.Content = $null
                    $syncHash.expanderTypeDisplay.Content = $null
                    $syncHash.compExpanderTypeDisplay.Content = $null
                }
            })
    }
    else {
        $SyncHash.Window.Dispatcher.Invoke([Action] {   
            if ($selection -match "$null|All|CompOnly") {         
                $SyncHash.compExpander.IsExpanded = $true
                $SyncHash.compExpanderProgressBar.Visibility = 'Collapsed' 
            }   
             if ($selection -match "$null|All|UserOnly") {               
                $SyncHash.userExpander.IsExpanded = $true 
                $SyncHash.expanderProgressBar.Visibility = 'Collapsed'
            }
        })
    }
}

function Start-ObjectSearch {
    param ($SyncHash, $ConfigHash, $queryHash, $Key)  
    
    $configHash.IsSearching = $true

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
            if ($match.SamAccountName -eq $configHash.currentTabItem -or $match.Name -eq $configHash.currentTabItem) {
                $configHash.IsSearching = $false
                $rsCmd.queue.Enqueue('Queried item is current item!')
                exit
            }

            Set-ItemExpanders -SyncHash $SyncHash -ConfigHash $ConfigHash -IsActive Disable
                        
            if ($match.ObjectClass -eq 'User') {
                $match = (Get-ADUser -Identity $match.SamAccountName -Properties @($ConfigHash.UserPropList.PropName.Where( { $_ -ne 'Non-AD Property' }) | Sort-Object -Unique))                 
                #Set-QueryVarsToUpdate -ConfigHash $ConfigHash -Type User
                Write-LogMessage -Path $ConfigHash.actionlogPath -Message Query -SubjectName $match.SamAccountName -ActionName 'Query' -SubjectType 'User' -ArrayList $ConfigHash.actionLog
            }

            elseif ($match.ObjectClass -eq 'Computer') {                       
                $match = (Get-ADComputer -Identity $match.SamAccountName -Properties @($ConfigHash.CompPropList.PropName.Where( { $_ -ne 'Non-AD Property' }) | Sort-Object -Unique))
                #Set-QueryVarsToUpdate -ConfigHash $ConfigHash -Type Comp
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

        $configHash.IsSearching = $false
    }
}

function Find-ObjectLogs {
    param (
        $SyncHash, $queryHash, $ConfigHash, $match,
        [ValidateSet('User', 'Comp')]$type)

    if ($type -eq 'User') { $idProp = 'SamAccountName' }
    else { $idProp = 'Name' }

    $queryHash.($match.$idProp) = @{ }
    $queryHash.($match.$idProp).QueryID = Get-Random
    
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
            Start-Sleep -Milliseconds 500
            $startID = $queryHash.$($match.SamAccountName).QueryID
        
            if ($ConfigHash.UserLogPath -and (Test-Path (Join-Path -Path $ConfigHash.UserLogPath -ChildPath "$($match.SamAccountName)`.*"))) {
                $queryHash.$($match.SamAccountName).LoginLogPath = (((Get-ChildItem (Join-Path -Path $ConfigHash.UserLogPath -ChildPath "$($match.SamAccountName)`.*") -Include *.txt, *.csv, *.log))[0]).FullName
                $queryHash.$($match.SamAccountName).LoginLogRaw = Get-Content $queryHash.$($match.SamAccountName).LoginLogPath |
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
                                                                              
                        if ($hostConnectivity.Online) {
                              $id = "RDRS$(Get-Random)"

                                $rsArgs = @{
                                    Name            =  $id
                                    ArgumentList    = $log.ComputerName, $match.SamAccountName
                                    ModulesToImport = $configHash.modList
                                } 
                                
                                Start-RSJob @rsArgs -ScriptBlock {
                                    param ($ComputerName, $UserName) 
                                    Get-RDSession -ComputerName $ComputerName -UserName $UserName
                                } | Out-Null
    
                                Wait-RSJob -Name $id -Timeout 10 | Out-Null
                                if ((Get-RSJob -Name $id).State -ne 'Completed') { $sessionInfo = 'Failed' }
                                else { $sessionInfo = (Get-RSJob -Name $id | Receive-RSJob) }
    
                        }

                        if ($hostConnectivity.IPV4Address) { $hostLocation = Resolve-Location -computerName $log.ComputerName -IPList $ConfigHash.netMapList -ErrorAction SilentlyContinue }
                        
                        if ($log.ClientName) { $clientOnline = Test-OnlineFast -ComputerName ($log.ClientName -replace ' ') }      
                        if ($clientOnline.IPV4Address) { $clientLocation = Resolve-Location -computerName ($log.ClientName -replace ' ') -IPList $ConfigHash.netMapList -ErrorAction SilentlyContinue }
                                            
                                            
                        $queryHash.$($match.SamAccountName).LoginLog.Add(( New-Object PSCustomObject -Property @{
                          
                                    logonTime       = Get-Date($log.DateTime) -Format MM/dd/yyyy
                            
                                    HostName        = $log.ComputerName
                            
                                    LoginDC         = $log.LoginDC -replace '\\'
                            
                                    UserName        = $match.SamAccountName
                            
                                    Connectivity    = ($hostConnectivity.Online).toString()
                            
                                    IPAddress       = $hostConnectivity.IPV4Address
                            
                                    userOnline      = if ($sessionInfo -eq 'Failed') { 'Failed' }
                                                      elseif ($sessionInfo) { $true }
                                                      else { $false }
                            
                                    sessionID       = if ($sessionInfo) { $sessionInfo.sessionID }
                                    else { $null }
                            
                                    IdleTime        = if ($sessionInfo) {
                                        if ('{0:dd\:hh\:mm}' -f $($sessionInfo.IdleTime) -eq '00:00:00') { 'Active' }
                                        else { '{0:dd\:hh\:mm}' -f $($sessionInfo.IdleTime) }   
                                    }
                                    else { $null }

                                    ClientName      = ($log.ClientName -replace ' ')
                            
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
                                                    if (([scriptblock]::Create($ConfigHash.nameMapList[$r].Condition)).Invoke()) {
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
                                                    if (([scriptblock]::Create($ConfigHash.nameMapList[$r].Condition)).Invoke()) {
                                                        $ConfigHash.nameMapList[$r].Name
                                                        break
                                                    }
                                                }
                                                catch { }
                                            }
                                        }
                                    }
                                }))
                        
                        
                        if ($match.SamAccountName -notin $queryHash.Keys -or
                            (!($queryHash[$match.SamAccountName].QueryID) -or ($queryHash[$match.SamAccountName].QueryID -ne $startID))) { break }
                                                        
                        if ($ConfigHash.userLogMapping.FieldSel -contains 'Custom') {
                            foreach ($customHeader in ($ConfigHash.userLogMapping | Where-Object { $_.FieldSel -eq 'Custom' })) {
                                foreach ($item in ($queryHash.$($match.SamAccountName).LoginLog)) { $item | Add-Member -Force -MemberType NoteProperty -Name $customHeader.Header -Value $log.($customHeader.Header) }
                            }
                        }

                        if (!$refreshTimer -or ($refreshTimer.Elapsed.TotalSeconds -ge 5)) {
                            $SyncHash.Window.Dispatcher.Invoke([Action] { $queryHash.$($match.SamAccountName).LoginLogListView.Refresh() })
                            $refreshTimer =  [system.diagnostics.stopwatch]::StartNew()
                        }
                            

                        if (($SyncHash.userCompGrid.Items | Measure-Object).Count -ge 1 -and !$gridIsPopulated) {
                            $gridIsPopulated = $true
                            $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.UserCompGrid.SelectedItem = $SyncHash.UserCompGrid.Items[0] })
                            Set-ItemExpanders -SyncHash $SyncHash -ConfigHash $ConfigHash -IsActive Enable -Selection CompOnly
                            $queryHash[$match.SamAccountName].logsSearched = $true
                        }
                    } 

                    if (!($queryHash[$match.SamAccountName].QueryID) -or ($queryHash[$match.SamAccountName].QueryID -ne $startID)) {
                        exit
                    }

                    if ($match.SamAccountName -in $queryHash.Keys) {
                        $SyncHash.Window.Dispatcher.Invoke([Action] { $queryHash.$($match.SamAccountName).LoginLogListView.Refresh() })
                    }

                   
                                                                                                 
                }
                else { $queryHash[$match.SamAccountName].logsSearched = $true }
            }                                    
            else { $queryHash[$match.SamAccountName].logsSearched = $true }
        }      
    }

    else {
        Start-RSJob @rsArgs -ScriptBlock {
            param($queryHash, $ConfigHash, $match, $SyncHash) 
            Start-Sleep -Milliseconds 500                               
            $startID = $queryHash.$($match.Name).QueryID

            if ($ConfigHash.compLogPath -and (Test-Path (Join-Path -Path $ConfigHash.compLogPath -ChildPath "$($match.Name)`.*"))) {
                $queryHash.$($match.Name).LoginLogPath = (((Get-ChildItem (Join-Path -Path $ConfigHash.compLogPath -ChildPath "$($match.name)`.*")-Include *.txt, *.csv, *.log))[0]).FullName
                $queryHash.$($match.Name).LoginLogRaw = Get-Content $queryHash.$($match.Name).LoginLogPath | 
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
                                    if (([scriptblock]::Create($ConfigHash.nameMapList[$r].Condition)).Invoke()) {
                                        $ConfigHash.nameMapList[$r].Name
                                        break
                                    }
                                }
                                catch {}                                     
                            }
                        }
                    }

                    $compPing = Test-OnlineFast $match.Name
                    if ($compPing.Online) { 
                        $id = "RDRS$(Get-Random)"

                        $rsArgs = @{
                            Name            =  $id
                            ArgumentList    = $match.Name
                            ModulesToImport = $configHash.modList
                        } 
                                
                        Start-RSJob @rsArgs -ScriptBlock {
                            param ($ComputerName) 
                            Get-RDSession -ComputerName $ComputerName
                        } | Out-Null
    
                        Wait-RSJob -Name $id -Timeout 10 | Out-Null
                        if ((Get-RSJob -Name $id).State -ne 'Completed') { $sessionInfo = 'Failed' }
                        else { $sessionInfo = (Get-RSJob -Name $id | Receive-RSJob) }                  
                    }

                    if ($compPing.IPV4Address) { $hostLocation = Resolve-Location -computerName $match.Name -IPList $ConfigHash.netMapList -ErrorAction SilentlyContinue }

                    foreach ($log in ($queryHash.$($match.Name).LoginLogRaw |
                                Group-Object User | ForEach-Object {$_ |
                                     Select-Object -ExpandProperty group |
                                         Sort-Object DateTime -Descending |
                                             Select-Object -First 1} ) | Sort-Object DateTime -Descending) {
                        Remove-Variable clientLocation -ErrorAction SilentlyContinue

                        if ($log.ClientName) { $clientOnline = Test-OnlineFast -ComputerName ($log.ClientName -replace ' ') }
    
                        $userSession = $sessionInfo | Where-Object { $_.UserName -eq $log.User }                                                                                                              
        
                        if ($clientOnline.IPV4Address) { $clientLocation = Resolve-Location -computerName ($log.ClientName -replace ' ') -IPList $ConfigHash.netMapList -ErrorAction SilentlyContinue}

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
                                                        if (([scriptblock]::Create($ConfigHash.nameMapList[$r].Condition)).Invoke()) {
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
                       
                        if ($match.Name -notin $queryHash.Keys -or
                        (!($queryHash[$match.Name].QueryID) -or ($queryHash[$match.Name].QueryID -ne $startID))) { break }

                        if (!$refreshTimer -or ($refreshTimer.Elapsed.TotalSeconds -ge 5)) {
                            $SyncHash.Window.Dispatcher.Invoke([Action] { $queryHash.$($match.Name).LoginLogListView.Refresh() })
                            $refreshTimer =  [system.diagnostics.stopwatch]::StartNew()
                        }
                        
                        if (($SyncHash.compUserGrid.Items | Measure-Object).Count -ge 1 -and !$gridIsPopulated) {
                            $gridIsPopulated = $true
                            $SyncHash.Window.Dispatcher.Invoke([Action] { $SyncHash.compUserGrid.SelectedItem = $SyncHash.compUserGrid.Items[0] })
                            Set-ItemExpanders -SyncHash $SyncHash -ConfigHash $ConfigHash -IsActive Enable -Selection CompOnly
                            $queryHash[$match.Name].logsSearched = $true
                        }
                    }
                    
                    if (!($queryHash[$match.Name].QueryID) -or ($queryHash[$match.Name].QueryID -ne $startID)) {
                        exit
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

   if ($type -eq 'User') { 
        $itemName = 'userComp' 
        $butCode = 'r'
    }
    else {
        $itemName = 'compUser' 
        $butCode = 'rc'
    }

    foreach ($button in ($SyncHash.Keys | Where-Object { $_ -like "*$($butcode)butbut*" })) {
        if (($SyncHash.($itemName + 'Grid').SelectedItem.ClientType -in $ConfigHash.rtConfig.($SyncHash[$button].Tag).Types) -and
            (!($ConfigHash.rtConfig.($SyncHash[$button].Tag).RequireOnline -and $SyncHash.($itemName + 'Grid').SelectedItem.ClientOnline -eq $false)) -and 
            (!($ConfigHash.rtConfig.($SyncHash[$button].Tag).RequireUser -eq $true))) { $SyncHash[$button].IsEnabled = $true }
        else { $SyncHash[$button].IsEnabled = $false }            
    }

    foreach ($button in $SyncHash.customRT.Keys) {
        if (($SyncHash.($itemName + 'Grid').SelectedItem.ClientType -in $ConfigHash.rtConfig.$button.Types) -and
            (!($ConfigHash.rtConfig.$button.RequireOnline -and $SyncHash.($itemName + 'Grid').SelectedItem.ClientOnline -eq $false)) -and 
            ($ConfigHash.rtConfig.$button.RequireUser -ne $true)) { $SyncHash.customRT.$button.($butCode + 'but').IsEnabled = $true }
        else { $SyncHash.customRT.$button.($butCode + 'but').IsEnabled = $false }            
    }

    foreach ($button in $SyncHash.customContext.Keys) {
        if (($SyncHash.($itemName + 'Grid').SelectedItem.ClientType -in $ConfigHash.contextConfig[($button -replace 'cxt') - 1].Types) -and
            (!($ConfigHash.contextConfig[($button -replace 'cxt') - 1].RequireOnline -and $SyncHash.($itemName + 'Grid').SelectedItem.ClientOnline -eq $false)) -and 
            ($ConfigHash.contextConfig[($button -replace 'cxt') - 1].RequireUser -ne $true)) {  $SyncHash.customContext.$button.(($butCode + 'but') + 'context' + ($button -replace 'cxt')).IsEnabled = $true }
        else { $SyncHash.customContext.$button.(($butCode + 'but') + 'context' + ($button -replace 'cxt')).IsEnabled = $false }            
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

    if ($SubjectName) { $queue.Enqueue("[$($toolName)]: $Status on $($SubjectName) - action $subStatus") }
    else { $queue.Enqueue("[$($toolName)]: $Status - standalone tool $subStatus") }
}
        
function Write-CustomSnackMsg {
    param ($Message, $Queue) 

    $queue.Enqueue([string]$Message)
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
            CxtSubName  = if ($ContextSubjectName) { ($ContextSubjectName -replace '[][]').toUpper()}
                          else { "[N/A]" }
            SubjectType = if ($SubjectType) { $textInfo.ToTitleCase(($SubjectType.ToLower() -replace '[][]' -replace "Comp$",'Computer')) }
                          else { "[N/A]" }
            Date        = (Get-Date -Format d)
            Time        = (Get-Date -Format t)
            DateFull    = Get-Date
            Admin       = ($env:USERNAME).ToLower()
            Error       = if ($Error) { $Error -replace 'Exception calling "Invoke" with "0" argument\(s\):' -replace '^ "' -replace '"$' }
                          else { '[none]' }
            ogValue     = if ( $OldValue ) {  $OldValue }
                          else { '[N/A]' }
            newValue    = if ( $newValue ) { $newValue }
                          else { '[N/A]' }

        }) 

    if ($logMsg.SubjectType -eq 'Computer') { $logMsg.SubjectName = ($logMsg.SubjectName).toUpper() }

    if ($syncHashWindow) { $syncHashWindow.Dispatcher.Invoke([Action] { $ArrayList.Insert(0, $logMsg) }) }
    else { $ArrayList.Insert(0, $logMsg) }

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
                Select-Object ActionName, Message, SubjectName, SubjectType, ContextSubject, Admin, Error, OldValue, NewValue, DateFull, @{Label = 'Date'; Expression = { Get-Date(Get-Date($_.DateFull)) -Format 'MM/dd/yyyy HH:mm:ss' } } |
                    Select-Object * -ExcludeProperty DateFull | ForEach-Object { $tempArray.Add($_) }
        
        
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
                New-HTMLContent  -HeaderText $Title { 
                    New-HTMLTable -DataTable $logList -DefaultSortColumn 'Date' -DateTimeSortingFormat 'M/D/YYYY HH:mm:SS tt' -DefaultSortOrder Descending -Style display -SearchBuilder {
                        New-HTMLTableStyle -FontFamily 'Segoe UI' -FontWeight 500 
                    }
                }


                New-HTMLContent -HeaderText 'General Metrics' {
                    New-HTMLPanel {
                        New-HTMLChart -Gradient -Title 'Actions' -TitleAlignment center {
                            New-ChartToolbar -Download               
                            $logList |
                                Where-Object { $_.ActionName -ne 'query' -and $_.ActionName -notlike "Refresh*" -and $_.ActionName -notlike "Remote Tool (*" } |
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
                
                 New-HTMLContent -HeaderText 'Remote Tool (RT) Metrics' {
                    New-HTMLPanel {
                        New-HTMLChart -Gradient -Title 'RT Usage' -TitleAlignment center {
                            New-ChartToolbar -Download               
                            $logList |
                                Where-Object { $_.ActionName -like "Remote Tool*" } |
                                    Group-Object -Property ActionName |
                                        ForEach-Object { New-ChartPie -Name ($_.Name -replace "remote tool " -replace "[)(]") -Value $_.Count }
                        }
                    }

                    New-HTMLPanel {
                        New-HTMLChart -Gradient -Title 'Top RT Computer Target' -TitleAlignment center {
                            New-ChartLegend -Name 'Computer'
                            $logList |
                                Where-Object { $_.ActionName -like "Remote Tool*" -and $_.SubjectName -notlike $null } |
                                    Group-Object -Property SubjectName |
                                        Sort-Object -Property Count -Descending |
                                            Select-Object -First 10 |
                                                ForEach-Object { New-ChartBar -Name $_.Name -Value $_.Count }
                        }
                    }

                     New-HTMLPanel {
                        New-HTMLChart -Gradient -Title 'Top RT User Target' -TitleAlignment center {
                            New-ChartLegend -Name 'Computer'
                            $logList |
                                Where-Object { $_.ActionName -like "Remote Tool*" -and $_.ContextSubject -ne '[N/A]' } |
                                    Group-Object -Property ContextSubject |
                                        Sort-Object -Property Count -Descending |
                                            Select-Object -First 10 |
                                                ForEach-Object { New-ChartBar -Name $_.Name -Value $_.Count }
                        }
                    }

                             
                }          
            }
        }
    }
}


#endregion

#endwindow