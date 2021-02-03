﻿# TODO
#
#- Color switches on glyph buttons???

#$SW_HIDE, $SW_SHOW = 0, 5
#$TypeDef = '[DllImport("User32.dll")]public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'
#Add-Type -MemberDefinition $TypeDef -Namespace Win32 -Name Functions
#$hWnd = (Get-Process -Id $PID).MainWindowHandle
#$Null = [Win32.Functions]::ShowWindow($hWnd,$SW_HIDE)
#####
#

Remove-Variable -Name * -ErrorAction SilentlyContinue
Add-Type -AssemblyName 'System.Windows.Forms'
$settingInfoHash = @{
    'settingCompPropContent' = [PSCustomObject]@{
        'Body'  = @' 
This table allows you to map the Active Directory (AD) properties displayed when querying a computer. Each item can be given a friendly name in the 'FIELD NAME'  box - this will be displayed as the header. The corresponding drop-down box will determine the AD property queried. The 'TYPE' drop down will determine the type of actions that can be performed with the item. These can be defined for each field with the button in the DEF row. The types are explained below.

The 'Non-AD Property' selection is an actionable-only field that allows its content to be populated using a custom query.
'@
        'Types' = [ordered]@{
            'ReadOnly'   = 'The field only shows the content as pulled from AD.'
            'Editable'   = 'The field value can be updated or cleared and then saved to AD.'
            'Actionable' = 'The field will allow up to two definable buttons to perform custom actions.'
            'Raw'        = 'Any raw field will display the value directly as pulled from AD. Otherwise, the presentation of the content can be defined.'
        }
    } 
    'settingUserPropContent' = [PSCustomObject]@{
        'Body'  = @' 
This table allows you to map the Active Directory (AD) properties displayed when querying a user. Each item can be given a friendly name in the 'FIELD NAME'  box - this will be displayed as the header. The corresponding drop-down box will determine the AD property queried. The 'TYPE' drop down will determine the type of actions that can be performed with the item. These can be defined for each field with the button in the DEF row. The types are explained below.

The 'Non-AD Property' selection is an actionable-only field that allows its content to be populated using a custom query.
'@
        'Types' = [ordered]@{
            'ReadOnly'   = 'The field only shows the content as pulled from AD.'
            'Editable'   = 'The field value can be updated or cleared and then saved to AD.'
            'Actionable' = 'The field will allow up to two definable buttons to perform custom actions.'
            'Raw'        = 'Any raw field will display the value directly as pulled from AD. Otherwise, the presentation of the content can be defined.'
        }
    } 
    'settingPropUserDefine'  = [PSCustomObject]@{
        'Body' = @'
These fields define how the selected property will function in regards to the button type selected in the previous table. 

The 'Result Presentation' scriptblock is present in any non-raw type. This will allow the property returned to be presented as an alternative value (e.g. a TRUE or FALSE value can be passed through an if statement and alternate text can be returned and displayed.

The 'Attached Actions I-II' are buttons that will attach to the returned value when queried. Their respective scriptblock will run when the button is processed. Along each attached action, an icon can be selected for use with the button. The 'Refresh Prop' option will requery Active Directory after the action completes and update the value in the display. The 'New Thread' option will run the action in a new thread, though this may not neccarrily be faster overall for quicker actions.

All script blocks must be validated using the 'Execute' button. This will only execute fully for the results scriptblock, but will altert for syntaxical errors for each button's respective scriptblock.

The variables below can be referenced or manipulated within the scriptblocks.
'@
        'Vars' = [ordered]@{          
            '$result [Result Presentation]'      = 'The actual value of the returned Active Directory property.'
            '$resultColor [Result Presentation]' = 'Setting the resultColor will determine the color of the returned value. This uses all valid .NET brush names and HEX values.'
            '$user [Actionable Items]'           = 'The value of the current, queried user. Only populated on action buttons for user properties.'
            '$comp [Actionable Items]'           = 'The value of the current, queried computer. Only populated on action buttons for computer properties.'
            '$actionObject [Actionable Items] '  = 'General variable containing hte name of queried object (i.e. the user or computer)'
            '$propName [Actionable Items]'       = 'The name of the property attached to the field. Not applicable to non-AD queries.'
            '$prop [Actionable Items]'           = 'The value of the queried property attached to the field.'
        }
    }
    'settingPropCompDefine'  = [PSCustomObject]@{
        'Body' = 'TestEdit'
        'Vars' = @{
            'One' = 'Test'
        }
    }
    'settingItemToolsContent' =  [PSCustomObject]@{
        'Body'  = @' 
To be completed.. settingItemToolsContent
'@
        'Types' = @{}
}
    'settingContextPropContent' =  [PSCustomObject]@{
        'Body'  = @' 
Contextual tools are actions accessible through the historical view. It uses the currently queried object and the selection made within the historical view to perform tasks defined within this configuration section. 
'@
      
}
    'settingVarContent' =  [PSCustomObject]@{
        'Body'  = @'  
Resources and variables are items updated as defined intervals that allow custom data to be accessed through out scriptblocks found throughout the tool.

Functions and modules are PowerShell items added to allow access to scriptblocks executing in new threads. If not added, they can not be accessed in the thread.
'@
        'Types' = @{}
}
    'settingModDataGrid' =  [PSCustomObject]@{
        'Body'  = @' 
To be completed.. settingModDataGrid
'@
        'Types' = @{}
}
    'settingOUDataGrid' =  [PSCustomObject]@{
        'Body'  = @' 
Add Active Directory OUs and containers - these will be the locations searched within during queries for users and computers.

Select the '+' button to add an additional OU definition. Select the button it its respective row to choose the OU. Select the search scope from the drop down - the search scope types are listed below.
'@
        'Types' = [ordered]@{
            'Subtree' = 'The default. This scope will search an OU and all OUs within it recursively.'
            'OneLevel' ='This will search an OU for objects only one level deep. Nested OUs are ignored.'}
}
    'settingQueryDefDataGrid' =  [PSCustomObject]@{
        'Body'  = @'
When querying for users or computers, each of the items added as query definitions will act as the properties searched to match against the provided search term. The NAME field is the 'friendly name' that will appear in the search settings (which can be toggled prior to querying). The AD Property field is the actual AD property queried.
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
To be completed.. settingNetContent
'@
        'Types' = @{}
}
    'settingNameContent' =  [PSCustomObject]@{
        'Body'  = @' 
To be completed.. settingNameContent
'@
        'Types' = @{}
}
    'settingGeneralContent' =  [PSCustomObject]@{
        'Body'  = @' 
These are general options related to querying and logging.
'@
}
   'settingMiscGrid' =  [PSCustomObject]@{
        'Body'  = @' 
Miscellaneous settings.

The logging path defines where this tool's actions are logged. Ideally, this should be the same network location for all administrators using this tool. 

Login log view depth refers to how far back login logs will be searched, analyzed, and displayed on the historical view. Larger depth will result in longer overall querying time, but this number may need to be adjusted to best fit your enviornment's usage.

Active directory mappings are an index of the entire list of the AD object properties and their related data. If the AD schema is updated, these should be refreshed using the button below.
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
Select directories that store the login logs for both clients and users. These should be .CSV, or data files otherwise delimited by commas. They should at least contain the username of the user, the date, and the computer name. After choosing the respective directory, the structure of the logs can be defined to map each attribute.

After selecting the directory with the gear button, the edit button will enable - from here, you can map each value to its respective type. The given values are pulled from the newest log entry in the previously defined logging path. 

The IGNORE selection will skip the property, while the CUSTOM selection will allow you to assign a custom friendly name to display this value in the historical view when this item type is queried.
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
                     
        }

        'Tips' = [ordered]@{
            'Confirmation Window' = @'
Adding the following line at any point within the scriptblock will launch a confirmation window before continuing the block:
    New-DialogWait -ConfirmWindow $confirmWindow -Window 
        $window -configHash $configHash 
'@         
        }
}
'settingObjectToolDefFlyout' =  [PSCustomObject]@{
        'Body'  = @' 
To be completed.. settingObjectToolDefFlyout
'@
        'Types' = @{}
}
'settingVarDataGrid' =  [PSCustomObject]@{
        'Body'  = @' 
Variables defined below will populate according to the set frequency and allow access from within scriptblocks defined elsewhere.

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
To be completed.. settingAdminContent
'@
}
}




$modList = @('C:\TempData\func\func.psm1', 'C:\TempData\internal\internal.psm1')
Import-Module $modList -Force -DisableNameChecking


$ConfirmPreference = 'None'
#Copy-Item -Path  C:\TempData\v3\MainWindow.xaml -Destination C:\TempData\MainWindow.xaml -Force
Copy-Item -Path '\\labtop\TempData\v3\v3\MainWindow.xaml'  -Destination C:\TempData\MainWindow.xaml -Force
$xamlPath = 'C:\TempData\MainWindow.xaml'
$helpXAMLPath = 'C:\TempData\HelpContent.xaml'
$glyphList = 'C:\TempData\segoeGlyphs.txt'

######

$savedConfig = 'C:\TempData\config.json'


# generated hash tables used throughout tool
New-HashTables

# Import from JSON and add to hash table
Set-Config -ConfigPath $savedConfig -Type Import -ConfigHash $configHash

# process loaded data or creates initial item templates for various config datagrids
@('userPropList', 'compPropList', 'contextConfig', 'objectToolConfig', 'nameMapList', 'netMapList', 'varListConfig', 'searchBaseConfig', 'queryDefConfig', 'modConfig') | Set-InitialValues -ConfigHash $configHash -PullDefaults
@('userLogMapping', 'compLogMapping') | Set-InitialValues -ConfigHash $configHash

# matches config'd user/comp logins with default headers, creates new headers from custom values
$defaultList = @('User', 'DateRaw', 'LoginDc', 'ClientName', 'ComputerName', 'Ignore', 'Custom')
@('userLogMapping', 'compLogMapping') | Set-LoggingStructure -DefaultList $defaultList -ConfigHash $configHash
Set-ActionLog -ConfigHash $configHash

$configHash.modConfig.modPath | ForEach-Object -Process { $modList += $_ }
$configHash.modList = $modList | Where-Object -FilterScript { [string]::IsNullOrWhiteSpace($_) -eq $false }

# Add default values if they are missing
@('MSRA', 'MSTSC') | Set-RTDefaults -ConfigHash $configHash

# loaded required DLLs
foreach ($dll in ((Get-ChildItem -Path C:\TempData\asm\ -Filter *.dll).FullName)) { $null = [System.Reflection.Assembly]::LoadFrom($dll) }



# read xaml and load wpf controls into synchash (named synchash)
Set-WPFControls -TargetHash $syncHash -XAMLPath $xamlPath

Get-Glyphs -ConfigHash $configHash -GlyphList $glyphList
$syncHash.settingLogo.Source = 'C:\tempdata\trident.png'
$syncHash.logoBackdrop.Source = 'C:\tempdata\trident.png'
#Set-WPFControls -TargetHash $helpHash -XAMLPath $helpXAMLPath

# builds custom WPF controls from whatever was defined and saved in ConfigHash
Add-CustomItemBoxControls -SyncHash $syncHash -ConfigHash $configHash
Add-CustomToolControls -SyncHash $syncHash -ConfigHash $configHash
Add-CustomRTControls -SyncHash $syncHash -ConfigHash $configHash
$sysCheckHash.missingCheckFail = $false


#region Item Tool Events
#region Item Tools - Grid
$syncHash.ItemToolGridADSelectionButton.Add_Click( { 
    switch ($syncHash.ItemToolGridADSelectionButton.Tag) {
        'AD' { Set-ADItemBox -ConfigHash $configHash -SyncHash $syncHash -Control Grid }
        'OU' { Set-OUItemBox -ConfigHash $configHash -SyncHash $syncHash -Control Grid }
        'Custom' { Set-CustomItemBox -ConfigHash $configHash -SyncHash $syncHash -Control Grid }
    }

})

$syncHash.itemToolGridItemsGrid.Add_SelectionChanged( {
        if ($syncHash.itemToolGridItemsGrid.SelectedItem.Image) { $syncHash.itemToolImageSource.Source = [byte[]]($syncHash.itemToolGridItemsGrid.SelectedItem.Image) }
    })

$syncHash.itemToolCustomConfirm.Add_Click({
        $configHash.customDialogClosed = $true
        $configHash.customInput = $syncHash.itemToolCustomContent.Text

})

$syncHash.itemToolCustomContent.add_KeyDown({ 
    if ($_.Key -eq 'Enter') {
        $configHash.customDialogClosed = $true
        $configHash.customInput = $syncHash.itemToolCustomContent.Text
    } 
})

$syncHash.itemToolGridItemsGrid.Add_AutoGeneratingColumn( {
        if ($_.Column.Header -eq 'Image') {
            $_.Cancel = $true 
            $syncHash.itemToolImageBorder.Visibility = 'Visible'
        }
    })

$syncHash.itemToolGridSearchBox.Add_TextChanged( {
        $syncHash.itemToolGridItemsGrid.ItemsSource.Filter = $null
        $syncHash.itemToolGridItemsGrid.ItemsSource.Filter = { param ($item) $item -match $syncHash.itemToolGridSearchBox.Text }
    })

$syncHash.itemToolGridSelectConfirmCancel.Add_Click( { $syncHash.itemToolDialog.IsOpen = $false })

$syncHash.itemToolGridSelectAllButton.Add_Click( {    
        if ($syncHash.itemToolGridItemsGrid.Items.Count -eq $syncHash.itemToolGridItemsGrid.SelectedItems.Count) { $syncHash.itemToolGridItemsGrid.UnselectAll() } 
        else { $syncHash.itemToolGridItemsGrid.SelectAll() }      
    })

#endregion

#region Item Tools - List
$syncHash.ItemToolADSelectionButton.Add_Click( {
    switch ($syncHash.ItemToolADSelectionButton.Tag) {
        'AD' { Set-ADItemBox -ConfigHash $configHash -SyncHash $syncHash -Control ListBox }
        'OU' { Set-OUItemBox -ConfigHash $configHash -SyncHash $syncHash -Control ListBox }
        'Custom' { Set-CustomItemBox -ConfigHash $configHash -SyncHash $syncHash -Control ListBox }
    }
})



$syncHash.itemToolListSearchBox.Add_TextChanged( {
        $syncHash.itemToolListSelectListBox.ItemsSource.Filter = $null
        $syncHash.itemToolListSelectListBox.ItemsSource.Filter = { param ($item) $item -match $syncHash.itemToolListSearchBox.Text }
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

$syncHash.settingNamingClick.add_Click( { Set-ChildWindow -Panel settingNameContent -Title 'Device Categorization' -SyncHash $syncHash -Height 275 })

$syncHash.settingVarClick.add_Click( { Set-ChildWindow -Panel settingVarContent -Title 'Resources and Variables' -SyncHash $syncHash -Height 275 -Width 530 })

$syncHash.settingGeneralClick.add_Click( { Set-ChildWindow -Panel settingGeneralContent -Title 'General Settings' -SyncHash $syncHash -Height 275 -Width 530 })

$syncHash.settingUserPropClick.add_Click( { Set-ChildWindow -Panel settingUserPropContent -Title 'User Property Mappings' -SyncHash $syncHash -Height 365 -Width 600 })

$syncHash.settingCompPropClick.add_Click( { Set-ChildWindow -Panel settingCompPropContent -Title 'Computer Property Mappings' -SyncHash $syncHash -Height 365 -Width 600 })

$syncHash.settingObjectToolsClick.add_Click( { Set-ChildWindow -Panel settingItemToolsContent -Title 'Object Tools Mappings' -SyncHash $syncHash -Height 365 -Width 600 })

$syncHash.settingContextClick.add_Click( { Set-ChildWindow -Panel settingContextPropContent -Title 'Contextual Actions Mappings'-SyncHash $syncHash -Height 365 -Width 390 })

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
                'Admin'                      = 'False'
                'Modules'                    = 'False'
                'ADDS'                       = 'False'
                'DelegatedGroupName'         =  if (Test-Path (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-delegatedGroup.json")) { (Get-Content (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-delegatedGroup.json") | ConvertFrom-Json).Name }
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
                
                Start-AdminCheck -SysCheckHash $sysCheckHash
                   
                # Check individual checks; mark parent categories as true is children are true       
                switch ($sysCheckHash.sysChecks) {
                    { $_.ADModule -eq $true -and $_.RSModule -eq $true } { $sysCheckHash.sysChecks[0].Modules = 'True' }
                    { $_.ADMember -eq $true -and $_.ADDCConnectivity -eq $true } { $sysCheckHash.sysChecks[0].ADDS = 'True' }
                    { $_.IsInAdmin -eq $true -and $_.IsDomainAdminOrDelegated -eq $true } { $sysCheckHash.sysChecks[0].Admin = 'True' }
                }

                @('settingADMemberLabel', 'settingADDCLabel', 'settingModADLabel', 'settingModRSLabel', 'settingDomainAdminLabel', 
                    'settingLocalAdminLabel', 'settingPermLabel', 'settingADLabel', 'settingModLabel', 'settingDelegatedPanel', 'settingDelegatedGroupSelection') | 
                    Set-RSDataContext -SyncHash $syncHash -DataContext $sysCheckHash.sysChecks
               
                $sysCheckHash.checkComplete = $true

                Start-Sleep -Seconds 2

                if ($sysCheckHash.sysChecks[0].ADDS -eq $false -or 
                    $sysCheckHash.sysChecks[0].Modules -eq $false -or 
                    $sysCheckHash.sysChecks[0].Admin -eq $false) { Suspend-FailedItems -SyncHash $syncHash -CheckedItems SysCheck }
                
                else { 
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
                
                Show-WPFWindow -SyncHash $syncHash     
            }
        }
        else {                                   
            Suspend-FailedItems -SyncHash $syncHash -CheckedItems RSCheck
            Show-WPFWindow -SyncHash $syncHash
        }
    })

$syncHash.settingDelegatedGroupPick.Add_Click({
    $configRoot = Split-Path $savedConfig 
    $sysCheckHash.sysChecks[0].DelegatedGroupName = (Select-ADObject -Type Groups).Name
    $syncHash.settingDelegatedGroupSelection.Content = $sysCheckHash.sysChecks[0].DelegatedGroupName
    $sysCheckHash.sysChecks[0].DelegatedGroupName | Select-Object @{Label = "Name"; Expression = {$_}} | 
        ConvertTo-Json | Out-File (Join-Path -Path $configRoot -ChildPath "$($env:USERDOMAIN)-delegatedGroup.json")
})
#endregion 



#region RemoteTools events

$syncHash.settingRtRDPClick.Add_Click( { Set-StaticRTContent -SyncHash $syncHash -ConfigHash $configHash -Tool MSTSC })

$syncHash.settingRtMSRAClick.Add_Click( { Set-StaticRTContent -SyncHash $syncHash -ConfigHash $configHash -Tool MSRA })

$syncHash.settingRemoteFlyout.Add_OpeningFinished( { 
$syncHash.settingChildWindow.ShowCloseButton = $false
Get-RTFlyoutContent -ConfigHash $configHash -SyncHash $syncHash })

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
                    QueryDefTypeList = $configHash.QueryADValues.Key
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

$syncHash.settingPermClick.add_Click( { Set-ChildWindow -SyncHash $syncHash -Panel settingAdminContent -Title 'Admin Permissions' }) 

#endregion Category button events
#region Action pane exit events

$syncHash.settingCloseClick.add_Click( { $syncHash.Window.Close() })
    
$syncHash.settingConfigCancelClick.add_Click( { $syncHash.Window.Close() })

$syncHash.settingImportClick.add_Click( { Import-Config -SyncHash $syncHash })

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

$syncHash.searchBoxHelp.add_MouseLeftButtonDown( { $syncHash.searchBoxHelp.Background = '#576573' })

$syncHash.searchBoxHelp.add_MouseLeftButtonUp( {
        $syncHash.childHelp.isOpen = $true  
        $syncHash.searchBoxHelp.Background = 'Transparent'
    })

$syncHash.HistoryToggle.Add_MouseLeftButtonUp( {
        if ($syncHash.historySidePane.IsOpen) { $syncHash.historySidePane.IsOpen = $false }
        else { $syncHash.historySidePane.IsOpen = $true }
        $syncHash.historySideDataGrid.Items.Refresh()
    })


$syncHash.tabControl.ItemsSource = New-Object -TypeName System.Collections.ObjectModel.ObservableCollection[Object]

$syncHash.tabMenu.add_Loaded( {
        ###SEARCHHERE#






        if ($sysCheckHash.sysChecks.RSModule -eq $true) {
            
            $rsArgs = @{
                Name            = 'menuLoad'
                ArgumentList    = @($syncHash, $savedConfig, $sysCheckHash)
                ModulesToImport = $configHash.modList
            }

            Start-RSJob @rsArgs -ScriptBlock {        
                Param($syncHash, $savedConfig, $sysCheckHash)

                do {} until ($sysCheckHash.checkComplete)
           
                if ($sysCheckHash.sysChecks.Admin -eq $false -or $sysCheckHash.sysChecks.ADDS -eq $false -or $sysCheckHash.sysChecks.Modules -eq $false ) { Suspend-FailedItems -SyncHash $syncHash -CheckedItems SysCheck }

                elseif (!(Test-Path $savedConfig)) { Suspend-FailedItems -SyncHash $syncHash -CheckedItems Config }
        
                else { $syncHash.Window.Dispatcher.invoke([action] { $syncHash.tabMenu.SelectedIndex = 0 }) }
            }
       
            New-VarUpdater -configHash $configHash
            Start-VarUpdater -configHash $configHash -varHash $varHash
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
                                                        Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName Edit -SubjectName $editObject -SubjectType $type -OldValue $oldValue -NewValue $changedValue -ArrayList  $configHash.actionLog -Error $_
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
                                                            Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName Clear -SubjectName $editObject -SubjectType $type -OldValue $oldValue -ArrayList $configHash.actionLog -Error $_
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
                                                            ArgumentList    = @($rsCmd, $queryHash, $b, $configHash, $syncHash.adHocConfirmWindow, $syncHash.Window, $varHash)
                                                            ModulesToImport = $configHash.modList
                                                        }

                                                        Start-RSJob @rsArgs -ScriptBlock {
                                                            Param($rsCmd, $queryHash, $b, $configHash, $confirmWindow, $window)
                                                        
                                                            Start-Sleep -Milliseconds 500
                                                           
                                                            Set-CustomVariables -VarHash $varHash

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

                                                                        $rsCmd.Window.Dispatcher.Invoke([Action] { $rsCmd.boxResources.($type[0] + 'box' + $rsCmd.id + 'TextBox').Text = $updatedValue })                                        
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

                                                        Set-CustomVariables -VarHash $varHash

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
        
                if ($configHash.($type + 'boxCount') -le 7) { $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.($type + 'detailGrid').Columns = 1 }) }

                elseif ($configHash.($type + 'boxCount') -gt 7 -and $configHash.($type + 'boxCount') -le 14) { $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.($type + 'detailGrid').Columns = 2 }) }
            
                else { $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.($type + 'detailGrid').Columns = 3 }) }
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

                            $user = $syncHash.UserCompGrid.SelectedItem.UserName
                        }

                        else {
                            if ($syncHash.compUserFocusUserToggle.IsChecked) { $comp = $syncHash.tabControl.SelectedItem.Name }
                
                            else { $comp = $syncHash.CompUserGrid.SelectedItem.ClientName }

                            $user = $syncHash.CompUserGrid.SelectedItem.UserName
                        }
            
                        mstsc.exe /v $comp /admin
                    })

                $syncHash.($buttonType + 2).Source = ([Convert]::FromBase64String($configHash.rtConfig.MSRA.Icon))
                $syncHash.($buttonType + 2).Parent.Tag = 'MSRA'
                $syncHash.($buttonType + 2).Parent.ToolTip = $configHash.rtConfig.MSRA.DisplayName
                $syncHash.($buttonType + 2).Parent.Add_Click( { 
                        param([Parameter(Mandatory)][Object]$sender)
                        if ($sender.Name -match 'rbut') {
                            if ($syncHash.userCompFocusHostToggle.IsChecked) { $comp = $syncHash.UserCompGrid.SelectedItem.HostName }
                            else { $comp = $syncHash.UserCompGrid.SelectedItem.ClientName }

                            $user = $syncHash.UserCompGrid.SelectedItem.UserName
                        }

                        else {
                            if ($syncHash.compUserFocusUserToggle.IsChecked) { $comp = $syncHash.tabControl.SelectedItem.Name }
                
                            else { $comp = $syncHash.CompUserGrid.SelectedItem.ClientName }

                            $user = $syncHash.CompUserGrid.SelectedItem.UserName
                        }
                        msra.exe /offerra $comp 
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
                            ([scriptblock]::Create($configHash.rtConfig.$id.cmd)).Invoke()
                   
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
                                    ArgumentList    = @($rsCmd, $syncHash, $configHash, $syncHash.adHocConfirmWindow, $syncHash.Window, $varHash)
                                    ModulesToImport = $configHash.modList
                                }

                                Start-RSJob @rsArgs -ScriptBlock {
                                    Param($rsCmd, $syncHash, $configHash, $confirmWindow, $window, $varHash)
                            
                                    $user = $rsCmd.user
                                    $comp = $rsCmd.comp
                                    $sessionID = $rsCmd.SessionID

                                    $combinedString = "$($user.ToLower()) on $($comp.ToLower())"
                                    $toolName = $rsCmd.buttonSettings.ActionName
                                    Set-CustomVariables -VarHash $varHash

                                    try {
                                        ([scriptblock]::Create($rsCmd.buttonSettings.actionCmd)).Invoke()
                                        Write-SnackMsg -Queue $rsCmd.Snackbar -ToolName $toolName -Status Success -SubjectName $combinedString                            
                                        Write-LogMessage -syncHashWindow $syncHash.Window -Path $configHash.actionlogPath -Message Succeed -ActionName $toolname -SubjectName $user -ContextSubjectName $comp -SubjectType 'Context' -ArrayList $configHash.actionLog
                                    }
                            
                                    catch {
                                        Write-SnackMsg -Queue $rsCmd.Snackbar -ToolName $toolName -Status Success -SubjectName $combinedString
                                        Write-LogMessage -syncHashWindow $syncHash.Window -Path $configHash.actionlogPath -Message Fail -ActionName $toolname -SubjectName $user -ContextSubjectName $comp -SubjectType 'Context' -ArrayList $configHash.actionLog -Error $_ 
                                    }
                                }
                            }

                            else {
                                $combinedString = "$($user.ToLower()) on $($comp.ToLower())"
                                $toolName = $configHash.contextConfig[$id - 1].ActionName
                                Set-CustomVariables -VarHash $varHash

                                try {
                                    ([scriptblock]::Create($configHash.contextConfig[$id - 1].actionCmd)).Invoke()
                                    Write-SnackMsg -Queue $syncHash.SnackMsg.MessageQueue -ToolName $toolName -Status Success -SubjectName $combinedString
                                    Write-LogMessage -Path $configHash.actionlogPath -Message Succeed -ActionName $toolName -SubjectName $user -ContextSubjectName $comp -SubjectType 'Context' -ArrayList $configHash.actionLog 
                                }
                            
                                catch {
                                    Write-SnackMsg -Queue $syncHash.SnackMsg.MessageQueue -ToolName $toolName -Status Success -SubjectName $combinedString
                                    Write-LogMessage -Path $configHash.actionlogPath -Message Fail -ActionName $toolName -SubjectName $user -ContextSubjectName $comp -SubjectType 'Context' -ArrayList $configHash.actionLog -Error $_
                                }
                            }
                        })
                
                    $syncHash.($buttonType + 'ContextGrid').AddChild($syncHash.customContext.('cxt' + $contextBut.IDNum).($buttonType + 'context' + $contextBut.IDNum))
                }
            }

            $customField = 0
            $sizeToExpand = ([math]::Ceiling(($configHash.userLogMapping | Where-Object -FilterScript { $_.FieldSel -ne 'Ignore' }).Count / 5) - 1) * 55 + $syncHash.userCompControlRow.Height.Value
            $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.userCompControlRow.Height = $sizeToExpand })
            $configHash.userLogMapping |
                Where-Object -FilterScript { $_.FieldSel -eq 'Custom' -and $_.Ignore -eq $false } |
                    ForEach-Object -Process {
                        # populate custom dock items
                        $customField++
                        $syncHash.Window.Dispatcher.Invoke([Action] {
                                $syncHash.('customPropDock' + $customField) = New-Object System.Windows.Controls.StackPanel
                                $syncHash.('customPropLabel' + $customField) = New-Object System.Windows.Controls.Label
                                $syncHash.('customPropText' + $customField) = New-Object System.Windows.Controls.Textbox
                                $syncHash.('customPropLabel' + $customField).Content = $_.Header
                                $syncHash.('customPropDock' + $customField).VerticalAlignment = 'Top'
                                $syncHash.('customPropDock' + $customField).Margin = '0,-10,0,0'
                

                                $syncHash.('customPropDock' + $customField).AddChild(($syncHash.('customPropLabel' + $customField)))
                                $syncHash.('customPropDock' + $customField).AddChild(($syncHash.('customPropText' + $customField)))
     

                                $syncHash.('customPropLabel' + $customField).FontSize = '10'
                                $syncHash.('customPropLabel' + $customField).Foreground = $syncHash.Window.FindResource('MahApps.Brushes.SystemControlBackgroundBaseMediumLow')
                                $syncHash.('customPropText' + $customField).Style = $syncHash.Window.FindResource('compItemBox') 
                                $syncHash.userLogExtraPropGrid.AddChild(($syncHash.('customPropDock' + $customField)))

                                # Create and set a binding on the textbox object
                                $Binding = New-Object System.Windows.Data.Binding
                                $Binding.UpdateSourceTrigger = 'PropertyChanged'
                                $Binding.Source = $syncHash.userCompGrid
                                $Binding.Path = "SelectedItem.$($_.Header)"
                                $Binding.Mode = [System.Windows.Data.BindingMode]::TwoWay
     

                                [void][System.Windows.Data.BindingOperations]::SetBinding(($syncHash.('customPropText' + $customField)), [System.Windows.Controls.TextBox]::TextProperty, $Binding)
                            })
                    }

            $customField = 0
            $sizeToExpand = ([math]::Ceiling(($configHash.compLogMapping | Where-Object -FilterScript { $_.FieldSel -ne 'Ignore' }).Count / 5) - 1) * 55 + $syncHash.compUserControlRow.Height.Value
            $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.compUserControlRow.Height = $sizeToExpand }) 
            $configHash.compLogMapping |
                Where-Object -FilterScript { $_.FieldSel -eq 'Custom' -and $_.Ignore -eq $false } |
                    ForEach-Object -Process {
                        # populate custom dock items
                        $customField++
                        $syncHash.Window.Dispatcher.Invoke([Action] {
                                $syncHash.('customCompPropDock' + $customField) = New-Object System.Windows.Controls.StackPanel
                                $syncHash.('customCompPropLabel' + $customField) = New-Object System.Windows.Controls.Label
                                $syncHash.('customCompPropText' + $customField) = New-Object System.Windows.Controls.Textbox
                                $syncHash.('customCompPropLabel' + $customField).Content = $_.Header
                                $syncHash.('customCompPropDock' + $customField).VerticalAlignment = 'Top'
                                $syncHash.('customCompPropDock' + $customField).Margin = '0,-10,0,0'

                                $syncHash.('customCompPropDock' + $customField).AddChild(($syncHash.('customCompPropLabel' + $customField)))
                                $syncHash.('customCompPropDock' + $customField).AddChild(($syncHash.('customCompPropText' + $customField)))
     

                                $syncHash.('customCompPropLabel' + $customField).FontSize = '10'
                                $syncHash.('customCompPropLabel' + $customField).Foreground = $syncHash.Window.FindResource('MahApps.Brushes.SystemControlBackgroundBaseMediumLow')
                                $syncHash.('customCompPropText' + $customField).Style = $syncHash.Window.FindResource('compItemBox') 
                                $syncHash.compLogExtraPropGrid.AddChild(($syncHash.('customCompPropDock' + $customField)))

                                # Create and set a binding on the textbox object
                                $Binding = New-Object System.Windows.Data.Binding
                                $Binding.UpdateSourceTrigger = 'PropertyChanged'
                                $Binding.Source = $syncHash.compUserGrid
                                $Binding.Path = "SelectedItem.$($_.Header)"
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
            $syncHash.settingToolParent.Visibility = 'Collapsed'
            $syncHash.userToolControlPanel.Visibility = 'Collapsed'
            $syncHash.compToolControlPanel.Visibility = 'Collapsed'
            $syncHash.tabControl.IsEnabled = $false
        }

        else {
            Get-RSJob -State Completed | Remove-RSJob
            $syncHash.tabControl.IsEnabled = $true
    
            $configHash.currentTabItem = $syncHash.tabControl.SelectedItem.Name
            $queryHash.Keys | ForEach-Object -Process { $queryHash.$_.ActiveItem = $false }
    
    
    
            $currentTabItem = $syncHash.tabControl.SelectedItem.Name

            $rsArgs = @{
                Name            = 'VisualChange'
                ArgumentList    = @($syncHash, $queryHash, $configHash, $currentTabItem)
                ModulesToImport = $configHash.modList
            }

            Start-RSJob @rsArgs -ScriptBlock {
                Param($syncHash, $queryHash, $configHash, $currentTabItem)
            
                $queryHash.$currentTabItem.ActiveItem = $true

                $syncHash.Window.Dispatcher.Invoke([Action] {
                        $syncHash.expanderDisplay.Content = $null
                        $syncHash.expanderTypeDisplay.Content = $null
                        $syncHash.compExpanderTypeDisplay.Content = $null
                        $syncHash.userExpander.IsExpanded = $false
                        $syncHash.compExpander.IsExpanded = $false
                        $syncHash.settingToolParent.Visibility = 'Collapsed'
                        $syncHash.userToolControlPanel.Visibility = 'Collapsed'
                        $syncHash.compToolControlPanel.Visibility = 'Collapsed'
                        $syncHash.compExpanderProgressBar.Visibility = 'Visible'
                        $syncHash.expanderProgressBar.Visibility = 'Visible'

                        if ($queryHash.($syncHash.tabControl.SelectedItem.Name).ObjectClass -eq 'Computer') {
                            if (($syncHash.compToolControlPanel.Children | Measure-Object).Count -eq 0) { $syncHash.settingTools.Visibility = 'Collapsed' }
                        }
                        else {
                            if (($syncHash.userToolControlPanel.Children | Measure-Object).Count -eq 0) { $syncHash.settingTools.Visibility = 'Collapsed' }
                        }
                    })
           
                if ($null -ne $currentTabItem) {
                
                    $rsArgs = @{
                        Name            = 'displayUpdate'
                        ArgumentList    = @($syncHash, $queryHash, $configHash, $currentTabItem)
                        ModulesToImport = $configHash.modList
                    }

                    Start-RSJob @rsArgs -ScriptBlock {        
                        Param($syncHash, $queryHash, $configHash, $currentTabItem)                     
                        #Start-Sleep -Milliseconds 500
               
                        $type = $queryHash[$currentTabItem].ObjectClass -replace 'Computer', 'Comp'
                        for ($i = 1; $i -le $configHash.boxMax; $i++) {
                            if ($i -le $configHash.($type + 'boxCount')) {
                            
                                $rsArgs = @{
                                    Name            = 'displayUpdateSub'
                                    Batch           = 'displayUpdateSub'
                                    Throttle        = 24
                                    ArgumentList    = @($syncHash, $queryHash, $configHash, $currentTabItem, $i)
                                    ModulesToImport = $configHash.modList
                                }

                                Start-RSJob @rsArgs -ScriptBlock {        
                                    Param($syncHash, $queryHash, $configHash, $currentTabItem, $i)   
                                    $type = $queryHash[$currentTabItem].ObjectClass -replace 'Computer', 'Comp'

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
   
                                    $syncHash.Window.Dispatcher.invoke([action] {

                                            $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Textbox').Text = $updateHash.Text
                                            if ($updateHash.Foreground) { $syncHash.($type[0] + 'box' + $i + 'resources').($type[0] + 'box' + $i + 'Textbox').Foreground = $updateHash.Foreground }        
                              
                                        })
                                }
                            }
                        }

                     
                        Wait-RSJob -Batch displayUpdateSub
                        if ($queryHash[$currentTabItem].ObjectClass -eq 'User') {  
                            $syncHash.Window.Dispatcher.invoke([action] {            
                                    $syncHash.expanderTypeDisplay.Content = 'USER   '
                                    $syncHash.compExpanderTypeDisplay.Content = 'COMPUTERS   '
                                    $syncHash.compDetailMainPanel.Visibility = 'Collapsed'
                                    $syncHash.userDetailMainPanel.Visibility = 'Visible'
                                    $syncHash.expanderDisplay.Content = "$($queryHash[$currentTabItem].Name)"
                                    $syncHash.userExpander.IsExpanded = $true
                                    $syncHash.userCompGrid.Visibility = 'Visible'
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
                                    $syncHash.expanderTypeDisplay.Content = 'COMPUTER   '
                                    $syncHash.userDetailMainPanel.Visibility = 'Collapsed'
                                    $syncHash.compDetailMainPanel.Visibility = 'Visible'
                                    $syncHash.compExpanderTypeDisplay.Content = 'USERS   '
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
                        ArgumentList    = @($syncHash, $queryHash, $configHash, $currentTabItem)
                        ModulesToImport = $configHash.modList
                    }

                    Start-RSJob @rsArgs  -ScriptBlock {        
                        Param($syncHash, $queryHash, $configHash, $currentTabItem)        

                        $type = $queryHash[$currentTabItem].ObjectClass -replace 'Computer', 'Comp'
           
                        do { } until ($queryHash[$currentTabItem].logsSearched -eq $true)

                        if ($null -ne ($queryHash[$currentTabItem]).LoginLogListView) {
                            if ($type -eq 'User') {
                                $syncHash.userCompGrid.Dispatcher.Invoke([Action] { $syncHash.userCompGrid.ItemsSource = ($queryHash[$currentTabItem]).LoginLogListView })
                                $syncHash.Window.Dispatcher.invoke([action] {
                                        $syncHash.userCompNotFound.Visibility = 'Hidden'
                                        $syncHash.userCompDockPanel.Visibility = 'Visible'
                                        $syncHash.compUserDockPanel.Visibility = 'Collapsed'
                                        $syncHash.compExpander.IsExpanded = $true
                                        $syncHash.compExpanderProgressBar.Visibility = 'Hidden'
                                        #$syncHash.userCompGrid.UpdateLayout()
                                    })
                            }
                            else {
                                $syncHash.userCompGrid.Dispatcher.Invoke([Action] { $syncHash.compUserGrid.ItemsSource = ($queryHash[$currentTabItem]).LoginLogListView })
                                $syncHash.Window.Dispatcher.invoke([action] {
                                        $syncHash.userCompNotFound.Visibility = 'Hidden'
                                        $syncHash.userCompDockPanel.Visibility = 'Collapsed'
                                        $syncHash.compUserDockPanel.Visibility = 'Visible'
                                        $syncHash.compExpander.IsExpanded = $true
                                        $syncHash.compExpanderProgressBar.Visibility = 'Hidden'
                                        #$syncHash.compUserGrid.UpdateLayout()
                                    })
                            }            
                        }

                        else {
                            $syncHash.Window.Dispatcher.invoke([action] {
                                    $syncHash.userCompNotFound.Visibility = 'Visible'
                                    $syncHash.compUserDockPanel.Visibility = 'Collapsed'
                                    $syncHash.userCompDockPanel.Visibility = 'Collapsed'
                                    $syncHash.compExpander.IsExpanded = $true
                                    $syncHash.compExpanderProgressBar.Visibility = 'Hidden'
                                })
                        }
                    }
                    #update valuebox
                }
            }
        }
    })
        
$syncHash.itemToolGridExport.Add_Click({

    
    $configHash.gridExportList = [System.Collections.ArrayList]@()
    $syncHash.itemToolGridItemsGrid.Items | ForEach-Object { $configHash.gridExportList.Add($_) | Out-Null}


    $rsArgs = @{
            Name            = 'ExportGrid'           
            ModulesToImport = $configHash.modList 
            ArgumentList    =  $configHash
        }

    $syncHash.snackMsg.MessageQueue.Enqueue("Grid contents exporting...")

    Start-RSJob @rsArgs -ScriptBlock {
        Param ($configHash) 
       $configHash.gridExportList | Out-HtmlView
    }

})

$syncHash.tabControl.add_tabItemClosingEvent( {
        $queryHash.Remove($configHash.currentTabItem)

        if ($syncHash.tabControl.Items.Count -le 1) { Set-ItemExpanders -SyncHash $syncHash -ConfigHash $configHash -IsActive Disable }
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
            if ($null -like $syncHash.SearchBox.Text -and $_.Key -ne 'Escape') { $syncHash.SnackMsg.MessageQueue.Enqueue('Empty!') }
           
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
            ArgumentList    = @($syncHash.snackMsg.MessageQueue, $syncHash.itemToolDialogConfirmButton.Tag, $configHash, $queryHash, $syncHash.Window, $syncHash.adHocConfirmWindow, $varHash)
            ModulesToImport = $configHash.modList
        }

        Start-RSJob @rsArgs -ScriptBlock {
            Param($queue, $toolID, $configHash, $queryHash, $window, $confirmWindow, $varHash)

            $item = ($configHash.currentTabItem).toLower()
            $targetType = $queryHash[$item].ObjectClass -replace 'c', 'C' -replace 'u', 'U'
            $toolName = ($configHash.objectToolConfig[$toolID - 1].toolName).ToUpper()
            Set-CustomVariables -VarHash $varHash

            try {              
                 ([scriptblock]::Create($configHash.objectToolConfig[$toolID - 1].toolAction)).Invoke()     
                          
                if ($configHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -Status Success
                    Write-LogMessage -syncHashWindow $window -Path $configHash.actionlogPath -Message Succeed -ActionName $toolName -SubjectType 'Standalone' -ArrayList $configHash.actionLog
                }

                else {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -Status Success -SubjectName $item
                    Write-LogMessage -syncHashWindow $window -Path $configHash.actionlogPath -Message Succeed -SubjectName $item -ActionName $toolName -SubjectType $targetType -ArrayList $configHash.actionLog 
                }
            }
            catch {
                if ($configHash.objectToolConfig[$toolID - 1].objectType -eq 'Standalone') {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -Status Fail
                    Write-LogMessage -syncHashWindow $window -Path $configHash.actionlogPath -Message Fail -ActionName $toolName -ArrayList $configHash.actionLog -Error $_
                }
                else {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -Status Fail -SubjectName $item
                    Write-LogMessage -syncHashWindow $window -Path $configHash.actionlogPath -Message Fail -SubjectName $item -ActionName $toolName -SubjectType $targetType -ArrayList $configHash.actionLog -Error $_
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
        $syncHash.itemToolGridADSelectedItem.Content = $null
        $syncHash.itemToolGridExport.Visibility = 'Collapsed'
        $syncHash.itemToolImageBorder.Visibility = 'Collapsed'
        $syncHash.itemToolGridSelect.Visibility = 'Collapsed'
        $syncHash.itemToolCommandGridPanel.Visibility = 'Collapsed'
        $syncHash.itemToolCustomDialog.Visibility = 'Collapsed'
        $syncHash.itemToolGridSelectAllButton.Visibility = 'Collapsed'
        $syncHash.itemToolListSelectAllButton.Visibility = 'Collapsed'
    })

$syncHash.compExpanderOpenLog.Add_Click({  Invoke-Item $queryHash[$configHash.currentTabItem].LoginLogPath })

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



$syncHash.toolsLogDialogClose.Add_Click( { $syncHash.toolsLogDialog.IsOpen = $false })
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
            ArgumentList    = @($syncHash.snackMsg.MessageQueue, $configHash, $queryHash, $syncHash, $rsCmd, $syncHash.adHocConfirmWindow, $syncHash.Window, $varHash)
            ModulesToImport = $configHash.modList
        }

        Start-RSJob @rsArgs -ScriptBlock {
            Param($queue, $configHash, $queryHash, $syncHash, $rsCmd, $confirmWindow, $window)

            Set-CustomVariables -VarHash $varHash

            $item = $rsCmd.item
            $targetType = $queryHash[$item].ObjectClass -replace 'c', 'C' -replace 'u', 'U'

            try {
                foreach ($gridItem in ($rsCmd.cmdGridItems | Where-Object { $_.Result -ne 'True' })) {
                    $actionCmd = ($configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].toolCommandGridConfig |
                            Where-Object { $_.actionCmdValid -eq 'True' -and $_.queryCmdValid -eq 'True' })[$gridItem.Index].actionCmd
                
                    $queryCmd = ($configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].toolCommandGridConfig |
                            Where-Object { $_.actionCmdValid -eq 'True' -and $_.queryCmdValid -eq 'True' })[$gridItem.Index].queryCmd
                
                    $result = $gridItem.result

                    ([scriptblock]::Create($actionCmd)).Invoke() 

                    $result = (([scriptblock]::Create($queryCmd)).Invoke()).ToString()

                    $syncHash.Window.Dispatcher.Invoke([action] { $syncHash.itemToolCommandGridDataGrid.Items[$gridItem.Index].Result = $result })
                }

                $syncHash.Window.Dispatcher.Invoke([action] { $syncHash.itemToolCommandGridDataGrid.Items.Refresh() })

                if (($syncHash.itemToolCommandGridDataGrid.Items.Result -match 'False' | Measure-Object).Count -eq 0) { $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.toolsCommandGridExecuteAll.Tag = 'False' }) }
           
                $toolName = $configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].ToolName

                if ($configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].objectType -eq 'Standalone') {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -Status Success
                    Write-LogMessage -syncHashWindow $syncHash.window -Path $configHash.actionlogPath -Message Succeed -ActionName $toolName -SubjectType 'Standalone' -ArrayList $configHash.actionLog
                }
                
                else {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -SubtoolName $rsCmd.cmdGridItemName -Status Success -SubjectName $item
                    Write-LogMessage -syncHashWindow $syncHash.window -Path $configHash.actionlogPath -Message Succeed -SubjectName $item -SubjectType $targetType -ActionName $toolName -ArrayList $configHash.actionLog 
                }
            }    
         

            catch {
                if ($configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].objectType -eq 'Standalone') {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -Status Fail
                    Write-LogMessage -syncHashWindow $syncHash.window -Path $configHash.actionlogPath -SubjectType 'Standalone' -Message Fail -ActionName $toolName -ArrayList $configHash.actionLog -Error $_
                }
                else {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -Status Fail -SubjectName $item
                    Write-LogMessage -syncHashWindow $syncHash.window -Path $configHash.actionlogPath -Message Fail -SubjectName $item -SubjectType $targetType -ActionName $toolName -ArrayList $configHash.actionLog -Error $_
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

    
$syncHash.settingSelectTextOption.Add_Click({
     $configHash.objectToolConfig[($syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID - 1)].toolActionSelectAD = $false
     $configHash.objectToolConfig[($syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID - 1)].toolActionSelectOU = $false
     $syncHash.settingSelectOUOption.IsChecked = $false
     $syncHash.settingSelectADOption.IsChecked = $false
})

$syncHash.settingSelectOUOption.Add_Click({
     $configHash.objectToolConfig[($syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID - 1)].toolActionSelectAD = $false
     $configHash.objectToolConfig[($syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID - 1)].toolActionSelectCustom = $false
     $syncHash.settingSelectTextOption.IsChecked = $false
     $syncHash.settingSelectADOption.IsChecked = $false
})

$syncHash.settingSelectADOption.Add_Click({
     $configHash.objectToolConfig[($syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID - 1)].toolActionSelectCustom = $false
     $configHash.objectToolConfig[($syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID - 1)].toolActionSelectOU = $false
     $syncHash.settingSelectOUOption.IsChecked = $false
     $syncHash.settingSelectTextOption.IsChecked = $false
})









$syncHash.settingNameDialogClose.Add_Click( {
        $syncHash.settingNameDialog.IsOpen = $false
        $configHash.nameMapListView.Refresh()
    })

$syncHash.settingInfoDialogClose.Add_Click( { $syncHash.settingInfoDialog.IsOpen = $false })

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
        if ($searchVal) {
            $syncHash.tabControl.ItemsSource.RemoveAt($syncHash.tabControl.SelectedIndex)
            $syncHash.SearchBox.Tag = $searchVal
            $syncHash.SearchBox.Focus()
            $wshell = New-Object -ComObject wscript.shell
            $wshell.SendKeys('{ESCAPE}')
        }
    })

[System.Windows.RoutedEventHandler]$global:addItemClick = {
    if ($syncHash.settingUserPropGrid.Visibility -eq 'Visible') { $type = 'User' }

    else { $type = 'Comp' }

    $i = ($configHash.($type + 'PropList').Field |
            Sort-Object -Descending |
                Select-Object -First 1) + 1

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
            ValidCmd          = $null
            ValidAction1      = $null
            ValidAction2      = $null
            actionCmd2        = 'Do-Something -User $user'
            actionCmd2ToolTip = 'Action name'
            actionCmd2Icon    = $null
            actionCmd2Refresh = $false
            actionCmd2Multi   = $false
            actionCmdsEnabled = $true
            transCmdsEnabled  = $true
            actionCmd2Enabled = $false
            PropType          = $null
            actionList        = @('ReadOnly', 'ReadOnly-Raw', 'Editable', 'Editable-Raw', 'Actionable', 'Actionable-Raw', 'Editable-Actionable', 'Editable-Actionable-Raw')
            ActionName        = 'null'
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
            toolActionToolTip      = $null
            toolActionIcon         = $null
            toolActionSelectAD     = $false
            toolActionSelectOU     = $false
            toolActionSelectCustom = $false
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
            item               = ($configHash.currentTabItem).toLower()
        }

        $rsArgs = @{
            Name            = 'CommandGridRun'
            ArgumentList    = @($syncHash.snackMsg.MessageQueue, $configHash, $queryHash, $syncHash, $rsCmd, $syncHash.adHocConfirmWindow, $syncHash.Window)
            ModulesToImport = $configHash.modList
        }

        Start-RSJob @rsArgs -ScriptBlock {
            Param($queue, $configHash, $queryHash, $syncHash, $rsCmd, $confirmWindow, $window)

            $item = $rsCmd.item
            $targetType = $queryHash[$item].ObjectClass -replace 'c', 'C' -replace 'u', 'U'

            $actionCmd = $configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].toolCommandGridConfig[$rsCmd.cmdGridItemIndex].actionCmd
            $queryCmd = $configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].toolCommandGridConfig[$rsCmd.cmdGridItemIndex].queryCmd
            $toolName = $configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].toolCommandGridConfig[$rsCmd.cmdGridItemIndex].SetName

            try {
                $result = $rsCmd.Result                     
                ([scriptblock]::Create($actionCmd)).Invoke() 
                $result = (([scriptblock]::Create($queryCmd)).Invoke() ).ToString()
                
                $syncHash.Window.Dispatcher.Invoke([action] {
                        $syncHash.itemToolCommandGridDataGrid.SelectedItem.Result = $result
                        $syncHash.itemToolCommandGridDataGrid.Items.Refresh()               
                    })
               
                if ($configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].objectType -eq 'Standalone') {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -SubtoolName $rsCmd.cmdGridItemName -Status Success
                    Write-LogMessage -syncHashWindow $syncHash.window -Path $configHash.actionlogPath -Message Succeed -ActionName ($($($toolName) + '-' + $($rsCmd.cmdGridItemName))) -SubjectType 'Standalone' -ArrayList $configHash.actionLog
                }
                
                else {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -SubtoolName $rsCmd.cmdGridItemName -Status Success -SubjectName $item
                    Write-LogMessage -syncHashWindow $syncHash.window -Path $configHash.actionlogPath -Message Succeed -SubjectName $item -SubjectType $targetType -ActionName ($($($toolName) + '-' + $($rsCmd.cmdGridItemName))) -ArrayList $configHash.actionLog 
                }
                
                if (($syncHash.itemToolCommandGridDataGrid.Items.Result -notmatch 'True' | Measure-Object).Count -eq 0) { $syncHash.Window.Dispatcher.Invoke([Action] { $syncHash.toolsCommandGridExecuteAll.Tag = 'False' }) }                                          
            }

            catch {
                if ($configHash.objectToolConfig[$rsCmd.cmdGridParentIndex].objectType -eq 'Standalone') {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -SubtoolName $rsCmd.cmdGridItemName -Status Fail
                    Write-LogMessage -syncHashWindow $syncHash.window -Path $configHash.actionlogPath -Message Fail -ActionName ($($($toolName) + '-' + $($rsCmd.cmdGridItemName))) -ArrayList $configHash.actionLog -Error $_
                }
                else {
                    Write-SnackMsg -Queue $queue -ToolName $toolName -SubtoolName $rsCmd.cmdGridItemName -Status Fail -SubjectName $item
                    Write-LogMessage -syncHashWindow $syncHash.window -Path $configHash.actionlogPath -Message Fail -SubjectName $item -SubjectType $targetType -ActionName ($($($toolName) + '-' + $($rsCmd.cmdGridItemName))) -SubjectType 'Standalone' -ArrayList $configHash.actionLog -Error $_
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
        $syncHash.settingCommandGridToolIcon.ItemsSource = $configHash.buttonGlyphs
        $syncHash.settingChildWindow.TitleBarBackground = ($syncHash.settingObjectToolDefFlyout.Background.Color).ToString()
        $syncHash.settingChildWindow.Title = 'Define Tool'
        $syncHash.settingChildWindow.ShowCloseButton = $false
        Set-CurrentPane -SyncHash $syncHash -Panel 'settingObjectToolDefFlyout'
    }

    elseif ($button.Name -match 'settingObjectToolsClearItem') {
        $num = $syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID 
        $configHash.objectToolConfig.RemoveAt($syncHash.settingObjectToolsPropGrid.SelectedItem.ToolID - 1)
        $configHash.objectToolConfig |
            Where-Object { $_.ToolID -gt $num } |
                ForEach-Object { $_.ToolID = $_.ToolID - 1 }       
        $syncHash.settingObjectToolsPropGrid.Items.Refresh()
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
        $searchVal = $syncHash.historySideDataGrid.SelectedItem.SubjectName       
        $syncHash.SearchBox.Tag = $searchVal
        $syncHash.historySidePane.IsOpen = $false
        $syncHash.SearchBox.Focus()
        $wshell = New-Object -ComObject wscript.shell
        $wshell.SendKeys('{ESCAPE}')
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
        $id = $syncHash.settingModDataGrid.SelectedItem.VarNum 
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


$syncHash.Window.ShowDialog() | Out-Null
$configHash.IsClosed = $true

