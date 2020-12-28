
function Resolve-Location {
    [CmdletBinding(DefaultParameterSetName = 'HostName')]
    param (

        [parameter(ValueFromPipeLine, Mandatory = $true, ParameterSetName = 'HostName')][string]$computerName,
        [parameter(Mandatory = $true, ParameterSetName = "IP")][IPAddress]$IPAddress,
        [parameter(Mandatory = $true)][PSCustomObject]$IPList

    )

    Begin {
        $ipMap = [System.Collections.ArrayList]@()
    }

    Process {
        if ($computerName) {
            try { $IPAddress = (Resolve-DnsName $computerName -Type A -ErrorAction Stop)[0].IPAddress }
            catch { Write-Error -Message "No valid DNS record for $computerName" }
        }

        else {
            try { $computername = (Resolve-DnsName $IPAddress -Type PTR -ErrorAction Stop)[0].HostName }
            catch { Write-Error -Message "No valid DNS record for $IPAddress" }
        }

        if ($computerName -and $IPAddress) {
               $ipMap.Add(($IPList | Where-Object {([System.Net.IpAddress]($_.Network.ToString())).Address -eq (([System.Net.IpAddress]$IPAddress).Address -band ([ipaddress]([uint32]::MaxValue-[math]::Pow(2,32-$_.Mask)+1)).Address)} |
                Select-Object *, @{Label = 'Computername'; Expression = { $computerName }}, @{Label = 'IPv4'; Expression = {$IPAddress }} -ErrorAction Stop)) | Out-Null
        }   
    }

    End {
        $ipMap
    }
}
function Test-OnlineFast {
    param
    (
        # make parameter pipeline-aware
        [Parameter(Mandatory,ValueFromPipeline)]
        [string[]]
        $ComputerName,

        $TimeoutMillisec = 1000
    )

    begin
    {
        # use this to collect computer names that were sent via pipeline
        [Collections.ArrayList]$bucket = @()
    
        # hash table with error code to text translation
        $StatusCode_ReturnValue = 
        @{
            0='Success'
            11001='Buffer Too Small'
            11002='Destination Net Unreachable'
            11003='Destination Host Unreachable'
            11004='Destination Protocol Unreachable'
            11005='Destination Port Unreachable'
            11006='No Resources'
            11007='Bad Option'
            11008='Hardware Error'
            11009='Packet Too Big'
            11010='Request Timed Out'
            11011='Bad Request'
            11012='Bad Route'
            11013='TimeToLive Expired Transit'
            11014='TimeToLive Expired Reassembly'
            11015='Parameter Problem'
            11016='Source Quench'
            11017='Option Too Big'
            11018='Bad Destination'
            11032='Negotiating IPSEC'
            11050='General Failure'
        }
    
    
        # hash table with calculated property that translates
        # numeric return value into friendly text

        $statusFriendlyText = @{
            # name of column
            Name = 'Status'
            # code to calculate content of column
            Expression = { 
                # take status code and use it as index into
                # the hash table with friendly names
                # make sure the key is of same data type (int)
                $StatusCode_ReturnValue[([int]$_.StatusCode)]
            }
        }

        # calculated property that returns $true when status -eq 0
        $IsOnline = @{
            Name = 'Online'
            Expression = { $_.StatusCode -eq 0 }
        }

        # do DNS resolution when system responds to ping
        $DNSName = @{
            Name = 'DNSName'
            Expression = { if ($_.StatusCode -eq 0) { 
                    if ($_.Address -like '*.*.*.*') 
                    { [Net.DNS]::GetHostByAddress($_.Address).HostName  } 
                    else  
                    { [Net.DNS]::GetHostByName($_.Address).HostName  } 
                }
            }
        }
    }
    
    process
    {
        # add each computer name to the bucket
        # we either receive a string array via parameter, or 
        # the process block runs multiple times when computer
        # names are piped
        $ComputerName | ForEach-Object {
            $null = $bucket.Add($_)
        }
    }
    
    end
    {
        # convert list of computers into a WMI query string
        $query = $bucket -join "' or Address='"
        
        Get-WmiObject -Class Win32_PingStatus -Filter "(Address='$query') and timeout=$TimeoutMillisec"|
        Select-Object -Property Address, IPV4Address, $IsOnline, $DNSName, $statusFriendlyText
    }
    
}
function New-FolderSelection {
    Param (
        [Parameter(Position = 0)]$Title)

    $AssemblyFullName = 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
    $Assembly = [System.Reflection.Assembly]::Load($AssemblyFullName)
   
    $openFolderDialog = [System.Windows.Forms.OpenFileDialog] @{
        Filter           = "Folders|`n"
        AddExtension     = $false
        CheckFileExists  = $false
        DereferenceLinks = $true
        Multiselect      = $false
        Title            = $Title
    }

  
    $FileDialogInterfaceType = $Assembly.GetType('System.Windows.Forms.FileDialogNative+IFileDialog')
    $IFileDialog = ($openFolderDialog.GetType()).GetMethod('CreateVistaDialog', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($openFolderDialog, $null)
    $null = ($openFolderDialog.GetType()).GetMethod('OnBeforeVistaDialog', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($openFolderDialog, $IFileDialog)
    [uint32]$PickFoldersOption = $Assembly.GetType('System.Windows.Forms.FileDialogNative+FOS').GetField('FOS_PICKFOLDERS').GetValue($null)
    $FolderOptions = ($openFolderDialog.GetType()).GetMethod('get_Options', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($openFolderDialog, $null) -bor $PickFoldersOption
    $null = $FileDialogInterfaceType.GetMethod('SetOptions', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, $FolderOptions)
    $VistaDialogEvent = [System.Activator]::CreateInstance($AssemblyFullName, 'System.Windows.Forms.FileDialog+VistaDialogEvents', $false, 0, $null, $openFolderDialog, $null, $null).Unwrap()
    [uint32]$AdviceCookie = 0
    $AdvisoryParameters = @($VistaDialogEvent, $AdviceCookie)
    $AdviseResult = $FileDialogInterfaceType.GetMethod('Advise', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, $AdvisoryParameters)

    $AdviceCookie = $AdvisoryParameters[1]

    $Result = $FileDialogInterfaceType.GetMethod('Show', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, [System.IntPtr]::Zero)

    $null = $FileDialogInterfaceType.GetMethod('Unadvise', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, $AdviceCookie)

    if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
        $FileDialogInterfaceType.GetMethod('GetResult', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, $null)
    }

    $openFolderDialog.FileName

}
Function Get-Icon {
    [cmdletbinding(
        DefaultParameterSetName = '__DefaultParameterSetName'
    )]
    Param (
        [parameter(ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullorEmpty()]
        [string]$Path,
        [parameter(ParameterSetName = 'Bytes')]
        [switch]$ToBytes,
        [parameter(ParameterSetName = 'Bitmap')]
        [switch]$ToBitmap,
        [parameter(ParameterSetName = 'Base64')]
        [switch]$ToBase64
    )
    Begin {
        If ($PSBoundParameters.ContainsKey('Debug')) {
            $DebugPreference = 'Continue'
        }
        Add-Type -AssemblyName System.Drawing
    }
    Process {
        $Path = Convert-Path -Path $Path
        Write-Debug $Path
        If (Test-Path -Path $Path) {
            $Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($Path) | 
                Add-Member -MemberType NoteProperty -Name FullName -Value $Path -PassThru
            If ($PSBoundParameters.ContainsKey('ToBytes')) {
                Write-Verbose "Retrieving bytes"
                $MemoryStream = New-Object System.IO.MemoryStream
                $Icon.save($MemoryStream)
                Write-Debug ($MemoryStream | Out-String)
                $MemoryStream.ToArray()   
                $MemoryStream.Flush()  
                $MemoryStream.Dispose()           
            }
            ElseIf ($PSBoundParameters.ContainsKey('ToBitmap')) {
                $Icon.ToBitMap()
            }
            ElseIf ($PSBoundParameters.ContainsKey('ToBase64')) {
                $MemoryStream = New-Object System.IO.MemoryStream
                $Icon.save($MemoryStream)
                Write-Debug ($MemoryStream | Out-String)
                $Bytes = $MemoryStream.ToArray()   
                $MemoryStream.Flush() 
                $MemoryStream.Dispose()
                [convert]::ToBase64String($Bytes)
            }
            Else {
                $Icon
            }
        }
        Else {
            Write-Warning "$Path does not exist!"
            Continue
        }
    }
}
function Select-ADObject {

Param (
        
    [Parameter(Mandatory=$false)]
    [Switch]
    $MultiSelect,

    [Parameter()]
    [ValidateSet('All','Users','Computers','Groups','UsersComputers')]
    [string[]]
    $Type,

    [Parameter()]
    [string[]]
    $SearchBase
    )

    Begin{

        $DialogPicker = New-Object CubicOrange.Windows.Forms.ActiveDirectory.DirectoryObjectPickerDialog

        if ($type -eq 'UsersComputers') {
            $DialogPicker.AllowedObjectTypes = [CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::Users, [CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::Computers
            $DialogPicker.DefaultObjectTypes = [CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::Users, [CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::Computers

        }
        
        elseif ($type -and $type -ne 'All') { 
                $DialogPicker.AllowedObjectTypes = [CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::$type 
                $DialogPicker.DefaultObjectTypes = [CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::$type
        }

        else {           
            $DialogPicker.AllowedObjectTypes = [CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::Groups, [CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::Users, [CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::Computers
            $DialogPicker.DefaultObjectTypes = [CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::Users
        }


        $DialogPicker.AllowedLocations = [CubicOrange.Windows.Forms.ActiveDirectory.Locations]::JoinedDomain       
        $DialogPicker.DefaultLocations = [CubicOrange.Windows.Forms.ActiveDirectory.Locations]::UserEntered       
        $DialogPicker.ShowAdvancedView = $true    
        $DialogPicker.SkipDomainControllerCheck = $true
        $DialogPicker.Providers = [CubicOrange.Windows.Forms.ActiveDirectory.ADsPathsProviders]::Default

        if ($MultiSelect) {
            $DialogPicker.MultiSelect = $true
        }

        $DialogPicker.AttributesToFetch.Add('samAccountName')

    }

    Process {
        $DialogPicker.ShowDialog()

    }
    
    End{

        if ($MultiSelect) {
            return $DialogPicker.SelectedObjects
        }

        else {
            $DialogPicker.Selectedobject
        }

    }
}
function Get-RDSession {
  

    [CmdletBinding()]
    [OutputType('Cassia.Impl.TerminalServicesSession')]

    param(
        [Parameter(
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('DSNHostName', 'Name', 'Computer')]
        [string]$ComputerName = 'localhost',

        [Parameter()]
        [ValidateSet('Active', 'Connected', 'ConnectQuery', 'Disconnected', 'Down', 'Idle', 'Initializing', 'Listening', 'Reset', 'Shadowing')]
        [Alias('ConnectionState')]
        [string]$State = '*',

        [Parameter()]
        [string]$ClientName = '*',

        [Parameter()]
        [string]$UserName = '*'
    )

    begin {
        try {
            Write-Verbose -Message 'Creating instance of the Cassia TSManager.'
            $TSManager = New-Object -TypeName Cassia.TerminalServicesManager
        }
        catch {
            throw
        }
    }

    process {
        Write-Verbose -Message ($LocalizedData.RemoteConnect -f $ComputerName)
        if (!(Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
            Write-Warning -Message ($LocalizedData.ComputerOffline -f $ComputerName)
            return
        }
        try {
            $TSRemoteServer = $TSManager.GetRemoteServer($ComputerName)
            $TSRemoteServer.Open()
            if (!($TSRemoteServer.IsOpen)) {
                throw ($LocalizedData.RemoteConnectError -f $ComputerName)
            }

            $Session = $TSRemoteServer.GetSessions()
            if ($Session) {
                $Session | Where-Object { $_.ConnectionState -like $State -and $_.UserName -like $UserName -and $_.ClientName -like $ClientName } |
                    Add-Member -MemberType AliasProperty -Name IPAddress -Value ClientIPAddress -PassThru |
                    Add-Member -MemberType AliasProperty State -Value ConnectionState -PassThru
            }
        }
        catch {
            throw
        }
        finally {
            $TSRemoteServer.Close()
            $TSRemoteServer.Dispose()
        }
    }

    end {}
}