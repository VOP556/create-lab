#Requires -Version 5
#Requires -RunasAdministrator
#.Author Joerg Zimmermann
#.License GPL
#.SYNOPSIS Deployment Framework 
#.Example Workflow Step 1
#   $Node = init-deployment
#   - gets the Configuration from Excel and builds the MyConfiguration-Object with all Nodes
#   - prepares the local FileSystem
#   - creates MasterImage, VHDs and VMs needed for Deployment
#   - installation of the first DomainController and a Windows-Deployment-Services and DHCP Server for Deployment of BareMetal-Machines
#   - configuration of DHCP and WDS
#.Example Workflow Reset
#   reset-deployment
#   - removes all Virtual Machines, Virtual Harddisks of the Deployment and clears the local hosts-file
param (
    $WORKDIR = "c:\deploy\scripts",
    $ConfigurationPath = ".\configuration.xlsx"
 )

#region Includes and imports
 . "C:\deploy\scripts\Convert-WindowsImage.ps1"
 . "C:\deploy\scripts\Set-AutoLogon.ps1"
 Import-Module Hyper-V -ErrorAction SilentlyContinue

#region Helperfunctions

        function ConvertTo-BinaryIP {
            <#
            .Synopsis
                Converts a Decimal IP address into a binary format.
            .Description
                ConvertTo-BinaryIP uses System.Convert to switch between decimal and binary format. The output from this function is dotted binary.
            .Parameter IPAddress
                An IP Address to convert.
            #>
 
            [CmdLetBinding()]
            param(
            [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
            [Net.IPAddress]$IPAddress
            )
 
            process {  
            return [String]::Join('.', $( $IPAddress.GetAddressBytes() |
                ForEach-Object { [Convert]::ToString($_, 2).PadLeft(8, '0') } ))
            }
        }
        function ConvertTo-DecimalIP {
            <#
            .Synopsis
              Converts a Decimal IP address into a 32-bit unsigned integer.
            .Description
              ConvertTo-DecimalIP takes a decimal IP, uses a shift-like operation on each octet and returns a single UInt32 value.
            .Parameter IPAddress
              An IP Address to convert.
          #>
  
          [CmdLetBinding()]
          param(
            [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
            [Net.IPAddress]$IPAddress
          )
 
          process {
            $i = 3; $DecimalIP = 0;
            $IPAddress.GetAddressBytes() | ForEach-Object { $DecimalIP += $_ * [Math]::Pow(256, $i); $i-- }
 
            return [UInt32]$DecimalIP
          }
        }
        function ConvertTo-DottedDecimalIP {
            <#
            .Synopsis
                Returns a dotted decimal IP address from either an unsigned 32-bit integer or a dotted binary string.
            .Description
                ConvertTo-DottedDecimalIP uses a regular expression match on the input string to convert to an IP address.
            .Parameter IPAddress
                A string representation of an IP address from either UInt32 or dotted binary.
            #>
 
            [CmdLetBinding()]
            param(
            [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
            [String]$IPAddress
            )
  
            process {
            Switch -RegEx ($IPAddress) {
                "([01]{8}.){3}[01]{8}" {
                return [String]::Join('.', $( $IPAddress.Split('.') | ForEach-Object { [Convert]::ToUInt32($_, 2) } ))
                }
                "\d" {
                $IPAddress = [UInt32]$IPAddress
                $DottedIP = $( For ($i = 3; $i -gt -1; $i--) {
                    $Remainder = $IPAddress % [Math]::Pow(256, $i)
                    ($IPAddress - $Remainder) / [Math]::Pow(256, $i)
                    $IPAddress = $Remainder
                    } )
       
                return [String]::Join('.', $DottedIP)
                }
                default {
                Write-Error "Cannot convert this format"
                }
            }
            }
        }
        function ConvertTo-MaskLength {
          <#
            .Synopsis
              Returns the length of a subnet mask.
            .Description
              ConvertTo-MaskLength accepts any IPv4 address as input, however the output value 
              only makes sense when using a subnet mask.
            .Parameter SubnetMask
              A subnet mask to convert into length
          #>
 
          [CmdLetBinding()]
          param(
            [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
            [Alias("Mask")]
            [Net.IPAddress]$SubnetMask
          )
 
          process {
            $Bits = "$( $SubnetMask.GetAddressBytes() | ForEach-Object { [Convert]::ToString($_, 2) } )" -replace '[\s0]'
 
            return $Bits.Length
          }
        }
        function ConvertTo-Mask {
          <#
            .Synopsis
              Returns a dotted decimal subnet mask from a mask length.
            .Description
              ConvertTo-Mask returns a subnet mask in dotted decimal format from an integer value ranging 
              between 0 and 32. ConvertTo-Mask first creates a binary string from the length, converts 
              that to an unsigned 32-bit integer then calls ConvertTo-DottedDecimalIP to complete the operation.
            .Parameter MaskLength
              The number of bits which must be masked.
          #>
  
          [CmdLetBinding()]
          param(
            [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
            [Alias("Length")]
            [ValidateRange(0, 32)]
            $MaskLength
          )
  
          Process {
            return ConvertTo-DottedDecimalIP ([Convert]::ToUInt32($(("1" * $MaskLength).PadRight(32, "0")), 2))
          }
        }
        function Get-NetworkAddress {
          <#
            .Synopsis
              Takes an IP address and subnet mask then calculates the network address for the range.
            .Description
              Get-NetworkAddress returns the network address for a subnet by performing a bitwise AND 
              operation against the decimal forms of the IP address and subnet mask. Get-NetworkAddress 
              expects both the IP address and subnet mask in dotted decimal format.
            .Parameter IPAddress
              Any IP address within the network range.
            .Parameter SubnetMask
              The subnet mask for the network.
          #>
  
          [CmdLetBinding()]
          param(
            [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
            [Net.IPAddress]$IPAddress,
    
            [Parameter(Mandatory = $true, Position = 1)]
            [Alias("Mask")]
            [Net.IPAddress]$SubnetMask
          )
 
          process {
            return ConvertTo-DottedDecimalIP ((ConvertTo-DecimalIP $IPAddress) -band (ConvertTo-DecimalIP $SubnetMask))
          }
        }
        function Get-BroadcastAddress {
          <#
            .Synopsis
              Takes an IP address and subnet mask then calculates the broadcast address for the range.
            .Description
              Get-BroadcastAddress returns the broadcast address for a subnet by performing a bitwise AND 
              operation against the decimal forms of the IP address and inverted subnet mask. 
              Get-BroadcastAddress expects both the IP address and subnet mask in dotted decimal format.
            .Parameter IPAddress
              Any IP address within the network range.
            .Parameter SubnetMask
              The subnet mask for the network.
          #>
  
          [CmdLetBinding()]
          param(
            [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
            [Net.IPAddress]$IPAddress, 
    
            [Parameter(Mandatory = $true, Position = 1)]
            [Alias("Mask")]
            [Net.IPAddress]$SubnetMask
          )
 
          process {
            return ConvertTo-DottedDecimalIP $((ConvertTo-DecimalIP $IPAddress) -bor `
              ((-bnot (ConvertTo-DecimalIP $SubnetMask)) -band [UInt32]::MaxValue))
          }
        }
        function Get-NetworkSummary ( [String]$IP, [String]$Mask ) {
          if ($IP.Contains("/")) {
            $Temp = $IP.Split("/")
            $IP = $Temp[0]
            $Mask = $Temp[1]
          }
 
          if (!$Mask.Contains(".")) {
            $Mask = ConvertTo-Mask $Mask
          }
 
          $DecimalIP = ConvertTo-DecimalIP $IP
          $DecimalMask = ConvertTo-DecimalIP $Mask
  
          $Network = $DecimalIP -band $DecimalMask
          $Broadcast = $DecimalIP -bor `
            ((-bnot $DecimalMask) -band [UInt32]::MaxValue)
          $NetworkAddress = ConvertTo-DottedDecimalIP $Network
          $RangeStart = ConvertTo-DottedDecimalIP ($Network + 1)
          $RangeEnd = ConvertTo-DottedDecimalIP ($Broadcast - 1)
          $BroadcastAddress = ConvertTo-DottedDecimalIP $Broadcast
          $MaskLength = ConvertTo-MaskLength $Mask
  
          $BinaryIP = ConvertTo-BinaryIP $IP; $Private = $False
          switch -regex ($BinaryIP) {
            "^1111"  { $Class = "E"; $SubnetBitMap = "1111"; break }
            "^1110"  { $Class = "D"; $SubnetBitMap = "1110"; break }
            "^110"   { 
              $Class = "C"
              if ($BinaryIP -match "^11000000.10101000") { $Private = $true }
              break
            }
            "^10"    { 
              $Class = "B"
              if ($BinaryIP -match "^10101100.0001") { $Private = $true }
              break
            }
            "^0"     { 
              $Class = "A" 
              if ($BinaryIP -match "^0000101") { $Private = $true }
            }
          }   
   
          $NetInfo = New-Object Object
          Add-Member NoteProperty "Network" -Input $NetInfo -Value $NetworkAddress
          Add-Member NoteProperty "Broadcast" -Input $NetInfo -Value $BroadcastAddress
          Add-Member NoteProperty "Range" -Input $NetInfo -Value "$RangeStart - $RangeEnd"
          Add-Member NoteProperty "Mask" -Input $NetInfo -Value $Mask
          Add-Member NoteProperty "MaskLength" -Input $NetInfo -Value $MaskLength
          Add-Member NoteProperty "Hosts" -Input $NetInfo -Value $($Broadcast - $Network - 1)
          Add-Member NoteProperty "Class" -Input $NetInfo -Value $Class
          Add-Member NoteProperty "IsPrivate" -Input $NetInfo -Value $Private
  
          return $NetInfo
        }
        function Get-NetworkRange( [String]$IP, [String]$Mask ) {
          if ($IP.Contains("/")) {
            $Temp = $IP.Split("/")
            $IP = $Temp[0]
            $Mask = $Temp[1]
          }
 
          if (!$Mask.Contains(".")) {
            $Mask = ConvertTo-Mask $Mask
          }
 
          $DecimalIP = ConvertTo-DecimalIP $IP
          $DecimalMask = ConvertTo-DecimalIP $Mask
  
          $Network = $DecimalIP -band $DecimalMask
          $Broadcast = $DecimalIP -bor ((-bnot $DecimalMask) -band [UInt32]::MaxValue)
 
          for ($i = $($Network + 1); $i -lt $Broadcast; $i++) {
            ConvertTo-DottedDecimalIP $i
          }
        }

 function Convert-MacAddress {
 <#
        .SYNOPSIS
                Converts a MAC address from one valid format to another.
 
        .DESCRIPTION
                The Convert-MacAddress function takes a valid hex MAC address and converts it to another valid hex format.
                Valid formats include the colon, dash, and dot delimiters as well as a raw address with no delimiter.
 
        .PARAMETER MacAddress
                Specifies the MAC address to be converted.
 
        .PARAMETER Delimiter
                Specifies a valid MAC address delimiting character. The format specified by the delimiter determines the conversion of the input string.
                Default value: ':'
 
        .EXAMPLE
                Convert-MacAddress 012345abcdef
                Converts the MAC address '012345abcdef' to '01:23:45:ab:cd:ef'.
 
        .EXAMPLE
                Convert-MacAddress 0123.45ab.cdef
                Converts the MAC address '0123.45ab.cdef' to '01:23:45:ab:cd:ef'.
               
        .EXAMPLE
                Convert-MacAddress 01:23:45:ab:cd:ef -Delimiter .
                Converts the MAC address '01:23:45:ab:cd:ef' to '0123.45ab.cdef'.
 
        .EXAMPLE
                Convert-MacAddress 01:23:45:ab:cd:ef -Delimiter ""
                Converts the dotted MAC address '01:23:45:ab:cd:ef' to '012345abcdef'.
 
        .INPUTS
                Sysetm.String
 
        .OUTPUTS
                System.String
 
        .NOTES
                Name: Convert-MacAddress
                Author: Rich Kusak
                Created: 2011-08-28
                LastEdit: 2011-08-29 10:02
                Version: 1.0.0.0
 
        .LINK
                http://en.wikipedia.org/wiki/MAC_address
       
        .LINK
                about_regular_expressions
 
 #>
 
        [CmdletBinding()]
        param (
                [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
                [ValidateScript({
                        $patterns = @(
                                '^([0-9a-f]{2}:){5}([0-9a-f]{2})$'
                                '^([0-9a-f]{2}-){5}([0-9a-f]{2})$'
                                '^([0-9a-f]{4}.){2}([0-9a-f]{4})$'
                                '^([0-9a-f]{12})$'
                        )
                        if ($_ -match ($patterns -join '|')) {$true} else {
                                throw "The argument '$_' does not match a valid MAC address format."
                        }
                })]
                [string]$MacAddress,
               
                [Parameter(ValueFromPipelineByPropertyName=$true)]
                [ValidateSet(':', '-', '.', $null)]
                [string]$Delimiter = ':'
        )
       
        process {
 
                $rawAddress = $MacAddress -replace '\W'
               
                switch ($Delimiter) {
                        {$_ -match ':|-'} {
                                for ($i = 2 ; $i -le 14 ; $i += 3) {
                                        $result = $rawAddress = $rawAddress.Insert($i, $_)
                                }
                                break
                        }
 
                        '.' {
                                for ($i = 4 ; $i -le 9 ; $i += 5) {
                                        $result = $rawAddress = $rawAddress.Insert($i, $_)
                                }
                                break
                        }
                       
                        default {
                                $result = $rawAddress
                        }
                } # switch
               
                $result
        } # process
 }
 function get-ExcelDataHashTable {
    #written by Jörg Zimmermann
    #www.burningmountain.de
    #imports all worksheets from excel into hashtables
    #define the header per parameter
    #define the first row per parameter
    param(
        [Parameter(Mandatory=$true)][string]$path,
        [int]$HeaderRow=3,
        [int]$FirstColumn=1
    )
    $path = (Get-Item -Path $path).FullName
    $DataHashTable = @{}
    $objExcel = New-Object -ComObject Excel.Application
    $WorkBook = $objExcel.Workbooks.Open($path)
    $WorkSheets = $WorkBook.sheets 
    $WorkSheets | ForEach-Object {
        $WorkSheet = $_
        $WorkSheetKey = $WorkSheet.Name
        $WorkSheetValue=@{}
        $Row=$HeaderRow+1
        while($WorkSheet.Cells($Row,$FirstColumn).text -ne ""){
            $ItemKey = $WorkSheet.Cells($Row, $FirstColumn).text
            $ItemValue=@{}
            $Column=$FirstColumn
            while($WorkSheet.Cells($HeaderRow, $Column).text -ne ""){
                $ItemDataKey = $WorkSheet.Cells($HeaderRow, $Column).text
                $ItemDataValue = $WorkSheet.Cells($Row,$Column).text
                $ItemValue.add($ItemDataKey,$ItemDataValue)
                $Column++
            }
            $WorkSheetValue.add($ItemKey,$ItemValue)
            $Row++
        }
      $DataHashTable.Add($WorkSheetKey,$WorkSheetValue)        
    }
    $WorkBook.Save()
    $WorkBook.Close()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WorkBook) | Out-Null
    $DataHashTable
 }

 function Set-VMNetworkConfiguration {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   Position=1,
                   ParameterSetName='DHCP',
                   ValueFromPipeline=$true)]
        [Parameter(Mandatory=$true,
                   Position=0,
                   ParameterSetName='Static',
                   ValueFromPipeline=$true)]
        [Microsoft.HyperV.PowerShell.VMNetworkAdapter]$NetworkAdapter,

        [Parameter(Mandatory=$true,
                   Position=1,
                   ParameterSetName='Static')]
        [String[]]$IPAddress=@(),

        [Parameter(Mandatory=$false,
                   Position=2,
                   ParameterSetName='Static')]
        [String[]]$Subnet=@(),

        [Parameter(Mandatory=$false,
                   Position=3,
                   ParameterSetName='Static')]
        [String[]]$DefaultGateway = @(),

        [Parameter(Mandatory=$false,
                   Position=4,
                   ParameterSetName='Static')]
        [String[]]$DNSServer = @(),

        [Parameter(Mandatory=$false,
                   Position=0,
                   ParameterSetName='DHCP')]
        [Switch]$Dhcp
    )

    $VM = Get-WmiObject -Namespace 'root\virtualization\v2' -Class 'Msvm_ComputerSystem' | Where-Object { $_.ElementName -eq $NetworkAdapter.VMName } 
    $VMSettings = $vm.GetRelated('Msvm_VirtualSystemSettingData') | Where-Object { $_.VirtualSystemType -eq 'Microsoft:Hyper-V:System:Realized' }    
    $VMNetAdapters = $VMSettings.GetRelated('Msvm_SyntheticEthernetPortSettingData') 

    $NetworkSettings = @()
    foreach ($NetAdapter in $VMNetAdapters) {
        if ($NetAdapter.Address -eq $NetworkAdapter.MacAddress) {
            $NetworkSettings = $NetworkSettings + $NetAdapter.GetRelated("Msvm_GuestNetworkAdapterConfiguration")
        }
    }

    $NetworkSettings[0].IPAddresses = $IPAddress
    $NetworkSettings[0].Subnets = $Subnet
    $NetworkSettings[0].DefaultGateways = $DefaultGateway
    $NetworkSettings[0].DNSServers = $DNSServer
    $NetworkSettings[0].ProtocolIFType = 4096

    if ($dhcp) {
        $NetworkSettings[0].DHCPEnabled = $true
    } else {
        $NetworkSettings[0].DHCPEnabled = $false
    }

    $Service = Get-WmiObject -Class "Msvm_VirtualSystemManagementService" -Namespace "root\virtualization\v2"
    $setIP = $Service.SetGuestNetworkAdapterConfiguration($VM, $NetworkSettings[0].GetText(1))

    if ($setip.ReturnValue -eq 4096) {
        $job=[WMI]$setip.job 

        while ($job.JobState -eq 3 -or $job.JobState -eq 4) {
            start-sleep 1
            $job=[WMI]$setip.job
        }

        if ($job.JobState -eq 7) {
            write-host "Success"
        }
        else {
            $job.GetError()
        }
    } elseif($setip.ReturnValue -eq 0) {
        Write-Verbose "Success"
    }
 }

#region Class-Definitions
class Network {

    [STRING]$IPAddress

#region Constructors 
    Network([String]$IPAddress) {
        $this.IPAddress = $IPAddress
    }

    Network([Hashtable]$Hash){
        $Hash.Keys | ForEach-Object {
            $this | Add-Member -MemberType NoteProperty -Name $_ -Value $Hash.$_ -Force
        }
    }
#region methods
    [hashtable] getConfigurationData(){
        $Hash = @{}
        $this | Get-Member -MemberType NoteProperty,Property | ForEach-Object{
            $Hash.Add($_.Name,$this.($_.Name))
        }
        return $Hash
    }
 }




class Node {
  static [scriptblock]$StartScript={
  
    $Ende={
        write-host "Konfiguration beendet"
        pause
    }
    $Start1 = {
           if ($Node.Description -match "x3750M4"){
                #Write-Warning "Inteltreiber werden installiert"
                #Start-Process -Wait -FilePath C:\deploy\software\Intel_treiber_06-15\APPS\PROSETDX\Winx64\DxSetup.exe -Verb runas -ArgumentList " /qn" | Wait-Process
                #Start-Sleep -Seconds 10
                Write-Warning "Broadcom-Treiber werden installiert"
                Start-Process -Wait -FilePath C:\deploy\software\BroadcomNetXtremeII\setup.exe -Verb runas -ArgumentList " /s /v/qn" | Wait-Process
                Start-Sleep -Seconds 10
                Get-NetAdapter | Format-Table -AutoSize
           }
           $MGMTNetwork = $Node.Networks | Where-Object {$_.name -like "Mgmt*"}
           Get-NetAdapter | Where-Object {$_.MacAddress -eq $Node.BootMac } | Rename-NetAdapter -NewName $MGMTNetwork.Name
           $MGMT= (Get-NetAdapter -Name $MGMTNetwork.Name)
           $MGMT | Get-NetAdapterBinding -ComponentID ms_tcpip6 | Disable-NetAdapterBinding
           Set-DnsClientServerAddress -InterfaceIndex $MGMT.InterfaceIndex -ServerAddresses $MGMTNetwork.DNSServer
           Get-ScheduledTask ScheduledDefrag | Disable-ScheduledTask
    }

    $wds =@(
        {
            . C:\deploy\scripts\create-lab.ps1
            Install-WindowsFeature WDS -IncludeAllSubFeature -IncludeManagementTools
            set-autologon -defaultusername ($Node.SMBDomain+"\Administrator") -defaultpassword $Node.Password -script 'powershell.exe -noexit -command {
            . C:\deploy\scripts\create-lab.ps1
            #WDS initialisieren            
            wdsutil.exe /initialize-server /reminst:c:\remoteinstall

            #ImageGroup erstellen
            $ImageGroup = (new-wdsinstallimagegroup -name "Deploy" -erroraction silentlycontinue).Name
            $ImageGroup = (get-wdsinstallimagegroup -name "Deploy").Name
        
            #Unattend WDS Setup bearbeiten und in c:\remoteinstall\wds_boot.xml speichern
            
            $BootUnattend="c:\remoteinstall\wds_boot.xml"
            $WDSBootXML = [xml] $Node.unattendXML()
            $WDSBootXML.unattend.settings[1].component[1].WindowsDeploymentServices.ImageSelection.InstallImage.ImageName="Windows Server 2012 R2 Datacenter"
            $WDSBootXML.unattend.settings[1].component[1].WindowsDeploymentServices.ImageSelection.InstallImage.ImageGroup="Deploy"
            $WDSBootXML.unattend.settings[1].component[1].WindowsDeploymentServices.ImageSelection.InstallImage.Filename="template_windows2012R2.vhdx"
            $WDSBootXML.Save($BootUnattend)
        
        
            #Unattend Install Setup bearbeiten und in c:\remoteinstall\wds_install.xml speichern
            $InstallUnattend="c:\remoteinstall\wds_install.xml"
            $WDSBootXML.Save($InstallUnattend)
        
            #Serveroptionen setzen
            #Bootprogramm setzen
            #Server authorisieren
            #allen Clients antworten
            #Standardimage auf x64 setzen
            #immer installieren, es sei denn, es wird durch ESC abgebrochen
            WDSUTIL.exe /Set-Server /BootProgram:boot\x64\wdsnbp.com /Architecture:x64
            wdsutil.exe /set-server /authorize:yes
            wdsutil.exe /set-server /answerclients:All
            wdsutil.exe /set-server /DefaultX86X64ImageType:x64
            wdsutil.exe /set-server /PxePromptPolicy /new:optout
            wdsutil.exe /set-server /PxePromptPolicy /known:optout

            #Boot- und Install-Image einrichten und Installation unattended machen
            #wdsutil.exe /progress /verbose /add-image /ImageFile:$BootWIM /ImageType:Boot /skipverify
            #wdsutil.exe /progress /verbose /add-image /ImageFile:$InstallWIM /ImageType:Install /ImageGroup:$ImageGroup /skipverify /unattendfile:$InstallUnattend /singleimage:$Edition
            #$BootImage=($BootWIM.Split("\")[($BootWIM.Split("\")).length-1])
            #wdsutil.exe /set-server /BootImage:$BootImage /Architecture:x64
            #$BootUnattendShort=($BootUnattend.Split("\")[($BootUnattend.Split("\")).length-1])
            #WDSUTIL.exe /Set-Server /WdsUnattend /Policy:Enabled /File:$BootUnattendShort /Architecture:x64

            #Multicast-Transmission einrichten, damit Deploy auf mehrere Rechner schneller geht
            #wdsutil.exe /new-multicasttransmission /friendlyname:"WDS Autocast Default" /image:$Edition /ImageType:Install /transmissiontype:autocast
            }'
            restart-computer -force
            })

    $hyperv = {
        Install-WindowsFeature Hyper-V,Hyper-V-PowerShell -IncludeManagementTools -Verbose
        Restart-Computer -Force
    }
    $member = {
        $SecureString = (ConvertTo-SecureString -AsPlainText -String $Node.Password -Force)
        $User = $Node.SMBDomain+"\"+$Node.User
        $mycreds = New-Object System.Management.Automation.PSCredential ($User, $SecureString)
        $MGMTNetwork = $Node.Networks | Where-Object {$_.name -like "Mgmt*"}
        $Result = $false
        while (!$Result) {
            try {
                $Register = Register-DnsClient
                $Resolve = Resolve-DnsName -Name $MGMTNetwork.IPAddress -ErrorAction Stop
                if ($Resolve.NameHost -eq $Node.Name) {
                    $Result = $true
                }
            }
            catch {
                Start-Sleep -Seconds 5
                $Result = $false
            }
        }


        while ($true) {
            if ($DomainJoin -eq 1) {
                break
            }
            try {
                $DomainJoin=1
                Add-Computer -DomainName $Node.Domain -Restart:$false -Credential $mycreds -ErrorAction Stop
                }
            catch {
                Write-Output $Error[0]
                $DomainJoin=0
                Start-Sleep -Seconds 10
            }

        }
        Start-Sleep -Seconds 2
        Restart-Computer -Force
    }

    $ad = {

        Install-WindowsFeature AD-Domain-Services, RSAT-DHCP -IncludeManagementTools

    }


    $pdc = @( {
            
            $SecureString = (ConvertTo-SecureString -AsPlainText -String $Node.Password -Force)
            $parms = @{
                'DomainName'= $Node.Domain;
                'SafeModeAdministratorPassword'=$SecureString;
                'DomainNetbiosName'=$Node.SMBDomain;
                'InstallDns'= $true;
                'NoRebootOnCompletion'= $true;
                'Confirm' = $false;
            }
            
            Install-ADDSForest @parms
            
            $MGMTNetwork = $Node.Networks | Where-Object {$_.name -like "Mgmt*"}
            $MGMT= (Get-NetAdapter -Name $MGMTNetwork.Name)
            Set-DnsClientServerAddress -InterfaceIndex $MGMT.InterfaceIndex -ServerAddresses $MGMTNetwork.DNSServer
            Restart-Computer -Force
    },
      { 
            while (!(get-service DNS).status -eq "running"){
                Write-Warning "waiting for DNSServer"
            }
            Import-Module DNSServer
            $MGMTNetwork = $Node.Networks | Where-Object {$_.name -like "Mgmt*"}
            $NetworkID=($MGMTNetwork.NetworkID+"/"+$MGMTNetwork.Prefix)
            Add-DnsServerPrimaryZone -NetworkId $NetworkID -ReplicationScope Forest
            Start-Sleep -Seconds 5
            Register-DnsClient | Out-Null
            #User zum Aufnehmen anlegen
            #$SecureString = (ConvertTo-SecureString -AsPlainText -String $Node.Password -Force)
            $user=get-ADUser -Identity $Node.User
            #Enable-ADAccount -Identity $user
            $group= Get-ADGroup -Identity "Domain Admins"
            Add-ADGroupMember -Identity $group -Members $user
            Install-WindowsFeature RSAT-DHCP
            $WDS = $MyConfiguration.nodes | Where-Object {$_.roles -match "wds"}
            $WDS_FQDN = $WDS.Name + "." +$WDS.domain
            $WDS_MgmtNetwork = $Node.networks | Where-Object {$_.Name -like "Mgmt*"}
            $WDS_IP = $WDS_MgmtNetwork.IPAddress
            Add-DhcpServerInDC -DnsName $WDS_FQDN -IPAddress $WDS_IP
            Get-DhcpServerInDC
      }
   )

    $dhcp = @( 
        {
            $SecureString = (ConvertTo-SecureString -AsPlainText -String $Node.Password -Force)
            $User = $Node.SMBDomain+"\Administrator"
            $mycreds = New-Object System.Management.Automation.PSCredential ($User, $SecureString)
            Install-WindowsFeature DHCP -IncludeManagementTools
            

            while ((get-service dhcpserver).status -ne "Running") {
                Write-Warning -Message "Waiting for DHCP-Server"
                start-Service DHCPServer
                Start-Sleep -Seconds 5
            }
        $MGMTNetwork = $Node.Networks | Where-Object {$_.name -like "Mgmt*"}

        #Scrope anlegen
        #Start- und EndIP werden anhand des WDS-Netzes bestimmt
        $Mask = (Get-NetworkSummary ($MGMTNetwork.networkID+"/"+$MgmtNetwork.Prefix)).Mask
        $IPRange = Get-NetworkRange ($MGMTNetwork.networkID+"/"+$MgmtNetwork.Prefix)
        $StartIP = $IPRange[0]
        $EndIP = $IPRange[-1]

        $Scope = Add-DhcpServerv4Scope -StartRange $StartIP -EndRange $EndIP -SubnetMask $Mask -Name $MGMTNetwork.name -PassThru
        

        #Scope ausschließen
        Add-DhcpServerv4ExclusionRange -ScopeId $Scope.ScopeId -StartRange $StartIP -EndRange $EndIP 
 

        #DHCP Brereichsoptionen anlegen
        Set-DhcpServerv4OptionValue -ScopeId $Scope.ScopeId -OptionId 006 -Value $MGMTNetwork.DNSServer.split(",") -Force
        Set-DhcpServerv4OptionValue -ScopeId $Scope.ScopeId -OptionId 015 -Value $MGMTNetwork.DNSSuffix
        Set-DhcpServerv4OptionValue -ScopeId $Scope.ScopeId -OptionId 066 -Value $MGMTNetwork.IPAddress
        Set-DhcpServerv4OptionValue -ScopeId $Scope.ScopeId -OptionId 067 -Value "boot\x64\wdsnbp.com"

        
        #Reservierungen einrichten
        $Nodes = $MyConfiguration.nodes | Where-Object {$_.BootMac -ne ""};
        foreach($Node in $Nodes) {
            $MGMTNetwork = $Node.Networks | Where-Object {$_.name -like "Mgmt*"};
            Add-DhcpServerv4Reservation -ScopeId $Scope.ScopeId -IPAddress $MGMTNetwork.IPAddress -ClientId $Node.BootMac -Name $Node.Name
        
        }


        }
    )
    $bdc = {
        Write-Host "test"
    }

    
    $iscsi={
        Write-Host "test"
    }

    $cluster= {

        Write-Host "test"
    }

    $finish = {

    }

   <#
    $RunOnceRestart = {
        
        #Dafür sorgen, dass das nächste Startscript ausgeführt wird
        Write-Host “Changing RunOnce script.” -foregroundcolor “magenta”
        $RunOnceKey = “HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce”
        $KeyContent1 = ('C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe -File ')
        $KeyContent2 = ("C:\deploy\scripts\<NEXT>.ps1")
        $KeyContent = $KeyContent1 + $KeyContent2
        set-itemproperty $RunOnceKey “NextRun” $KeyContent
        Set-AutoLogon -DefaultUsername "Administrator" -DefaultPassword $Node.Password
        #Aufräumen und Restart
        Remove-Item ("C:\deploy\scripts\<CURRENT>.ps1") -Force
        Restart-Computer -Force
    }

    $RemoveRunOnceRestart = {
        
        #Dafür sorgen, dass der RunOnceKey gelöscht wird
        Write-Host “Changing RunOnce script.” -foregroundcolor “magenta”
        #$RunOnceKey = “HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce”
        #Remove-ItemProperty $RunOnceKey “NextRun”
        #Aufräumen und Restart
        Remove-Item ("C:\deploy\scripts\<CURRENT>.ps1") -Force
        Restart-Computer -Force
    }

   #> 

    function get-StartConfig {
        #bekommt ein Array von Strings und gibt den Replace zurück
        param(
            [Parameter(Mandatory=$True,Position=1)][STRING[]]$Config

        ) 

        $Config | ForEach-Object {

        $ConfigItem = $_.toString()

        

        #Wenn der String nicht der letzte ist, dann append $RunOnceRestart mit dem nächsten String im Array
        #Hier werden Scriptblöcke miteinander verglichen
        if ( $ConfigItem -ne ($Config[($Config.Length -1)]).ToString()) {
    
            #herausfinden an welcher Reihe im Array der Scriptblock steht und den namen des nächsten ermitteln
            $RunOnceRestarttemp = $RunOnceRestart.ToString()
            $RunOnceRestarttemp = $RunOnceRestarttemp.Replace('<CURRENT>',("Start"+($Config.IndexOf($ConfigItem)).toString()))
            $RunOnceRestarttemp = $RunOnceRestarttemp.Replace('<NEXT>',("Start"+($Config.IndexOf($ConfigItem)+1)).toString())
            #Replace falls restart-computer, damit das selbe script nochmal ausgeführt wird
            $RunOnceRestarttemp1 = $RunOnceRestart.ToString()
            $RunOnceRestarttemp1 = $RunOnceRestarttemp1.Replace('<CURRENT>',("Start"+($Config.IndexOf($ConfigItem)-1).toString()))
            $RunOnceRestarttemp1 = $RunOnceRestarttemp1.Replace('<NEXT>',("Start"+($Config.IndexOf($ConfigItem))).toString())
            $ConfigItem = $ConfigItem.Replace("Restart-Computer -Force",$RunOnceRestarttemp1)
            #ende replace
            $ConfigItem = [STRING]( $ConfigItem,$RunOnceRestarttemp -join " ")
            #$ConfigItem= ('. c:\deploy\scripts\create-lab.ps1 ´n'+$ConfigItem)
        }
        else {
            $RunOnceRestarttemp = $RemoveRunOnceRestart.ToString()
            $RunOnceRestarttemp = $RunOnceRestarttemp.Replace('<CURRENT>',("Start"+($Config.IndexOf($ConfigItem)).toString()))
            $ConfigItem = [STRING]( $ConfigItem,$RunOnceRestarttemp -join " ")
        }


        return $ConfigItem

        }

    }
 
    #. C:\deploy\scripts\create-lab.ps1
    $Roles = ($Node.Roles -replace '([A-z]\w+)','$ $1').Replace("$ ","$")
        
    [int]$i=0

    $Node.Roles.Split(",") | ForEach-Object {
    
         $i++
        
        }


    $ConfigItems=@()
    $ConfigItems +=$Start1
    #$ConfigItems[$ConfigItems.Length-1]=$Ende
   
    $Node.Roles.Split(",") | ForEach-Object {
        $ConfigItems += (Get-Variable -Name $_).Value
    }
    $ConfigItems += $finish
    $StartConfig = $ConfigItems
    Write-Output $StartConfig
   <#
    $StartConfig | ForEach-Object {
        $FilePath = ("c:\deploy\scripts\Start"+($StartConfig.indexof($_))+".ps1")          
        $_ | Out-File -FilePath $FilePath
    }

    Invoke-Expression c:\deploy\scripts\Start0.ps1
   #>

   }
    [STRING]$Name
#region Constructors
    Node([STRING]$Name) {
          
          if((Get-Item -Path $Name -ErrorAction SilentlyContinue).exists){
            $item = Get-Item -Path $Name
            [PSObject]$obj = $Null
            if($item.extension -match 'xml'){
            #xml
                $obj = Import-Clixml -Path ($item.fullname)
            }
            elseif($item.extension -match 'csv'){
            #csv
                $obj= Import-Csv -Path ($item.fullname)

            }
            else{
                Write-Error "Fileformat not supported"
                break
            }

            $this.init($obj)
          }
          else {
            #if not a file
            $this.Name = $Name
          }
    }

    Node([Hashtable]$Hash){
        $this.init($Hash)

    }

    Node([PSObject]$obj){

        $this.init($obj)        

    }

    Node([STRING]$Name,[Hashtable]$Hash){
        if($Hash.Contains($Name)){
            $this.init($Hash.$Name)

        }
        else{
            $Hash.keys | ForEach-Object{
                if (($Hash.$_).contains($Name)) {
                    $this.init(($Hash.$_).$Name)
                }
            }
        }


        #to allow unencrypted password for dsc-configuration (THIS IS NOT SECURE!!! But we only create a LAB ;) )
        $this | Add-Member -MemberType NoteProperty -Name "PSDscAllowPlainTextPassword" -Value $true
        $this | Add-Member -MemberType NoteProperty -Name "PSDscAllowDomainUser" -Value $true


        if($Hash.contains("Networks")){
            $this | Add-Member -MemberType NoteProperty -Name "Networks" -value @() -force
            $this | Get-Member -MemberType NoteProperty | ForEach-Object {
                $NoteProperty = $_.Name
                if ($Hash.Networks.contains($NoteProperty)) {
                    $Contains = $false
                    $this.Networks | ForEach-Object{
                        $Network = $_
                        if ($Network.Name -match $NoteProperty) {
                            $Contains = $true
                        }
                    }
                    if (!$Contains) {
                        $NewNetwork = [Network]::new($Hash.Networks.$NoteProperty)
                        $NewNetwork.IPAddress = $this.$NoteProperty
                        $this.networks += $NewNetwork
                    }
                }
            }
        }

        if ($Hash.contains("ConfigurationStandards")) {
            $this | Add-Member -MemberType NoteProperty -Name "ConfigurationStandards" -value $Hash.ConfigurationStandards -force
            $Paths = ($this.ConfigurationStandards.keys).where({$_ -match "Path"})
            $Paths | ForEach-Object{
                if ( !(Test-Path -Path $this.ConfigurationStandards.$_.value) ) {
                    mkdir -Path $this.ConfigurationStandards.$_.value -Force
                } 
            }
        }
    }


#region methods
    init([Hashtable]$Hash){
        
        $Hash.Keys | ForEach-Object {
            $this | Add-Member -MemberType NoteProperty -Name $_ -Value $Hash.$_ -force
        }
        if ($Hash.NodeName[3] -eq 'V'){
            $this | Add-Member -MemberType NoteProperty -Name isVirtual -Value $true -force 
        }else {
            $this | Add-Member -MemberType NoteProperty -Name isVirtual -Value $false -force
        }

        if (($Hash.NodeName[4,5] -join("")) -eq "IN" ) {
            $this | Add-Member -MemberType NoteProperty -Name Layer -Value "INFRA" -force
        }

        if (($Hash.NodeName[4,5] -join("")) -eq "PF" ) {
            $this | Add-Member -MemberType NoteProperty -Name Layer -Value "HSA" -force
        }

        if (($Hash.NodeName[4,5] -join("")) -eq "GW" ) {
            $this | Add-Member -MemberType NoteProperty -Name Layer -Value "SSA" -force
        }

        if (($Hash.NodeName[4,5] -join("")) -eq "CM" ) {
            $this | Add-Member -MemberType NoteProperty -Name Layer -Value "CentralManagment" -force
        }

        $this.Name = $Hash.Nodename

     }

    
    init([PSObject]$obj){
        $obj | get-member -MemberType NoteProperty | ForEach-Object {
            $this | Add-Member -MemberType NoteProperty -Name $_.Name -Value $obj.($_.Name) -Force

        }
        $this.Name = $this.nodename

     }

    [hashtable] getConfigurationData(){
         $ConfigurationData = @{
            AllNodes =
            @(
                @{
                    
                }

            )
         
         }

         $this | Get-Member -MemberType NoteProperty,Property | ForEach-Object {
             $NoteProperty = $_.Name
             if ($this.$NoteProperty -is [array]) {
                 $Array = @()
                 
                 $this.$NoteProperty | ForEach-Object {
                     $ArrayObject = $_
                     if ($ArrayObject -is [hashtable]) {
                        $Array += $ArrayObject
                     }
                     elseif ($ArrayObject -is [Network]) {
                         $Array += ($ArrayObject.getConfigurationData())
                     }
                 }
                 $ConfigurationData.allnodes[0].add($NoteProperty,$Array)
             }
             else {
                 $ConfigurationData.AllNodes[0].add($NoteProperty,$this.$NoteProperty)
             }
            
         }

         return $ConfigurationData

         
     }

    exportXML([string]$path){
        $this | Export-Clixml -Path $path
    
     }
    [psobject]getMasterImage(){
    
        return (get-vhd -Path $this.TemplatePath)
       
     }

    [psobject]createMasterImage(){
        
        #create MasterImage from ISO and integration of Packages
        $ISO = $null
        $VHDPATH = $null
        if ($this.OS -eq "Windows Server 2012 R2") {
            $ISO = join-path -Path $this.ConfigurationStandards.ISOPath.Value -ChildPath $this.ConfigurationStandards.ISO_Windows2012R2.value
            $VHDPATH = join-path -path $this.ConfigurationStandards.HyperVVHDPath.Value -childpath $this.ConfigurationStandards.template_Windows2012R2.Value
        }
        elseif ($this.OS -eq "Windows Server 2012") {
            $ISO= join-path -path $this.ConfigurationStandards.ISOPath.Value -childpath $this.ConfigurationStandards.ISO_Windows2012.value
            $VHDPATH = join-path -path $this.ConfigurationStandards.HyperVVHDPath.Value -childpath $this.ConfigurationStandards.template_Windows2012.Value
        }
        else {
            Write-Error ($this.OS+" is not a valid OS")
        }
        
        $ISOItem = Test-Path -Path $ISO
        $VHDPATHItem = Test-Path -Path $VHDPATH
        if (!$ISOItem) {
            Write-Error "$ISO not found"
        }

        if ($VHDPATHItem) {
            Write-Information "$VHDPATH exists"
            $this | Add-Member -MemberType NoteProperty -Name TemplatePath -Value $VHDPATH
        }

        if (($ISOItem) -and (!$VHDPATHItem) ) {
            $ISOImageMount = (Mount-DiskImage -ImagePath $ISO -StorageType ISO -PassThru | Get-Volume).DriveLetter
            $WIM = ($ISOImageMount+":\sources\install.wim")
            $BootWIM = ($ISOImagemount + ":\sources\boot.wim")
            $PackagePath = $this.ConfigurationStandards.PackagePath.Value
            #$Packages = Get-ChildItem $PackagePath
            Convert-WindowsImage -SourcePath $WIM -VHDPath $VHDPATH -Edition ($this.OS+" SERVERDATACENTER") -VHDFormat VHDX -Verbose -VHDType Dynamic -SizeBytes 50GB -VHDPartitionStyle MBR
            Copy-Item -Path $BootWIM -Destination ($this.ConfigurationStandards.HyperVVHDPath.Value+"\") -Verbose
            dismount-DiskImage $ISO -Verbose
            $this | Add-Member -MemberType NoteProperty -Name TemplatePath -Value $VHDPATH
            #Integration of Packages
            try {
                $this.getMasterImage() | Mount-VHD
                Start-Sleep -Seconds 2
                $ScripDest= (($this.getMasterImage() | Get-Partition | Where-Object { $_.Size -gt 350MB}).DriveLetter+":\")
                $ScriptPath = $this.ConfigurationStandards.ScriptPath.Value
                $PackagePath = $this.ConfigurationStandards.PackagePath.Value
                Get-ChildItem "$PackagePath" | ForEach-Object {
                    Add-WindowsPackage -Path $ScripDest -PackagePath $_.FullName -Verbose    
                }
                $ScriptPathDest = mkdir -path (Join-Path -Path $ScripDest -ChildPath ($this.ConfigurationStandards.ScriptPath.Value.split(":")[1])) -Force
                Copy-Item -Path "$ScriptPath\*" -Destination $ScriptPathDest -Recurse -Force
                Copy-Item -Path $this.configsource -Destination $ScriptPathDest -Force
                Copy-Item -Path $this.ConfigurationStandards.ModulePath.Value -Destination (Join-Path -Path $ScripDest -ChildPath "Program Files\WindowsPowerShell\") -Recurse -Force
            }
            catch {
                Write-Host $Error[0]
            }
            finally {
                Dismount-VHD -Path $this.templatePath
            }
        }
        $this.save()
        return $this.getMasterImage()
     }

    [psobject]createVHD(){
        #copy the MasterImage with new name to new filelocation
        $VHDPath = Join-Path -path $this.ConfigurationStandards.HyperVVHDPath.Value -ChildPath ($this.Name,"vhdx" -join(".")) 
        $TempplatePath = $this.templatepath
        if (!$TempplatePath) {
            Write-Error "$TempplatePath not found"
        }
        $VHDPathItem = Test-Path $VHDPath
        if (!$VHDPathItem) {
            Copy-Item -Path $TempplatePath -Destination $VHDPath -Verbose -Force
            
        }
        $this | Add-Member -MemberType NoteProperty -Name VHDPath -Value $VHDPath -Force
        $this.save()
        #mounting VHD and copy some stuff
        $MyDisk=( $this.getVHD() | Mount-VHD -Passthru)
        $ScripDest= ((Get-Disk $MyDisk.DiskNumber | Get-Partition | Where-Object { $_.Size -gt 350MB}).DriveLetter+":")
        #DriverIntegration, and copy of scriptsources for non virtual Nodes
        if (!$this.isvirtual) {
            $DriverPath = $this.ConfigurationStandards.DriverPath.Value
            Add-WindowsDriver -Recurse -ForceUnsigned -Driver $DriverPath -Path (Join-Path -path $ScripDest -childpath "\")
            $SoftwarePath = $this.ConfigurationStandards.SoftwarePath.value
            $SoftwarePathDestination = Join-Path -Path $ScripDest -ChildPath ($SoftwarePath.split(":"))[1]
            mkdir -Force -Path $SoftwarePathDestination
            Copy-Item -Path $SoftwarePath  -Destination $SoftwarePathDestination -Force -Verbose -Recurse
        }
        #create baremetal-template and move into wds-vhd
        if ($this.Roles -match "wds") {
            $MyConfiguration = [MyConfiguration]::new($this.configsource)
            $BareMetalImage = Join-Path -path $this.ConfigurationStandards.HyperVVHDPath.Value -ChildPath "BareMetalImage.vhdx" 
            move-item -Path $BareMetalImage -Destination "$ScripDest\deploy\template_windows2012R2.vhdx" -Verbose -Force
            copy-item -path ($this.ConfigurationStandards.HyperVVHDPath.Value+"\boot.wim") -Destination "$ScripDest\deploy\" -Verbose
        }

        $MyDisk | Dismount-VHD 
        return $this.getVHD()
     }

    [psobject]getVHD(){
        return (get-vhd -path $this.VHDPath)        
     }

    [psobject]createVM(){
        $VHDPath = $this.getVHD().Path
        $MgmtNetwork = $this.networks | Where-Object {$_.name -like "Mgmt*"}
        $VM = $null
        if ($this.isVirtual) {
            try {
                $VMId = (get-vm -Name $this.Name -ErrorAction Stop).VMId
                $this | Add-Member -MemberType NoteProperty -name VMId -Value $VMId -Force
            }
            catch {
                #$Error[0] | Out-Host
                $VM = New-VM -Name $this.Name -MemoryStartupBytes ([int64]$this.memory * 1024 * 1024) -Generation 1 -Path $this.ConfigurationStandards.hypervvmpath.value
                $VM | Start-VM -Passthru | Stop-VM -Force
                $VM | Add-VMHardDiskDrive -Path $VHDPath
                $VM | Set-VM -ProcessorCount $this.Cores -notes $this.Description
                $this | Add-Member -MemberType NoteProperty -name VMId -Value $VM.VMId -Force
                $VMNetworkAdapter = $VM | Get-VMNetworkAdapter | Rename-VMNetworkAdapter -NewName $MgmtNetwork.Name -Passthru
                $VMNetworkAdapter | Connect-VMNetworkAdapter -SwitchName $MgmtNetwork.team
                $MAC = Convert-MacAddress $VMNetworkAdapter.MacAddress -Delimiter "-"
                $this | Add-Member -MemberType NoteProperty -Name BOOTMAC -Value $MAC -Force
                #injecting MAC-Address in unattended installation
                $unattend = $this.unattendXML()
                #mounting VHD, placing unattend.xml and copy some stuff
                $MyDisk=( $this.getVHD() | Mount-VHD -Passthru)
                $ScripDest= ((Get-Disk $MyDisk.DiskNumber | Get-Partition | Where-Object { $_.Size -gt 350MB}).DriveLetter+":")
                $SysprepPath = Join-Path -Path $ScripDest -ChildPath "\Windows\System32\Sysprep\unattend.xml"
                $unattend.Save($SysprepPath)
                $ScriptPathDest = mkdir -path (Join-Path -Path $ScripDest -ChildPath ($this.ConfigurationStandards.ScriptPath.Value.split(":")[1])) -Force
                $ScriptPath = $this.ConfigurationStandards.ScriptPath.Value
                Copy-Item -Path "$ScriptPath\*" -Destination $ScriptPathDest -Recurse -Force
                $this.save()
                Copy-Item -Path $this.configsource -Destination $ScriptPathDest -Force
                Copy-Item -Path $this.configsource -Destination $ScripDest -Force
                
                $MyDisk | Dismount-VHD
            }
        }
        $this.save()
        return $this.getVM()
     }

    [psobject]getVM(){
        return (get-vm -Id $this.VMId)
     }
     
    [xml] unattendXML(){
        [xml] $unattend= [xml]"<?xml version='1.0' encoding='utf-8'?>
        <unattend xmlns='urn:schemas-microsoft-com:unattend'>
    <settings pass='specialize'>
        <component name='Microsoft-Windows-Shell-Setup' processorArchitecture='amd64' publicKeyToken='31bf3856ad364e35' language='neutral' versionScope='nonSxS' xmlns:wcm='http://schemas.microsoft.com/WMIConfig/2002/State' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>
            <OEMInformation>
                <Manufacturer></Manufacturer>
                <SupportURL></SupportURL>
            </OEMInformation>
            <ComputerName>*</ComputerName>
            <ProductKey></ProductKey>
            <TimeZone></TimeZone>
        </component>
 
        <component name='Microsoft-Windows-TCPIP' processorArchitecture='amd64' publicKeyToken='31bf3856ad364e35' language='neutral' versionScope='nonSxS' xmlns:wcm='http://schemas.microsoft.com/WMIConfig/2002/State' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>
            <Interfaces>
                <Interface wcm:action='add'>
                    <Identifier></Identifier>
                    <Ipv4Settings>
                        <DhcpEnabled>false</DhcpEnabled>
                        <Metric>10</Metric>
                        <RouterDiscoveryEnabled>false</RouterDiscoveryEnabled>
                    </Ipv4Settings>
                    <UnicastIpAddresses>
                    <IpAddress wcm:action='add' wcm:keyValue='1'>192.168.5.201/24</IpAddress>
                    </UnicastIpAddresses>
                </Interface>
            </Interfaces>
        </component>

    </settings>
    <settings pass='windowsPE'>
        <component name='Microsoft-Windows-International-Core-WinPE' processorArchitecture='amd64' publicKeyToken='31bf3856ad364e35' language='neutral' versionScope='nonSxS' xmlns:wcm='http://schemas.microsoft.com/WMIConfig/2002/State' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>
            <SetupUILanguage>
                <UILanguage></UILanguage>
            </SetupUILanguage>
            <InputLocale></InputLocale>
            <SystemLocale></SystemLocale>
            <UILanguage></UILanguage>
            <UserLocale></UserLocale>
        </component>
        <component name='Microsoft-Windows-Setup' processorArchitecture='amd64' publicKeyToken='31bf3856ad364e35' language='neutral' versionScope='nonSxS' xmlns:wcm='http://schemas.microsoft.com/WMIConfig/2002/State' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>
      <DiskConfiguration>
        <Disk wcm:action='add'>
          <CreatePartitions>
            <CreatePartition wcm:action='add'>
              <Type>Primary</Type>
              <Size>350</Size>
              <Order>1</Order>
            </CreatePartition>
            <CreatePartition wcm:action='add'>
              <Type>Primary</Type>
              <Size>51000</Size>
              <Order>2</Order>
            </CreatePartition>
            <CreatePartition wcm:action='add'>
              <Order>3</Order>
              <Type>Primary</Type>
              <Extend>true</Extend>
            </CreatePartition>
          </CreatePartitions>
          <ModifyPartitions>
            <ModifyPartition wcm:action='add'>
              <Active>true</Active>
              <Format>NTFS</Format>
              <Label>Boot</Label>
              <Order>1</Order>
              <PartitionID>1</PartitionID>
            </ModifyPartition>
            <ModifyPartition wcm:action='add'>
              <Format>NTFS</Format>
              <Label>Temp</Label>
              <Order>2</Order>
              <PartitionID>2</PartitionID>
            </ModifyPartition>
            <ModifyPartition wcm:action='add'>
              <Format>NTFS</Format>
              <Label>System</Label>
              <Order>3</Order>
              <PartitionID>3</PartitionID>
            </ModifyPartition>
          </ModifyPartitions>
          <WillWipeDisk>true</WillWipeDisk>
          <DiskID>0</DiskID>
        </Disk>
        <WillShowUI>OnError</WillShowUI>
      </DiskConfiguration>
            <WindowsDeploymentServices>
                <Login>
                    <WillShowUI>OnError</WillShowUI>
                    <Credentials>
                        <Username></Username>
                        <Domain></Domain>
                        <Password></Password>
                    </Credentials>
                </Login>
                <ImageSelection>
                    <InstallImage>
                        <ImageName></ImageName>
                        <ImageGroup></ImageGroup>
                        <Filename></Filename>
                    </InstallImage>
                    <WillShowUI>OnError</WillShowUI>
                    <InstallTo>
                        <DiskID>0</DiskID>
                        <PartitionID>3</PartitionID>
                    </InstallTo>
                </ImageSelection>
            </WindowsDeploymentServices>
            <UserData>
                <ProductKey>
                    <Key></Key>
                    <WillShowUI>Never</WillShowUI>
                </ProductKey>
                <AcceptEula>true</AcceptEula>
            </UserData>
        </component>
    </settings>
          <settings pass='oobeSystem'>
            <component name='Microsoft-Windows-International-Core' processorArchitecture='amd64' publicKeyToken='31bf3856ad364e35' language='neutral' versionScope='nonSxS' xmlns:wcm='http://schemas.microsoft.com/WMIConfig/2002/State' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>
              <InputLocale></InputLocale>
              <SystemLocale></SystemLocale>
              <UILanguage></UILanguage>
              <UserLocale></UserLocale>
            </component>
            <component name='Microsoft-Windows-Shell-Setup' processorArchitecture='amd64' publicKeyToken='31bf3856ad364e35' language='neutral' versionScope='nonSxS' xmlns:wcm='http://schemas.microsoft.com/WMIConfig/2002/State' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>
              <OOBE>
                <HideEULAPage>true</HideEULAPage>
              </OOBE>
              <AutoLogon>
                <Password>
                  <Value></Value>
                  <PlainText>true</PlainText>
                </Password>
                <Enabled>true</Enabled>
                <LogonCount>1</LogonCount>
                <Username>Administrator</Username>
              </AutoLogon>
            <FirstLogonCommands>
                <SynchronousCommand wcm:action='add'>
                  <CommandLine>%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy RemoteSigned -File c:\deploy\scripts\start-config.ps1</CommandLine>
                  <Description>Configuration in Progress</Description>
                  <Order>2</Order>
                  <RequiresUserInput>false</RequiresUserInput>
                </SynchronousCommand>
                <SynchronousCommand wcm:action='add'>
                  <CommandLine>cmd /c reg add 'HKLM\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell' /v ExecutionPolicy /t REG_SZ /d RemoteSigned /f</CommandLine>
                  <Description>setting Executionpolicy</Description>
                  <Order>1</Order>
                </SynchronousCommand>
              </FirstLogonCommands>
              <UserAccounts>
                <AdministratorPassword>
                  <Value></Value>
                  <PlainText>true</PlainText>
                </AdministratorPassword>
                <LocalAccounts>
                    <LocalAccount wcm:action='add'>
                    <Password>
                        <Value></Value>
                        <PlainText>true</PlainText>
                    </Password>
                    <Description>Local Administrator</Description>
                    <DisplayName></DisplayName>
                    <Group>Administrators</Group>
                    <Name></Name>
                </LocalAccount>
               </LocalAccounts>
              </UserAccounts>
              <RegisteredOwner>-</RegisteredOwner>
              <RegisteredOrganization>-</RegisteredOrganization>
            </component>
          </settings>
          <cpi:offlineImage cpi:source='wim:d:/clusterinstall/install.wim#Windows Server 2012 SERVERDATACENTERCORE' xmlns:cpi='urn:schemas-microsoft-com:cpi' />
        </unattend>"  
        $ProductKey = ""
        if ($this.os -eq "Windows Server 2012 R2") {
            $ProductKey = $this.ConfigurationStandards.ProductKey_Windows2012R2.value
        }
        elseif ($this.os -eq "Windows Server 2012") {
            $ProductKey = $this.ConfigurationStandards.ProductKey_Windows2012.value
        }
        $MgmtNetwork = $this.networks | Where-Object {$_.Name -like "Mgmt*"}
        $IPAddress = $MgmtNetwork.IPAddress+"/"+$MgmtNetwork.Prefix

        #Microsoft-Windows-Shell-Setup
        $unattend.unattend.settings.component[0].ComputerName = $this.Name
        $unattend.unattend.settings.component[0].ProductKey = $ProductKey
        $unattend.unattend.settings.component[0].TimeZone = $this.configurationstandards.TimeZone.Value
        
        #Microsoft-Windows-TCPIP
        $unattend.unattend.settings.component[1].Interfaces.Interface.Identifier = $this.BOOTMAC
        $unattend.unattend.settings.component[1].Interfaces.Interface.UnicastIpAddresses.IpAddress.'#text' = $IPAddress

        #Microsoft-Windows-International-Core-WinPE
        $unattend.unattend.settings.component[2].SetupUILanguage.UILanguage = $this.configurationstandards.UILanguage.Value
        $unattend.unattend.settings.component[2].InputLocale = $this.configurationstandards.InputLanguage.Value
        $unattend.unattend.settings.component[2].SystemLocale = $this.configurationstandards.InputLanguage.Value
        $unattend.unattend.settings.component[2].UILanguage = $this.configurationstandards.UILanguage.Value
        $unattend.unattend.settings.component[2].UserLocale = $this.configurationstandards.InputLanguage.Value

        #Microsoft-Windows-Setup
        $unattend.unattend.settings.component[3].WindowsDeploymentServices.Login.Credentials.Username = $this.User
        $unattend.unattend.settings.component[3].WindowsDeploymentServices.Login.Credentials.Domain = $this.Domain
        $unattend.unattend.settings.component[3].WindowsDeploymentServices.Login.Credentials.Password = $this.Password
        $unattend.unattend.settings.component[3].WindowsDeploymentServices.ImageSelection.InstallImage.ImageName = "Install"
        $unattend.unattend.settings.component[3].WindowsDeploymentServices.ImageSelection.InstallImage.ImageGroup = "Deploy"
        $unattend.unattend.settings.component[3].WindowsDeploymentServices.ImageSelection.InstallImage.Filename = "template_windows2012R2.vhdx"
        $unattend.unattend.settings.component[3].WindowsDeploymentServices.ImageSelection.InstallTo.DiskID = "0"
        $unattend.unattend.settings.component[3].WindowsDeploymentServices.ImageSelection.InstallTo.PartitionID = "3"
        $unattend.unattend.settings.component[3].userdata.ProductKey.key = $ProductKey


        #oobeSystem
        #Microsoft-Windows-International-Core
        $unattend.unattend.settings.component[4].InputLocale = $this.configurationstandards.InputLanguage.Value
        $unattend.unattend.settings.component[4].SystemLocale = $this.configurationstandards.InputLanguage.Value
        $unattend.unattend.settings.component[4].UILanguage = $this.configurationstandards.UILanguage.Value
        $unattend.unattend.settings.component[4].UserLocale = $this.configurationstandards.InputLanguage.Value
    
        #Microsoft-Windows-Shell-Setup
        $unattend.unattend.settings.component[5].AutoLogon.Password.Value = $this.Password
        $unattend.unattend.settings.component[5].AutoLogon.Username = $this.User
        $unattend.unattend.settings.component[5].AutoLogon.LogonCount = $this.configurationstandards.AutoLoginCount.Value
        $unattend.unattend.settings.component[5].UserAccounts.AdministratorPassword.Value = $this.Password
        $unattend.unattend.settings.component[5].UserAccounts.LocalAccounts.LocalAccount.Name = $this.User
        $unattend.unattend.settings.component[5].UserAccounts.LocalAccounts.LocalAccount.DisplayName = $this.User
        $unattend.unattend.settings.component[5].UserAccounts.LocalAccounts.LocalAccount.Password.Value = $this.Password
        $unattend.unattend.settings.component[5].RegisteredOrganization = $this.configurationstandards.RegisteredOrganisation.Value
        $unattend.unattend.settings.component[5].RegisteredOwner = $this.configurationstandards.RegisteredOwner.Value
        return $unattend
     }
    
    [bool]isOnline(){
        $isOnline = $false
        $MgmtNetwork = ($this.networks).where({$_.Name -match "Mgmt"})
        if (Test-NetConnection -ComputerName $MgmtNetwork.IPAddress -CommonTCPPort WINRM -Verbose -InformationLevel Quiet) {
            if (Test-NetConnection -ComputerName $MgmtNetwork.IPAddress -CommonTCPPort WINRM -Verbose -InformationLevel Quiet) {
                $isOnline = $true
            }
        }
        return $isOnline
     }

    save(){
        $MyConfiguration = [MyConfiguration]::new($this.ConfigSource)
        $Node = $MyConfiguration.getNodesbyName($this.Name)
        $Node.init($this)
        $MyConfiguration.save()
        $MyConfiguration = $null
     }
    [psobject]PSSession(){
        $MgmtNetwork = $this.networks | Where-Object {$_.name -like "Mgmt*"}
        $Password = ConvertTo-SecureString -String $this.Password -AsPlainText -Force
        $Session = $false
        Enable-WSManCredSSP -Role Client -DelegateComputer $env:COMPUTERNAME -Force
        try {
            if ($this.isOnline()) {
                Write-Warning -Message "Trying Domain Account..."
                $MyCred = New-Object System.Management.Automation.PSCredential (($this.SMBDomain+"\Administrator"), $Password)
                $Session = New-PSSession -Name $this.Name -ComputerName $this.FQDN() -Credential $MyCred  -ErrorAction Stop
                Invoke-Command -Session $Session -ScriptBlock {. "C:\deploy\scripts\create-lab.ps1"} -ErrorAction Stop
                invoke-command -Session $Session -ScriptBlock {Enable-WSManCredSSP -Role Server -Force}
            }
        }
        catch {
            if ($this.isOnline()) {
                Write-Warning -Message "Trying local Account..."
                $MyCred = New-Object System.Management.Automation.PSCredential ("Administrator", $Password)
                $Session = New-PSSession -Name $this.Name -ComputerName $this.FQDN() -Credential $MyCred 
                Invoke-Command -Session $Session -ScriptBlock {. "C:\deploy\scripts\create-lab.ps1"} 
                invoke-command -Session $Session -ScriptBlock {Enable-WSManCredSSP -Role Server -Force}
            }

        }

        return $Session
     }
    wait(){
        Start-Sleep -Seconds 3
        while (!$this.isOnline()) {
            Start-Sleep -Seconds 5    
        }
     }

    [PSObject]run([ScriptBlock]$ScriptBlock){
         Write-Warning ("Waiting for "+ $this.Name)
         $this.wait()
         Write-Warning ($this.name+": Trying to get PSSession...")
         $Result = $false
         $Session = $this.PSSession()
         if ($Session) {
             Write-Warning ($this.Name+": Trying to Invoke-Command...")
             Write-Warning $ScriptBlock.ToString()
             $Result = Invoke-Command -Session $Session -ScriptBlock $ScriptBlock
         }else {
             $Result = $false
         }
         return $Result
     }

    applyConfig(){
        $Scriptblock = [Scriptblock]::Create( 
            {   
                param($NodeName)
                . C:\deploy\scripts\create-lab.ps1
                $Node = $MyConfiguration.getNodesbyName($Nodename)
                $StartConfig = [scriptblock]::Create([Node]::StartScript)
                $StartConfig = $Node.run($StartConfig)
                #Write-Warning ($NodeName + $StartConfig.ToString())
                foreach ( $Scriptblock in $StartConfig){
                    $Node.run([scriptblock]::Create($Scriptblock)) | Out-Default
                    $Node.wait()
                }
        })
        Start-Job -ScriptBlock $Scriptblock -Name ($this.name+".applyConfig()") -ArgumentList $this.Name
     }
    [String]FQDN() {
        return ($this.Name + "." + $this.Domain)
     }
    [System.Management.Automation.PSCredential]getLocalCredentials(){
        $User = "Administrator"
        $Password = ConvertTo-SecureString -String $this.Password -AsPlainText -Force
        return [System.Management.Automation.PSCredential]::new($User,$Password)
     }
    [System.Management.Automation.PSCredential]getDomainCredentials(){
        $User = $this.SMBDomain + "\Administrator"
        $Password = ConvertTo-SecureString -String $this.Password -AsPlainText -Force
        return [System.Management.Automation.PSCredential]::new($User,$Password)
     }


 }

class MyConfiguration {
    [Node[]] $Nodes
    [String] $ConfigSource
#region Constructors
    MyConfiguration(){

    }
    MyConfiguration([Hashtable]$Hash){
        $this.init($Hash)
        $this.ConfigSource = "from Hash"
        $this.setNodesConfigSource()
    }

    MyConfiguration([string]$path){
        $Item = Get-Item -Path $path
        if ($Item.Extension -match 'xls') {
            $this.init((get-ExcelDataHashTable($path)))
            $this.ConfigSource = $Item.FullName
        }
        elseif ($Item.Extension -match 'xml') {
            $xmlObjects = Import-Clixml -Path $Item.FullName
            $xmlObjects.Nodes | ForEach-Object {
                $this.add([Node]::new($_))
            }
            $this.ConfigSource = $Item.FullName
        }
        else {
            Write-Error -Message 'Format nicht unterstützt'
            break
        }
        $this.setNodesConfigSource()
    }

    init([Hashtable]$Hash){
        $Hash.keys | ForEach-Object {
            if ( ($_ -notmatch "Networks") -and ($_ -notmatch "ConfigurationStandards")){
                $Layer = $_
                ($Hash.$Layer).keys | ForEach-Object {
                    $this.add($_,$Hash)
                }
            }
        }
        if ($Hash.Contains("ConfigurationStandards")) {
            #find VMHosts for Infra an set it in Node-Object ConfigurationStandards
            $VMHosts = ($this.nodes.where({($_.Layer -match "INFRA") -and ($_.roles -match "HyperV")})).Name -join(",")
            $this.nodes.where({($_.Layer -match "INFRA") -or ($_.Layer -match "HSA") -or ($_.Layer -match "CM")}) | ForEach-Object {
                $_ | Add-Member -MemberType NoteProperty -Name VMHosts -Value $VMHosts
            }
            #find VMHosts for SSA an set it in Node-Object ConfigurationStandards
            $VMHosts = $this.nodes.where({($_.Layer -match "SSA") -and ($_.roles -match "HyperV")}) -join(",")
            $this.nodes.where({($_.Layer -match "SSA")}) | ForEach-Object {
                $_ | Add-Member -MemberType NoteProperty -Name VMHosts -Value $VMHosts
            }
        }
    }


#region methods
    add([Node]$Node){
        $this.Nodes += $Node
    }

    add([string]$Name,[Hashtable]$Hash){
        $this.Nodes += ([Node]::new($Name,$Hash))
    }

    
    [Network[]] getNetworks(){
        [Network[]]$Networks = @()

        $this.Nodes | ForEach-Object {
            $Node = $_

            $Node.Networks | ForEach-Object {
                $Network = $_
                if(!($Networks.where({$_.NetworkID -match $Network.NetworkID}))){
                    $Networks += $Network
                }

            }

        }
        $Networks | ForEach-Object{
            $Network = $_
            $Network.IPAddress = $null
        }
        return $Networks
    }

    [Node[]]getNodesbyName([String[]]$NodeName){
        #returns all Node Objects where $NodeName matches the NodeObjects NodeName Atribute
        $NodeArray = @()
        foreach ($Node in $NodeName) {
            $NodeArray += ($this.Nodes).where({$_.Nodename -match $Node})
        }
        return $NodeArray
    }
    [Hashtable]getConfigurationData(){
        [Hashtable]$ConfigurationData = @{
            AllNodes = @(

            );
            NoneNodeData = ""
        }

        $this.Nodes | ForEach-Object {
            $ConfigurationData.AllNodes += $_.getConfigurationData().AllNodes
        }


        return $ConfigurationData
    }

    save([String]$Path){
        $this.ConfigSource = $path
        $this.setNodesConfigSource()
        $this | Export-Clixml -Path $path -Force
    }
    save(){
        $item = Get-Item $this.ConfigSource
        
        if ($item.Extension -match "xml") {
            $this.save($this.ConfigSource)
        }
        else {
            $FileName = Read-Host -Prompt "Filename"
            $Path = Join-Path -Path ".\" -ChildPath $FileName
            $this.save($Path)
            $this.ConfigSource = (Get-Item -Path $Path).FullName
        }
    }

    load([String]$Path){
        $MyConfiguration = [MyConfiguration]::new($Path)
        $this.Nodes = $MyConfiguration.Nodes
        $this.ConfigSource = (get-item -path $Path).FullName
        $this.setNodesConfigSource()
    }

    load(){
        $item = Get-Item -Path $this.ConfigSource
        if ($item) {
            if ($item.Extension -match "xml") {
                $this.load($this.ConfigSource)
            }
            else {
                Write-Error -Message "Format not supported"
            }
        }
        else {
            Write-Error -Message "does not exists"
        }    
    }

    setNodesConfigSource(){
        foreach($Node in $this.Nodes){
            $Node | Add-Member -MemberType NoteProperty -Name "ConfigSource" -Value $this.ConfigSource -Force
        }
    }

 }

#region Configuration

Configuration InitializeDeployment {
    param(
        [System.Management.Automation.PSCredential]$LocalCredential,
        [System.Management.Automation.PSCredential]$DomainCredential
    )
    Import-DscResource -ModuleName xPSDesiredStateConfiguration
    Import-DscResource -ModuleName xComputerManagement
    Import-DscResource -ModuleName xnetworking
    Import-DscResource -ModuleName xActiveDirectory
    Import-DscResource -ModuleName xDNSServer
    Import-DscResource -ModuleName xDHCPServer
    Import-DscResource -ModuleName xPendingReboot


    #both nodes
    Node $AllNodes.where{(($_.Roles -match "PDC") -or ($_.Roles -match "WDS")) -and ($_.Layer -match "INFRA")}.Nodename {
        LocalConfigurationManager{
            RebootNodeIfNeeded = $true
            ActionAfterReboot = 'StopConfiguration'
        }
        xScript NetworkAdapterName {
            GetScript = {
                . C:\deploy\scripts\create-lab.ps1
                (Get-NetAdapter | Where-Object {$_.MacAddress -eq $Node.BootMac }).Name
            }

            SetScript = {
                . C:\deploy\scripts\create-lab.ps1
                $MGMTNetwork = $Node.Networks | Where-Object {$_.name -like "Mgmt*"}
                Get-NetAdapter | Where-Object {$_.MacAddress -eq $Node.BootMac } | Rename-NetAdapter -NewName $MGMTNetwork.Name -Verbose
                $MGMT= (Get-NetAdapter -Name $MGMTNetwork.Name)
                $MGMT | Get-NetAdapterBinding -ComponentID ms_tcpip6 | Disable-NetAdapterBinding -Verbose
            }

            TestScript = {
                . C:\deploy\scripts\create-lab.ps1
                $NetWorkAdapter = Get-NetAdapter | Where-Object {$_.MacAddress -eq $Node.BootMac }
                $MGMTNetwork = $Node.Networks | Where-Object {$_.name -like "Mgmt*"}
                if ($NetWorkAdapter.Name -eq $MgmtNetwork.Name ) {
                    return $true
                }
                else {
                    return $false
                }
            }
        }

        xScript DefragTask {
            GetScript = {
                                
                (Get-ScheduledTask ScheduledDefrag).State
            }
            SetScript = {
                Get-ScheduledTask ScheduledDefrag | Disable-ScheduledTask -Verbose
            }
            TestScript = {
                $Task = Get-ScheduledTask ScheduledDefrag
                if ($Task.state -eq "Disabled") {
                    return $true
                }else {
                    return $false
                }


            }
        }

        xDNSServerAddress DNSServer {
            Address = (($Node.Networks | Where-Object {$_.name -like "Mgmt*"}).DNSServer).split(",")
            InterfaceAlias = ($Node.Networks | Where-Object {$_.name -like "Mgmt*"}).Name
            AddressFamily = "IPv4"
            DependsOn = '[xscript]NetworkAdapterName'
        }

        xDnsConnectionSuffix DNSSuffix {
            InterfaceAlias = ($Node.Networks | Where-Object {$_.name -like "Mgmt*"}).Name
            ConnectionSpecificSuffix = ($Node.Networks | Where-Object {$_.name -like "Mgmt*"}).DNSSuffix
            DependsOn = '[xscript]NetworkAdapterName'
        }

        xNetAdapterBinding IPv6 {
            InterfaceAlias = ($Node.Networks | Where-Object {$_.name -like "Mgmt*"}).Name
            ComponentId    = 'ms_tcpip6'
            State          = 'Disabled'
            DependsOn = '[xscript]NetworkAdapterName'
        }

    }
    #only PDC
    Node $AllNodes.where{($_.Roles -match "PDC") -and ($_.Layer -match "INFRA")}.Nodename {
        LocalConfigurationManager{
            RebootNodeIfNeeded = $true
            ActionAfterReboot = 'StopConfiguration'
        }
        xWindowsFeatureSet ADDS {
            Name = "AD-Domain-Services","RSAT-AD-Powershell"
            Ensure = "Present"
        }
        xADDomain Forest {
            DomainName = $Node.Domain
            DomainNetbiosName = $Node.SMBDomain
            SafemodeAdministratorPassword = $LocalCredential
            DomainAdministratorCredential = $DomainCredential
            DependsOn = '[xwindowsfeatureset]ADDS'
        }
        xPendingReboot RebootAfterForest {
            Name = "AfterForest"
            DependsOn = '[xADDomain]Forest'
        }

        xWaitForADDomain ADDSWait {
            DomainName = $Node.Domain
            RetryCount = "10"
            RetryIntervalSec = "10"
            DependsOn = '[xADDomain]Forest'
        }
        xScript ReverseZone {
            GetScript = {
                . C:\deploy\scripts\create-lab.ps1
                (get-dnsserverzone).where({($_.isreverselookupzone -eq $true) -and ($_.isautocreated -eq $false)})
            }
            SetScript = {
                . C:\deploy\scripts\create-lab.ps1
                Import-Module DNSServer
                $MGMTNetwork = $Node.Networks | Where-Object {$_.name -like "Mgmt*"}
                $NetworkID=($MGMTNetwork.NetworkID+"/"+$MGMTNetwork.Prefix)
                Add-DnsServerPrimaryZone -NetworkId $NetworkID -ReplicationScope Forest -Verbose
                Start-Sleep -Seconds 5
                Register-DnsClient -Verbose | Out-Null
            }
            TestScript = {
                . C:\deploy\scripts\create-lab.ps1
                $ReverseZone = (get-dnsserverzone).where({($_.isreverselookupzone -eq $true) -and ($_.isautocreated -eq $false)})
                if ($ReverseZone) {
                    return $true
                }else {
                    return $false
                }
            }
            DependsOn = '[xWaitForADDomain]ADDSWait'
        }
        xScript "Domain Admins" {
            GetScript = {
                . C:\deploy\scripts\create-lab.ps1
                $groups = get-aduser | Get-ADPrincipalGroupMembership
                $groups.name
            }
            SetScript = {
                . C:\deploy\scripts\create-lab.ps1
                $user = get-ADUser -Identity $Node.User
                Add-ADGroupMember -Identity "Domain Admins" -Members $user
            }
            TestScript = {
                . C:\deploy\scripts\create-lab.ps1
                $result = $false
                $Groups= Get-ADUser $Node.User | Get-ADPrincipalGroupMembership
                foreach($Name in $Groups.name) {
                    if ($Name -eq "Domain Admins") {
                        $result = $true
                    }
                }
                return $result
            }
            DependsOn = '[xWaitForADDomain]ADDSWait'
        }


    }
    #only WDS
    Node $AllNodes.where{($_.Roles -match "WDS") -and ($_.Layer -match "INFRA")}.Nodename {
        LocalConfigurationManager{
            RebootNodeIfNeeded = $true
            ActionAfterReboot = 'StopConfiguration'
        }
        xWindowsFeatureSet DHCP_WDS {
            Name = "DHCP","WDS"
            Ensure = "Present"
            IncludeAllSubFeature = $true
        }
        
        xWaitForADDomain ADDSWait {
            DomainName = $Node.Domain
            DomainUserCredential = $DomainCredential
            RetryCount = "30"
            RetryIntervalSec = "30"
        }

        xComputer DomainJoin {
            Name = $Node.Name
            DomainName = $Node.Domain
            Credential = $DomainCredential
            DependsOn = '[xWaitForADDomain]ADDSWait'
        }

        xPendingReboot RebootAfterJoin {
            Name = "AfterJoin"
            DependsOn = '[xComputer]DomainJoin'
        }

        xDhcpServerAuthorization Authorization {
            Ensure = "Present"
            PsDscRunAsCredential = $DomainCredential
            DependsOn = '[xWaitForADDomain]ADDSWait'
        }

        xScript "Configuring DHCPScope" {
            GetScript = {
                Get-DhcpServerv4Scope -Verbose
            }

            SetScript = {
                . C:\deploy\scripts\create-lab.ps1
                #reset of DHCPScope
                Get-DhcpServerv4Scope | Remove-DhcpServerv4Scope -Force -Verbose
                #wait for dhcp
                while ((get-service dhcpserver).status -ne "Running") {
                    Write-Warning -Message "Waiting for DHCP-Server"
                    start-Service DHCPServer -Verbose
                    Start-Sleep -Seconds 5
                }
                $MGMTNetwork = $Node.Networks | Where-Object {$_.name -like "Mgmt*"}

                #create Scope
                
                $Mask = (Get-NetworkSummary ($MGMTNetwork.networkID+"/"+$MgmtNetwork.Prefix)).Mask
                $IPRange = Get-NetworkRange ($MGMTNetwork.networkID+"/"+$MgmtNetwork.Prefix)
                $StartIP = $IPRange[0]
                $EndIP = $IPRange[-1]
                $Scope = Add-DhcpServerv4Scope -StartRange $StartIP -EndRange $EndIP -SubnetMask $Mask -Name $MGMTNetwork.name -PassThru -Verbose
        

                #Scope ausschließen
                Add-DhcpServerv4ExclusionRange -ScopeId $Scope.ScopeId -StartRange $StartIP -EndRange $EndIP -Verbose
 

                #DHCP Brereichsoptionen anlegen
                Set-DhcpServerv4OptionValue -ScopeId $Scope.ScopeId -OptionId 006 -Value $MGMTNetwork.DNSServer.split(",") -Force
                Set-DhcpServerv4OptionValue -ScopeId $Scope.ScopeId -OptionId 015 -Value $MGMTNetwork.DNSSuffix -Verbose
                Set-DhcpServerv4OptionValue -ScopeId $Scope.ScopeId -OptionId 066 -Value $MGMTNetwork.IPAddress -Verbose
                Set-DhcpServerv4OptionValue -ScopeId $Scope.ScopeId -OptionId 067 -Value "boot\x64\wdsnbp.com" -Verbose

        
                #Reservierungen einrichten
                $Nodes = $MyConfiguration.nodes | Where-Object {$_.BootMac -ne ""};
                foreach($Node in $Nodes) {
                    $MGMTNetwork = $Node.Networks | Where-Object {$_.name -like "Mgmt*"};
                    Add-DhcpServerv4Reservation -ScopeId $Scope.ScopeId -IPAddress $MGMTNetwork.IPAddress -ClientId $Node.BootMac -Name $Node.Name -Verbose
                }
            }

            TestScript = {
                return $false
                
            }
            DependsOn = "[xDhcpServerAuthorization]Authorization"
        }
        xRegistry "Mark DHCP as configured" {
            Key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ServerManager\Roles\12"
            ValueName = "ConfigurationState"
            ValueData = "2"
            Force = $true
            DependsOn = "[xScript]Configuring DHCPScope"
        }

        xscript  "Configuring WDSServer" {
            GetScript = {
                WDSUTIL /Get-Server /Show:Config
            }
            SetScript = {
                . C:\deploy\scripts\create-lab.ps1
                wdsutil.exe /initialize-server /reminst:c:\remoteinstall
                while ((Get-Service *WDS*).status -ne "running") {
                    start-sleep -Seconds 3
                    get-service *WDS* | start-service -verbose
                }
                WDSUTIL.exe /Uninitialize-Server
                Remove-Item -Force -Recurse -Path "C:\remoteinstall" -Verbose 
                start-sleep -Seconds 2 -Verbose
                wdsutil.exe /initialize-server /reminst:c:\remoteinstall
                
                #Create Imagegroup
                new-wdsinstallimagegroup -name "Deploy" -erroraction silentlycontinue -verbose | Out-Null
                $ImageGroup = (get-wdsinstallimagegroup -name "Deploy").Name
        
                #Unattend WDS Setup file creation
            
                $BootUnattend="c:\remoteinstall\wds_boot.xml"
                $WDSBootXML = [xml] $Node.unattendXML()
                $WDSBootXML.unattend.settings[1].component[1].WindowsDeploymentServices.ImageSelection.InstallImage.ImageName="Install"
                $WDSBootXML.unattend.settings[1].component[1].WindowsDeploymentServices.ImageSelection.InstallImage.ImageGroup="Deploy"
                $WDSBootXML.unattend.settings[1].component[1].WindowsDeploymentServices.ImageSelection.InstallImage.Filename= $Node.ConfigurationStandards.template_Windows2012R2.value
                $WDSBootXML.Save($BootUnattend)
        
        
                #Unattend Install Setup creation
                $InstallUnattend="c:\remoteinstall\wds_install.xml"
                $WDSBootXML.Save($InstallUnattend)
        
                #Serveroptions
                #Bootprogram
                #Server Authorization
                #awnser all clients
                #Standardimage settings on x64
                #allways install - abort with ESC
                WDSUTIL.exe /Set-Server /BootProgram:boot\x64\wdsnbp.com /Architecture:x64 
                wdsutil.exe /set-server /authorize:yes
                wdsutil.exe /set-server /answerclients:All
                wdsutil.exe /set-server /DefaultX86X64ImageType:x64
                wdsutil.exe /set-server /PxePromptPolicy /new:optout
                wdsutil.exe /set-server /PxePromptPolicy /known:optout

                #Boot- und Install-Image einrichten und Installation unattended machen
                wdsutil.exe /progress /verbose /add-image /ImageFile:C:\deploy\boot.wim /ImageType:Boot /skipverify
                Mount-DiskImage C:\deploy\template_windows2012R2.vhdx -Verbose
                get-disk 1 | Set-Disk -IsOffline:$false
                Start-Sleep -Seconds 2
                Dismount-DiskImage C:\deploy\template_windows2012R2.vhdx
                wdsutil.exe /progress /verbose /add-image /ImageFile:C:\deploy\template_windows2012R2.vhdx /ImageType:Install /ImageGroup:Deploy /skipverify
                Get-WdsInstallImage | Set-WdsInstallImage -NewImageName Install -verbose
                #$BootImage=($BootWIM.Split("\")[($BootWIM.Split("\")).length-1])
                wdsutil.exe /set-server /BootImage:boot.wim /Architecture:x64
                $BootUnattendShort=($BootUnattend.Split("\")[($BootUnattend.Split("\")).length-1])
                WDSUTIL.exe /Set-Server /WdsUnattend /Policy:Enabled /File:$BootUnattendShort /Architecture:x64

                #Multicast-Transmission einrichten, damit Deploy auf mehrere Rechner schneller geht
                wdsutil.exe /new-multicasttransmission /friendlyname:"WDS Autocast Default" /image:"Install" /ImageType:Install /transmissiontype:autocast
            }
            TestScript = {
                return $false
            }

            DependsOn = "[xScript]Configuring DHCPScope"
        }
        xScript "Prestaging" {
            GetScript = {
                get-wdsclient
            }
            SetScript = {
                . C:\deploy\scripts\create-lab.ps1
                #get all Nodes with bootmac and create prestaged clients
                $Nodes = $MyConfiguration.nodes | Where-Object {$_.BootMac -ne ""};
                foreach($Node in $Nodes) {
                    $MGMTNetwork = $Node.Networks | Where-Object {$_.name -like "Mgmt*"};
                    $ScriptDest = $Node.ConfigurationStandards.ScriptPath.value
                    $BootUnattendDest = Join-Path -Path "c:\remoteinstall\WdsClientUnattend" -ChildPath ($Node.Name + ".xml")
                    $BootUnattend = $Node.unattendXML()
                    $BootUnattend.Save($BootUnattendDest)
                    New-WDSClient -DeviceID $Node.BootMac -DeviceName $Node.Name -WdsClientUnattend $BootUnattendDest -JoinRights Full -JoinDomain $true -User ($Node.Domain+"\Administrator") -ErrorAction silentlycontinue -verbose
                }
            }
            TestScript = {

                return $false
            }
            PsDscRunAsCredential = $DomainCredential
            DependsOn = "[xscript]Configuring WDSServer"
        }
    }
    

 }

Configuration DeployHosts {
        param(
        [System.Management.Automation.PSCredential]$LocalCredential,
        [System.Management.Automation.PSCredential]$DomainCredential
    )
    Import-DscResource -ModuleName xPSDesiredStateConfiguration
    Import-DscResource -ModuleName xComputerManagement
    Import-DscResource -ModuleName xnetworking
    Import-DscResource -ModuleName xActiveDirectory
    Import-DscResource -ModuleName xDNSServer
    Import-DscResource -ModuleName xDHCPServer
    Import-DscResource -ModuleName xPendingReboot
    Import-DscResource -modulename xNetworking

    Node $AllNodes.Where{($_.Roles -match "hyperV") -and ($_.BootMac -ne "")}.NodeName {
        LocalConfigurationManager{
            RebootNodeIfNeeded = $true
            ActionAfterReboot = 'StopConfiguration'
        }

        xScript NetworkDriver {
            GetScript = {
                Get-NetAdapter
            }
            SetScript = {
              . C:\deploy\scripts\create-lab.ps1
              $FilePath = Join-Path -Path $Node.ConfigurationStandards.SoftwarePath.value -childpath "BroadcomNetXtremeII\setup.exe"
              Write-Verbose "Broadcom-Treiber werden installiert"
              Start-Process -Wait -FilePath $FilePath -Verb runas -ArgumentList " /s /v/qn" | Wait-Process
              $FilePath = Join-Path -path $Node.ConfigurationStandards.ScriptPath.Value -childpath "NetworkDriver.txt"
              Get-NetAdapter | out-file -FilePath $FilePath -encoding UTF8
            }
            TestScript = {
                . C:\deploy\scripts\create-lab.ps1
                $FilePath = Join-Path -path $Node.ConfigurationStandards.ScriptPath.Value -childpath "NetworkDriver.txt"
                return (test-path -Path $FilePath)
            }

        }

        xScript NetworkAdapterName {
            GetScript = {
                . C:\deploy\scripts\create-lab.ps1
                (Get-NetAdapter | Where-Object {$_.MacAddress -eq $Node.BootMac }).Name
            }

            SetScript = {
                . C:\deploy\scripts\create-lab.ps1
                $MGMTNetwork = $Node.Networks | Where-Object {$_.name -like "Mgmt*"}
                Get-NetAdapter | Where-Object {$_.MacAddress -eq $Node.BootMac } | Rename-NetAdapter -NewName $MGMTNetwork.Name -Verbose
                $MGMT= (Get-NetAdapter -Name $MGMTNetwork.Name)
                $MGMT | Get-NetAdapterBinding -ComponentID ms_tcpip6 | Disable-NetAdapterBinding -Verbose
            }

            TestScript = {
                . C:\deploy\scripts\create-lab.ps1
                $NetWorkAdapter = Get-NetAdapter | Where-Object {$_.MacAddress -eq $Node.BootMac }
                $MGMTNetwork = $Node.Networks | Where-Object {$_.name -like "Mgmt*"}
                if ($NetWorkAdapter.Name -eq $MgmtNetwork.Name ) {
                    return $true
                }
                else {
                    return $false
                }
            }
            DependsOn = "[xScript]NetworkDriver"
        }

        xScript DefragTask {
            GetScript = {
                                
                (Get-ScheduledTask ScheduledDefrag).State
            }
            SetScript = {
                Get-ScheduledTask ScheduledDefrag | Disable-ScheduledTask -Verbose
            }
            TestScript = {
                $Task = Get-ScheduledTask ScheduledDefrag
                if ($Task.state -eq "Disabled") {
                    return $true
                }else {
                    return $false
                }


            }
        }

        xDNSServerAddress DNSServer {
            Address = (($Node.Networks | Where-Object {$_.name -like "Mgmt*"}).DNSServer).split(",")
            InterfaceAlias = ($Node.Networks | Where-Object {$_.name -like "Mgmt*"}).Name
            AddressFamily = "IPv4"
            DependsOn = '[xscript]NetworkAdapterName'
        }

        xDnsConnectionSuffix DNSSuffix {
            InterfaceAlias = ($Node.Networks | Where-Object {$_.name -like "Mgmt*"}).Name
            ConnectionSpecificSuffix = ($Node.Networks | Where-Object {$_.name -like "Mgmt*"}).DNSSuffix
            DependsOn = '[xscript]NetworkAdapterName'
        }

        xNetAdapterBinding IPv6 {
            InterfaceAlias = ($Node.Networks | Where-Object {$_.name -like "Mgmt*"}).Name
            ComponentId    = 'ms_tcpip6'
            State          = 'Disabled'
            DependsOn = '[xscript]NetworkAdapterName'
        }
        xWindowsFeatureSet ADDS {
            Name = "Hyper-V","Hyper-V-Tools","Hyper-V-Powershell","Failover-Clustering"
            Ensure = "Present"
        }
    }
}

#region Workflow Functions
function get-MyConfiguration {
    <#
    .Synopsis
     gets the Configuration from existing excel file and outputs it as MyConfiguration Object
    .Example
     $MyConfiguration = get-MyConfiguration
    #>

    $MyConfiguration = [MyConfiguration]::new($ConfigurationPath)
    Write-Output $MyConfiguration
}

function init-Deployment {
    param()
    Copy-Item -Path ".\create-lab.ps1" -Destination "C:\deploy\scripts\" -Force
    Set-Item -Path WSMan:\localhost\Client\TrustedHosts -Value * -Force
    $MyConfiguration = [MyConfiguration]::new(".\Configuration.xlsx")
    $MyConfiguration.save(".\Configuration.xml")
    $MyConfiguration.save()
    $Nodes = $MyConfiguration.nodes | where-Object {($_.layer -eq "INFRA") -and (($_.Roles -match "pdc") -or ($_.Roles -match "wds"))}

    #create baremetal-template
    $BareMetalNode = ($MyConfiguration.nodes).where({$_.isvirtual -ne $true})[0]
    $BareMetalNode.createMasterImage()

    $BareMetalImage = $BareMetalNode.createVHD()
    $BareMetalDestination = join-path -path $BareMetalNode.ConfigurationStandards.HyperVVHDPath.Value -childpath "BareMetalImage.vhdx"
    Move-Item -Path $BareMetalImage.Path -Destination $BareMetalDestination -Force


    #create Nodes and start initial configuration    
    foreach ($Node in $Nodes ) {
        # create entry in local hosts-file for node
        $MgmtNetwork = $Node.networks | Where-Object {$_.name -like "Mgmt*"}
        $IPAddress = $MgmtNetwork.IPAddress
        $Entry = $IPAddress,$Node.fqdn() -join(" ")
        $Entry | Out-File -Append -FilePath C:\Windows\System32\drivers\etc\hosts -Encoding utf8
        #create the VMs and apply configuration
        $Node.createMasterImage() | Out-Host
        $Node.createVHD() | Out-Host
        $Node.createVM() | Out-Host
        $Node.getVM() | Start-VM
        vmconnect.exe localhost $Node.Name
        Copy-Item -Path ".\configuration.xml" -Destination "C:\deploy\scripts\"
        
    }
    #start the dsc-magic till WDS is ready
    $Nodes.wait()
    InitializeDeployment -LocalCredential $MyConfiguration.Nodes[0].getLocalCredentials() -DomainCredential $MyConfiguration.Nodes[0].getDomainCredentials()  -ConfigurationData $MyConfiguration.getConfigurationData()
    Set-DscLocalConfigurationManager -Verbose -Path .\InitializeDeployment -Credential $Node[0].getLocalCredentials() -Force
    start-DscConfiguration -Verbose -Path .\InitializeDeployment\ -Credential $Node[0].getLocalCredentials() -wait -Force
    $Nodes.wait()
    start-DscConfiguration -Verbose -Path .\InitializeDeployment\ -Credential $Node[0].getLocalCredentials() -wait -Force

    $nodes
}


function reset-Deployment {
    param()
    get-content -Path C:\Windows\System32\drivers\etc\hosts | Select-String -SimpleMatch "t02s02" -notmatch | Out-File C:\Windows\System32\drivers\etc\hosts -Encoding utf8
    Set-Item -Path WSMan:\localhost\Client\TrustedHosts -Value "" -Force
    get-vhd 'C:\Hyper-V\Virtual Disks\*' | dismount-vhd -ErrorAction SilentlyContinue
    get-vm Y* | stop-vm -Passthru -Force | remove-vm -Force
    get-process *vmconnect* | stop-process
    get-job | Stop-Job 
    get-job | remove-job
    Remove-Item 'C:\hyper-v\Virtual Disks\Y*' -Force
    Remove-Item 'C:\hyper-v\Virtual Machines\*' -Recurse -Force

}

function start-deployment {
    Import-Module -Name IMM-Module -Force
    $MyConfiguration = [MyConfiguration]::new(".\configuration.xml")
    $Nodes = ($MyConfiguration.nodes).where({ ($_.Roles -match "hyperv") -and ($_.Layer -match "INFRA")})
    Reboot-IMMServerOS -IMM $Nodes.Name



}

$MyConfiguration = [MyConfiguration]::new("c:\deploy\scripts\Configuration.xml")
$Node = $MyConfiguration.getNodesbyName($env:COMPUTERNAME)