#Requires -Version 5
#Requires -Modules xComputerManagement
#Requires -RunasAdministrator
param (
    $WORKDIR = ".\",
    $ConfigurationPath = ".\configuration.xlsx"
 )

#region Includes and imports
 . .\convert-windowsimage.ps1

function get-ExcelDataHashTable {
    #written by Jörg Zimmermann
    #www.burningmountain.de
    #imports all worksheets from excel into hashtables
    #define the header per parameter
    #define the first row per parameter
    param(
        [Parameter(Mandatory=$true)][string]$path,
        [int]$HeaderRow=1,
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

Function Set-VMNetworkConfiguration {
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
     [Microsoft.Vhd.PowerShell.VirtualHardDisk]getMasterImage(){
    
        return (get-vhd -Path $this.TemplatePath)
       
     }

    [Microsoft.Vhd.PowerShell.VirtualHardDisk]createMasterImage(){
        
        #create MasterImage from ISO and integration of Packages
        $ISO = $null
        $VHDPATH = $null
        if ($this.OS -eq "Windows Server 2012 R2") {
            $ISO=$this.ConfigurationStandards.ISOPath.Value+"\"+$this.ConfigurationStandards.ISO_Windows2012R2.value
            $VHDPATH = $this.ConfigurationStandards.HyperVVHDPath.Value+"\"+$this.ConfigurationStandards.template_Windows2012R2.Value
        }
        elseif ($this.OS -eq "Windows Server 2012") {
            $ISO=$this.ConfigurationStandards.ISOPath.Value+"\"+$this.ConfigurationStandards.ISO_Windows2012.value
            $VHDPATH = $this.ConfigurationStandards.HyperVVHDPath.Value+"\"+$this.ConfigurationStandards.template_Windows2012.Value
        }
        else {
            Write-Error ($this.OS+" is not a valid OS")
        }
        
        $ISOItem = Test-Path -Path $ISO
        $VHDPATHItem = Test-Path -Path $VHDPATH
        if (!$ISOItem) {
            Write-Error "$ISO not found"
        }

        if ($VHDPATH) {
            Write-Information "$VHDPATH exists"
        }

        if (($ISOItem) -and (!$VHDPATHItem) ) {
            $WIM = ((Mount-DiskImage -ImagePath $ISO -StorageType ISO -PassThru | Get-Volume).DriveLetter+":\sources\install.wim")
            Convert-WindowsImage -SourcePath $WIM -VHDPath $VHDPATH -Edition ($this.OS+" SERVERDATACENTER") -VHDFormat VHDX -Verbose -VHDType Dynamic -SizeBytes 50GB -VHDPartitionStyle MBR
            dismount-DiskImage $ISO
            <# ToDo integration of Packages and Updates 
            Put the Dism-Magic here:
            Mount VHDX
            integrate Packages and Drivers
            dismount vhdx
            #>
            $this | Add-Member -MemberType NoteProperty -Name TemplatePath -Value $VHDPATH
        }


        return $this.getMasterImage()
    }

    [Microsoft.Vhd.PowerShell.VirtualHardDisk] createVHD(){
        #copy the MasterImage with new name to new filelocation
        $VHDPath = Join-Path -path $this.ConfigurationStandards.HyperVVHDPath.Value -ChildPath ($this.Name,"vhdx" -join(".")) 
        $TempplatePath = ($this.getMasterImage()).Path
        if (!$TempplatePath) {
            Write-Error "$TempplatePath not found"
        }
        if (!$this.getVHD()) {
            Copy-Item -Path $TempplatePath -Destination $VHDPath -Verbose -Force
            $this | Add-Member -MemberType NoteProperty -Name VHDPath -Value $VHDPath
        }
        return $this.getVHD()
    }

    [Microsoft.Vhd.PowerShell.VirtualHardDisk] getVHD(){
        return (get-vhd -path $this.VHDPath)        
    }

    [Microsoft.HyperV.PowerShell.VirtualMachine]createVM(){
        $VHDPath = $this.getVHD().Path
        $VM = $null
        if ($this.isVirtual) {
            try {
                get-vm -id $this.VMId -ErrorAction Stop
            }
            catch [Microsoft.HyperV.PowerShell.VirtualizationManagementException] {
                $VM = New-VM -Name $this.Name -MemoryStartupBytes ($this.memory+"MB") -Generation 1 -Path $this.ConfigurationStandards.hypervvmpath.value -VHDPath $VHDPath
                $this | Add-Member -MemberType NoteProperty -name VMId -Value $VM.VMId -Force
            }
        }

        return $this.getVM()
    }

    [Microsoft.HyperV.PowerShell.VirtualMachine]getVM(){
        return (get-vm -Name $this.Name)
    }
     
    [xml] unattendXML(){
        [xml] $unattend= [xml]"<?xml version='1.0' encoding='utf-8'?>
        <unattend xmlns='urn:schemas-microsoft-com:unattend'>
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
                  <PlainText></PlainText>
                </Password>
                <Enabled>true</Enabled>
                <LogonCount>1</LogonCount>
                <Username>Administrator</Username>
              </AutoLogon>
            <FirstLogonCommands>
                <SynchronousCommand wcm:action='add'>
                  <CommandLine>%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy RemoteSigned -File c:\scripts\start-config.ps1</CommandLine>
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
              </UserAccounts>
              <RegisteredOwner>-</RegisteredOwner>
              <RegisteredOrganization>-</RegisteredOrganization>
            </component>
          </settings>
          <cpi:offlineImage cpi:source='wim:d:/clusterinstall/install.wim#Windows Server 2012 SERVERDATACENTERCORE' xmlns:cpi='urn:schemas-microsoft-com:cpi' />
        </unattend>"  
        
        return $unattend
    }
    
}

class MyConfiguration {
    [Node[]] $Nodes


#region Constructors
    MyConfiguration([Hashtable]$Hash){
        $this.init($Hash)
    }

    MyConfiguration([string]$path){
        $Item = Get-Item -Path $path
        if ($Item.Extension -match 'xls') {
            $this.init((get-ExcelDataHashTable($path)))
        }
        elseif ($Item.Extension -match 'xml') {
            $xmlObjects = Import-Clixml -Path $Item.FullName
            $xmlObjects | ForEach-Object {
                $this.add([Node]::new($_))
            }
        }
        else {
            Write-Error -Message 'Format nicht unterstützt'
            break
        }
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
            $VMHosts = ($this.nodes.where({($_.Layer -match "INFRA") -and ($_.roles -match "Hyper-V")})).Name -join(",")
            $this.nodes.where({($_.Layer -match "INFRA") -or ($_.Layer -match "HSA") -or ($_.Layer -match "CM")}) | ForEach-Object {
                $_ | Add-Member -MemberType NoteProperty -Name VMHosts -Value $VMHosts
            }
            #find VMHosts for SSA an set it in Node-Object ConfigurationStandards
            $VMHosts = $this.nodes.where({($_.Layer -match "SSA") -and ($_.roles -match "Hyper-V")}) -join(",")
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

    exportXML([String]$Path){
        $this | Export-Clixml -Path $path

    }

 }

#region Configurations 
Configuration set-Computername {
    param(
        [String]$Computername,
        [String]$NodeName
    )
    Import-DscResource -Module xComputerManagement
   
    Node $NodeName {
        xComputer NewName {
            Name = $Computername
        }
        
    }
}


