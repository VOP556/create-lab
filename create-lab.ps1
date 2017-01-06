param (
    $WORKDIR = ".\",
    $ConfigurationPath = ".\configuration.xlsx"
 )

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
                
                if($this.($_.Name) -is [STRING]){
                    $LastIndexOf = ($this.($_.Name)).lastindexof(".")
                    if ($LastIndexOf -gt 0) {
                        $Value = ($this.($_.Name)).substring(0,$LastIndexOf)
                        $Hash.Networks.keys | ForEach-Object {
                            if($Hash.Networks.$_ -match $Value){
                                $Network = [Network]::new($Hash.Networks.$_)
                                $Network.IPAddress = $this.$_
                            }
                        }
                    }
                }
            }
        }

        if ($Hash.contains("ConfigurationStandards")) {
            $this | Add-Member -MemberType NoteProperty -Name "ConfigurationStandards" -value $Hash.ConfigurationStandards -force
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

         $this | Get-Member -MemberType NoteProperty | ForEach-Object {
            $ConfigurationData.AllNodes[0].add($_.Name,$this.($_.Name))
         }

         $this | add-member -membertype NoteProperty -name ConfigurationData -value $ConfigurationData

         return $this.ConfigurationData
    }

    serializeToXML([string]$path){
        $this | Export-Clixml -Path $path
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
        return $Networks
    }


    [Hashtable]ConfigurationData(){
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



