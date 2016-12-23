param (
    $WORKDIR = "d:\temp",
    $ConfigurationPath = ".\configuration.xlsx",
    $NetworkDefinitionsPath = ".\NetworkDefinitions.csv"
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
    $WorkBook.Close()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WorkBook)
    $DataHashTable
}


class Network {
    [ValidateSet("PLATTFORM", "INFRA", IgnoreCase = $true)][String]$Layer
    [string]$IPv4Address
    [int16]$Prefix
    [String]$Name
    [int16]$VLAN
    [string]$Description
    [String]$DNSSuffix
    [String]$DNSServer

    setIPAddress(){

    }

    setDNSConfiguration(){

    }

    serialize(){

    }
}

class Node {
    [STRING]$Name
    [bool]$isVirtual
    [bool]$isPhysical
    [System.Management.Automation.Runspaces.PSSession]$Session
    [Network[]]$Network
    [String]$Description
    [string]$Roles

    createSession(){

    }
    exists(){
        $Result = $false
        if($this.isPhysical){
            $Result = (Test-Connection -Quiet -Count 1 -ComputerName $this.Name)
         
        }
        if ($this.isVirtual){
            $Result 
        }
        $Result
    }
    create(){
        # create vm if virtual or create prepare wim and wds to boot from
    }

    serialize(){

    }
}
$Data=get-ExcelDataHashTable -path .\Configuration.xlsx
$Data.Networks



