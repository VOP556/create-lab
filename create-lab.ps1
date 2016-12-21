param (
    $WORKDIR = "d:\temp",
    $ConfigurationPath = ".\configuration.csv",
    $NetworkDefinitionsPath = ".\NetworkDefinitions.csv"
)

class Network {
    [string]$IPv4Address
    [int16]$Prefix
    [String]$Name
    [int16]$VLAN
    [String]$DNSSuffix
    [String]$DNSServer

    setIPAddress(){

    }

    setDNSConfiguration(){

    }

    
}

class Node {
    [STRING]$Name
    [bool]$isVirtual
    [bool]$isPhysical
    [System.Management.Automation.Runspaces.PSSession]$Session
    [array]$Network
    [String]$Description
    [string]$Roles


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

    save(){
        # save to csv 
    }
}

