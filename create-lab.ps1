param (
    $WORKDIR = "d:\temp",
    $ConfigurationPath = ".\configuration.csv",
    $NetworkDefinitionsPath = ".\NetworkDefinitions.csv"
)

class MyConfiguration {
    [STRING]$Name
    [bool]$isVirtual
    [bool]$isPhysical
    [System.Management.Automation.Runspaces.PSSession]$Session
    MyConfiguration([STRING]$Name,[STRING]$ConfigurationPath){
        $this.Name = $Name
        $Configuration = Import-Csv -Path $ConfigurationPath -Delimiter ";" 
        
        $Configuration | foreach {
            $config = $_
            #search for name
            if($Config.Hostname -match $this.Name){
                $Config | Get-Member -MemberType NoteProperty | foreach {
                    $this | Add-Member -Name $_.Name -MemberType $_.MemberType -Value $_.value
                    $this.($_.Name) = $Config.($_.Name)
                }
            }
            #search for bootmac
            elseif ($config.BOOTMAC -match $this.Name) {
                $this.Name = $Config.Hostname
                $Config | Get-Member -MemberType NoteProperty | foreach {
                    $this | Add-Member -Name $_.Name -MemberType $_.MemberType -Value $_.value
                    $this.($_.Name) = $Config.($_.Name)
                }
            }
        }
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

    append(){
        # append to configuration if it not exists
    }

    save(){
        # save to csv 
    }
}

