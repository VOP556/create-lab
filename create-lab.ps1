$WORKDIR = "D:\ClusterInstall"
$ConfigurationPath = "D:\Cloud\upload\Deploy\IP_Passwords.csv"
$NetworkDefinitionsPath = "D:\doku\scripts\Deploy\NetworkDefinitions.csv"

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
            #suche nach dem namen
            if($Config.Hostname -match $this.Name){
                $Config | Get-Member -MemberType NoteProperty | foreach {
                    $this | Add-Member -Name $_.Name -MemberType $_.MemberType -Value $_.value
                    $this.($_.Name) = $Config.($_.Name)
                }
            }
            #suche nach der mac
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


}
