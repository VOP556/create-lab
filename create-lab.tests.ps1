
. ./create-lab.ps1

Describe "InputString ConfigurationPath"{
    it "should contain a String" {
        $ConfigurationPath = ".\Configuration.xlsx"
        $ConfigurationPath | should beOfType [String]
        
    }
    it "should get an item" {
        get-item $ConfigurationPath | should beOfType [System.IO.FileInfo]
    }

    it "should be the Item of that string"{
        (get-item $ConfigurationPath).FullName | Should beLike "*$ConfigurationPath*"
    }
}

Describe "get-exceldatahashtable" {
    Context "get-exceldatahashtable returns"{
        $Configuration = get-exceldatahashtable -path $configurationPath
        it 'get-exceldatahashtable -path [$configurationPath] should return a Hashtable' {            
            $Configuration | Should beOfType [Hashtable]
        }
        $Configuration.keys | ForEach-Object {
            $Key = $_
            it "$Key should return a Hashtable" {
                $Configuration.$Key | Should beOfType [Hashtable]
            }
            $Configuration.$Key.Keys | ForEach-Object {
                it "$_ should return a Hashtable" {
                    $Configuration.$Key.$_ | Should beOfType [Hashtable]
                }
            }
        }
    }
}

Describe "Class node" {
    Context "Constructors" {
        $Configuration = get-exceldatahashtable -path $configurationPath
        it '[Node]::new("Name") should return a Node Object' {
            [Node]::new("Name") | should beOfType [Node]
        }
        it '[Node]::new("Y02PINHST001",$Configuration) should return a Node Object' {
            
            [Node]::new("Y02VINPDC001",$Configuration) | should beOfType [Node]
        }
    }
    Context "Methods" {
        $Configuration = get-exceldatahashtable -path $configurationPath
        $Node = [Node]::new("Y02VINPDC001",$Configuration)
        it '$Node should be a Node Object'{
            $Node | should beOfType [Node]
        }
        it '$Node.getMasterImage() should return a VHD'{
            $Node.getmasterimage() | Should beOfType [Microsoft.Vhd.PowerShell.VirtualHardDisk]
        }
        it '$Node.getvhd() should return a VHD' {
            $Node.getvhd() | Should beOfType [Microsoft.Vhd.PowerShell.VirtualHardDisk]
        }
    }
}

Describe "Class MyConfiguration"{
    Context "Constructors"{
        $MyConfiguration = [MyConfiguration]::new($ConfigurationPath)
        it 'new([String]) should return Type MyConfiguration' {
            $MyConfiguration | should beOfType [MyConfiguration]
        }
        it "Property Nodes should be of Type [Node]" {
            $MyConfiguration.Nodes | should beOfType [Node]
        }

        it "every Node has to have a property ConfigurationStandards" {
            $MyConfiguration.nodes | ForEach-Object {
                ($_ | Get-Member -Name ConfigurationStandards).Name | Should be "ConfigurationStandards"
            }
        }
        it "every Node has to have a property Networks" {
            $MyConfiguration.nodes | ForEach-Object {
                ($_ | Get-Member -Name Networks).Name | Should be "Networks"
            }
        }
        it "every Node's Networks Property has to have a Network Object inside" {
            $MyConfiguration.nodes | ForEach-Object {
            $_.Networks | Should beOfType [Network]
            }
        }
    }
}