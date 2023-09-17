import subprocess


# create file for output
output = open("output.txt", "w")

# get pc serial number and model
def PcInfoScraper():
    serial = subprocess.check_output('powershell.exe Get-WmiObject win32_bios | select Serialnumber | Format-Table -HideTableHeaders').decode("utf-8")
    model = subprocess.check_output('powershell.exe Get-CimInstance Win32_ComputerSystemProduct | Select Name | Format-Table -HideTableHeaders').decode("utf-8")
    result = "Model        : " + model.strip() + "\nSerial       : " + serial.strip() + '\n'
    return result
                    
# get monitors serial and model name
def MonitorInfoScraper():
    result = subprocess.check_output(["powershell.exe", """
    function Decode {
        If ($args[0] -is [System.Array]) {
            [System.Text.Encoding]::ASCII.GetString($args[0])
        }
        Else {
            "Not Found"
        }
    }

    ForEach ($Monitor in Get-WmiObject WmiMonitorID -Namespace root\wmi) {  
        $Manufacturer = Decode $Monitor.ManufacturerName -notmatch 0
        $Name = Decode $Monitor.UserFriendlyName -notmatch 0
        $Serial = Decode $Monitor.SerialNumberID -notmatch 0
        
        echo "`nManufacturer : $Manufacturer`nName         : $Name`nSerial Number: $Serial"
    }
    """]).decode("utf-8")
    return result.strip()


# write variables to file
output.write("------ Komputer: ------\n")
output.write(PcInfoScraper())
# output.write(model.strip()+'\n')
# output.write(serial.strip()+'\n')
output.write("\n\n------ Monitory: ------\n")
output.write(MonitorInfoScraper())

