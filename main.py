import subprocess


# create file for output
output = open("output.txt", "w")

# get pc serial number and model
serial = subprocess.check_output('powershell.exe Get-WmiObject win32_bios | select Serialnumber').decode("utf-8") 
model = subprocess.check_output('powershell.exe Get-CimInstance Win32_ComputerSystemProduct | Select Name').decode("utf-8")

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
        
        echo "Manufacturer: $Manufacturer`nName: $Name`nSerial Number: $Serial`n"
    }
    """]).decode("utf-8")
    return result

# write variables to file
output.write("------ Komputer: ------")
output.write(model)
output.write(serial)
output.write("------ Monitor: ------\n")
output.write(MonitorInfoScraper())
