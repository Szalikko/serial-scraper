import subprocess, sys, ctypes
from datetime import datetime   

# first messagebox
def MessageBox():
    mbox = ctypes.windll.user32.MessageBoxW(0, "Uruchomic skrypt?", "serial-scraper d-_-b", 3 | 0x40)
    return mbox

# second messagebox
def FinalBox():
    mbox = ctypes.windll.user32.MessageBoxW(0, "Gotowe", "serial-scraper d-_-b", 0 | 0x40)
    return mbox

# get pc serial number and model
def PcInfoScraper():
    serial = subprocess.check_output('powershell.exe -WindowStyle Hidden Get-WmiObject win32_bios | select Serialnumber | Format-Table -HideTableHeaders').decode("utf-8")
    manufacturer = subprocess.check_output('powershell.exe -WindowStyle Hidden Get-WmiObject win32_bios | select Manufacturer | Format-Table -HideTableHeaders').decode("utf-8")
    model = subprocess.check_output('powershell.exe -WindowStyle Hidden Get-CimInstance Win32_ComputerSystemProduct | Select Name | Format-Table -HideTableHeaders').decode("utf-8")
    result = "Manufacturer : " + manufacturer.strip() + "\nModel        : " + model.strip() + "\nSerial       : " + serial.strip() + '\n'
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

# get time
def time():
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    return dt_string

# save all gathered data into output.txt
def output():
    # create file for output
    output = open("output.txt", "w")

    # write variables to file
    output.write("------ Komputer: ------\n")
    output.write(PcInfoScraper())
    output.write("\n\n------ Monitory: ------\n")
    output.write(MonitorInfoScraper().replace('\x00', ''))
    output.write("\n"*3 + "------   Czas:   ------\n")
    output.write(time())

# main function
def main():
    # message boxes
    # Yes = 6
    # No = 7
    # Cancel = 2
    # Okcancel = 1

    input = MessageBox()

    # OK case
    if (input == 6):
        PcInfoScraper()
        MonitorInfoScraper()
        output()
        FinalBox()
        sys.exit()
    # NO case
    elif (input == 7):
        sys.exit()
    # Cancel case
    else:
        sys.exit()
    
# main function
main()

# d-_-b