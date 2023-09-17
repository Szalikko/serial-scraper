
# create file for output
output = open("output.txt", "w")

# get pc serial number and model
serial = subprocess.check_output('powershell.exe Get-WmiObject win32_bios | select Serialnumber').decode("utf-8") 
model = subprocess.check_output('powershell.exe Get-CimInstance Win32_ComputerSystemProduct | Select Name').decode("utf-8")

