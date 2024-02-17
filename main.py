import math

import PySimpleGUI as sg
import platform
import psutil
import os
import wmi
import pandas as pd
import pyautogui,time

# --------------------
# os info
# -------------------


platform.node()
sys = platform.uname()

computer = wmi.WMI()
computer_info = computer.Win32_ComputerSystem()[0]
os_info = computer.Win32_OperatingSystem()[0]
proc_info = computer.Win32_Processor()[0]
gpu_info = computer.Win32_VideoController()[0]

os_name = os_info.Name.split('|')[0].strip()
os_version = ' '.join([os_info.Version, os_info.BuildNumber])
system_ram = float(os_info.TotalVisibleMemorySize) / 1048576  #
new_system_ram = str(system_ram)[:-13]

ws = wmi.WMI(namespace='root/Microsoft/Windows/Storage')


# -------------------------------------------------------------------------


totalSSD = int()
freeSSD  = int()
totalHDD = int()
freeHDD  = int()
total = int()
used= int()
free = int()


wss = wmi.WMI(namespace="root\\CIMv2")
disks = wss.Win32_DiskDrive()
logical_disks = wss.Win32_LogicalDisk()

for psu,disk in zip(psutil.disk_partitions(),disks):
    if disk.MediaType == "Fixed hard disk media":
        for d in ws.MSFT_PhysicalDisk():
            if d.Model == disk.Model:


                if d.MediaType == 4:
                    totalSSD = totalSSD + int(disk.Size) // (1024 ** 3)
                    freeSSD = freeSSD + free // (2 ** 30)

                if d.MediaType == 3:
                    totalHDD = totalHDD + int(disk.Size) // (1024 ** 3)
                    freeHDD = freeHDD + free // (2 ** 30)

ssdGB = totalSSD
ssdFree = freeSSD
hddGB = totalHDD
hddFree = freeHDD


if ssdGB ==0:
    ssdGB = 'x'
if ssdFree == 0:
    ssdFree = 'x'
if hddGB == 0:
    hddGB = 'x'
if hddFree == 0:
    hddFree = 'x'


# ----------------------
# RAM-------------------
# ----------------------

import subprocess

def run_powershell_command(command):
    try:
        result = subprocess.run(["powershell", "-Command", command], capture_output=True, text=True, check=True)
        return result.stdout.strip()
    except subprocess.CalledProcessError as e:
        return None

powershell_command = 'Get-WmiObject -Class "Win32_PhysicalMemoryArray"'
output = run_powershell_command(powershell_command)


c = wmi.WMI()
x = int()
def get_memory_slots():
    c = wmi.WMI()
    for mem_module in c.Win32_PhysicalMemory():
        x = 1
    return len(c.Win32_PhysicalMemory())

memory_slots = get_memory_slots()

def run_powershell_command_for_DDR(command):
    try:
        result = subprocess.run(["powershell", "-Command", command], capture_output=True, text=True, check=True)
        return result.stdout.strip()
    except subprocess.CalledProcessError as e:
        return None

run_powershell_command_ddr = 'wmic memorychip get SMBIOSMemoryType'
output_ddr = run_powershell_command_for_DDR(run_powershell_command_ddr)

value_error_message = output_ddr
number_string = value_error_message.split()[-1]  # Extract the last part of the string which contains the number
number_ddr = int(number_string.strip())  # Remove leading/trailing whitespace and convert to integer

DDR_out = int()

if number_ddr == 26:
    DDR_out = 4
elif number_ddr == 24:
    DDR_out = 3
elif number_ddr == 21:
    DDR_out = 2
elif number_ddr == 20:
    DDR_out = 1
else:
    DDR_out = 0

# --------------------
#GUI
# --------------------


sg.theme('BrownBlue')


layout = [
    [sg.Text('FileBrowse',justification='center')],
    [sg.Input(expand_x=True, key="-FilePath-"),sg.FileBrowse(file_types=(("MIDI files", "*.xlsx"),))],

    [sg.Text('Location', size=(17, 1)), sg.InputText('', key='Location')],

    [sg.Text('User Name', size=(17, 1)), sg.InputText('', key='User Name')],

    [sg.Text('PC/Laptop-Model', size=(17, 1)), sg.Combo(['Pc', 'Laptop', 'AllInOne'], key='Pc/Laptop-Model')],

    [sg.Text('OS Type',size=(17,1)), sg.InputText(os_name,key='Os Name')],

    [sg.Text('PC Name',size=(17,1)), sg.InputText(sys.node,key='PC Name')],

    [sg.Text('CPU Model',size=(17,1)), sg.InputText(format(proc_info.Name),key='CPU Model')],
    [sg.Text('CPU Date',size=(17,1)), sg.InputText('',key='CPU Date')],

    [sg.Text('Ram (GB)',size=(17,1)), sg.InputText(format(int(round(float(new_system_ram)))),key='Ram (GB)')],
    [sg.Text('Ram Tech.',size=(17,1)), sg.InputText('DDR' + str(DDR_out),key='Ram Tech.')],
    [sg.Text('Ram Slots',size=(17,1)), sg.InputText(str(memory_slots) +'/'+output[-1],key='Ram Slots')],


    [sg.Text('SSD (GB)',size=(17,1)), sg.InputText(ssdGB,key='SSD (GB)')],
    [sg.Text('Free Space(GB)',size=(17,1)), sg.InputText(ssdFree,key='Free Space(GB)')],

    [sg.Text('HDD (GB)',size=(17,1)), sg.InputText(hddGB,key='HDD (GB)')],
    [sg.Text('HDD Free Space (GB)',size=(17,1)), sg.InputText(hddFree,key='HDD Free Space (GB)')],


    [sg.Text('GPU',size=(17,1)), sg.InputText(format(gpu_info.Name),key='GPU')],

    [sg.Text('Description',size=(17,1)), sg.InputText('',key='Description')],

    [sg.Save(size=(15,1)),sg.Exit(size=(15,1))]
          ]


window = sg.Window('Hello World!',layout)

oneTimeIf = int(0)

while True:
    event,values = window.read()
    try:
        if oneTimeIf == 0:
            FilePath = values["-FilePath-"]
            EXCEL_FILE = FilePath
            df = pd.read_excel(EXCEL_FILE)
            oneTimeIf = 1

        if event == sg.WIN_CLOSED or event == 'Exit':
            break

        if event == 'Save':
            df = df._append(values, ignore_index=True)
            df.to_excel(EXCEL_FILE, index=False)
            sg.popup('Data Saved')
            time.sleep(1.5)
            os.startfile(EXCEL_FILE)
            break

    except:
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        sg.popup('Error')
        break


window.close()
# -------------------