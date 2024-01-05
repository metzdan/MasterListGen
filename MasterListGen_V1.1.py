import PySimpleGUI as sg
import os
import shutil
import csv
import pandas as pd


sg.theme("SystemDefault")

layout = [
    [sg.T("")],
    [sg.Text("Blancco Report.zip:   "),
     sg.Input(key="-IN-"),
     sg.FileBrowse(file_types=(("Zip Files", "*.zip"), ("ALL Files", "*.*"), ))],
    [sg.Text("Serial Numbers.xlsx: "),
     sg.Input(key="-IN1-"),
     sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"), ("ALL Files", "*.*"), ))],
     [sg.T("")],
    [sg.Button("Submit")],
    [sg.T("")],
    [sg.Text("POSRG Data Destruction Lab 01-02-24")],
    [sg.Text("Created by Dan Metzler. Tested by Jack Turner.")]
]

window = sg.Window('Master Report Generator', layout, size=(600,230))

filename = ""
while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, "Exit"):
        break
    elif event == "Submit":
        blistin = values['-IN-']
        break
plistin = values['-IN1-']

window.close()
print(filename)

dir = os.path.dirname(os.path.abspath(__file__))

output = os.path.join(dir, r'Output')
if not os.path.exists(output):
    os.makedirs(output)


#Variables
#blistin = ["IN"]#"/home/dmetzler/Desktop/Programs/Input/reports (14).zip"
blist1 = os.path.join(output, r"reports.csv")
#output="/home/dmetzler/Desktop/Programs/Output"
#plistin="/home/dmetzler/Desktop/Programs/Input/Serial Numbers.xlsx"


#Extracts Blancco report and puts it in the output file
shutil.unpack_archive(filename=values["-IN-"], extract_dir=output)
#Places Brackets around blist csv
results = []
with open(blist1) as csvfile:
    reader = csv.reader(csvfile, quoting=csv.QUOTE_NONNUMERIC)
    for row in reader:
        results.append(row)

#This Reads and filters the P-List.xlsx and grabs only the serial number column
pl = pd.read_excel(plistin, sheet_name=0)
#pl1 = pl[['Serial']]
pl.to_excel(os.path.join(output, r'Filtered_P-List_1.xlsx'), index=False)

rp = pd.read_csv(os.path.join(output, r'reports.csv'), index_col=False)
rp.to_excel(os.path.join(output, r'output.xlsx'))

#########################

blistcomp=pd.read_excel(os.path.join(output, r'output.xlsx'))


plistcomp = pd.read_excel(os.path.join(output, r'Filtered_P-List_1.xlsx'))

pd.merge(blistcomp, plistcomp, on='blancco_hardware_report.system.serial', how='inner')

outer_common=pd.merge(blistcomp, plistcomp, on='blancco_hardware_report.system.serial', how='outer', indicator=True)
outer_common_nc=outer_common[outer_common['_merge']!="both"]
outer_common_nc.iloc[:,0:3]



outer_common.to_excel(os.path.join(output, r'output.xlsx'))
###########################
op = pd.read_excel(os.path.join(output, r'output.xlsx'))
op = op[op['_merge'].str.contains('both')]
op.to_excel(os.path.join(output, r'outputfinal.xlsx'))
#### op1 is the list created for any serial numbers found only on the profiled list and not the blancco list. Anything that shows up on this list means a mistake happend
op1 = pd.read_excel(os.path.join(output,r'output.xlsx'))
op1 = op1[op1['_merge'].str.contains('right_only')]
op1.to_excel(os.path.join(output, r'Missing.xlsx'))

#########################Creating Master List and Manifest
masterlist = pd.read_excel(os.path.join(output, r'outputfinal.xlsx'), usecols=['report.report_date', 'blancco_hardware_report.system.manufacturer', 'blancco_hardware_report.system.model', 'blancco_hardware_report.system.serial', 'blancco_hardware_report.disks.disk.capacity', 'blancco_hardware_report.processors.processor.model', 'blancco_hardware_report.memory.total_memory', 'blancco_hardware_report.memory.memory_bank.type', 'blancco_hardware_report.memory.memory_bank.hz', 'blancco_hardware_report.disks.disk.model', 'blancco_hardware_report.disks.disk.interface_type', 'blancco_hardware_report.optical_drives.optical_drive.model', 'blancco_hardware_report.video_cards.video_card.model', 'user_data.fields.R2 Cosmetic', 'user_data.fields.R2 Functionality'])
masterlist.to_excel(os.path.join(output, r'Masterlist.xlsx'))
manifest = pd.read_excel(os.path.join(output, r'Masterlist.xlsx'), usecols=['report.report_date', 'blancco_hardware_report.system.manufacturer', 'blancco_hardware_report.system.model', 'blancco_hardware_report.system.serial'])
manifest.to_excel(os.path.join(output, r'Manifest.xlsx'))

#Unused File removal
os.remove(os.path.join(output, r'Filtered_P-List_1.xlsx'))
os.remove(os.path.join(output, r'output.xlsx'))
os.remove(os.path.join(output, r'outputfinal.xlsx'))
os.remove(os.path.join(output, r'reports.csv'))

print('Success!')