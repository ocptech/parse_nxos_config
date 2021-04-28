import xlwt
from xlwt import Workbook
from ciscoconfparse import CiscoConfParse
from pathlib import Path
import re
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

path = filedialog.askopenfilename()

print ("File: {}".format(path))
# NXOS running config file
try:
    parse = CiscoConfParse(path, syntax='nxos')
except UnicodeDecodeError as uni:
    print ("USE A TEXT FILE!!! Closing...")
    exit()

wb = Workbook()

# VLANs information page
sheet1 = wb.add_sheet('VLANs')
row1 = 0
sheet1.write(row1, 0, "VLAN")
sheet1.write(row1, 1, "NAME")
for p_obj in parse.find_objects('^vlan')[1:]:
    row1 += 1
    sheet1.write(row1, 0, str(p_obj)[str(p_obj).find("vlan") + 4:-2])
    for c_obj in p_obj.children:
        sheet1.write(
            row1,
            1,
            str(c_obj)[
                str(c_obj).find("name") +
                5:str(c_obj).find("' (parent")])

# VLANs Interfaces information page
sheet2 = wb.add_sheet('SVIs')
row1 = 0
sheet2.write(row1, 0, "SVI")
sheet2.write(row1, 1, "Description")
sheet2.write(row1, 2, "IP")
sheet2.write(row1, 3, "VIP")
sheet2.write(row1, 4, "VRF")
for p_obj in parse.find_objects('^interface Vlan')[1:]:
    row1 += 1
    sheet2.write(
        row1,
        0,
        str(p_obj)[
            str(p_obj).find("interface Vlan") +
            14:-
            2])
    for c_obj in p_obj.children:
        if "description" in str(c_obj):
            sheet2.write(
                row1,
                1,
                str(c_obj)[
                    str(c_obj).find("description") +
                    12:str(c_obj).find("' (parent")])
        if "ip address" in str(c_obj):
            sheet2.write(
                row1,
                2,
                str(c_obj)[
                    str(c_obj).find("ip address") +
                    11:str(c_obj).find("' (parent")])
        if "hsrp" in str(c_obj):
            for h_obj in c_obj.children:
                if "ip" in str(h_obj):
                    sheet2.write(
                        row1,
                        3,
                        str(h_obj)[
                            str(h_obj).find("ip") +
                            3:str(h_obj).find("' (parent")])
        if "vrf member" in str(c_obj):
            sheet2.write(
                row1,
                4,
                str(c_obj)[
                    str(c_obj).find("vrf member") +
                    11:str(c_obj).find("' (parent")])

# Interfaces information page
sheet3 = wb.add_sheet('Ints')
row1 = 0
sheet3.write(row1, 0, "Interface")
sheet3.write(row1, 1, "Description")
sheet3.write(row1, 2, "Type")
sheet3.write(row1, 3, "VLANs/IP")
sheet3.write(row1, 4, "Po")
sheet3.write(row1, 5, "Status")
sheet3.write(row1, 6, "VRF")
for p_obj in parse.find_objects('^interface Ethernet')[1:]:
    row1 += 1
    sheet3.write(row1, 0, str(p_obj)[str(p_obj).find("interface") + 10:-2])
    for c_obj in p_obj.children:
        if "description" in str(c_obj):
            sheet3.write(
                row1,
                1,
                str(c_obj)[
                    str(c_obj).find("description") +
                    12:str(c_obj).find("' (parent")])
        try:
            if "access vlan" in str(c_obj):
                sheet3.write(
                    row1,
                    3,
                    str(c_obj)[
                        str(c_obj).find("access vlan") +
                        12:str(c_obj).find("' (parent")])
                sheet3.write(row1, 2, "swichport")
            if "trunk allowed vlan" in str(c_obj):
                sheet3.write(
                    row1,
                    3,
                    str(c_obj)[
                        str(c_obj).find("trunk allowed vlan") +
                        19:str(c_obj).find("' (parent")])
                sheet3.write(row1, 2, "swichport")
            if "ip address" in str(c_obj):
                sheet3.write(
                    row1,
                    3,
                    str(c_obj)[
                        str(c_obj).find("ip address") +
                        11:str(c_obj).find("' (parent")])
                sheet3.write(row1, 2, "routed")
        except Exception as ex:
            print("Problem with line ", str(p_obj), ex)
            sheet3.write(row1, 7, "incoherence in col: {}".format(
                int(str(ex)[str(ex).find("colx=") + 5:]) + 1))
        if "channel-group" in str(c_obj):
            sheet3.write(
                row1,
                4,
                str(c_obj)[
                    str(c_obj).find("channel-group") +
                    14:str(c_obj).find("' (parent")])
        if "no shutdown" in str(c_obj):
            sheet3.write(row1, 5, "no shutdown")
        if "'  shutdown' (" in str(c_obj):
            sheet3.write(row1, 5, "no shutdown")
        if "vrf member" in str(c_obj):
            sheet3.write(
                row1,
                6,
                str(c_obj)[
                    str(c_obj).find("vrf member") +
                    11:str(c_obj).find("' (parent")])

# Port-channels information page
sheet4 = wb.add_sheet('Po')
row1 = 0
sheet4.write(row1, 0, "Interface")
sheet4.write(row1, 1, "Description")
sheet4.write(row1, 2, "Type")
sheet4.write(row1, 3, "VLANs/IP")
sheet4.write(row1, 4, "Status")
sheet4.write(row1, 5, "VRF")
for p_obj in parse.find_objects('^interface port-channel')[1:]:
    row1 += 1
    sheet4.write(row1, 0, str(p_obj)[str(p_obj).find("interface") + 10:-2])
    for c_obj in p_obj.children:
        if "description" in str(c_obj):
            sheet4.write(
                row1,
                1,
                str(c_obj)[
                    str(c_obj).find("description") +
                    12:str(c_obj).find("' (parent")])
        try:
            if "access vlan" in str(c_obj):
                sheet4.write(
                    row1,
                    3,
                    str(c_obj)[
                        str(c_obj).find("access vlan") +
                        12:str(c_obj).find("' (parent")])
                sheet4.write(row1, 2, "swichport")
            if "trunk allowed vlan" in str(c_obj):
                sheet4.write(
                    row1,
                    3,
                    str(c_obj)[
                        str(c_obj).find("trunk allowed vlan") +
                        19:str(c_obj).find("' (parent")])
                sheet4.write(row1, 2, "swichport")
            if "ip address" in str(c_obj):
                sheet4.write(
                    row1,
                    3,
                    str(c_obj)[
                        str(c_obj).find("ip address") +
                        11:str(c_obj).find("' (parent")])
                sheet4.write(row1, 2, "routed")
        except Exception as ex:
            print("Problem with line ", str(p_obj), ex)
            sheet4.write(row1, 7, "incoherence in col: {}".format(
                int(str(ex)[str(ex).find("colx=") + 5:]) + 1))
        if "no shutdown" in str(c_obj):
            sheet4.write(row1, 4, "no shutdown")
        if "'  shutdown' (" in str(c_obj):
            sheet4.write(row1, 4, "shutdown")
        if "vrf member" in str(c_obj):
            sheet4.write(
                row1,
                5,
                str(c_obj)[
                    str(c_obj).find("vrf member") +
                    11:str(c_obj).find("' (parent")])

# Save to file
wbname = path.split("/")[-1] + ".xls"
print("The file was generated:  {}".format(wbname))
wb.save(path + ".xls")
