import sys
import os
import datetime
import xlsxwriter
import re
import csv

EVIDENCES = None
MEMORY = None
REPORTS = None
SHADOWCOPIES = None

def read_file(file) -> list:
    if not file:
        return None
    f = open(file, encoding='cp1252', errors='replace')
    c = f.read()
    f.close()
    return c

def read_lines(file) -> list:
    if not file:
        return None
    f = open(file, 'r')
    l = f.readlines()
    f.close()
    return l

def path_options(root, options, none=None):
    if type(options) == str:
        options = [options]
    for opt in options:
        if os.path.exists(root + "\\" + opt):
            return root + "\\" + opt
    return None

def value_from_tag_options(lines, options):
    if type(lines) == str:
        lines = lines.split("\n")
    if type(options) == str:
        options = [options]
    for opt in options:
        for line in lines:
            try:
                return line[line.index(opt)+len(opt):].strip()
            except ValueError:
                continue
    return ""

def values_from_table(lines, column_separator):
    if type(lines) == str:
        lines = lines.split("\n")
    table = []
    for line in lines:
        if column_separator in line:
            row = [val.strip() for val in line.split(column_separator)]
            table.append(row)
        else:
            table.append([line.strip()]) # table[-1][-1] += "\n" + line.strip()
    return table

def executed_command_output(data, command):
    exec_comm = ["EXECUTED COMMAND:", "COMANDO EJECUTADO:"]
    head = None
    # find commadn header
    for e_c in exec_comm:
        if e_c in data:
            head = e_c
            break
    # get command output
    outputs = data.split(head)
    for out in outputs:
        try:
            if out and command == out.split("\n")[0].strip():
                return out[out.index(command)+len(command):].strip()
        except ValueError:
            continue
    return ""

def reports(workbook):

    bold = workbook.add_format({'bold': True})
    
    ws_general = workbook.add_worksheet("General")
    ws_general_i = 0

    ws_general.set_column(0, 0, 20)
    ws_general.set_column(1, 1, 70)

    print ("[+] Analyze Reports...")

    # system_date_time.txt
    system_date_time = read_file(path_options(REPORTS, ["system_date_time.txt"]))
    date = executed_command_output(system_date_time, "date /T")
    time = executed_command_output(system_date_time, "time /T")
    timezone = value_from_tag_options(system_date_time, ["Caption="])
    WINTRIAGE_DATETIME = datetime_object = datetime.datetime.strptime(date + " " + time, '%d/%m/%Y %H:%M')
    print("[i] WinTriage Executed at:", WINTRIAGE_DATETIME, timezone)
    ws_general.write(ws_general_i, 0, "WinTriage Execution"); ws_general.write(ws_general_i, 1, str(WINTRIAGE_DATETIME)); ws_general_i += 1
    ws_general.write(ws_general_i, 0, "WinTriage Timezone"); ws_general.write(ws_general_i, 1, timezone); ws_general_i += 1

    # system_info.txt
    system_info = read_file(path_options(REPORTS, ["system_info.txt"]))
    systeminfo = executed_command_output(system_info, "systeminfo")
    hostname = value_from_tag_options(systeminfo, ["Host Name:", "Nombre de host:"])
    winos = value_from_tag_options(systeminfo, ["OS Name:", "Nombre del sistema operativo:"])
    domain = value_from_tag_options(systeminfo, ["Domain:", "Dominio:"])
    manufacturer = value_from_tag_options(systeminfo, ["System Manufacturer:", "Fabricante del sistema:"])
    
    ws_general.write(ws_general_i, 0, "Hostname"); ws_general.write(ws_general_i, 1, hostname); ws_general_i += 1
    ws_general.write(ws_general_i, 0, "Domain"); ws_general.write(ws_general_i, 1, domain); ws_general_i += 1
    ws_general.write(ws_general_i, 0, "OS"); ws_general.write(ws_general_i, 1, winos); ws_general_i += 1
    ws_general.write(ws_general_i, 0, "Manufacturer"); ws_general.write(ws_general_i, 1, manufacturer); ws_general_i += 1
    
    ws_systeminfo = workbook.add_worksheet("System Info")
    ws_systeminfo_i = 0

    ws_systeminfo.set_column(0, 0, 35)
    ws_systeminfo.set_column(1, 1, 60)
    ws_systeminfo.set_column(2, 2, 12)

    properties = values_from_table(systeminfo, ": ")
    for prop in properties:
        # print (prop)
        for i, val in enumerate(prop):
            ws_systeminfo.write(ws_systeminfo_i, i, val)
        ws_systeminfo_i += 1

    # Usuarios.txt Users.txt

    users = read_file(path_options(REPORTS, ["Usuarios.txt", "Users.txt"]))

    ws_users = workbook.add_worksheet("Users & Groups")
    ws_users_i = 0

    ws_users.set_column(0, 0, 12)
    ws_users.set_column(1, 1, 15)
    ws_users.set_column(2, 2, 50)
    ws_users.set_column(3, 3, 45)
    ws_users.set_column(4, 4, 10)
    ws_users.set_column(5, 5, 7)

    ws_users.write(ws_users_i, 0, "Local Account", bold)
    ws_users.write(ws_users_i, 1, "Domain", bold)
    ws_users.write(ws_users_i, 2, "Name", bold)
    ws_users.write(ws_users_i, 3, "SID", bold)
    ws_users.write(ws_users_i, 4, "InstallDate", bold)
    ws_users.write(ws_users_i, 5, "Status", bold)
    ws_users_i += 1

    useraccount = executed_command_output(users, "wmic useraccount get caption, sid")

    for i, line in enumerate(useraccount.split("\n")):
        line = line.strip()
        if line:
            if i == 0: continue # ignore headers
            res = re.split("\s+", line)
            if len(res) != 2:
                continue
            ws_users.write(ws_users_i, 1, res[0].split("\\")[0])
            ws_users.write(ws_users_i, 2, res[0].split("\\")[1])
            ws_users.write(ws_users_i, 3, res[1])
            ws_users_i += 1

    group = executed_command_output(users, "wmic group get Caption, InstallDate, LocalAccount, Domain, SID, Status")

    for i, line in enumerate(group.split("\n")):
        line = line.strip()
        if line:
            if i == 0: continue # ignore headers
            res = re.split("\s\s+", line)
            offset = 0
            if len(res) < 5:
                continue
            if len(res) > 5:
                offset = 1
            ws_users.write(ws_users_i, 0, res[2+offset])
            ws_users.write(ws_users_i, 1, res[1])
            ws_users.write(ws_users_i, 2, res[0].split("\\")[1])
            ws_users.write(ws_users_i, 3, res[3+offset])
            ws_users.write(ws_users_i, 4, res[2] if offset else "")
            ws_users.write(ws_users_i, 5, res[4+offset])
            ws_users_i += 1

    # network.txt
    network = read_file(path_options(REPORTS, ["network.txt"]))
    ipconfig = executed_command_output(network, "ipconfig /all")
    ip = value_from_tag_options(ipconfig, ["Dirección IP. . . . . . . . . :", "IP Address"])
    ws_general.write(ws_general_i, 0, "IP"); ws_general.write(ws_general_i, 1, ip); ws_general_i += 1

    netstat = executed_command_output(network, "netstat -noab")

    ws_network = workbook.add_worksheet("Active Connections")
    ws_network_i = 0

    ws_network.set_column(0, 0, 8)
    ws_network.set_column(1, 1, 18)
    ws_network.set_column(2, 2, 18)
    ws_network.set_column(3, 3, 13)
    ws_network.set_column(4, 4, 8)
    ws_network.set_column(5, 5, 25)

    ws_network.write(ws_network_i, 0, "Protocol", bold)
    ws_network.write(ws_network_i, 1, "Local Adress", bold)
    ws_network.write(ws_network_i, 2, "Remote Address", bold)
    ws_network.write(ws_network_i, 3, "Status", bold)
    ws_network.write(ws_network_i, 4, "PID", bold)
    ws_network.write(ws_network_i, 5, "Process", bold)
    ws_network_i += 1

    for block in netstat.strip().split("\n\n"):
        if ws_network_i == 0: # ignore header
            pass
        # print (block)
        for line in block.split("\n"):
            res = re.split("\s+", line)
            if len(res) == 6:
                ws_network.write(ws_network_i, 0, res[1])
                ws_network.write(ws_network_i, 1, res[2])
                ws_network.write(ws_network_i, 2, res[3])
                ws_network.write(ws_network_i, 3, res[4])
                ws_network.write(ws_network_i, 4, res[5])
                ws_network_i += 1
            elif len(res) == 5:
                ws_network.write(ws_network_i, 0, res[1])
                ws_network.write(ws_network_i, 1, res[2])
                ws_network.write(ws_network_i, 2, res[3])
                ws_network.write(ws_network_i, 4, res[4])
                ws_network_i += 1
            else:
                res = re.findall("\[.+\]", line)
                ws_network.write(ws_network_i - 1, 5, ", ".join(res).replace("[", "").replace("]", ""))
    
    # etc hosts

    ws_hosts = workbook.add_worksheet("Network Hosts")
    ws_hosts_i = 0

    ws_hosts.write(ws_hosts_i, 0, "IP", bold)
    ws_hosts.write(ws_hosts_i, 1, "Hostname", bold)
    ws_hosts_i += 1

    ws_hosts.set_column(0, 0, 18)
    ws_hosts.set_column(1, 1, 35)

    hosts = executed_command_output(network, "type C:\WINDOWS\system32\drivers\etc\hosts")
    for line in hosts.split("\n"):
        if line.strip() and line.strip()[0] != '#':
            res = re.split("\s+", line)
            # print (res)
            ws_hosts.write(ws_hosts_i, 0, res[0])
            ws_hosts.write(ws_hosts_i, 1, res[1])
            ws_hosts_i += 1

    # processes.csv

    ws_processes = workbook.add_worksheet("Processes")
    ws_processes_i = 0

    ws_processes.set_column(0, 0, 30)
    ws_processes.set_column(1, 1, 6)
    ws_processes.set_column(2, 2, 15)
    ws_processes.set_column(3, 3, 15)
    ws_processes.set_column(4, 4, 15)
    ws_processes.set_column(5, 5, 12)
    ws_processes.set_column(6, 6, 30)
    ws_processes.set_column(7, 7, 14)
    ws_processes.set_column(8, 8, 45)

    ws_processes.write(ws_processes_i, 0, "Image Name", bold)
    ws_processes.write(ws_processes_i, 1, "PID", bold)
    ws_processes.write(ws_processes_i, 2, "Session Name", bold)
    ws_processes.write(ws_processes_i, 3, "Session#", bold)
    ws_processes.write(ws_processes_i, 4, "Mem Usage", bold)
    ws_processes.write(ws_processes_i, 5, "Status", bold)
    ws_processes.write(ws_processes_i, 6, "User Name", bold)
    ws_processes.write(ws_processes_i, 7, "CPU Time", bold)
    ws_processes.write(ws_processes_i, 8, "Window Title", bold)

    with open(path_options(REPORTS, ["processes.csv"]), newline='') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
        for row in spamreader:
            if ws_processes_i == 0: ws_processes_i += 1; continue # avoid header
            for i, rv in enumerate(row):
                ws_processes.write(ws_processes_i, i, rv)
            ws_processes_i += 1

    # scheduled tasks

    ws_scheduled = workbook.add_worksheet("Scheduled Tasks")
    ws_scheduled_i = 0

    ws_scheduled.set_column(2, 2, 18)
    ws_scheduled.set_column(7, 7, 15)
    ws_scheduled.set_column(9, 9, 40)

    headers = ["HostName","TaskName","Next Run Time","Status","Logon Mode","Last Run Time","Last Result","Creator","Schedule","Task To Run","Start In","Comment","Scheduled Task State","Scheduled Type","Start Time","Start Date","End Date","Days","Months","Run As User","Delete Task If Not Rescheduled","Stop Task If Runs X Hours and X Mins","Repeat: Every","Repeat: Until: Time","Repeat: Until: Duration","Repeat: Stop If Still Running","Idle Time","Power Management"]
    for i, h in enumerate(headers):
        ws_scheduled.write(ws_scheduled_i, i, h, bold)

    with open(path_options(REPORTS, ["programmed_tasks.csv"]), newline='') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
        for row in spamreader:
            if ws_scheduled_i == 0: ws_scheduled_i += 1; continue # avoid header
            for i, rv in enumerate(row):
                ws_scheduled.write(ws_scheduled_i, i, rv)
            ws_scheduled_i += 1

    # loaded_dlls

    ws_loaded_dlls = workbook.add_worksheet("Loaded DLLs")
    ws_loaded_dlls_i = 0

    ws_loaded_dlls.set_column(0, 0, 30)
    ws_loaded_dlls.set_column(1, 1, 6)
    ws_loaded_dlls.set_column(2, 2, 70)

    ws_loaded_dlls.write(ws_loaded_dlls_i, 0, "Image Name", bold)
    ws_loaded_dlls.write(ws_loaded_dlls_i, 1, "PID", bold)
    ws_loaded_dlls.write(ws_loaded_dlls_i, 2, "Loaded DLLs", bold)

    with open(path_options(REPORTS, ["loaded_dlls.csv"]), newline='') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
        for row in spamreader:
            if ws_loaded_dlls_i == 0: ws_loaded_dlls_i += 1; continue # avoid header
            for i, rv in enumerate(row):
                ws_loaded_dlls.write(ws_loaded_dlls_i, i, rv)
            ws_loaded_dlls_i += 1

    # complete_file_listing_x

    for file in os.listdir(REPORTS):

        if file.startswith("complete_file_listing_"):
            drive = file.split(".")[0][-1].upper()

            ws_drive = workbook.add_worksheet("Drive " + drive + " Files")
            ws_drive_i = 0

            headers = ["Date", "Type", "Size (bytes)", "File"]
            for i, h in enumerate(headers):
                ws_drive.write(ws_drive_i, i, h, bold)
            ws_drive_i += 1

            ws_drive.set_column(0, 0, 15)
            ws_drive.set_column(1, 1, 5)
            ws_drive.set_column(2, 2, 13)
            ws_drive.set_column(3, 3, 90)
            
            data = read_file(REPORTS + "\\" + file)

            dir_of = ["Directorio de", "Directory of"]
            head = None
            # find directory header
            for d_o in dir_of:
                if d_o in data:
                    head = d_o
                    break
            # get directory blocks
            for dir_block in data.split(head):
                dir_block = dir_block.strip().split("\n")
                dir_root = dir_block[0].strip()
                rows = dir_block[2:-1]
                for row in rows:
                    res = re.split("\s+", row)
                    if len(res) < 3:
                        continue
                    if res[3] in [".", ".."]:
                        continue
                    ws_drive.write(ws_drive_i, 0, res[0] + " " + res[1])
                    ws_drive.write(ws_drive_i, 1, "DIR" if res[2] == "<DIR>" else "FILE")
                    ws_drive.write(ws_drive_i, 2, "" if res[2] == "<DIR>" else res[2])
                    ws_drive.write(ws_drive_i, 3, dir_root + "\\" + res[3])
                    ws_drive_i += 1


    return

if __name__ == "__main__":

    # Initial Checks

    if len(sys.argv) != 2:
        print ("Usage:", sys.argv[0], "D:\\PATH\\TO\\WinTriage\n")
        exit(1)

    ROOT = sys.argv[1]
    print ("[+] WinTriage root:", ROOT)


    # Check for files

    EVIDENCES = path_options(ROOT, ["Evidences", "Evidencias"])
    MEMORY = path_options(ROOT, ["Memory", "Memoria"])
    REPORTS = path_options(ROOT, ["Reports", "Informes"])
    SHADOWCOPIES = path_options(ROOT, ["ShadowCopies"])

    if not EVIDENCES and not MEMORY and not REPORTS and not SHADOWCOPIES:
        print ("Error: WinTriage not found in the given path.\n")
        exit(1)

    if EVIDENCES:  print ("    * Found Evidences!")
    if MEMORY: print ("    * Found Memory!")
    if REPORTS: print ("    * Found Reports!")
    if SHADOWCOPIES: print ("    * Found Shadow Copies!")


    # Spreadsheet

    workbook = xlsxwriter.Workbook(ROOT + '\Extract.xlsx')

    # Run modules

    if EVIDENCES:
        pass

    if MEMORY:
        pass

    if REPORTS:
        reports(workbook)
    
    if SHADOWCOPIES:
        pass

    workbook.close()

    print("[+] Done!")