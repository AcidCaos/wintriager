import sys
import os
import datetime
import dateutil.parser
import xlsxwriter
import re
import csv

EVIDENCES = None
MEMORY = None
REPORTS = None
SHADOWCOPIES = None

HOSTNAME = None

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
            line = "\n" + line # so that we can have an option that sets the start-of-line: "\nVersion="
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

    # formats
    bold = workbook.add_format({'bold': True})
    date_format = workbook.add_format({'num_format': "dd/mm/yy hh:mm:ss", 'align': 'left'})
    unusual = workbook.add_format({'bold': True})
    suspicious = workbook.add_format({'bold': True})
    
    ws_general = workbook.add_worksheet("General")
    ws_general_i = 0

    ws_general.set_column(0, 0, 20)
    ws_general.set_column(1, 1, 70)

    # system_date_time.txt
    print ("[+] REPORTS: System Date Time...")
    system_date_time = read_file(path_options(REPORTS, ["system_date_time.txt"]))
    date = executed_command_output(system_date_time, "date /T")
    time = executed_command_output(system_date_time, "time /T")
    timezone = value_from_tag_options(system_date_time, ["Caption="])
    WINTRIAGE_DATETIME = dateutil.parser.parse(date + " " + time, dayfirst=True, yearfirst=False)
    print("    * WinTriage Executed at:", WINTRIAGE_DATETIME, timezone)
    ws_general.write(ws_general_i, 0, "WinTriage Execution"); ws_general.write(ws_general_i, 1, str(WINTRIAGE_DATETIME)); ws_general_i += 1
    ws_general.write(ws_general_i, 0, "WinTriage Timezone"); ws_general.write(ws_general_i, 1, timezone); ws_general_i += 1

    # system_info.txt
    print ("[+] REPORTS: System Info...")
    system_info = read_file(path_options(REPORTS, ["system_info.txt"]))

    systeminfo = executed_command_output(system_info, "systeminfo")
    hostname = value_from_tag_options(systeminfo, ["Host Name:", "Nombre de host:"])
    winos = value_from_tag_options(systeminfo, ["OS Name:", "Nombre del sistema operativo:"])
    winver = value_from_tag_options(systeminfo, ["OS Version:", "n del sistema operativo:"])
    domain = value_from_tag_options(systeminfo, ["Domain:", "Dominio:"])
    manufacturer = value_from_tag_options(systeminfo, ["System Manufacturer:", "Fabricante del sistema:"])

    wmic_os = executed_command_output(system_info, "wmic os get Version, Caption, CountryCode, CSName, Description, InstallDate, SerialNumber, ServicePackMajorVersion, WindowsDirectory /format:list")
    if not hostname: hostname = value_from_tag_options(wmic_os, ["CSName="])
    if not winos: winos = value_from_tag_options(wmic_os, ["Caption="])
    if not winver: winver = value_from_tag_options(wmic_os, ["\nVersion="]) + " Service Pack " + value_from_tag_options(wmic_os, ["ServicePackMajorVersion="])

    if hostname: print("    * Hostname:", hostname)
    if domain: print("    * Domain:", domain)
    if winos: print("    * OS:", winos, winver)

    global HOSTNAME
    if hostname: HOSTNAME = hostname

    ws_general.write(ws_general_i, 0, "Hostname"); ws_general.write(ws_general_i, 1, hostname); ws_general_i += 1
    ws_general.write(ws_general_i, 0, "Domain"); ws_general.write(ws_general_i, 1, domain); ws_general_i += 1
    ws_general.write(ws_general_i, 0, "OS"); ws_general.write(ws_general_i, 1, winos + " " + winver); ws_general_i += 1
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
    print ("[+] REPORTS: Users and Groups...")
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

    ws_users.autofilter(0, 0, 0, 5) 

    useraccount = executed_command_output(users, "wmic useraccount get caption, sid")

    for i, line in enumerate(useraccount.split("\n")):
        line = line.strip()
        if line:
            if i == 0: continue # ignore headers
            res = re.split("\s\s+", line)
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
    print ("[+] REPORTS: Network...")
    network = read_file(path_options(REPORTS, ["network.txt"]))
    ipconfig = executed_command_output(network, "ipconfig /all")
    ip = value_from_tag_options(ipconfig, ["Dirección IP", "Dirección IPv4", "Direcci¢n IPv4", "Direcci¢n IPv4", "IP Address", "IPv4 Address"])
    ip = ip.replace(". ", "").replace(":", "").strip()

    if ip: print("    * IP:", ip)
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

    ws_network.autofilter(0, 0, 0, 5)

    for block in netstat.strip().split("\n\n"):

        for line in block.split("\n"):

            res = re.split("\s\s+", line)
            res_proc = re.findall("\[.+\]", line)

            if res and len(res)>1 and res[1] == "Proto": # Ignore header
                continue

            if not res_proc and len(res) < 4 : # Title OR Can not obtain ownership information
                continue

            elif res_proc: # Process line
                ws_network.write(ws_network_i - 1, 5, ", ".join(res_proc).replace("[", "").replace("]", ""))

            elif len(res) == 6: # Line with Status
                ws_network.write(ws_network_i, 0, res[1])
                ws_network.write(ws_network_i, 1, res[2])
                ws_network.write(ws_network_i, 2, res[3])
                ws_network.write(ws_network_i, 3, res[4])
                try:
                    ws_network.write_number(ws_network_i, 4, int(res[5]))
                except:
                    ws_network.write(ws_network_i, 4, res[5])
                ws_network_i += 1

            elif len(res) == 5: # Line without Status
                ws_network.write(ws_network_i, 0, res[1])
                ws_network.write(ws_network_i, 1, res[2])
                ws_network.write(ws_network_i, 2, res[3])
                try:
                    ws_network.write_number(ws_network_i, 4, int(res[4]))
                except:
                    ws_network.write(ws_network_i, 4, res[4])
                ws_network_i += 1
            
    
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
    print ("[+] REPORTS: Processes...")
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

    ws_processes.autofilter(0, 0, 0, 8)

    with open(path_options(REPORTS, ["processes.csv"]), newline='') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
        for row in spamreader:
            if not row: # empty rows
                continue
            if ws_processes_i == 0: # avoid header
                ws_processes_i += 1
                continue
            for i, rv in enumerate(row):
                if i == 1 or i == 3:
                    try:
                        ws_processes.write_number(ws_processes_i, i, int(rv))
                        continue
                    except:
                        pass
                ws_processes.write(ws_processes_i, i, rv)
            ws_processes_i += 1

    # scheduled tasks
    print ("[+] REPORTS: Scheduled Tasks...")
    ws_scheduled = workbook.add_worksheet("Scheduled Tasks")
    ws_scheduled_i = 0

    ws_scheduled.set_column(1, 1, 25)
    ws_scheduled.set_column(2, 2, 16)
    ws_scheduled.set_column(5, 5, 16)
    ws_scheduled.set_column(8, 8, 40)

    headers = ["HostName","TaskName","Next Run Time","Status","Logon Mode","Last Run Time","Last Result","Creator","Schedule","Task To Run","Start In","Comment","Scheduled Task State","Scheduled Type","Start Time","Start Date","End Date","Days","Months","Run As User","Delete Task If Not Rescheduled","Stop Task If Runs X Hours and X Mins","Repeat: Every","Repeat: Until: Time","Repeat: Until: Duration","Repeat: Stop If Still Running","Idle Time","Power Management"]
    for i, h in enumerate(headers):
        ws_scheduled.write(ws_scheduled_i, i, h, bold)

    ws_scheduled.autofilter(0, 0, 0, len(headers)-1)

    with open(path_options(REPORTS, ["programmed_tasks.csv"]), newline='') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
        for row in spamreader:
            if not row: # empty rows
                continue
            if ws_scheduled_i == 0: # avoid header
                ws_scheduled_i += 1
                continue
            for i, rv in enumerate(row):
                if i == 2 or i == 5: # dates
                    try:
                        # date_time = datetime.datetime.strptime(rv, '%d/%m/%Y %H:%M:%S')
                        date_time = dateutil.parser.parse(rv, dayfirst=True, yearfirst=False)
                        ws_scheduled.write_datetime(ws_scheduled_i, i, date_time, date_format)
                        continue
                    except:
                        pass
                if i == 6:
                    try:
                        ws_scheduled.write_number(ws_scheduled_i, i, int(rv))
                        continue
                    except:
                        pass
                ws_scheduled.write(ws_scheduled_i, i, rv)
            ws_scheduled_i += 1

    # loaded_dlls
    print ("[+] REPORTS: Loaded DLLs...")
    ws_loaded_dlls = workbook.add_worksheet("Loaded DLLs")
    ws_loaded_dlls_i = 0

    ws_loaded_dlls.set_column(0, 0, 30)
    ws_loaded_dlls.set_column(1, 1, 6)
    ws_loaded_dlls.set_column(2, 2, 70)

    ws_loaded_dlls.write(ws_loaded_dlls_i, 0, "Image Name", bold)
    ws_loaded_dlls.write(ws_loaded_dlls_i, 1, "PID", bold)
    ws_loaded_dlls.write(ws_loaded_dlls_i, 2, "Loaded DLLs", bold)

    ws_loaded_dlls.autofilter(0, 0, 0, 2)

    with open(path_options(REPORTS, ["loaded_dlls.csv"]), newline='') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
        for row in spamreader:
            if not row: # empty rows
                continue
            if ws_loaded_dlls_i == 0: # avoid header
                ws_loaded_dlls_i += 1
                continue
            for i, rv in enumerate(row):
                if i == 1:
                    try:
                        ws_loaded_dlls.write_number(ws_loaded_dlls_i, i, int(rv))
                        continue
                    except:
                        pass
                ws_loaded_dlls.write(ws_loaded_dlls_i, i, rv)
            ws_loaded_dlls_i += 1

    # complete_file_listing_x
    print ("[+] REPORTS: Drives Files...")
    for file in os.listdir(REPORTS):

        if file.startswith("complete_file_listing_"):
            drive = file.split(".")[0][-1].upper()

            ws_drive = workbook.add_worksheet("Drive " + drive + " Files")
            ws_drive_i = 0

            headers = ["Date", "Type", "Size (bytes)", "File"]
            for i, h in enumerate(headers):
                ws_drive.write(ws_drive_i, i, h, bold)
            ws_drive_i += 1

            ws_drive.autofilter(0, 0, 0, len(headers)-1)

            ws_drive.set_column(0, 0, 16)
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
            dir_blocks = data.split(head)
            for i, dir_block in enumerate(dir_blocks):
                if i == len(dir_blocks)-1: # If last block, remove the "Total Files Listed:" section
                    dir_block = "\n\n".join(dir_block.strip().split("\n\n")[0:-1]) # remove last section
                    # print("LAST BLOCK:", i, "--->\n", dir_block)
                dir_block = dir_block.strip().split("\n")
                dir_root = dir_block[0].strip()
                rows = dir_block[2:-1]
                for row in rows:
                    res = re.split("\s+", row)
                    if len(res) < 4:
                        continue
                    # Check AM/PM date format
                    date_offset = 0
                    if res[2].strip().upper() in ["AM", "PM"]:
                        date_offset = 1
                        res[1] += " " + res[2]
                    if res[3+date_offset] in [".", ".."]:
                        continue
                    try:
                        # date_time = datetime.datetime.strptime(res[0] + " " + res[1], '%d/%m/%Y %H:%M')
                        date_time = dateutil.parser.parse(res[0] + " " + res[1], dayfirst=True, yearfirst=False)
                        ws_drive.write_datetime(ws_drive_i, 0, date_time, date_format)
                    except:
                        ws_drive.write(ws_drive_i, 0, res[0] + " " + res[1])
                    ws_drive.write(ws_drive_i, 1, "DIR" if res[2+date_offset] == "<DIR>" else "FILE")
                    ws_drive.write(ws_drive_i, 2, "" if res[2+date_offset] == "<DIR>" else res[2+date_offset])
                    ws_drive.write(ws_drive_i, 3, dir_root + "\\" + res[3+date_offset])
                    ws_drive_i += 1
    
    print ("[+] REPORTS: Done!")
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
    
    print ("[+] Saving Extract...")
    try:
        if HOSTNAME:
            workbook.filename = ROOT + '\\Extract_' + HOSTNAME + '.xlsx'
        workbook.close()
    except xlsxwriter.workbook.FileCreateError:
        print ("[!] Could not create Extract file... It is likely being used.")
    except Exception as e:
        print ("[!] Could not save Extract. Error: ", str(e))
    print("[+] Done!")