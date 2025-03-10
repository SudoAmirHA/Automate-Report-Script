import openpyxl
import paramiko
import time


switch_ips = ['172.27.4.100' , '172.27.242.7', '172.26.8.5']

# Get username and password for login device
username = "test"
password = "pass"

desktop_path = "C:/Users/a.amiri/Desktop/"
file_name = "switch_output.xlsx"
excel_file = openpyxl.Workbook()
sheet = excel_file.active
sheet.append(["Switch Name", "Command", "Output"])

def ssh_and_execute(switch_ips, username, password, command):
    ssh_client = paramiko.SSHClient()
    ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        ssh_client.connect(hostname=switch_ips, username=username, password=password)
        time.sleep(3)
        stdin, stdout, stderr = ssh_client.exec_command(command)
        output = stdout.read().decode('utf-8')
        sheet.append([switch_ips, command, output])
    except paramiko.AuthenticationException:
        print(f"Authentication failed for {switch_ips}")
    except paramiko.SSHException as e:
        print(f"SSH error for {switch_ips}: {e}")
    except Exception as e:
        print(f"Error connecting to {switch_ips}: {e}")
    finally:
        ssh_client.close()

for switch_ip in switch_ips:
    ssh_and_execute(switch_ip, username, password, "show ru | inc hostname")
    time.sleep(1)
    ssh_and_execute(switch_ip, username, password, "sho inte | inc CRC")
    time.sleep(1)
    ssh_and_execute(switch_ip, username, password, "sh int des | exc admin down")
    time.sleep(1)
    ssh_and_execute(switch_ip, username, password, "sho run | in trap-source")
    time.sleep(1)
    ssh_and_execute(switch_ip, username, password, "sho run | in tacacs-server host")
    time.sleep(1)
    ssh_and_execute(switch_ip, username, password, "sh env all")
    time.sleep(1)

excel_file.save(desktop_path + file_name)