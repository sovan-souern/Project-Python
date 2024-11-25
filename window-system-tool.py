import os
import psutil
import shutil
import speedtest
import subprocess
import platform
import socket
import os
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font
from openpyxl import load_workbook
import time

#setch system information

def get_wifi_profiles():
    """Get a list of saved Wi-Fi profiles."""
    try:
        profiles_data = subprocess.check_output("netsh wlan show profiles", shell=True, encoding="mbcs")
        return [line.split(":")[1].strip() for line in profiles_data.splitlines() if "All User Profile" in line]
    except subprocess.CalledProcessError:
        print("Error: Unable to retrieve Wi-Fi profiles.")
        return []

def get_wifi_password(profile):
    """Retrieve the password for a specific Wi-Fi profile."""
    try:
        profile_info = subprocess.check_output(f'netsh wlan show profile name="{profile}" key=clear', shell=True, encoding="mbcs")
        for line in profile_info.splitlines():
            if "Key Content" in line:
                return line.split(":")[1].strip()
        return None
    except subprocess.CalledProcessError:
        return None

def display_saved_wifi_credentials():
    """Display saved Wi-Fi networks and their passwords."""
    profiles = get_wifi_profiles()
    if profiles:
        print("\nSaved Wi-Fi Networks and Passwords:")
        print("-----------------------------------")
        for profile in profiles:
            password = get_wifi_password(profile)
            print(f"Network: {profile} | Password: {'None' if password is None else password}")
    else:
        print("No Wi-Fi profiles found.")

def get_serial_number():
    try:
        # Run the wmic command to get the serial number
        serial_number = subprocess.check_output("wmic bios get serialnumber").decode().strip().split('\n')[1]
        return serial_number
    except subprocess.CalledProcessError:
        return "Unable to retrieve serial number"


def get_system_info():
    # Get the system information
    system_info = platform.uname()

    return {
        "System Manufacturer": system_info.system,
        "Asset Tag": system_info.node,
        "Release": system_info.release,
        "Version": system_info.version,
        "System Model": system_info.machine,
        "Processor": system_info.processor
    }
def get_cpu_info():
    cpu_count = psutil.cpu_count(logical=True)  # Logical cores (threads)
    cpu_freq = psutil.cpu_freq()  # CPU frequency
    cpu_percent = psutil.cpu_percent(interval=1)  # CPU usage percent over a 1-second interval

    return f"""
CPU Count (Logical Cores): {cpu_count}
CPU Frequency: {cpu_freq.current} MHz (Current), {cpu_freq.min} MHz (Min), {cpu_freq.max} MHz (Max)
CPU Usage: {cpu_percent}%
    """

# Display CPU Information
print(get_cpu_info())
def get_memory_info():
    # Get memory information
    virtual_memory = psutil.virtual_memory()
    memory_info = {
        "Total Memory (GB)": virtual_memory.total / (1024 ** 3),
        "Used Memory (GB)": virtual_memory.used / (1024 ** 3),
        "Available Memory (GB)": virtual_memory.available / (1024 ** 3),
        "Memory Usage (%)": virtual_memory.percent
    }
    return memory_info

import socket


def get_local_ip():
    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)
    return local_ip

local_ip = get_local_ip()
print("Local IP Address:",local_ip)

def get_disk_info():
    # Get disk usage information
    disk_usage = psutil.disk_usage('/')
    disk_info = {
        "Total Disk Space (GB)": disk_usage.total / (1024 ** 3),
        "Used Disk Space (GB)": disk_usage.used / (1024 ** 3),
        "Free Disk Space (GB)": disk_usage.free / (1024 ** 3),
        "Disk Usage (%)": disk_usage.percent
    }
    return disk_info

# Example Usage:

# Get and print serial number
serial_number = get_serial_number()
print(f"Serial Number: {serial_number}")

# Get and print system information
system_info = get_system_info()
for key, value in system_info.items():
    print(f"{key}: {value}")

# Get and print memory information
memory_info = get_memory_info()
for key, value in memory_info.items():
    print(f"{key}: {value:.2f}")

# Get and print disk information
disk_info = get_disk_info()
for key, value in disk_info.items():
    print(f"{key}: {value:.2f}")


#username


username = os.getlogin()
print(f"Username: {username}")


# Run the wmic command to get the system's description (remark)
try:
    remark = subprocess.check_output("wmic computersystem get description", shell=True).decode().strip().split("\n")[1]
    print(f"System Remark: {remark}")
except subprocess.CalledProcessError:
    print("Failed to get system remark/description.")
# ------------------------------
# Cleanup Functions
# ------------------------------



# Function to get battery status
def get_battery_status():
    battery = psutil.sensors_battery()
    if battery:
        percent = battery.percent
        charging = battery.power_plugged
        return f"Battery: {percent}% {'(Charging)' if charging else '(Not Charging)'}"
    else:
        return "Battery information not available"



# Function to get the number of processes running
def get_process_count():
    process_count = len(psutil.pids())  # Get the list of process IDs
    return f"Number of processes: {process_count}"

# Function to get system uptime
def get_system_uptime():
    uptime_seconds = time.time() - psutil.boot_time()
    uptime_hours = uptime_seconds // 3600
    uptime_minutes = (uptime_seconds % 3600) // 60
    uptime = f"{int(uptime_hours)} hours and {int(uptime_minutes)} minutes"
    return f"System Uptime: {uptime}"

# Example usage:
battery_status = get_battery_status()
process_count = get_process_count()
system_uptime = get_system_uptime()

# Display the information
print(battery_status)
print(process_count)
print(system_uptime)




def get_wifi_passwords():
    # Run the command to list all Wi-Fi profiles
    profiles_cmd = "netsh wlan show profiles"
    profiles = subprocess.check_output(profiles_cmd, shell=True, text=True)

    # Extract the profile names
    profile_names = []
    for line in profiles.splitlines():
        if "All User Profile" in line:
            profile_names.append(line.split(":")[1].strip())

    # Retrieve passwords for each profile
    wifi_passwords = {}
    for profile in profile_names:
        # Run the command to show the Wi-Fi profile details
        password_cmd = f"netsh wlan show profile \"{profile}\" key=clear"
        profile_info = subprocess.check_output(password_cmd, shell=True, text=True)

        # Look for the password in the profile info
        for line in profile_info.splitlines():
            if "Key Content" in line:
                password = line.split(":")[1].strip()
                wifi_passwords[profile] = password
                break

    return wifi_passwords

# Get Wi-Fi passwords
wifi_passwords = get_wifi_passwords()
for profile, password in wifi_passwords.items():
    print("Profile: ",profile, "Password:" ,password)



#cleaing 



def clean_temp_files():
    """Clean temporary files from system temp directories."""
    # Get the temp directories
    temp_dirs = [os.environ.get('TEMP'), os.environ.get('TMP')]
    # Loop over temp directories and remove files
    for temp_dir in temp_dirs:
        if temp_dir and os.path.exists(temp_dir):
            for filename in os.listdir(temp_dir):
                file_path = os.path.join(temp_dir, filename)
                try:
                    if os.path.isfile(file_path):
                        os.remove(file_path) 
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)  
                except Exception as e:
                    print(f"Failed to remove {file_path}: {e}")

    print("Temporary files cleaned.")

def clean_recycle_bin():
    """Clear the Recycle Bin using PowerShell command (Windows only)."""
    try:
        # Run the PowerShell command to empty the Recycle Bin
        subprocess.run('PowerShell.exe -Command "Clear-RecycleBin -Force"', shell=True, check=True)
        print("Recycle Bin cleaned.")
    except subprocess.CalledProcessError as e:
        print(f"Failed to clear Recycle Bin: {e}")

def main():
    """Main function to run cleaning tasks."""
    print("Cleaning temporary files and Recycle Bin...")
    clean_temp_files()
    clean_recycle_bin()

# Corrected the if condition to check for '__main__'
if __name__ == "__main__":
    main()


def check_wifi_speed():
    """Check the download and upload speed of the active Wi-Fi connection."""
    st = speedtest.Speedtest()

    # List available servers and choose the first one
    servers = st.get_servers()
    server = list(servers.keys())[0] 

    st.get_best_server()

    # Measure download and upload speeds
    download_speed = st.download() / 1_000_000  
    upload_speed = st.upload() / 1_000_000  

    # Get ping (latency)
    ping = st.results.ping

    # Print the results
    print(f"Download Speed: {download_speed:.2f} Mbps")
    print(f"Upload Speed: {upload_speed:.2f} Mbps")
    print(f"Ping: {ping} ms")

def shutdown():
    """Shutdown the computer."""
    confirm = input("Are you sure you want to shut down the computer? (yes/no): ").strip().lower()
    if confirm == 'yes':
        os.system("shutdown /s /t 1")  # Windows shutdown command
        print("Shutting down...")
    else:
        print("Shutdown canceled.")

def restart():
    """Restart the computer."""
    confirm = input("Are you sure you want to restart the computer? (yes/no): ").strip().lower()
    if confirm == 'yes':
        os.system("shutdown /r /t 1")  # Windows restart command
        print("Restarting...")
    else:
        print("Restart canceled.")

def main():
    """Main menu for system control."""
    while True:
        print("\nSystem Control Tool")
        print("1. Shutdown")
        print("2. Restart")
        print("3. Exit")
        choice = input("Enter your choice (1/2/3): ").strip()
        
        if choice == '1':
            shutdown()
        elif choice == '2':
            restart()
        elif choice == '3':
            print("Exiting program...")
            break
        else:
            print("Invalid choice. Please try again.")

if __name__ == "__main__":
    main()
check_wifi_speed()


