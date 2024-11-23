import os
import psutil
import shutil
import speedtest
import subprocess
import platform

import subprocess
import psutil
import platform
import os
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font
from openpyxl import load_workbook

# Function to get the serial number
def get_serial_number():
    try:
        result = subprocess.check_output("wmic bios get serialnumber", shell=True).decode().strip().split("\n")
        return result[1] if len(result) > 1 else "No Serial Number Found"
    except Exception as e:
        return f"Error: {e}"

# Function to get memory info
def get_memory_info():
    virtual_memory = psutil.virtual_memory()
    return {
        "Total Memory (GB)": virtual_memory.total / (1024 ** 3),
        "Used Memory (GB)": virtual_memory.used / (1024 ** 3),
        "Available Memory (GB)": virtual_memory.available / (1024 ** 3),
        "Memory Usage (%)": virtual_memory.percent
    }

# Function to get disk info
def get_disk_info():
    disk_usage = psutil.disk_usage('/')
    return {
        "Total Disk Space (GB)": disk_usage.total / (1024 ** 3),
        "Used Disk Space (GB)": disk_usage.used / (1024 ** 3),
        "Free Disk Space (GB)": disk_usage.free / (1024 ** 3),
        "Disk Usage (%)": disk_usage.percent
    }

# Function to get the username
def get_username():
    try:
        return os.getlogin()
    except Exception as e:
        return f"Error: {e}"

# Function to get the system remark
def get_system_remark():
    try:
        result = subprocess.check_output("wmic computersystem get description", shell=True).decode().strip().split("\n")
        return result[1] if len(result) > 1 else "No Description Found"
    except Exception as e:
        return f"Error: {e}"

# Gather all data
memory_info = get_memory_info()
disk_info = get_disk_info()
specification_values = (
    f"Total Memory: {memory_info['Total Memory (GB)']:.2f} GB, "
    f"Used Memory: {memory_info['Used Memory (GB)']:.2f} GB, "
    f"Total Disk Space: {disk_info['Total Disk Space (GB)']:.2f} GB"
)

data = {
    "Username": get_username(),
    "Serial Number": get_serial_number(),
    "System Model": platform.uname().machine,
    "System Manufacturer": platform.uname().system,
    "Asset Tag": platform.uname().node,
    "System Remark": get_system_remark(),
    "Specification": specification_values,
}

# Export to Excel
def export_to_excel(data, filename="window.xlsx"):
    try:
        # Check if the file already exists
        if os.path.exists(filename):
            wb = load_workbook(filename)
            sheet = wb.active
        else:
            # Create a new workbook and set up headers
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet.title = "System Info"
            headers = [
                "N",
                "Specification",
                "Username",
                "Serial Number",
                "System Model",
                "System Manufacturer",
                "Asset Tag",
                "System Remark",
            ]
            header_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
            header_font = Font(bold=True)

            for col_num, header in enumerate(headers, start=1):
                cell = sheet.cell(row=1, column=col_num)
                cell.value = header
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")

        # Find the next empty row
        next_row = sheet.max_row + 1

        # Populate the row
        sheet[f"A{next_row}"] = next_row - 1  # Numbering starts at 1
        sheet[f"B{next_row}"] = data.get("Specification", "N/A")
        sheet[f"C{next_row}"] = data.get("Username", "N/A")
        sheet[f"D{next_row}"] = data.get("Serial Number", "N/A")
        sheet[f"E{next_row}"] = data.get("System Model", "N/A")
        sheet[f"F{next_row}"] = data.get("System Manufacturer", "N/A")
        sheet[f"G{next_row}"] = data.get("Asset Tag", "N/A")
        sheet[f"H{next_row}"] = data.get("System Remark", "N/A")

        # Center-align all data
        for row in sheet.iter_rows(min_row=next_row, max_row=next_row, min_col=1, max_col=8):
            for cell in row:
                cell.alignment = Alignment(horizontal="center")

        # Save the workbook
        wb.save(filename)
        print(f"Data appended to {filename}")
    except Exception as e:
        print(f"Error exporting to Excel: {e}")

# Save results to Excel
export_to_excel(data)


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


