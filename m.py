import os
import platform
import psutil
import tkinter as tk
from tkinter import ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import speedtest
import subprocess
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

def fetch_system_info():
    """Fetch and display basic system information."""
    info = [
        f"OS: {platform.system()} {platform.release()}",
        f"Architecture: {platform.architecture()[0]}",
        f"Processor: {platform.processor()}",
        f"RAM: {round(psutil.virtual_memory().total / (1024 ** 3), 2)} GB",
        f"Disk: {round(psutil.disk_usage('/').total / (1024 ** 3), 2)} GB",
    ]
    return "\n".join(info)


def cleanup_temp_files():
    """Placeholder for cleaning up temporary files."""
    print("Cleanup functionality triggered!")
    tk.messagebox.showinfo("Cleanup", "Temporary files cleaned successfully!")


def show_system_info(info_label):
    """Update the info section with system details."""
    info = fetch_system_info()
    info_label.config(text=info)


def update_task_list(tree):
    """Update the task list in the UI."""
    for row in tree.get_children():
        tree.delete(row)
    for proc in psutil.process_iter(['pid', 'name', 'cpu_percent']):
        try:
            tree.insert("", "end", values=(proc.info['pid'], proc.info['name'], proc.info['cpu_percent']))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue


def update_performance(cpu_line, mem_line, disk_line, x_data, cpu_data, mem_data, disk_data, canvas):
    """Update performance metrics dynamically."""
    cpu_data.append(psutil.cpu_percent())
    mem_data.append(psutil.virtual_memory().percent)
    disk_data.append(psutil.disk_usage('/').percent)
    x_data.append(len(x_data) + 1)

    if len(x_data) > 30:
        x_data.pop(0)
        cpu_data.pop(0)
        mem_data.pop(0)
        disk_data.pop(0)

    cpu_line.set_data(x_data, cpu_data)
    mem_line.set_data(x_data, mem_data)
    disk_line.set_data(x_data, disk_data)

    ax = cpu_line.axes
    ax.set_xlim(max(0, len(x_data) - 30), len(x_data))
    ax.set_ylim(0, 100)
    canvas.draw()
    canvas.get_tk_widget().after(1000, update_performance, cpu_line, mem_line, disk_line, x_data, cpu_data, mem_data, disk_data, canvas)


def check_wifi_speed():
    """Check the download and upload speed of the active Wi-Fi connection."""
    st = speedtest.Speedtest()
    st.get_best_server()
    download_speed = st.download() / 1_000_000  
    upload_speed = st.upload() / 1_000_000  
    ping = st.results.ping
    return download_speed, upload_speed, ping


def get_wifi_passwords():
    """Retrieve stored Wi-Fi passwords."""
    profiles_cmd = "netsh wlan show profiles"
    profiles = subprocess.check_output(profiles_cmd, shell=True, text=True)
    profile_names = [line.split(":")[1].strip() for line in profiles.splitlines() if "All User Profile" in line]

    wifi_passwords = {}
    for profile in profile_names:
        password_cmd = f"netsh wlan show profile \"{profile}\" key=clear"
        profile_info = subprocess.check_output(password_cmd, shell=True, text=True)
        for line in profile_info.splitlines():
            if "Key Content" in line:
                password = line.split(":")[1].strip()
                wifi_passwords[profile] = password
                break
    return wifi_passwords


def get_serial_number():
    try:
        serial_number = subprocess.check_output("wmic bios get serialnumber").decode().strip().split('\n')[1]
        return serial_number
    except subprocess.CalledProcessError:
        return "Unable to retrieve serial number"


def get_cpu_info():
    cpu_count = psutil.cpu_count(logical=True)
    cpu_freq = psutil.cpu_freq()
    cpu_percent = psutil.cpu_percent(interval=1)
    return f"CPU Count (Logical Cores): {cpu_count}\nCPU Frequency: {cpu_freq.current} MHz\nCPU Usage: {cpu_percent}%"


def get_memory_info():
    virtual_memory = psutil.virtual_memory()
    return {
        "Total Memory (GB)": virtual_memory.total / (1024 ** 3),
        "Used Memory (GB)": virtual_memory.used / (1024 ** 3),
        "Available Memory (GB)": virtual_memory.available / (1024 ** 3),
        "Memory Usage (%)": virtual_memory.percent
    }


def get_disk_info():
    disk_usage = psutil.disk_usage('/')
    return {
        "Total Disk Space (GB)": disk_usage.total / (1024 ** 3),
        "Used Disk Space (GB)": disk_usage.used / (1024 ** 3),
        "Free Disk Space (GB)": disk_usage.free / (1024 ** 3),
        "Disk Usage (%)": disk_usage.percent
    }


def create_full_interface():
    root = tk.Tk()
    root.title("Complete System Dashboard")
    root.geometry("1400x800")

    sidebar = ttk.Frame(root, width=200, relief="raised", padding=10)
    sidebar.pack(side=tk.LEFT, fill=tk.Y)
    ttk.Label(sidebar, text="Utilities", font=("Segoe UI", 14, "bold")).pack(pady=10)
    ttk.Button(sidebar, text="Wi-Fi Credentials", command=lambda: show_wifi_passwords(info_label)).pack(fill=tk.X, pady=5)
    ttk.Button(sidebar, text="Wi-Fi Speed Test", command=check_wifi_speed).pack(fill=tk.X, pady=5)
    ttk.Button(sidebar, text="System Info", command=lambda: show_system_info(info_label)).pack(fill=tk.X, pady=5)
    ttk.Button(sidebar, text="CPU Info", command=lambda: show_cpu_info(info_label)).pack(fill=tk.X, pady=5)
    ttk.Button(sidebar, text="Memory Info", command=lambda: show_memory_info(info_label)).pack(fill=tk.X, pady=5)
    ttk.Button(sidebar, text="Disk Info", command=lambda: show_disk_info(info_label)).pack(fill=tk.X, pady=5)
    ttk.Button(sidebar, text="Cleanup", command=cleanup_temp_files).pack(fill=tk.X, pady=5)

    main_frame = ttk.Frame(root, padding=10)
    main_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

    graph_frame = ttk.LabelFrame(main_frame, text="Performance Graphs")
    graph_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.set_title("System Performance Metrics")
    ax.set_xlabel("Time (s)")
    ax.set_ylabel("Usage (%)")
    ax.grid(True)

    x_data, cpu_data, mem_data, disk_data = [], [], [], []
    cpu_line, = ax.plot([], [], label="CPU", color="red")
    mem_line, = ax.plot([], [], label="Memory", color="blue")
    disk_line, = ax.plot([], [], label="Disk", color="green")
    ax.legend()
    canvas = FigureCanvasTkAgg(fig, master=graph_frame)
    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    info_frame = ttk.LabelFrame(main_frame, text="System Information")
    info_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    info_label = ttk.Label(info_frame, text="Click 'System Info' to fetch details", anchor="w", justify="left")
    info_label.pack(fill=tk.BOTH, padx=5, pady=5)

    task_frame = ttk.LabelFrame(main_frame, text="Task Manager")
    task_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    tree = ttk.Treeview(task_frame, columns=("PID", "Name", "CPU"), show="headings")
    tree.heading("PID", text="PID")
    tree.heading("Name", text="Name")
    tree.heading("CPU", text="CPU (%)")
    tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    ttk.Button(task_frame, text="Refresh Tasks", command=lambda: update_task_list(tree)).pack(pady=5)

    update_performance(cpu_line, mem_line, disk_line, x_data, cpu_data, mem_data, disk_data, canvas)

    root.mainloop()


def show_wifi_passwords(info_label):
    passwords = get_wifi_passwords()
    output = "\n".join(f"Profile: {key}, Password: {value}" for key, value in passwords.items())
    info_label.config(text=output)
def update_task_list(tree):
    """Update the task list in the UI with CPU, memory, and disk usage."""
    for row in tree.get_children():
        tree.delete(row)
    
    # Iterate over the processes and fetch information
    for proc in psutil.process_iter(['pid', 'name', 'cpu_percent', 'memory_info']):
        try:
            # Get memory usage in MB and disk usage
            memory_usage = proc.info['memory_info'].rss / (1024 ** 2)  # Convert to MB
            disk_usage = 0 

            tree.insert("", "end", values=(proc.info['pid'], proc.info['name'], proc.info['cpu_percent'], round(memory_usage, 2), disk_usage))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue



if __name__ == "__main__":
    create_full_interface()
