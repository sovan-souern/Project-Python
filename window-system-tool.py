import tkinter as tk
from tkinter import ttk
import platform
import psutil
import shutil
import socket


# Function to update system information
def update_info():
    # Disk Usage
    total, used, free = shutil.disk_usage("/")
    total_gb = total / (1024 ** 3)
    used_gb = used / (1024 ** 3)
    free_gb = free / (1024 ** 3)
    disk_label.config(
        text=f"Total: {total_gb:.2f} GB\n"
             f"Used: {used_gb:.2f} GB\n"
             f"Free: {free_gb:.2f} GB"
    )

    # CPU Information
    cpu_cores = psutil.cpu_count(logical=True)
    cpu_usage = psutil.cpu_percent(interval=1)
    cpu_label.config(
        text=f"Cores: {cpu_cores}\n"
             f"Usage: {cpu_usage:.2f} %"
    )

    # Memory Information
    memory_info = psutil.virtual_memory()
    total_mem = memory_info.total // (1024 ** 2)
    available_mem = memory_info.available // (1024 ** 2)
    mem_usage = memory_info.percent
    memory_label.config(
        text=f"Total: {total_mem} MB\n"
             f"Available: {available_mem} MB\n"
             f"Usage: {mem_usage:.2f} %"
    )

    # Network Information
    net_info = psutil.net_io_counters()
    bytes_sent = net_info.bytes_sent // (1024 ** 2)
    bytes_recv = net_info.bytes_recv // (1024 ** 2)
    network_label.config(
        text=f"Sent: {bytes_sent} MB\n"
             f"Received: {bytes_recv} MB"
    )

    # System Information
    windows_version = platform.version()
    windows_release = platform.release()
    device_name = platform.node()
    processor = platform.processor()
    system_type = platform.architecture()[0]
    ip_address = socket.gethostbyname(socket.gethostname())
    system_label.config(
        text=f"Version: {windows_version}\n"
             f"Release: {windows_release}\n"
             f"Device: {device_name}\n"
             f"Processor: {processor}\n"
             f"Type: {system_type}\n"
             f"IP: {ip_address}"
    )


# Create the main window
root = tk.Tk()
root.title("Advanced System Information")
root.geometry("800x600")
root.configure(bg="#34495e")

# Title
title_label = tk.Label(
    root, text="Advanced System Information", font=("Helvetica", 18, "bold"),
    bg="#2c3e50", fg="#ecf0f1", pady=10
)
title_label.pack(fill="x")

# Notebook (Tab Control)
notebook = ttk.Notebook(root)
notebook.pack(pady=20, padx=10, expand=True, fill="both")

# Disk Tab
disk_tab = ttk.Frame(notebook)
notebook.add(disk_tab, text="Disk Info")
disk_label = tk.Label(disk_tab, font=("Helvetica", 14), justify="left")
disk_label.pack(pady=10, padx=10, anchor="w")

# CPU Tab
cpu_tab = ttk.Frame(notebook)
notebook.add(cpu_tab, text="CPU Info")
cpu_label = tk.Label(cpu_tab, font=("Helvetica", 14), justify="left")
cpu_label.pack(pady=10, padx=10, anchor="w")

# Memory Tab
memory_tab = ttk.Frame(notebook)
notebook.add(memory_tab, text="Memory Info")
memory_label = tk.Label(memory_tab, font=("Helvetica", 14), justify="left")
memory_label.pack(pady=10, padx=10, anchor="w")

# Network Tab
network_tab = ttk.Frame(notebook)
notebook.add(network_tab, text="Network Info")
network_label = tk.Label(network_tab, font=("Helvetica", 14), justify="left")
network_label.pack(pady=10, padx=10, anchor="w")

# System Tab
system_tab = ttk.Frame(notebook)
notebook.add(system_tab, text="System Info")
system_label = tk.Label(system_tab, font=("Helvetica", 14), fg="black", justify="left")
system_label.pack(pady=10, padx=10, anchor="w")

# Refresh Button
refresh_button = ttk.Button(root, text="Refresh Info", command=update_info)
refresh_button.pack(pady=10)

# Update Info Initially
update_info()

# Run the application
root.mainloop()
