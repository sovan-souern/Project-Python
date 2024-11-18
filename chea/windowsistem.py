import psutil

# Get CPU usage
cpu_usage = psutil.cpu_percent(interval=1)
print(f"CPU Usage: {cpu_usage}%")

# Get Memory usage
memory = psutil.virtual_memory()
print(f"Memory Usage: {memory.percent}%")
print(f"Total Memory: {memory.total / (1024 ** 3):.2f} GB")
print(f"Available Memory: {memory.available / (1024 ** 3):.2f} GB")

# Get Disk usage
disk_usage = psutil.disk_usage('/')
print(f"Disk Usage: {disk_usage.percent}%")
print(f"Total Disk Space: {disk_usage.total / (1024 ** 3):.2f} GB")
print(f"Used Disk Space: {disk_usage.used / (1024 ** 3):.2f} GB")
print(f"Free Disk Space: {disk_usage.free / (1024 ** 3):.2f} GB")

# Get Network stats
net_io = psutil.net_io_counters()
print(f"Total Bytes Sent: {net_io.bytes_sent / (1024 ** 2):.2f} MB")
print(f"Total Bytes Received: {net_io.bytes_recv / (1024 ** 2):.2f} MB")


