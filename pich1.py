#•	Shutdown/Restart: The tool should allow users to shut down or restart their Windows machine using Python commands.
import os

# To shut down the computer
os.system("shutdown /s /t 1")  # For Windows
os.system("shutdown now")      # For Linux

# To restart the computer
os.system("shutdown /r /t 1")  # For Windows
os.system("reboot")            # For Linux

# # restart their Windows machine using Python commands.
# import os
# import platform

# def restart_computer():
#     # Check the operating system
#     current_os = platform.system()

#     if current_os == "Windows":
#         # Restart command for Windows
#         os.system("shutdown /r /t 1")
#     elif current_os == "Linux":
#         # Restart command for Linux
#         os.system("reboot")
#     else:
#         print("Unsupported operating system.")

# # Call the function
# restart_computer()
