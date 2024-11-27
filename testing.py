import os

processor_info = os.popen("wmic cpu get name").read()
print("Processor Details:",processor_info)

