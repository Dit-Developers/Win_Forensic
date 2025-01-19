Win_Forensic
============

Win_Forensic is a Python-based Windows forensic tool designed to gather critical system information, user accounts, installed programs, event logs, running processes, and network connections. It generates a comprehensive HTML report styled with Bootstrap for easy analysis.

Features
--------
- Collects detailed **system information** (OS, hostname, IP, boot time, etc.).
- Retrieves **user accounts** and **installed programs**.
- Extracts the most recent **event logs**.
- Lists all **running processes** with PID, name, and username.
- Displays **network connections** (local/remote address, status).
- Generates an HTML report styled with **Bootstrap**.

Requirements
------------
- Python 3.x
- Windows operating system
- Administrator privileges

Installation
------------
1. Clone the repository:
   git clone https://github.com/Dit-Developers/Win_Forensic.git

2. Navigate to the project directory:
   cd Win_Forensic

Usage
-----
1. Ensure you run the script as an administrator.
2. Execute the tool:
   python forensic_tool.py

3. The forensic report will be saved as `forensic_report.html` in the project directory.

License
-------
This project is licensed under the MIT License.
