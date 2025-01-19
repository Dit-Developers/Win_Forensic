import os
import platform
import ctypes
import socket
import subprocess
import psutil
from datetime import datetime
import time
from termcolor import colored
from tqdm import tqdm


# Developer Credits
DEVELOPER_NAME = "Muhammad Sudais Usmani"

# Yellow and Green ASCII Banner
ASCII_BANNER = colored(""" 
██╗    ██╗██╗███╗   ██╗    ███████╗ ██████╗ ██████╗ ███████╗███╗   ██╗███████╗██╗ ██████╗
██║    ██║██║████╗  ██║    ██╔════╝██╔═══██╗██╔══██╗██╔════╝████╗  ██║██╔════╝██║██╔════╝
██║ █╗ ██║██║██╔██╗ ██║    █████╗  ██║   ██║██████╔╝█████╗  ██╔██╗ ██║███████╗██║██║     
██║███╗██║██║██║╚██╗██║    ██╔══╝  ██║   ██║██╔══██╗██╔══╝  ██║╚██╗██║╚════██║██║██║     
╚███╔███╔╝██║██║ ╚████║    ██║     ╚██████╔╝██║  ██║███████╗██║ ╚████║███████║██║╚██████╗
 ╚══╝╚══╝ ╚═╝╚═╝  ╚═══╝    ╚═╝      ╚═════╝ ╚═╝  ╚═╝╚══════╝╚═╝  ╚═══╝╚══════╝╚═╝ ╚═════╝
""", 'yellow')

def is_windows():
    """Check if the operating system is Windows."""
    return platform.system() == "Windows"

def is_admin():
    """Check if the script is running with administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def clear_screen():
    """Clear the screen based on the operating system."""
    if platform.system() == "Windows":
        os.system('cls')
    else:
        os.system('clear')

def get_system_info():
    return {
        "OS": platform.system(),
        "OS Version": platform.version(),
        "OS Release": platform.release(),
        "Architecture": platform.architecture()[0],
        "Hostname": socket.gethostname(),
        "IP Address": socket.gethostbyname(socket.gethostname()),
        "Processor": platform.processor(),
        "Boot Time": datetime.fromtimestamp(psutil.boot_time()).strftime("%Y-%m-%d %H:%M:%S"),
    }

def get_user_accounts():
    try:
        result = subprocess.run(
            ["net", "user"], capture_output=True, text=True, check=True
        )
        return result.stdout.strip()
    except Exception as e:
        return f"Error retrieving user accounts: {e}"

def get_installed_programs():
    try:
        result = subprocess.run(
            ["wmic", "product", "get", "name,version"], capture_output=True, text=True
        )
        return result.stdout.strip()
    except Exception as e:
        return f"Error retrieving installed programs: {e}"

def get_event_logs():
    try:
        result = subprocess.run(
            ["wevtutil", "qe", "System", "/f:text", "/c:10"], capture_output=True, text=True
        )
        return result.stdout.strip()
    except Exception as e:
        return f"Error retrieving event logs: {e}"

def get_running_processes():
    processes = []
    for proc in psutil.process_iter(attrs=["pid", "name", "username"]):
        try:
            processes.append(proc.info)
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return processes

def get_network_connections():
    try:
        connections = psutil.net_connections()
        return [
            {
                "Type": conn.type,
                "Local Address": f"{conn.laddr.ip}:{conn.laddr.port}",
                "Remote Address": f"{conn.raddr.ip}:{conn.raddr.port}" if conn.raddr else "N/A",
                "Status": conn.status,
            }
            for conn in connections
        ]
    except Exception as e:
        return f"Error retrieving network connections: {e}"

def generate_html_report(system_info, user_accounts, programs, logs, processes, connections):
    html_template = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Windows Forensic Report</title>
    </head>
    <body>
        <h1>Windows Forensic Report</h1>
        <p>Developed by <strong>{DEVELOPER_NAME}</strong></p>
        <a href="https://www.linkedin.com/in/muhammad-sudais-usmani-950889311/">LinkedIn</a>
        <a href="https://github.com/Dit-Developers/">GitHub</a>
        <hr>
        <h2>System Information</h2>
        <table border="1">
            <tr><th>Key</th><th>Value</th></tr>
    """
    # Add system info to the HTML template
    for key, value in system_info.items():
        html_template += f"<tr><td>{key}</td><td>{value}</td></tr>"

    html_template += """
        </table>
        <h2>User Accounts</h2>
        <pre>{user_accounts}</pre>
        <h2>Installed Programs</h2>
        <pre>{programs}</pre>
        <h2>Event Logs</h2>
        <pre>{logs}</pre>
        <h2>Running Processes</h2>
        <table border="1">
            <tr><th>PID</th><th>Name</th><th>Username</th></tr>
    """
    # Add running processes to the HTML template
    for proc in processes:
        html_template += f"""
        <tr>
            <td>{proc['pid']}</td>
            <td>{proc['name']}</td>
            <td>{proc['username']}</td>
        </tr>
        """

    html_template += """
        </table>
        <h2>Network Connections</h2>
        <table border="1">
            <tr><th>Type</th><th>Local Address</th><th>Remote Address</th><th>Status</th></tr>
    """
    # Add network connections to the HTML template
    for conn in connections:
        html_template += f"""
        <tr>
            <td>{conn['Type']}</td>
            <td>{conn['Local Address']}</td>
            <td>{conn['Remote Address']}</td>
            <td>{conn['Status']}</td>
        </tr>
        """

    html_template += """
        </table>
    </body>
    </html>
    """
    # Return formatted HTML
    return html_template.format(
        DEVELOPER_NAME=DEVELOPER_NAME,
        user_accounts=user_accounts, 
        programs=programs, 
        logs=logs
    )

def save_report(html_content, file_name="forensic_report.html"):
    with open(file_name, "w") as file:
        file.write(html_content)
    print(colored(f"Report saved as {file_name}", 'green'))

if __name__ == "__main__":
    clear_screen()  # Clear the screen at the start
    print(ASCII_BANNER)

    if not is_windows():
        print(colored("This tool can only run on Windows platforms.", 'red'))
        exit(1)
    if not is_admin():
        print(colored("Please run this script as an administrator.", 'red'))
        exit(1)

    tasks = [
        ("Collecting system information...", get_system_info),
        ("Collecting user accounts...", get_user_accounts),
        ("Collecting installed programs...", get_installed_programs),
        ("Collecting event logs...", get_event_logs),
        ("Collecting running processes...", get_running_processes),
        ("Collecting network connections...", get_network_connections),
    ]

    system_info = {}
    user_accounts = ""
    programs = ""
    logs = ""
    processes = []
    connections = []

    for task_desc, task_func in tasks:
        print(colored(task_desc, 'yellow'))
        for _ in tqdm(range(1), desc=task_desc, ncols=100, colour='yellow'):
            time.sleep(0.5)  # Simulate loading time
        if task_func == get_system_info:
            system_info = task_func()
        elif task_func == get_user_accounts:
            user_accounts = task_func()
        elif task_func == get_installed_programs:
            programs = task_func()
        elif task_func == get_event_logs:
            logs = task_func()
        elif task_func == get_running_processes:
            processes = task_func()
        elif task_func == get_network_connections:
            connections = task_func()

    print(colored("Generating HTML report...", 'yellow'))
    html_report = generate_html_report(system_info, user_accounts, programs, logs, processes, connections)

    print(colored("Saving report...", 'yellow'))
    save_report(html_report)

    print(colored("Done!", 'green'))
