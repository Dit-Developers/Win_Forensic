import os
import platform
import ctypes
import socket
import subprocess
import psutil
from datetime import datetime

def is_windows():
    """Check if the operating system is Windows."""
    return platform.system() == "Windows"

def is_admin():
    """Check if the script is running with administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

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
    html_template = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Windows Forensic Report</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    </head>
    <body>
        <div class="container my-4">
            <h1 class="text-center mb-4">Windows Forensic Report</h1>
            
            <h2>System Information</h2>
            <table class="table table-bordered">
                <tbody>
    """
    for key, value in system_info.items():
        html_template += f"<tr><th>{key}</th><td>{value}</td></tr>"

    html_template += """
                </tbody>
            </table>

            <h2>User Accounts</h2>
            <pre>{user_accounts}</pre>

            <h2>Installed Programs</h2>
            <pre>{programs}</pre>

            <h2>Recent Event Logs</h2>
            <pre>{logs}</pre>

            <h2>Running Processes</h2>
            <table class="table table-bordered">
                <thead>
                    <tr>
                        <th>PID</th>
                        <th>Name</th>
                        <th>Username</th>
                    </tr>
                </thead>
                <tbody>
    """
    for proc in processes:
        html_template += f"""
        <tr>
            <td>{proc['pid']}</td>
            <td>{proc['name']}</td>
            <td>{proc['username']}</td>
        </tr>
        """

    html_template += """
                </tbody>
            </table>

            <h2>Network Connections</h2>
            <table class="table table-bordered">
                <thead>
                    <tr>
                        <th>Type</th>
                        <th>Local Address</th>
                        <th>Remote Address</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
    """
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
                </tbody>
            </table>
        </div>
    </body>
    </html>
    """
    return html_template.format(
        user_accounts=user_accounts, programs=programs, logs=logs
    )

def save_report(html_content, file_name="forensic_report.html"):
    with open(file_name, "w") as file:
        file.write(html_content)
    print(f"Report saved as {file_name}")

if __name__ == "__main__":
    if not is_windows():
        print("This tool can only run on Windows platforms.")
        exit(1)
    if not is_admin():
        print("Please run this script as an administrator.")
        exit(1)

    print("Collecting system information...")
    system_info = get_system_info()

    print("Collecting user accounts...")
    user_accounts = get_user_accounts()

    print("Collecting installed programs...")
    programs = get_installed_programs()

    print("Collecting event logs...")
    logs = get_event_logs()

    print("Collecting running processes...")
    processes = get_running_processes()

    print("Collecting network connections...")
    connections = get_network_connections()

    print("Generating HTML report...")
    html_report = generate_html_report(system_info, user_accounts, programs, logs, processes, connections)

    print("Saving report...")
    save_report(html_report)

    print("Done!")
