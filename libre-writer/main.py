#!/usr/bin/env python3
import os
import sys
import subprocess
import socket
import time
import platform


def get_office_path():
    """Get the Collabora Office or LibreOffice executable path based on the operating system"""
    system = platform.system().lower()
    if system == "windows":
        # Windows paths
        possible_paths = [
            r"C:\Program Files\Collabora Office\program\soffice.exe",
            r"C:\Program Files (x86)\Collabora Office\program\soffice.exe",
            r"C:\Program Files\LibreOffice\program\soffice.exe",  # Fallback to LibreOffice
        ]
    elif system == "linux":
        # Linux paths - Collabora Office is typically installed in these locations
        possible_paths = [
            "/usr/bin/coolwsd",  # Collabora Online WebSocket Daemon
            "/usr/bin/collaboraoffice",  # Collabora Office main executable
            "/opt/collaboraoffice/program/soffice",
            "/usr/lib/collaboraoffice/program/soffice",
            # # Fallback to standard LibreOffice paths
            # '/usr/bin/soffice',
            # '/usr/lib/libreoffice/program/soffice',
            # '/opt/libreoffice/program/soffice'
        ]
    else:
        raise OSError(f"Unsupported operating system: {system}")

    for path in possible_paths:
        if os.path.exists(path):
            return path
    raise FileNotFoundError(
        "Neither Collabora Office nor LibreOffice executable found. Please install either office suite."
    )


def get_python_path():
    """Get the Python executable path based on the operating system"""
    system = platform.system().lower()
    if system == "windows":
        possible_paths = [
            r"C:\Program Files\Collabora Office\program\python.exe",
            r"C:\Program Files (x86)\Collabora Office\program\python.exe",
            r"C:\Program Files\LibreOffice\program\python.exe",
        ]
        for path in possible_paths:
            if os.path.exists(path):
                return path
        return sys.executable  # Fallback to system Python
    else:
        return sys.executable  # Use the current Python interpreter on Linux


def is_port_in_use(port):
    """Check if a port is already in use"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("localhost", port)) == 0


def start_office():
    """Start Collabora Office in headless mode with socket"""
    if not is_port_in_use(2002):
        print("Starting Collabora Office with socket...", file=sys.stderr)
        soffice_path = get_office_path()
        subprocess.Popen(
            [
                soffice_path,
                "-env:UserInstallation=file:///C:/Temp/LibreOfficeHeadlessProfile",
                "--headless",
                "--accept=socket,host=localhost,port=2002;urp;",
                "--norestore",
                "--nodefault",
                "--nologo",
            ]
        )
        time.sleep(3)  # Give it time to start
    else:
        print("Office socket already running on port 2002", file=sys.stderr)


def start_helper():
    """Start the Office helper script"""
    if not is_port_in_use(8765):
        print("Starting Office helper...", file=sys.stderr)
        helper_script = os.path.join(os.path.dirname(__file__), "helper.py")
        python_path = get_python_path()
        subprocess.Popen([python_path, helper_script])
        time.sleep(3)
    else:
        print("Helper script already running on port 8765", file=sys.stderr)


def start_mcp_server():
    """Start the MCP server"""
    print("Starting Office MCP server...", file=sys.stderr)
    server_script = os.path.join(os.path.dirname(__file__), "libre.py")
    subprocess.run(
        [
            sys.executable,  # Use the same Python interpreter
            server_script,
        ]
    )


def main():
    try:
        start_office()
        start_helper()
        start_mcp_server()
    except Exception as e:
        print(f"An error occurred: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
