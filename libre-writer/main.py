#!/usr/bin/env python3
import os
import sys
import subprocess
import socket
import time

def is_port_in_use(port):
    """Check if a port is already in use"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('localhost', port)) == 0

def start_libreoffice():
    """Start LibreOffice in headless mode with socket"""
    if not is_port_in_use(2002):
        print("Starting LibreOffice with socket...", file=sys.stderr)
        subprocess.Popen([
            r"C:\Program Files\LibreOffice\program\soffice.exe", 
            "--headless", 
            "--accept=socket,host=localhost,port=2002;urp;",
            "--norestore", 
            "--nodefault", 
            "--nologo"
        ])
        time.sleep(3)  # Give it time to start
    else:
        print("LibreOffice socket already running on port 2002", file=sys.stderr)

def start_helper():
    """Start the LibreOffice helper script"""
    print("Starting LibreOffice helper...", file=sys.stderr)
    helper_script = os.path.join(os.path.dirname(__file__), 'helper.py')
    subprocess.Popen([
        r"C:\Program Files\LibreOffice\program\python.exe", 
        helper_script
    ])
    time.sleep(3)

def start_mcp_server():
    """Start the MCP server"""
    print("Starting LibreOffice MCP server...", file=sys.stderr)
    server_script = os.path.join(os.path.dirname(__file__), 'libre.py')
    subprocess.run([
        sys.executable,  # Use the same Python interpreter
        server_script
    ])

def main():
    try:
        start_libreoffice()
        start_helper()
        start_mcp_server()
    except Exception as e:
        print(f"An error occurred: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()