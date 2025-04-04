# LibreOffice MCP Server

A Model Context Protocol (MCP) server that enables Claude and other AI assistants to interact with LibreOffice documents directly.


## Overview

This MCP server provides a bridge between AI assistants like Claude and LibreOffice, allowing the AI to create, read, edit, and format LibreOffice documents. Built on the [Model Context Protocol](https://modelcontextprotocol.io/), it exposes LibreOffice functionality as tools that can be called by compatible AI applications.

## Features

### LibreOffice Suite Support

- [x] **LibreWriter**
  - **Implemented:**
    - [x] Create new documents
    - [x] Open and read text documents
    - [x] Copy documents
    - [x] List documents in a directory
    - [x] Get document properties and metadata
    - [x] Add text to documents
    - [x] Add headings with different levels
    - [x] Add formatted paragraphs
    - [x] Add tables with data
    - [x] Format tables with borders, colors, and header rows
    - [x] Search and replace text
    - [x] Delete specific text
    - [x] Delete paragraphs
    - [x] Insert page breaks (partial implementation)
    - [x] Insert images
    - [x] Format specific text (bold, italic, color, size)
  
    
- [ ] **LibreCalc**
- [ ] **LibreImpress**

## Installation

### Requirements
- Python 3.10 or higher
- LibreOffice installed
- Claude for Desktop (or another MCP-compatible client)

### Setup Instructions

1. **Clone the repository**
   ```bash
   git clone https://github.com/harshithb3304/libre-office-mcp.git
   cd libre-office-mcp/libre-writer
   ```

2. **Set up Python environment**
   ```bash
   UV Installation 

   #For Windows  
   powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
   #For MacOS/Linux 
   curl -LsSf https://astral.sh/uv/install.sh | sh

   # With uv (recommended)
   uv init
   uv venv
   uv add mcp[cli] httpx
   
   # With pip
   pip install "mcp[cli]" httpx
   ```

3. **Install LibreOffice**
   - Windows: [Download from LibreOffice website](https://www.libreoffice.org/download/download-libreoffice/)
   - macOS: 
     ```bash
     brew install --cask libreoffice
     ```

## Usage with Claude for Desktop

1. **Configure Claude for Desktop**

   Add the following to your Claude for Desktop configuration file:
   - Windows: `%APPDATA%\Claude\claude_desktop_config.json`
   - macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`

   ```json
   {
     "mcpServers": {
       "libreoffice-server": {
         "command": "uv",
         "args": [
           "--directory",
           "C:\\path\\to\\libreoffice-mcp-server",
           "run",
           "main.py"
         ]
       }
     }
   }
   ``` 
   or 
   ```json 
   {
     "mcpServers": {
       "libreoffice-server": {
          "command": "python",
          "args": [
            "C:\\path\\to\\the\\main.py"
          ]
        }
      }
   }
   
   ```

   > ⚠️ **Note**: Replace `C:\\path\\to\\libreoffice-mcp-server` with the actual path to where you cloned the repository.

2. **Restart Claude for Desktop**

3. **Interact with LibreOffice**

   You can now ask Claude to perform actions like:
   - "Create a new document with a title and three paragraphs"
   - "List all the documents in my Documents folder"
   - "Open my report.odt file and add a table with 4 columns and 5 rows"

![Claude LibreOffice Interaction](assets\image.png)

## How It Works

The server consists of three main components:

1. **LibreOffice Helper (`helper.py`)**: A Python script that communicates directly with LibreOffice using the UNO bridge API.

2. **MCP Server (`libre.py`)**: The main MCP server that exposes LibreOffice functionality as tools for AI assistants.

3. **Launcher (`main.py`)**: A script that launches both the helper and the MCP server.

When a request comes in from Claude:

1. The MCP server receives the request
2. It forwards the command to the LibreOffice helper via a socket connection
3. The helper executes the command using LibreOffice's UNO API
4. The result is sent back to the MCP server and then to Claude

## Configuration Options

The server will use your Documents folder as the default location for creating new documents if no path is specified. You can change this by modifying the `get_default_document_path` function in `libre.py`.

## Troubleshooting

### Common Issues

- **"LibreOffice helper is not running"**  
  Make sure LibreOffice is installed and the path to the LibreOffice Python executable is correct in `main.py`.

- **"Connection refused"**  
  The helper script may not have started correctly. Check if port 8765 is already in use.

- **"Failed to connect to LibreOffice desktop"**  
  LibreOffice may not be running in headless mode. Check if port 2002 is available.

## Contributing

Contributions are welcome! Here are some areas that need improvement:

- Fixing image insertion functionality
- Implementing robust text formatting
- Supporting additional LibreOffice features
- Improving error handling

