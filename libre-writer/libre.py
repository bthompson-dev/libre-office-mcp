#!/usr/bin/env python3
import os
import sys
from typing import Optional, Any, List, Dict
import socket
import json
import time
import logging

log_path = os.path.join(os.path.dirname(__file__), "libre.log")
logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s"
)

# Extensions
writer_extensions = (
            ".odt",
            ".docx",
            ".dotx",
            ".xml",
            ".doc",
            ".dot",
            ".rtf",
            ".wpd",
        )

impress_extensions = (
            ".odp",
            ".pptx",
            ".ppsx",
            ".ppmx",
            ".potx",
            ".pomx",
            ".ppt",
            ".pps",
            ".ppm",
            ".pot",
            ".pom",
        )


# Helper function for directory management
def ensure_directory_exists(file_path: str) -> None:
    """Ensure the directory for a file exists, creating it if necessary."""
    directory = os.path.dirname(file_path)
    if directory and not os.path.exists(directory):
        try:
            os.makedirs(directory, exist_ok=True)
            print(f"Created directory: {directory}")
        except Exception as e:
            print(f"Failed to create directory {directory}: {str(e)}")
            raise

# Helper function to normalize file paths
def normalize_path(file_path: str) -> str:
    """Convert a relative path to an absolute path."""
    if not file_path:
        raise ValueError("Empty file path provided")
    
    # Expand user directory if path starts with ~
    if file_path.startswith('~'):
        file_path = os.path.expanduser(file_path)
        
    # Make absolute if relative
    if not os.path.isabs(file_path):
        file_path = os.path.abspath(file_path)
        
    print(f"Normalized path: {file_path}")
    return file_path

# Function to communicate with the LibreOffice helper
def call_libreoffice_helper(command: dict) -> dict:
    """
    Send a command to the LibreOffice helper process.
    
    Args:
        command: Dictionary with command details
        
    Returns:
        Dictionary with response from helper
    """
    try:
        logging.info("call_libreoffice_helper function called")

        client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        client_socket.settimeout(30)  # 30 second timeout
        client_socket.connect(('localhost', 8765))
        
        logging.info(client_socket)

        # Send command
        request_data = json.dumps(command).encode('utf-8')
        client_socket.send(request_data)

        logging.info(request_data)
        
        # Receive response
        response_data = client_socket.recv(16384).decode('utf-8')
        client_socket.close()
        
        logging.info(response_data)

        if not response_data:
            return {"status": "error", "message": "Empty response from helper"}
            
        return json.loads(response_data)
    except socket.timeout:
        return {"status": "error", "message": "Connection to helper timed out"}
    except ConnectionRefusedError:
        return {"status": "error", "message": "Connection refused. Is the helper script running?"}
    except Exception as e:
        return {"status": "error", "message": f"Error communicating with helper: {str(e)}"}

# Now import MCP dependencies
try:
    from mcp.server.fastmcp import FastMCP, Context
except ImportError as e:
    print(f"Dependency Import Error: {e}")
    print("Ensure you have installed:")
    print("- mcp[cli]")
    sys.exit(1)

# Initialize MCP server
mcp = FastMCP("libreoffice-server")

# Helper function to get default document path on Windows
def get_default_document_path(filename: str) -> str:
    """Get default path for documents on Windows."""
    # Use Documents folder on Windows
    docs_path = os.path.join(os.path.expanduser("~"), "Documents")
    return os.path.join(docs_path, filename)

# Document Management Tools

@mcp.tool()
async def create_blank_document(
    filename: str,
    title: str = "",
    author: str = "",
    subject: str = "",
    keywords: str = "",
) -> str:
    """
    Create a new LibreOffice Writer document.

    Args:
        filename: Name of the document to create
        title: Document title metadata
        author: Document author metadata
        subject: Document subject metadata
        keywords: Document keywords metadata (comma-separated)
    """
    try:
        # Check for supported file extensions
        if not filename.lower().endswith(writer_extensions):
            filename += ".odt"

        # If filename doesn't include a path, add default path
        if os.path.basename(filename) == filename:
            save_path = get_default_document_path(filename)
        else:
            save_path = filename

        # Normalize path
        save_path = normalize_path(save_path)
        print(f"Creating document at: {save_path}")

        # Prepare metadata if provided
        metadata = {}
        if title:
            metadata["Title"] = title
        if author:
            metadata["Author"] = author
        if subject:
            metadata["Subject"] = subject
        if keywords:
            # Split by comma and strip whitespace
            keywords = [k.strip() for k in keywords.split(",")]
            metadata["Keywords"] = keywords

        logging.info(save_path)
        logging.info(metadata)

        # Send command to helper
        response = call_libreoffice_helper(
            {
                "action": "create_document",
                "doc_type": "text",
                "file_path": save_path,
                "metadata": metadata,
            }
        )

        logging.info(response["status"])
        logging.info(response["message"])

        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"

    except Exception as e:
        logging.error(f"Error in create_document: {str(e)}")
        return f"Failed to create document: {str(e)}"

@mcp.tool()
async def read_text_document(file_path: str) -> str:
    """
    Open and read a text document.
    
    Args:
        file_path: Path to the document
    """
    try:
        # Normalize path
        file_path = normalize_path(file_path)
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "open_text_document",
            "file_path": file_path
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in open_text_document: {str(e)}")
        return f"Failed to open document: {str(e)}"

@mcp.tool()
async def get_document_properties(file_path: str) -> str:
    """
    Get document properties and statistics, including author, description, keywords, word count, etc.
    
    Args:
        file_path: Path to the document
    """
    try:
        # Normalize path
        file_path = normalize_path(file_path)
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "get_document_properties",
            "file_path": file_path
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in get_document_properties: {str(e)}")
        return f"Failed to get document properties: {str(e)}"

@mcp.tool()
async def list_documents(directory: str) -> str:
    """
    List all documents in a directory.
    
    Args:
        directory: Path to the directory to scan
    """
    try:
        # Normalize path
        directory = normalize_path(directory)
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "list_documents",
            "directory": directory
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in list_documents: {str(e)}")
        return f"Failed to list documents: {str(e)}"

@mcp.tool()
async def copy_document(source_path: str, target_path: str) -> str:
    """
    Create a copy of an existing document.
    
    Args:
        source_path: Path to the document to copy
        target_path: Path where to save the copy
    """
    try:
        # Normalize paths
        source_path = normalize_path(source_path)
        target_path = normalize_path(target_path)
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "copy_document",
            "source_path": source_path,
            "target_path": target_path
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in copy_document: {str(e)}")
        return f"Failed to copy document: {str(e)}"

# Content Creation Tools

@mcp.tool()
async def add_text(file_path: str, text: str, position: Optional[str] = "end") -> str:
    """
    Add text to a LibreOffice document.
    
    Args:
        file_path: Path to the document
        text: Text to add
        position: Where to add text (start, end, or cursor)
    """
    try:
        # Normalize path
        file_path = normalize_path(file_path)
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "add_text",
            "file_path": file_path,
            "text": text,
            "position": position
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in add_text: {str(e)}")
        return f"Failed to add text: {str(e)}"

@mcp.tool()
async def add_heading(file_path: str, text: str, level: int = 1) -> str:
    """
    Add a heading to a document.
    
    Args:
        file_path: Path to the document
        text: Heading text
        level: Heading level (1-6, where 1 is the highest level)
    """
    try:
        # Normalize path
        file_path = normalize_path(file_path)
        
        # Validate heading level
        if level < 1 or level > 6:
            return f"Invalid heading level: {level}. Choose a level between 1 and 6."
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "add_heading",
            "file_path": file_path,
            "text": text,
            "level": level
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in add_heading: {str(e)}")
        return f"Failed to add heading: {str(e)}"

@mcp.tool()
async def add_paragraph(file_path: str, text: str, style: Optional[str] = None, alignment: Optional[str] = None) -> str:
    """
    Add a paragraph with optional styling.
    
    Args:
        file_path: Path to the document
        text: Paragraph text
        style: Paragraph style name (if available in document)
        alignment: Text alignment (left, center, right, justify)
    """
    try:
        # Normalize path
        file_path = normalize_path(file_path)
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "add_paragraph",
            "file_path": file_path,
            "text": text,
            "style": style,
            "alignment": alignment
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in add_paragraph: {str(e)}")
        return f"Failed to add paragraph: {str(e)}"

@mcp.tool()
async def add_table(file_path: str, rows: int, columns: int, data: Optional[List[List[str]]] = None, header_row: bool = False) -> str:
    """
    Add a table to a LibreOffice text document.
    
    Args:
        file_path: Path to the document
        rows: Number of rows
        columns: Number of columns
        data: Optional 2D list of data to populate the table
        header_row: Whether to format the first row as a header
    """
    try:
        # Normalize path
        file_path = normalize_path(file_path)
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "add_table",
            "file_path": file_path,
            "rows": rows,
            "columns": columns,
            "data": data,
            "header_row": header_row
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in add_table: {str(e)}")
        return f"Failed to add table: {str(e)}"

@mcp.tool()
async def insert_image(file_path: str, image_path: str, width: Optional[int] = None, height: Optional[int] = None) -> str:
    """
    Insert an image into a LibreOffice document with optional resizing.
    
    Args:
        file_path: Path to the target document
        image_path: Path to the image file to insert
        width: Optional width in 100ths of mm (maintains aspect ratio if only width is specified)
        height: Optional height in 100ths of mm (maintains aspect ratio if only height is specified)
    """
    try:
        # Normalize paths
        file_path = normalize_path(file_path)
        image_path = normalize_path(image_path)
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "insert_image",
            "file_path": file_path,
            "image_path": image_path,
            "width": width,
            "height": height
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in insert_image: {str(e)}")
        return f"Failed to insert image: {str(e)}"

@mcp.tool()
async def insert_page_break(file_path: str) -> str:
    """
    Insert a page break at the end of a document.
    
    Args:
        file_path: Path to the document
    """
    try:
        # Normalize path
        file_path = normalize_path(file_path)
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "insert_page_break",
            "file_path": file_path
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in insert_page_break: {str(e)}")
        return f"Failed to insert page break: {str(e)}"

# Text Formatting Tools

@mcp.tool()
async def format_text(file_path: str, text_to_find: str, bold: bool = False, italic: bool = False,
                     underline: bool = False, color: Optional[str] = None,
                     font: Optional[str] = None, size: Optional[float] = None) -> str:
    """Format specific text in a document."""
    try:
        # Normalize path
        file_path = normalize_path(file_path)
        
        # Send command to helper with all parameters using the correct names
        response = call_libreoffice_helper({
            "action": "format_text",
            "file_path": file_path,  # Note: using file_path, not filepath
            "text_to_find": text_to_find,
            "bold": bold,
            "italic": italic,
            "underline": underline,
            "color": color,
            "font": font,
            "size": size
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in format_text: {str(e)}")
        return f"Failed to format text: {str(e)}"

@mcp.tool()
async def search_replace_text(file_path: str, search_text: str, replace_text: str) -> str:
    """
    Search and replace text throughout the document.
    
    Args:
        file_path: Path to the document
        search_text: Text to search for
        replace_text: Text to replace with
    """
    try:
        # Normalize path
        file_path = normalize_path(file_path)
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "search_replace_text",
            "file_path": file_path,
            "search_text": search_text,
            "replace_text": replace_text
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in search_replace_text: {str(e)}")
        return f"Failed to search and replace text: {str(e)}"

@mcp.tool()
async def delete_text(file_path: str, text_to_delete: str) -> str:
    """
    Delete specific text from the document.
    
    Args:
        file_path: Path to the document
        text_to_delete: Text to search for and delete
    """
    try:
        # Normalize path
        file_path = normalize_path(file_path)
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "delete_text",
            "file_path": file_path,
            "text_to_delete": text_to_delete
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in delete_text: {str(e)}")
        return f"Failed to delete text: {str(e)}"

# Table Formatting Tools

@mcp.tool()
async def format_table(file_path: str, table_index: int = 0, border_width: Optional[int] = None, 
                      background_color: Optional[str] = None, header_row: bool = False) -> str:
    """
    Format a table with borders, shading, etc.
    
    Args:
        file_path: Path to the document
        table_index: Index of the table to format (0 = first table)
        border_width: Border width in points
        background_color: Background color (hex format, e.g., "#F0F0F0")
        header_row: Whether to format the first row as a header
    """
    try:
        # Normalize path
        file_path = normalize_path(file_path)
        
        # Prepare format options
        format_options = {}
        
        if border_width is not None:
            format_options["border_width"] = border_width
        
        if background_color:
            format_options["background_color"] = background_color

        format_options["header_row"] = header_row
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "format_table",
            "file_path": file_path,
            "table_index": table_index,
            "format_options": format_options
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in format_table: {str(e)}")
        return f"Failed to format table: {str(e)}"

# Advanced Document Manipulation Tools

@mcp.tool()
async def create_custom_style(file_path: str, style_name: str, font_name: Optional[str] = None, 
                             font_size: Optional[float] = None, bold: bool = False, italic: bool = False,
                             underline: bool = False, color: Optional[str] = None, 
                             alignment: Optional[str] = None) -> str:
    """
    Create a custom paragraph style.
    
    Args:
        file_path: Path to the document
        style_name: Name for the style
        font_name: Font name
        font_size: Font size in points
        bold: Apply bold formatting
        italic: Apply italic formatting
        underline: Apply underline formatting
        color: Text color (hex format, e.g., "#000000")
        alignment: Paragraph alignment (left, center, right, justify)
    """
    try:
        # Normalize path
        file_path = normalize_path(file_path)
        
        # Prepare style properties
        style_properties = {}
        
        if font_name:
            style_properties["font_name"] = font_name
        
        if font_size is not None:
            style_properties["font_size"] = font_size
        
        if bold:
            style_properties["bold"] = True
        
        if italic:
            style_properties["italic"] = True
        
        if underline:
            style_properties["underline"] = True
        
        if color:
            style_properties["color"] = color
        
        if alignment:
            style_properties["alignment"] = alignment
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "create_custom_style",
            "file_path": file_path,
            "style_name": style_name,
            "style_properties": style_properties
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in create_custom_style: {str(e)}")
        return f"Failed to create custom style: {str(e)}"

@mcp.tool()
async def delete_paragraph(file_path: str, paragraph_index: int) -> str:
    """
    Delete a paragraph at the given index.
    
    Args:
        file_path: Path to the document
        paragraph_index: Index of the paragraph to delete (0 = first paragraph)
    """
    try:
        # Normalize path
        file_path = normalize_path(file_path)
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "delete_paragraph",
            "file_path": file_path,
            "paragraph_index": paragraph_index
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in delete_paragraph: {str(e)}")
        return f"Failed to delete paragraph: {str(e)}"

@mcp.tool()
async def apply_document_style(file_path: str, font_name: Optional[str] = None, 
                              font_size: Optional[float] = None, color: Optional[str] = None,
                              alignment: Optional[str] = None) -> str:
    """
    Apply consistent formatting throughout the document.
    
    Args:
        file_path: Path to the document
        font_name: Font name
        font_size: Font size in points
        color: Text color (hex format, e.g., "#000000")
        alignment: Alignment (left, center, right, justify)
    """
    try:
        # Normalize path
        file_path = normalize_path(file_path)
        
        # Prepare style
        style = {}
        
        if font_name:
            style["font_name"] = font_name
        
        if font_size:
            style["font_size"] = font_size
        
        if color:
            style["color"] = color
        
        if alignment:
            style["alignment"] = alignment
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "apply_document_style",
            "file_path": file_path,
            "style": style
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in apply_document_style: {str(e)}")
        return f"Failed to apply document style: {str(e)}"

# Impress Presentation functions
@mcp.tool()
async def create_blank_presentation(
    filename: str,
    title: str = "",
    author: str = "",
    subject: str = "",
    keywords: str = "",
) -> str:
    """
    Create a new LibreOffice Impress presentation.

    Args:
        filename: Name of the presentation
        title: Presentation title metadata
        author: Presentation author metadata
        subject: Presentation subject metadata
        keywords: Presentation keywords metadata (comma-separated)
    """
    try:
        # Check for supported file extensions
        if not filename.lower().endswith(impress_extensions):
            filename += ".odp"

        # If filename doesn't include a path, add default path
        if os.path.basename(filename) == filename:
            save_path = get_default_document_path(filename)
        else:
            save_path = filename

        # Normalize path
        save_path = normalize_path(save_path)
        print(f"Creating document at: {save_path}")

        # Prepare metadata if provided
        metadata = {}
        if title:
            metadata["Title"] = title
        if author:
            metadata["Author"] = author
        if subject:
            metadata["Subject"] = subject
        if keywords:
            # Split by comma and strip whitespace
            keywords = [k.strip() for k in keywords.split(",")]
            metadata["Keywords"] = keywords

        logging.info(save_path)
        logging.info(metadata)

        # Send command to helper
        response = call_libreoffice_helper(
            {
                "action": "create_document",
                "doc_type": "impress",
                "file_path": save_path,
                "metadata": metadata,
            }
        )

        logging.info(response["status"])
        logging.info(response["message"])

        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"

    except Exception as e:
        logging.error(f"Error in create_document: {str(e)}")
        return f"Failed to create document: {str(e)}"

@mcp.tool()
async def read_presentation(file_path: str) -> str:
    """
    Open and read the text of an Impress presentation.
    
    Args:
        file_path: Path to the presentation
    """
    try:
        # Normalize path
        file_path = normalize_path(file_path)
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "open_presentation",
            "file_path": file_path
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in open_presentation: {str(e)}")
        return f"Failed to open presentation: {str(e)}"

@mcp.tool()
async def add_slide(file_path: str, slide_index: Optional[int] = None, title: Optional[str] = None, content: Optional[str] = None, layout = 'TitleContent') -> str:
    """
    Add a new slide to an Impress presentation using the built-in layout.
    Args:
        file_path: Path to the presentation file.
        slide_index: Index to insert the slide (None = append at end).
        title: Optional title text for the slide.
        content: Optional content text for the slide.
    """

    if not file_path.endswith(impress_extensions):
        return "Error: file_path is not a presentation."

    file_path = normalize_path(file_path)
    response = call_libreoffice_helper({
        "action": "add_slide",
        "file_path": file_path,
        "slide_index": slide_index,
        "title": title,
        "content": content,
        "layout": layout
    })
    if response["status"] == "success":
        return response["message"]
    else:
        return f"Error: {response['message']}"

@mcp.tool()
async def apply_presentation_template(file_path: str, template_name: str) -> str:
    """
    Apply a built-in LibreOffice presentation template/master slide to a presentation.
    
    Args:
        file_path: Path to the presentation file
        template_name: Name of the built-in template to apply (can be exact name, partial match, or numeric index)
    """
    try:
        # Validate file extension
        if not file_path.lower().endswith(impress_extensions):
            return "Error: file_path is not a presentation file."
        
        # Normalize path
        file_path = normalize_path(file_path)
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "apply_presentation_template",
            "file_path": file_path,
            "template_name": template_name
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Error: {response['message']}"
    except Exception as e:
        print(f"Error in apply_presentation_template: {str(e)}")
        return f"Failed to apply presentation template: {str(e)}"

# Resource for document access
@mcp.resource("libreoffice:{path}")
async def document_resource(path: str) -> str:
    """
    Resource for accessing LibreOffice document content.
    
    Args:
        path: Path to the document
    
    Returns:
        Document content as text
    """
    try:
        normalized_path = normalize_path(path)
        
        # Send command to helper
        response = call_libreoffice_helper({
            "action": "open_text_document",
            "file_path": normalized_path
        })
        
        if response["status"] == "success":
            return response["message"]
        else:
            return f"Failed to access document resource: {response['message']}"
    except Exception as e:
        print(f"Error in document_resource: {str(e)}")
        return f"Failed to access document resource: {str(e)}"

def main():
    """
    Main entry point for the LibreOffice MCP server.
    """
    print("Starting LibreOffice MCP server with stdio transport")
    
    # Check if helper is running
    try:
        response = call_libreoffice_helper({"action": "ping"})
        if response["status"] == "error" and "Connection refused" in response["message"]:
            print("⚠️ WARNING: LibreOffice helper is not running!")
            print("Please start the helper first with:")
            print('"C:\\Program Files\\LibreOffice\\program\\python.exe" helper.py')
            print("Continuing anyway, but LibreOffice operations will fail...")
    except:
        print("Failed to check helper status")
    
    # Run the server using stdio transport
    mcp.run(transport='stdio')

if __name__ == "__main__":
    main()