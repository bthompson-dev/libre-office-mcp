#!/usr/bin/env python
import os
import sys
import json
import time
import socket
import traceback
from datetime import datetime
import logging

log_path = os.path.join(os.path.dirname(__file__), "helper.log")
logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s")

print("Starting LibreOffice Helper Script...")

try:
    print("Importing UNO...")
    logging.info("Importing UNO...")
    import uno
    from com.sun.star.beans import PropertyValue
    from com.sun.star.text import ControlCharacter
    from com.sun.star.text.TextContentAnchorType import AS_CHARACTER
    from com.sun.star.awt import Size
    from com.sun.star.lang import Locale
    from com.sun.star.style.ParagraphAdjust import CENTER, LEFT, RIGHT, BLOCK
    from com.sun.star.style.BreakType import PAGE_BEFORE
    from com.sun.star.table import BorderLine2, TableBorder2
    from com.sun.star.table.BorderLineStyle import SOLID
    from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
    from com.sun.star.connection import NoConnectException
    print("UNO imported successfully!")
    logging.info("UNO imported successfully!")
except ImportError as e:
    print(f"UNO Import Error: {e}")
    logging.error(f"UNO Import Error: {e}")
    print("This script must be run with LibreOffice's Python.")
    logging.error("This script must be run with LibreOffice's Python.")
    sys.exit(1)

# Create a server socket
server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
server_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
server_socket.bind(('localhost', 8765))
server_socket.listen(1)

print("LibreOffice helper listening on port 8765")
logging.info("LibreOffice helper listening on port 8765")

print(f"Socket bound to localhost:8765")
logging.info("Socket bound to localhost:8765")

class HelperError(Exception):
    pass

# Helper functions

def ensure_directory_exists(file_path):
    """Ensure the directory for a file exists, creating it if necessary."""
    directory = os.path.dirname(file_path)
    if directory and not os.path.exists(directory):
        try:
            os.makedirs(directory, exist_ok=True)
            print(f"Created directory: {directory}")
        except Exception as e:
            print(f"Failed to create directory {directory}: {str(e)}")
            return False
    return True

def normalize_path(file_path):
    """Convert a relative path to an absolute path."""
    if not file_path:
        return None
    
    # If file path is already complete, return it
    if file_path.startswith(('file://', 'http://', 'https://', 'ftp://')):
        return file_path

    # Expand user directory if path starts with ~
    if file_path.startswith('~'):
        file_path = os.path.expanduser(file_path)
        
    # Make absolute if relative
    if not os.path.isabs(file_path):
        file_path = os.path.abspath(file_path)
        
    print(f"Normalized path: {file_path}")
    return file_path

def get_uno_desktop():
    """Get LibreOffice desktop object."""
    try:
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context)
        
        # Try both localhost and 127.0.0.1
        try:
            context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        except NoConnectException:
            context = resolver.resolve("uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext")
            
        desktop = context.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", context)
        return desktop
    except Exception as e:
        print(f"Failed to get UNO desktop: {str(e)}")
        print(traceback.format_exc())
        return None

def create_property_value(name, value):
    """Create a PropertyValue with given name and value."""
    prop = PropertyValue()
    prop.Name = name
    prop.Value = value
    return prop

def open_document(file_path, read_only=False):
    """Open a LibreOffice document and return it."""
    print(f"Opening document: {file_path} (read_only: {read_only})")
    
    # Normalize path
    normalized_path = normalize_path(file_path)
    
    # For URLs, don't check file existence with os.path.exists
    if not normalized_path.startswith(('file://', 'http://', 'https://', 'ftp://')):
        # It's a local path, check if it exists
        if not os.path.exists(normalized_path):
            raise HelperError(f"Document not found: {normalized_path}")
        # Convert local path to file URL
        file_url = uno.systemPathToFileUrl(normalized_path)
    else:
        # It's already a URL
        file_url = normalized_path

    # Get desktop
    desktop = get_uno_desktop()
    if not desktop:
        raise HelperError("Failed to connect to LibreOffice desktop")
    
    try:
        # Open document
        props = [
            create_property_value("Hidden", True),
            create_property_value("ReadOnly", read_only)
        ]
        
        doc = desktop.loadComponentFromURL(file_url, "_blank", 0, tuple(props))
        if not doc:
            raise HelperError(f"Failed to load document: {file_path}")
            
        return doc, "Success"
    except Exception as e:
        print(f"Error opening document: {str(e)}")
        print(traceback.format_exc())
        raise

# General functions

def create_document(doc_type, file_path, metadata=None):
    """Create a new LibreOffice document with optional metadata."""
    logging.info(f"Creating {doc_type} document at {file_path}")
    
    # Normalize path and ensure directory exists
    file_path = normalize_path(file_path)
    if not ensure_directory_exists(file_path):
        raise HelperError(f"Failed to create directory for {file_path}")
    
    # Get desktop
    desktop = get_uno_desktop()
    if not desktop:
        raise HelperError("Failed to connect to LibreOffice desktop")
    
    # Map document types
    type_map = {
        "text": "private:factory/swriter",
        "calc": "private:factory/scalc",
        "impress": "private:factory/simpress"
    }
    
    if doc_type not in type_map:
        raise HelperError(f"Invalid document type. Choose from: {list(type_map.keys())}")
    
    try:
        # Create document
        doc = desktop.loadComponentFromURL(type_map[doc_type], "_blank", 0, ())
        if not doc:
            raise HelperError(f"Failed to create {doc_type} document")
        
        # Add metadata if provided
        if metadata and hasattr(doc, "DocumentProperties"):
            doc_info = doc.DocumentProperties
            logging.info(doc_info)
            for key, value in metadata.items():
                logging.info(f"{key} {value} {type(value)}")
                if hasattr(doc_info, key):
                    setattr(doc_info, key, value)
        
        # Save document
        file_url = uno.systemPathToFileUrl(file_path)
        logging.info(f"Saving to URL: {file_url}")
        
        props = [create_property_value("Overwrite", True)]
        doc.storeToURL(file_url, tuple(props))
        doc.close(True)
        
        # Verify file was created
        time.sleep(1)  # Wait for filesystem sync
        if os.path.exists(file_path):
            return f"Successfully created {doc_type} document at: {file_path}"
        else:
            raise HelperError(f"Document creation attempted, but file not found at: {file_path}")

    except Exception as e:
        logging.error(f"Error creating document: {str(e)}")
        logging.error(traceback.format_exc())
        raise

def list_documents(directory):
    """List all documents in a directory."""
    dir_path = normalize_path(directory)
    if not os.path.exists(dir_path) or not os.path.isdir(dir_path):
        raise HelperError(f"Directory not found: {dir_path}")
    
    try:
        docs = []
        # Extensions for LibreOffice and MS Office documents
        extensions = [
            '.odt', '.ods', '.odp', '.odg',  # LibreOffice/OpenOffice
            '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx',  # MS Office
            '.rtf', '.txt', '.csv', '.pdf'  # Other common document types
        ]
        
        for file in os.listdir(dir_path):
            file_path = os.path.join(dir_path, file)
            if os.path.isfile(file_path):
                ext = os.path.splitext(file)[1].lower()
                if ext in extensions:
                    # Get file stats
                    stats = os.stat(file_path)
                    size = stats.st_size
                    
                    # Format last modified time
                    mod_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(stats.st_mtime))
                    
                    # Map extension to document type
                    doc_type = "unknown"
                    if ext in ['.odt', '.doc', '.docx', '.rtf', '.txt']:
                        doc_type = "text"
                    elif ext in ['.ods', '.xls', '.xlsx', '.csv']:
                        doc_type = "spreadsheet"
                    elif ext in ['.odp', '.ppt', '.pptx']:
                        doc_type = "presentation"
                    elif ext in ['.odg']:
                        doc_type = "drawing"
                    elif ext in ['.pdf']:
                        doc_type = "pdf"
                    
                    docs.append({
                        "name": file,
                        "path": file_path,
                        "size": size,
                        "modified": mod_time,
                        "type": doc_type,
                        "extension": ext[1:]  # Remove leading dot
                    })
        
        # Sort by name
        docs = sorted(docs, key=lambda x: x["name"])
        
        # Format as a readable string
        if not docs:
            return "No documents found in the directory."
        
        result = f"Found {len(docs)} documents in {dir_path}:\n\n"
        for doc in docs:
            size_kb = doc["size"] / 1024
            size_display = f"{size_kb:.1f} KB" if size_kb < 1024 else f"{size_kb/1024:.1f} MB"
            result += f"Name: {doc['name']}\n"
            result += f"Type: {doc['type']} ({doc['extension']})\n"
            result += f"Size: {size_display}\n"
            result += f"Modified: {doc['modified']}\n"
            result += f"Path: {doc['path']}\n"
            result += "---\n"
        
        return result
    except Exception as e:
        print(f"Error listing documents: {str(e)}")
        print(traceback.format_exc())
        raise

def copy_document(source_path, target_path):
    """Create a copy of an existing document."""
    source_path = normalize_path(source_path)
    target_path = normalize_path(target_path)
    
    if not os.path.exists(source_path):
        raise HelperError(f"Source document not found: {source_path}")
    
    if not ensure_directory_exists(target_path):
        raise HelperError(f"Failed to create directory for target: {target_path}")
    
    # First try to open and save through LibreOffice
    doc, message = open_document(source_path, read_only=True)
    if not doc:
        raise HelperError(message)
        
    try:
        # Save to new location
        target_url = uno.systemPathToFileUrl(target_path)
        props = [create_property_value("Overwrite", True)]
        doc.storeToURL(target_url, tuple(props))
        doc.close(True)
            
        if os.path.exists(target_path):
            return f"Successfully copied document to: {target_path}"
        else:
            # If LibreOffice method failed, try direct file copy
            import shutil
            shutil.copy2(source_path, target_path)
            return f"Successfully copied document to: {target_path}"
    except Exception as e:
        doc.close(True)
        # Fall back to direct file copy
        import shutil
        shutil.copy2(source_path, target_path)
        return f"Successfully copied document to: {target_path} (using fallback method)"
    
def get_document_properties(file_path):
    """Extract document properties and statistics."""
    doc, message = open_document(file_path, read_only=True)
    if not doc:
        raise HelperError(message)
    
    try:
        props = {}
        
        # Get basic document properties
        if hasattr(doc, "DocumentProperties"):
            doc_props = doc.DocumentProperties
            for prop in ["Title", "Subject", "Author", "Description", "Keywords", "ModifiedBy"]:
                if hasattr(doc_props, prop):
                    props[prop] = getattr(doc_props, prop)
            
            # Get dates
            for date_prop in ["CreationDate", "ModificationDate"]:
                if hasattr(doc_props, date_prop):
                    date_val = getattr(doc_props, date_prop)
                    if date_val:
                        props[date_prop] = date_val.isoformat() if hasattr(date_val, 'isoformat') else str(date_val)
        
        # Get document statistics
        if hasattr(doc, "WordCount") and hasattr(doc.WordCount, "getWordCount"):
            props["WordCount"] = doc.WordCount.getWordCount()
        
        if hasattr(doc, "getText"):
            text = doc.getText()
            props["CharacterCount"] = len(text.getString())
            
            # Count paragraphs
            paragraph_count = 0
            enum = text.createEnumeration()
            while enum.hasMoreElements():
                paragraph_count += 1
                enum.nextElement()
            props["ParagraphCount"] = paragraph_count
        
        doc.close(True)
        return json.dumps(props, indent=2)
    except Exception as e:
        try:
            doc.close(True)
        except:
            pass
        raise
 
# Writer functions           

def extract_text(file_path):
    """Extract text from a document."""
    doc, message = open_document(file_path, read_only=True)
    if not doc:
        raise HelperError("Document not opened successfully")
    
    try:
        if hasattr(doc, "getText"):
            text_content = doc.getText().getString()
            doc.close(True)
            return text_content
        else:
            doc.close(True)
            raise HelperError("Document does not support text extraction")
    except Exception as e:
        try:
            doc.close(True)
        except:
            pass
        raise

def add_text(file_path, text, position="end"):
    """Add text to a document."""
    doc, message = open_document(file_path)
    if not doc:
        raise HelperError(message)
    
    try:
        if hasattr(doc, "getText"):
            text_obj = doc.getText()
            
            if position == "start":
                text_obj.insertString(text_obj.getStart(), text, False)
            elif position == "cursor":
                cursor = text_obj.createTextCursor()
                text_obj.insertString(cursor, text, False)
            else:  # default to end
                text_obj.insertString(text_obj.getEnd(), text, False)
            
            # Save document
            doc.store()
            doc.close(True)
            return f"Text added to {file_path}"
        else:
            doc.close(True)
            raise HelperError("Document does not support text insertion")
    except Exception as e:
        try:
            doc.close(True)
        except:
            pass
        raise

def add_heading(file_path, text, level=1):
    """Add a heading to a document."""
    doc, message = open_document(file_path)
    if not doc:
        raise HelperError(message)
    
    try:
        if hasattr(doc, "getText"):
            text_obj = doc.getText()
            cursor = text_obj.createTextCursor()

            # Add paragraph break
            text_obj.insertControlCharacter(text_obj.getEnd(), PARAGRAPH_BREAK, False)

            text_obj.insertString(text_obj.getEnd(), text, False)
            
            # Move cursor to the added text
            cursor.gotoEnd(False)
            cursor.goLeft(len(text), True)  # Select the text
            
            # Apply heading style
            heading_style = f"Heading {level}"
            if cursor.ParaStyleName != heading_style:
                cursor.ParaStyleName = heading_style
            
            # Add paragraph break
            text_obj.insertControlCharacter(text_obj.getEnd(), PARAGRAPH_BREAK, False)
            
            # Save document
            doc.store()
            doc.close(True)
            return f"Heading added to {file_path}"
        else:
            doc.close(True)
            raise HelperError("Document does not support headings")
    except Exception as e:
        try:
            doc.close(True)
        except:
            pass
        raise

def add_paragraph(file_path, text, style=None, alignment=None):
    """Add a paragraph with optional styling."""
    doc, message = open_document(file_path)
    if not doc:
        raise HelperError(message)
    
    try:
        if hasattr(doc, "getText"):
            text_obj = doc.getText()
            cursor = text_obj.createTextCursor()
            
            # Go to the end of the document
            cursor.gotoEnd(False)
            
            # Insert the paragraph text
            text_obj.insertString(cursor, text, False)
            
            # Select the inserted text
            cursor.gotoEnd(False)
            cursor.goLeft(len(text), True)
            
            # Apply style if specified
            if style:
                try:
                    cursor.ParaStyleName = style
                except Exception as style_error:
                    raise HelperError(f"Error applying style: {style_error}")
            
            # Apply alignment if specified
            alignment_map = {
                "left": LEFT,
                "center": CENTER,
                "right": RIGHT,
                "justify": BLOCK
            }
            
            if alignment and alignment.lower() in alignment_map:
                cursor.ParaAdjust = alignment_map[alignment.lower()]
            
            # Add paragraph break
            text_obj.insertControlCharacter(text_obj.getEnd(), PARAGRAPH_BREAK, False)
            
            # Save document
            doc.store()
            doc.close(True)
            return f"Paragraph added to {file_path}"
        else:
            doc.close(True)
            raise HelperError("Document does not support paragraphs")
    except Exception as e:
        try:
            doc.close(True)
        except:
            pass
        raise

def format_text(file_path, text_to_find, format_options):
    """Format specific text in a document."""
    doc, message = open_document(file_path)
    if not doc:
        raise HelperError(message)
    
    try:
        if hasattr(doc, "getText"):
            text = doc.getText()
            document_text = text.getString()
            
            # Check if text exists in document
            if text_to_find not in document_text:
                doc.close(True)
                return f"Text '{text_to_find}' not found in document"
            
            # Manual cursor approach - more reliable
            cursor = text.createTextCursor()
            cursor.gotoStart(False)
            
            found_count = 0
            search_pos = document_text.find(text_to_find)
            
            while search_pos >= 0:
                found_count += 1
                
                # Navigate to the position
                cursor.gotoStart(False)
                cursor.goRight(search_pos, False)
                cursor.goRight(len(text_to_find), True)  # Select the text
                
                # Apply formatting
                if format_options.get("bold"):
                    cursor.CharWeight = 150
                
                if format_options.get("italic"):
                    cursor.CharPosture = 2
                
                if format_options.get("underline"):
                    cursor.CharUnderline = 1
                
                if format_options.get("color"):
                    try:
                        color = format_options["color"]
                        if isinstance(color, str) and color.startswith("#"):
                            color = int(color[1:], 16)
                        cursor.CharColor = color
                    except Exception as e:
                        raise HelperError(f"Color error: {e}")
                
                if format_options.get("font"):
                    cursor.CharFontName = format_options["font"]
                
                if format_options.get("size"):
                    cursor.CharHeight = float(format_options["size"])
                
                # Find next occurrence
                search_pos = document_text.find(text_to_find, search_pos + len(text_to_find))
            
            # Save document
            doc.store()
            doc.close(True)
            return f"Formatted {found_count} occurrences of '{text_to_find}' in {file_path}"
        else:
            doc.close(True)
            raise HelperError("Document does not support text formatting")
    except Exception as e:
        try:
            doc.close(True)
        except:
            pass
        print(f"Error in format_text: {str(e)}")
        print(traceback.format_exc())
        raise

def search_replace_text(file_path, search_text, replace_text):
    """Search and replace text throughout the document."""
    doc, message = open_document(file_path)
    if not doc:
        raise HelperError(message)
    
    try:
        if hasattr(doc, "getText"):
            text_obj = doc.getText()
            document_text = text_obj.getString()
            
            # Check if text exists in document
            if search_text not in document_text:
                doc.close(True)
                raise HelperError(f"Text '{search_text}' not found in document")
            
            # Create replace descriptor
            replace_desc = doc.createReplaceDescriptor()
            replace_desc.SearchString = search_text
            replace_desc.ReplaceString = replace_text
            replace_desc.SearchCaseSensitive = False
            replace_desc.SearchWords = False
            
            # Perform replacement
            count = doc.replaceAll(replace_desc)
            
            # Save document
            doc.store()
            doc.close(True)
            return f"Replaced {count} occurrences of '{search_text}' with '{replace_text}' in {file_path}"
        else:
            doc.close(True)
            raise HelperError("Document does not support search and replace")
    except Exception as e:
        try:
            doc.close(True)
        except:
            pass
        print(f"Error in search_replace_text: {str(e)}")
        print(traceback.format_exc())
        raise

def delete_text(file_path, text_to_delete):
    """Delete specific text from the document."""
    return search_replace_text(file_path, text_to_delete, "")

def add_table(file_path, rows, columns, data=None, header_row=False):
    """Add a table to a document."""
    doc, message = open_document(file_path)
    if not doc:
        raise HelperError(message)
    
    try:
        if hasattr(doc, "getText"):
            text = doc.getText()
            cursor = text.createTextCursor()
            cursor.gotoEnd(False)  # Move to end of document
            
            # Create table
            table = doc.createInstance("com.sun.star.text.TextTable")
            table.initialize(rows, columns)
            text.insertTextContent(cursor, table, False)
            
            # Populate table if data is provided
            if data:
                try:
                    for row_idx, row_data in enumerate(data):
                        if row_idx >= rows:
                            break
                        for col_idx, cell_value in enumerate(row_data):
                            if col_idx >= columns:
                                break
                            cell_name = chr(65 + col_idx) + str(row_idx + 1)  # A1, B1, etc.
                            cell = table.getCellByName(cell_name)
                            cell_text = cell.getText()
                            cell_text.setString(str(cell_value))
                except Exception as table_error:
                    raise HelperError(f"Error populating table: {str(table_error)}")
            
            # Format header row if requested
            if header_row and rows > 0:
                try:
                    # Format cells in first row as bold
                    for col_idx in range(columns):
                        cell_name = chr(65 + col_idx) + "1"  # A1, B1, etc.
                        cell = table.getCellByName(cell_name)
                        cursor = cell.getText().createTextCursor()
                        cursor.gotoStart(False)
                        cursor.gotoEnd(True)
                        cursor.CharWeight = 150  # Bold
                except Exception as header_error:
                    raise HelperError(f"Error formatting header row: {header_error}")
            
            # Save document
            doc.store()
            doc.close(True)
            return f"Table added to {file_path}"
        else:
            doc.close(True)
            raise HelperError("Document does not support tables")
    except Exception as e:
        try:
            doc.close(True)
        except:
            pass
        raise
    
def format_table(file_path, table_index, format_options):
    """Format a table with borders, shading, etc."""
    doc, message = open_document(file_path)
    if not doc:
        raise HelperError(message)
    
    try:
        if not hasattr(doc, "getTextTables"):
            doc.close(True)
            raise HelperError("Document does not support table formatting")
        
        tables = doc.getTextTables()
        if tables.getCount() <= table_index:
            doc.close(True)
            raise HelperError(f"Table index {table_index} is out of range (document has {tables.getCount()} tables)")
        
        table = tables.getByIndex(table_index)
        
        # Apply table formatting options
        if "border_width" in format_options:
            try:
                width = int(format_options["border_width"])
                # Create border line
                border_line = BorderLine2()
                border_line.LineWidth = width
                border_line.LineStyle = SOLID
                
                # Create table border
                table_border = TableBorder2()
                table_border.TopLine = border_line
                table_border.BottomLine = border_line
                table_border.LeftLine = border_line
                table_border.RightLine = border_line
                table_border.HorizontalLine = border_line
                table_border.VerticalLine = border_line
                
                # Apply border to table
                table.TableBorder2 = table_border
            except Exception as border_error:
                raise HelperError(f"Error applying table borders: {border_error}")
        
        if "background_color" in format_options:
            try:
                color = format_options["background_color"]
                if isinstance(color, str) and color.startswith("#"):
                    color = int(color[1:], 16)
                table.BackColor = color
            except Exception as color_error:
                raise HelperError(f"Error applying table background color: {color_error}")
        
        # Format specific rows if requested
        if "header_row" in format_options:
            try:
                if format_options["header_row"]:
                    row = table.getRows().getByIndex(0)
                    row.BackColor = 13421772  # Light gray
                
                    # Format header cells
                    for col_idx in range(table.getColumns().getCount()):
                        cell = table.getCellByPosition(col_idx, 0)
                        cursor = cell.getText().createTextCursor()
                        cursor.gotoStart(False)
                        cursor.gotoEnd(True)
                        cursor.CharWeight = 150  # Bold
                else:
                    row = table.getRows().getByIndex(0)
                    row.BackColor = 16777215  # White
                
                    # Format header cells
                    for col_idx in range(table.getColumns().getCount()):
                        cell = table.getCellByPosition(col_idx, 0)
                        cursor = cell.getText().createTextCursor()
                        cursor.gotoStart(False)
                        cursor.gotoEnd(True)
                        cursor.CharWeight = 100  # Normal
            except Exception as header_error:
                raise HelperError(f"Error formatting header row: {header_error}")
        
        # Save document
        doc.store()
        doc.close(True)
        return f"Table formatted in {file_path}"
    except Exception as e:
        try:
            doc.close(True)
        except:
            pass
        raise

def insert_image(file_path, image_path, width=None, height=None):
    """Insert an image into a document using dispatch."""
    doc, message = open_document(file_path)
    if not doc:
        raise HelperError(message)
    
    try:
        # Normalize image path
        image_path = normalize_path(image_path)
        if not os.path.exists(image_path):
            doc.close(True)
            raise HelperError (f"Image not found: {image_path}")
        
        # Get component context and necessary services
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        
        # Create dispatch helper for UNO commands
        dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
        
        # Get frame from document controller
        frame = doc.getCurrentController().getFrame()
        
        # Create properties for InsertGraphic command
        props = []
        
        # Add filename property - convert to URL format
        filename_prop = PropertyValue()
        filename_prop.Name = "FileName"
        filename_prop.Value = uno.systemPathToFileUrl(image_path)
        props.append(filename_prop)
        
        # Add AsLink property
        aslink_prop = PropertyValue()
        aslink_prop.Name = "AsLink"
        aslink_prop.Value = False  # Set to True if you want to link instead of embed
        props.append(aslink_prop)
        
        # Execute the InsertGraphic command
        dispatcher.executeDispatch(frame, ".uno:InsertGraphic", "", 0, tuple(props))
        
        # Optional: Resize the image if width/height provided
        if width is not None or height is not None:
            # Try to get the inserted image as the current selection
            current_selection = doc.getCurrentController().getSelection()
            if current_selection and current_selection.getCount() > 0:
                shape = current_selection.getByIndex(0)
                
                # Calculate new size preserving aspect ratio
                if width is not None and height is not None:
                    # Use both dimensions as provided
                    size = Size(width, height)
                    shape.setSize(size)
                elif width is not None:
                    # Maintain aspect ratio based on width
                    ratio = shape.Size.Height / shape.Size.Width
                    new_height = int(width * ratio)
                    shape.setSize(Size(width, new_height))
                elif height is not None:
                    # Maintain aspect ratio based on height
                    ratio = shape.Size.Width / shape.Size.Height
                    new_width = int(height * ratio)
                    shape.setSize(Size(new_width, height))
        
        # Save document
        doc.store()
        doc.close(True)
        return f"Image inserted into {file_path}"
        
    except Exception as e:
        try:
            doc.close(True)
        except:
            pass
        raise

def insert_page_break(file_path):
    """Insert a page break at the end of the document."""
    doc, message = open_document(file_path)
    if not doc:
        raise HelperError(message)
    
    try:
        if hasattr(doc, "getText"):
            text_obj = doc.getText()
            
            # Insert page break at the end of the document
            text_obj.insertControlCharacter(text_obj.getEnd(), ControlCharacter.PARAGRAPH_BREAK, False)
            cursor = text_obj.createTextCursor()
            cursor.gotoEnd(False)
            cursor.BreakType = PAGE_BEFORE
            
            # Save document
            doc.store()
            doc.close(True)
            return f"Page break inserted in {file_path}"
        else:
            doc.close(True)
            raise HelperError("Document does not support page breaks")
    except Exception as e:
        try:
            doc.close(True)
        except:
            pass
        raise

def create_custom_style(file_path, style_name, style_properties):
    """Create a custom paragraph style."""
    doc, message = open_document(file_path)
    if not doc:
        raise HelperError(message)
    
    try:
        # Check if document supports styles
        if not hasattr(doc, "StyleFamilies"):
            doc.close(True)
            raise HelperError("Document does not support custom styles")
        
        # Get paragraph styles
        para_styles = doc.StyleFamilies.getByName("ParagraphStyles")
        
        # Create new style or modify existing style
        style = None
        if para_styles.hasByName(style_name):
            style = para_styles.getByName(style_name)
        else:
            style = doc.createInstance("com.sun.star.style.ParagraphStyle")
            para_styles.insertByName(style_name, style)
        
        # Apply style properties
        for prop, value in style_properties.items():
            if prop == "font_name":
                style.CharFontName = value
            elif prop == "font_size":
                style.CharHeight = float(value)
            elif prop == "bold":
                style.CharWeight = 150 if value else 100
            elif prop == "italic":
                style.CharPosture = uno.getConstantByName("com.sun.star.awt.FontSlant.ITALIC") if value else uno.getConstantByName("com.sun.star.awt.FontSlant.NONE")
            elif prop == "underline":
                style.CharUnderline = 1 if value else 0
            elif prop == "color":
                if isinstance(value, str) and value.startswith("#"):
                    value = int(value[1:], 16)
                style.CharColor = value
            elif prop == "alignment":
                alignment_map = {
                    "left": LEFT,
                    "center": CENTER,
                    "right": RIGHT,
                    "justify": BLOCK
                }
                if value.lower() in alignment_map:
                    style.ParaAdjust = alignment_map[value.lower()]
        
        # Save document
        doc.store()
        doc.close(True)
        return f"Custom style '{style_name}' created/updated in {file_path}"
    except Exception as e:
        try:
            doc.close(True)
        except:
            pass
        raise

def delete_paragraph(file_path, paragraph_index):
    """Delete a paragraph at the given index."""
    doc, message = open_document(file_path)
    if not doc:
        raise HelperError(message)
    
    try:
        if hasattr(doc, "getText"):
            text = doc.getText()
            
            # Get all paragraphs
            paragraphs = []
            enum = text.createEnumeration()
            while enum.hasMoreElements():
                paragraphs.append(enum.nextElement())
            
            # Check if index is valid
            if paragraph_index < 0 or paragraph_index >= len(paragraphs):
                doc.close(True)
                raise HelperError(f"Paragraph index {paragraph_index} is out of range (document has {len(paragraphs)} paragraphs)")
            
            # Get paragraph cursor
            paragraph = paragraphs[paragraph_index]
            paragraph_cursor = text.createTextCursorByRange(paragraph)
            
            # Delete paragraph
            text.removeTextContent(paragraph)
            
            # Save document
            doc.store()
            doc.close(True)
            return f"Paragraph at index {paragraph_index} deleted from {file_path}"
        else:
            doc.close(True)
            raise HelperError("Document does not support paragraph deletion")
    except Exception as e:
        try:
            doc.close(True)
        except:
            pass
        raise

def apply_document_style(file_path, style):
    """Apply consistent formatting throughout the document."""
    doc, message = open_document(file_path)
    if not doc:
        raise HelperError(message)
    
    try:
        if not hasattr(doc, "getText"):
            doc.close(True)
            raise HelperError("Document does not support style application")
        
        # Apply styles to all paragraphs
        text = doc.getText()
        cursor = text.createTextCursor()
        cursor.gotoStart(False)
        cursor.gotoEnd(True)

        # Apply character formatting
        if "font_name" in style:
            cursor.CharFontName = style["font_name"]
        
        if "font_size" in style:
            cursor.CharHeight = float(style["font_size"])
        
        if "color" in style:
            color = style["color"]
            if isinstance(color, str) and color.startswith("#"):
                color = int(color[1:], 16)
            cursor.CharColor = color
        
        # Apply paragraph formatting
        if "alignment" in style:
            alignment_map = {
                "left": LEFT,
                "center": CENTER,
                "right": RIGHT,
                "justify": BLOCK
            }
            if style["alignment"].lower() in alignment_map:
                cursor.ParaAdjust = alignment_map[style["alignment"].lower()]
        
        # Save document
        doc.store()
        doc.close(True)
        return f"Style applied to document {file_path}"
    except Exception as e:
        try:
            doc.close(True)
        except:
            pass
        raise

# Impress functions

def extract_impress_text(file_path):
    """Extract all text from an Impress presentation (.odp)."""
    doc, message = open_document(file_path, read_only=True)
    if not doc:
        raise HelperError("Presentation not opened successfully")
    
    try:
        # Get all slides (DrawPages)
        if not hasattr(doc, "getDrawPages"):
            doc.close(True)
            raise HelperError("Document does not support slides/pages")
        
        draw_pages = doc.getDrawPages()
        slide_count = draw_pages.getCount()
        all_text = []
        
        for i in range(slide_count):
            slide = draw_pages.getByIndex(i)
            slide_texts = []
            # Iterate over all shapes on the slide
            for shape_idx in range(slide.getCount()):
                shape = slide.getByIndex(shape_idx)
                # Some shapes have getString(), some have getText()
                if hasattr(shape, "getString"):
                    text = shape.getString()
                    if text:
                        slide_texts.append(text)
                elif hasattr(shape, "getText"):
                    text_obj = shape.getText()
                    if hasattr(text_obj, "getString"):
                        text = text_obj.getString()
                        if text:
                            slide_texts.append(text)
            all_text.append(f"Slide {i+1}:\n" + "\n".join(slide_texts))
        
        doc.close(True)
        return "\n\n".join(all_text) if all_text else "No text found in presentation."
    except Exception as e:
        try:
            doc.close(True)
        except:
            pass
        raise

def add_slide(file_path, slide_index=None, title=None, content=None):
    """
    Add a new slide to an Impress presentation using a built-in layout.
    Args:
        file_path: Path to the presentation file.
        slide_index: Index to insert the slide (None = append at end).
        title: Optional title text for the slide.
        content: Optional content text for the slide.
    """
    logging.info(f"add_slide called with: file_path={file_path}, slide_index={slide_index}, title={title}, content={content}")
    
    try:
        doc, message = open_document(file_path)
        if not doc:
            error_msg = f"Failed to open document: {message}"
            logging.error(error_msg)
            raise HelperError(error_msg)

        logging.info("Document opened successfully")

        if not hasattr(doc, "getDrawPages"):
            doc.close(True)
            error_msg = "Document does not support slides/pages"
            logging.error(error_msg)
            raise HelperError(error_msg)

        draw_pages = doc.getDrawPages()
        num_slides = draw_pages.getCount()
        logging.info(f"Current number of slides: {num_slides}")
        
        # Determine where to insert the slide
        insert_index = num_slides if slide_index is None else max(0, min(slide_index, num_slides))
        logging.info(f"Inserting slide at index: {insert_index}")
        
        # Insert new slide
        draw_pages.insertNewByIndex(insert_index)
        new_slide = draw_pages.getByIndex(insert_index)
        logging.info("New slide created")
        
        # Apply slide layout first
        layout_applied = False
        try:
            layout_type = 1  # TitleContent layout
            logging.info(f"Applying layout type: {layout_type}")
            
            # Apply the layout using different methods
            if hasattr(new_slide, "setLayout"):
                new_slide.setLayout(layout_type)
                layout_applied = True
                logging.info("Layout applied using setLayout")
            elif hasattr(new_slide, "Layout"):
                new_slide.Layout = layout_type
                layout_applied = True
                logging.info("Layout applied using Layout property")
            else:
                logging.warning("No layout method found")
                
        except Exception as layout_error:
            logging.warning(f"Could not apply layout: {layout_error}")

        # Give LibreOffice time to create the placeholder shapes
        if layout_applied:
            time.sleep(0.5)
        
        # Now look for the actual placeholder shapes that were created by the layout
        title_shape = None
        content_shape = None
        
        logging.info(f"Number of shapes on slide: {new_slide.getCount()}")
        
        # Examine all shapes on the slide
        for i in range(new_slide.getCount()):
            shape = new_slide.getByIndex(i)
            shape_type = shape.getShapeType()
            logging.info(f"Shape {i}: {shape_type}")
            
            # Check if this shape has text capabilities
            if hasattr(shape, "getText"):
                try:
                    # Check for LibreOffice presentation shapes by type
                    if shape_type == "com.sun.star.presentation.TitleTextShape":
                        title_shape = shape
                        logging.info(f"  Found title shape at index {i}")
                    elif shape_type == "com.sun.star.presentation.OutlinerShape":
                        # This is a content placeholder - use the first one we find
                        if not content_shape:
                            content_shape = shape
                            logging.info(f"  Found content shape at index {i}")
                        else:
                            logging.info(f"  Found additional content shape at index {i} (ignoring)")
                    
                    # Fallback: Try to get presentation object type
                    elif hasattr(shape, "PresentationObject"):
                        pres_obj = shape.PresentationObject
                        logging.info(f"  PresentationObject: {pres_obj}")
                        
                        # Check for title placeholder
                        if pres_obj in [0, 1] and not title_shape:  # Title placeholders
                            title_shape = shape
                            logging.info(f"  Found title placeholder at shape {i}")
                        # Check for content placeholder
                        elif pres_obj in [2, 3, 4, 5] and not content_shape:  # Content placeholders
                            content_shape = shape
                            logging.info(f"  Found content placeholder at shape {i}")
                    
                    # Additional fallback: check shape name or position
                    else:
                        if hasattr(shape, "Name"):
                            shape_name = shape.Name.lower()
                            logging.info(f"  Shape name: '{shape_name}'")
                            if "title" in shape_name and not title_shape:
                                title_shape = shape
                                logging.info(f"  Found title shape by name at shape {i}")
                            elif any(keyword in shape_name for keyword in ["content", "text", "outline"]) and not content_shape:
                                content_shape = shape
                                logging.info(f"  Found content shape by name at shape {i}")
                        
                        # Position-based fallback (title usually at top)
                        if hasattr(shape, "Position") and not title_shape and not content_shape:
                            y_pos = shape.Position.Y
                            if y_pos < 5000:  # Top area - likely title
                                title_shape = shape
                                logging.info(f"  Assuming title shape by position at shape {i} (Y: {y_pos})")
                            elif y_pos > 5000 and not content_shape:  # Lower area - likely content
                                content_shape = shape
                                logging.info(f"  Assuming content shape by position at shape {i} (Y: {y_pos})")
                                
                except Exception as shape_error:
                    logging.warning(f"  Error examining shape {i}: {shape_error}")

        # If we still don't have placeholders and text was requested, create manual shapes
        if title and not title_shape:
            logging.info("Creating manual title shape")
            title_shape = doc.createInstance("com.sun.star.drawing.TextShape")
            title_shape.setSize(Size(24000, 3000))
            title_shape.setPosition(uno.createUnoStruct("com.sun.star.awt.Point"))
            title_shape.Position.X = 2000
            title_shape.Position.Y = 2000
            new_slide.add(title_shape)

        if content and not content_shape:
            logging.info("Creating manual content shape")
            content_shape = doc.createInstance("com.sun.star.drawing.TextShape")
            content_shape.setSize(Size(24000, 14000))
            content_shape.setPosition(uno.createUnoStruct("com.sun.star.awt.Point"))
            content_shape.Position.X = 2000
            content_shape.Position.Y = 6000
            new_slide.add(content_shape)

        # Set title text
        if title and title_shape:
            logging.info(f"Setting title text: {title}")
            try:
                title_text = title_shape.getText()
                title_text.setString(title)
                
                # Format title
                title_cursor = title_text.createTextCursor()
                title_cursor.gotoStart(False)
                title_cursor.gotoEnd(True)
                title_cursor.CharHeight = 28
                title_cursor.CharWeight = 150
                title_cursor.ParaAdjust = CENTER
                logging.info("Title text set and formatted")
            except Exception as title_error:
                logging.error(f"Error setting title: {title_error}")

        # Set content text
        if content and content_shape:
            logging.info(f"Setting content text: {content}")
            try:
                content_text = content_shape.getText()
                content_text.setString(content)
                
                # Format content
                content_cursor = content_text.createTextCursor()
                content_cursor.gotoStart(False)
                content_cursor.gotoEnd(True)
                content_cursor.CharHeight = 18
                content_cursor.ParaAdjust = LEFT
                logging.info("Content text set and formatted")
            except Exception as content_error:
                logging.error(f"Error setting content: {content_error}")

        # Save and close
        logging.info("Saving document...")
        doc.store()
        doc.close(True)
        
        success_msg = f"Slide added at index {insert_index} with TitleContent layout in {file_path}"
        logging.info(success_msg)
        return success_msg
        
    except Exception as e:
        error_msg = f"Error in add_slide: {str(e)}"
        logging.error(error_msg)
        logging.error(traceback.format_exc())
        try:
            if 'doc' in locals():
                doc.close(True)
        except:
            pass
        raise HelperError(error_msg)

def apply_presentation_template(file_path, template_name):
    """Apply a presentation template to an existing presentation."""
    logging.info(f"Attempting to apply template: {template_name} to {file_path}")
    
    try:
        # Try different template paths
        template_paths = [
            f"file:///C:/Program Files/LibreOffice/share/template/common/presnt/{template_name}.otp",
        ]
                
        template_doc = None
        for template_path in template_paths:
            try:
                logging.info(f"Trying template path: {template_path}")
                template_doc, template_message = open_document(template_path, read_only=True)
                if template_doc:
                    logging.info(f"Successfully loaded template from: {template_path}")
                    break
            except Exception as path_error:
                logging.info(f"Failed to load from {template_path}: {path_error}")
                continue

        if not template_doc:
            raise HelperError(f"Could not find template '{template_name}' in any of the standard locations")

        # Load target presentation
        target_doc, target_message = open_document(file_path)
        if not target_doc:
            template_doc.close(True)
            raise HelperError(f"Failed to open target presentation: {target_message}")

        success = False
        new_doc = None

        # Create new presentation from template and copy content
        try:
            logging.info("Creating new presentation from template...")
                
            # Get desktop
            desktop = get_uno_desktop()
            if not desktop:
                raise HelperError("Failed to get UNO desktop")
                
            # Create new presentation from template
            template_path = None
            for path in template_paths:
                if path.startswith("file:///"):
                    local_path = path[8:].replace("/", "\\")
                    if os.path.exists(local_path):
                        template_path = path
                        break
                
            if not template_path:
                raise HelperError("No valid template path found")
                
            # Create new document from template
            props = [
                create_property_value("AsTemplate", True),
                create_property_value("Hidden", True)
            ]
                
            new_doc = desktop.loadComponentFromURL(template_path, "_blank", 0, tuple(props))
            if not new_doc:
                raise HelperError("Failed to create new document from template")
                
            logging.info("Created new document from template")
                
            # Get slides from target and new document
            target_slides = target_doc.getDrawPages()
            new_slides = new_doc.getDrawPages()
            
            logging.info(f"Target has {target_slides.getCount()} slides")
            logging.info(f"New document has {new_slides.getCount()} slides")
                
            target_slide_count = target_slides.getCount()
            new_slide_count = new_slides.getCount()
            
            # Validation: Ensure we have slides to work with
            if target_slide_count == 0:
                raise HelperError("Target presentation has no slides")
            
            # Analyze what layouts to use based on target slides
            target_slide_layouts = []
            for i in range(target_slide_count):
                target_slide = target_slides.getByIndex(i)
                has_title = False
                has_content = False
                
                for j in range(target_slide.getCount()):
                    shape = target_slide.getByIndex(j)
                    shape_type = shape.getShapeType()
                    
                    if shape_type == "com.sun.star.presentation.TitleTextShape":
                        text = shape.getText().getString() if hasattr(shape, "getText") else ""
                        if text.strip():
                            has_title = True
                    elif shape_type == "com.sun.star.presentation.OutlinerShape":
                        text = shape.getText().getString() if hasattr(shape, "getText") else ""
                        if text.strip():
                            has_content = True
                
                # Determine what layout is needed
                if has_title and has_content:
                    needed_layout = 1  # TitleContent layout
                    logging.info(f"Target slide {i} needs TitleContent layout (has both title and content)")
                elif has_title:
                    needed_layout = 0  # Title only layout
                    logging.info(f"Target slide {i} needs Title layout (has title only)")
                else:
                    needed_layout = 1  # Default to TitleContent for safety
                    logging.info(f"Target slide {i} needs default TitleContent layout")
                
                target_slide_layouts.append(needed_layout)
            
            # Determine the template's default layout
            template_layout = None
            if new_slide_count > 0:
                try:
                    first_template_slide = new_slides.getByIndex(0)
                    if hasattr(first_template_slide, "Layout"):
                        template_layout = first_template_slide.Layout
                        logging.info(f"Template default layout detected: {template_layout}")
                    else:
                        template_layout = 1  # Default to TitleContent
                        logging.info("Could not detect template layout, defaulting to TitleContent")
                except Exception as layout_detect_error:
                    logging.warning(f"Could not detect template layout: {layout_detect_error}")
                    template_layout = 1  # Default to TitleContent layout
            
            # Add more slides to new document if needed, with appropriate layouts
            while new_slide_count < target_slide_count:
                try:
                    new_slides.insertNewByIndex(new_slide_count)
                    added_slide = new_slides.getByIndex(new_slide_count)
                    
                    # Use the layout needed for this specific slide
                    needed_layout = target_slide_layouts[new_slide_count]
                    
                    try:
                        if hasattr(added_slide, "setLayout"):
                            added_slide.setLayout(needed_layout)
                            logging.info(f"Applied layout {needed_layout} to added slide {new_slide_count}")
                        elif hasattr(added_slide, "Layout"):
                            added_slide.Layout = needed_layout
                            logging.info(f"Set layout {needed_layout} on added slide {new_slide_count}")
                        
                        # Give LibreOffice time to create the placeholder shapes
                        time.sleep(0.3)
                        
                    except Exception as layout_error:
                        logging.warning(f"Could not apply layout {needed_layout} to slide {new_slide_count}: {layout_error}")
                    
                    new_slide_count += 1
                    logging.info(f"Added slide {new_slide_count} with layout {needed_layout}")
                    
                except Exception as slide_add_error:
                    raise HelperError(f"Failed to add slide {new_slide_count}: {slide_add_error}")
            
            # Track copying success for validation
            copy_errors = []
            slides_processed = 0
            
            # Copy content from target slides to new slides
            for i in range(target_slide_count):
                try:
                    logging.info(f"Processing slide {i + 1} of {target_slide_count}")
                    
                    target_slide = target_slides.getByIndex(i)
                    new_slide = new_slides.getByIndex(i)
                    
                    # Find and categorize shapes on both slides
                    target_title_shape = None
                    target_content_shape = None
                    target_other_shapes = []
                    
                    new_title_shape = None
                    new_content_shape = None
                    
                    # Analyze target slide shapes
                    target_shape_count = target_slide.getCount()
                    logging.info(f"Target slide {i} has {target_shape_count} shapes")
                    
                    for j in range(target_shape_count):
                        try:
                            shape = target_slide.getByIndex(j)
                            shape_type = shape.getShapeType()
                            
                            if shape_type == "com.sun.star.presentation.TitleTextShape":
                                target_title_shape = shape
                                logging.info(f"Found target title shape on slide {i}")
                            elif shape_type == "com.sun.star.presentation.OutlinerShape":
                                if not target_content_shape:
                                    target_content_shape = shape
                                    logging.info(f"Found target content shape on slide {i}")
                            else:
                                target_other_shapes.append(shape)
                                logging.info(f"Found target other shape: {shape_type}")
                        except Exception as shape_error:
                            error_msg = f"Failed to analyze target shape {j} on slide {i}: {shape_error}"
                            logging.error(error_msg)
                            copy_errors.append(error_msg)
                    
                    # Analyze new slide shapes
                    new_shape_count = new_slide.getCount()
                    logging.info(f"New slide {i} has {new_shape_count} shapes")
                    
                    for j in range(new_shape_count):
                        try:
                            shape = new_slide.getByIndex(j)
                            shape_type = shape.getShapeType()
                            
                            if shape_type == "com.sun.star.presentation.TitleTextShape":
                                new_title_shape = shape
                                logging.info(f"Found new slide title shape on slide {i}")
                            elif shape_type == "com.sun.star.presentation.OutlinerShape":
                                if not new_content_shape:
                                    new_content_shape = shape
                                    logging.info(f"Found new slide content shape on slide {i}")
                        except Exception as shape_error:
                            error_msg = f"Failed to analyze new slide shape {j} on slide {i}: {shape_error}"
                            logging.error(error_msg)
                            copy_errors.append(error_msg)
                    
                    # Copy title text (critical operation)
                    if target_title_shape:
                        target_title_text = target_title_shape.getText().getString() if hasattr(target_title_shape, "getText") else ""
                        if target_title_text.strip():  # Only if there's actual text to copy
                            if not new_title_shape:
                                error_msg = f"Target slide {i} has title text but new slide has no title placeholder"
                                logging.error(error_msg)
                                copy_errors.append(error_msg)
                            else:
                                try:
                                    new_title_shape.getText().setString(target_title_text)
                                    logging.info(f"Copied title text: '{target_title_text[:50]}...'")
                                    
                                    # Verify the text was actually set
                                    verification_text = new_title_shape.getText().getString()
                                    if verification_text != target_title_text:
                                        error_msg = f"Title text verification failed on slide {i}: expected '{target_title_text}', got '{verification_text}'"
                                        logging.error(error_msg)
                                        copy_errors.append(error_msg)
                                except Exception as title_error:
                                    error_msg = f"Failed to copy title text on slide {i}: {title_error}"
                                    logging.error(error_msg)
                                    copy_errors.append(error_msg)
                        else:
                            logging.info(f"No title text to copy on slide {i}")
                    
                    # Copy content text (critical operation)
                    if target_content_shape:
                        target_content_text = target_content_shape.getText().getString() if hasattr(target_content_shape, "getText") else ""
                        if target_content_text.strip():  # Only if there's actual text to copy
                            if not new_content_shape:
                                error_msg = f"Target slide {i} has content text but new slide has no content placeholder"
                                logging.error(error_msg)
                                copy_errors.append(error_msg)
                            else:
                                try:
                                    new_content_shape.getText().setString(target_content_text)
                                    logging.info(f"Copied content text: '{target_content_text[:50]}...'")
                                    
                                    # Verify the text was actually set
                                    verification_text = new_content_shape.getText().getString()
                                    if verification_text != target_content_text:
                                        error_msg = f"Content text verification failed on slide {i}: expected '{target_content_text}', got '{verification_text}'"
                                        logging.error(error_msg)
                                        copy_errors.append(error_msg)
                                except Exception as content_error:
                                    error_msg = f"Failed to copy content text on slide {i}: {content_error}"
                                    logging.error(error_msg)
                                    copy_errors.append(error_msg)
                        else:
                            logging.info(f"No content text to copy on slide {i}")
                    
                    # Copy other shapes (non-critical, but track errors)
                    other_shapes_copied = 0
                    for k, source_shape in enumerate(target_other_shapes):
                        try:
                            # Create a new shape of the same type
                            shape_type = source_shape.getShapeType()
                            cloned_shape = new_doc.createInstance(shape_type)
                            
                            # Copy basic properties
                            if hasattr(source_shape, "Position"):
                                cloned_shape.Position = source_shape.Position
                            if hasattr(source_shape, "Size"):
                                cloned_shape.Size = source_shape.Size
                            
                            # Copy style properties
                            style_properties = [
                                "FillColor", "FillStyle", "LineColor", "LineStyle", "LineWidth"
                            ]
                            for prop in style_properties:
                                if hasattr(source_shape, prop) and hasattr(cloned_shape, prop):
                                    try:
                                        setattr(cloned_shape, prop, getattr(source_shape, prop))
                                    except:
                                        pass  # Non-critical property copy failure
                            
                            # Copy text content if it's a text shape
                            if hasattr(source_shape, "getText") and hasattr(cloned_shape, "getText"):
                                source_text = source_shape.getText().getString()
                                if source_text:
                                    cloned_shape.getText().setString(source_text)
                                    logging.info(f"Copied text to other shape: '{source_text[:30]}...'")
                            
                            # Add the cloned shape to the new slide
                            new_slide.add(cloned_shape)
                            other_shapes_copied += 1
                            logging.info(f"Successfully copied other shape {k}: {shape_type}")
                            
                        except Exception as clone_error:
                            error_msg = f"Failed to copy other shape {k} on slide {i}: {clone_error}"
                            logging.warning(error_msg)
                            copy_errors.append(error_msg)
                    
                    logging.info(f"Copied {other_shapes_copied} of {len(target_other_shapes)} other shapes on slide {i}")
                    slides_processed += 1
                    
                except Exception as slide_error:
                    error_msg = f"Critical error processing slide {i}: {slide_error}"
                    logging.error(error_msg)
                    copy_errors.append(error_msg)
            
            # Remove any extra slides from new document
            extra_slides_removed = 0
            while new_slides.getCount() > target_slide_count:
                try:
                    last_slide = new_slides.getByIndex(new_slides.getCount() - 1)
                    new_slides.remove(last_slide)
                    extra_slides_removed += 1
                    logging.info(f"Removed extra slide")
                except Exception as remove_slide_error:
                    error_msg = f"Failed to remove extra slide: {remove_slide_error}"
                    logging.error(error_msg)
                    copy_errors.append(error_msg)
                    break
            
            # CRITICAL VALIDATION: Check if all operations succeeded
            if copy_errors:
                error_summary = f"Content copying failed with {len(copy_errors)} errors. No changes will be applied to preserve data integrity."
                logging.error(error_summary)
                for error in copy_errors:
                    logging.error(f"  - {error}")
                raise HelperError(f"{error_summary} First error: {copy_errors[0]}")
            
            if slides_processed != target_slide_count:
                raise HelperError(f"Only processed {slides_processed} of {target_slide_count} slides. Aborting to prevent data loss.")
            
            # Verify final slide count matches
            if new_slides.getCount() != target_slide_count:
                raise HelperError(f"Final slide count mismatch: expected {target_slide_count}, got {new_slides.getCount()}")
            
            logging.info("All content copied successfully. Proceeding with file replacement.")
            
            # Only now that everything is verified, save new document over the target
            file_url = uno.systemPathToFileUrl(normalize_path(file_path))
            save_props = [create_property_value("Overwrite", True)]
            
            try:
                new_doc.storeToURL(file_url, tuple(save_props))
                logging.info("Successfully saved new document over target file")
            except Exception as save_error:
                raise HelperError(f"Failed to save templated document: {save_error}")
            
            success = True
            logging.info("Template applied successfully with all content preserved")
                    
        except Exception as process_error:
            logging.error(f"Template application failed: {process_error}")
            logging.error(traceback.format_exc())
            # Don't set success = True, so no changes are applied
            raise process_error

        finally:
            # Clean up documents
            try:
                if new_doc:
                    new_doc.close(True)
                    logging.info("Closed new document")
            except:
                pass
            
            try:
                target_doc.close(True)
                logging.info("Closed target document")
            except:
                pass
                
            try:
                template_doc.close(True)
                logging.info("Closed template document")
            except:
                pass
        
        if success:
            logging.info(f"Successfully applied template '{template_name}' to {file_path}")
            return f"Successfully applied template '{template_name}' to presentation with all content preserved"
        else:
            logging.warning("Template application failed - original file unchanged")
            return f"Failed to apply template '{template_name}' to presentation - original file preserved"
        
    except Exception as e:
        error_msg = f"Error applying presentation template: {str(e)}"
        logging.error(error_msg)
        logging.error(traceback.format_exc())
        
        # Clean up any open documents in case of error
        try:
            if 'template_doc' in locals() and template_doc:
                template_doc.close(True)
        except:
            pass
        try:
            if 'target_doc' in locals() and target_doc:
                target_doc.close(True)
        except:
            pass
        try:
            if 'new_doc' in locals() and new_doc:
                new_doc.close(True)
        except:
            pass
        
        raise HelperError(error_msg)

# Main command handler
def handle_command(command):
    """Process commands from the MCP server."""
    try:
        logging.info("handle_command called")
        action = command.get("action", "")
        logging.info(f"action: {action}")
        
        # Document creation and management
        if action == "create_document":
            return create_document(
                command.get("doc_type", "text"), 
                command.get("file_path", ""),
                command.get("metadata", None)
            )
        
        elif action == "open_text_document":
            return extract_text(command.get("file_path", ""))
        
        elif action == "get_document_properties":
            return get_document_properties(command.get("file_path", ""))
        
        elif action == "list_documents":
            return list_documents(command.get("directory", ""))
        
        elif action == "copy_document":
            return copy_document(
                command.get("source_path", ""),
                command.get("target_path", "")
            )
        
        # Impress presentation creation and management

        elif action == "open_presentation":
            return extract_impress_text(command.get("file_path", ""))

        elif action == "add_slide":
            return add_slide(
                command.get("file_path", ""),
                command.get("slide_index", None),
                command.get("title", None),
                command.get("content", None),
            )

        elif action == "apply_presentation_template":
            return apply_presentation_template(
                command.get("file_path", ""),
                command.get("template_name", "")
            )

        # Writer content creation
        elif action == "add_text":
            return add_text(
                command.get("file_path", ""), 
                command.get("text", ""), 
                command.get("position", "end")
            )
        
        elif action == "add_heading":
            return add_heading(
                command.get("file_path", ""),
                command.get("text", ""),
                command.get("level", 1)
            )
        
        elif action == "add_paragraph":
            return add_paragraph(
                command.get("file_path", ""),
                command.get("text", ""),
                command.get("style", None),
                command.get("alignment", None)
            )
        
        elif action == "add_table":
            return add_table(
                command.get("file_path", ""),
                command.get("rows", 2),
                command.get("columns", 2),
                command.get("data", None),
                command.get("header_row", False)
            )
        
        elif action == "insert_image":
            return insert_image(
                command.get("file_path", ""),
                command.get("image_path", ""),
                command.get("width", None),
                command.get("height", None)
            )
        
        elif action == "insert_page_break":
            return insert_page_break(command.get("file_path", ""))
        
        # Text formatting
        elif action == "format_text":
            return format_text(
                command.get("file_path", ""),  # Use file_path consistently (not filepath)
                command.get("text_to_find", ""),
                {  # Single format_options dictionary
                    "bold": command.get("bold", False),
                    "italic": command.get("italic", False),
                    "underline": command.get("underline", False),
                    "color": command.get("color", None),
                    "font": command.get("font", None),
                    "size": command.get("size", None)
                }
            )
        
        elif action == "search_replace_text":
            return search_replace_text(
                command.get("file_path", ""),
                command.get("search_text", ""),
                command.get("replace_text", "")
            )
        
        elif action == "delete_text":
            return delete_text(
                command.get("file_path", ""),
                command.get("text_to_delete", "")
            )
        
        # Table formatting
        elif action == "format_table":
            return format_table(
                command.get("file_path", ""),
                command.get("table_index", 0),
                command.get("format_options", {})
            )
        
        # Advanced document manipulation
        elif action == "create_custom_style":
            return create_custom_style(
                command.get("file_path", ""),
                command.get("style_name", "CustomStyle"),
                command.get("style_properties", {})
            )
        
        elif action == "delete_paragraph":
            return delete_paragraph(
                command.get("file_path", ""),
                command.get("paragraph_index", 0)
            )
        
        elif action == "apply_document_style":
            return apply_document_style(
                command.get("file_path", ""),
                command.get("style", {})
            )
        
        elif action == "ping":
            return "LibreOffice helper is running"
            
        else:
            return f"Unknown action: {action}"
            
    except Exception as e:
        print(f"Error handling command: {str(e)}")
        print(traceback.format_exc())
        raise


# Main server loop
print("Starting command processing loop...")
try:
    while True:
        print("Waiting for connection...")
        logging.info("Waiting for connection...")
        logging.info(server_socket)

        client_socket, address = server_socket.accept()
        print(f"Connection from {address}")
        logging.info(f"Connection from {address}")
        
        try:
            # Receive data with timeout
            client_socket.settimeout(30)
            data = client_socket.recv(16384).decode('utf-8')
            
            if not data:
                print("Empty data received, closing connection")
                client_socket.close()
                continue
                
            print(f"Received data: {data[:100]}...")
            logging.info(f"Received data: {data[:100]}...")
            
            try:
                command = json.loads(data)
                result = handle_command(command)
                
                response = {
                    "status": "success",
                    "message": result
                }
            except json.JSONDecodeError:
                response = {
                    "status": "error",
                    "message": "Invalid JSON received"
                }
            except Exception as e:
                print(f"Error processing command: {str(e)}")
                print(traceback.format_exc())
                response = {
                    "status": "error",
                    "message": f"Error: {str(e)}"
                }
                
            # Send response
            client_socket.send(json.dumps(response).encode('utf-8'))
            print("Response sent")
            logging.info("Response sent")
            
        except socket.timeout:
            print("Connection timed out")
            logging.error("Connection timed out")
            response = {
                "status": "error",
                "message": "Connection timed out"
            }
            try:
                client_socket.send(json.dumps(response).encode('utf-8'))
            except:
                pass
        except Exception as e:
            print(f"Error handling client: {str(e)}")
            logging.error(f"Error handling client: {str(e)}")
            print(traceback.format_exc())
            logging.error(traceback.format_exc())
            try:
                response = {
                    "status": "error",
                    "message": f"Server error: {str(e)}"
                }
                client_socket.send(json.dumps(response).encode('utf-8'))
            except:
                pass
        finally:
            client_socket.close()
            print("Connection closed")

except KeyboardInterrupt:
    print("Helper server shutting down...")
except Exception as e:
    print(f"Fatal error: {str(e)}")
    logging.fatal(f"Fatal error: {str(e)}")
    print(traceback.format_exc())
    logging.fatal(traceback.format_exc())
finally:
    server_socket.close()
    print("Server socket closed")