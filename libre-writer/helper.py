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

def open_document(file_path, read_only=False):
    """Open a LibreOffice document and return it."""
    print(f"Opening document: {file_path} (read_only: {read_only})")
    
    # Normalize path
    file_path = normalize_path(file_path)
    if not os.path.exists(file_path):
        raise HelperError(f"Document not found: {file_path}")
    
    # Get desktop
    desktop = get_uno_desktop()
    if not desktop:
        raise HelperError("Failed to connect to LibreOffice desktop")
    
    try:
        # Open document
        file_url = uno.systemPathToFileUrl(file_path)
        
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
        
        # Content creation
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