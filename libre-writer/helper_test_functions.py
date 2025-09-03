import logging
import json
from helper_utils import managed_document, HelperError


# Functions for testing
def get_text_formatting(file_path, text_to_find):
    """Get formatting information for specific text in a document."""
    with managed_document(file_path, read_only=True) as doc:
        if hasattr(doc, "getText"):
            search = doc.createSearchDescriptor()
            search.SearchString = text_to_find
            search.SearchCaseSensitive = False

            found = doc.findFirst(search)
            if not found:
                raise HelperError(f"Text '{text_to_find}' not found in document")

            # Extract formatting properties from the found text
            formatting_info = {}

            # Character formatting
            try:
                formatting_info["font_name"] = getattr(found, "CharFontName", "Unknown")
                formatting_info["font_size"] = getattr(found, "CharHeight", 0)

                # Bold (CharWeight: 100=normal, 150=bold)
                char_weight = getattr(found, "CharWeight", 100)
                formatting_info["bold"] = char_weight >= 150

                # Italic (CharPosture: 0=none, 2=italic)
                char_posture = getattr(found, "CharPosture", 0)
                formatting_info["italic"] = char_posture == 2

                # Underline (CharUnderline: 0=none, 1=single, etc.)
                char_underline = getattr(found, "CharUnderline", 0)
                formatting_info["underline"] = char_underline > 0

                # Text color
                char_color = getattr(found, "CharColor", 0)
                formatting_info["color"] = (
                    f"#{char_color:06X}" if char_color > 0 else "#000000"
                )

                # Background color
                char_back_color = getattr(found, "CharBackColor", -1)
                if char_back_color != -1:
                    formatting_info["background_color"] = f"#{char_back_color:06X}"
                else:
                    formatting_info["background_color"] = "transparent"

            except Exception as char_error:
                logging.warning(f"Error reading character formatting: {char_error}")

            # Paragraph formatting
            try:
                # Alignment - convert enum to string
                para_adjust = getattr(found, "ParaAdjust", None)
                if para_adjust is not None:
                    # Convert the enum value to its numeric value first, then to string
                    try:
                        adjust_value = int(para_adjust)
                        alignment_map = {
                            0: "left",  # LEFT
                            1: "right",  # RIGHT
                            2: "block",  # BLOCK
                            3: "center",  # CENTER
                            4: "stretch",  # STRETCH
                        }
                        formatting_info["alignment"] = alignment_map.get(
                            adjust_value, f"unknown({adjust_value})"
                        )
                    except (ValueError, TypeError):
                        formatting_info["alignment"] = "unknown"

                # Line spacing
                para_line_spacing = getattr(found, "ParaLineSpacing", None)
                if para_line_spacing and hasattr(para_line_spacing, "Height"):
                    # Height is in percentage (100 = single spacing)
                    spacing_value = para_line_spacing.Height / 100.0
                    formatting_info["line_spacing"] = spacing_value
                else:
                    formatting_info["line_spacing"] = 1.0

                # Paragraph margins
                formatting_info["left_margin"] = getattr(found, "ParaLeftMargin", 0)
                formatting_info["right_margin"] = getattr(found, "ParaRightMargin", 0)
                formatting_info["top_margin"] = getattr(found, "ParaTopMargin", 0)
                formatting_info["bottom_margin"] = getattr(found, "ParaBottomMargin", 0)

                # First line indent
                formatting_info["first_line_indent"] = getattr(
                    found, "ParaFirstLineIndent", 0
                )

            except Exception as para_error:
                logging.warning(f"Error reading paragraph formatting: {para_error}")

            # Style information
            try:
                formatting_info["paragraph_style"] = getattr(
                    found, "ParaStyleName", "Standard"
                )
                formatting_info["character_style"] = getattr(found, "CharStyleName", "")

            except Exception as style_error:
                logging.warning(f"Error reading style information: {style_error}")

            # Additional properties
            try:
                # Strikethrough
                char_strikeout = getattr(found, "CharStrikeout", 0)
                formatting_info["strikethrough"] = char_strikeout > 0

                # Superscript/Subscript
                char_escapement = getattr(found, "CharEscapement", 0)
                if char_escapement > 0:
                    formatting_info["script"] = "superscript"
                elif char_escapement < 0:
                    formatting_info["script"] = "subscript"
                else:
                    formatting_info["script"] = "normal"

                # Font style (italic alternative) - convert to numeric value
                font_slant = getattr(found, "CharPosture", 0)
                try:
                    formatting_info["font_slant"] = int(font_slant)
                except (ValueError, TypeError):
                    formatting_info["font_slant"] = 0

                # Character scaling
                formatting_info["char_scale_width"] = getattr(
                    found, "CharScaleWidth", 100
                )

            except Exception as additional_error:
                logging.warning(
                    f"Error reading additional formatting: {additional_error}"
                )

            # Count occurrences
            occurrence_count = 0
            current_found = found
            while current_found:
                occurrence_count += 1
                current_found = doc.findNext(current_found.End, search)

            formatting_info["occurrences_found"] = occurrence_count
            formatting_info["search_text"] = text_to_find

            return json.dumps(formatting_info, indent=2)

        else:
            raise HelperError("Document does not support text formatting retrieval")


def get_table_info(file_path, table_index=0):
    """Get detailed information about a table in a document."""
    with managed_document(file_path, read_only=True) as doc:
        if not hasattr(doc, "getTextTables"):
            raise HelperError("Document does not support tables")

        tables = doc.getTextTables()
        table_count = tables.getCount()

        if table_count == 0:
            raise HelperError("No tables found in document")

        if table_index < 0 or table_index >= table_count:
            raise HelperError(
                f"Table index {table_index} is out of range (document has {table_count} tables)"
            )

        table = tables.getByIndex(table_index)

        # Get table dimensions
        rows = table.getRows()
        columns = table.getColumns()
        row_count = rows.getCount()
        column_count = columns.getCount()

        # Extract table data
        table_data = []
        for row_idx in range(row_count):
            row_data = []
            for col_idx in range(column_count):
                try:
                    cell_name = chr(65 + col_idx) + str(row_idx + 1)  # A1, B1, etc.
                    cell = table.getCellByName(cell_name)
                    cell_text = (
                        cell.getText().getString() if hasattr(cell, "getText") else ""
                    )
                    row_data.append(cell_text)
                except Exception as cell_error:
                    logging.warning(f"Error reading cell: {cell_error}")
                    row_data.append("")
            table_data.append(row_data)

        # Get table formatting information
        table_info = {
            "table_index": table_index,
            "rows": row_count,
            "columns": column_count,
            "data": table_data,
            "total_tables": table_count,
        }

        # Try to get additional table properties
        try:
            if hasattr(table, "TableBorder2"):
                border_info = table.TableBorder2
                table_info["has_borders"] = (
                    border_info.TopLine.LineWidth > 0
                    or border_info.BottomLine.LineWidth > 0
                    or border_info.LeftLine.LineWidth > 0
                    or border_info.RightLine.LineWidth > 0
                )
            else:
                table_info["has_borders"] = False

            if hasattr(table, "BackColor"):
                table_info["background_color"] = (
                    f"#{table.BackColor:06X}"
                    if table.BackColor != -1
                    else "transparent"
                )

            if hasattr(table, "Width"):
                table_info["width"] = table.Width

        except Exception as prop_error:
            logging.warning(f"Error reading table properties: {prop_error}")

        return json.dumps(table_info, indent=2)


def has_image(file_path):
    """Check if a document contains at least one image and get dimensions of the first image."""
    with managed_document(file_path, read_only=True) as doc:
        if not hasattr(doc, "getGraphicObjects"):
            raise HelperError("Document does not support graphics/images")

        graphic_objects = doc.getGraphicObjects()
        image_count = graphic_objects.getCount()

        result = {"has_image": image_count > 0, "image_count": image_count}

        # If there's at least one image, get dimensions of the first one
        if image_count > 0:
            try:
                first_image = graphic_objects.getByIndex(0)
                if hasattr(first_image, "Size"):
                    size = first_image.Size
                    result["first_image_width"] = size.Width
                    result["first_image_height"] = size.Height
                    # Also provide dimensions in mm for convenience
                    result["first_image_width_mm"] = size.Width / 100.0
                    result["first_image_height_mm"] = size.Height / 100.0
                else:
                    result["first_image_width"] = None
                    result["first_image_height"] = None
            except Exception as e:
                logging.warning(f"Error reading first image dimensions: {e}")
                result["first_image_width"] = None
                result["first_image_height"] = None

        return json.dumps(result, indent=2)


def get_page_break_info(file_path):
    """Get information about page breaks in a document."""
    with managed_document(file_path, read_only=True) as doc:
        if not hasattr(doc, "getText"):
            raise HelperError("Document does not support text content")

        text = doc.getText()

        # Method 1: Check paragraphs for BreakType property
        text_enum = text.createEnumeration()
        page_break_count = 0
        paragraph_details = []

        paragraph_index = 0
        while text_enum.hasMoreElements():
            paragraph = text_enum.nextElement()
            paragraph_info = {
                "index": paragraph_index,
                "has_break": False,
                "break_type": None,
            }

            try:
                if hasattr(paragraph, "BreakType"):
                    break_type = paragraph.BreakType
                    paragraph_info["break_type"] = str(break_type)

                    # Handle different ways the enum might be represented
                    if hasattr(break_type, "value"):
                        break_value = break_type.value
                    else:
                        # Convert string representation to check for page break types
                        break_str = str(break_type).upper()
                        if "PAGE_BEFORE" in break_str:
                            break_value = "PAGE_BEFORE"
                        elif "PAGE_AFTER" in break_str:
                            break_value = "PAGE_AFTER"
                        elif "PAGE_BOTH" in break_str:
                            break_value = "PAGE_BOTH"
                        else:
                            break_value = "NONE"

                    paragraph_info["break_value"] = break_value

                    # Check for any page break type
                    if break_value in ["PAGE_BEFORE", "PAGE_AFTER", "PAGE_BOTH"]:
                        page_break_count += 1
                        paragraph_info["has_break"] = True

                # Also check the paragraph text content
                if hasattr(paragraph, "getString"):
                    para_text = paragraph.getString()
                    paragraph_info["text"] = (
                        para_text[:50] + "..." if len(para_text) > 50 else para_text
                    )
                    paragraph_info["length"] = len(para_text)

            except Exception as e:
                paragraph_info["error"] = str(e)
                logging.debug(f"Error checking paragraph {paragraph_index}: {e}")

            paragraph_details.append(paragraph_info)
            paragraph_index += 1

        # Method 2: Check for manual page breaks (form feed characters)
        text_content = text.getString()
        manual_breaks = text_content.count("\x0c")  # Form feed character

        # Method 3: Try to get actual page count from view
        estimated_pages = None
        try:
            if hasattr(doc, "getCurrentController"):
                controller = doc.getCurrentController()
                if hasattr(controller, "getPageCount"):
                    estimated_pages = controller.getPageCount()
        except Exception as e:
            logging.debug(f"Could not get page count: {e}")

        result = {
            "paragraph_page_breaks": page_break_count,
            "manual_page_breaks": manual_breaks,
            "total_page_breaks": page_break_count + manual_breaks,
            "estimated_page_count": estimated_pages,
            "total_paragraphs": len(paragraph_details),
            "paragraph_details": paragraph_details,
            "text_length": len(text_content),
            "contains_form_feed": "\x0c" in text_content,
            "debug_text_sample": text_content[:200] + "..."
            if len(text_content) > 200
            else text_content,
        }

        return json.dumps(result, indent=2)
