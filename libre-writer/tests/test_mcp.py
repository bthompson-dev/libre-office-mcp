import os
import sys
import subprocess
import pytest
import time
import json
import socket
import tempfile
import shutil
import asyncio
from fastmcp import Client
from mcp.types import TextContent

from main import start_office, is_port_in_use, get_python_path
from libre import mcp


def start_helper():
    """Start the Office helper script with LibreOffice"""
    if not is_port_in_use(8765):
        print("Starting Office helper...", file=sys.stderr)
        current_dir = os.path.dirname(__file__)
        parent_dir = os.path.dirname(current_dir)
        helper_script = os.path.join(parent_dir, "helper.py")
        python_path = get_python_path()
        subprocess.Popen([python_path, helper_script])
        time.sleep(3)
    else:
        print("Helper script already running on port 8765", file=sys.stderr)


def send_command_to_helper(command):
    """Send a command to the LibreOffice helper and get response"""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.connect(("localhost", 8765))
            sock.send(json.dumps(command).encode("utf-8"))
            response = sock.recv(16384).decode("utf-8")
            return json.loads(response)
    except Exception as e:
        return {"status": "error", "message": str(e)}


@pytest.fixture(scope="session")
def libreoffice_server():
    start_office()
    start_helper()
    yield mcp


@pytest.fixture(scope="function")
def temp_dir():
    """Create a temporary directory for test files"""
    temp_dir = tempfile.mkdtemp()
    yield temp_dir
    # Cleanup
    shutil.rmtree(temp_dir, ignore_errors=True)


@pytest.fixture(scope="function")
def test_document(libreoffice_server, temp_dir):
    """Create a basic test document file"""

    async def create_document():
        async with Client(libreoffice_server) as client:
            # Create a document with some content
            filename = os.path.join(temp_dir, "test_document.odt")
            await client.call_tool(
                "create_blank_document",
                {"filename": filename, "title": "Test Document", "author": "pytest"},
            )
            await client.call_tool(
                "add_text",
                {
                    "file_path": filename,
                    "text": "This is a test document with sample content. This is some formatted text.",
                },
            )
            return filename

    # Run the async function and return the result
    loop = asyncio.get_event_loop()
    return loop.run_until_complete(create_document())


@pytest.fixture(scope="function")
def test_presentation(libreoffice_server, temp_dir):
    """Create a basic test presentation file"""

    async def create_presentation():
        async with Client(libreoffice_server) as client:
            # Create a presentation with some content
            filename = os.path.join(temp_dir, "test_presentation.odp")
            await client.call_tool(
                "create_blank_presentation",
                {
                    "filename": filename,
                    "title": "Test Presentation",
                    "author": "pytest",
                },
            )
            # Add a slide with content
            await client.call_tool(
                "add_slide",
                {
                    "file_path": filename,
                    "title": "Test Slide",
                    "content": "This is test slide content.",
                },
            )
            return filename

    # Run the async function and return the result
    loop = asyncio.get_event_loop()
    return loop.run_until_complete(create_presentation())


@pytest.fixture(scope="function")
def test_image(temp_dir):
    """Create a simple test image file"""
    try:
        from PIL import Image

        image_path = os.path.join(temp_dir, "test_image.png")
        # Create a simple 100x100 red image
        img = Image.new("RGB", (100, 100), color="red")
        img.save(image_path)
        return image_path
    except ImportError:
        pytest.skip("PIL not available for image tests")


# Document Management Tests


@pytest.mark.asyncio
async def test_create_blank_document(libreoffice_server, temp_dir):
    """Test creating a blank Writer document"""
    async with Client(libreoffice_server) as client:
        filename = os.path.join(temp_dir, "test_document.odt")
        result = await client.call_tool(
            "create_blank_document",
            {
                "filename": filename,
                "title": "Test Document",
                "author": "pytest",
                "subject": "Testing",
                "keywords": "test, pytest, document",
            },
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "Successfully created text document" in result_content.text
        assert os.path.exists(filename)


@pytest.mark.asyncio
async def test_read_text_document(libreoffice_server, test_document):
    """Test reading a text document"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "read_text_document", {"file_path": test_document}
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "This is a test document with sample content." in result_content.text


@pytest.mark.asyncio
async def test_get_document_properties(libreoffice_server, test_document):
    """Test getting document properties"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "get_document_properties", {"file_path": test_document}
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert (
            "Test Document" in result_content.text
            or "properties" in result_content.text.lower()
        )


@pytest.mark.asyncio
async def test_list_documents(libreoffice_server, temp_dir):
    """Test listing documents in a directory"""
    async with Client(libreoffice_server) as client:
        # Create test documents first
        filename = os.path.join(temp_dir, "list_test.odt")
        filename2 = os.path.join(temp_dir, "list_test2.odt")

        await client.call_tool("create_blank_document", {"filename": filename})
        await client.call_tool("create_blank_document", {"filename": filename2})

        result = await client.call_tool("list_documents", {"directory": temp_dir})

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert (
            "list_test.odt" in result_content.text
            and "list_test2.odt" in result_content.text
        )


@pytest.mark.asyncio
async def test_copy_document(libreoffice_server, test_document, temp_dir):
    """Test copying a document"""
    async with Client(libreoffice_server) as client:
        target_file = os.path.join(temp_dir, "copied_document.odt")

        # Copy the test document
        result = await client.call_tool(
            "copy_document", {"source_path": test_document, "target_path": target_file}
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "copied" in result_content.text.lower()
        assert os.path.exists(target_file)

        # Check the content has also been copied
        read_result = await client.call_tool(
            "read_text_document", {"file_path": target_file}
        )

        read_result_content = read_result.content[0]
        assert isinstance(read_result_content, TextContent)
        assert (
            "This is a test document with sample content." in read_result_content.text
        )


# Content Creation Tests


@pytest.mark.asyncio
async def test_add_text(libreoffice_server, test_document):
    """Test adding text to a document using the test_document fixture"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "add_text",
            {
                "file_path": test_document,
                "text": " Additional text added!",
                "position": "end",
            },
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "text added" in result_content.text.lower()

        # Check the content has been added
        read_result = await client.call_tool(
            "read_text_document", {"file_path": test_document}
        )

        read_result_content = read_result.content[0]
        assert isinstance(read_result_content, TextContent)
        assert "Additional text added!" in read_result_content.text


@pytest.mark.asyncio
async def test_add_heading(libreoffice_server, test_document):
    """Test adding a heading to a document using the test_document fixture"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "add_heading", {"file_path": test_document, "text": "Chapter 1", "level": 1}
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "heading" in result_content.text.lower()

        # Check the content has been added
        formatting_result = send_command_to_helper(
            {
                "action": "get_text_formatting",
                "file_path": test_document,
                "text_to_find": "Chapter 1",
            }
        )

        if formatting_result["status"] == "success":
            formatting_data = json.loads(formatting_result["message"])

            assert formatting_data["occurrences_found"] >= 1
            assert formatting_data["paragraph_style"] == "Heading 1"
        else:
            assert False, "Failed to get formatting data"


@pytest.mark.asyncio
async def test_add_paragraph(libreoffice_server, test_document):
    """Test adding a paragraph to a document using the test_document fixture"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "add_paragraph",
            {
                "file_path": test_document,
                "text": "This is a test paragraph.",
                "alignment": "center",
            },
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "paragraph" in result_content.text.lower()

        # Check the content has been added
        formatting_result = send_command_to_helper(
            {
                "action": "get_text_formatting",
                "file_path": test_document,
                "text_to_find": "This is a test paragraph.",
            }
        )

        if formatting_result["status"] == "success":
            formatting_data = json.loads(formatting_result["message"])

            assert formatting_data["occurrences_found"] >= 1
            assert formatting_data["paragraph_style"] == "Standard"
            assert formatting_data["alignment"] == "center"
        else:
            assert False, "Failed to get formatting data"


@pytest.mark.asyncio
async def test_add_table(libreoffice_server, test_document):
    """Test adding a table to a document using the test_document fixture"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "add_table",
            {
                "file_path": test_document,
                "rows": 3,
                "columns": 2,
                "data": [
                    ["Header 1", "Header 2"],
                    ["Data 1", "Data 2"],
                    ["Data 3", "Data 4"],
                ],
                "header_row": True,
            },
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert (
            "table" in result_content.text.lower()
            or "success" in result_content.text.lower()
        )

        # Check the table has been added
        table_result = send_command_to_helper(
            {
                "action": "get_table_info",
                "file_path": test_document,
            }
        )

        if table_result["status"] == "success":
            table_data = json.loads(table_result["message"])

            assert table_data["rows"] == 3
            assert table_data["columns"] == 2
        else:
            assert False, "Failed to get table data"


@pytest.mark.asyncio
async def test_insert_image(libreoffice_server, test_document, test_image):
    """Test inserting an image into a document using the test_document fixture"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "insert_image",
            {
                "file_path": test_document,
                "image_path": test_image,
                "width": 5000,  # 50mm in 100ths of mm
                "height": 5000,
            },
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "image" in result_content.text.lower()

        # Check the image has been added
        image_result = send_command_to_helper(
            {
                "action": "has_image",
                "file_path": test_document,
            }
        )

        if image_result["status"] == "success":
            image_data = json.loads(image_result["message"])

            assert image_data["image_count"] > 0
            assert (image_data["first_image_width"] / 5000) - 1 < 0.01
            assert (image_data["first_image_height"] / 5000) - 1 < 0.01
        else:
            assert False, "Failed to get image data"


@pytest.mark.asyncio
async def test_insert_page_break(libreoffice_server, test_document):
    """Test inserting a page break using the test_document fixture"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "insert_page_break", {"file_path": test_document}
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "page break" in result_content.text.lower()

        # Check the page break has been added
        page_break_result = send_command_to_helper(
            {
                "action": "get_page_break_info",
                "file_path": test_document,
            }
        )

        if page_break_result["status"] == "success":
            page_break_data = json.loads(page_break_result["message"])

            print(page_break_data)
            assert page_break_data["total_page_breaks"] > 0
        else:
            assert False, "Failed to get page break data"


# Text Formatting Tests


@pytest.mark.asyncio
async def test_format_text(libreoffice_server, test_document):
    """Test formatting specific text in a document using the test_document fixture"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "format_text",
            {
                "file_path": test_document,
                "text_to_find": "This is some formatted text.",
                "bold": True,
                "italic": True,
                "color": "#FF0000",
            },
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "format" in result_content.text.lower()

        # Check the content has been formatted
        formatting_result = send_command_to_helper(
            {
                "action": "get_text_formatting",
                "file_path": test_document,
                "text_to_find": "This is some formatted text.",
            }
        )

        if formatting_result["status"] == "success":
            formatting_data = json.loads(formatting_result["message"])

            print(formatting_data)
            assert formatting_data["occurrences_found"] >= 1
            assert formatting_data["color"] == "#FF0000"
            assert formatting_data["bold"]
            assert formatting_data["italic"]
        else:
            assert False, "Failed to get formatting data"


@pytest.mark.asyncio
async def test_search_replace_text(libreoffice_server, test_document):
    """Test search and replace functionality using the test_document fixture"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "search_replace_text",
            {
                "file_path": test_document,
                "search_text": "sample",
                "replace_text": "modified",
            },
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "replace" in result_content.text.lower()

        result = await client.call_tool(
            "read_text_document", {"file_path": test_document}
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "modified" in result_content.text


@pytest.mark.asyncio
async def test_delete_text(libreoffice_server, test_document):
    """Test deleting specific text from a document using the test_document fixture"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "delete_text",
            {"file_path": test_document, "text_to_delete": "sample content"},
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "replaced" in result_content.text.lower()

        read_result = await client.call_tool(
            "read_text_document", {"file_path": test_document}
        )

        read_result_content = read_result.content[0]
        assert isinstance(read_result_content, TextContent)
        assert "sample content" not in read_result_content.text


# Table Formatting Tests


@pytest.mark.asyncio
async def test_format_table(libreoffice_server, test_document):
    """Test formatting a table using the test_document fixture"""
    async with Client(libreoffice_server) as client:
        # Add a table first
        await client.call_tool(
            "add_table", {"file_path": test_document, "rows": 2, "columns": 2}
        )

        border_width = 5

        result = await client.call_tool(
            "format_table",
            {
                "file_path": test_document,
                "table_index": 0,
                "border_width": border_width,
                "background_color": "#F0F0F0",
                "header_row": True,
            },
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "formatted" in result_content.text.lower()

        # Check the table has been added
        table_result = send_command_to_helper(
            {
                "action": "get_table_info",
                "file_path": test_document,
            }
        )

        if table_result["status"] == "success":
            table_data = json.loads(table_result["message"])
            avg_border_width = sum(table_data["border_widths"].values()) / 40
            print(table_data)
            assert table_data["background_color"] == "#F0F0F0"
            assert abs(avg_border_width - border_width) <= border_width * 0.1
        else:
            assert False, "Failed to get table data"


# Advanced Document Manipulation Tests


@pytest.mark.asyncio
async def test_delete_paragraph(libreoffice_server, test_document):
    """Test deleting a paragraph by index"""
    async with Client(libreoffice_server) as client:
        # Add some paragraphs first
        await client.call_tool(
            "add_paragraph", {"file_path": test_document, "text": "First paragraph"}
        )
        await client.call_tool(
            "add_paragraph",
            {"file_path": test_document, "text": "Second paragraph to delete"},
        )

        result = await client.call_tool(
            "delete_paragraph", {"file_path": test_document, "paragraph_index": 1}
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "delete" in result_content.text.lower()

        read_result = await client.call_tool(
            "read_text_document", {"file_path": test_document}
        )

        read_result_content = read_result.content[0]
        assert isinstance(read_result_content, TextContent)
        assert "Second paragraph" not in read_result_content.text


@pytest.mark.asyncio
async def test_apply_document_style(libreoffice_server, test_document):
    """Test applying document-wide styling using the test_document fixture"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "apply_document_style",
            {
                "file_path": test_document,
                "font_name": "Arial",
                "font_size": 12,
                "color": "#000080",
                "alignment": "justify",
            },
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "style" in result_content.text.lower()

        # Check the content has been formatted
        formatting_result = send_command_to_helper(
            {
                "action": "get_text_formatting",
                "file_path": test_document,
                "text_to_find": "This is a test document",
            }
        )

        if formatting_result["status"] == "success":
            formatting_data = json.loads(formatting_result["message"])

            assert formatting_data["occurrences_found"] >= 1
            assert formatting_data["color"] == "#000080"
            assert formatting_data["font_name"] == "Arial"
            assert formatting_data["alignment"] == "block"
        else:
            assert False, "Failed to get formatting data"


# Presentation Tests


@pytest.mark.asyncio
async def test_create_blank_presentation(libreoffice_server, temp_dir):
    """Test creating a blank Impress presentation"""
    async with Client(libreoffice_server) as client:
        filename = os.path.join(temp_dir, "test_presentation.odp")
        result = await client.call_tool(
            "create_blank_presentation",
            {"filename": filename, "title": "Test Presentation", "author": "pytest"},
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "presentation" in result_content.text.lower()
        assert os.path.exists(filename)


@pytest.mark.asyncio
async def test_read_presentation(libreoffice_server, test_presentation):
    """Test reading a presentation using the test_presentation fixture"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "read_presentation", {"file_path": test_presentation}
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "This is test slide content." in result_content.text


@pytest.mark.asyncio
async def test_add_slide(libreoffice_server, test_presentation):
    """Test adding a slide to a presentation using the test_presentation fixture"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "add_slide",
            {
                "file_path": test_presentation,
                "title": "New Slide",
                "content": "This is added slide content.",
            },
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "slide added" in result_content.text.lower()

        read_result = await client.call_tool(
            "read_presentation", {"file_path": test_presentation}
        )

        read_result_content = read_result.content[0]
        assert isinstance(read_result_content, TextContent)
        assert "This is added slide content." in read_result_content.text


@pytest.mark.asyncio
async def test_edit_slide_content(libreoffice_server, test_presentation):
    """Test editing slide content using the test_presentation fixture"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "edit_slide_content",
            {
                "file_path": test_presentation,
                "slide_index": 0,
                "new_content": "Updated slide content",
            },
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "edited" in result_content.text.lower()

        read_result = await client.call_tool(
            "read_presentation", {"file_path": test_presentation}
        )

        read_result_content = read_result.content[0]
        assert isinstance(read_result_content, TextContent)
        assert "Updated slide content" in read_result_content.text


@pytest.mark.asyncio
async def test_edit_slide_title(libreoffice_server, test_presentation):
    """Test editing slide title using the test_presentation fixture"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "edit_slide_title",
            {
                "file_path": test_presentation,
                "slide_index": 0,
                "new_title": "Updated Title",
            },
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "edited" in result_content.text.lower()

        read_result = await client.call_tool(
            "read_presentation", {"file_path": test_presentation}
        )

        read_result_content = read_result.content[0]
        assert isinstance(read_result_content, TextContent)
        assert "Updated Title" in read_result_content.text


@pytest.mark.asyncio
async def test_delete_slide(libreoffice_server, test_presentation):
    """Test deleting a slide using the test_presentation fixture"""
    async with Client(libreoffice_server) as client:
        # Add another slide first
        await client.call_tool(
            "add_slide",
            {
                "file_path": test_presentation,
                "title": "Slide to delete",
                "content": "Deleted content",
            },
        )

        result = await client.call_tool(
            "delete_slide", {"file_path": test_presentation, "slide_index": 2}
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "deleted" in result_content.text.lower()

        read_result = await client.call_tool(
            "read_presentation", {"file_path": test_presentation}
        )

        read_result_content = read_result.content[0]
        assert isinstance(read_result_content, TextContent)
        assert "Deleted content" not in read_result_content.text


@pytest.mark.asyncio
async def test_apply_presentation_template(libreoffice_server, test_presentation):
    """Test applying a presentation template using the test_presentation fixture"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "apply_presentation_template",
            {"file_path": test_presentation, "template_name": "Beehive"},
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "applied template" in result_content.text.lower()

        template_result = send_command_to_helper(
            {
                "action": "get_presentation_template_info",
                "file_path": test_presentation,
            }
        )

        if template_result["status"] == "success":
            template = template_result["message"]

            assert "Beehive" in template
        else:
            assert False, "Failed to get template data"


@pytest.mark.asyncio
async def test_format_slide_content(libreoffice_server, test_presentation):
    """Test formatting slide content using the test_presentation fixture"""
    async with Client(libreoffice_server) as client:
        slide_index = 1

        result = await client.call_tool(
            "format_slide_content",
            {
                "file_path": test_presentation,
                "slide_index": slide_index,
                "font_name": "Arial",
                "font_size": 18,
                "bold": True,
                "color": "#FF0000",
            },
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "formatted" in result_content.text.lower()

        formatting_result = send_command_to_helper(
            {
                "action": "get_presentation_text_formatting",
                "file_path": test_presentation,
                "text_to_find": "This is test slide content.",
            }
        )

        if formatting_result["status"] == "success":
            formatting_data = json.loads(formatting_result["message"])

            print(formatting_data)
            assert formatting_data["found_on_slide"] == slide_index
            assert formatting_data["color"] == "#FF0000"
            assert formatting_data["bold"]
            assert formatting_data["font_size"] == 18
        else:
            assert False, "Failed to get template data"


@pytest.mark.asyncio
async def test_format_slide_title(libreoffice_server, test_presentation):
    """Test formatting slide title using the test_presentation fixture"""
    async with Client(libreoffice_server) as client:
        slide_index = 1

        result = await client.call_tool(
            "format_slide_title",
            {
                "file_path": test_presentation,
                "slide_index": slide_index,
                "font_name": "Times New Roman",
                "font_size": 24,
                "underline": True,
                "alignment": "center",
            },
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "formatted" in result_content.text.lower()

        formatting_result = send_command_to_helper(
            {
                "action": "get_presentation_text_formatting",
                "file_path": test_presentation,
                "text_to_find": "Test Slide",
            }
        )

        if formatting_result["status"] == "success":
            formatting_data = json.loads(formatting_result["message"])

            assert formatting_data["alignment"] == "center"
            assert formatting_data["found_on_slide"] == slide_index
            assert formatting_data["font_name"] == "Times New Roman"
            assert formatting_data["underline"]
            assert formatting_data["font_size"] == 24
        else:
            assert False, "Failed to get template data"


@pytest.mark.asyncio
async def test_insert_slide_image(libreoffice_server, test_presentation, test_image):
    """Test inserting an image into a slide using the test_presentation fixture"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "insert_slide_image",
            {
                "file_path": test_presentation,
                "slide_index": 0,
                "image_path": test_image,
                "max_width": 10000,  # 100mm in 1/100mm units
                "max_height": 10000,
            },
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert "inserted image" in result_content.text.lower()

        image_result = send_command_to_helper(
            {
                "action": "get_slide_image_info",
                "file_path": test_presentation,
                "slide_index": 0,
            }
        )

        if image_result["status"] == "success":
            image_data = json.loads(image_result["message"])
            image = image_data["images"][0]

            assert image_data["has_images"]
            assert image["is_centered_horizontally"]
            assert image["is_centered_vertically"]
            
        else:
            assert False, "Failed to get image data"


# Error Handling Tests


@pytest.mark.asyncio
async def test_invalid_file_path(libreoffice_server):
    """Test handling of invalid file paths"""
    async with Client(libreoffice_server) as client:
        result = await client.call_tool(
            "read_text_document", {"file_path": "/nonexistent/path/file.odt"}
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert (
            "error" in result_content.text.lower()
            or "not found" in result_content.text.lower()
        )


@pytest.mark.asyncio
async def test_invalid_heading_level(libreoffice_server, temp_dir):
    """Test handling of invalid heading levels"""
    async with Client(libreoffice_server) as client:
        filename = os.path.join(temp_dir, "invalid_heading.odt")
        await client.call_tool("create_blank_document", {"filename": filename})

        result = await client.call_tool(
            "add_heading",
            {
                "file_path": filename,
                "text": "Invalid Heading",
                "level": 10,  # Invalid level
            },
        )

        result_content = result.content[0]
        assert isinstance(result_content, TextContent)
        assert (
            "invalid" in result_content.text.lower()
            or "error" in result_content.text.lower()
        )
