import azure.functions as func # type: ignore
import logging
import json
import base64
import io
import os
import tempfile
from docx import Document # type: ignore
from docx.shared import Pt, RGBColor # type: ignore
from docx.oxml import parse_xml # type: ignore
from docx.oxml.ns import nsdecls # type: ignore
import pypandoc # type: ignore

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

@app.route(route="receiveDoc")


def receiveDoc(req: func.HttpRequest) -> func.HttpResponse:
    """
    Handles incoming HTTP request with a base64-encoded DOCX file,
    processes it to merge table columns where column 2 == 'section',
    formats the merged cells, and returns the modified DOCX as base64.
    """
    try:
        req_body = req.get_json()
        input_base64 = req_body.get("base64", "")

        if not input_base64:
            return func.HttpResponse(json.dumps({"message": "No input base64 provided", "base64": None}), status_code=400)

        # Decode base64 to DOCX bytes
        doc_bytes = base64.b64decode(input_base64)

        # Process document
        modified_docx_bytes, message = process_docx(doc_bytes)

        # If processing failed
        if modified_docx_bytes is None:
            return func.HttpResponse(json.dumps({"message": message, "base64": None}), status_code=500)

        # Encode modified DOCX as base64
        output_base64 = base64.b64encode(modified_docx_bytes).decode("utf-8")

        return func.HttpResponse(json.dumps({"message": message, "base64": output_base64}), status_code=200)

    except Exception as e:
        return func.HttpResponse(json.dumps({"message": f"Unexpected error: {str(e)}", "base64": None}), status_code=500)


def process_docx(doc_bytes):
    """
    Processes the DOCX document:
    - Finds the first table
    - If column 2 contains "section", removes "section" before merging with column 1
    - Format the merged cell (bold, white text, font size 14, dark blue background)
    - Returns modified DOCX as bytes
    """
try:
        # Load the DOCX from bytes
        doc = Document(io.BytesIO(doc_bytes))

        # Process the first table in the document
        if doc.tables:
            table = doc.tables[0]  # Assume the first table is the target

            for row in table.rows:
                if len(row.cells) >= 2:
                    col1, col2 = row.cells[0], row.cells[1]

                    # Check if column 2 contains "section" (case-insensitive)
                    col2_text = col2.text.strip().lower()
                    if "section" in col2_text:
                        # Remove "section" and clean up text
                        cleaned_text = col2.text.replace("Section", "").replace("section", "").strip()

                        # Merge column 2 into column 1
                        col1.merge(col2)

                        # Preserve column 1's original text and only add cleaned text if it's not empty
                        col1_text = col1.text.strip()
                        if cleaned_text:
                            col1.text = col1_text + " " + cleaned_text
                        else:
                            col1.text = col1_text  # Keep col1 as it is if cleaned_text is empty

                        # Apply formatting only to merged rows
                        for paragraph in col1.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = Pt(14)
                                run.font.bold = True
                                run.font.color.rgb = RGBColor(255, 255, 255)  # White text

                        # Apply dark blue background
                        shading_xml = parse_xml(r'<w:shd {} w:fill="002060"/>'.format(nsdecls('w')))
                        col1._element.get_or_add_tcPr().append(shading_xml)

        # Save modified DOCX to a buffer
        output_docx = io.BytesIO()
        doc.save(output_docx)
        output_docx.seek(0)

        return output_docx.read(), "Success"

    except Exception as e:
        return None, f"Processing error: {str(e)}"