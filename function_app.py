import azure.functions as func # type: ignore
import logging
import json
import base64
import io
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
    formats the merged cells, converts the document to PDF, and returns the base64-encoded PDF.
    """
    try:
        req_body = req.get_json()
        input_base64 = req_body.get("base64", "")

        if not input_base64:
            return func.HttpResponse(json.dumps({"message": "No input base64 provided", "base64": None}), status_code=400)

        # Decode base64 to DOCX bytes
        doc_bytes = base64.b64decode(input_base64)

        # Process document
        pdf_bytes, message = process_docx(doc_bytes)

        # If processing failed
        if pdf_bytes is None:
            return func.HttpResponse(json.dumps({"message": message, "base64": None}), status_code=500)

        # Encode output PDF as base64
        output_base64 = base64.b64encode(pdf_bytes).decode("utf-8")

        return func.HttpResponse(json.dumps({"message": message, "base64": output_base64}), status_code=200)

    except Exception as e:
        return func.HttpResponse(json.dumps({"message": f"Unexpected error: {str(e)}", "base64": None}), status_code=500)


def process_docx(doc_bytes):
    """
    Processes the DOCX document:
    - Finds the first table
    - If column 2 contains "section", merge with column 1
    - Format the merged cell (bold, white text, font size 14, dark blue background)
    - Convert DOCX to PDF and return the PDF bytes
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

                    # Check if column 2 exactly matches "section" (case-insensitive)
                    if col2.text.strip().lower() == "section":
                        # Merge column 2 into column 1
                        col1.merge(col2)

                        # Preserve column 1's original text
                        col1.text = col1.text.strip()

                        # Apply formatting only to merged rows
                        for paragraph in col1.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = Pt(14)
                                run.font.bold = True
                                run.font.color.rgb = RGBColor(255, 255, 255)  # White text

                        # Apply dark blue background
                        shading_xml = parse_xml(r'<w:shd {} w:fill="002060"/>'.format(nsdecls('w')))
                        col1._element.get_or_add_tcPr().append(shading_xml)

        # Use a temporary file for conversion
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
            doc.save(tmp_docx.name)  # Save DOCX file
            tmp_docx_path = tmp_docx.name  # Get the file path

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf_path = tmp_pdf.name  # Get file path for PDF

        # Convert DOCX to PDF using pypandoc
        pypandoc.convert_file(tmp_docx_path, 'pdf', format='docx', outputfile=tmp_pdf_path)

        # Read the generated PDF
        with open(tmp_pdf_path, "rb") as pdf_file:
            pdf_bytes = pdf_file.read()

        # Cleanup temp files
        os.remove(tmp_docx_path)
        os.remove(tmp_pdf_path)

        return pdf_bytes, "Success"

    except Exception as e:
        return None, f"Processing error: {str(e)}"