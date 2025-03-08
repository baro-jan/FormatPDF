import azure.functions as func
import json
import base64
import io
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import pypandoc

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

@app.route(route="receiveDoc")
def receiveDoc(req: func.HttpRequest) -> func.HttpResponse:
  

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

        # Save modified DOCX to a buffer
        output_docx = io.BytesIO()
        doc.save(output_docx)
        output_docx.seek(0)

        # Convert DOCX to PDF using pypandoc
        output_pdf = io.BytesIO()
        pypandoc.convert_file(output_docx, 'pdf', format='docx', outputfile=output_pdf)

        # Return PDF as bytes
        output_pdf.seek(0)
        return output_pdf.read(), "Success"

    except Exception as e:
        return None, f"Processing error: {str(e)}"