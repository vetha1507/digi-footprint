
from flask import Flask, render_template, send_from_directory, request, send_file
from PIL import Image, ExifTags
from pypdf import PdfReader, PdfWriter
from docx import Document
from pptx import Presentation
from openpyxl import load_workbook
from io import BytesIO
import mimetypes
import os

app = Flask(__name__, template_folder=".")
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/style.css")
def style():
    return send_from_directory("static", "style.css")


@app.route("/hero-image")
def hero_image():
    image_path = os.path.join(app.root_path, "static", "digi_footprint_hero_art.jpeg")
    if os.path.exists(image_path):
        return send_file(image_path)
    return"Image not found: " + image_path, 404


@app.route("/file-info")
def file_info():
    return render_template("file_info.html")

@app.route("/analyse-file", methods=["POST"])
def analyse_file():

    uploaded_file = request.files.get("file")

    if not uploaded_file or uploaded_file.filename == "":
        return "Please select a file.", 400

    filename = uploaded_file.filename
    extension = filename.rsplit(".", 1)[-1].lower()

    allowed_formats = [
        "jpg", "jpeg", "png", "pdf", "docx", "pptx", "xlsx"
    ]

    if extension not in allowed_formats:
        return "Unsupported file format.", 400

    file_data = uploaded_file.read()

    if not file_data:
        return "The selected file is empty.", 400

    file_size = f"{len(file_data) / 1024:.2f} KB"
    file_type = mimetypes.guess_type(filename)[0] or extension.upper()

    metadata = {}

    try:

        if extension in ["jpg", "jpeg", "png"]:

            with Image.open(BytesIO(file_data)) as img:

                metadata["Image Format"] = img.format
                metadata["Dimensions"] = f"{img.width} x {img.height} pixels"

                exif_data = img.getexif()

                for tag_id, value in exif_data.items():
                    tag_name = ExifTags.TAGS.get(tag_id, str(tag_id))
                    metadata[tag_name] = str(value)

                for key, value in img.info.items():
                    if isinstance(value, str) and value.strip():
                        metadata.setdefault(key, value[:500])

        elif extension == "pdf":

            reader = PdfReader(BytesIO(file_data))

            metadata["Page Count"] = len(reader.pages)

            if reader.metadata:
                for key, value in reader.metadata.items():
                    if value is not None:
                        metadata[str(key).lstrip("/")] = str(value)

        elif extension == "docx":

            document = Document(BytesIO(file_data))
            properties = document.core_properties

            fields = {
                "Title": properties.title,
                "Author": properties.author,
                "Subject": properties.subject,
                "Keywords": properties.keywords,
                "Comments": properties.comments,
                "Category": properties.category,
                "Created": properties.created,
                "Modified": properties.modified,
                "Last Modified By": properties.last_modified_by
            }

            for key, value in fields.items():
                if value:
                    metadata[key] = str(value)

        elif extension == "pptx":

            presentation = Presentation(BytesIO(file_data))
            properties = presentation.core_properties

            metadata["Slide Count"] = len(presentation.slides)

            fields = {
                "Title": properties.title,
                "Author": properties.author,
                "Subject": properties.subject,
                "Keywords": properties.keywords,
                "Comments": properties.comments,
                "Category": properties.category,
                "Created": properties.created,
                "Modified": properties.modified,
                "Last Modified By": properties.last_modified_by
            }

            for key, value in fields.items():
                if value:
                    metadata[key] = str(value)

        elif extension == "xlsx":

            workbook = load_workbook(
                BytesIO(file_data),
                read_only=True,
                data_only=False
            )

            properties = workbook.properties

            metadata["Sheet Count"] = len(workbook.sheetnames)
            metadata["Sheet Names"] = ", ".join(workbook.sheetnames[:10])

            for index, sheet in enumerate(workbook.worksheets):
                if index >= 10:
                    break

                metadata[f"{sheet.title} - Rows"] = sheet.max_row
                metadata[f"{sheet.title} - Columns"] = sheet.max_column

            fields = {
                "Title": properties.title,
                "Author": properties.creator,
                "Subject": properties.subject,
                "Keywords": properties.keywords,
                "Description": properties.description,
                "Category": properties.category,
                "Created": properties.created,
                "Modified": properties.modified,
                "Last Modified By": properties.lastModifiedBy
            }

            for key, value in fields.items():
                if value:
                    metadata[key] = str(value)

            workbook.close()

    except Exception:
        return "Unable to read this file. Please check that it is a valid supported file.", 400

    return render_template(
        "result.html",
        filename=filename,
        file_type=file_type,
        file_size=file_size,
        metadata=metadata
    )

@app.route("/risk-analysis")
def risk_analysis():
    return render_template("risk_analysis.html")

@app.route("/analyse-risk", methods=["POST"])
def analyse_risk():

    uploaded_file = request.files.get("file")

    if not uploaded_file or uploaded_file.filename == "":
        return "Please select a file.", 400

    filename = uploaded_file.filename
    extension = filename.rsplit(".", 1)[-1].lower()

    allowed_formats = [
        "jpg", "jpeg", "png", "pdf", "docx", "pptx", "xlsx"
    ]

    if extension not in allowed_formats:
        return "Unsupported file format.", 400

    file_data = uploaded_file.read()

    if not file_data:
        return "The selected file is empty.", 400

    file_size = f"{len(file_data) / 1024:.2f} KB"
    metadata = {}

    try:

        if extension in ["jpg", "jpeg", "png"]:

            with Image.open(BytesIO(file_data)) as img:

                metadata["Image Format"] = img.format
                metadata["Dimensions"] = f"{img.width} x {img.height}"

                exif_data = img.getexif()

                for tag_id, value in exif_data.items():
                    tag_name = ExifTags.TAGS.get(tag_id, str(tag_id))
                    metadata[tag_name] = str(value)

                for key, value in img.info.items():
                    if isinstance(value, str) and value.strip():
                        metadata.setdefault(key, value[:500])

        elif extension == "pdf":

            reader = PdfReader(BytesIO(file_data))

            metadata["Page Count"] = len(reader.pages)

            if reader.metadata:
                for key, value in reader.metadata.items():
                    if value is not None:
                        metadata[str(key).lstrip("/")] = str(value)

        elif extension == "docx":

            document = Document(BytesIO(file_data))
            properties = document.core_properties

            metadata = {
                "Title": properties.title,
                "Author": properties.author,
                "Subject": properties.subject,
                "Keywords": properties.keywords,
                "Comments": properties.comments,
                "Created": str(properties.created),
                "Modified": str(properties.modified),
                "Last Modified By": properties.last_modified_by
            }

        elif extension == "pptx":

            presentation = Presentation(BytesIO(file_data))
            properties = presentation.core_properties

            metadata = {
                "Slide Count": len(presentation.slides),
                "Title": properties.title,
                "Author": properties.author,
                "Subject": properties.subject,
                "Keywords": properties.keywords,
                "Comments": properties.comments,
                "Created": str(properties.created),
                "Modified": str(properties.modified),
                "Last Modified By": properties.last_modified_by
            }

        elif extension == "xlsx":

            workbook = load_workbook(
                BytesIO(file_data),
                read_only=True,
                data_only=False
            )

            properties = workbook.properties

            metadata = {
                "Sheet Count": len(workbook.sheetnames),
                "Sheet Names": ", ".join(workbook.sheetnames[:10]),
                "Title": properties.title,
                "Author": properties.creator,
                "Subject": properties.subject,
                "Keywords": properties.keywords,
                "Created": str(properties.created),
                "Modified": str(properties.modified),
                "Last Modified By": properties.lastModifiedBy
            }

            workbook.close()

    except Exception:
        return "Unable to read this file. Please check that it is valid.", 400

    sensitive_keywords = [
        "gps", "latitude", "longitude", "location",
        "author", "creator", "last modified by",
        "email", "owner", "device", "camera",
        "company", "comments"
    ]

    findings = []

    for key, value in metadata.items():

        if value and str(value).strip():

            if any(word in key.lower() for word in sensitive_keywords):

                findings.append({
                    "name": key,
                    "value": str(value)
                })

    if any(
        word in item["name"].lower()
        for item in findings
        for word in ["gps", "latitude", "longitude", "location"]
    ):
        risk_level = "High"
        suggestion = "Review location-related metadata before sharing this file."

    elif findings:
        risk_level = "Medium"
        suggestion = "Review personal or identifying metadata before sharing."

    else:
        risk_level = "Low"
        suggestion = "No potentially sensitive metadata was detected by this basic check."

    return render_template(
        "risk_result.html",
        filename=filename,
        file_size=file_size,
        risk_level=risk_level,
        findings=findings,
        suggestion=suggestion
    )

if __name__ == "__main__":
    app.run(debug=True)
