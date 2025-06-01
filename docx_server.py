from flask import Flask, request, send_file
from docx import Document
from docx.shared import Pt, Inches
import os
import json

app = Flask(__name__)

@app.route("/generate-docx", methods=["POST"])
def generate_docx():
    # Safely parse JSON whether it's a dict or string
    data = request.get_json()
    if isinstance(data, str):
        data = json.loads(data)

    student_name = data.get("student_name", "Student")
    title = data.get("title", "Untitled Project")
    content = data.get("content", {})

    doc = Document()

    # Header
    p1 = doc.add_paragraph()
    r1 = p1.add_run(f"Student name: {student_name}")
    r1.bold = True
    r1.font.size = Pt(20)

    p2 = doc.add_paragraph()
    r2 = p2.add_run(f"Title: {title}")
    r2.bold = True
    r2.font.size = Pt(20)

    doc.add_paragraph()

    for section, body in content.items():
        heading = doc.add_paragraph()
        heading_run = heading.add_run(f"{section}:")
        heading_run.bold = True
        heading_run.font.size = Pt(16)

        if isinstance(body, str):
            doc.add_paragraph(body)
        elif isinstance(body, dict):
            if "text" in body:
                doc.add_paragraph(body["text"])
            if "bullets" in body and isinstance(body["bullets"], list):
                for bullet in body["bullets"]:
                    bullet_paragraph = doc.add_paragraph(style='List Bullet')
                    bullet_paragraph.paragraph_format.left_indent = Inches(0.5)
                    bullet_paragraph.add_run(bullet)

    filename = f"{title.replace(' ', '_')}.docx"
    filepath = os.path.join("/tmp", filename)
    doc.save(filepath)

    return send_file(filepath, as_attachment=True)

@app.route("/", methods=["GET"])
def health():
    return "Docx Generator is running!"

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
