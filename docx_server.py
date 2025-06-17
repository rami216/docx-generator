from flask import Flask, request, send_file
from docx import Document
from docx.shared import Pt, Inches
import os
import json
import re # Import regex module for cleaning

app = Flask(__name__)

# Function to remove markdown links like ([text](url))
def clean_markdown_links(text):
    if not isinstance(text, str):
        return text
    # Regex to find patterns like ([any text](any url))
    # It will remove the entire pattern.
    cleaned_text = re.sub(r'\s*\(\[.*?\]\(.*?\)\)', '', text)
    return cleaned_text

@app.route("/generate-docx", methods=["POST"])
def generate_docx():
    # Safely parse JSON from the request body
    # request.get_json() handles different content types and ensures it's a dict
    data = request.get_json()

    if not data:
        return {"error": "Invalid or empty JSON payload"}, 400

    student_name = data.get("student_name", "Student")
    title = data.get("title", "Untitled Project")
    
    # content should now be a dict directly from n8n's JSON.parse()
    content = data.get("content", {}) 

    # Ensure content is a dictionary. If for some reason it's still a string, try parsing it.
    # This acts as a fallback, but the n8n fix should make it unnecessary.
    if isinstance(content, str):
        try:
            content = json.loads(content)
        except json.JSONDecodeError as e:
            app.logger.error(f"Error decoding content JSON: {e}")
            app.logger.error(f"Content received: {content}") # Log the problematic content
            return {"error": "Invalid JSON format for 'content' field"}, 400
    
    if not isinstance(content, dict):
        app.logger.error(f"Content is not a dictionary after parsing: {type(content)}")
        return {"error": "Content field must be a dictionary"}, 400

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

    doc.add_paragraph() # Add a blank line after title

    for section, body in content.items():
        heading = doc.add_paragraph()
        heading_run = heading.add_run(f"{section}:")
        heading_run.bold = True
        heading_run.font.size = Pt(16)

        if isinstance(body, str):
            # Clean links from string paragraphs
            doc.add_paragraph(clean_markdown_links(body))
        elif isinstance(body, dict):
            # Handle 'text' field (and potentially 'text2' etc. if they appear)
            combined_text = ""
            for key in sorted(body.keys()): # Process keys in order, e.g., text, then text2
                if key.startswith("text") and isinstance(body[key], str):
                    combined_text += " " + clean_markdown_links(body[key]) # Add space for concatenation
            
            if combined_text.strip(): # Add paragraph only if there's content
                doc.add_paragraph(combined_text.strip())

            # Handle 'bullets'
            if "bullets" in body and isinstance(body["bullets"], list):
                for bullet in body["bullets"]:
                    if isinstance(bullet, str): # Ensure bullet is a string
                        bullet_paragraph = doc.add_paragraph(style='List Bullet')
                        bullet_paragraph.paragraph_format.left_indent = Inches(0.5)
                        bullet_paragraph.add_run(clean_markdown_links(bullet))

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