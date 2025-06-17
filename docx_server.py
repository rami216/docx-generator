from flask import Flask, request, send_file
from docx import Document
from docx.shared import Pt, Inches
import os
import json
import re # <-- New: Import re for regex
import logging

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

# Function to remove markdown links like ([text](url)) - THIS IS CRUCIAL
def clean_markdown_links(text):
    if not isinstance(text, str):
        return text
    # More aggressive regex to ensure all variations of markdown links are removed
    text = re.sub(r'\(\[.*?\]\(.*?\)\)', '', text) # Catches ([text](url))
    text = re.sub(r'\[.*?\]\([^\s)]*\)', '', text) # Catches [text](url)
    text = re.sub(r'https?://[^\s]+', '', text) # Catches bare URLs
    text = re.sub(r'\(\s*www\.[^\s)]+\s*\)|\(\s*[a-zA-Z0-9-]+\.(com|org|net|io|co|ai|xyz)\s*\)', '', text) # Catches (example.com)
    text = re.sub(r'[\(\)\[\]]', '', text) # Remove any remaining stray parentheses or brackets
    return text.strip() # Strip whitespace after cleaning

@app.route("/generate-docx", methods=["POST"])
def generate_docx():
    app.logger.info("--- Incoming Request ---")
    app.logger.info(f"Headers: {request.headers}")
    app.logger.info(f"Content-Type: {request.headers.get('Content-Type')}")

    # Attempt to get JSON. If it fails, something is severely wrong with n8n's payload
    # As per your n8n setup, request.get_json() *should* return a dict.
    data = request.get_json()

    if data is None or not isinstance(data, dict):
        app.logger.error(f"request.get_json() failed or returned non-dict: {type(data)}. Value: {data}")
        # Try to decode raw data for more debugging info if it's not JSON
        try:
            raw_data = request.data.decode('utf-8')
            app.logger.error(f"Raw received body (not JSON): {raw_data[:500]}...")
        except Exception as e:
            app.logger.error(f"Could not decode raw request data: {e}")
        return {"error": "Invalid or unparseable JSON payload at the top level from n8n."}, 400

    student_name = data.get("student_name", "Student")
    title = data.get("title", "Untitled Project")
    
    # content_raw will now be a STRING (containing the problematic markdown links)
    content_raw_string = data.get("content", "{}") # Default to empty JSON string

    parsed_content = {}
    if isinstance(content_raw_string, str):
        app.logger.info("Content field is a string. Attempting to clean and parse it.")
        # Apply the cleaning here!
        cleaned_content_string = clean_markdown_links(content_raw_string)
        
        # Log the cleaned string before parsing
        app.logger.info(f"Cleaned content string before final parse: {cleaned_content_string[:500]}...")

        try:
            parsed_content = json.loads(cleaned_content_string)
            app.logger.info("Successfully parsed content string after cleaning.")
        except json.JSONDecodeError as e:
            app.logger.error(f"Error decoding cleaned content string JSON: {e}")
            app.logger.error(f"Cleaned content that failed parsing: {cleaned_content_string[:500]}...")
            return {"error": "Invalid JSON format for 'content' field after cleaning in Flask"}, 400
    elif isinstance(content_raw_string, dict):
        # This case is less likely now, but good to have.
        parsed_content = content_raw_string
        app.logger.info("Content field was already a dictionary.")
    else:
        app.logger.error(f"Content field has unexpected type: {type(content_raw_string)}")
        return {"error": "Content field has an invalid type"}, 400

    # Ensure parsed_content is a dictionary for the rest of the logic
    if not isinstance(parsed_content, dict):
        app.logger.error(f"Parsed content is not a dictionary: {type(parsed_content)}")
        return {"error": "Content field must result in a dictionary after parsing"}, 400

    # Now use parsed_content for document generation
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

    for section, body in parsed_content.items(): # Use parsed_content here
        heading = doc.add_paragraph()
        heading_run = heading.add_run(f"{section}:")
        heading_run.bold = True
        heading_run.font.size = Pt(16)

        if isinstance(body, str):
            doc.add_paragraph(clean_markdown_links(body)) # Apply cleaning to any nested strings
        elif isinstance(body, dict):
            combined_text = ""
            # Combine 'text' fields if they exist
            for key in sorted(body.keys()):
                if key.startswith("text") and isinstance(body[key], str):
                    combined_text += " " + clean_markdown_links(body[key])
            
            if combined_text.strip():
                doc.add_paragraph(combined_text.strip())

            if "bullets" in body and isinstance(body["bullets"], list):
                for bullet in body["bullets"]:
                    if isinstance(bullet, str):
                        bullet_paragraph = doc.add_paragraph(style='List Bullet')
                        bullet_paragraph.paragraph_format.left_indent = Inches(0.5)
                        bullet_paragraph.add_run(clean_markdown_links(bullet)) # Apply cleaning to bullets

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