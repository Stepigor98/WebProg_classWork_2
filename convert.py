
from docx import Document
from html import escape

# Open Word document
doc = Document("article.docx")

# Create empty HTML content
html_content = ""

# Read each paragraph from Word
for paragraph in doc.paragraphs:
    text = paragraph.text.strip()

    if not text:
        continue

    # Convert special characters for HTML
    text = escape(text)

    # Convert Word headings to HTML headings
    if paragraph.style.name == "Title":
        html_content += f"<h1>{text}</h1>\n"

    elif paragraph.style.name.startswith("Heading"):
        html_content += f"<h2>{text}</h2>\n"

    else:
        html_content += f"<p>{text}</p>\n"

# Open default HTML template
with open("default.html", "r", encoding="utf-8") as file:
    template = file.read()

# Insert converted text into HTML template
template = template.replace(
    "<!-- your content here... -->",
    html_content
)

# Save the new HTML file
with open("article.html", "w", encoding="utf-8") as file:
    file.write(template)

print("Conversion completed successfully!")
