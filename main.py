from docx import Document
import os
import subprocess #for opening document on wsl ubuntu
from docx.shared import Inches  # margins
from docx.shared import Pt, RGBColor  # styling headings
from docx.oxml.ns import qn  # page nums
from docx.oxml import OxmlElement  # page nums
from docx.enum.text import WD_ALIGN_PARAGRAPH  # For justification
from docx.oxml import OxmlElement  # For horizontal line


from languages import languagesData

leftLang = "english"
rightLang = "spanish"

books = ["1-nephi", "enos", "moroni"]
chapters = {
    "1-nephi": 6, "enos": 1, "moroni": 6
}

# Create a new Word document
document = Document()

# Set smaller margins
sections = document.sections
for section in sections:
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.2)
    section.right_margin = Inches(0.2)
    section.footer_distance = Inches(0.2)  # make footer with page number smaller 

# Define a function to customize heading style
def customize_heading_style(doc, level, font_name='Arial', font_size=14, font_color=RGBColor(0, 0, 0)):
    style = doc.styles[f'Heading {level}']
    font = style.font
    font.name = font_name
    font.size = Pt(font_size)
    font.color.rgb = font_color
    paragraph_format = style.paragraph_format
    paragraph_format.alignment = 1  # Center alignment

# Customize the heading styles
customize_heading_style(document, level=1, font_name='Arial', font_size=18, font_color=RGBColor(0, 0, 0))
customize_heading_style(document, level=2, font_name='Arial', font_size=16, font_color=RGBColor(0, 0, 0))
customize_heading_style(document, level=3, font_name='Arial', font_size=14, font_color=RGBColor(0, 0, 0))

def style_cell_text(cell, text, font_name='Georgia', font_size=11, font_color=RGBColor(0, 0, 0)):
    # Clear existing text
    cell.text = ''
    # Create a new run for the cell
    run = cell.paragraphs[0].add_run(text.strip())
    # Apply the styles
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.color.rgb = font_color
    # Set the alignment to justify
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    # Adjust paragraph spacing
    paragraph_format = cell.paragraphs[0].paragraph_format
    paragraph_format.space_before = Pt(0)  # No space before the paragraph
    paragraph_format.space_after = Pt(4)  # Small space after the paragraph
    paragraph_format.line_spacing = Pt(12)  # Adjust line spacing (this can be fine-tuned)

def add_verses_to_table(row_cells, english_verse, spanish_verse):
    # Add English verse without a new paragraph
    style_cell_text(row_cells[0], f"{english_verse}")
    # Add Spanish verse without a new paragraph
    style_cell_text(row_cells[1], f"{spanish_verse}")

def add_horizontal_line(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    hr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')  # Border size
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    hr.append(bottom)
    p._element.get_or_add_pPr().append(hr)

def add_title_page(doc):
    # Title page
    doc.add_paragraph("\n\n\n") # Add spacing before the title
    # Add the main title in large, bold font
    main_title = doc.add_paragraph()
    main_title.alignment = 1  # Center alignment
    run = main_title.add_run(f'{languagesData[rightLang]["book-of-mormon"]}')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(36)
    run.bold = True
    # Add subtitle in a slightly smaller font and italics
    subtitle = doc.add_paragraph()
    subtitle.alignment = 1  # Center alignment
    run = subtitle.add_run(f'{languagesData[rightLang]["another-testament-of-jesus-christ"]}')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(24)
    run.italic = True
    # Add spacing between title and subtitle
    doc.add_paragraph("\n\n")
    # Book of Mormon in second langaueg
    second_title = doc.add_paragraph()
    second_title.alignment = 1  # Center alignment
    run = second_title.add_run(f'{languagesData[leftLang]["book-of-mormon"]}')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(36)
    run.bold = True
    # Add another subtitle for the translation language in a smaller font
    subtitle_3 = doc.add_paragraph()
    subtitle_3.alignment = 1  # Center alignment
    run = subtitle_3.add_run(f'{languagesData[leftLang]["another-testament-of-jesus-christ"]}')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(24)
    run.italic = True
    # Add more spacing before the side-by-side description
    doc.add_paragraph("\n\n")
    # Add a description below the titles
    description = doc.add_paragraph()
    description.alignment = 1  # Center alignment
    run = description.add_run(f'{languagesData[leftLang]["side-by-side-version"]}')
    run.font.name = 'Arial'
    run.font.size = Pt(18)
    run.bold = True
    # Add language pairing description in smaller font
    language_pair = doc.add_paragraph()
    language_pair.alignment = 1  # Center alignment
    run = language_pair.add_run(f'{languagesData[leftLang][leftLang].capitalize()} | {languagesData[rightLang][rightLang].capitalize()}')
    run.font.name = 'Arial'
    run.font.size = Pt(16)
    # Add more spacing before the page break
    doc.add_paragraph("\n\n\n")
    # Page break
    doc.add_page_break()

#add_title_page(document)

# Iterate through each book
for book in books:
    add_horizontal_line(document) # line before book title
    document.add_heading(f"{languagesData[leftLang][book]} | {languagesData[rightLang][book]}", level=2)
    document.add_paragraph("") #space after book title

    # Iterate through each chapter
    for chapter in range(1, chapters[book] + 1):
        eng_path = f'bom2/bom-{leftLang}/{book}/{chapter}.txt'
        spa_path = f'bom2/bom-{rightLang}/{book}/{chapter}.txt'
        
        # Check if both files exist
        if os.path.exists(eng_path) and os.path.exists(spa_path):
            with open(eng_path, 'r', encoding='utf-8') as eng_file:
                english_verses = [line.strip() for line in eng_file.readlines() if line.strip()] #removes new linee characters

            with open(spa_path, 'r', encoding='utf-8') as spa_file:
                 spanish_verses = [line.strip() for line in spa_file.readlines() if line.strip()] #removes new linee characters

            # Create a table with two columns
            table = document.add_table(rows=0, cols=2)

            # Have the first row of the cols be "Chapter X" and "Chapitre X"
            row_cells = table.add_row().cells
            style_cell_text(row_cells[0], f"{languagesData[leftLang]['chapter']} {chapter}", font_name='Daytona', font_size=12, font_color=RGBColor(0, 0, 0))
            style_cell_text(row_cells[1], f"{languagesData[rightLang]['chapter']} {chapter}", font_name='Daytona', font_size=12, font_color=RGBColor(0, 0, 0))

            # Ensure both files have the same number of verses
            min_len = min(len(english_verses), len(spanish_verses))

            # Add verses to the table with verse numbers
            for i in range(min_len):
                row_cells = table.add_row().cells
                add_verses_to_table(row_cells, f"{i+1} {english_verses[i].strip()}", f"{i+1} {spanish_verses[i].strip()}")

            # Add a space after each chapter
            document.add_paragraph("")
        else:
            document.add_paragraph("")

# Page numbers
def add_page_numbers(document):
    sections = document.sections
    for section in sections:
        footer = section.footer
        p = footer.add_paragraph()
        p.alignment = 1  # Center-align the page number
        run = p.add_run()
        field = OxmlElement('w:fldSimple')
        field.set(qn('w:instr'), 'PAGE')  # PAGE is the instruction for page number
        run._element.append(field)

# Add page numbers to the footer
add_page_numbers(document)

# Save the document
document.save("side_by_side_bom.docx")
subprocess.call(['powershell.exe', 'Start-Process', 'side_by_side_bom.docx']) #opens doc
