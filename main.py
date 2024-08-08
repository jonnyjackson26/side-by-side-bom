from docx import Document
import os
from docx.shared import Inches

leftLang = "english"
rightLang = "spanish"
#books = ["1-nephi", "2-nephi", "jacob", "enos", "jarom", "omni", "words-of-mormon", "mosiah", "alma", "helaman", "3-nephi", "4-nephi", "mormon", "ether", "moroni"]
#chapters = {
#    "1-nephi": 22, "2-nephi": 33, "jacob": 7, "enos": 1, "jarom": 1, "omni": 1, "words-of-mormon": 1,
#    "mosiah": 29, "alma": 63, "helaman": 16, "3-nephi": 30, "4-nephi": 1,
#    "mormon": 9, "ether": 15, "moroni": 10
#}
books = ["1-nephi", "2-nephi", "enos", "moroni"]
chapters = {
    "1-nephi": 22, "2-nephi": 33,"enos": 1, "moroni": 10
}


document = Document() # Create a new Word document
# smaller margins
sections = document.sections
for section in sections:
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    
#title page
document.add_heading("Side-by-Side Verses of the Book of Mormon", level=1)
document.add_paragraph(f'{rightLang} | {leftLang}')
document.add_page_break()

# Iterate through each book
for book in books:
    document.add_heading(book.replace("-", " ").title(), level=2)

    # Iterate through each chapter
    for chapter in range(1, chapters[book] + 1):
        document.add_heading(f"Chapter {chapter}", level=3)

        eng_path = f'bom/bom-{leftLang}/{book}/{chapter}.txt'
        spa_path = f'bom/bom-{rightLang}/{book}/{chapter}.txt'
        
        # Check if both files exist
        if os.path.exists(eng_path) and os.path.exists(spa_path):
            with open(eng_path, 'r', encoding='utf-8') as eng_file:
                english_verses = eng_file.readlines()

            with open(spa_path, 'r', encoding='utf-8') as spa_file:
                spanish_verses = spa_file.readlines()

            # Create a table with two columns
            table = document.add_table(rows=0, cols=2)

            # Ensure both files have the same number of verses
            min_len = min(len(english_verses), len(spanish_verses))

            # Add verses to the table with verse numbers
            for i in range(min_len):
                row_cells = table.add_row().cells
                row_cells[0].text = f"{i+1} {english_verses[i].strip()}"
                row_cells[1].text = f"{i+1} {spanish_verses[i].strip()}"
        else:
            document.add_paragraph(f"Chapter {chapter} of {book} is missing in one or both languages.")

# Save the document
document.save("side_by_side_bom.docx")
