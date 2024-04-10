import os
from docx import Document


# _______ pre section _______

# Defining directory containing the documents
directory = '/path/to/your/directory/'

# Iterate over all files in the directory
for filename in os.listdir(directory):
    # Check if the file is a .docx file
    if filename.endswith(".docx"):
        # Construct the full path to the document
        filepath = os.path.join(directory, filename)
        # Open the document
        document = Document(filepath)

# Defining styles, needed for the more complicated and custom styles (e.g. lists, bold, italic, etc.) so far.
styles = document.styles

# _______ pre section end _______


# _______ built-in section _______

# Following loops iterate for Word built-in-styles and adds a corresponding <tags>. looping for
#   Title
#   Subtitle
#   Heading 1; Heading 2; Heading 3; Heading 4; Heading 5; Heading 6;
#   Normal
# 

# Title
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Title':
        paragraph.text = '<?xml version="1.0" encoding="UTF-8"?><doc><meta><title>' + paragraph.text + '</title>'

# Subtitle
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Subtitle':
        paragraph.text = '<subtitle>' + paragraph.text + '</subtitle>'

# Heading 1
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Heading 1':
        paragraph.text = '<h1>' + paragraph.text + '</h1>'

# Heading 2
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Heading 2':
        paragraph.text = '<h2>' + paragraph.text + '</h2>'

# Heading 3
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Heading 3':
        paragraph.text = '<h3>' + paragraph.text + '</h3>'

# Heading 4
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Heading 4':
        paragraph.text = '<h4>' + paragraph.text + '</h4>'

# Heading 5
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Heading 5':
        paragraph.text = '<h5>' + paragraph.text + '</h5>'

# Heading 6
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Heading 6':
        paragraph.text = '<h6>' + paragraph.text + '</h6>'

# Normal
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Normal':
        paragraph.text = '<p>' + paragraph.text + '</p>'

# _______ built-in section end _______


# _______ custom style section _______

# Following loops iterate for Word custom styles (styles created by the user) and lists and adds desired <tags>. Loops must be given the exact style name from the word doc. <tags> must be defined in the "paragraph.text" argument.

# List Bullet ('Aufzählungszeichen1') und formatiert den paragraphen als "normalen" Text und fügt "<ul>" davor und "</ul>" danach ein.
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Aufzählungszeichen1':
        paragraph.style = styles['Normal'] # could be that this is of no use, if the Lists can be converted with built-in-Style. The "styles = document.styles" argument could be deleted if so.
        paragraph.text = '<ul>' + paragraph.text + '</ul>'

# List Number ('Listenabsatz1') und formatiert den paragraphen als "normalen" Text und fügt "<ol>" davor und "</ol>" danach ein.
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Listenabsatz1':
        paragraph.style = styles['Normal'] # could be that this is of no use, if the Lists can be converted with built-in-Style. The "styles = document.styles" argument could be deleted if so.
        paragraph.text = '<ol>' + paragraph.text + '</ol>'

# Italic ('Standard kursiv') und fügt vor der Run "<italic>" ein und danach "</italic>"
for paragraph in document.paragraphs:
    for run in paragraph.runs:
        if run.italic:
            run.text = '<italic>' + run.text + '</italic>'

# work in progress
# durchläuft den Text und sucht nach fettgedrucktem Text und fügt davor "<bold>" ein und danach "</bold>"
for paragraph in document.paragraphs:
    for run in paragraph.runs:
        if run.bold:
            run.text = '<bold>' + run.text + '</bold>'

# author
for paragraph in document.paragraphs:
    if paragraph.style.name == 'author':
        paragraph.text = '<author>' + paragraph.text + '</author></meta>'

# doc_end
for paragraph in document.paragraphs:
    if paragraph.style.name == 'doc_end':
        paragraph.text = '</doc>'

# _______ custom style section end _______


# _______ post section _______

# save edited file to a output directory
document.save('/path/to/your/directory/' + filename) # Don't forget to change the path to your desired output directory.

# _______ post section end _______
