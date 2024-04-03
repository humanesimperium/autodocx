import os
from docx import Document

# Directory containing the documents
directory = '/mnt/c/Users/aecke/Desktop/python/input/'

# Iterate over all files in the directory
for filename in os.listdir(directory):
    # Check if the file is a .docx file
    if filename.endswith(".docx"):
        # Construct the full path to the document
        filepath = os.path.join(directory, filename)
        
        # Open the document
        document = Document(filepath)


styles = document.styles
# Bei den folgenden Schleifen wird nach den jeweiligen Styles gesucht und die entsprechenden Tags hinzugefügt. Beim Style.name muss dann der Name der Formatvorlage eingetragen werden.


# durchläuft alle Absätze und sucht nach Title und fügt davor "<title>" ein und danach "</title>".
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Title':
        paragraph.text = '<?xml version="1.0" encoding="UTF-8"?><doc><meta><title>' + paragraph.text + '</title>'

# durchläuft alle Absätze und sucht nach Subtitle und fügt davor "<subtitle>" ein und danach "</subtitle>".
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Subtitle':
        paragraph.text = '<subtitle>' + paragraph.text + '</subtitle>'

# durchläuft alle Absätze und sucht nach author und fügt davor "<author>" ein und danach "</author>".
for paragraph in document.paragraphs:
    if paragraph.style.name == 'author':
        paragraph.text = '<author>' + paragraph.text + '</author></meta>'

# durchläuft alle Absätze und sucht nach Heading 1 und fügt davor "<h1>" ein und danach "</h1>".
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Heading 1':
        paragraph.text = '<h1>' + paragraph.text + '</h1>'

# durchläuft alle Absätze und sucht nach Heading 2 und fügt davor "<h2>" ein und danach "</h2>"
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Heading 2':
        paragraph.text = '<h2>' + paragraph.text + '</h2>'

# durchläuft alle Absätze und sucht nach Heading 3 und fügt davor "<h3>" ein und danach "</h3>"
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Heading 3':
        paragraph.text = '<h3>' + paragraph.text + '</h3>'

# durchläuft alle Absätze und sucht nach Heading 4 und fügt davor "<h4>" ein und danach "</h4>"
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Heading 4':
        paragraph.text = '<h4>' + paragraph.text + '</h4>'

# durchläuft alle Absätze und sucht nach Heading 5 und fügt davor "<h5>" ein und danach "</h5>"
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Heading 5':
        paragraph.text = '<h5>' + paragraph.text + '</h5>'

# durchläuft alle Absätze und sucht nach Heading 6 und fügt davor "<h6>" ein und danach "</h6>"
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Heading 6':
        paragraph.text = '<h6>' + paragraph.text + '</h6>'

# durchläuft alle Absätze und sucht nach Paragraph oder "Fließtext" und fügt davor "<p>" ein und danach "</p>"
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Normal':
        paragraph.text = '<p>' + paragraph.text + '</p>'

# durchläuft alle Absätze und sucht Paragraphen mit 'Aufzählungszeichen1' und formatiert den paragraphen als "normalen" Text und fügt "<ul>" davor und "</ul>" danach ein.
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Aufzählungszeichen1':
        paragraph.style = styles['Normal']
        paragraph.text = '<ul>' + paragraph.text + '</ul>'

# durchläuft alle Absätze und sucht Paragraphen mit 'Listenabsatz1' und formatiert den paragraphen als "normalen" Text und fügt "<ol>" davor und "</ol>" danach ein.
for paragraph in document.paragraphs:
    if paragraph.style.name == 'Listenabsatz1':
        paragraph.style = styles['Normal']
        paragraph.text = '<ol>' + paragraph.text + '</ol>'

# durchläuft die Text runs und sucht nach der Formatvorlage 'Standard kursiv' und fügt vor der Run "<italic>" ein und danach "</italic>"
for paragraph in document.paragraphs:
    for run in paragraph.runs:
        if run.italic:
            run.text = '<italic>' + run.text + '</italic>'


# durchläuft den Text und sucht nach fettgedrucktem Text und fügt davor "<bold>" ein und danach "</bold>"
for paragraph in document.paragraphs:
    for run in paragraph.runs:
        if run.bold:
            run.text = '<bold>' + run.text + '</bold>'



# durchläuft den Text und sucht nach der Formatierung "doc_end" und ersetzt den Absatz mit "</doc>".
for paragraph in document.paragraphs:
    if paragraph.style.name == 'doc_end':
        paragraph.text = '</doc>'



# speichert das bearb. Dokument im "output" Ordner ohne den Namen zu ändern
document.save('/mnt/c/Users/aecke/Desktop/python/output/' + filename) # Hier muss der Pfad angegeben werden, wo die bearbeiteten Dateien gespeichert werden sollen.
