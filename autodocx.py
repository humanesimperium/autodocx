from docx import Document

def docx_to_xml(docx_path, xml_path):
    doc = Document(docx_path)
    xml_content = ""

    # Built-in Styles to convert
    for paragraph in doc.paragraphs:
        style_name = paragraph.style.name
        text = paragraph.text
        
        # title
        if style_name == 'Title':
            xml_content += "<title>{}</title>".format(text)

        # subtitle
        elif style_name == 'Subtitle':
            xml_content += "<subtitle>{}</subtitle>".format(text)

        # headings
        elif style_name.startswith('Heading'):
            level = style_name.split()[1]
            xml_content += "<h{}>{}</h{}>".format(level, text, level)

        # normal text
        elif style_name == 'Normal':
            xml_content += "<p>{}</p>".format(text)

    # custom styles to convert. enter the exact style name as it appears in the docx file

        # List Bullet (unordered list)
        elif style_name == 'ListBullet':
            xml_content += "<ul><li>{}</li></ul>".format(text)

        # List Number
        elif style_name == 'Listenabsatz1':
            xml_content += "<ol><li>{}</li></ol>".format(text)

        # add more custom styles here  

        # bold text
        for run in paragraph.runs:
            if run.style.name == 'fett':
                xml_content = xml_content.replace(run.text, "<b>{}</b>".format(run.text))

        # italic text (for italic text, search for custom style, due to a word bug, italic itself is not recognized)
        for run in paragraph.runs:
            if run.style.name == 'kursiv':
                xml_content = xml_content.replace(run.text, "<i>{}</i>".format(run.text))


    # save xml file
    with open(xml_path, "w") as xml_file:
        xml_file.write("<document>{}</document>".format(xml_content))

# Replace 'input.docx' with the path to your input.docx file
# Replace 'output.xml' with the path where you want to save the XML file
docx_to_xml('input.docx', 'output.xml')

# some script to clean the lists in the xml file
exec(open('/mnt/c/Users/aecke/Desktop/python/autodocx/subprocesses/clean_lists.py').read())
exit()
