from docx import Document


def docx_to_xml(docx_path, xml_path):
    doc = Document(docx_path)
    xml_content = ""

    # Built-in Styles to convert
    for paragraph in doc.paragraphs:
        if paragraph.style.name == 'Title':
            xml_content += "<title>{}</title>".format(paragraph.text)

        elif paragraph.style.name == 'Subtitle':
            xml_content += "<subtitle>{}</subtitle>".format(paragraph.text)

        elif paragraph.style.name == 'Heading 1':
            xml_content += "<h1>{}</h1>".format(paragraph.text)

        elif paragraph.style.name == 'Heading 2':
            xml_content += "<h2>{}</h2>".format(paragraph.text)

        elif paragraph.style.name == 'Heading 3':
            xml_content += "<h3>{}</h3>".format(paragraph.text)

        elif paragraph.style.name == 'Heading 4':
            xml_content += "<h4>{}</h4>".format(paragraph.text)

        elif paragraph.style.name == 'Heading 5':
            xml_content += "<h5>{}</h5>".format(paragraph.text)

        elif paragraph.style.name == 'Heading 6':
            xml_content += "<h6>{}</h6>".format(paragraph.text)

        elif paragraph.style.name == 'Normal':
            xml_content += "<p>{}</p>".format(paragraph.text)

    ### custom styles to convert. enter the exact style name as it appears in the docx file

        # List Bullet (unordered list)
        elif paragraph.style.name == 'ListBullet':
            xml_content += "<ul><li>{}</li></ul>".format(paragraph.text)

        # List Number
        elif paragraph.style.name == 'Listenabsatz1':
            xml_content += "<ol><li>{}</li></ol>".format(paragraph.text)

### Converting Italic and Bold. Doesnt work yet



    with open(xml_path, "w") as xml_file:
        xml_file.write("<document>{}</document>".format(xml_content))

# Replace 'input.docx' with the path to your input .docx file
# Replace 'output.xml' with the path where you want to save the XML file
docx_to_xml('input.docx', 'output.xml')

exec(open('/mnt/c/Users/aecke/Desktop/python/autodocx/subprocesses/clean_lists.py').read())
exit()