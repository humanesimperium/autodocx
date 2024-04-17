# autodocx
a word-to-xml-to-print transformation project

## Description
*autodocx* will be a word-to-XML python script. It should transform a docx-file out of a hot-folder into a plain XML-file using the [python-docx](https://github.com/python-openxml/python-docx) library by (c) Steve Canny.

## AI-Disclaimer
This project is a non-professional project, using the AI Github-Copilot to complement the lack of programming expertise. Feedback is highly appreciated.

## Requirements
This script is tested in an Windows 10 WSL Environment, which is the Linux Kernel running on Ubuntu ... Running on native Linux Ubuntu should be no problem. Other Unix-Environment should also work, although it is not tested. No testing nor installation on native Windows intended. 

## How to use (so far)
1. Drop any docx-file into desired directory.
2. Implement file path in the autodocx.py to where you dropped the docx. 
3. Implement file path in the autodocx.py to where you want to save the output.
4. Run.

## Troubleshooting
### single styles don´t get converted
Due to the way python-docx and word working, converting single styles may fail if the target style is not implemented in the document ([look here for further information](https://python-docx.readthedocs.io/en/latest/user/styles-understanding.html?highlight=understanding%20style)). To check if target style is implemented in your input-document (and therefore possible to find for the script), add the following argument anywhere in the script:
```
for style in document.styles:
    print("style.name == %s" % style.name)
```
This prints a list of all implemented styles in the document. If your style isn´t there, it won´t get converted. No error message or something else, it just won´t work.
