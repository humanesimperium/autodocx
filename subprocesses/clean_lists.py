# clean up for not entirely well made conversion of the main script

import re

# opens a specific xml file, parse through and search for occurences of "</ul>\n<ul>" and "</ol>\n<ol>" and remove them
with open("path/to/your/file/XY.xml", "r") as file:
    content = file.read()
    content = re.sub(r"</ul>\s*<ul>", "", content, flags=re.DOTALL)
    content = re.sub(r"</ol>\s*<ol>", "", content, flags=re.DOTALL)

# save the cleaned xml file
with open("path/to/your/file/XY.xml", "w") as file:
    file.write(content)

