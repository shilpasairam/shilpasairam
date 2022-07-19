# Import the required library
from lxml import html, etree
import os


def reportHtmlToXml(htmlReportFilePath):
    # Main Function
    # Provide the path of the html file
    # file = htmlReportFilePath
    file = htmlReportFilePath

    # Open the html file and Parse it,
    # returning a single element/document.
    with open(file, 'r', encoding='utf-8') as inp:
        htmldoc = html.fromstring(inp.read())

    # Open a output.xml file and write the
    # element/document to an encoded string
    # representation of its XML tree.
    with open(f"{os.getcwd()}\\Reports\\output.xml", 'wb') as out:
        out.write(etree.tostring(htmldoc))
