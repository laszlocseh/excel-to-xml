from pyexcel_xls import get_data
from lxml import etree
import re
import os
import datetime

# settings
inputFile = os.path.join("inputs", "TEST-COMPANIES-Licences.xlsx")
outputFile = os.path.join("outputs", "test-companies-new.xml")
sheetsToTransform = ("Sheet2", )
headersToTransform = ()
headersToNotTransform = ()
mainElementName = "licence"

inputFileData = get_data(afile=inputFile)
xml = etree.Element('dataroot')
root = etree.ElementTree(xml)

xml.set("generated", str(datetime.datetime.now().replace(microsecond=0).isoformat()))
for sheet in sheetsToTransform:
    nodeNames = inputFileData[sheet][0]
    for values in inputFileData[sheet][1:]:
        mainElement = etree.SubElement(xml, mainElementName)
        for index, nodeValue in enumerate(values):
            if not headersToTransform and \
                    (nodeNames[index] in headersToTransform
                     or nodeNames[index] not in headersToNotTransform):
                nodeNameNormalized = re.sub(r'[\W\s]', '', nodeNames[index])
                elem = etree.SubElement(mainElement, nodeNameNormalized)
                if isinstance(nodeValue, float):
                    elem.text = "%.4f" % nodeValue
                else:
                    elem.text = str(nodeValue)
        # item = etree.SubElement(node, "Year")
        # item.text = "2015"
    root.write(outputFile, encoding=None, method="xml", pretty_print=True)
