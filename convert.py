from pyexcel_xls import get_data
from lxml import etree
import re

# settings
inputFile = "LCP_extract_v3.1_xlsx.xlsx"
outputFile = "LCP_v2_plantsdb.xml"
sheetsToTransform = ("2015", )
headersToTransform = ("MemberState", "ReferenceYear", "Unique_Plant_ID", "PlantName")
mainElementName = "Plant"

inputFileData = get_data(afile=inputFile)
xml = etree.Element('dataroot')
root = etree.ElementTree(xml)


for sheet in sheetsToTransform:
    nodeNames = inputFileData[sheet][0]
    for values in inputFileData[sheet][1:]:
        mainElement = etree.SubElement(xml, mainElementName)
        for index, nodeValue in enumerate(values):
            if nodeNames[index] in headersToTransform:
                nodeNameNormalized = re.sub(r'[\W\s]', '', nodeNames[index])
                elem = etree.SubElement(mainElement, nodeNameNormalized)
                if isinstance(nodeValue, float):
                    elem.text = "%.2f" % nodeValue
                else:
                    elem.text = str(nodeValue)
        # item = etree.SubElement(node, "Year")
        # item.text = "2015"
    root.write(outputFile, encoding=None, method="xml", pretty_print=True)
