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
document_node = etree.Element('dataroot')
root = etree.ElementTree(document_node)

document_node.set("generated", str(datetime.datetime.now().replace(microsecond=0).isoformat()))


def node_is_needed(name):
    return (
       not headersToTransform
       or name in headersToTransform
       ) \
       and name not in headersToNotTransform


def main():
    for sheet in sheetsToTransform:
        node_names = inputFileData[sheet][0]
        sheet_data_rows = inputFileData[sheet][1:]
        for row in sheet_data_rows:
            main_element = etree.SubElement(document_node, mainElementName)
            for index, node_value in enumerate(row):
                node_name = node_names[index]
                if node_is_needed(node_name):
                    node_name_normalized = re.sub(r'[\W\s]', '', node_name)
                    node = etree.SubElement(main_element, node_name_normalized)
                    if isinstance(node_value, float):
                        node.text = "%.4f" % node_value
                    else:
                        node.text = str(node_value)
        root.write(outputFile, encoding=None, method="xml", pretty_print=True)


if __name__ == "__main__":
    main()
