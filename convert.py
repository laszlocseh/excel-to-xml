from pyexcel_xls import get_data
from lxml import etree
import re
import os
import datetime

# settings
inputFolder = "inputs"
inputFile = "PollutantCodeListValues_v2.xlsx"
outputName = "EPRTR-LCP_"
sheetsToTransform = ("E-PRTR", )
columnsToTransform = () # leave empty to transform all columns
columnsToNotTransform = ()
dataElementName = "row"


# init the xml file
inputFileData = get_data(afile=os.path.join(inputFolder, inputFile))
document_node = etree.Element('dataroot')
root = etree.ElementTree(document_node)
# add 'generated' attribute in xml whith current datetime
document_node.set("generated", str(datetime.datetime.now().replace(microsecond=0).isoformat()))


# check if the column from excel is needed
def node_is_needed(name):
    return (
       not columnsToTransform
       or name in columnsToTransform
       ) \
       and name not in columnsToNotTransform


def main():
    for sheet in sheetsToTransform:
        outputFileName = os.path.join("outputs", "{}_{}.xml".format(outputName, sheet))
        node_names = inputFileData[sheet][0]
        sheet_data_rows = inputFileData[sheet][1:]
        for row in sheet_data_rows:
            main_element = etree.SubElement(document_node, dataElementName)
            for index, node_value in enumerate(row):
                node_name = node_names[index]
                if node_is_needed(node_name):
                    node_name_normalized = re.sub(r'[\W\s]', '', node_name)
                    node = etree.SubElement(main_element, node_name_normalized)
                    # if isinstance(node_value, float):
                    #     node.text = "%.5f" % node_value
                    # else:
                    #     node.text = str(node_value)
                    node.text = str(node_value)
        root.write(outputFileName, encoding=None, method="xml", pretty_print=True)


if __name__ == "__main__":
    main()
