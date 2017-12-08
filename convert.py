from pyexcel_xls import get_data
from lxml import etree
import re

result = get_data(afile='emissions1.xls')

xml = etree.Element('dataroot')
root = etree.ElementTree(xml)

for sheet in result:
    nodeNames = result[sheet][0]
    for values in result[sheet][1:]:
        node = etree.SubElement(xml, 'emissions')
        for index, nodeValue in enumerate(values):
            tmp = re.sub(r'[\W\s]', '', nodeNames[index])
            item = etree.SubElement(node, tmp)
            if isinstance(nodeValue, float):
                item.text = "%.2f" % nodeValue
            else:
                item.text = str(nodeValue)
        item = etree.SubElement(node, "Year")
        item.text = "2015"
    root.write("output.xml", encoding=None, method="xml", pretty_print=True)