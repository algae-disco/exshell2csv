#!/bin/python3

import sys
import xml.etree.ElementTree as ET

root = ET.fromstring(sys.stdin.read())

for c in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
    cell_name = c.attrib.get('r')
    cell_type = c.attrib.get('t', 'v')
    v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
    cell_value = v.text if v is not None else ""

    print(f"{cell_name} {cell_type} {cell_value} ")
