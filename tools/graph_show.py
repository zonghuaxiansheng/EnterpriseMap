#! /usr/bin/env python
import openpyxl as xlsx
import networkx as nx

excel = xlsx.load_workbook("../data/EnterpriseMap.xlsx")

sheet = excel.get_sheet_by_name("EnterpriseMap")

print(sheet["C8"].value)
