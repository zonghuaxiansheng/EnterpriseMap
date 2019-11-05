#! /usr/bin/env python
# coding=utf-8
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import openpyxl as xlsx
import networkx as nx
import matplotlib.pyplot as plt

excel = xlsx.load_workbook("../data/EnterpriseMap.xlsx")
sheet = excel.get_sheet_by_name("EnterpriseMap")
# excel = xlsx.load_workbook("../data/Test.xlsx")
# sheet = excel.get_sheet_by_name("Sheet1")

# row between 133 - 10006
# col C:company D:LE E:SH
column = ["C", "D", "E"]
rows = 180

# for col in column:
node_dict = {}
print("* Read xlsx")
for row in range(133, rows + 1):
    from_node = sheet[column[2] + str(row)].value
    to_node = sheet[column[0] + str(row)].value
    node_dict[from_node] = to_node

node_graph = nx.DiGraph()
node_graph.clear()

print("* NX build graph")
for from_nodes, to_node in node_dict.items():
    from_node_list = str(from_nodes).split(",")
    for from_node in from_node_list:
        # print from_node,"->",to_node
        node_graph.add_edge(from_node, to_node)

print("* NX draw graph")
# nx.draw_networkx(node_graph, pos=nx.random_layout(node_graph), with_labels=True, font_size=8, node_color="r", edge_color="black")
nx.draw_networkx(node_graph,
                 pos=nx.random_layout(node_graph),
                 with_labels=False,
                 font_size=6,
                 node_color="r",
                 edge_color="b")
print("* NX save graph")
pic_name = "node_graph.png"
plt.savefig(pic_name)
plt.show()
