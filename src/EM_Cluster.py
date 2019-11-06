#! /usr/bin/env python
# coding=utf-8
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import openpyxl as xlsx
import networkx as nx
import matplotlib.pyplot as plt

class EM_Cluster:
    def __init__(self):
        self.cluster_ = []
        self.hash_ = {}
    def get_hash_index(self, keys):
        for key in keys:
            if self.hash_.has_key(key):
                return self.hash_.get(key)
        return None
    def new_cluster(self, node):
        self.cluster_.append([node])
        return len(self.cluster_) - 1
    def push_into_cluster(self, index, node):
        self.cluster_[index].append(node)
    def push_into_hash(self, keys, index):
        for key in keys:
            if self.hash_.has_key(key):
                if self.hash_.get(key) != index:
                    print("* EM_Cluster: Hash get a same key but with different value !")
                    print("* EM_Cluster: Key " + key)
                    print("* EM_Cluster: Value1 " + self.hash_.get(key))
                    print("* EM_Cluster: Value2 " + index)
                    return False
            else:
                self.hash_[key] = index
        return True
    def build_cluster(self, nodes):
        for node in nodes:
            keys = node["keys"]
            index = self.get_hash_index(keys)
            if index:
                self.push_into_cluster(index, node)
                success = self.push_into_hash(keys, index)
                if not success:
                    print("* EM_Cluster: Build cluster failed !")
                    exit(1)
            else:
                index_ = self.new_cluster(node)
                success = self.push_into_hash(keys, index_)
                if not success:
                    print("* EM_Cluster: Build cluster failed !")
                    exit(1)
    def compute_cluster(self):
        for cluster in self.cluster_:
            relation = len(cluster)
            for node in cluster:
                node["relation"] = relation - 1
    def show_result(self):
        for cluster in self.cluster_:
            for node in cluster:
                print("* Node: " + node["name"] + " Relation: " + node["relation"])

# Build nodes by read EnterpriseMap.xlsx
# node dict structure:
#   name: the company name
#   keys: the names who invested in the company
#   relation: the investment-count
#   point: use this point to write investment-count back into the xlsx
def build_nodes(sheet, rows, column):
    nodes = []
    for row in range(rows[0], rows[1] + 1):
        LE = sheet[column[1] + str(row)].value
        SH = sheet[column[2] + str(row)].value
        LE_list = LE.split(",")
        SH_list = SH.split(",")
        SH_list.extend(LE_list)
        from_nodes = list(set(SH_list))
        to_node = sheet[column[0] + str(row)].value
        node = {}
        node["name"] = str(to_node)
        node["keys"] = from_nodes
        node["relation"] = 0
        node["point"] = column[3] + str(row)
        nodes.append(node)
    return nodes
        

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
