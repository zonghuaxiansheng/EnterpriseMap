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
        self.attr_idx_ = {}
        self.idx_attr_ = {}
    def get_idx_by_attr(self, keys):
        indexes = []
        for key in keys:
            if self.attr_idx_.has_key(key):
                indexes.append(self.attr_idx_.get(key))
        return indexes
    def get_attr_by_idx(self, index):
        if self.idx_attr_.has_key(index):
            return self.idx_attr_.get(index)
    def put_idx_by_attr(self, keys, index):
        for key in keys:
            if self.attr_idx_.has_key(key):
                if self.attr_idx_.get(key) != index:
                    print("* EM_Cluster: Attr to Idx get a same key but with different value !")
                    print("* EM_Cluster: Key " + key)
                    print("* EM_Cluster: Value1 " + self.hash_.get(key))
                    print("* EM_Cluster: Value2 " + index)
                    return False
            else:
                self.attr_idx_[key] = index
        return True
    def update_idx_by_attr(self, keys, index):
        for key in keys:
            self.attr_idx_[key] = index
    def put_attr_by_idx(self, index, attrs):
        if self.idx_attr_.has_key(index):
            self.idx_attr_[index].extend(attrs)
            self.idx_attr_[index] = list(set(self.idx_attr_[index]))
        else:
            self.idx_attr_[index] = attrs
    def del_attr_by_idx(self, indexes):
        for idx in indexes:
            del self.idx_attr_[idx]
    def new_cluster(self, node):
        self.cluster_.append([node])
        return len(self.cluster_) - 1
    def put_cluster(self, index, node):
        self.cluster_[index].append(node)
    def merge_cluster(self, indexes):
        merge_nodes = []
        for idx in indexes:
            merge_nodes.extend(self.cluster_[idx])
        merge_nodes = list(set(merge_nodes))
        self.cluster_.append(merge_nodes)
        return len(self.cluster_) - 1
    def build_cluster(self, nodes):
        for node in nodes:
            keys = node["keys"]
            indexes = self.get_idx_by_attr(keys)
            if indexes:
                merge_idx = list(set(indexes))
                if len(merge_idx) == 1:
                    # Step1. Put node into cluster by index
                    self.put_cluster(merge_idx, node)
                    # Step2. Put all node's attrs into idx_attr_
                    self.put_attr_by_idx(merge_idx, keys)
                    # Step3. Put the attr->idx relationship into attr_idx_
                    success = self.put_idx_by_attr(keys, merge_idx)
                    if not success:
                        print("* EM_Cluster: Build cluster failed !")
                        exit(1)
                else:
                    # Find different indexes by node's attrs.
                    # Step1. We need to merge the different clusters, because of the investment relationship.
                    new_idx = self.merge_cluster(merge_idx)
                    # Step2. Update attr_idx_ by idx_attr_
                    attrs = []
                    for idx_ in merge_idx:
                        attrs_ = self.get_attr_by_idx(idx_)
                        attrs.append(attrs_)
                    attrs = attrs + keys
                    attrs = list(set(attrs))
                    self.update_idx_by_attr(attrs, new_idx)
                    # Step3. Update idx_attr_
                    self.put_attr_by_idx(new_idx, attrs)
                    # Step4. Delete old indexes
                    self.del_attr_by_idx(merge_idx)
                    # Step5. Put node into cluster by new_idx
                    slef.put_cluster(new_idx, node)
            else:
                new_idx = self.new_cluster(node)
                self.put_attr_by_idx(new_idx, keys)
                success = self.put_idx_by_attr(keys, new_idx)
                if not success:
                    print("* EM_Cluster: Build cluster failed !")
                    exit(1)
    def compute_cluster(self):
        for idx in self.idx_attr_.keys():
        # for cluster in self.cluster_:
            relation = len(self.cluster_[idx])
            for node in self.cluster_[idx]:
                node["relation"] = relation - 1
    def show_result(self):
        for idx in self.idx_attr_.keys():
        # for cluster in self.cluster_:
            for node in self.cluster_[idx]:
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
        
if __name__ == "__main__":
    excel = xlsx.load_workbook("../data/EnterpriseMap.xlsx")
    sheet = excel.get_sheet_by_name("EnterpriseMap")
    # excel = xlsx.load_workbook("../data/Test.xlsx")
    # sheet = excel.get_sheet_by_name("Sheet1")

    # row between 133 - 10006
    # col C:company D:LE E:SH
    column = ["C", "D", "E"]
    rows = [1, 10006]
    nodes = build_nodes(sheet, rows, column)

    em_cluster = EM_Cluster()
    em_cluster.build_cluster(nodes)
    em_cluster.compute_cluster()
    em_cluster.show_result()

