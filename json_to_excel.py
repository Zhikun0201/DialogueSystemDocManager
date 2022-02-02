import os
import json
from enum import Enum, unique


class JTEData():
    work_book = None
    work_sheet = None
    flow = []
    flow_order = []
    flow_all_paths = []
    dlg_json = {}
    node_relations = {}
    sheet_row = None
    next_str = ["Next", "Skip", "Finish"]


@unique
class XlHead(Enum):
    '''Excel表格的表头枚举'''
    GUID = 1
    TechIndex = 2
    TechJump = 3
    WorkIndex = 4
    Owner = 5
    Content = 6
    WorkJump = 7


@unique
class DlgNodeType(Enum):
    '''DlgSystem中的节点类型
    Start: -1, Sequence: 0, Speech: 1, Selector: 2, End: 3.
    '''
    Start = -1
    Sequence = 0
    Speech = 1
    Selector = 2
    End = 3


class JsonToExcel(object):
    def __init__(self):
        super(JsonToExcel, self).__init__()
        print("JsonToExcel Init.")
        # get json file path
        json_dir = os.path.abspath(os.path.join(os.getcwd(), "../.."))
        json_path = '%s/Dialogue/Dlg_TestFile.dlg.json' % json_dir
        # factory
        self.dlg_json_loader(json_path)

    @staticmethod
    def dlg_json_loader(json_file):
        """加载DlgSystem资产的json文件
        """
        if not os.path.exists(json_file):
            print("Exit: Dlg json path is not valid.")
            return
        
        with open(json_file, encoding="UTF-8") as f:
            dj_file = json.load(f)
        if not dj_file:
            return

        # 获取开始节点
        JTEData.dlg_json = dj_file
        print("Dlg json has been loaded.")
        f.close()

    @staticmethod
    def get_node_relations():
        '''记录索引与children的关系（StartNode的索引记作-1）
        '''
        dlgfile = JTEData.dlg_json
        if not dlgfile:
            print("Exit: Dlg json is not valid.")
            return

        node_relations = {}
        start_children = dlgfile["StartNode"]["Children"]
        start_targets = []
        # StartNode
        for child in start_children:
            start_targets.append(child["TargetIndex"])
        node_relations[-1] = start_targets
        # Other nodes
        all_nodes = dlgfile["Nodes"]
        for node in all_nodes:
            node_index = node["__index__"]
            node_targets = []
            node_children = node["Children"]
            for child in node_children:
                node_targets.append(child["TargetIndex"])
            node_relations[node_index] = node_targets
        JTEData.node_relations = node_relations

        print("Debug Message <node relations>:", node_relations)

    @staticmethod
    def flow_order():
        """将node按flow的先后顺序排序
        """
        node_relations = JTEData.node_relations.copy()
        if not node_relations:
            print("Exit: Node relations is not valid.")
            return

        flow_dict = {}
        start_node = node_relations.pop(-1)
        flow_dict[-1] = start_node

        while node_relations != {}:
            # 拿到flow中最后一个有效的chidren的第一个child
            flow_revers = list(flow_dict.values())
            flow_revers.reverse()
            for target_list in flow_revers:
                if target_list == []:
                    continue
                flow_value_last = target_list.pop(0)
                break

            # 检查child是否也是其他节点的child
            same_elm_found = False
            for node_children in node_relations.values():
                if flow_value_last in node_children:
                    same_elm_found = True

            # 如果其他节点没有该child，则将其置入flow
            if not same_elm_found:
                popitem = node_relations.pop(flow_value_last)
                flow_dict[flow_value_last] = popitem

        flow_order = list(flow_dict.keys())
        JTEData.flow_order = flow_order
        print("Debug Message <flow>:", flow_order)

    @staticmethod
    def indent_level():
        """计算每个节点的缩进等级
        """
        flow_order = JTEData.flow_order
        if not flow_order:
            print("Exit: Flow order is not valid.")
            return
        
        


if __name__ == "__main__":
    JsonToExcel()
    JsonToExcel.get_node_relations()
    JsonToExcel.flow_order()
