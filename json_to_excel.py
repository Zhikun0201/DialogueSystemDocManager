import os
import json
from enum import Enum


class XlData():
    work_book = None
    work_sheet = None
    flow = []
    flow_all_paths = []
    dlg_json = {}
    sheet_row = None
    next_str = ["Next", "Skip", "Finish"]

class XlHead(Enum):
    '''Excel表格的表头枚举'''
    GUID = 1
    TechIndex = 2
    TechJump = 3
    WorkIndex = 4
    Owner = 5
    Content = 6
    WorkJump = 7
    
class DlgNodeType(Enum):
    StartNode = -1
    SpeechNode = 0
    

class JsonToExcel(object):
    def __init__(self):
        super(JsonToExcel, self).__init__()
        json_path = os.path.abspath(os.path.join(os.getcwd(), "../.."))
        print(json_path)
        json_file = '%s/Dialogue/Dlg_TestFile.dlg.json' % json_path
        self.dlg_json_loader(json_file)
        print("JsonToExcel Init")

    @staticmethod
    def dlg_json_loader(json_file):
        file = None

        with open(json_file, encoding="UTF-8") as f:
            file = json.load(f)
        if not file:
            return

        # 获取开始节点
        XlData.dlg_json = file
        print("Dlg json has been loaded.")
        print(XlHead.GUID.value)
        f.close()
        
        
if __name__ == "__main__":
    JsonToExcel()