class DebugMessage(object):
    def __init__(self, name, object=None):
        print(f"Debug Message <{name}>:", object)
        
class DebugExit(object):
    def __init__(self, name):
        print(f"Debug Exit: <{name}> is not valid.")