from ariba_ope import Ariba_Auto
from cm_ui import CM_UI
if __name__ == "__main__":
    ariba_case1=Ariba_Auto()
    ui=CM_UI(ariba_case1)
    ui.run()


