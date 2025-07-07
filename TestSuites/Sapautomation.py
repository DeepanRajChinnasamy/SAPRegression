

import time
import win32com.client

class Sapautomation:

    def run_sap_navigation(self):
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        # session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").expandNode("N6")
        session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").selectNode("N7")

    def get_sap_table_cell_value(self, row_index, column_id):
        """
        Gets a cell value from an SAP table in the first session.
        row_index: int (e.g., 0)
        column_id: string (e.g., 'VBELN')
        """
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)

        # Reference the table control
        table = session.findById("wnd[0]/usr/cntlCUST_100/shellcont/shell")

        # Get cell value
        value = table.GetCellValue(int(row_index), column_id)
        print(f"Cell Value at Row {row_index}, Column {column_id}: {value}")
        return value

    def select_empty_dropdown(self, elementid):
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        session.findById(elementid).key = " "
        session.findById(elementid).setFocus()

    def select_dropdown(self, elementid, value):
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        session.findById(elementid).key = value
        session.findById(elementid).setFocus()

    def popup_btn(self, button_id):
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        popup_button = session.findById(button_id)
        popup_button.press()
