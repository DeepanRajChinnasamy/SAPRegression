

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


    def get_table_row_count(self):
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        table = session.FindById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC")
        return table.RowCount

    def find_field_in_table(self, field_label):
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)

        table = session.FindById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC")
        total_rows = table.RowCount
        visible_rows = table.VisibleRowCount

        for row in range(total_rows):
            # Focus a cell in the current row to scroll it into view
            try:
                # Focus the first column (0) in the row
                cell_id = f"/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtRSRD1-FIELDNAME[0,{row}]"
                cell = session.FindById(cell_id)
                cell.SetFocus()

                # Now get the cell value and compare
                value = cell.Text
                if value == field_label:
                    return f"Field '{field_label}' found at row {row}"

            except Exception as e:
                continue

        return f"Field '{field_label}' not found"
