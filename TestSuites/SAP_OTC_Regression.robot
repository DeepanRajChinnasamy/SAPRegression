*** Settings ***
Resource    ../Resource/ObjectRepositories/CustomVariables.robot
Library     ../Resource/ObjectRepositories/CustomLib.py
Library     ../TestSuites/Sapautomation.py
Library     ../Resource/ObjectRepositories/Response.py

*** Variables ***
${InputExcel}    ${execdir}\\UploadExcel\\SAP_OTC_Regression.xlsx
${SAPGUIPATH}    C:/Program Files (x86)/SAP/FrontEnd/SAPgui/saplogon.exe
${CONNECTION}    EQ2-Load balancer
${SAP_CLIENT}      100
${SAP_USER}    SAPQA_APP1
${SAP_PASSWORD}    Quality75#
${ENTERBUTTON}    /app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[0]
${ExecuteButton}    /app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[8]
${popup1enterbutton}    /app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[0]
${popup2enterbutton}    /app/con[0]/ses[0]/wnd[2]/tbar[0]/btn[0]
${SaveButton}    /app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[11]
${StatusbarText}    /app/con[0]/ses[0]/wnd[0]/sbar/pane[0]
${btn1}    /app/con[0]/ses[0]/wnd[1]/usr/btn%#AUTOTEXT001
${SAPGUIPATH}     C:/Program Files (x86)/SAP/FrontEnd/SAPgui/saplogon.exe
${screenshotpath}    C:\\Users\\dchinnasam\\OneDrive\\Documents\\SAP Test Automation\\SAP-Test-Automation\\TestSuites\\Screenshots
${SAPSYSTEMNAME}  EQ2-Load balancer
${SAPCLIENT}      400
${SAPUSERNAME}    SAPQA_APP1
${SAPPASSWORD}    Quality75#
${ShipToBP}      /app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR
${SoldToBP}      /app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR
${EnterButton}   /app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[0]
${ItemOveriewTab}    /app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02
${element_exists}
${InputExcelPath}    ${execdir}\\UploadExcel\\TAX-EQ2.xlsx
${InputExcelSheet}    Inputs
${MaterialList}
${green}    00FF00
${red}    FF0000
${SAPPopUpElement}   /app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[0]
${orderTypeElementID}    /app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-AUART
${sellingOrdElementID}    /app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-VKORG
${distributionChannelElementID}   /app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-VTWEG
${divisonElemID}    /app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-SPART
${statusbarTextID}    /app/con[0]/ses[0]/wnd[0]/sbar/pane[0]
${MaterialTableId}    /app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4427/subSUBSCREEN_TC:SAPMV45A:4908/tblSAPMV45ATCTRL_U_ERF_KONTRAKT/ctxtRV45A-MABNR[1,0]
${TargetQtyId}    /app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4427/subSUBSCREEN_TC:SAPMV45A:4908/tblSAPMV45ATCTRL_U_ERF_KONTRAKT/txtVBAP-ZMENG[2,0]
${headerButton}    /app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD
${salesHeaderTab}    /app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01
${orderHeaderTab}    /app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11
${sellingGrpTextbox}    /app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-VKBUR
${PurchaseGrpTextbox}    /app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/ssubSUBSCREEN_BODY:SAPMV45A:4351/ctxtVBKD-BSARK
${SE16NTableEntriesID}    /app/con[0]/ses[0]/wnd[0]/usr/cntlRESULT_LIST/shellcont/shell
${VA43OrderTextBox}    /app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-VBELN
${ItemOverViewTable}    /app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4427/subSUBSCREEN_TC:SAPMV45A:4908/tblSAPMV45ATCTRL_U_ERF_KONTRAKT
${ItemConditionTab}    /app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\07
${NetAmountTextBox}    /app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\07/ssubSUBSCREEN_BODY:SAPLV69A:6201/txtKOMP-NETWR
${PercentageAmountTexBox}    /app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\07/ssubSUBSCREEN_BODY:SAPLV69A:6201/txtKOMP-MWSBP
${NetPriceTextBox}    /app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBAK-NETWR
${screeshotname}    0
${jpgdoctype}    .jpg
${FolderPath}   C:\\Users\\dchinnasam\\OneDrive\\Documents\\01. Offical\\Robot\\ExecutionDoc\\
${PopupCancelButton}    /app/con[0]/ses[0]/wnd[1]/usr/btnCANCEL
${PopupEditButton}    /app/con[0]/ses[0]/wnd[1]/usr/btnSPOP-VAROPTION2

*** Test Cases ***
KBE_51_001
   [Tags]    id=KBE_51_001
    Connect To SAP
    Read All Input Values From Excel    ${InputExcel}    Data
    ${ListIndexIterator}    set variable    0
    ${DataIndexIterator}    set variable    0
    ${IdocIDCount}=    get length    ${IdocNumberList}
    ${RowCounter}    set variable    2
    # Start VA01 transaction
    FOR    ${ScenarioIterator}    IN RANGE    ${IdocIDCount}
        ${IdocNumber}=    get from list    ${IdocNumberList}    ${ListIndexIterator}
        #Trigger the Existing Idoc in WE19
        run transaction    we19
        sapguilibrary.input text    /app/con[0]/ses[0]/wnd[0]/usr/ctxtMSED7START-EXIDOCNUM    ${IdocNumber}
        send vkey    F8
        sapguilibrary.click element    ${ExecuteButton}
        sapguilibrary.click element   ${popup1enterbutton}
        ${IdocNumber}=    sapguilibrary.get value    /app/con[0]/ses[0]/wnd[2]/usr/txtMESSTXT1
        @{arrIdocText}=    split string  ${IdocNumber}  ${SPACE}
        ${IdocNumber}=  get from list    ${arrIdocText}    1
        log to console  ${IdocNumber}
        write output excel    Data    NewIdocNumber    ${RowCounter}    ${IdocNumber}
        sapguilibrary.click element    ${popup2enterbutton}
        #Process the IDoc in Bd87
        Run Transaction    /nBd87
        sapguilibrary.input text  /app/con[0]/ses[0]/wnd[0]/usr/ctxtSX_DOCNU-LOW    ${IdocNumber}
        sapguilibrary.click element    ${ExecuteButton}
        run sap navigation
        sapguilibrary.click element    ${ExecuteButton}
        sleep    3s
        ${errorText}=   sapguilibrary.get value    ${StatusbarText}
        should contain    ${errorText}    saved
        log to console    ${errorText}
        @{errorText}=    split string    ${errorText}    ${SPACE}
        ${SubOrderNumber}=  get from list    ${errorText}    2
        write output excel    Data    SubscriptionOrder    ${RowCounter}    ${SubOrderNumber}
        #Credit Release Flag
        run transaction    /nVKM3
        sapguilibrary.input text    /app/con[0]/ses[0]/wnd[0]/usr/ctxtVBELN-LOW    ${SubOrderNumber}
        send vkey    F8
        sapguilibrary.select checkbox    /app/con[0]/ses[0]/wnd[0]/usr/chk[1,3]
        sapguilibrary.click element    /app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[34]
        sapguilibrary.click element    ${SaveButton}
        #Remove Block Manually
        run transaction    /nVA42
        sapguilibrary.input text    /app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-VBELN    ${SubOrderNumber}
        sapguilibrary.click element    ${ENTERBUTTON}
        select empty dropdown    /app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4426/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-FAKSK
        sleep    3s
        sapguilibrary.click element    ${SaveButton}
        sapguilibrary.click element   ${popup1enterbutton}
        sapguilibrary.input text    /app/con[0]/ses[0]/wnd[1]/usr/cntlLONGTEXT/shellcont/shell    Test
        sapguilibrary.click element   ${popup1enterbutton}
        sleep    3s
        sapguilibrary.click element    ${SaveButton}
        #Create Billing Document
        run transaction    /nVF01
        sapguilibrary.click element    ${ENTERBUTTON}
        sapguilibrary.click element    ${SaveButton}
        ${statustext}=    sapguilibrary.get value    /app/con[0]/ses[0]/wnd[0]/sbar
        @{statustext}=    split string    ${statustext}    ${SPACE}
        ${DocumentNumber}=    get from list    ${statustext}    1
        should contain    ${statustext}    saved
        write output excel    Data    BillingDocumentNumber    ${RowCounter}    ${DocumentNumber}
        #Verfiy the Document Flow
        run transaction    /nVa43
        sapguilibrary.input text    /app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-VBELN    ${SubOrderNumber}
        sapguilibrary.click element    /app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[17]
    END
    save excel document    ${InputExcel}
    close all excel documents
    close sap connection
KBE_51_002
    [Tags]    id=KBE_51_002
    Connect To SAP
    run transaction    /nBP
    sapguilibrary.click element    /app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[5]
    Sapautomation.select dropdown   /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7013/subA02P01:SAPLBUD0:1130/cmbBUS000FLDS-TITLE_MEDI    0002
#    sapguilibrary.input text    /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7013/subA02P01:SAPLBUD0:1130/cmbBUS000FLDS-TITLE_MEDI    Mr.
    ${rdn}=    gettimestamp
    sapguilibrary.input text    /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7013/subA02P03:SAPLBUD0:1301/txtBUT000-NAME_FIRST    Automation${rdn}
    sapguilibrary.input text    /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7013/subA02P04:SAPLBUD0:1302/txtBUT000-NAME_LAST    Auto
    sapguilibrary.input text    /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7013/subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA7:0600/subCOUNTRY_SCREEN:SAPLSZA7:0601/txtADDR2_DATA-STREET    H-2-34
    sapguilibrary.input text    /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7013/subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA7:0600/subCOUNTRY_SCREEN:SAPLSZA7:0601/txtADDR2_DATA-POST_CODE1    500055
    sapguilibrary.input text    /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7013/subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA7:0600/subCOUNTRY_SCREEN:SAPLSZA7:0601/txtADDR2_DATA-CITY1    HYDERABAD
    sapguilibrary.input text    /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7013/subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA7:0600/subCOUNTRY_SCREEN:SAPLSZA7:0601/ctxtADDR2_DATA-COUNTRY    IN
    sapguilibrary.input text    /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7013/subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA7:0600/subCOUNTRY_SCREEN:SAPLSZA7:0601/ctxtADDR2_DATA-REGION    36
    sapguilibrary.input text    /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7013/subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA7:0600/subCOUNTRY_SCREEN:SAPLSZA7:0601/txtSZA7_D0400-SMTP_ADDR    Test@Test.com
    sapguilibrary.click element    ${SaveButton}
    ${popupvisible}=    run keyword and return status    sapguilibrary.element should be present   /app/con[0]/ses[0]/wnd[1]/usr/
    IF    '${popupvisible}' == 'True'
        customvariables.click sap popup button if present     ${btn1}
        ${popupvisible}=    run keyword and return status    sapguilibrary.element should be present   /app/con[0]/ses[0]/wnd[2]/usr/
        IF    '${popupvisible}' == 'True'
            customvariables.click sap popup button if present    /app/con[0]/ses[0]/wnd[2]/usr/btnSPOP-OPTION1
            select table row    /app/con[0]/ses[0]/wnd[1]/usr/cntlALV_CONTAINER/shellcont/shell    0
            customvariables.click sap popup button if present     ${btn1}
        END
        customvariables.click sap popup button if present    /app/con[0]/ses[0]/wnd[1]/usr/btn%#AUTOTEXT012
    END
    ${BPID}=    sapguilibrary.get value    /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/subSCREEN_1000_HEADER_AREA:SAPLBUPA_DIALOG_JOEL:1510/ctxtBUS_JOEL_MAIN-CHANGE_NUMBER
    log to console    ${BPID}
    send vkey    F6
    Sapautomation.select dropdown    /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/subSCREEN_1100_ROLE_AND_TIME_AREA:SAPLBUPA_DIALOG_JOEL:1110/cmbBUS_JOEL_MAIN-PARTNER_ROLE    ISM000
    sapguilibrary.click element      /app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[27]
    sapguilibrary.click element    /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/subSCREEN_1100_SUB_HEADER_AREA:SAPLCVI_FS_UI_CUSTOMER_SALES:0001/btnPUSH_SA
    sapguilibrary.input text    /app/con[0]/ses[0]/wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLCVI_FS_UI_CUSTOMER_SALES:0002/tblSAPLCVI_FS_UI_CUSTOMER_SALESTCTRL_SALES_AREA/ctxtCVIS_SALES_AREA_DYNPRO-SALES_ORG[0,0]    1001
    sapguilibrary.click element    /app/con[0]/ses[0]/wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLCVI_FS_UI_CUSTOMER_SALES:0002/btnPUSH_SA_OKAY
    customvariables.click sap popup button if present    /app/con[0]/ses[0]/wnd[2]/tbar[0]/btn[0]
    sapguilibrary.click element    /app/con[0]/ses[0]/wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLCVI_FS_UI_CUSTOMER_SALES:0002/btnPUSH_SA_OKAY
    sapguilibrary.click element    /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_05
    sapguilibrary.input text    /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_05/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLCVI_FS_UI_CUSTOMER_SALES:0089/ctxtGS_KNVV-KVGR1    C0A
    sapguilibrary.click element    /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01
    sapguilibrary.input text    /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7011/subA02P01:SAPLCVI_FS_UI_CUSTOMER_SALES:0071/ctxtGS_KNVV-KDGRP    01
    sapguilibrary.input text   /app/con[0]/ses[0]/wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7011/subA02P02:SAPLCVI_FS_UI_CUSTOMER_SALES:0072/ctxtGS_KNVV-VKBUR  0080
    sapguilibrary.click element    ${SaveButton}
    customvariables.click sap popup button if present    /app/con[0]/ses[0]/wnd[1]/usr/btnSPOP-OPTION1
    run transaction    /nVA41
    sapguilibrary.input text    ${orderTypeElementID}  ZSUB
    sapguilibrary.input text    ${sellingOrdElementID}    1001
    sapguilibrary.input text    ${distributionChannelElementID}    00
    sapguilibrary.input text    ${divisonElemID}    00
    sapguilibrary.click element    ${ENTERBUTTON}
    sapguilibrary.input text    ${ShipToBP}    ${BPID}
    sapguilibrary.input text    ${SoldToBP}    ${BPID}
    sapguilibrary.click element    ${ENTERBUTTON}
    ${statusText}=    sapguilibrary.get value    ${statusbarTextID}
    IF    '${statusText}' == '${EMPTY}'
        click sap popup button if present    ${SAPPopUpElement}
        sapguilibrary.click element    ${ItemOveriewTab}
        sapguilibrary.input text    ${MaterialTableId}    BESTP
        ${MaterialErrorText}=    sapguilibrary.get value    ${statusbarTextID}
        IF    '${MaterialErrorText}' == '${EMPTY}'
            sapguilibrary.input text    ${TargetQtyId}    1
            send vkey    0
            click sap popup button if present    ${SAPPopUpElement}
            #sleep    5s
            sapguilibrary.click element    ${headerButton}
            #sleep    3s
            sapguilibrary.click element    ${salesHeaderTab}
            sapguilibrary.input text    ${sellingGrpTextbox}    0050
            sapguilibrary.click element    ${orderHeaderTab}
            sapguilibrary.input text    ${PurchaseGrpTextbox}    0020
            send vkey    11
            #${element_exists}  set variable
            click sap popup button if present    ${SAPPopUpElement}
            click sap popup button if present    ${SAPPopUpElement}
            click sap popup button if present   ${PopupEditButton}
            ${OrderIssue}    sapguilibrary.get value    ${statusbarTextID}
            ${Order}=    sapguilibrary.get value    ${statusbarTextID}
            sleep    3s
#            ${OrderCheck}    should contain    ${Order}    saved
#            IF    '${OrderCheck}' == 'True'
                ${OrderNumList}    split string    ${Order}    ${SPACE}
                log to console    ${OrderNumList}[2]
                ${OrderNumber}    set variable    ${OrderNumList}[2]
#                write and color excel cell   ${Inputexcelsheet}    ContractNo    ${RowCounter}    ${OrderNumber}    ${green}
#                save excel document    ${InputExcelPath}
#            END
            #Credit Release Flag
            run transaction    /nVKM3
            sapguilibrary.input text    /app/con[0]/ses[0]/wnd[0]/usr/ctxtVBELN-LOW    ${OrderNumList}[2]
            send vkey    F8
            sapguilibrary.select checkbox    /app/con[0]/ses[0]/wnd[0]/usr/chk[1,3]
            sapguilibrary.click element    /app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[34]
            sapguilibrary.click element    ${SaveButton}
            run transaction    /nVF01
            sapguilibrary.click element    ${ENTERBUTTON}
            sapguilibrary.click element    ${SaveButton}
            ${statustext}=    sapguilibrary.get value    /app/con[0]/ses[0]/wnd[0]/sbar
            @{statustext}=    split string    ${statustext}    ${SPACE}
            ${DocumentNumber}=    get from list    ${statustext}    1
            should contain    ${statustext}    saved
#            write output excel    Data    BillingDocumentNumber    ${RowCounter}    ${DocumentNumber}
            #Verfiy the Document Flow
            run transaction    /nVA43
            sapguilibrary.input text    /app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-VBELN    ${OrderNumList}[2]
            sapguilibrary.click element    /app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[17]
        END
    END
    close sap connection


*** Keywords ***
Connect To SAP
    [Documentation]    Keyword to establish connection to SAP system
    Start Process    ${SAPGUIPATH}    saplogon
    sleep    5s
    SapGuiLibrary.Connect To Session
    open connection    ${CONNECTION}
#    Send Vkey    F12    # Get rid of any initial popups if they appear
#    Wait Until Element Is Visible    wnd[0]/usr/txtRSYST-BNAME
    SapGuiLibrary.input text        wnd[0]/usr/txtRSYST-BNAME    ${SAP_USER}
    SapGuiLibrary.input password    wnd[0]/usr/pwdRSYST-BCODE    ${SAP_PASSWORD}
#    Send Vkey    F8    # Use F8 instead of Enter for login
    sapguilibrary.click element    ${ENTERBUTTON}
#    Click SAP PopUp Button If Present    ${continuebutton}
    sapguilibrary.click element    ${ENTERBUTTON}
    # Handle potential session expired or information dialogs
    ${status}    ${value}=    Run Keyword And Ignore Error    Element Should Be Visible    wnd[1]
    Run Keyword If    '${status}'=='PASS'    Click Element    wnd[1]/usr/btnSPOP-OPTION1



Read All Input Values From Excel
    [Arguments]    ${InputExcel}    ${InputSheet}
    ${ExcelDictionary}    ReadAllValuesFromExcel    ${InputExcel}    ${InputSheet}
    ${IdocNumberList}=    get from dictionary    ${ExcelDictionary}    IdocNumber
    set suite variable    ${IdocNumberList}    ${IdocNumberList}
    open excel document    ${InputExcel}    docID