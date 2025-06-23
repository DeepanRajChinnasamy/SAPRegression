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