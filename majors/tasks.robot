*** Settings ***
Library     RPA.Browser.Selenium    auto_close=${FALSE}
Library     RPA.Excel.Files
Library     RPA.PDF
Library     RPA.Windows
Library     Collections
Library     RPA.FileSystem
Library     DateTime
Library     String


*** Variables ***
${DATA_SOURCE}      ${CURDIR}${/}DataSheet.xlsx
${DRIVER}           ${CURDIR}${/}msedgedriver.exe

*** Tasks ***
Majors
    Open Workbook    ${DATA_SOURCE}
    ${excel_rows}=    Read Worksheet    Majors
    Log    ${excel_rows}
    ${length}=    Get Length    ${excel_rows}
    ${list_records}=    Create List
    Set Global Variable    ${list_records}
    ${current_record}=    Create Dictionary
    Set Task Variable    ${current_record}
    ${list_items}=    Create List
    Set Task Variable    ${list_items}
    ${max_row}=    Find Empty Row
    FOR    ${counter}    IN RANGE    2    ${max_row}-1
        Log    ${counter}
        ${current_row}=    Get From List    ${excel_rows}    ${counter}
        Log    ${current_row}
        Set Global Variable    ${current_row}
        IF    ${counter} == 2
            Login to App
        ELSE
            Continue on Next Risk
        END
        Flow to execute
        Exit For Loop
    END


*** Keywords ***
Flow to execute
    Risk Creation    True    False    False
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]
    ...    Submission (Pricing Completed)
	Copy Risk	True
    Sandbox Quote
    Manage Sandbox
    Finalize Tech Prem
    View Risk

Login to App
    ${URL}=    Get From Dictionary    ${current_row}    A
    ${NUSER}=    Get From Dictionary    ${current_row}    B
    ${NUSERPASSWORD}=    Get From Dictionary    ${current_row}    C
    Open Browser    ${URL}    chrome    options=add_argument("--inprivate")
    Maximize Browser Window
    Set Selenium Implicit Wait    30s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    xpath=//a[normalize-space()='Create New Risk >']
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    xpath=//a[normalize-space()='US E&U Plus']
    Sleep    10s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    id:ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlClearanceSearch_CreateInsured
    Sleep    5s
    ${EMAIL}=    Set Variable    ${NUSER}@libertymutual.com
 
    Wait Until Element Is Visible    xpath=//input[@type='email']    30s
    Clear Element Text               xpath=//input[@type='email']
    Safe Input Text                       xpath=//input[@type='email']    ${EMAIL}
 
    Sleep    1s
    Press Keys                      xpath=//input[@type='email']    ENTER
 
    Wait Until Element Is Visible    xpath=//input[@type='password']    30s
    Clear Element Text               xpath=//input[@type='password']
    Safe Input Text                       xpath=//input[@type='password']    ${NUSERPASSWORD}
 
    Sleep    1s
    Press Keys                      xpath=//input[@type='password']    ENTER
 
    Sleep    5s

Risk Creation
    [Arguments]    ${firstRun}    ${renewal}    ${copyRisk}
    Fill Insured Details    ${firstRun}    ${renewal}    ${copyRisk}
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_labelStatus'])[1]
    ...    Submission
    Fill Pricing Details    ${firstRun}    ${renewal}

Fill Insured Details
    [Arguments]    ${firstRun}    ${renewal}    ${copyRisk}
    IF    ${renewal} != True
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
        Sleep    1s
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    END
    IF    ${firstRun} == True
        ${iname}=    Get From Dictionary    ${current_row}    D
        Safe Input Text
        ...    //input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_TextBoxFirstNamedInsured1']
        ...    ${iname}
        ${state}=    Get From Dictionary    ${current_row}    E
        Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
        ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_ddUSState'])[1]
        ...    ${state}
        ${adl1}=    Get From Dictionary    ${current_row}    F
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSAddressLine1'])[1]
        ...    ${adl1}
        ${city1}=    Get From Dictionary    ${current_row}    G
        Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
        ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_ddUSCity'])[1]
        ...    ${city1}
        ${zip1}=    Get From Dictionary    ${current_row}    H
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'])[1]
        ${zipelm}=    Get WebElement
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'])[1]
        ${attribute}=    Get Element Attribute
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'])[1]
        ...    value
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'])[1]
        ...    ${zip1}
        Press Keys
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'])[1]
        ...    1+2+3+4+5+6+7+8+9
        ${return_value}=    Execute Javascript    arguments[0].value='12345-6789'    ARGUMENTS    ${zipelm}
        ${attribute}=    Get Element Attribute
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'])[1]
        ...    value
        Sleep    1s
        ${pobox}=    Get From Dictionary    ${current_row}    I
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSPoBox'])[1]
        ...    ${pobox}
    END
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Wait Until Element Is Visible    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]    15s
    ${risk_number}=    RPA.Browser.Selenium.Get Text    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]
    ${flow_type}=    Set Variable If    ${firstRun} == True    First Run    ${copyRisk} == True    Copy Risk    Unknown Flow
    Evaluate    open(r'${CURDIR}${/}risk_numbers.txt', 'a').write('${flow_type} - Risk Number: ${risk_number}\\n')
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]
    ...    LSI2
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    IF    ${firstRun} == True
        ${effdate}=    Get From Dictionary    ${current_row}    J
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
        ...    ${effdate}
        ${Indcode}=    Get From Dictionary    ${current_row}    K
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]
        ...    ${Indcode}
    END
    IF    ${copyRisk} == True
        ${effdate1}=    Get From Dictionary    ${current_row}    J
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
        ...    ${effdate1}
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]
        ...    4491 Marine cargo handling
    END
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxXsMajorsCCSNumber'])[1]
    ...    12345678
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    IF    ${firstRun} == True
        ${brokfirm}=    Get From Dictionary    ${current_row}    M
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_TextBoxBrokerFirm'])[1]
        ...    ${brokfirm}
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_ButtonBrokerFirmSearchStartsWith'])[1]
        Sleep    4s
        Wait Until Element Is Visible
        ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]    30s
        ${xpath_locator}=    Set Variable    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]
        ${table_elements}=    Get WebElements    ${xpath_locator}
        Log    Found ${table_elements.__len__()} elements in the broker info table
        Safe Click WebElements By Index    ${xpath_locator}    0
        ${value1}=    Is Element Visible    (//span[normalize-space()='Yes'])[1]
        IF    ${value1} == True
            Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[normalize-space()='Yes'])[1]
        END
    END
    Press Keys    None    PAGE_DOWN
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s
    ${variableExist}=    Run Keyword And Return Status
    ...    Element Should Be Visible
    ...    (//span[normalize-space()='Continue'])[1]
    IF    ${variableExist} == True
        Wait Until Element Is Visible    (//span[normalize-space()='Continue'])[1]    10s
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    IF    ${firstRun} == True
        ${brokcont}=    Get From Dictionary    ${current_row}    N
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]
        ...    ${brokcont}
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_ButtonBrokerContactSearchStartsWith'])[1]
        Sleep    4s
        Wait Until Element Is Visible
        ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]    30s
        ${xpath_locator}=    Set Variable    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]
        ${table_elements}=    Get WebElements    ${xpath_locator}
        Log    Found ${table_elements.__len__()} elements in the broker contact table
        Safe Click WebElements By Index    ${xpath_locator}    2
        Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]
    END
    IF    ${renewal} == True
        ${brokcont}=    Get From Dictionary    ${current_row}    N
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]
        ...    ${brokcont}
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_ButtonBrokerContactSearchStartsWith'])[1]
        Sleep    4s
        Wait Until Element Is Visible
        ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]    30s
        ${xpath_locator}=    Set Variable    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]
        ${table_elements}=    Get WebElements    ${xpath_locator}
        Log    Found ${table_elements.__len__()} elements in the broker contact table
        Safe Click WebElements By Index    ${xpath_locator}    0
        Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]
    END
    ${EditVisible}=    Is Element Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_repeaterBrokerContact_ctl01_buttonEditBroker'])[1]
    IF    ${EditVisible} == True
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_repeaterBrokerContact_ctl01_buttonEditBroker'])[1]
        Sleep    3s
        ${is_visible}=    Run Keyword And Return Status    Element Should Be Visible    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_chkPAndCLicenseHolder'])[1]
        IF    ${is_visible} == True
            ${PCLicHolder}=    Is Checkbox Selected
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_chkPAndCLicenseHolder'])[1]
            IF    ${PCLicHolder} == False
                Run Keyword And Ignore Error    Select Checkbox
                ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_chkPAndCLicenseHolder'])[1]
            END
        END
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]
        Press Keys    None    PAGE_DOWN
        ${yesVisible}=    Is Element Visible    (//span[normalize-space()='Yes'])[1]
        IF    ${yesVisible} == True
            Sleep    2s
            Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[normalize-space()='Yes'])[1]
        END
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    ELSE
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]
        Press Keys    None    PAGE_DOWN
        ${yesVisible}=    Is Element Visible    (//span[normalize-space()='Yes'])[1]
        IF    ${yesVisible} == True
            Sleep    2s
            Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[normalize-space()='Yes'])[1]
        END
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    END
    Sleep    10s
    ${value2}=    Is Element Visible    (//span[normalize-space()='Continue'])[1]
    IF    ${value2} == True
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END

Fill Pricing Details
    [Arguments]    ${firstRun}    ${renewal}
    Wait Until Element Is Visible    (//legend[normalize-space()='Expiring Policy'])[1]    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    IF    ${renewal} == True
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Ok'])[1]
    END
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info'])[1]
    ${form}=    Get From Dictionary    ${current_row}    O
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlForm'])[1]
    ...    ${form}
    ${attachment}=    Get From Dictionary    ${current_row}    P
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlAttachment'])[1]
    ...    ${attachment}
    ${industry}=    Get From Dictionary    ${current_row}    Q
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlIndustry'])[1]
    ...    ${industry}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[normalize-space()='Non Project'])[1]
    ${industry_segment}=    Get From Dictionary    ${current_row}    R
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlIndustrySegment'])[1]
    ...    ${industry_segment}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[normalize-space()='Correct'])[1]
    ${primary_class_code}=    Generate Random String    5    [NUMBERS]
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_txtPrimaryClassCode'])[1]
    ...    ${primary_class_code}
    ${category}=    Get From Dictionary    ${current_row}    S
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlCategory'])[1]
    ...    ${category}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    ${quote_rater}=    Get From Dictionary    ${current_row}    T
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlQuoteRater'])[1]
    ...    ${quote_rater}
    ${exposure1}=    Get From Dictionary    ${current_row}    U
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_txtExposure'])[1]
    ...    ${exposure1}
    ${exposure_base}=    Get From Dictionary    ${current_row}    V
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlExposureBase'])[1]
    ...    ${exposure_base}
    ${premium_basis}=    Get From Dictionary    ${current_row}    W
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlPremiumBasis'])[1]
    ...    ${premium_basis}
    ${min_earned_premium}=    Get From Dictionary    ${current_row}    X
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_txtMinEarnedPremium'])[1]
    ...    ${min_earned_premium}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    ${total_limit_each_occur}=    Get From Dictionary    ${current_row}    Y
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_txtTotalLimitEachOccur'])[1]
    ...    ${total_limit_each_occur}
    ${attachment_point}=    Get From Dictionary    ${current_row}    Z
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_txtAttachmentPoint'])[1]
    ...    ${attachment_point}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_cblGaragedStates_0'])[1]
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Underlyers'])[1]    20s
    IF    ${firstRun} == True
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_underLyerSched_underLyerList_12'])[1]
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_underLyerSched_ctrlCommercialExcessLiability_btnAdd'])[1]
        Wait Until Element Is Visible
        ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctl00_lblHeader'])[1]
        ...    20s
        ${carrier}=    Get From Dictionary    ${current_row}    AA
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctl00_txtCarrier'])[1]
        ...    ${carrier}
        ${excess_of}=    Get From Dictionary    ${current_row}    AB
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctl00_tbExcessOf'])[1]
        ...    ${excess_of}
        ${policy_number}=    Get From Dictionary    ${current_row}    AC
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctl00_txtPolicyNumber'])[1]
        ...    ${policy_number}
        ${premium}=    Get From Dictionary    ${current_row}    AD
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctl00_txtPremium'])[1]
        ...    ${premium}
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnAddNewEntry'])[1]
    END
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_underLyerSched_ctrlCommercialExcessLiability_underlyerRepeater_ctl01_btnEdit'])[1]
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[2]
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnAddNewEntry'])[1]
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='GL Info'])[1]
    ...    20s
    ${premesis_ops_occurence}=    Get From Dictionary    ${current_row}    AE
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGLInfo_txtPremisesOpsOccurence'])[1]
    ...    ${premesis_ops_occurence}
    ${hazard_type}=    Get From Dictionary    ${current_row}    AF
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGLInfo_ddlGLHazardType'])[1]
    ...    ${hazard_type}
    ${rating_type}=    Get From Dictionary    ${current_row}    AG
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGLInfo_ddlGLRatingType'])[1]
    ...    ${hazard_type}
    ${underlying_primary}=    Get From Dictionary    ${current_row}    AH
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGLInfo_txtUnderlyingPrimary'])[1]
    ...    ${underlying_primary}
    ${primary_treat_def}=    Get From Dictionary    ${current_row}    AI
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGLInfo_ddlPrimaryTreatDef'])[1]
    ...    ${primary_treat_def}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Exposure Information'])[1]    20s
    ${industry_code}=    Get From Dictionary    ${current_row}    AJ
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_formGLExposures_listViewProspectivepolicyclasscode_ctrl0_TextBoxIndustryCode'])[1]
    ...    10120
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//a[normalize-space()='10120 - Bathhouses or Bathing Pavilions'])[1]
    Sleep    2s
    ${state}=    Get From Dictionary    ${current_row}    AK
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_formGLExposures_listViewProspectivepolicyclasscode_ctrl0_ddlState'])[1]
    ...    ${state}
    ${territory}=    Get From Dictionary    ${current_row}    AL
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_formGLExposures_listViewProspectivepolicyclasscode_ctrl0_ddlTerritory'])[1]
    ...    ${territory}
    ${amount}=    Get From Dictionary    ${current_row}    AM
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_formGLExposures_listViewProspectivepolicyclasscode_ctrl0_txtAmount'])[1]
    ...    ${amount}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_formGLExposures_btnGetDefaults'])[1]
    Sleep    5s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    ...    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    ...    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAutoLiability_ddlPrimaryAutoRetention'])[1]
    ...    Guaranteed Costs
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAutoLiability_txtPriAutoLimit'])[1]
    ...    500K
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAutoLiability_ddlSigHighHazardExposure'])[1]
    ...    1
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAutoLiability_btnCalculateAuto'])[1]
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    ${al_exposure}=    Get From Dictionary    ${current_row}    BF
    FOR    ${index}    IN RANGE    0    5
        ${is_element_visible}=    Is Element Visible
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrAlExposure_listViewExposure_ctrl${index}_TextBoxALExposureAmount'])[1]
        IF    ${is_element_visible} == True
            Safe Input Text
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrAlExposure_listViewExposure_ctrl${index}_TextBoxALExposureAmount'])[1]
            ...    10000
        END
    END
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='AL Losses'])[1]    20s
    Sleep    1s
    ${loss_valuation_date}=    Get From Dictionary    ${current_row}    BF
            Safe Input Text
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAlLosses_listViewGroundUpLosses_ctrl0_datePickerLossValuationDate_textDate'])[1]
            ...    ${loss_valuation_date}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Click Element If Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[2]
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Rating Summary'])[1]    20s
    Sleep    10s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_btnCalculateLayers'])[1]
    ${primary_premium_comment}=    Get From Dictionary    ${current_row}    BG
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtPrimaryPremiumComment'])[1]
    ...    ${primary_premium_comment}
    ${deviation_comment}=    Get From Dictionary    ${current_row}    BH
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtIlfDeviationComment'])[1]
    ...    ${deviation_comment}
    ${final_pricing_comment}=    Get From Dictionary    ${current_row}    BI
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtFinalPricingRationaleComment'])[1]
    ...    ${final_pricing_comment}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s
    Wait Until Element Is Visible    (//a[normalize-space()='Rate Monitor'])[1]    20s
    ${rate_monitor_coment}=    Get From Dictionary    ${current_row}    BI
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRateMonitor_txtRateMonitorComment'])[1]
    ...    ${rate_monitor_coment}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnComplete'])[1]

Fill Pricing Details[Sandbox]
    Wait Until Element Is Visible    (//legend[normalize-space()='Expiring Policy'])[1]    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info'])[1]
    ${form}=    Get From Dictionary    ${current_row}    O
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlForm'])[1]
    ...    ${form}
    ${attachment}=    Get From Dictionary    ${current_row}    P
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlAttachment'])[1]
    ...    ${attachment}
    ${industry}=    Get From Dictionary    ${current_row}    Q
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlIndustry'])[1]
    ...    ${industry}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[normalize-space()='Non Project'])[1]
    ${industry_segment}=    Get From Dictionary    ${current_row}    R
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlIndustrySegment'])[1]
    ...    ${industry_segment}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[normalize-space()='Correct'])[1]
    ${primary_class_code}=    Generate Random String    5    [NUMBERS]
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_txtPrimaryClassCode'])[1]
    ...    ${primary_class_code}
    ${category}=    Get From Dictionary    ${current_row}    S
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlCategory'])[1]
    ...    ${category}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    ${quote_rater}=    Get From Dictionary    ${current_row}    T
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlQuoteRater'])[1]
    ...    ${quote_rater}
    ${exposure1}=    Get From Dictionary    ${current_row}    U
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_txtExposure'])[1]
    ...    ${exposure1}
    ${exposure_base}=    Get From Dictionary    ${current_row}    V
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlExposureBase'])[1]
    ...    ${exposure_base}
    ${premium_basis}=    Get From Dictionary    ${current_row}    W
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlPremiumBasis'])[1]
    ...    ${premium_basis}
    ${min_earned_premium}=    Get From Dictionary    ${current_row}    X
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_txtMinEarnedPremium'])[1]
    ...    ${min_earned_premium}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    ${total_limit_each_occur}=    Get From Dictionary    ${current_row}    Y
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_txtTotalLimitEachOccur'])[1]
    ...    ${total_limit_each_occur}
    ${attachment_point}=    Get From Dictionary    ${current_row}    Z
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_txtAttachmentPoint'])[1]
    ...    ${attachment_point}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Underlyers'])[1]    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='GL Info'])[1]
    ...    20s
    ${premesis_ops_occurence}=    Get From Dictionary    ${current_row}    AE
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGLInfo_txtPremisesOpsOccurence'])[1]
    ...    ${premesis_ops_occurence}
    ${hazard_type}=    Get From Dictionary    ${current_row}    AF
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGLInfo_ddlGLHazardType'])[1]
    ...    ${hazard_type}
    ${rating_type}=    Get From Dictionary    ${current_row}    AG
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGLInfo_ddlGLRatingType'])[1]
    ...    ${hazard_type}
    ${underlying_primary}=    Get From Dictionary    ${current_row}    AH
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGLInfo_txtUnderlyingPrimary'])[1]
    ...    ${underlying_primary}
    ${primary_treat_def}=    Get From Dictionary    ${current_row}    AI
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGLInfo_ddlPrimaryTreatDef'])[1]
    ...    ${primary_treat_def}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Exposure Information'])[1]    20s
    ${industry_code}=    Get From Dictionary    ${current_row}    AJ
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_formGLExposures_listViewProspectivepolicyclasscode_ctrl0_TextBoxIndustryCode'])[1]
    ...    10120
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//a[normalize-space()='10120 - Bathhouses or Bathing Pavilions'])[1]
    Sleep    2s
    ${state}=    Get From Dictionary    ${current_row}    AK
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_formGLExposures_listViewProspectivepolicyclasscode_ctrl0_ddlState'])[1]
    ...    ${state}
    ${territory}=    Get From Dictionary    ${current_row}    AL
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_formGLExposures_listViewProspectivepolicyclasscode_ctrl0_ddlTerritory'])[1]
    ...    ${territory}
    ${amount}=    Get From Dictionary    ${current_row}    AM
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_formGLExposures_listViewProspectivepolicyclasscode_ctrl0_txtAmount'])[1]
    ...    ${amount}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_formGLExposures_btnGetDefaults'])[1]
    Sleep    5s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    ...    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    ...    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAutoLiability_ddlPrimaryAutoRetention'])[1]
    ...    Guaranteed Costs
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAutoLiability_txtPriAutoLimit'])[1]
    ...    500K
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAutoLiability_ddlSigHighHazardExposure'])[1]
    ...    1
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAutoLiability_btnCalculateAuto'])[1]
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    ${al_exposure}=    Get From Dictionary    ${current_row}    BF
    FOR    ${index}    IN RANGE    0    5
        ${is_element_visible}=    Is Element Visible
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrAlExposure_listViewExposure_ctrl${index}_TextBoxALExposureAmount'])[1]
        IF    ${is_element_visible} == True
            Safe Input Text
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrAlExposure_listViewExposure_ctrl${index}_TextBoxALExposureAmount'])[1]
            ...    10000
        END
    END
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='AL Losses'])[1]    20s
    Sleep    1s
    ${loss_valuation_date}=    Get From Dictionary    ${current_row}    BF
            Safe Input Text
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAlLosses_listViewGroundUpLosses_ctrl0_datePickerLossValuationDate_textDate'])[1]
            ...    ${loss_valuation_date}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Rating Summary'])[1]    20s
    Sleep    10s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_btnCalculateLayers'])[1]
    ${primary_premium_comment}=    Get From Dictionary    ${current_row}    BG
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtPrimaryPremiumComment'])[1]
    ...    ${primary_premium_comment}
    ${deviation_comment}=    Get From Dictionary    ${current_row}    BH
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtIlfDeviationComment'])[1]
    ...    ${deviation_comment}
    ${final_pricing_comment}=    Get From Dictionary    ${current_row}    BI
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtFinalPricingRationaleComment'])[1]
    ...    ${final_pricing_comment}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s
    Wait Until Element Is Visible    (//a[normalize-space()='Rate Monitor'])[1]    20s
    ${rate_monitor_coment}=    Get From Dictionary    ${current_row}    BI
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRateMonitor_txtRateMonitorComment'])[1]
    ...    ${rate_monitor_coment}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnComplete'])[1]

Risk Creation[CopyRisk]
    Wait Until Element Is Visible    (//legend[normalize-space()='Expiring Policy'])[1]    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info'])[1]    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[normalize-space()='Correct'])[1]
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Schedule of Underlyers'])[1]    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='GL Info'])[1]    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='GL Exposures'])[1]    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='GL Losses'])[1]    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Underlying GL Loss Rating'])[1]    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Underlying Primary GL Selection'])[1]    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='AL Selection'])[1]    20s
    ${high_hazard_exposure}=    Get From Dictionary    ${current_row}    AR
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAutoLiability_ddlSigHighHazardExposure'])[1]
    ...    ${high_hazard_exposure}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAutoLiability_btnCalculateAuto'])[1]
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='AL Exposures'])[1]    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='AL Losses'])[1]    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rating Summary'])[1]    20s
    Sleep    3s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rate Monitor'])[1]    20s
    ${rate_monitor_coment}=    Get From Dictionary    ${current_row}    BI
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRateMonitor_txtRateMonitorComment'])[1]
    ...    ${rate_monitor_coment}
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnComplete'])[1]

Copy Risk
    [Arguments]    ${continue}
    Wait Until Element Is Visible    (//a[normalize-space()='Copy Risk'])[1]    30s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//a[normalize-space()='Copy Risk'])[1]
    Sleep    5s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[4]
    Sleep    10s
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_labelStatus'])[1]
    ...    Incomplete Submission
    IF    ${continue} == True
        Fill Insured Details    False    False    True
        Risk Creation[CopyRisk]
    ELSE
        Wait Until Keyword Succeeds    3x    5s    Safe Click Element    xpath=//div[@class='RiskHeaderColumnOne']//div[1]
    END

Sandbox Quote
    Wait Until Element Is Visible    (//a[normalize-space()='Sandbox Quote'])[1]    30s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//a[normalize-space()='Sandbox Quote'])[1]
    Sleep    5s
    ${rate_monitor_coment}=    Get From Dictionary    ${current_row}    BI
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_repeaterQuoteOptions_ctl01_quoteTab_txtSandboxTitle'])[1]
    ...    test
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[12]
    Fill Pricing Details[Sandbox]

Manage Sandbox
    Wait Until Element Is Visible    (//a[normalize-space()='Manage Sandbox'])[1]    30s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//a[normalize-space()='Manage Sandbox'])[1]
    Sleep    10s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[normalize-space()='To Quote'])[1]
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[normalize-space()='Yes'])[1]
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonReturn'])[1]

Finalize Tech Prem
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//a[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_repeaterQuoteOptions_ctl01_quoteTab_btnQuoteFinalized_lnkButton'])[1]
    ${risk_status}=    Get From Dictionary    ${current_row}    O
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlQuoteFinalized_ddlRiskStatus'])[1]
    ...    Technical Pricing Ready
    ${comments}=    Get From Dictionary    ${current_row}    BI
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlQuoteFinalized_tbComments'])[1]
    ...    test
    Sleep    10s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnSubmit'])[1]
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]
    ...    Submission (Technical Pricing Ready)
    Sleep    10s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//a[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_repeaterQuoteOptions_ctl01_quoteTab_btnQuoteFinalized_lnkButton'])[1]
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[19]
    ${risk_status}=    Get From Dictionary    ${current_row}    O
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlQuoteFinalized_ddlRiskStatus'])[1]
    ...    Quoted by Majors
    ${comments}=    Get From Dictionary    ${current_row}    BI
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlQuoteFinalized_tbComments'])[1]
    ...    test
    Sleep    10s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnSubmit'])[1]
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]
    ...    Submission (Quoted by Majors)
    Sleep    10s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//a[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_repeaterQuoteOptions_ctl01_quoteTab_btnQuoteFinalized_lnkButton'])[1]
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[19]
    ${risk_status}=    Get From Dictionary    ${current_row}    O
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlQuoteFinalized_ddlRiskStatus'])[1]
    ...    Sold by Majors
    ${comments}=    Get From Dictionary    ${current_row}    BI
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlQuoteFinalized_tbComments'])[1]
    ...    test
    Sleep    10s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnSubmit'])[1]
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]
    ...    Submission (Sold by Majors)

View Risk
    Wait Until Element Is Visible    (//a[normalize-space()='View Risk'])[1]    30s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    (//a[normalize-space()='View Risk'])[1]
    Sleep    5s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonViewCurrent'])[1]
    Sleep    5s
    Verify Status    (//legend[@id='fsMainContentLegend'])[1]    View Risk

Continue on Next Risk
    Wait Until Element Is Visible    xpath=//a[normalize-space()='Create New Risk >']    20s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    xpath=//a[normalize-space()='Create New Risk >']
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    xpath=//a[normalize-space()='Create New Risk >']
    Wait Until Element Is Visible    xpath=//a[normalize-space()='US E&U Plus']    30s
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element    xpath=//a[normalize-space()='US E&U Plus']
    Wait Until Keyword Succeeds    3x    5s    Safe Click Element
    ...    id:ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlClearanceSearch_CreateInsured
    Wait Until Element Is Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]    40s

Verify Status
    [Arguments]    ${location}    ${status_expected}
    ${status_actual}=    RPA.Browser.Selenium.Get Text    ${location}
    Log    "status_expected ${status_expected}"
    Log    "status_actual ${status_actual}"
    ${value}=    Evaluate    "${status_expected}"=="${status_actual}"
    IF    ${value} != ${TRUE}
        Fail    "Status not Macthing: Expected ${status_expected} but Current Status is ${status_actual}"
    END

Safe Click Element
    [Arguments]    ${locator}
    Wait Until Keyword Succeeds    5x    2s    Click Element When Visible    ${locator}

Safe Input Text
    [Arguments]    ${locator}    ${text}
    Wait Until Keyword Succeeds    5x    2s    Input Text When Element Is Visible    ${locator}    ${text}

Safe Select From List By Label
    [Arguments]    ${locator}    ${label}
    Wait Until Keyword Succeeds    5x    2s    Select From List By Label    ${locator}    ${label}

Safe Select From List By Index
    [Arguments]    ${locator}    ${index}
    Wait Until Keyword Succeeds    5x    2s    Select From List By Index    ${locator}    ${index}

Safe Click Element If Visible
    [Arguments]    ${locator}
    Run Keyword And Ignore Error    Wait Until Keyword Succeeds    5x    2s    Click Element When Visible    ${locator}

Safe Click WebElements By Index
    [Arguments]    ${locator}    ${index}
    Wait Until Keyword Succeeds    5x    2s    Click Element From WebElements    ${locator}    ${index}

Click Element From WebElements
    [Arguments]    ${locator}    ${index}
    ${elements}=    Get WebElements    ${locator}
    Click Element When Visible    ${elements}[${index}]
