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


*** Tasks ***
US Excess and Umbrella
    Open Workbook    ${DATA_SOURCE}
    ${excel_rows}=    Read Worksheet    EU
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
        Flow to execute
    END


*** Keywords ***
Flow to execute
    Login to App
    Risk Creation    True    False    False    True
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Submission (Pricing In Progress)
    Pre Bind Endorsement
    Quote
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Quoted
    Copy Quote
    Ready To Bind
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Bound (Pending)
    Book    False
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Bound
    Copy Risk
    Reissue
    Post Bind Endorsement
    Renewal
    View Risk
    Close Browser

Login to App
    ${URL}=    Get From Dictionary    ${current_row}    A
    ${NUSER}=    Get From Dictionary    ${current_row}    B
    ${NUSERPASSWORD}=    Get From Dictionary    ${current_row}    C
    Open Browser    ${URL}    chrome    options=add_argument("--inprivate")
    Maximize Browser Window
    Set Selenium Implicit Wait    30s
    Click Element When Visible    xpath=//a[normalize-space()='Create New Risk >']
    Click Element When Visible    xpath=//a[normalize-space()='US Excess and Umbrella']
    Wait Until Element Is Visible
    ...    id:ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlClearanceSearch_CreateInsured
    ...    30s
    Click Element When Visible
    ...    id:ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlClearanceSearch_CreateInsured
    ${EMAIL}=    Set Variable    ${NUSER}@libertymutual.com
    Wait Until Element Is Visible    xpath=//input[@type='email']    30s
    Clear Element Text    xpath=//input[@type='email']
    Input Text Safely    xpath=//input[@type='email']    ${EMAIL}
    Sleep    1s
    Press Keys    xpath=//input[@type='email']    ENTER
    Wait Until Element Is Visible    xpath=//input[@type='password']    30s
    Clear Element Text    xpath=//input[@type='password']
    Input Text Safely    xpath=//input[@type='password']    ${NUSERPASSWORD}
    Sleep    1s
    Press Keys    xpath=//input[@type='password']    ENTER
    Wait Until Element Is Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]    30s

Risk Creation
    [Arguments]    ${firstRun}    ${renewal}    ${copyRisk}    ${Puerto_Rico}
    Fill Insured Details    ${firstRun}    ${renewal}    ${copyRisk}    ${Puerto_Rico}
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Submission
    Fill Pricing Details    ${firstRun}    ${renewal}

Fill Insured Details
    [Arguments]    ${firstRun}    ${renewal}    ${copyRisk}    ${Puerto_Rico}
    IF    ${renewal} != True
        Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
        Sleep    1s
        Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    END
    IF    ${firstRun} == True
        ${ifirstname}=    Get From Dictionary    ${current_row}    D
        Input Text Safely
        ...    //input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_TextBoxFirstNamedInsured1']
        ...    ${ifirstname}
        ${country}=    Get From Dictionary    ${current_row}    D
        Select Option By Label
        ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_ddCountry'])[1]
        ...    Puerto Rico
        ${State_visible}=    Is Element Visible
        ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_ddUSState'])[1]
        IF    ${State_visible} == True
            ${state}=    Get From Dictionary    ${current_row}    E
            Select Option By Label
            ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_ddUSState'])[1]
            ...    ${state}
        END
        ${Address_line_visible}=    Is Element Visible
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSAddressLine1'])[1]
        IF    ${Address_line_visible} == True
            ${address_line_1}=    Get From Dictionary    ${current_row}    F
            Input Text Safely
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSAddressLine1'])[1]
            ...    ${address_line_1}
        END
        ${Address_line_visible1}=    Is Element Visible
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxROWAddressLine1'])[1]
        IF    ${Address_line_visible1} == True
            ${address_line_1}=    Get From Dictionary    ${current_row}    F
            Input Text Safely
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxROWAddressLine1'])[1]
            ...    ${address_line_1}
        END
        IF    ${Puerto_Rico} == True
            ${city1}=    Get From Dictionary    ${current_row}    G
            Input Text Safely
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxROWAddressLine4'])[1]
            ...    ${city1}
        ELSE
            ${city1}=    Get From Dictionary    ${current_row}    G
            Select Option By Label
            ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_ddUSCity'])[1]
            ...    ${city1}
        END
        IF    ${Puerto_Rico} == False
            ${zip1}=    Get From Dictionary    ${current_row}    H
            Click Element When Visible
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'])[1]
            ${zipelm}=    Get WebElement
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'])[1]
            ${attribute}=    Get Element Attribute
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'])[1]
            ...    value
            Input Text Safely
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
            Input Text Safely
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSPoBox'])[1]
            ...    ${pobox}
        END
    END
    Click Next And Wait For Element
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    Click Next And Wait For Element    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]
    Wait Until Element Is Visible    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]    15s
    ${risk_number}=    RPA.Browser.Selenium.Get Text    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]
    ${flow_type}=    Set Variable If    ${firstRun} == True    First Run    ${copyRisk} == True    Copy Risk    Renewal
    Evaluate    open(r'${CURDIR}${/}risk_numbers.txt', 'a').write('${flow_type} - Risk Number: ${risk_number}\\n')
	Select Option By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListBranch'])[1]
    ...    Caribbean
    Select Option By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]
    ...    LMLATAM
    Select Option By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    IF    ${firstRun} == True
        ${effective_date}=    Get From Dictionary    ${current_row}    J
        ${effective_date}=    Format Excel Date    ${effective_date}
        Input Text Safely
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
        ...    ${effective_date}
        ${Industry_code}=    Get From Dictionary    ${current_row}    K
        Input Text Safely
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]
        ...    ${Industry_code}
    END
    IF    ${copyRisk} == True
        ${effective_date1}=    Get From Dictionary    ${current_row}    L
        ${effective_date1}=    Format Excel Date    ${effective_date1}
        Input Text Safely
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
        ...    ${effective_date1}
        ${Industry_code1}=    Get From Dictionary    ${current_row}    K
        Input Text Safely
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]
        ...    ${Industry_code1}
    END
    Click Next And Wait For Element    (//legend[normalize-space()='Broker Info'])[1]
    IF    ${firstRun} == True or ${copyRisk} == True
        ${broker_firm}=    Get From Dictionary    ${current_row}    M
        Input Text Safely
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_TextBoxBrokerFirm'])[1]
        ...    ${broker_firm}
        Click Element When Visible
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_ButtonBrokerFirmSearchContains'])[1]
        Sleep    4s
        Wait Until Element Is Visible
        ...    (//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']//input[@type="submit"])[1]
        ...    30s
        Wait Until Keyword Succeeds
        ...    5x
        ...    2s
        ...    Click Element
        ...    (//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']//input[@type="submit"])[1]
        Sleep    3s
        Sleep    2s
        ${yes_visible}=    Run Keyword And Return Status
        ...    Wait Until Keyword Succeeds    3x    2s    Element Should Be Visible    (//span[normalize-space()='Yes'])[1]
        IF    ${yes_visible} == True
            Wait Until Keyword Succeeds
            ...    3x
            ...    2s
            ...    Click Element When Visible    (//span[normalize-space()='Yes'])[1]
        END
        Sleep    2s
    END
    Sleep    3s
    Wait Until Keyword Succeeds
    ...    5x
    ...    3s
    ...    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s
    ${variableExist}=    Run Keyword And Return Status
    ...    Wait Until Element Is Visible
    ...    (//span[normalize-space()='Continue'])[1]
    ...    5s
    IF    ${variableExist} == True
        Wait Until Keyword Succeeds
        ...    3x
        ...    5s
        ...    Click Element When Visible
        ...    (//span[normalize-space()='Continue'])[1]
    END
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    IF    ${firstRun} == True or ${copyRisk} == True
        ${broker_contact}=    Get From Dictionary    ${current_row}    N
        Input Text Safely
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]
        ...    ${broker_contact}
        Click Element When Visible
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_ButtonBrokerContactSearchStartsWith'])[1]
        Sleep    4s
        Wait Until Element Is Visible    (//table[@id='TableBrokerContactSearch']//input[@value="Select"])[1]    30s
        Wait Until Keyword Succeeds
        ...    5x
        ...    2s
        ...    Click Element
        ...    (//table[@id='TableBrokerContactSearch']//input[@value="Select"])[1]
        Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]
        Click Element When Visible
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]
    END
    Click Next And Wait For Element
    ...    xpath=//span[contains(@id, 'labelStatus')]

Fill Pricing Details
    [Arguments]    ${firstRun}    ${renewal}
    IF    ${renewal} == True
        Wait Until Element Is Visible    (//a[normalize-space()='Expiring Policy Info'])[1]    20s
        Click Next And Wait For Element    (//a[normalize-space()='New Policy Info'])[1]
    END
    Wait Until Element Is Visible    (//a[normalize-space()='New Policy Info'])[1]    20s
    ${risk_state}=    Get From Dictionary    ${current_row}    O
    Select Option By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlRiskState'])[1]
    ...    ${risk_state}
    ${treaty}=    Get From Dictionary    ${current_row}    P
    Select Option By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlTreaty'])[1]
    ...    ${treaty}
    Click Element When Visible    (//span[normalize-space()='Admitted'])[1]
    ${quote_type}=    Get From Dictionary    ${current_row}    Q
    Select Option By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlQuoteType'])[1]
    ...    ${quote_type}
    ${policy_type}=    Get From Dictionary    ${current_row}    R
    Select Option By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlPolicyType'])[1]
    ...    ${policy_type}
    ${industry}=    Get From Dictionary    ${current_row}    S
    Select Option By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlIndustry'])[1]
    ...    ${industry}
    ${exposure}=    Get From Dictionary    ${current_row}    T
    Input Text Safely
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbExposure'])[1]
    ...    ${exposure}
    ${exposure_base}=    Get From Dictionary    ${current_row}    U
    Select Option By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlExposureBase'])[1]
    ...    ${exposure_base}
    Click Element When Visible    (//span[normalize-space()='Yes'])[1]
    ${total_limit}=    Get From Dictionary    ${current_row}    V
    Input Text Safely
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbTotalLimit'])[1]
    ...    ${total_limit}
    ${quota_share}=    Get From Dictionary    ${current_row}    W
    Input Text Safely
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbQuotaShare'])[1]
    ...    ${quota_share}
    ${attachment_point}=    Get From Dictionary    ${current_row}    X
    Input Text Safely
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbAttachmentPoint'])[1]
    ...    ${attachment_point}
    Click Next And Wait For Element    (//a[normalize-space()='Additional Policy Info'])[1]
    Click Element When Visible    (//span[normalize-space()='Non Project'])[1]
    ${minimum_earned_premium}=    Get From Dictionary    ${current_row}    Y
    Input Text Safely
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbMinEarnedPremium'])[1]
    ...    ${minimum_earned_premium}
    ${cyber_type}=    Get From Dictionary    ${current_row}    Y
    Select Option By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlCyberType'])[1]
    ...    Full Exclusion
    ${cyber_limlit}=    Get From Dictionary    ${current_row}    Y
    Input Text Safely
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbCyberLimit'])[1]
    ...    0
    ${exclusion_list}=    Get From Dictionary    ${current_row}    Y
    Select Option By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlExclusionList'])[1]
    ...    Tier 1
    Click Next And Wait For Element    (//legend[normalize-space()='Underlyers'])[1]
    IF    ${firstRun} == True
        Click Element When Visible
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_underLyerSched_underLyerList_12'])[1]
        Click Element When Visible
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_underLyerSched_ctrlCommercialExcessLiability_btnAdd'])[1]
        Wait Until Element Is Visible
        ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctl00_lblHeader'])[1]
        ...    20s
        ${carrier}=    Get From Dictionary    ${current_row}    Z
        Input Text Safely
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctl00_txtCarrier'])[1]
        ...    ${carrier}
        ${excess_of}=    Get From Dictionary    ${current_row}    AA
        Input Text Safely
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctl00_tbExcessOf'])[1]
        ...    ${excess_of}
        ${policy_number}=    Get From Dictionary    ${current_row}    AB
        Input Text Safely
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctl00_txtPolicyNumber'])[1]
        ...    ${policy_number}
        ${premium}=    Get From Dictionary    ${current_row}    AC
        Input Text Safely
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctl00_txtPremium'])[1]
        ...    ${premium}
        Click Element When Visible
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnAddNewEntry'])[1]
        Wait Until Element Is Visible    (//legend[normalize-space()='Underlyers'])[1]    20s
    END
    Click Next And Wait For Element    (//legend[normalize-space()='Rating Input'])[1]
    ${hazard_level}=    Get From Dictionary    ${current_row}    AD
    Select Option By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlGLHazardLvl'])[1]
    ...    ${hazard_level}
    ${category}=    Get From Dictionary    ${current_row}    AE
    Select Option By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlCategory'])[1]
    ...    ${category}
    ${primary_g1_limit}=    Get From Dictionary    ${current_row}    AF
    Input Text Safely
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbPrimaryGlLimit'])[1]
    ...    ${primary_g1_limit}
    ${primary_g1_premium}=    Get From Dictionary    ${current_row}    AG
    Input Text Safely
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbPrimaryGlPremium'])[1]
    ...    ${primary_g1_premium}
    ${metholodgy}=    Get From Dictionary    ${current_row}    AH
    Select Option By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlMetholodgy'])[1]
    ...    ${metholodgy}
    Click Next And Wait For Element    (//legend[normalize-space()='Rating Summary'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ButtonCalculateUnderlyers'])[1]
    Wait Until Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ButtonCalculateLayers'])[1]
    ...    30s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ButtonCalculateLayers'])[1]
    Wait Until Element Is Visible
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewGlProducts_ctrl0_DropDownListGlDescription'])[1]
    ...    30s
    ${description}=    Get From Dictionary    ${current_row}    AI
    Select Option By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewGlProducts_ctrl0_DropDownListGlDescription'])[1]
    ...    ${description}
    ${creditdebit}=    Get From Dictionary    ${current_row}    AJ
    Select Option By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewGlProducts_ctrl0_DropDownListGlCreditDebit'])[1]
    ...    ${creditdebit}
    ${justification}=    Get From Dictionary    ${current_row}    AK
    Input Text Safely
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewGlProducts_ctrl0_TextBoxGlJustification'])[1]
    ...    ${justification}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ButtonCalculateScheduleMods'])[1]
    Wait Until Element Is Visible
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_TxtPrimaryPremiumComment'])[1]
    ...    30s
    Scroll Element Into View
    ...    //div[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_panelComments']
    ${primary_premium_comment}=    Get From Dictionary    ${current_row}    AL
    Input Text Safely
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_TxtPrimaryPremiumComment'])[1]
    ...    ${primary_premium_comment}
    ${deviation_comment}=    Get From Dictionary    ${current_row}    AM
    Input Text Safely
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_TxtIlfDeviationComment'])[1]
    ...    ${deviation_comment}
    ${final_pricing_comment}=    Get From Dictionary    ${current_row}    AN
    Input Text Safely
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_TxtFinalPricingRationaleComment'])[1]
    ...    ${final_pricing_comment}
    Click Next And Wait For Element
    ...    xpath=//span[contains(@id, 'labelStatus')]

Risk Creation[Reissue]
    Wait Until Element Is Visible    (//a[normalize-space()='New Policy Info'])[1]    20s
    Click Next And Wait For Element    (//a[normalize-space()='Additional Policy Info'])[1]
    Click Next And Wait For Element    (//legend[normalize-space()='Underlyers'])[1]
    Click Next And Wait For Element    (//legend[normalize-space()='Rating Input'])[1]
    Click Next And Wait For Element    (//legend[normalize-space()='Rating Summary'])[1]
    Wait Until Keyword Succeeds
    ...    3x
    ...    5s
    ...    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Risk Creation[CopyRisk]
    Wait Until Element Is Visible    (//a[normalize-space()='New Policy Info'])[1]    20s
    Click Next And Wait For Element    (//a[normalize-space()='Additional Policy Info'])[1]
    Click Next And Wait For Element    (//legend[normalize-space()='Underlyers'])[1]
    Click Next And Wait For Element    (//legend[normalize-space()='Rating Input'])[1]
    Click Next And Wait For Element    (//legend[normalize-space()='Rating Summary'])[1]
    Wait Until Keyword Succeeds
    ...    3x
    ...    5s
    ...    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Pre Bind Endorsement
    ${category_list}=    Get From Dictionary    ${current_row}    AO
    Select Option By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]
    ...    ${category_list}
    Wait Until Element Is Visible
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlManuscriptTemplates'])[1]
    ...    30s
    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    AP
    Select Option By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlManuscriptTemplates'])[1]
    ...    2
    Wait Until Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btAddEndorsement'])[1]
    ...    15s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btAddEndorsement'])[1]
    Wait Until Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_EndorsementStandardButtons_btnSubmit'])[1]
    ...    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_EndorsementStandardButtons_btnSubmit'])[1]
    Click Next And Wait For Element
    ...    (//legend[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_legTitle'])[1]
    FOR    ${index}    IN RANGE    1    10
        ${subjective_exist}=    Is Element Visible
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_repeaterSubjectivities_ctl0${index}_cbMandatory'])[1]
        IF    ${subjective_exist} == True
            ${subjective_checked}=    Is Checkbox Selected
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_repeaterSubjectivities_ctl0${index}_cbMandatory'])[1]
            IF    ${subjective_checked} == True
                Select Checkbox
                ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_repeaterSubjectivities_ctl0${index}_cbSatisfied'])[1]
            END
        ELSE
            BREAK
        END
    END
    FOR    ${index}    IN RANGE    10    20
        ${subjective_exist}=    Is Element Visible
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_repeaterSubjectivities_ctl${index}_cbMandatory'])[1]
        IF    ${subjective_exist} == True
            ${subjective_checked}=    Is Checkbox Selected
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_repeaterSubjectivities_ctl${index}_cbMandatory'])[1]
            IF    ${subjective_checked} == True
                Select Checkbox
                ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_repeaterSubjectivities_ctl${index}_cbSatisfied'])[1]
            END
        ELSE
            BREAK
        END
    END
    Click Next And Wait For Element    (//legend[normalize-space()='Quote Final'])[1]

Pre Bind Endorsement[Reissue]
    Click Next And Wait For Element
    ...    (//legend[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_legTitle'])[1]
    Click Next And Wait For Element    (//legend[normalize-space()='Quote Final'])[1]

Quote
    Wait Until Element Is Visible    (//legend[normalize-space()='Quote Final'])[1]    20s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    ${comments_under_quote}=    Get From Dictionary    ${current_row}    AQ
    Input Text Safely
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbComments'])[1]
    ...    ${comments_under_quote}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnGenerateQuote'])[1]

Ready To Bind
    Wait Until Element Is Visible    (//a[normalize-space()='Bind'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Bind'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlCheckSubjectivities_legTitle'])[1]
    ...    30s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s

Book
    [Arguments]    ${download_policy}
    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Book and Issue'])[1]
    Wait Until Element Is Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]    30s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    ${pollution_code}=    Get From Dictionary    ${current_row}    AR
    Select Option By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlBook_ddlPollutionCode'])[1]
    ...    ${pollution_code}
    ${days_to_cancel}=    Get From Dictionary    ${current_row}    AS
    Select Option By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlBook_ddlDaysToCancel'])[1]
    ...    ${days_to_cancel}
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[4]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[5]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[6]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonCheck'])[1]

Download Policy
    Wait Until Element Is Visible    (//a[normalize-space()='Risk UW eFile'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Risk UW eFile'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s
    Sleep    30s
    Wait Until Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnReload'])[1]
    ...    20s
    Click Element If Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnReload'])[1]
    Sleep    30s
    Wait Until Element Is Visible    (//td[@role='gridcell'])[5]//a    120s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnReload'])[1]
    IF    ${download_policy} == True
        Wait Until Element Is Visible    (//a[normalize-space()='Policy.pdf'])[1]    240s
        Click Element When Visible    (//a[normalize-space()='Policy.pdf'])[1]
        Click Element If Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    ELSE
        Sleep    240s
    END
    Click Element When Visible    xpath=//div[@class='RiskHeaderColumnOne']//div[1]

Copy Quote
    Wait Until Element Is Visible    (//a[normalize-space()='Copy Quote'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Copy Quote'])[1]
    Wait Until Element Is Visible
    ...    xpath=(//span[contains(@id, 'HeaderQuote')])[2]
    ...    10s
    Verify Status
    ...    xpath=(//span[contains(@id, 'HeaderQuote')])[2]
    ...    Option 2 - [Status : Quoted (In Revision)]
    Click Element When Visible
    ...    xpath=(//span[contains(@id, 'HeaderQuote')])[1]

View Risk
    Wait Until Element Is Visible    (//a[normalize-space()='View Risk'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='View Risk'])[1]
    Wait Until Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonViewCurrent'])[1]
    ...    30s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonViewCurrent'])[1]
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    20s
    Verify Status    (//legend[@id='fsMainContentLegend'])[1]    View Risk

Copy Risk
    Wait Until Element Is Visible    (//a[normalize-space()='Copy Risk'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Copy Risk'])[1]
    Wait Until Element Is Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[6]    30s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[6]
    Wait Until Element Is Visible
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    30s
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Incomplete Submission
    Fill Insured Details    False    False    True    True
    Risk Creation[CopyRisk]
    Pre Bind Endorsement
    Quote
    Ready To Bind
    Book    False

Post Bind Endorsement
    Wait Until Element Is Visible    (//a[normalize-space()='Endorsement'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Endorsement'])[1]
    ${category_list}=    Get From Dictionary    ${current_row}    AT
    Select Option By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]
    ...    ${category_list}
    Sleep    1s
    Wait Until Element Is Enabled
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdEndorsementTypeList'])[1]
    ...    30s
    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    AU
    Select Option By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdEndorsementTypeList'])[1]
    ...    ${endorsement_type_list}
    Sleep    2s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btAddEndorsement'])[1]
    Wait Until Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementEffectiveDate_datepickerEffectiveDate_textDate'])[1]
    ...    20s
    ${effective_date_post_bind}=    Get From Dictionary    ${current_row}    AW
    ${effective_date_post_bind}=    Format Excel Date    ${effective_date_post_bind}
    Input Text Safely
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementEffectiveDate_datepickerEffectiveDate_textDate'])[1]
    ...    ${effective_date_post_bind}
    ${Expiry_date}=    Get Current Date    increment=7 days    result_format=%m/%d/%Y
    Wait Until Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerPolicyExpiryDate_textDate'])[1]
    ...    15s
    Input Text Safely
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerPolicyExpiryDate_textDate'])[1]
    ...    ${Expiry_date}
    Sleep    3s
    Wait Until Keyword Succeeds
    ...    3x
    ...    2s
    ...    Click Element When Visible
    ...    (//span[normalize-space()='No Premium'])[1]
    Wait Until Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_EndorsementStandardButtons_btnSubmit'])[1]
    ...    15s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_EndorsementStandardButtons_btnSubmit'])[1]
    Sleep    8s
    Wait Until Keyword Succeeds
    ...    3x
    ...    2s
    ...    Click Element When Visible    xpath=//div[@class='RiskHeaderColumnOne']//div[1]

Renewal
    Wait Until Element Is Visible    (//a[normalize-space()='Renew'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Renew'])[1]
    Wait Until Element Is Visible
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    30s
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Submission
    Risk Creation    False    True    False    True
    Pre Bind Endorsement
    Quote
    Ready To Bind
    Book    False

Reissue
    Wait Until Element Is Visible    (//a[normalize-space()='Reissue'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Reissue'])[1]
    Wait Until Element Is Visible
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    30s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[7]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Quoted (Pending, In Revision, RI)
    Click Element When Visible
    ...    (//a[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_repeaterQuoteOptions_ctl01_quoteTab_btnEditQuote_lnkButton'])[1]
    Risk Creation[Reissue]
    Pre Bind Endorsement[Reissue]
    Quote
    Ready To Bind
    Book    False

Continue on Next Risk
    Wait Until Element Is Visible    xpath=//a[normalize-space()='Create New Risk >']    20s
    Click Element When Visible    xpath=//a[normalize-space()='Create New Risk >']
    Click Element When Visible    xpath=//a[normalize-space()='Create New Risk >']
    Wait Until Element Is Visible    xpath=//a[normalize-space()='US E&U Plus']    30s
    Click Element When Visible    xpath=//a[normalize-space()='US E&U Plus']
    Click Element When Visible
    ...    id:ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlClearanceSearch_CreateInsured
    Wait Until Element Is Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]    40s

Verify Status
    [Arguments]    ${location}    ${status_expected}
    Wait Until Keyword Succeeds    45x    2s    Element Should Contain    ${location}    ${status_expected}

Click Next And Wait For Element
    [Arguments]    ${expected_element_locator}
    Wait Until Keyword Succeeds    3x    1s    Attempt Next Page Transition    ${expected_element_locator}

Attempt Next Page Transition
    [Arguments]    ${expected_element_locator}
    Press Keys    None    PAGE_DOWN
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    ${expected_element_locator}    15s

Select Option By Label
    [Arguments]    ${locator}    ${label}
    Wait Until Keyword Succeeds    3x    2s    Attempt Select By Label    ${locator}    ${label}

Attempt Select By Label
    [Arguments]    ${locator}    ${label}
    Wait Until Element Is Enabled    ${locator}    15s
    Select From List By Label    ${locator}    ${label}

Select Option By Index
    [Arguments]    ${locator}    ${index}
    Wait Until Keyword Succeeds    3x    2s    Attempt Select By Index    ${locator}    ${index}

Attempt Select By Index
    [Arguments]    ${locator}    ${index}
    Wait Until Element Is Enabled    ${locator}    15s
    Select From List By Index    ${locator}    ${index}

Input Text Safely
    [Arguments]    ${locator}    ${text}
    Wait Until Keyword Succeeds    3x    2s    Attempt Input Text    ${locator}    ${text}

Attempt Input Text
    [Arguments]    ${locator}    ${text}
    Wait Until Element Is Visible    ${locator}    15s
    Input Text    ${locator}    ${text}

Format Excel Date
    [Arguments]    ${date_val}
    ${formatted}=    Evaluate    (lambda s: s[5:7]+'/'+s[8:10]+'/'+s[0:4] if '-' in s else (str(int(s.split('/')[0])).zfill(2)+'/'+str(int(s.split('/')[1])).zfill(2)+'/'+s.split('/')[2] if '/' in s else s))(str($date_val)) if not hasattr($date_val, 'strftime') else $date_val.strftime('%m/%d/%Y')
    [Return]    ${formatted}

Check All Subjectivities
    FOR    ${index}    IN RANGE    1    20
        ${idx}=    Convert To String    ${index}
        ${idx}=    Run Keyword If    ${index} < 10    Set Variable    0${index}    ELSE    Set Variable    ${index}
        
        ${mandatory_xpath}=    Set Variable    //input[contains(@id, 'repeaterSubjectivities_ctl${idx}_cbMandatory')]
        ${satisfied_xpath}=    Set Variable    //input[contains(@id, 'repeaterSubjectivities_ctl${idx}_cbSatisfied')]
        
        ${exists}=    Run Keyword And Return Status    Page Should Contain Element    ${mandatory_xpath}
        IF    ${exists} == True
            ${mandatory_checked}=    Execute Javascript    return document.evaluate("${mandatory_xpath}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.checked;
            IF    ${mandatory_checked} == False
                Execute Javascript    document.evaluate("${mandatory_xpath}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();
                Sleep    1s
            END
            
            ${satisfied_checked}=    Execute Javascript    return document.evaluate("${satisfied_xpath}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.checked;
            IF    ${satisfied_checked} == False
                Execute Javascript    document.evaluate("${satisfied_xpath}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();
                Sleep    1s
            END
        ELSE
            BREAK
        END
    END
