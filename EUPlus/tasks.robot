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

US E&U Plus

    Open Workbook    ${DATA_SOURCE}

    ${excel_rows}=    Read Worksheet    EUPlus

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

    Risk Creation    True    False    False

    Verify Status

    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_labelStatus'])[1]

    ...    Submission (Pricing In Progress)

    Pre Bind Endorsement

    Quote

    Verify Status

    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]

    ...    Quoted

    Copy Quote

    Ready To Bind

    Verify Status

    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]

    ...    Bound (Pending)

    Book    False

    Verify Status

    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]

    ...    Bound

    Copy Risk    True

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

    Click Element When Visible    xpath=//a[normalize-space()='US E&U Plus']

    Sleep    10s

    Click Element When Visible

    ...    id:ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlClearanceSearch_CreateInsured

    Sleep    5s

    ${EMAIL}=    Set Variable    ${NUSER}@libertymutual.com

    Wait Until Element Is Visible    xpath=//input[@type='email']    30s

    Clear Element Text    xpath=//input[@type='email']

    Input Text    xpath=//input[@type='email']    ${EMAIL}

    Sleep    1s

    Press Keys    xpath=//input[@type='email']    ENTER

    Wait Until Element Is Visible    xpath=//input[@type='password']    30s

    Clear Element Text    xpath=//input[@type='password']

    Input Text    xpath=//input[@type='password']    ${NUSERPASSWORD}

    Sleep    1s

    Press Keys    xpath=//input[@type='password']    ENTER

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

        Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]

        Sleep    1s

        Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[3]

    END

    IF    ${firstRun} == True

        ${iname}=    Get From Dictionary    ${current_row}    D

        Input Text

        ...    //input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_TextBoxFirstNamedInsured1']

        ...    ${iname}

        ${state}=    Get From Dictionary    ${current_row}    E

        Select From List By Label

        ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_ddUSState'])[1]

        ...    ${state}

        ${adl1}=    Get From Dictionary    ${current_row}    F

        Input Text

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSAddressLine1'])[1]

        ...    ${adl1}

        ${city1}=    Get From Dictionary    ${current_row}    G

        Select From List By Label

        ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_ddUSCity'])[1]

        ...    ${city1}

        ${zip1}=    Get From Dictionary    ${current_row}    H

        Click Element When Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'])[1]

        ${zipelm}=    Get WebElement

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'])[1]

        ${attribute}=    Get Element Attribute

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'])[1]

        ...    value

        Input Text

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

        Input Text

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSPoBox'])[1]

        ...    ${pobox}

    END

    Click Next And Wait For Element

    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]

    Click Next And Wait For Element    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]

    Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]

    ...    LIUI

    Select From List By Index

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]

    ...    1

    IF    ${firstRun} == True

        ${effdate}=    Get From Dictionary    ${current_row}    J

        Input Text

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]

        ...    ${effdate}

        ${Indcode}=    Get From Dictionary    ${current_row}    K

        Input Text

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]

        ...    ${Indcode}

    END

    IF    ${copyRisk} == True

        ${effdate1}=    Get From Dictionary    ${current_row}    L

        Input Text

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]

        ...    ${effdate1}

        Input Text

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]

        ...    4491 Marine cargo handling

    END

    Click Next And Wait For Element    (//legend[normalize-space()='Broker Info'])[1]

    IF    ${firstRun} == True

        ${brokfirm}=    Get From Dictionary    ${current_row}    M

        Input Text

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_TextBoxBrokerFirm'])[1]

        ...    ${brokfirm}

        Click Element When Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_ButtonBrokerFirmSearchStartsWith'])[1]

        Wait Until Element Is Visible

        ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]

        ...    30s

        ${table_elements}=    Get WebElements

        ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]

        Log    Found ${table_elements.__len__()} elements in the broker info table

        Click Element    ${table_elements}[2]

        Sleep    5s

        ${value1}=    Is Element Visible    (//span[normalize-space()='Yes'])[1]

        IF    ${value1} == True

            Click Element When Visible    (//span[normalize-space()='Yes'])[1]

            Sleep    2s

        END

    END

    Execute Javascript    window.scrollTo(0, document.body.scrollHeight)

    Wait Until Keyword Succeeds

    ...    3x

    ...    5s

    ...    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    Sleep    5s

    ${variableExist}=    Run Keyword And Return Status

    ...    Element Should Be Visible

    ...    (//span[normalize-space()='Continue'])[1]

    IF    ${variableExist} == True

        Wait Until Element Is Visible    (//span[normalize-space()='Continue'])[1]    10s

        Click Element When Visible    (//span[normalize-space()='Continue'])[1]

    END

    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s

    IF    ${firstRun} == True

        ${brokcont}=    Get From Dictionary    ${current_row}    N

        Input Text

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]

        ...    ${brokcont}

        Click Element When Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_ButtonBrokerContactSearchStartsWith'])[1]

        Wait Until Element Is Visible

        ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]

        ...    30s

        ${table_elements}=    Get WebElements

        ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]

        Log    Found ${table_elements.__len__()} elements in the broker contact table

        Click Element    ${table_elements}[0]

        Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]

    END

    IF    ${renewal} == True

        ${brokcont}=    Get From Dictionary    ${current_row}    N

        Input Text

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]

        ...    ${brokcont}

        Click Element When Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_ButtonBrokerContactSearchStartsWith'])[1]

        Wait Until Element Is Visible

        ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]

        ...    30s

        ${table_elements}=    Get WebElements

        ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]

        Log    Found ${table_elements.__len__()} elements in the broker contact table

        Click Element    ${table_elements}[0]

        Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]

    END

    ${EditVisible}=    Is Element Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_repeaterBrokerContact_ctl01_buttonEditBroker'])[1]

    IF    ${EditVisible} == True

        Click Element When Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_repeaterBrokerContact_ctl01_buttonEditBroker'])[1]

        ${PCLicHolder}=    Is Checkbox Selected

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_chkPAndCLicenseHolder'])[1]

        IF    ${PCLicHolder} == False

            Select Checkbox

            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_chkPAndCLicenseHolder'])[1]

            Click Element When Visible

            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]

            Press Keys    None    PAGE_DOWN

            Wait Until Keyword Succeeds

            ...    3x

            ...    5s

            ...    Click Element When Visible

            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

        END

    ELSE

        Click Element When Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]

        Press Keys    None    PAGE_DOWN

        Wait Until Keyword Succeeds

        ...    3x

        ...    5s

        ...    Click Element When Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    END

    Sleep    10s

    ${value2}=    Is Element Visible    (//span[normalize-space()='Continue'])[1]

    IF    ${value2} == True

        Click Element When Visible    (//span[normalize-space()='Continue'])[1]

    END


Fill Pricing Details

    [Arguments]    ${firstRun}    ${renewal}

    Wait Until Element Is Visible    (//legend[normalize-space()='Expiring Policy'])[1]    20s

    Wait Until Keyword Succeeds

    ...    3x

    ...    5s

    ...    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    IF    ${renewal} == True

        ${value2}=    Is Element Visible    (//span[@class='ui-button-text'][normalize-space()='Ok'])[1]

        IF    ${value2} == True

            Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Ok'])[1]

        END

    END

    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info'])[1]

    ${form}=    Get From Dictionary    ${current_row}    O

    Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlForm'])[1]

    ...    ${form}

    ${attachment}=    Get From Dictionary    ${current_row}    P

    Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlAttachment'])[1]

    ...    ${attachment}

    ${industry}=    Get From Dictionary    ${current_row}    Q

    Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlIndustry'])[1]

    ...    ${industry}

    Click Element When Visible    (//span[normalize-space()='Non Project'])[1]

    Sleep    10s

    ${industry_segment}=    Get From Dictionary    ${current_row}    R

    Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlIndustrySegment'])[1]

    ...    ${industry_segment}

    Click Element When Visible    (//span[normalize-space()='Correct'])[1]

    ${category}=    Get From Dictionary    ${current_row}    S

    Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlCategory'])[1]

    ...    ${category}

    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]

    ${quote_rater}=    Get From Dictionary    ${current_row}    T

    Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlQuoteRater'])[1]

    ...    ${quote_rater}

    ${exposure1}=    Get From Dictionary    ${current_row}    U

    Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_txtExposure'])[1]

    ...    ${exposure1}

    ${exposure_base}=    Get From Dictionary    ${current_row}    V

    Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlExposureBase'])[1]

    ...    ${exposure_base}

    ${premium_basis}=    Get From Dictionary    ${current_row}    W

    Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_ddlPremiumBasis'])[1]

    ...    ${premium_basis}

    ${min_earned_premium}=    Get From Dictionary    ${current_row}    X

    Input Text When Element Is Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_txtMinEarnedPremium'])[1]

    ...    ${min_earned_premium}

    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[3]

    ${total_limit_each_occur}=    Get From Dictionary    ${current_row}    Y

    Input Text When Element Is Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_txtTotalLimitEachOccur'])[1]

    ...    ${total_limit_each_occur}

    ${attachment_point}=    Get From Dictionary    ${current_row}    Z

    Input Text When Element Is Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_txtAttachmentPoint'])[1]

    ...    ${attachment_point}

    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrPolicyInfo_cblGaragedStates_0'])[1]

    Click Next And Wait For Element    (//legend[normalize-space()='Underlyers'])[1]

    IF    ${firstRun} == True

        Click Element When Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_underLyerSched_underLyerList_12'])[1]

        Click Element When Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_underLyerSched_ctrlCommercialExcessLiability_btnAdd'])[1]

        Wait Until Element Is Visible

        ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctl00_lblHeader'])[1]

        ...    20s

        ${carrier}=    Get From Dictionary    ${current_row}    AA

        Input Text When Element Is Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctl00_txtCarrier'])[1]

        ...    ${carrier}

        ${excess_of}=    Get From Dictionary    ${current_row}    AB

        Input Text When Element Is Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctl00_tbExcessOf'])[1]

        ...    ${excess_of}

        ${policy_number}=    Get From Dictionary    ${current_row}    AC

        Input Text When Element Is Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctl00_txtPolicyNumber'])[1]

        ...    ${policy_number}

        ${premium}=    Get From Dictionary    ${current_row}    AD

        Input Text When Element Is Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctl00_txtPremium'])[1]

        ...    ${premium}

        Click Element When Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnAddNewEntry'])[1]

    END

    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_underLyerSched_ctrlCommercialExcessLiability_underlyerRepeater_ctl01_btnEdit'])[1]

    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[2]

    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnAddNewEntry'])[1]

    Click Next And Wait For Element

    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='GL Info'])[1]

    ${premesis_ops_occurence}=    Get From Dictionary    ${current_row}    AE

    Input Text When Element Is Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGLInfo_txtPremisesOpsOccurence'])[1]

    ...    ${premesis_ops_occurence}

    ${hazard_type}=    Get From Dictionary    ${current_row}    AF

    Select From List By Index

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGLInfo_ddlGLHazardType'])[1]

    ...    ${hazard_type}

    ${rating_type}=    Get From Dictionary    ${current_row}    AG

    Select From List By Index

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGLInfo_ddlGLRatingType'])[1]

    ...    ${hazard_type}

    ${underlying_primary}=    Get From Dictionary    ${current_row}    AH

    Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGLInfo_txtUnderlyingPrimary'])[1]

    ...    ${underlying_primary}

    ${primary_treat_def}=    Get From Dictionary    ${current_row}    AI

    Select From List By Index

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGLInfo_ddlPrimaryTreatDef'])[1]

    ...    ${primary_treat_def}

    Click Next And Wait For Element    (//legend[normalize-space()='Exposure Information'])[1]

    ${industry_code}=    Get From Dictionary    ${current_row}    AJ

    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_formGLExposures_listViewProspectivepolicyclasscode_ctrl0_TextBoxIndustryCode'])[1]
    ...    10120
    Sleep    3s
    Press Keys    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_formGLExposures_listViewProspectivepolicyclasscode_ctrl0_TextBoxIndustryCode'])[1]    ARROW_DOWN
    Sleep    1s
    Press Keys    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_formGLExposures_listViewProspectivepolicyclasscode_ctrl0_TextBoxIndustryCode'])[1]    ENTER

    Sleep    2s

    ${state}=    Get From Dictionary    ${current_row}    AK

    Select From List By Index

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_formGLExposures_listViewProspectivepolicyclasscode_ctrl0_ddlState'])[1]

    ...    ${state}

    ${territory}=    Get From Dictionary    ${current_row}    AL

    Select From List By Index

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_formGLExposures_listViewProspectivepolicyclasscode_ctrl0_ddlTerritory'])[1]

    ...    ${territory}

    ${amount}=    Get From Dictionary    ${current_row}    AM

    Input Text When Element Is Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_formGLExposures_listViewProspectivepolicyclasscode_ctrl0_txtAmount'])[1]

    ...    ${amount}

    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_formGLExposures_btnGetDefaults'])[1]

    Sleep    5s

    Click Next And Wait For Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    Click Next And Wait For Element    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]

    Click Next And Wait For Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    Click Next And Wait For Element    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]

    Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAutoLiability_ddlPrimaryAutoRetention'])[1]

    ...    Guaranteed Costs

    Input Text When Element Is Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAutoLiability_txtPriAutoLimit'])[1]

    ...    500K

    Select From List By Index

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAutoLiability_ddlSigHighHazardExposure'])[1]

    ...    1

    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAutoLiability_btnCalculateAuto'])[1]

    Wait Until Keyword Succeeds

    ...    3x

    ...    5s

    ...    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s

    ${al_exposure}=    Get From Dictionary    ${current_row}    BF

    FOR    ${index}    IN RANGE    0    5

        ${is_element_visible}=    Is Element Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrAlExposure_listViewExposure_ctrl${index}_TextBoxALExposureAmount'])[1]

        IF    ${is_element_visible} == True

            Input Text When Element Is Visible

            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrAlExposure_listViewExposure_ctrl${index}_TextBoxALExposureAmount'])[1]

            ...    10000

        END

    END

    Wait Until Keyword Succeeds

    ...    3x

    ...    5s

    ...    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    Wait Until Element Is Visible    (//legend[normalize-space()='AL Losses'])[1]    20s

    Sleep    1s

    ${loss_valuation_date}=    Get From Dictionary    ${current_row}    BF

    FOR    ${index}    IN RANGE    0    5

        ${locator}=    Set Variable

        ...    //input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAlLosses_listViewGroundUpLosses_ctrl${index}_datePickerLossValuationDate_textDate']

        ${is_visible}=    Run Keyword And Return Status    Element Should Be Visible    ${locator}

        IF    ${is_visible}

            ${elm}=    Get WebElement    ${locator}

            Execute Javascript    arguments[0].value = arguments[1];    ARGUMENTS    ${elm}    ${loss_valuation_date}

        END

    END

    Press Keys

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAlLosses_listViewGroundUpLosses_ctrl4_datePickerLossValuationDate_textDate'])[1]

    ...    TAB

    Wait Until Keyword Succeeds

    ...    3x

    ...    5s

    ...    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    Click Element If Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[2]

    Click Next And Wait For Element    (//legend[normalize-space()='Rating Summary'])[1]

    Sleep    10s

    Wait Until Keyword Succeeds
    ...    3x
    ...    2s
    ...    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_btnCalculateLayers'])[1]

    Sleep    10s

    ${primary_premium_comment}=    Get From Dictionary    ${current_row}    BG

    Wait Until Keyword Succeeds
    ...    3x
    ...    2s
    ...    Input Text When Element Is Visible
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtPrimaryPremiumComment'])[1]
    ...    ${primary_premium_comment}

    ${deviation_comment}=    Get From Dictionary    ${current_row}    BH

    Wait Until Keyword Succeeds
    ...    3x
    ...    2s
    ...    Input Text When Element Is Visible
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtIlfDeviationComment'])[1]
    ...    ${deviation_comment}

    ${final_pricing_comment}=    Get From Dictionary    ${current_row}    BI

    Wait Until Keyword Succeeds
    ...    3x
    ...    2s
    ...    Input Text When Element Is Visible
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtFinalPricingRationaleComment'])[1]
    ...    ${final_pricing_comment}

    Click Next And Wait For Element    (//a[normalize-space()='State Info'])[1]

    Sleep    2s

    Wait Until Keyword Succeeds
    ...    3x
    ...    2s
    ...    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]

    Wait Until Keyword Succeeds
    ...    3x
    ...    2s
    ...    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[2]

    Click Next And Wait For Element    (//a[normalize-space()='Rate Monitor'])[1]

    ${rate_monitor_coment}=    Get From Dictionary    ${current_row}    BI

    Input Text When Element Is Visible

    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRateMonitor_txtRateMonitorComment'])[1]

    ...    ${rate_monitor_coment}

    Wait Until Keyword Succeeds

    ...    3x

    ...    5s

    ...    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]


Risk Creation[CopyRisk]

    Wait Until Element Is Visible    (//legend[normalize-space()='Expiring Policy'])[1]    20s

    Click Next And Wait For Element    (//a[normalize-space()='Policy Info'])[1]

    Click Element When Visible    (//span[normalize-space()='Correct'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Schedule of Underlyers'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='GL Info'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='GL Exposures'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='GL Losses'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Underlying GL Loss Rating'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Underlying Primary GL Selection'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='AL Selection'])[1]

    ${high_hazard_exposure}=    Get From Dictionary    ${current_row}    AR

    Select From List By Index

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAutoLiability_ddlSigHighHazardExposure'])[1]

    ...    ${high_hazard_exposure}

    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAutoLiability_btnCalculateAuto'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='AL Exposures'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='AL Losses'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Rating Summary'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='State Info'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Rate Monitor'])[1]

    Wait Until Keyword Succeeds

    ...    3x

    ...    5s

    ...    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]


Risk Creation[Reissue]

    FOR    ${i}    IN RANGE    12

        Press Keys    None    PAGE_DOWN

        Wait Until Keyword Succeeds

        ...    3x

        ...    5s

        ...    Click Element When Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

        Sleep    3s

    END


Pre Bind Endorsement

    ${category_list}=    Get From Dictionary    ${current_row}    BJ

    Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]

    ...    ${category_list}

    Sleep    10s

    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    BK

    Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdEndorsementTypeList'])[1]

    ...    ${endorsement_type_list}

    Sleep    2s

    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btAddEndorsement'])[1]

    Sleep    3s

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

    Wait Until Keyword Succeeds

    ...    3x

    ...    5s

    ...    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]


Pre Bind Endorsement[Reissue]

    Click Next And Wait For Element

    ...    (//legend[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_legTitle'])[1]

    Click Next And Wait For Element    (//legend[normalize-space()='Quote Final Info'])[1]


Pre Bind Endorsement[Renewal]

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

    Wait Until Keyword Succeeds

    ...    3x

    ...    5s

    ...    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]


Quote

    Wait Until Element Is Visible    (//legend[normalize-space()='Quote Final Info'])[1]    20s

    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]

    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]

    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[3]

    ${comments_under_quote}=    Get From Dictionary    ${current_row}    BL

    Input Text When Element Is Visible

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

    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s

    Sleep    2s

    Wait Until Keyword Succeeds
    ...    3x
    ...    2s
    ...    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]

    Wait Until Keyword Succeeds
    ...    3x
    ...    2s
    ...    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]

    Wait Until Keyword Succeeds
    ...    3x
    ...    2s
    ...    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]


Ready To Bind[Renewal]

    Wait Until Element Is Visible    (//a[normalize-space()='Bind'])[1]    30s

    Click Element When Visible    (//a[normalize-space()='Bind'])[1]

    Wait Until Element Is Visible

    ...    (//legend[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlCheckSubjectivities_legTitle'])[1]

    ...    30s

    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]

    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s

    Sleep    2s

    Wait Until Keyword Succeeds
    ...    3x
    ...    2s
    ...    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]

    Wait Until Keyword Succeeds
    ...    3x
    ...    2s
    ...    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]

    Wait Until Keyword Succeeds
    ...    3x
    ...    2s
    ...    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]


Book

    [Arguments]    ${download_policy}

    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s

    Click Element When Visible    (//a[normalize-space()='Book and Issue'])[1]

    Wait Until Element Is Visible    (//legend[normalize-space()='Reinsurance'])[1]    30s

    ${days_to_cancel}=    Get From Dictionary    ${current_row}    BM

    Select From List By Index

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlBook_ddlDaysToCancel'])[1]

    ...    ${days_to_cancel}

    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]

    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[3]

    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[4]

    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonCheck'])[1]

    Sleep    10s


Risk UW eFile

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

    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_repeaterQuoteOptions_ctl00_HeaderQuote'])[2]

    ...    10s

    Verify Status

    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_repeaterQuoteOptions_ctl00_HeaderQuote'])[2]

    ...    Option 2 - [Status : Quoted (In Revision)]

    Click Element When Visible

    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_repeaterQuoteOptions_ctl00_HeaderQuote'])[1]


View Risk

    Wait Until Element Is Visible    (//a[normalize-space()='View Risk'])[1]    30s

    Click Element When Visible    (//a[normalize-space()='View Risk'])[1]

    Sleep    5s

    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonViewCurrent'])[1]

    Sleep    5s

    Verify Status    (//legend[@id='fsMainContentLegend'])[1]    View Risk


Copy Risk

    [Arguments]    ${continue}

    Wait Until Element Is Visible    (//a[normalize-space()='Copy Risk'])[1]    30s

    Click Element When Visible    (//a[normalize-space()='Copy Risk'])[1]

    Sleep    5s

    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[5]

    Sleep    10s

    Verify Status

    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_labelStatus'])[1]

    ...    Incomplete Submission

    IF    ${continue} == True

        Fill Insured Details    False    False    True

        Risk Creation[CopyRisk]

        Pre Bind Endorsement

        Quote

        Ready To Bind

        Book    False

    ELSE

        Click Element When Visible    xpath=//div[@class='RiskHeaderColumnOne']//div[1]

    END


Post Bind Endorsement

    Wait Until Element Is Visible    (//a[normalize-space()='Endorsement'])[1]    30s

    Click Element When Visible    (//a[normalize-space()='Endorsement'])[1]

    ${category_list}=    Get From Dictionary    ${current_row}    BN

    Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]

    ...    ${category_list}

    Sleep    10s

    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    BO

    Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdEndorsementTypeList'])[1]

    ...    ${endorsement_type_list}

    Sleep    2s

    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btAddEndorsement'])[1]

    Sleep    3s

    Click Element When Visible    (//span[normalize-space()='Additional Premium'])[1]

    ${endorsement_premium}=    Get From Dictionary    ${current_row}    BP

    Input Text When Element Is Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementPremium_tbEndorsementPremium'])[1]

    ...    ${endorsement_premium}

    ${effective_date_post_bind}=    Get From Dictionary    ${current_row}    BQ

    Input Text When Element Is Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementEffectiveDate_datepickerEffectiveDate_textDate'])[1]

    ...    ${effective_date_post_bind}

    TRY

        ${Expiry_date}=    Get From Dictionary    ${current_row}    BR

        Input Text When Element Is Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerPolicyExpiryDate_textDate'])[1]

        ...    ${Expiry_date}

    EXCEPT

        ${Expiry_date}=    Get From Dictionary    ${current_row}    BR

        Input Text When Element Is Visible

        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerPolicyExpiryDate_textDate'])[1]

        ...    ${Expiry_date}

    END

    Sleep    3s

    TRY

        Click Element When Visible    (//span[normalize-space()='No Premium'])[1]

    EXCEPT

        Click Element When Visible    (//span[normalize-space()='No Premium'])[1]

    END

    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_EndorsementStandardButtons_btnSubmit'])[1]

    Click Element When Visible    xpath=//div[@class='RiskHeaderColumnOne']//div[1]


Renewal

    Wait Until Element Is Visible    (//a[normalize-space()='Renew'])[1]    30s

    Click Element When Visible    (//a[normalize-space()='Renew'])[1]

    Sleep    10s

    Verify Status

    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_labelStatus'])[1]

    ...    Submission

    Risk Creation    False    True    False

    Pre Bind Endorsement[Renewal]

    Quote

    Ready To Bind

    Book    False


Reissue

    Wait Until Element Is Visible    (//a[normalize-space()='Reissue'])[1]    30s

    Click Element When Visible    (//a[normalize-space()='Reissue'])[1]

    Sleep    10s

    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[7]

    Sleep    5s

    Verify Status

    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]

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

    Wait Until Keyword Succeeds    10x    3s    Element Text Should Be    ${location}    ${status_expected}


Click Next And Wait For Element

    [Arguments]    ${expected_element_locator}

    Wait Until Keyword Succeeds    3x    1s    Attempt Next Page Transition    ${expected_element_locator}


Attempt Next Page Transition

    [Arguments]    ${expected_element_locator}

    Press Keys    None    PAGE_DOWN

    Click Element When Visible

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    4s
    Wait Until Element Is Visible    ${expected_element_locator}    15s

