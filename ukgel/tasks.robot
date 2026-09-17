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
UK GEL
    Open Workbook    ${DATA_SOURCE}
    ${excel_rows}=    Read Worksheet    UKGel
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
    Risk Creation[First Run]
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
    Sleep    100s
    Renewal
    View Risk
    Close Browser

Login to App
    ${URL}=    Get From Dictionary    ${current_row}    A
    ${NUSER}=    Get From Dictionary    ${current_row}    B
    ${NUSERPASSWORD}=    Get From Dictionary    ${current_row}    C
    Open Browser    ${URL}    edge    executable_path=${DRIVER}    options=add_argument("--inprivate")
    Maximize Browser Window
    Set Selenium Implicit Wait    30s
    Safe Click Element    xpath=//a[normalize-space()='Create New Risk >']
    Safe Click Element    xpath=//a[normalize-space()='UK GEL']
    Sleep    5s
    Safe Click Element
    ...    id:ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlClearanceSearch_CreateInsured
    Sleep    5s
    ${EMAIL}=    Set Variable    ${NUSER}@libertymutual.com
    Wait Until Element Is Visible    xpath=//input[@type='email']    30s
    Clear Element Text    xpath=//input[@type='email']
    Safe Input Text    xpath=//input[@type='email']    ${EMAIL}
    Sleep    1s
    Press Keys    xpath=//input[@type='email']    ENTER
    Wait Until Element Is Visible    xpath=//input[@type='password']    30s
    Clear Element Text    xpath=//input[@type='password']
    Safe Input Text    xpath=//input[@type='password']    ${NUSERPASSWORD}
    Sleep    1s
    Press Keys    xpath=//input[@type='password']    ENTER
    Sleep    5s

Risk Creation[First Run]
    Fill Insured Details[First Run]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Submission
    Fill Pricing Details[First Run]

Fill Insured Details[First Run]
    Safe Click Element    (//span[normalize-space()='No'])[1]
    ${iname}=    Get From Dictionary    ${current_row}    D
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_TextBoxFirstNamedInsured1'])[1]
    ...    ${iname}
    ${adl1}=    Get From Dictionary    ${current_row}    E
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUKAddressLine1'])[1]
    ...    ${adl1}
    ${zip1}=    Get From Dictionary    ${current_row}    F
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUKPostCode'])[1]
    ${zipelm}=    Get WebElement
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUKPostCode'])[1]
    ${attribute}=    Get Element Attribute
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUKPostCode'])[1]
    ...    value
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUKPostCode'])[1]
    ...    ${zip1}
    Press Keys
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUKPostCode'])[1]
    ...    A+C+1+3+ +2+A+A
    ${return_value}=    Execute Javascript    arguments[0].value='AC13 2AA'    ARGUMENTS    ${zipelm}
    ${attribute}=    Get Element Attribute
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUKPostCode'])[1]
    ...    value
    Sleep    1s
    ${pobox}=    Get From Dictionary    ${current_row}    G
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUKPoBox'])[1]
    ...    ${pobox}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListBranch'])[1]
    ...    Bristol
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]
    ...    BRISTLUK
    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    ${effdate}=    Get From Dictionary    ${current_row}    H
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
    ...    ${effdate}
    ${Indcode}=    Get From Dictionary    ${current_row}    I
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]
    ...    ${Indcode}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Store Risk Number    First Run
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    ${brokfirm}=    Get From Dictionary    ${current_row}    K
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_TextBoxBrokerFirm'])[1]
    ...    ${brokfirm}
    Safe Click Element
    ...    (//input[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_ButtonBrokerFirmSearchStartsWith"])[1]
    Wait Until Element Is Visible
    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]
    ...    30s
    ${table_elements}=    Get WebElements
    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]
    Log    Found ${table_elements.__len__()} elements in the broker info table
    Safe Click Element
    ...    xpath=(//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"])[3]
    Sleep    3s
    Execute Javascript    window.scrollTo(0, document.body.scrollHeight)
    ${value1}=    Is Element Visible    (//span[normalize-space()='Yes'])[1]
    IF    ${value1} == True
        Safe Click Element    (//span[normalize-space()='Yes'])[1]
    END
    Execute Javascript    window.scrollTo(0, document.body.scrollHeight)
    Press Keys    None    PAGE_DOWN
    Sleep    2s
    Wait Until Keyword Succeeds
    ...    5x
    ...    5s
    ...    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s
    ${variableExist}=    Run Keyword And Return Status
    ...    Element Should Be Visible
    ...    (//span[normalize-space()='Continue'])[1]
    IF    ${variableExist} == True
        Wait Until Element Is Visible    (//span[normalize-space()='Continue'])[1]    10s
        Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    ${brokcont}=    Get From Dictionary    ${current_row}    L
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]
    ...    ${brokcont}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_ButtonBrokerContactSearchStartsWith'])[1]
    Wait Until Element Is Visible
    ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]
    ...    30s
    ${table_elements}=    Get WebElements
    ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]
    Log    Found ${table_elements.__len__()} elements in the broker contact table
    Safe Click Element
    ...    xpath=(//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"])[2]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]
    ${EditVisible}=    Is Element Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_repeaterBrokerContact_ctl01_buttonEditBroker'])[1]
    IF    ${EditVisible} == True
        Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_repeaterBrokerContact_ctl01_buttonEditBroker'])[1]
        ${PCLicHolder}=    Is Checkbox Selected
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_chkPAndCLicenseHolder'])[1]
        IF    ${PCLicHolder} == False
            Safe Select Checkbox
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_chkPAndCLicenseHolder'])[1]
            Safe Click Element
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]
            Press Keys    None    PAGE_DOWN
            Wait Until Keyword Succeeds
            ...    3x
            ...    5s
            ...    Safe Click Element
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
        END
    ELSE
        Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]
        Press Keys    None    PAGE_DOWN
        Wait Until Keyword Succeeds
        ...    3x
        ...    5s
        ...    Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    END
    Sleep    10s
    ${value2}=    Is Element Visible    (//span[normalize-space()='Continue'])[1]
    IF    ${value2} == True
        Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s

Fill Pricing Details[First Run]
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    ${documentation_type}=    Get From Dictionary    ${current_row}    M
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_GeneralInfo_ddlDocumentationType'])[1]
    ...    ${documentation_type}
    ${line_of_business}=    Get From Dictionary    ${current_row}    N
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_GeneralInfo_ddlLOB'])[1]
    ...    ${line_of_business}
    Sleep    2s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    ${currency}=    Get From Dictionary    ${current_row}    O
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_GeneralInfo_ddlCurrency'])[1]
    ...    ${currency}
    Sleep    2s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[5]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Policy Info'])[1]    10s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Individual Claims'])[2]
    ${claim_experience_date}=    Get From Dictionary    ${current_row}    P
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_eLPolicyInfo_ctrlPolicyInfoEL_dpClaimsExperienceDateForNoExperienceRating_textDate'])[1]
    ...    ${claim_experience_date}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Exposures'])[1]    10s
    ${industry}=    Get From Dictionary    ${current_row}    Q
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlMainIndustry'])[1]
    ...    ${industry}
    Sleep    10s
    ${sub_industry}=    Get From Dictionary    ${current_row}    R
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlMainSubIndustry'])[1]
    ...    ${sub_industry}
    ${exposure}=    Get From Dictionary    ${current_row}    S
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewOccupations_ctrl0_ddlExposureType'])[1]
    ...    ${exposure}
    ${exposure_amount}=    Get From Dictionary    ${current_row}    T
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewOccupations_ctrl0_tbWageroll'])[1]
    ...    ${exposure_amount}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    ${risk_management}=    Get From Dictionary    ${current_row}    U
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlRiskManagementGrade'])[1]
    ...    ${risk_management}
    ${claim_history}=    Get From Dictionary    ${current_row}    V
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbClaimsHistoryDiscount'])[1]
    ...    ${claim_history}
    ${comments}=    Get From Dictionary    ${current_row}    W
    Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbDiscountComments'])[1]
    ...    ${comments}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Hist Exposure'])[1]    20s
    ${exposure_type}=    Get From Dictionary    ${current_row}    X
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_HistoricalExposureForm_ddlExposureType'])[1]
    ...    ${exposure_type}
    ${exposure_amount}=    Get From Dictionary    ${current_row}    Y
    Sleep    5s
    Safe Input Exposure Amount    xpath=(//input[contains(@id, 'txtExposureAmount')])[1]    ${exposure_amount}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Individual Claims'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_IndividualClaimsControl_radGridIndividualClaims_ctl00_ctl03_ctl01_btnInsert'])[1]
    Sleep    5s
    ${policy_year}=    Get From Dictionary    ${current_row}    Z
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_IndividualClaimsControl_radGridIndividualClaims_ctl00_ctl04_txtPolicyYear'])[1]
    ...    ${policy_year}
    ${incident_date}=    Get From Dictionary    ${current_row}    AA
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_IndividualClaimsControl_radGridIndividualClaims_ctl00_ctl04_txtIncidentDate_textDate'])[1]
    ...    ${incident_date}
    ${currency}=    Get From Dictionary    ${current_row}    AB
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_IndividualClaimsControl_radGridIndividualClaims_ctl00_ctl04_ddlOrigCurrency'])[1]
    ...    ${currency}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Override'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    AC
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_OverrideControl_rptLiabilityOverrideOccupation_ctl01_tbRateOverride'])[1]
    ...    ${rate_override}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Policy Info'])[1]    10s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element    (//span[normalize-space()='Occurrence'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Individual Claims'])[2]
    ${claim_experience_date}=    Get From Dictionary    ${current_row}    AD
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_pLPolicyInfo_ctrlPolicyInfoPL_dpClaimsExperienceDateForNoExperienceRating_textDate'])[1]
    ...    ${claim_experience_date}
    ${retention_type}=    Get From Dictionary    ${current_row}    AE
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_pLPolicyInfo_ctrlPolicyInfoPL_ddlRetentionType'])[1]
    ...    ${retention_type}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Exposures'])[1]    10s
    ${exposure1}=    Get From Dictionary    ${current_row}    AF
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewOccupations_ctrl0_ddlExposureType'])[1]
    ...    ${exposure1}
    ${exposure_amount1}=    Get From Dictionary    ${current_row}    AG
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewOccupations_ctrl0_tbWageroll'])[1]
    ...    ${exposure_amount1}
    ${exposure2}=    Get From Dictionary    ${current_row}    AH
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewOccupations_ctrl1_ddlExposureType'])[1]
    ...    ${exposure2}
    ${exposure_amount2}=    Get From Dictionary    ${current_row}    AI
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewOccupations_ctrl1_tbWageroll'])[1]
    ...    ${exposure_amount2}
    ${exposure3}=    Get From Dictionary    ${current_row}    AJ
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewOccupations_ctrl2_ddlExposureType'])[1]
    ...    ${exposure3}
    ${exposure_amount3}=    Get From Dictionary    ${current_row}    AK
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewOccupations_ctrl2_tbWageroll'])[1]
    ...    ${exposure_amount3}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    ${risk_management}=    Get From Dictionary    ${current_row}    AL
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlRiskManagementGrade'])[1]
    ...    ${risk_management}
    ${claim_history}=    Get From Dictionary    ${current_row}    AM
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbClaimsHistoryDiscount'])[1]
    ...    ${claim_history}
    ${comments}=    Get From Dictionary    ${current_row}    AN
    Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbDiscountComments'])[1]
    ...    ${comments}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Hist Exposure'])[1]    20s
    ${exposure_type}=    Get From Dictionary    ${current_row}    AO
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_HistoricalExposureForm_ddlExposureType'])[1]
    ...    ${exposure_type}
    ${exposure_amount}=    Get From Dictionary    ${current_row}    AP
    Sleep    5s
    Safe Input Exposure Amount    xpath=(//input[contains(@id, 'txtExposureAmount')])[1]    ${exposure_amount}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Individual Claims'])[1]    20s
    Safe Click Element    (//span[normalize-space()='Yes'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_IndividualClaimsControl_radGridIndividualClaims_ctl00_ctl03_ctl01_btnInsert'])[1]
    ${policy_year}=    Get From Dictionary    ${current_row}    AQ
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_IndividualClaimsControl_radGridIndividualClaims_ctl00_ctl04_txtPolicyYear'])[1]
    ...    ${policy_year}
    ${incident_date}=    Get From Dictionary    ${current_row}    AR
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_IndividualClaimsControl_radGridIndividualClaims_ctl00_ctl04_txtIncidentDate_textDate'])[1]
    ...    ${incident_date}
    ${currency}=    Get From Dictionary    ${current_row}    AS
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_IndividualClaimsControl_radGridIndividualClaims_ctl00_ctl04_ddlOrigCurrency'])[1]
    ...    ${currency}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Override'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    AT
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_OverrideControl_rptLiabilityOverrideOccupation_ctl01_tbRateOverride'])[1]
    ...    ${rate_override}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    AU
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ControlSummary_tbModelPremiumEL'])[1]
    ...    ${rate_override}
    ${rate_override}=    Get From Dictionary    ${current_row}    AV
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ControlSummary_tbModelPremiumPL'])[1]
    ...    ${rate_override}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Insured Details[Copy Risk]
    Wait Until Element Is Visible    (//a[normalize-space()='Insured Details'])[1]    20s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListBranch'])[1]
    ...    Bristol
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]
    ...    BRISTLUK
    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    ${effdate}=    Get From Dictionary    ${current_row}    H
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
    ...    ${effdate}
    ${Indcode}=    Get From Dictionary    ${current_row}    I
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]
    ...    ${Indcode}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
   Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    ${brokfirm}=    Get From Dictionary    ${current_row}    K
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_TextBoxBrokerFirm'])[1]
    ...    ${brokfirm}
    Safe Click Element
    ...    (//input[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_ButtonBrokerFirmSearchStartsWith"])[1]
    Wait Until Element Is Visible
    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]
    ...    30s
    ${table_elements}=    Get WebElements
    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]
    Log    Found ${table_elements.__len__()} elements in the broker info table
    Safe Click Element
    ...    xpath=(//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"])[3]
    Sleep    3s
    ${value1}=    Is Element Visible    (//span[normalize-space()='Yes'])[1]
    IF    ${value1} == True
        Safe Click Element    (//span[normalize-space()='Yes'])[1]
    END
    Execute Javascript    window.scrollTo(0, document.body.scrollHeight)
    Press Keys    None    PAGE_DOWN
    Sleep    2s
    Wait Until Keyword Succeeds
    ...    5x
    ...    5s
    ...    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s
    ${variableExist}=    Run Keyword And Return Status
    ...    Element Should Be Visible
    ...    (//span[normalize-space()='Continue'])[1]
    IF    ${variableExist} == True
        Wait Until Element Is Visible    (//span[normalize-space()='Continue'])[1]    10s
        Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    ${brokcont}=    Get From Dictionary    ${current_row}    L
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]
    ...    ${brokcont}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_ButtonBrokerContactSearchStartsWith'])[1]
    Wait Until Element Is Visible
    ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]
    ...    30s
    ${table_elements}=    Get WebElements
    ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]
    Log    Found ${table_elements.__len__()} elements in the broker contact table
    Safe Click Element
    ...    xpath=(//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"])[2]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]
    ${EditVisible}=    Is Element Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_repeaterBrokerContact_ctl01_buttonEditBroker'])[1]
    IF    ${EditVisible} == True
        Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_repeaterBrokerContact_ctl01_buttonEditBroker'])[1]
        ${PCLicHolder}=    Is Checkbox Selected
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_chkPAndCLicenseHolder'])[1]
        IF    ${PCLicHolder} == False
            Safe Select Checkbox
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_chkPAndCLicenseHolder'])[1]
            Safe Click Element
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]
            Press Keys    None    PAGE_DOWN
            Wait Until Keyword Succeeds
            ...    3x
            ...    5s
            ...    Safe Click Element
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
        END
    ELSE
        Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]
        Press Keys    None    PAGE_DOWN
        Wait Until Keyword Succeeds
        ...    3x
        ...    5s
        ...    Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    END
    Sleep    10s
    ${value2}=    Is Element Visible    (//span[normalize-space()='Continue'])[1]
    IF    ${value2} == True
        Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s

Fill Pricing Details[Copy Risk]
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Policy Info'])[1]    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Exposures'])[1]    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Hist Exposure'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Individual Claims'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Override'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    BU
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_OverrideControl_rptLiabilityOverrideOccupation_ctl01_tbRateOverride'])[1]
    ...    ${rate_override}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Policy Info'])[1]    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Exposures'])[1]    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Hist Exposure'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Individual Claims'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Override'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    BV
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_OverrideControl_rptLiabilityOverrideOccupation_ctl01_tbRateOverride'])[1]
    ...    ${rate_override}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Risk Creation[Renewal]
    Fill Insured Details[Renewal]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Submission
    Fill Pricing Details[Renewal]

Fill Insured Details[Renewal]
    Wait Until Element Is Visible    (//a[normalize-space()='Insured Details'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Store Risk Number    Renewal
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    Press Keys    None    PAGE_DOWN
    Wait Until Keyword Succeeds
    ...    3x
    ...    5s
    ...    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s
    ${variableExist}=    Run Keyword And Return Status
    ...    Element Should Be Visible
    ...    (//span[normalize-space()='Continue'])[1]
    IF    ${variableExist} == True
        Wait Until Element Is Visible    (//span[normalize-space()='Continue'])[1]    10s
        Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Press Keys    None    PAGE_DOWN
    Wait Until Keyword Succeeds
    ...    3x
    ...    5s
    ...    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    10s
    ${value2}=    Is Element Visible    (//span[normalize-space()='Continue'])[1]
    IF    ${value2} == True
        Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END

Fill Pricing Details[Renewal]
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Policy Info'])[1]    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Exposures'])[1]    10s
    ${risk_management}=    Get From Dictionary    ${current_row}    U
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlRiskManagementGrade'])[1]
    ...    ${risk_management}
    ${claim_history}=    Get From Dictionary    ${current_row}    V
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbClaimsHistoryDiscount'])[1]
    ...    ${claim_history}
    ${comments}=    Get From Dictionary    ${current_row}    W
    Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbDiscountComments'])[1]
    ...    ${comments}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Hist Exposure'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Individual Claims'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Override'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    BU
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_OverrideControl_rptLiabilityOverrideOccupation_ctl01_tbRateOverride'])[1]
    ...    ${rate_override}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Policy Info'])[1]    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Exposures'])[1]    10s
    ${risk_management}=    Get From Dictionary    ${current_row}    AL
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlRiskManagementGrade'])[1]
    ...    ${risk_management}
    ${claim_history}=    Get From Dictionary    ${current_row}    AM
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbClaimsHistoryDiscount'])[1]
    ...    ${claim_history}
    ${comments}=    Get From Dictionary    ${current_row}    AN
    Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbDiscountComments'])[1]
    ...    ${comments}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Hist Exposure'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Individual Claims'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Override'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    BV
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_OverrideControl_rptLiabilityOverrideOccupation_ctl01_tbRateOverride'])[1]
    ...    ${rate_override}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rate Monitor EL'])[1]
    ${comments}=    Get From Dictionary    ${current_row}    AN
    Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ExposureRatedMonitorControl_tbDiscountComments'])[1]
    ...    ${comments}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rate Monitor PL'])[1]
    ${comments}=    Get From Dictionary    ${current_row}    AN
    Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ExposureRatedMonitorControl_tbDiscountComments'])[1]
    ...    ${comments}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Pricing Details[Reissue]
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Policy Info'])[1]    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Exposures'])[1]    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Hist Exposure'])[1]    20s
    ${exposure_amount}=    Get From Dictionary    ${current_row}    Y
    Sleep    5s
    Sleep    15s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Individual Claims'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Override'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    BU
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_OverrideControl_rptLiabilityOverrideOccupation_ctl01_tbRateOverride'])[1]
    ...    ${rate_override}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Policy Info'])[1]    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Exposures'])[1]    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Hist Exposure'])[1]    20s
    ${exposure_amount}=    Get From Dictionary    ${current_row}    AP
    Sleep    5s
    Sleep    15s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Individual Claims'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Override'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    BV
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_OverrideControl_rptLiabilityOverrideOccupation_ctl01_tbRateOverride'])[1]
    ...    ${rate_override}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Pre Bind Endorsement
    ${category_list1}=    Get From Dictionary    ${current_row}    AW
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]
    ...    ${category_list1}
    Sleep    2s
    ${endorsement_type_list1}=    Get From Dictionary    ${current_row}    AX
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlManuscriptTemplates'])[1]
    ...    ${endorsement_type_list1}
    Sleep    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btAddEndorsement'])[1]
    Sleep    3s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_EndorsementStandardButtons_btnSubmit'])[1]
    Sleep    3s
    ${category_list2}=    Get From Dictionary    ${current_row}    AY
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]
    ...    ${category_list2}
    Sleep    10s
    ${endorsement_type_list2}=    Get From Dictionary    ${current_row}    AZ
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdEndorsementTypeList'])[1]
    ...    ${endorsement_type_list2}
    Sleep    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btAddEndorsement'])[1]
    Sleep    3s
    ${amount_1}=    Get From Dictionary    ${current_row}    BA
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_GenericEndorsement_ctl_1567'])[1]
    ...    1000
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_EndorsementStandardButtons_btnSubmit'])[1]
    Sleep    3s
    ${category_list3}=    Get From Dictionary    ${current_row}    BB
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]
    ...    ${category_list3}
    Sleep    2s
    ${endorsement_type_list3}=    Get From Dictionary    ${current_row}    BC
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdEndorsementTypeList'])[1]
    ...    ${endorsement_type_list3}
    Sleep    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btAddEndorsement'])[1]
    Sleep    3s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_EndorsementStandardButtons_btnSubmit'])[1]
    Sleep    3s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Remove'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    Sleep    3s
    Check All Endorsements
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Sleep    5s

Quote
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Quote Final'])[1]    20s
    ${cyber_type}=    Get From Dictionary    ${current_row}    BF
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlCyberType'])[1]
    ...    ${cyber_type}
    ${cyber_limit}=    Get From Dictionary    ${current_row}    BG
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbCyberLimit'])[1]
    ...    ${cyber_limit}
    ${comments_under_quote}=    Get From Dictionary    ${current_row}    BH
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbFinalQuoteComments'])[1]
    ...    ${comments_under_quote}
    ${business_description}=    Get From Dictionary    ${current_row}    BH
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_TextBoxBusinessDescription'])[1]
    ...    ${business_description}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_BtnGenQuote'])[1]

Quote[Copy Risk]
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Quote Final'])[1]    20s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_BtnGenQuote'])[1]

Pre Bind Endorsement[Reissue]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Sleep    5s
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Sleep    5s

Quote[Reissue]
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Quote Final'])[1]    20s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_BtnGenQuote'])[1]

Ready To Bind
    Safe Click Element    (//a[normalize-space()='Bind'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_buttonBind'])[1]

Book
    [Arguments]    ${download_policy}
    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Book and Issue'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Multi-Currency Taxes'])[1]
    ...    30s
    ${country}=    Get From Dictionary    ${current_row}    BJ
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_PremiumAndTaxControlEl_ctrlMultiCurrencyBreakdown_rptPremiumBreakdown_ctl00_ddlCountry'])[1]
    ...    ${country}
    ${exchange_rate}=    Get From Dictionary    ${current_row}    BK
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_PremiumAndTaxControlEl_ctrlMultiCurrencyBreakdown_rptPremiumBreakdown_ctl00_txtExchangeRate'])[1]
    ...    ${exchange_rate}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_PremiumAndTaxControlPl_legendPremTax'])[1]
    ${country1}=    Get From Dictionary    ${current_row}    BL
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_PremiumAndTaxControlPl_ctrlMultiCurrencyBreakdown_rptPremiumBreakdown_ctl00_ddlCountry'])[1]
    ...    ${country1}
    ${exchange_rate1}=    Get From Dictionary    ${current_row}    BM
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_PremiumAndTaxControlPl_ctrlMultiCurrencyBreakdown_rptPremiumBreakdown_ctl00_txtExchangeRate'])[1]
    ...    ${exchange_rate1}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[5]
    Safe Click Element    (//span[normalize-space()='Next'])[1]
    ${jurisdiction}=    Get From Dictionary    ${current_row}    BN
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_BookControl_ddlJurisdictionEl'])[1]
    ...    ${jurisdiction}
    ${jurisdiction1}=    Get From Dictionary    ${current_row}    BO
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_BookControl_ddlJurisdictionPl'])[1]
    ...    ${jurisdiction1}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Contract Certainty'])[1]
    ...    30s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[2]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[3]
    ${source_of_business}=    Get From Dictionary    ${current_row}    BP
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_ddlSourceofBusiness'])[1]
    ...    ${source_of_business}
    ${placement_type}=    Get From Dictionary    ${current_row}    BQ
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_ddlPlacementType'])[1]
    ...    ${placement_type}
    ${ncb_applicable}=    Get From Dictionary    ${current_row}    BR
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_ddlNCBApplicable'])[1]
    ...    ${ncb_applicable}
    ${risk_level}=    Get From Dictionary    ${current_row}    BS
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_ddlRiskLevel'])[1]
    ...    ${risk_level}
    ${eea_exposure}=    Get From Dictionary    ${current_row}    BT
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_txtEEAExposure'])[1]
    ...    ${eea_exposure}
    Sleep    60s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonBook'])[1]

Book[Reissue]
    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Book and Issue'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_buttonBook'])[1]

Copy Quote
    Wait Until Element Is Visible    (//a[normalize-space()='Copy Quote'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Copy Quote'])[1]
    Wait Until Element Is Visible
    ...    xpath=(//span[contains(@id, 'HeaderQuote')])[2]
    ...    10s
    Verify Status
    ...    xpath=(//span[contains(@id, 'HeaderQuote')])[2]
    ...    Option 2 - [Status : Quoted (In Revision)]
    Safe Click Element
    ...    xpath=(//span[contains(@id, 'HeaderQuote')])[1]

View Risk
    Wait Until Element Is Visible    (//a[normalize-space()='View Risk'])[1]    30s
    Safe Click Element    (//a[normalize-space()='View Risk'])[1]
    Sleep    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonViewCurrent'])[1]
    Verify Status    (//legend[@id='fsMainContentLegend'])[1]    View Risk

Post Bind Endorsement
    Wait Until Element Is Visible    (//a[normalize-space()='Endorsement'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Endorsement'])[1]
    ${category_list}=    Get From Dictionary    ${current_row}    BW
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]
    ...    ${category_list}
    Sleep    2s
    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    BX
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdEndorsementTypeList'])[1]
    ...    ${endorsement_type_list}
    Sleep    2s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_chknewEL'])[1]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_chknewPL'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btEditEndorsement'])[1]
    Sleep    3s
    ${effective_date_post_bind}=    Get From Dictionary    ${current_row}    BY
    ${effective_date_post_bind}=    Format Date For Input    ${effective_date_post_bind}
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementEffectiveDate_datepickerEffectiveDate_textDate'])[1]
    ...    ${effective_date_post_bind}
    ${policy_expiry_date}=    Get From Dictionary    ${current_row}    BZ
    ${policy_expiry_date}=    Format Date For Input    ${policy_expiry_date}
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_datepickerPolicyExpiryDate_textDate'])[1]
    ...    ${policy_expiry_date}
    Sleep    3s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_EndorsementStandardButtons_btnSubmit'])[1]
    Safe Click Element    xpath=//div[@class='RiskHeaderColumnOne']//div[1]
    Sleep    30s

Renewal
    Wait Until Element Is Visible    (//a[normalize-space()='Renew'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Renew'])[1]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Submission
    Risk Creation[Renewal]
    Pre Bind Endorsement[Reissue]
    Quote[Reissue]
    Ready To Bind
    Book    False

Reissue
    Wait Until Element Is Visible    (//a[normalize-space()='Reissue'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Reissue'])[1]
    Sleep    10s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[7]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Quoted (Pending, In Revision, RI)
    Wait Until Element Is Visible    (//a[normalize-space()='Edit Quote'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Edit Quote'])[1]
    Fill Pricing Details[Reissue]
    Pre Bind Endorsement[Reissue]
    Quote[Reissue]
    Ready To Bind
    Book    False

Copy Risk
    Wait Until Element Is Visible    (//a[normalize-space()='Copy Risk'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Copy Risk'])[1]
    Sleep    10s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[4]
    Fill Insured Details[Copy Risk]
    Fill Pricing Details[Copy Risk]
    Pre Bind Endorsement
    Quote[Copy Risk]
    Ready To Bind
    Book    False

Continue on Next Risk
    Wait Until Element Is Visible    xpath=//a[normalize-space()='Create New Risk >']    20s
    Safe Click Element    xpath=//a[normalize-space()='Create New Risk >']
    Safe Click Element    xpath=//a[normalize-space()='Create New Risk >']
    Wait Until Element Is Visible    xpath=//a[normalize-space()='US Primary GL']    30s
    Safe Click Element    xpath=//a[normalize-space()='US Primary GL']
    Safe Click Element
    ...    id:ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlClearanceSearch_CreateInsured
    Wait Until Element Is Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]    40s

Verify Status
    [Arguments]    ${location}    ${status_expected}
    Wait Until Keyword Succeeds    45x    2s    Element Should Contain    ${location}    ${status_expected}

Safe Click Element
    [Arguments]    ${locator}
    Wait Until Keyword Succeeds    15x    2s    Click Element When Visible    ${locator}

Safe Input Text
    [Arguments]    ${locator}    ${text}
    Wait Until Keyword Succeeds    5x    2s    Input Text When Element Is Visible    ${locator}    ${text}

Safe Select From List By Label
    [Arguments]    ${locator}    ${label}
    Wait Until Keyword Succeeds    5x    2s    Select From List By Label    ${locator}    ${label}

Safe Select From List By Index
    [Arguments]    ${locator}    ${index}
    Wait Until Keyword Succeeds    5x    2s    Select From List By Index    ${locator}    ${index}

Safe Select Checkbox
    [Arguments]    ${locator}
    Wait Until Keyword Succeeds    5x    2s    Select Checkbox    ${locator}

Format Date For Input
    [Arguments]    ${raw_date}
    ${status}    ${result}=    Run Keyword And Ignore Error    Convert Date    ${raw_date}    result_format=%d/%m/%Y
    ${final_date}=    Set Variable If    '${status}' == 'PASS'    ${result}    ${raw_date}
    [Return]    ${final_date}

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

Check All Endorsements
    FOR    ${index}    IN RANGE    1    20
        ${idx}=    Convert To String    ${index}
        ${idx}=    Run Keyword If    ${index} < 10    Set Variable    0${index}    ELSE    Set Variable    ${index}
        
        ${el_xpath}=    Set Variable    //input[contains(@id, 'repeaterEndorsement_ctl${idx}_chkEL')]
        ${pl_xpath}=    Set Variable    //input[contains(@id, 'repeaterEndorsement_ctl${idx}_chkPL')]
        
        ${el_exists}=    Run Keyword And Return Status    Page Should Contain Element    ${el_xpath}
        IF    ${el_exists} == True
            ${el_checked}=    Execute Javascript    return document.evaluate("${el_xpath}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.checked;
            IF    ${el_checked} == False
                Execute Javascript    document.evaluate("${el_xpath}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();
                Sleep    1s
            END
        ELSE
            BREAK
        END
        
        ${pl_exists}=    Run Keyword And Return Status    Page Should Contain Element    ${pl_xpath}
        IF    ${pl_exists} == True
            ${pl_checked}=    Execute Javascript    return document.evaluate("${pl_xpath}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.checked;
            IF    ${pl_checked} == False
                Execute Javascript    document.evaluate("${pl_xpath}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();
                Sleep    1s
            END
        END
    END

Safe Input Exposure Amount
    [Arguments]    ${locator}    ${expected_amount}
    Wait Until Keyword Succeeds    15x    3s    Attempt Input Exposure Amount    ${locator}    ${expected_amount}

Attempt Input Exposure Amount
    [Arguments]    ${locator}    ${expected_amount}
    Wait Until Element Is Visible    ${locator}    15s
    Run Keyword And Ignore Error    Wait Until Page Does Not Contain Element    css=.ui-widget-overlay    15s
    Click Element When Visible    ${locator}
    Sleep    1s
    RPA.Browser.Selenium.Clear Element Text    ${locator}
    Sleep    1s
    RPA.Browser.Selenium.Press Keys    ${locator}    ${expected_amount}
    RPA.Browser.Selenium.Press Keys    ${locator}    TAB
    Sleep    2s
    Run Keyword And Ignore Error    Wait Until Page Does Not Contain Element    css=.ui-widget-overlay    15s
    ${val}=    RPA.Browser.Selenium.Get Value    ${locator}
    ${clean_val}=    Evaluate    str('${val}').replace(',', '').replace('.00', '').strip()
    ${clean_expected}=    Evaluate    str('${expected_amount}').replace(',', '').replace('.00', '').strip()
    Should Be Equal As Strings    ${clean_val}    ${clean_expected}

Store Risk Number
    [Arguments]    ${flow_type}
    Wait Until Element Is Visible    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]    15s
    ${risk_number}=    RPA.Browser.Selenium.Get Text    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]
    Evaluate    open(r'${CURDIR}${/}risk_numbers.txt', 'a').write('${flow_type} - Risk Number: ${risk_number}\\n')
