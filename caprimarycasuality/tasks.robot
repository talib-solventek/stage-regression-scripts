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
Canada Primary Casualty
    Open Workbook    ${DATA_SOURCE}
    ${excel_rows}=    Read Worksheet    CAPrimaryCasuality
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
    Safe Click Element    xpath=//a[normalize-space()='Create New Risk >']
    Safe Click Element    xpath=//a[normalize-space()='CA Primary Casualty']
    Sleep    10s
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
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Sleep    1s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[4]
    ${iname}=    Get From Dictionary    ${current_row}    D
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_TextBoxFirstNamedInsured1'])[1]
    ...    ${iname}
    ${province}=    Get From Dictionary    ${current_row}    E
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_ddCanProvince'])[1]
    ...    ${province}
    ${adl1}=    Get From Dictionary    ${current_row}    F
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxCanAddressLine1'])[1]
    ...    ${adl1}
    ${city1}=    Get From Dictionary    ${current_row}    G
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxCanCity'])[1]
    ...    ${city1}
    ${zip1}=    Get From Dictionary    ${current_row}    H
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxCanPostalCode'])[1]
    ${zipelm}=    Get WebElement
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxCanPostalCode'])[1]
    ${attribute}=    Get Element Attribute
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxCanPostalCode'])[1]
    ...    value
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxCanPostalCode'])[1]
    ...    ${zip1}
    Press Keys
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxCanPostalCode'])[1]
    ...    T+5+T+ +5+T+5
    ${return_value}=    Execute Javascript    arguments[0].value='T5T 5T5'    ARGUMENTS    ${zipelm}
    ${attribute}=    Get Element Attribute
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxCanPostalCode'])[1]
    ...    value
    Sleep    1s
    Sleep    10s
    ${pobox}=    Get From Dictionary    ${current_row}    I
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxCanPoBox'])[1]
    ...    ${pobox}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListBranch'])[1]
    ...    Calgary
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]
    ...    CA-CBC
    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    ${effdate}=    Get From Dictionary    ${current_row}    J
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
    ...    ${effdate}
    ${Indcode}=    Get From Dictionary    ${current_row}    K
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]
    ...    ${Indcode}
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_txtNAICSCode'])[1]
    ...    111110 SOYBEAN FARMING
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    ${brokfirm}=    Get From Dictionary    ${current_row}    M
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_TextBoxBrokerFirm'])[1]
    ...    ${brokfirm}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_ButtonBrokerFirmSearchStartsWith'])[1]
    Wait Until Element Is Visible
    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]
    ...    30s
    ${table_elements}=    Get WebElements
    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]
    Log    Found ${table_elements.__len__()} elements in the broker info table
    Safe Click Element    xpath=(//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"])[1]
    Sleep    2s
    ${value1}=    Run Keyword And Return Status
    ...    Wait Until Keyword Succeeds    3x    2s    Element Should Be Visible    (//span[normalize-space()='Yes'])[1]
    IF    ${value1} == True
        Wait Until Keyword Succeeds    3x    2s    Safe Click Element    (//span[normalize-space()='Yes'])[1]
    END
    Sleep    2s
    Press Keys    None    PAGE_DOWN
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
        Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    ${brokcont}=    Get From Dictionary    ${current_row}    N
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
    Safe Click Element    xpath=(//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]
    ${EditVisible}=    Is Element Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_repeaterBrokerContact_ctl01_buttonEditBroker'])[1]
    IF    ${EditVisible} == True
        Safe Click Element
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_repeaterBrokerContact_ctl01_buttonEditBroker'])[1]
        ${PCLicHolder}=    Is Checkbox Selected
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_chkPAndCLicenseHolder'])[1]
        IF    ${PCLicHolder} == False
            Select Checkbox
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_chkPAndCLicenseHolder'])[1]
            Safe Click Element
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]
            Press Keys    None    PAGE_DOWN
            Wait Until Keyword Succeeds
            ...    3x
            ...    5s
            ...    Click Element When Visible
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
        END
    ELSE
        Safe Click Element
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
        Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END

Fill Pricing Details[First Run]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info'])[1]    20s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Sleep    2s
    ${risk_exposure}=    Get From Dictionary    ${current_row}    O
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlRiskExposure'])[1]
    ...    ${risk_exposure}
    ${line_of_business}=    Get From Dictionary    ${current_row}    P
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlLineOfBusiness'])[1]
    ...    ${line_of_business}
    Sleep    5s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element    (//span[normalize-space()='Non Project'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Deductible'])[1]
    Sleep    2s
    ${coverage_form}=    Get From Dictionary    ${current_row}    Q
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlCoverageForm'])[1]
    ...    ${coverage_form}
    ${type_of_business}=    Get From Dictionary    ${current_row}    R
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlTypeOfRisk'])[1]
    ...    ${type_of_business}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Options'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_option1_ctrlSirDeductible_listViewSirDeductibleFrench_ctrl0_btnDelete'])[1]
    Sleep    5s
    ${deductible_type}=    Get From Dictionary    ${current_row}    S
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_option1_ctrlSirDeductible_listViewSirDeductibleFrench_ctrl0_ddlSirDeductibleType'])[1]
    ...    ${deductible_type}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_option1_ctrlSirDeductible_listViewSirDeductibleFrench_ctrl0_cbRatingIdentifierSirDeductible'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Proj Expo'])[1]    10s
    ${class_code}=    Get From Dictionary    ${current_row}    T
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_TextBoxClassCode'])[1]
    ...    ${class_code}
    ${class_code_full}=    Get From Dictionary    ${current_row}    U
    Safe Click Element    (//a[normalize-space()='${class_code_full}'])[1]
    ${loss_cost}=    Get From Dictionary    ${current_row}    V
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_txtLossCostPremOps'])[1]
    ...    ${loss_cost}
    ${ILF}=    Get From Dictionary    ${current_row}    W
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_ddlPremOpsILFTable'])[1]
    ...    ${ILF}
    ${exposures}=    Get From Dictionary    ${current_row}    X
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_txtExposures'])[1]
    ...    ${exposures}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Loss Expo History'])[1]    20s
    ${projected_exposure_amount}=    Get From Dictionary    ${current_row}    Y
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_tbProjectedExposureAmount'])[1]
    ...    ${projected_exposure_amount}
    ${rate_per}=    Get From Dictionary    ${current_row}    Z
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlRatePer'])[1]
    ...    ${rate_per}
    ${exposure_base}=    Get From Dictionary    ${current_row}    AA
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlExposureBase'])[1]
    ...    ${exposure_base}
    ${exposure_amount}=    Get From Dictionary    ${current_row}    AB
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_TextBoxExposureAmount'])[1]
    ...    ${exposure_amount}
    ${aggregate_loss}=    Get From Dictionary    ${current_row}    AC
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_TextBoxGroundUpAggregateLoss'])[1]
    ...    ${aggregate_loss}
    ${aggregate_alae}=    Get From Dictionary    ${current_row}    AD
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_TextBoxGroundUpAggregateAlae'])[1]
    ...    ${aggregate_alae}
    ${number_of_claims}=    Get From Dictionary    ${current_row}    AE
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_TextBoxNumberOfClaims'])[1]
    ...    ${number_of_claims}
    ${valuation_date}=    Get From Dictionary    ${current_row}    AF
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_TextBoxValuationDate_textDate'])[1]
    ...    ${valuation_date}
    ${coverage_form}=    Get From Dictionary    ${current_row}    AG
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_DropDownListCoverageForm'])[1]
    ...    ${coverage_form}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Man Rating'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Large Loss Info'])[1]    5s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Loss Rating'])[1]    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Exp Rating'])[1]    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]    5s
    ${quoted_premium}=    Get From Dictionary    ${current_row}    AH
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbQuotedPremiumValue'])[1]
    ...    ${quoted_premium}
    ${justification}=    Get From Dictionary    ${current_row}    AI
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_textboxJustification'])[1]
    ...    ${justification}
    ${exposure_premium}=    Get From Dictionary    ${current_row}    AJ
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_TextBoxExposurePremium'])[1]
    ...    ${exposure_premium}
    ${rate_per}=    Get From Dictionary    ${current_row}    AK
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListRatePer'])[1]
    ...    ${rate_per}
    ${exposure_base}=    Get From Dictionary    ${current_row}    AL
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListExposureBase'])[1]
    ...    ${exposure_base}
    ${comments_english}=    Get From Dictionary    ${current_row}    AM
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_TextBoxExposureComments'])[1]
    ...    ${comments_english}
    ${comments_french}=    Get From Dictionary    ${current_row}    AN
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_TextBoxExposureCommentsFrench'])[1]
    ...    ${comments_french}
    ${exposure_premium}=    Get From Dictionary    ${current_row}    AO
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListExposureType'])[1]
    ...    ${exposure_premium}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]

Risk Creation[Renewal]
    Fill Insured Details[Renewal]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Submission
    Fill Pricing Details[Renewal]

Fill Insured Details[Renewal]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    3s
    Wait Until Element Is Visible    (//a[normalize-space()='Submission Details'])[1]    20s
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_txtNAICSCode'])[1]
    ...    111110 SOYBEAN FARMING
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    3s
    Wait Until Element Is Visible    (//a[normalize-space()='Broker Firm'])[1]    20s
    Press Keys    None    PAGE_DOWN
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s
    ${variableExist}=    Run Keyword And Return Status
    ...    Element Should Be Visible
    ...    (//span[normalize-space()='Continue'])[1]
    IF    ${variableExist} == True
        Wait Until Element Is Visible    (//span[normalize-space()='Continue'])[1]    10s
        Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END
    Wait Until Element Is Visible    (//a[normalize-space()='Broker Contacts'])[1]    20s
    Press Keys    None    PAGE_DOWN
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    10s
    ${value2}=    Is Element Visible    (//span[normalize-space()='Continue'])[1]
    IF    ${value2} == True
        Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END

Fill Pricing Details[Renewal]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Options'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Proj Expo'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Loss Expo History'])[1]    20s
    ${projected_exposure_amount}=    Get From Dictionary    ${current_row}    Y
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_tbProjectedExposureAmount'])[1]
    ...    ${projected_exposure_amount}
    ${rate_per}=    Get From Dictionary    ${current_row}    Z
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlRatePer'])[1]
    ...    ${rate_per}
    ${exposure_base}=    Get From Dictionary    ${current_row}    AA
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlExposureBase'])[1]
    ...    ${exposure_base}
    ${exposure_amount}=    Get From Dictionary    ${current_row}    AB
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_TextBoxExposureAmount'])[1]
    ...    ${exposure_amount}
    ${aggregate_loss}=    Get From Dictionary    ${current_row}    AC
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_TextBoxGroundUpAggregateLoss'])[1]
    ...    ${aggregate_loss}
    ${aggregate_alae}=    Get From Dictionary    ${current_row}    AD
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_TextBoxGroundUpAggregateAlae'])[1]
    ...    ${aggregate_alae}
    ${number_of_claims}=    Get From Dictionary    ${current_row}    AE
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_TextBoxNumberOfClaims'])[1]
    ...    ${number_of_claims}
    ${valuation_date}=    Get From Dictionary    ${current_row}    AF
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_TextBoxValuationDate_textDate'])[1]
    ...    ${valuation_date}
    ${coverage_form}=    Get From Dictionary    ${current_row}    AG
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_DropDownListCoverageForm'])[1]
    ...    ${coverage_form}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Man Rating'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Large Loss Info'])[1]    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Loss Rating'])[1]    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Exp Rating'])[1]    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]    5s
    ${quoted_premium}=    Get From Dictionary    ${current_row}    AH
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbQuotedPremiumValue'])[1]
    ...    ${quoted_premium}
    ${justification}=    Get From Dictionary    ${current_row}    AI
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_textboxJustification'])[1]
    ...    ${justification}
    ${exposure_premium}=    Get From Dictionary    ${current_row}    AJ
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_TextBoxExposurePremium'])[1]
    ...    ${exposure_premium}
    ${rate_per}=    Get From Dictionary    ${current_row}    AK
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListRatePer'])[1]
    ...    ${rate_per}
    ${exposure_base}=    Get From Dictionary    ${current_row}    AL
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListExposureBase'])[1]
    ...    ${exposure_base}
    ${comments_english}=    Get From Dictionary    ${current_row}    AM
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_TextBoxExposureComments'])[1]
    ...    ${comments_english}
    ${comments_french}=    Get From Dictionary    ${current_row}    AN
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_TextBoxExposureCommentsFrench'])[1]
    ...    ${comments_french}
    ${exposure_premium}=    Get From Dictionary    ${current_row}    AO
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListExposureType'])[1]
    ...    ${exposure_premium}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Sleep    5s

Fill Insured Details[Copy Risk]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Sleep    1s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[4]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListBranch'])[1]
    ...    Calgary
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]
    ...    CA-CBC
    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    ${effdate}=    Get From Dictionary    ${current_row}    J
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
    ...    ${effdate}
    ${Indcode}=    Get From Dictionary    ${current_row}    K
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]
    ...    ${Indcode}
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_txtNAICSCode'])[1]
    ...    111110 SOYBEAN FARMING
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    Press Keys    None    PAGE_DOWN
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
        Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Press Keys    None    PAGE_DOWN
    Wait Until Keyword Succeeds
    ...    3x
    ...    5s
    ...    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    10s
    ${value2}=    Is Element Visible    (//span[normalize-space()='Continue'])[1]
    IF    ${value2} == True
        Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END

Fill Pricing Details[Copy Risk]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Options'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Proj Expo'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Loss Expo History'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Man Rating'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Large Loss Info'])[1]    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Loss Rating'])[1]    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Exp Rating'])[1]    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Sleep    5s

Fill Pricing Details[Reissue]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Options'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Proj Expo'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Loss Expo History'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Man Rating'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Large Loss Info'])[1]    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Loss Rating'])[1]    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Exp Rating'])[1]    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Sleep    5s

Pre Bind Endorsement
    Wait Until Element Is Visible    (//a[normalize-space()='Canada Pre Bind Endorsements'])[1]    20s
    Sleep    2s
    FOR    ${index}    IN RANGE    1    5
        Safe Click Element    (//div[normalize-space()='Sample'])[1]
        Safe Click Element    (//div[normalize-space()='Sample'])[1]
        ${sample_yes}=    Is Element Visible    (//span[@class='ui-button-text'][normalize-space()='Edit'])[1]
        IF    ${sample_yes} == True
            Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Remove'])[1]
            Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
        END
    END
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    Check All Subjectivities
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s

Quote
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Quote Final'])[1]    20s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    ${comments_under_quote}=    Get From Dictionary    ${current_row}    AR
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxFinalQuoteComments'])[1]
    ...    ${comments_under_quote}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonCheckQuote'])[1]

Pre Bind Endorsement[Reissue]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s

Ready To Bind
    Wait Until Element Is Visible    (//a[normalize-space()='Bind'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Bind'])[1]
    Sleep    2s
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Safe Click Element    (//input[contains(@id, 'BtnNext')] | //span[normalize-space()='Next'])[1]
    Sleep    2s
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonCheckBind'])[1]
    Sleep    20s

Book
    [Arguments]    ${download_policy}
    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Book and Issue'])[1]
    Sleep    10s
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    ${insured_type}=    Get From Dictionary    ${current_row}    AS
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownInsuredType'])[1]
    ...    ${insured_type}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    ${Jurisdiction}=    Get From Dictionary    ${current_row}    AT
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownJurisdiction'])[1]
    ...    ${Jurisdiction}
    ${market_hazard}=    Get From Dictionary    ${current_row}    AU
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownMarketHazard'])[1]
    ...    ${market_hazard}
    ${pollution_code}=    Get From Dictionary    ${current_row}    AV
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownPollutionCode'])[1]
    ...    ${pollution_code}
    ${business_description}=    Get From Dictionary    ${current_row}    AW
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxBusinessDescription'])[1]
    ...    ${business_description}
    ${coverage_territory}=    Get From Dictionary    ${current_row}    AX
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownCoverageTerritory'])[1]
    ...    ${coverage_territory}
    ${financial_rating}=    Get From Dictionary    ${current_row}    AY
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownFinancialRating'])[1]
    ...    ${financial_rating}
    ${defense_code}=    Get From Dictionary    ${current_row}    AZ
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownDefenseCode'])[1]
    ...    ${defense_code}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[4]
    Safe Click Element    (//span[normalize-space()='Français'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonCheckBook'])[1]

Risk UW eFile
    Wait Until Element Is Visible    (//a[normalize-space()='Risk UW eFile'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Risk UW eFile'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s
    Sleep    30s
    Wait Until Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnReload'])[1]
    ...    20s
    Safe Click Element If Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnReload'])[1]
    Sleep    30s
    Wait Until Element Is Visible    (//td[@role='gridcell'])[5]//a    120s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnReload'])[1]
    IF    ${download_policy} == True
        Wait Until Element Is Visible    (//a[normalize-space()='Policy.pdf'])[1]    60s
        Safe Click Element    (//a[normalize-space()='Policy.pdf'])[1]
        Safe Click Element If Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    ELSE
        Sleep    240s
    END
    Safe Click Element    xpath=//div[@class='RiskHeaderColumnOne']//div[1]

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
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='OK'])[1]
    ${category_list}=    Get From Dictionary    ${current_row}    BA
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]
    ...    ${category_list}
    Sleep    10s
    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    BB
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdEndorsementTypeList'])[1]
    ...    ${endorsement_type_list}
    Sleep    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btAddEndorsement'])[1]
    Sleep    3s
    ${effective_date_post_bind}=    Get From Dictionary    ${current_row}    BC
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementEffectiveDate_datepickerEffectiveDate_textDate'])[1]
    ...    ${effective_date_post_bind}
    ${policy_expiry_date}=    Get From Dictionary    ${current_row}    BD
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_datepickerPolicyExpiryDate_textDate'])[1]
    ...    ${policy_expiry_date}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_EndorsementStandardButtons_btnSubmit'])[1]
    Safe Click Element    xpath=//div[@class='RiskHeaderColumnOne']//div[1]

Renewal
    Wait Until Element Is Visible    (//a[normalize-space()='Renew'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Renew'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='OK'])[1]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Submission
    Risk Creation[Renewal]
    Pre Bind Endorsement[Reissue]
    Quote
    Ready To Bind
    Book    False

Reissue
    Wait Until Element Is Visible    (//a[normalize-space()='Reissue'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Reissue'])[1]
    Sleep    10s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='OK'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[7]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Quoted (Pending, In Revision, RI)
    Wait Until Element Is Visible    (//a[normalize-space()='Edit Quote'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Edit Quote'])[1]
    Fill Pricing Details[Reissue]
    Pre Bind Endorsement[Reissue]
    Quote
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
    Quote
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
    Wait Until Keyword Succeeds    10x    3s    Verify Status Matches    ${location}    ${status_expected}

Verify Status Matches
    [Arguments]    ${location}    ${status_expected}
    ${status_actual}=    RPA.Browser.Selenium.Get Text    ${location}
    ${match}=    Evaluate    '${status_expected}' == '${status_actual}' or ('${status_expected}' == 'Submission (Pricing In Progress)' and '${status_actual}' == 'Submission (Pending)')
    Should Be True    ${match}    Expected status '${status_expected}' but got '${status_actual}'


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
