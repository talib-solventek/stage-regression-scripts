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
Primary GL
    Open Workbook    ${DATA_SOURCE}
    ${excel_rows}=    Read Worksheet    GL
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
    Open Browser    ${URL}    edge    executable_path=${DRIVER}    options=add_argument("--inprivate")
    Maximize Browser Window
    Set Selenium Implicit Wait    30s
    Safe Click Element    xpath=//a[normalize-space()='Create New Risk >']
    Safe Click Element    xpath=//a[normalize-space()='US Primary GL']
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
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    ${iname}=    Get From Dictionary    ${current_row}    D
    Safe Input Text
    ...    //input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_TextBoxFirstNamedInsured1']
    ...    ${iname}
    ${state}=    Get From Dictionary    ${current_row}    E
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_ddUSState'])[1]
    ...    ${state}
    ${adl1}=    Get From Dictionary    ${current_row}    F
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSAddressLine1'])[1]
    ...    ${adl1}
    ${city1}=    Get From Dictionary    ${current_row}    G
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_ddUSCity'])[1]
    ...    ${city1}
    ${zip1}=    Get From Dictionary    ${current_row}    H
    Safe Click Element
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
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListBranch'])[1]
    ...    Los Angeles
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]
    ...    LSI2
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
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    ${brokfirm}=    Get From Dictionary    ${current_row}    M
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_TextBoxBrokerFirm'])[1]
    ...    ${brokfirm}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_ButtonBrokerFirmSearchStartsWith'])[1]
    Sleep    10s
    ${table_elements}=    Get WebElements
    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]
    Log    Found ${table_elements.__len__()} elements in the broker info table
    Safe Click Element
    ...    xpath=(//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"])[1]
    Sleep    2s
    ${value1}=    Run Keyword And Return Status
    ...    Wait Until Keyword Succeeds    3x    2s    Element Should Be Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    IF    ${value1} == True
        Wait Until Keyword Succeeds    3x    2s    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    END
    Sleep    2s
    Scroll Element Into View
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
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
    ${brokcont}=    Get From Dictionary    ${current_row}    N
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]
    ...    ${brokcont}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_ButtonBrokerContactSearchStartsWith'])[1]
    Wait Until Element Is Visible
    ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]
    ...    30s
    Safe Click Element
    ...    xpath=(//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]
    Sleep    2s
    Press Keys    None    PAGE_DOWN
    ${yes_visible}=    Run Keyword And Return Status
    ...    Wait Until Keyword Succeeds    3x    2s    Element Should Be Visible
    ...    xpath=//*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContactSurplus_rblSurplusLinesBroker"]/label[1]/span
    IF    ${yes_visible} == True
        Wait Until Keyword Succeeds    3x    2s    Safe Click Element
        ...    xpath=//*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContactSurplus_rblSurplusLinesBroker"]/label[1]/span
    END
    Wait Until Keyword Succeeds
    ...    3x
    ...    5s
    ...    Safe Click Element
    ...    xpath=//*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext"]
    Sleep    2s
    ${continue_visible}=    Is Element Visible    (//span[normalize-space()='Continue'])[1]
    IF    ${continue_visible} == True
        Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END

Fill Pricing Details[First Run]
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Sleep    2s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Non Project'])[1]
    ${line_of_business}=    Get From Dictionary    ${current_row}    O
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlLineOfBusiness'])[1]
    ...    ${line_of_business}
    Sleep    2s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Deductible'])[1]
    Sleep    2s
    ${coverage_form}=    Get From Dictionary    ${current_row}    P
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlCoverageForm'])[1]
    ...    ${coverage_form}
    ${type_of_business}=    Get From Dictionary    ${current_row}    Q
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlTypeOfBusiness'])[1]
    ...    ${type_of_business}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[5]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//span[normalize-space()='Option 1'])[1]    10s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[4]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[5]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[7]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Proj Expo'])[1]    10s
    ${zip}=    Get From Dictionary    ${current_row}    R
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_txtZip'])[1]
    ...    ${zip}
    ${class_code}=    Get From Dictionary    ${current_row}    S
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_TextBoxClassCode'])[1]
    ...    ${class_code}
    Safe Click Element
    ...    (//a[normalize-space()='11201 - Contractors Equipment--Cranes, Derricks, Power Shovels and Equipment Incidental Thereto--Rented to Others with Operators'])[1]
    ${exposures}=    Get From Dictionary    ${current_row}    U
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_txtExposures'])[1]
    ...    ${exposures}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonCalculate'])[1]
    Sleep    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Loss Expo History'])[1]    20s
    Sleep    5s
    ${projected_exposure_amount}=    Get From Dictionary    ${current_row}    V
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_tbProjectedExposureAmount'])[1]
    ...    3000000
    Sleep    5s
    ${rate_per}=    Get From Dictionary    ${current_row}    W
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlRatePer'])[1]
    ...    ${rate_per}
    ${exposure_base}=    Get From Dictionary    ${current_row}    X
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlExposureBase'])[1]
    ...    ${exposure_base}
    ${exposure_amount}=    Get From Dictionary    ${current_row}    Y
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_TextBoxExposureAmount'])[1]
    ...    ${exposure_amount}
    ${aggregate_loss}=    Get From Dictionary    ${current_row}    Z
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_TextBoxGroundUpAggregateLoss'])[1]
    ...    ${aggregate_loss}
    ${aggregate_alae}=    Get From Dictionary    ${current_row}    AA
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_TextBoxGroundUpAggregateAlae'])[1]
    ...    ${aggregate_alae}
    ${number_of_claims}=    Get From Dictionary    ${current_row}    AB
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_TextBoxNumberOfClaims'])[1]
    ...    ${number_of_claims}
    ${valuation_date}=    Get From Dictionary    ${current_row}    AC
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_TextBoxValuationDate_textDate'])[1]
    ...    ${valuation_date}
    ${large_losses}=    Get From Dictionary    ${current_row}    AC
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlLargeLosses_txtLossCap'])[1]
    ...    1000
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Man Rating'])[1]    10s
    ${NAICS_Code}=    Get From Dictionary    ${current_row}    AD
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlManualRatingInfo1_txtNAICSCode'])[1]
    ...    ${NAICS_Code}
    Safe Click Element    (//a[normalize-space()='113310 LOGGING'])[1]
    Safe Click Element    (//span[normalize-space()='Correct'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonCalculate'])[1]
    Sleep    5s
    ${variableExist}=    Is Element Visible    (//span[normalize-space()='Ok'])[1]
    IF    ${variableExist} == True
        Safe Click Element    (//span[normalize-space()='Ok'])[1]
    END
    Sleep    5s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Log To Console    \n[Summary of Pricing] Starting...
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]    5s
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//legend[normalize-space()='Summary of Pricing'])[1]    30s
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbTotalPolicyPremium'])[1]    30s
    Sleep    2s
    Log To Console    [Summary of Pricing] Entering Total Policy Premium...
    ${total_policy_premium}=    Get From Dictionary    ${current_row}    AG
    Wait Until Keyword Succeeds    3x    5s    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbTotalPolicyPremium'])[1]
    ...    100000
    Sleep    2s
    Log To Console    [Summary of Pricing] Entering Exposure Premium...
    ${exposure_premium}=    Get From Dictionary    ${current_row}    AH
    Wait Until Keyword Succeeds    3x    5s    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_TextBoxExposurePremium'])[1]
    ...    ${exposure_premium}
    Sleep    2s
    Log To Console    [Summary of Pricing] Selecting Rate Per...
    ${rate_per}=    Get From Dictionary    ${current_row}    AI
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListRatePer'])[1]
    ...    ${rate_per}
    Sleep    2s
    ${exposure_base}=    Get From Dictionary    ${current_row}    AJ
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListExposureBase'])[1]
    ...    ${exposure_base}
    Sleep    2s
    ${exposure_premium}=    Get From Dictionary    ${current_row}    AK
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListExposureType'])[1]
    ...    ${exposure_premium}
    Sleep    2s
    ${premium_percentage}=    Get From Dictionary    ${current_row}    AL
    Wait Until Keyword Succeeds    3x    5s    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbMinAuditPremiumPercentage'])[1]
    ...    ${premium_percentage}
    Sleep    2s
    ${tria_premium}=    Get From Dictionary    ${current_row}    AM
    Wait Until Keyword Succeeds    3x    5s    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbTriaPremium'])[1]
    ...    ${tria_premium}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Sleep    5s

Risk Creation[Renewal]
    Fill Insured Details[Renewal]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Submission
    Fill Pricing Details[Renewal]

Fill Insured Details[Renewal]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Submission Details'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Broker Firm'])[1]    20s
    Sleep    2s
    Scroll Element Into View
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Keyword Succeeds
    ...    3x
    ...    5s
    ...    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    2s
    ${variableExist}=    Is Element Visible    (//span[normalize-space()='Continue'])[1]
    IF    ${variableExist} == True
        Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END
    Wait Until Element Is Visible    (//a[normalize-space()='Broker Contacts'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    2s
    ${continue_visible}=    Is Element Visible    (//span[normalize-space()='Continue'])[1]
    IF    ${continue_visible} == True
        Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END

Fill Insured Details[Copy Risk]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Sleep    1s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Submission Details'])[1]    20s
    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    ${effdate}=    Get From Dictionary    ${current_row}    J
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
    ...    ${effdate}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Broker Firm'])[1]    20s
    Sleep    2s
    Scroll Element Into View
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Keyword Succeeds
    ...    3x
    ...    5s
    ...    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    2s
    ${variableExist}=    Is Element Visible    (//span[normalize-space()='Continue'])[1]
    IF    ${variableExist} == True
        Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END
    Wait Until Element Is Visible    (//a[normalize-space()='Broker Contacts'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    2s
    ${continue_visible}=    Is Element Visible    (//span[normalize-space()='Continue'])[1]
    IF    ${continue_visible} == True
        Safe Click Element    (//span[normalize-space()='Continue'])[1]
    END

Fill Pricing Details[Renewal]
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//span[normalize-space()='Option 1'])[1]    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Proj Expo'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonCalculate'])[1]
    Sleep    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Loss Expo History'])[1]    20s
    Sleep    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_ctrl0_CheckBoxSingle'])[1]
    Sleep    2s
    Safe Click Element
    ...    (//a[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_listViewExposure_deleteLossExposure'])[1]
    Sleep    5s
    ${variableExist}=    Is Element Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[5]
    IF    ${variableExist} == 2
        Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[5]
    END
    Sleep    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Man Rating'])[1]    10s
    Safe Click Element    (//span[normalize-space()='Correct'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonCalculate'])[1]
    Sleep    5s
    ${variableExist}=    Is Element Visible    (//span[normalize-space()='Ok'])[1]
    IF    ${variableExist} == True
        Safe Click Element    (//span[normalize-space()='Ok'])[1]
    END
    Sleep    5s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]    5s
    Wait Until Element Is Visible    (//legend[normalize-space()='Summary of Pricing'])[1]    30s
    Wait Until Element Is Visible    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbTotalPolicyPremium'])[1]    30s
    Sleep    10s
    TRY
        ${total_policy_premium}=    Get From Dictionary    ${current_row}    AG
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbTotalPolicyPremium'])[1]
        ...    100000
    EXCEPT
        ${total_policy_premium}=    Get From Dictionary    ${current_row}    AG
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbTotalPolicyPremium'])[1]
        ...    100000
    END
    Sleep    10s
    TRY
        ${exposure_premium}=    Get From Dictionary    ${current_row}    AH
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_TextBoxExposurePremium'])[1]
        ...    ${exposure_premium}
    EXCEPT
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_TextBoxExposurePremium'])[1]
        ...    ${exposure_premium}
    END
    ${rate_per}=    Get From Dictionary    ${current_row}    AI
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListRatePer'])[1]
    ...    ${rate_per}
    Sleep    2s
    ${exposure_base}=    Get From Dictionary    ${current_row}    AJ
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListExposureBase'])[1]
    ...    ${exposure_base}
    Sleep    2s
    ${exposure_premium}=    Get From Dictionary    ${current_row}    AK
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListExposureType'])[1]
    ...    ${exposure_premium}
    Sleep    2s
    ${premium_percentage}=    Get From Dictionary    ${current_row}    AL
    Wait Until Keyword Succeeds    3x    5s    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbMinAuditPremiumPercentage'])[1]
    ...    ${premium_percentage}
    Sleep    2s
    ${tria_premium}=    Get From Dictionary    ${current_row}    AM
    Wait Until Keyword Succeeds    3x    5s    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbTriaPremium'])[1]
    ...    ${tria_premium}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Sleep    5s
    Wait Until Element Is Visible    (//a[normalize-space()='Rate Monitor'])[1]    5s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]

Fill Pricing Details[Copy Risk]
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//span[normalize-space()='Option 1'])[1]    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Proj Expo'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonCalculate'])[1]
    Sleep    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Loss Expo History'])[1]    20s
    Sleep    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Man Rating'])[1]    10s
    Safe Click Element    (//span[normalize-space()='Correct'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonCalculate'])[1]
    Sleep    5s
    ${variableExist}=    Is Element Visible    (//span[normalize-space()='Ok'])[1]
    IF    ${variableExist} == True
        Safe Click Element    (//span[normalize-space()='Ok'])[1]
    END
    Sleep    5s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Sleep    5s
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]    5s
    Wait Until Element Is Visible    (//legend[normalize-space()='Summary of Pricing'])[1]    30s
    Wait Until Element Is Visible    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbTotalPolicyPremium'])[1]    30s
    Sleep    10s
    TRY
        ${total_policy_premium}=    Get From Dictionary    ${current_row}    AG
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbTotalPolicyPremium'])[1]
        ...    100000
    EXCEPT
        ${total_policy_premium}=    Get From Dictionary    ${current_row}    AG
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbTotalPolicyPremium'])[1]
        ...    100000
    END
    Sleep    10s
    TRY
        ${exposure_premium}=    Get From Dictionary    ${current_row}    AH
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_TextBoxExposurePremium'])[1]
        ...    ${exposure_premium}
    EXCEPT
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_TextBoxExposurePremium'])[1]
        ...    ${exposure_premium}
    END
    ${rate_per}=    Get From Dictionary    ${current_row}    AI
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListRatePer'])[1]
    ...    ${rate_per}
    Sleep    2s
    ${exposure_base}=    Get From Dictionary    ${current_row}    AJ
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListExposureBase'])[1]
    ...    ${exposure_base}
    Sleep    2s
    ${exposure_premium}=    Get From Dictionary    ${current_row}    AK
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListExposureType'])[1]
    ...    ${exposure_premium}
    Sleep    2s
    ${premium_percentage}=    Get From Dictionary    ${current_row}    AL
    Wait Until Keyword Succeeds    3x    5s    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbMinAuditPremiumPercentage'])[1]
    ...    ${premium_percentage}
    Sleep    2s
    ${tria_premium}=    Get From Dictionary    ${current_row}    AM
    Wait Until Keyword Succeeds    3x    5s    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbTriaPremium'])[1]
    ...    ${tria_premium}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Sleep    5s

Fill Pricing Details[Reissue]
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//span[normalize-space()='Option 1'])[1]    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Proj Expo'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonCalculate'])[1]
    Sleep    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Loss Expo History'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Man Rating'])[1]    10s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Sleep    5s
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]    5s
    Wait Until Element Is Visible    (//legend[normalize-space()='Summary of Pricing'])[1]    30s
    Wait Until Element Is Visible    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbTotalPolicyPremium'])[1]    30s
    Sleep    10s
    TRY
        ${total_policy_premium}=    Get From Dictionary    ${current_row}    AG
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbTotalPolicyPremium'])[1]
        ...    100000
    EXCEPT
        ${total_policy_premium}=    Get From Dictionary    ${current_row}    AG
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbTotalPolicyPremium'])[1]
        ...    100000
    END
    Sleep    10s
    TRY
        ${exposure_premium}=    Get From Dictionary    ${current_row}    AH
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_TextBoxExposurePremium'])[1]
        ...    ${exposure_premium}
    EXCEPT
        ${exposure_premium}=    Get From Dictionary    ${current_row}    AH
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_TextBoxExposurePremium'])[1]
        ...    ${exposure_premium}
    END
    ${rate_per}=    Get From Dictionary    ${current_row}    AI
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListRatePer'])[1]
    ...    ${rate_per}
    Sleep    2s
    ${exposure_base}=    Get From Dictionary    ${current_row}    AJ
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListExposureBase'])[1]
    ...    ${exposure_base}
    Sleep    2s
    ${exposure_premium}=    Get From Dictionary    ${current_row}    AK
    Wait Until Keyword Succeeds    3x    5s    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_listViewSummaryOfPricingExposure_ctrl0_DropDownListExposureType'])[1]
    ...    ${exposure_premium}
    Sleep    2s
    ${premium_percentage}=    Get From Dictionary    ${current_row}    AL
    Wait Until Keyword Succeeds    3x    5s    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbMinAuditPremiumPercentage'])[1]
    ...    ${premium_percentage}
    Sleep    2s
    ${tria_premium}=    Get From Dictionary    ${current_row}    AM
    Wait Until Keyword Succeeds    3x    5s    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_Option1SummaryOfPricing_tbTriaPremium'])[1]
    ...    ${tria_premium}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Sleep    5s

Pre Bind Endorsement
    ${category_list}=    Get From Dictionary    ${current_row}    AN
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]
    ...    ${category_list}
    Sleep    2s
    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    AO
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlManuscriptTemplates'])[1]
    ...    ${endorsement_type_list}
    Sleep    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btAddEndorsement'])[1]
    Sleep    3s
    ${form_number}=    Get From Dictionary    ${current_row}    AM
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxFormNo'])[1]
    ...    CGL 12 34 56 78
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_EndorsementStandardButtons_btnSubmit'])[1]
    Sleep    3s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s
    ${variableExist}=    Is Element Visible    (//span[normalize-space()='Ok'])[1]
    IF    ${variableExist} == True
        Safe Click Element    (//span[normalize-space()='Ok'])[1]
    END
    Safe Click Element    xpath=//*[@id="EndorsementTable"]/thead/tr/th[3]
    Sleep    2s
    Safe Click Element    xpath=//*[@id="EndorsementTable"]/thead/tr/th[3]
    Sleep    2s
    ${yes_checkboxes}=    Get WebElements    xpath=//tr[td[normalize-space()='Yes']]//input[@type='checkbox']
    ${count}=    Get Length    ${yes_checkboxes}
    IF    ${count} > 0
        FOR    ${checkbox}    IN    @{yes_checkboxes}
            Click Element    ${checkbox}
        END
        Safe Click Element    xpath=//*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_repeaterEndorsement_ctl00_deleteEndorsement"]/span
        Sleep    5s
        ${variableExist}=    Is Element Visible    xpath=/html/body/div[3]/div[3]/div/button[1]
        IF    ${variableExist} == True
            Safe Click Element    xpath=/html/body/div[3]/div[3]/div/button[1]
        END
    END
    Safe Click Element    xpath=//*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext"]
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    Sleep    5s
    Check All Subjectivities
    FOR    ${index}    IN RANGE    10    20
        ${subjective_exist}=    Execute Javascript    return document.getElementById('ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_repeaterSubjectivities_ctl${index}_cbMandatory') != null;
        IF    ${subjective_exist} == True
            ${subjective_checked}=    Execute Javascript    return document.getElementById('ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_repeaterSubjectivities_ctl${index}_cbMandatory').checked;
            IF    ${subjective_checked} == True
                Wait Until Keyword Succeeds    3x    2s    Execute Javascript    document.getElementById('ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_repeaterSubjectivities_ctl${index}_cbSatisfied').click()
                Sleep    1s
            END
        ELSE
            BREAK
        END
    END
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s

Pre Bind Endorsement[Copy Risk]
    ${category_list}=    Get From Dictionary    ${current_row}    AN
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]
    ...    ${category_list}
    Sleep    2s
    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    AO
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlManuscriptTemplates'])[1]
    ...    ${endorsement_type_list}
    Sleep    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btAddEndorsement'])[1]
    Sleep    3s
    ${form_number}=    Get From Dictionary    ${current_row}    AM
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxFormNo'])[1]
    ...    CGL 12 34 56 78
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_EndorsementStandardButtons_btnSubmit'])[1]
    Sleep    3s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s
    ${variableExist}=    Is Element Visible    (//span[normalize-space()='Ok'])[1]
    IF    ${variableExist} == True
        Safe Click Element    (//span[normalize-space()='Ok'])[1]
    END
    Safe Click Element    xpath=//*[@id="EndorsementTable"]/thead/tr/th[3]
    Sleep    2s
    Safe Click Element    xpath=//*[@id="EndorsementTable"]/thead/tr/th[3]
    Sleep    2s
    ${yes_checkboxes}=    Get WebElements    xpath=//tr[td[normalize-space()='Yes']]//input[@type='checkbox']
    ${count}=    Get Length    ${yes_checkboxes}
    IF    ${count} > 0
        FOR    ${checkbox}    IN    @{yes_checkboxes}
            Click Element    ${checkbox}
        END
        Safe Click Element    xpath=//*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_repeaterEndorsement_ctl00_deleteEndorsement"]/span
        Sleep    5s
        ${variableExist}=    Is Element Visible    xpath=/html/body/div[3]/div[3]/div/button[1]
        IF    ${variableExist} == True
            Safe Click Element    xpath=/html/body/div[3]/div[3]/div/button[1]
        END
    END
    Safe Click Element    xpath=//*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext"]
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    Sleep    5s
    Check All Subjectivities
    FOR    ${index}    IN RANGE    10    20
        ${subjective_exist}=    Execute Javascript    return document.getElementById('ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_repeaterSubjectivities_ctl${index}_cbMandatory') != null;
        IF    ${subjective_exist} == True
            ${subjective_checked}=    Execute Javascript    return document.getElementById('ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_repeaterSubjectivities_ctl${index}_cbMandatory').checked;
            IF    ${subjective_checked} == True
                Wait Until Keyword Succeeds    3x    2s    Execute Javascript    document.getElementById('ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_repeaterSubjectivities_ctl${index}_cbSatisfied').click()
                Sleep    1s
            END
        ELSE
            BREAK
        END
    END
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s

Quote
    Wait Until Element Is Visible    (//a[normalize-space()='Quote Final'])[1]    20s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[3]
    ${comments_under_quote}=    Get From Dictionary    ${current_row}    AP
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxFinalQuoteComments'])[1]
    ...    ${comments_under_quote}
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_BtnGenQuote'])[1]

Pre Bind Endorsement[Reissue]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s

Pre Bind Endorsement[Renewal]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    Sleep    5s
    Check All Subjectivities
    FOR    ${index}    IN RANGE    10    20
        ${subjective_exist}=    Execute Javascript    return document.getElementById('ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_repeaterSubjectivities_ctl${index}_cbMandatory') != null;
        IF    ${subjective_exist} == True
            ${subjective_checked}=    Execute Javascript    return document.getElementById('ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_repeaterSubjectivities_ctl${index}_cbMandatory').checked;
            IF    ${subjective_checked} == True
                Wait Until Keyword Succeeds    3x    2s    Execute Javascript    document.getElementById('ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlSubjectivities_repeaterSubjectivities_ctl${index}_cbSatisfied').click()
                Sleep    1s
            END
        ELSE
            BREAK
        END
    END
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s

Quote[Reissue]
    Wait Until Element Is Visible    (//a[normalize-space()='Quote Final'])[1]    20s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_BtnGenQuote'])[1]

Ready To Bind
    Wait Until Element Is Visible    (//a[normalize-space()='Bind'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Bind'])[1]
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Check All Subjectivities
    FOR    ${index}    IN RANGE    10    20
        ${subjective_exist}=    Execute Javascript    return document.getElementById('ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlCheckSubjectivities_repeaterSubjectivities_ctl${index}_cbMandatory') != null;
        IF    ${subjective_exist} == True
            ${subjective_checked}=    Execute Javascript    return document.getElementById('ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlCheckSubjectivities_repeaterSubjectivities_ctl${index}_cbMandatory').checked;
            IF    ${subjective_checked} == True
                Wait Until Keyword Succeeds    3x    2s    Execute Javascript    document.getElementById('ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlCheckSubjectivities_repeaterSubjectivities_ctl${index}_cbSatisfied').click()
                Sleep    1s
            END
        ELSE
            BREAK
        END
    END
    Safe Click Element    xpath=//*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_BtnNext"]/span
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_buttonBind'])[1]

Book
    [Arguments]    ${download_policy}
    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Book and Issue'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s
    ${insured_type}=    Get From Dictionary    ${current_row}    AQ
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlInsuredType'])[1]
    ...    ${insured_type}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    ${Jurisdiction}=    Get From Dictionary    ${current_row}    AR
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownJurisdiction'])[1]
    ...    ${Jurisdiction}
    ${market_hazard}=    Get From Dictionary    ${current_row}    AS
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownMarketHazard'])[1]
    ...    ${market_hazard}
    ${pollution_code}=    Get From Dictionary    ${current_row}    AT
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownPollutionCode'])[1]
    ...    ${pollution_code}
    ${business_description}=    Get From Dictionary    ${current_row}    AU
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxBusinessDescription'])[1]
    ...    ${business_description}
    ${coverage_territory}=    Get From Dictionary    ${current_row}    AV
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownCoverageTerritory'])[1]
    ...    ${coverage_territory}
    ${financial_rating}=    Get From Dictionary    ${current_row}    AW
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownFinancialRating'])[1]
    ...    ${financial_rating}
    ${defense_code}=    Get From Dictionary    ${current_row}    AX
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownDefenseCode'])[1]
    ...    ${defense_code}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[4]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[5]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_buttonBook'])[1]

Risk UW eFile
    Wait Until Element Is Visible    (//a[normalize-space()='Risk UW eFile'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Risk UW eFile'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s
    Sleep    30s
    Wait Until Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnReload'])[1]
    ...    20s
    Click Element If Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnReload'])[1]
    Sleep    30s
    Wait Until Element Is Visible    (//td[@role='gridcell'])[5]//a    120s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnReload'])[1]
    IF    ${download_policy} == 2
        Wait Until Element Is Visible    (//a[normalize-space()='Policy.pdf'])[1]    60s
        Safe Click Element    (//a[normalize-space()='Policy.pdf'])[1]
        Click Element If Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    ELSE
        Sleep    240s
    END
    Safe Click Element    xpath=//div[@class='RiskHeaderColumnOne']//div[1]

Book[Reissue]
    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Book and Issue'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[5]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_buttonBook'])[1]
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
    ${category_list}=    Get From Dictionary    ${current_row}    AY
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]
    ...    ${category_list}
    Sleep    2s
    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    AZ
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlManuscriptTemplates'])[1]
    ...    ${endorsement_type_list}
    Sleep    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btAddEndorsement'])[1]
    Sleep    3s
    ${form_number}=    Get From Dictionary    ${current_row}    AM
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxFormNo'])[1]
    ...    CGL 12 34 56 78
    Safe Click Element    (//span[normalize-space()='Additional Premium'])[1]
    ${endorsement_premium}=    Get From Dictionary    ${current_row}    BA
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementPremium_tbEndorsementPremium'])[1]
    ...    ${endorsement_premium}
    ${effective_date_post_bind}=    Get From Dictionary    ${current_row}    BB
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementEffectiveDate_datepickerEffectiveDate_textDate'])[1]
    ...    ${effective_date_post_bind}
    Sleep    3s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_EndorsementStandardButtons_btnSubmit'])[1]
    Safe Click Element    xpath=//div[@class='RiskHeaderColumnOne']//div[1]

Renewal
    Wait Until Element Is Visible    (//a[normalize-space()='Endorsement'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Endorsement'])[1]
    ${Renewal_category_list}=    Get From Dictionary    ${current_row}    BC
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]
    ...    ${Renewal_category_list}
    Sleep    2s
    ${renewal_endorsement_type_list}=    Get From Dictionary    ${current_row}    BD
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdEndorsementTypeList'])[1]
    ...    ${renewal_endorsement_type_list}
    Sleep    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btEditEndorsement'])[1]
    Sleep    3s
    ${renewal_effective_date_post_bind}=    Get From Dictionary    ${current_row}    BE
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementEffectiveDate_datepickerEffectiveDate_textDate'])[1]
    ...    ${renewal_effective_date_post_bind}
    ${renewal_policy_expiry_date}=    Get From Dictionary    ${current_row}    BF
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_datepickerPolicyExpiryDate_textDate'])[1]
    ...    ${renewal_policy_expiry_date}
    Sleep    3s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_EndorsementStandardButtons_btnSubmit'])[1]
    Safe Click Element    xpath=//div[@class='RiskHeaderColumnOne']//div[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Renew'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Renew'])[1]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Submission
    Risk Creation[Renewal]
    Pre Bind Endorsement[Renewal]
    Quote[Reissue]
    Ready To Bind
    Book[Reissue]
    Safe Click Element    xpath=//div[@class='RiskHeaderColumnOne']//div[1]

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
    Book[Reissue]
    Safe Click Element    xpath=//div[@class='RiskHeaderColumnOne']//div[1]

Copy Risk
    Wait Until Element Is Visible    (//a[normalize-space()='Copy Risk'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Copy Risk'])[1]
    Sleep    10s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[4]
    Fill Insured Details[Copy Risk]
    Fill Pricing Details[Copy Risk]
    Pre Bind Endorsement[Copy Risk]
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
    Wait Until Keyword Succeeds    15x    2s    Element Should Contain    ${location}    ${status_expected}

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

