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
UK HIGH EXCESS
	Open Workbook    ${DATA_SOURCE}
    ${excel_rows}=    Read Worksheet    UKHighExcess
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
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_labelStatus'])[1]
    ...    Submission (Pricing In Progress)
    Subjectivities
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
    Safe Click Element    xpath=(//a[normalize-space()='UK High Excess'])[1]
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
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_labelStatus'])[1]
    ...    Submission
    Fill Pricing Details[First Run]

Fill Insured Details[First Run]
    Safe Click Element    (//span[normalize-space()='No'])[1]
    ${iname}=    Get From Dictionary    ${current_row}    D
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_TextBoxFirstNamedInsured1'])[1]
    ...    ${iname}
    ${adl1}=    Get From Dictionary    ${current_row}    E
    Safe Input Text
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
    Safe Input Text
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
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUKPoBox'])[1]
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
    ...    Bristol
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]
    ...    BRISTLUK
    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    ${line_of_business}=    Get From Dictionary    ${current_row}    H
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlLOB'])[1]
    ...    ${line_of_business}
    ${effdate}=    Get From Dictionary    ${current_row}    I
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
    ...    ${effdate}
    ${industry_code}=    Get From Dictionary    ${current_row}    J
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]
    ...    ${industry_code}
    ${industry_code_full}=    Get From Dictionary    ${current_row}    K
    Safe Click Element    (//a[normalize-space()='${industry_code_full}'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    ${brokfirm}=    Get From Dictionary    ${current_row}    M
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_TextBoxBrokerFirm'])[1]
    ...    ${brokfirm}
    Safe Click Element
    ...    (//input[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_ButtonBrokerFirmSearchStartsWith"])[1]
    Sleep    10s
    ${table_elements}=    Get WebElements
    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]
    Log    Found ${table_elements.__len__()} elements in the broker info table
    Safe Click Element    ${table_elements}[2]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_ButtonCreateBrokerContact'])[1]
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_TextBoxLastName'])[1]
    ...    broker
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_TextBoxFirstName'])[1]
    ...    test
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_AddressBrokerContact_textBoxUKAddressLine1'])[1]
    ...    addl1
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactInfo_TextBoxEmail'])[1]
    ...    abcde@email.com
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Pricing Details[First Run]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info 1'])[1]    20s
    ${industry}=    Get From Dictionary    ${current_row}    O
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlIndustry'])[1]
    ...    ${industry}
    Sleep    5s
    ${sub_industry}=    Get From Dictionary    ${current_row}    P
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlSubIndustry'])[1]
    ...    ${sub_industry}
    Select Checkbox
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_cblLinesOfBusiness_0'])[1]
    ${policy_trigger}=    Get From Dictionary    ${current_row}    Q
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlPolicyTrigger'])[1]
    ...    ${policy_trigger}
    ${limit}=    Get From Dictionary    ${current_row}    R
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtLimit'])[1]
    ...    ${limit}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Sleep    2s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    ${attachment}=    Get From Dictionary    ${current_row}    S
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtAttachment'])[1]
    ...    ${attachment}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Sleep    2s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[4]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[5]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[7]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[9]
    Sleep    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info 2'])[1]    10s
    ${hazard_code}=    Get From Dictionary    ${current_row}    T
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlHazardCode'])[1]
    ...    ${hazard_code}
    ${pollution_code}=    Get From Dictionary    ${current_row}    U
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlPollutionCode'])[1]
    ...    ${pollution_code}
    ${defense_costs}=    Get From Dictionary    ${current_row}    V
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlDefenceCosts'])[1]
    ...    ${defense_costs}
    ${jurisdiction}=    Get From Dictionary    ${current_row}    W
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlJurisdiction'])[1]
    ...    ${jurisdiction}
    ${coverage_territory}=    Get From Dictionary    ${current_row}    X
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlCoverageTerritory'])[1]
    ...    ${coverage_territory}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Exposure Details'])[1]    20s
    ${revenue_us}=    Get From Dictionary    ${current_row}    Y
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExposureDetails_ctrlMainExpoMeasures_txtRevenueUs'])[1]
    ...    ${revenue_us}
    ${revenue_row}=    Get From Dictionary    ${current_row}    Z
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExposureDetails_ctrlMainExpoMeasures_txtRevenueRow'])[1]
    ...    ${revenue_row}
    ${expiring_revenue_us}=    Get From Dictionary    ${current_row}    AA
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExposureDetails_ctrlMainExpoMeasures_txtExpireRevenueUs'])[1]
    ...    ${expiring_revenue_us}
    ${expiring_revenue_row}=    Get From Dictionary    ${current_row}    AB
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExposureDetails_ctrlMainExpoMeasures_txtExpireRevenueRow'])[1]
    ...    ${expiring_revenue_row}
    ${amount_usa}=    Get From Dictionary    ${current_row}    AC
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExposureDetails_ctrlOtherExpoMeasures_rptExposures_ctl00_txtAmounts'])[1]
    ...    ${amount_usa}
    ${amount_row}=    Get From Dictionary    ${current_row}    AD
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExposureDetails_ctrlOtherExpoMeasures_rptExposures_ctl01_txtAmounts'])[1]
    ...    ${amount_row}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Discount & Loadings'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rating View 1'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]    10s
    ${underwriter_premium}=    Get From Dictionary    ${current_row}    AE
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtUnderwriterSelectedPremium'])[1]
    ...    ${underwriter_premium}
    ${written_line}=    Get From Dictionary    ${current_row}    AF
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtWrittenLine'])[1]
    ...    ${written_line}
    ${signed_line}=    Get From Dictionary    ${current_row}    AG
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtSignedLine'])[1]
    ...    ${signed_line}
    ${reason_of_selection}=    Get From Dictionary    ${current_row}    AH
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtReasonForSelection'])[1]
    ...    ${reason_of_selection}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Insured Details[Copy Risk]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
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
    ...    London - Head Office
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]
    ...    LMUK
    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    ${line_of_business}=    Get From Dictionary    ${current_row}    H
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlLOB'])[1]
    ...    ${line_of_business}
    ${effdate}=    Get From Dictionary    ${current_row}    I
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
    ...    ${effdate}
    ${industry_code}=    Get From Dictionary    ${current_row}    J
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]
    ...    ${industry_code}
    ${industry_code_full}=    Get From Dictionary    ${current_row}    K
    Safe Click Element    (//a[normalize-space()='${industry_code_full}'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    ${brokfirm}=    Get From Dictionary    ${current_row}    M
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_TextBoxBrokerFirm'])[1]
    ...    ${brokfirm}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_ButtonBrokerFirmSearchContains'])[1]
    Sleep    10s
    ${table_elements}=    Get WebElements
    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]
    Log    Found ${table_elements.__len__()} elements in the broker info table
    Safe Click Element    ${table_elements}[2]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    ${brokcont}=    Get From Dictionary    ${current_row}    N
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]
    ...    ${brokcont}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_ButtonBrokerContactSearchStartsWith'])[1]
    Sleep    20s
    ${table_elements}=    Get WebElements
    ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]
    Log    Found ${table_elements.__len__()} elements in the broker contact table
    Safe Click Element    ${table_elements}[0]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Pricing Details[Copy Risk]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info 1'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
	Sleep	10s
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info 2'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Exposure Details'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Discount & Loadings'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rating View 1'])[1]
	Sleep	10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]    10s
    ${underwriter_premium}=    Get From Dictionary    ${current_row}    AP
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtUnderwriterSelectedPremium'])[1]
    ...    ${underwriter_premium}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Risk Creation[Renewal]
    Fill Insured Details[Renewal]
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_labelStatus'])[1]
    ...    Submission
    Fill Pricing Details[Renewal]

Fill Insured Details[Renewal]
    Wait Until Element Is Visible    (//a[normalize-space()='Insured Details'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Pricing Details[Renewal]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info 1'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
	Sleep	10s
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info 2'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Exposure Details'])[1]    20s
    ${revenue_us}=    Get From Dictionary    ${current_row}    AU
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExposureDetails_ctrlMainExpoMeasures_txtRevenueUs'])[1]
    ...    ${revenue_us}
    ${revenue_row}=    Get From Dictionary    ${current_row}    AV
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExposureDetails_ctrlMainExpoMeasures_txtRevenueRow'])[1]
    ...    ${revenue_row}
    ${expiring_revenue_us}=    Get From Dictionary    ${current_row}    AW
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExposureDetails_ctrlMainExpoMeasures_txtExpireRevenueUs'])[1]
    ...    ${expiring_revenue_us}
    ${expiring_revenue_row}=    Get From Dictionary    ${current_row}    AX
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExposureDetails_ctrlMainExpoMeasures_txtExpireRevenueRow'])[1]
    ...    ${expiring_revenue_row}
    ${amount_usa}=    Get From Dictionary    ${current_row}    AY
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExposureDetails_ctrlOtherExpoMeasures_rptExposures_ctl00_txtAmounts'])[1]
    ...    ${amount_usa}
    ${amount_row}=    Get From Dictionary    ${current_row}    AZ
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExposureDetails_ctrlOtherExpoMeasures_rptExposures_ctl01_txtAmounts'])[1]
    ...    ${amount_row}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Discount & Loadings'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rating View 1'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]    10s
    ${underwriter_premium}=    Get From Dictionary    ${current_row}    BA
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtUnderwriterSelectedPremium'])[1]
    ...    ${underwriter_premium}
    ${written_line}=    Get From Dictionary    ${current_row}    BB
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtWrittenLine'])[1]
    ...    ${written_line}
    ${signed_line}=    Get From Dictionary    ${current_row}    BC
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtSignedLine'])[1]
    ...    ${signed_line}
    ${reason_of_selection}=    Get From Dictionary    ${current_row}    BD
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtReasonForSelection'])[1]
    ...    ${reason_of_selection}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRateMonitor_txtOtherChanges'])[1]
    ...    2
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Pricing Details[Reissue]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info 1'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
	Sleep	10s
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info 2'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Exposure Details'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Discount & Loadings'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rating View 1'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]    10s
    ${industry_code_full}=    Get From Dictionary    ${current_row}    K
    ${underwriter_premium}=    Get From Dictionary    ${current_row}    AP
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_txtUnderwriterSelectedPremium'])[1]
    ...    ${underwriter_premium}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Subjectivities
    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    ${cyber_type}=    Get From Dictionary    ${current_row}    AI
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlUkHighExcSubjectivities_ddlCyberType'])[1]
    ...    ${cyber_type}
    ${cyber_limit}=    Get From Dictionary    ${current_row}    AJ
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlUkHighExcSubjectivities_tbCyberLimit'])[1]
    ...    ${cyber_limit}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnGenQuote'])[1]

Subjectivities[Copy Risk]
    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnGenQuote'])[1]

Subjectivities[Reissue]
    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnGenQuote'])[1]

Ready To Bind
    Safe Click Element    (//a[normalize-space()='Bind'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_buttonBind'])[1]

Book
    [Arguments]    ${download_policy}
    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Book and Issue'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Taxes'])[1]
    ...    30s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Book'])[1]
    ...    30s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    ${result}=    Generate Random String    4
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlBook_txtUniqueMarketRef'])[1]
    ...    ${result}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Contract Certainty'])[1]
    ...    30s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[2]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[3]
    ${source_of_business}=    Get From Dictionary    ${current_row}    AK
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_ddlSourceofBusiness'])[1]
    ...    ${source_of_business}
    ${placement_type}=    Get From Dictionary    ${current_row}    AL
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_ddlPlacementType'])[1]
    ...    ${placement_type}
    ${ncb_applicable}=    Get From Dictionary    ${current_row}    AM
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_ddlNCBApplicable'])[1]
    ...    ${ncb_applicable}
    ${risk_level}=    Get From Dictionary    ${current_row}    AN
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_ddlRiskLevel'])[1]
    ...    ${risk_level}
    ${eea_exposure}=    Get From Dictionary    ${current_row}    AO
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_txtEEAExposure'])[1]
    ...    ${eea_exposure}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonBook'])[1]

Copy Quote
    Wait Until Element Is Visible    (//a[normalize-space()='Copy Quote'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Copy Quote'])[1]
    Wait Until Element Is Visible
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_repeaterQuoteOptions_ctl00_HeaderQuote'])[2]
    ...    10s
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_repeaterQuoteOptions_ctl00_HeaderQuote'])[2]
    ...    Option 2 - [Status : Quoted (In Revision)]
    Safe Click Element
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_repeaterQuoteOptions_ctl00_HeaderQuote'])[1]
    Sleep    5s

View Risk
    Wait Until Element Is Visible    (//a[normalize-space()='View Risk'])[1]    30s
    Safe Click Element    (//a[normalize-space()='View Risk'])[1]
    Sleep    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonViewCurrent'])[1]
    Sleep    5s
    Verify Status    (//legend[@id='fsMainContentLegend'])[1]    View Risk

Post Bind Endorsement
    Wait Until Element Is Visible    (//a[normalize-space()='Endorsement'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Endorsement'])[1]
    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    AR
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_txtEndorsementType'])[1]
    ...    ${endorsement_type_list}
    Sleep    2s
    Safe Click Element    (//a[normalize-space()='${endorsement_type_list}'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btAddEndorsement'])[1]
    Sleep    3s
    ${effective_date_post_bind}=    Get From Dictionary    ${current_row}    AS
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementEffectiveDate_datepickerEffectiveDate_textDate'])[1]
    ...    ${effective_date_post_bind}
    TRY
        ${expiry_date}=    Get From Dictionary    ${current_row}    AT
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_DatePickerPolicyExpiryDate_textDate'])[1]
        ...    ${expiry_date}
    EXCEPT
        ${expiry_date}=    Get From Dictionary    ${current_row}    AT
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_DatePickerPolicyExpiryDate_textDate'])[1]
        ...    ${expiry_date}
        Safe Click Element    (//span[normalize-space()='Additional Premium'])[1]
        ${endorsement_premium}=    Get From Dictionary    ${current_row}    AM
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_endorsementPremium_tbEndorsementPremium'])[1]
        ...    1000
        Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
        Safe Click Element
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementStandardButtons_btnSubmit'])[1]
    END
    Sleep    3s
    Safe Click Element    xpath=//div[@class='RiskHeaderColumnOne']//div[1]
    Sleep    600s

Renewal
    Wait Until Element Is Visible    (//a[normalize-space()='Renew'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Renew'])[1]
    Sleep    10s
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_labelStatus'])[1]
    ...    Submission
    Risk Creation[Renewal]
    Subjectivities[Reissue]
    Ready To Bind
    Book    False

Reissue
    Wait Until Element Is Visible    (//a[normalize-space()='Reissue'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Reissue'])[1]
    Sleep    10s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[7]
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]
    ...    Quoted (Pending, In Revision, RI)
    Wait Until Element Is Visible    (//a[normalize-space()='Edit Quote'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Edit Quote'])[1]
    Fill Pricing Details[Reissue]
    Subjectivities[Reissue]
    Ready To Bind
    Book    False

Copy Risk
    Wait Until Element Is Visible    (//a[normalize-space()='Copy Risk'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Copy Risk'])[1]
    Sleep    10s
    Safe Click Element    //span[contains(.,"Original Risks's Account")]
    Fill Insured Details[Copy Risk]
    Fill Pricing Details[Copy Risk]
    Subjectivities[Copy Risk]
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
    Wait Until Keyword Succeeds    15x    2s    Check Status Match    ${location}    ${status_expected}

Check Status Match
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
