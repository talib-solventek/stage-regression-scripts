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
UK COMBINED
    Open Workbook    ${DATA_SOURCE}
    ${excel_rows}=    Read Worksheet    UKCombined
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
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_RiskHeader_labelStatus'])[1]
    ...    Submission (Pricing In Progress)
    Pre Bind Endorsement
    Quote
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]
    ...    Quoted
    Copy Quote
    Ready To Bind
    Sleep    60s
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]
    ...    Bound (Pending)
    Book    False
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]
    ...    Bound
    Post Bind Endorsement
    Renewal
    View Risk
    Close Browser

Login to App
    ${URL}=    Get From Dictionary    ${current_row}    A
    ${NUSER}=    Get From Dictionary    ${current_row}    B
    ${NUSERPASSWORD}=    Get From Dictionary    ${current_row}    C
    Open Available Browser    ${URL}    maximized=${TRUE}    browser_selection=edge
    Set Selenium Implicit Wait    30s
    Click Element When Visible    xpath=//a[normalize-space()='Create New Risk >']
    Click Element When Visible    xpath=(//a[normalize-space()='UK Combined'])[1]
    Sleep    10s
    Click Element When Visible
    ...    id:ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlClearanceSearch_CreateInsured
    Sleep    5s
    ${EMAIL}=    Set Variable    ${NUSER}@libertymutual.com
    Input Text When Element Is Visible    xpath=//input[@type='email']    ${EMAIL}
    Sleep    1s
    Press Keys    xpath=//input[@type='email']    ENTER
    Input Text When Element Is Visible    xpath=//input[@type='password']    ${NUSERPASSWORD}
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
    Click Element When Visible    (//span[normalize-space()='No'])[1]
    ${iname}=    Get From Dictionary    ${current_row}    D
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_TextBoxFirstNamedInsured1'])[1]
    ...    ${iname}
    ${adl1}=    Get From Dictionary    ${current_row}    E
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUKAddressLine1'])[1]
    ...    ${adl1}
    ${zip1}=    Get From Dictionary    ${current_row}    F
    Click Element When Visible
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
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListBranch'])[1]
    ...    London - Head Office
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]
    ...    LMUK
    Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    ${department}=    Get From Dictionary    ${current_row}    H
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListDepartment'])[1]
    ...    ${department}
    ${effdate}=    Get From Dictionary    ${current_row}    I
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
    ...    ${effdate}
    ${industry_code}=    Get From Dictionary    ${current_row}    J
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]
    ...    ${industry_code}
    ${industry_code_full}=    Get From Dictionary    ${current_row}    K
    Click Element When Visible    (//a[normalize-space()='${industry_code_full}'])[1]
    ${major_minor}=    Get From Dictionary    ${current_row}    M
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxMajorMinorCode'])[1]
    ...    ${major_minor}
    ${major_minor_full}=    Get From Dictionary    ${current_row}    M
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    ${brokfirm}=    Get From Dictionary    ${current_row}    O
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_TextBoxBrokerFirm'])[1]
    ...    ${brokfirm}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_ButtonBrokerFirmSearchContains'])[1]
    Sleep    10s
    ${table_elements}=    Get WebElements
    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]
    Log    Found ${table_elements.__len__()} elements in the broker info table
    Click Element    ${table_elements}[2]
    Execute Javascript    window.scrollTo(0, document.body.scrollHeight)
    Press Keys    None    PAGE_DOWN
    Sleep    2s
    Wait Until Keyword Succeeds
    ...    5x
    ...    5s
    ...    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    10s

Fill Pricing Details[First Run]
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    ${documentation_type}=    Get From Dictionary    ${current_row}    Q
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGeneralInformation_ddlDocumentationType'])[1]
    ...    ${documentation_type}
    ${line_of_business}=    Get From Dictionary    ${current_row}    R
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGeneralInformation_ddlLOB'])[1]
    ...    ${line_of_business}
    ${currency}=    Get From Dictionary    ${current_row}    S
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGeneralInformation_ddlCurrency'])[1]
    ...    ${currency}
    Sleep    2s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Property Setup'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Extensions'])[1]    20s
    ${personal_effects}=    Get From Dictionary    ${current_row}    T
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExtensions_rptExtensionGroups_ctl00_ctrlExtensionGroup_rptStandardExtensions_ctl00_tbLimit'])[1]
    ...    ${personal_effects}
    ${fire_extinguishing}=    Get From Dictionary    ${current_row}    U
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExtensions_rptExtensionGroups_ctl00_ctrlExtensionGroup_rptStandardExtensions_ctl04_tbLimit'])[1]
    ...    ${fire_extinguishing}
    ${unspecified_customers}=    Get From Dictionary    ${current_row}    V
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExtensions_rptExtensionGroups_ctl01_ctrlExtensionGroup_rptStandardExtensions_ctl04_tbLimit'])[1]
    ...    ${unspecified_customers}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Location Management'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlLocationManagement_btnImportTop'])[1]
    Sleep    10s
    Choose File
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlLocationManagement_fuLocationsCSV'])[1]
    ...    ${CURDIR}${/}Locations.csv
    Click Element When Visible    (//span[normalize-space()='Import'])[1]
    Sleep    5s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ctrlLocationManagement$ctl01'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Location Details'])[1]    10s
    ${adl1}=    Get From Dictionary    ${current_row}    W
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlLocationDetails_ctrlAddress_txtAddressLine1'])[1]
    ...    ${adl1}
    ${postal_code}=    Get From Dictionary    ${current_row}    X
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlLocationDetails_ctrlAddress_txtPostalCode'])[1]
    ...    ${postal_code}
    ${construction_description}=    Get From Dictionary    ${current_row}    Y
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlLocationDetails_ddlConstructionDescription'])[1]
    ...    ${construction_description}
    ${occupancy_description}=    Get From Dictionary    ${current_row}    Z
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlLocationDetails_txtOccupancyDescription'])[1]
    ...    ${occupancy_description}
    ${occupancy_description_full}=    Get From Dictionary    ${current_row}    AA
    Click Element When Visible    (//a[normalize-space()='${occupancy_description_full}'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnSave'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Aggregate Locations'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rating Summary'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary of Pricing'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Policy Info'])[1]    10s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Individual Claims'])[2]
    ${claim_experience_date}=    Get From Dictionary    ${current_row}    AB
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_elPolicyInfo_ctrlPolicyInfoEL_dpClaimsExperienceDateForNoExperienceRating_textDate'])[1]
    ...    ${claim_experience_date}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Exposures'])[1]    10s
    ${industry}=    Get From Dictionary    ${current_row}    AC
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlMainIndustry'])[1]
    ...    ${industry}
    Sleep    10s
    ${sub_industry}=    Get From Dictionary    ${current_row}    AD
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlMainSubIndustry'])[1]
    ...    ${sub_industry}
    ${exposure}=    Get From Dictionary    ${current_row}    AE
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewOccupations_ctrl0_ddlExposureType'])[1]
    ...    ${exposure}
    ${exposure_amount}=    Get From Dictionary    ${current_row}    AF
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewOccupations_ctrl0_tbWageroll'])[1]
    ...    ${exposure_amount}
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    ${risk_management}=    Get From Dictionary    ${current_row}    AG
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlRiskManagementGrade'])[1]
    ...    ${risk_management}
    ${claim_history}=    Get From Dictionary    ${current_row}    AH
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbClaimsHistoryDiscount'])[1]
    ...    ${claim_history}
    ${comments}=    Get From Dictionary    ${current_row}    AI
    Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbDiscountComments'])[1]
    ...    ${comments}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Hist Exposure'])[1]    20s
    ${exposure_type}=    Get From Dictionary    ${current_row}    AJ
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_HistoricalExposureForm_ddlExposureType'])[1]
    ...    ${exposure_type}
    ${exposure_amount}=    Get From Dictionary    ${current_row}    AK
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_HistoricalExposureForm_rptHistoricalExposures_ctl00_txtExposureAmount'])[1]
    ...    ${exposure_amount}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Individual Claims'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_IndividualClaimsControl_radGridIndividualClaims_ctl00_ctl03_ctl01_btnInsert'])[1]
    Sleep    5s
    ${policy_year}=    Get From Dictionary    ${current_row}    AL
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_IndividualClaimsControl_radGridIndividualClaims_ctl00_ctl04_txtPolicyYear'])[1]
    ...    ${policy_year}
    ${incident_date}=    Get From Dictionary    ${current_row}    AM
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_IndividualClaimsControl_radGridIndividualClaims_ctl00_ctl04_txtIncidentDate_textDate'])[1]
    ...    ${incident_date}
    ${currency}=    Get From Dictionary    ${current_row}    AN
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_IndividualClaimsControl_radGridIndividualClaims_ctl00_ctl04_ddlOrigCurrency'])[1]
    ...    ${currency}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Override'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    AO
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_OverrideControl_rptLiabilityOverrideOccupation_ctl01_tbRateOverride'])[1]
    ...    ${rate_override}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Policy Info'])[1]    10s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Individual Claims'])[2]
    ${claim_experience_date}=    Get From Dictionary    ${current_row}    AP
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_plPolicyInfo_ctrlPolicyInfoPL_dpClaimsExperienceDateForNoExperienceRating_textDate'])[1]
    ...    ${claim_experience_date}
    ${retention_type}=    Get From Dictionary    ${current_row}    AQ
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_plPolicyInfo_ctrlPolicyInfoPL_ddlRetentionType'])[1]
    ...    ${retention_type}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Exposures'])[1]    10s
    ${exposure1}=    Get From Dictionary    ${current_row}    AR
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewOccupations_ctrl0_ddlExposureType'])[1]
    ...    ${exposure1}
    ${exposure_amount1}=    Get From Dictionary    ${current_row}    AS
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewOccupations_ctrl0_tbWageroll'])[1]
    ...    ${exposure_amount1}
    ${exposure2}=    Get From Dictionary    ${current_row}    AT
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewOccupations_ctrl1_ddlExposureType'])[1]
    ...    ${exposure2}
    ${exposure_amount2}=    Get From Dictionary    ${current_row}    AU
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewOccupations_ctrl1_tbWageroll'])[1]
    ...    ${exposure_amount2}
    ${exposure3}=    Get From Dictionary    ${current_row}    AV
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewOccupations_ctrl2_ddlExposureType'])[1]
    ...    ${exposure3}
    ${exposure_amount3}=    Get From Dictionary    ${current_row}    AW
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_listViewOccupations_ctrl2_tbWageroll'])[1]
    ...    ${exposure_amount3}
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    ${risk_management}=    Get From Dictionary    ${current_row}    AX
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlRiskManagementGrade'])[1]
    ...    ${risk_management}
    ${claim_history}=    Get From Dictionary    ${current_row}    AY
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbClaimsHistoryDiscount'])[1]
    ...    ${claim_history}
    ${comments}=    Get From Dictionary    ${current_row}    AZ
    Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbDiscountComments'])[1]
    ...    ${comments}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Hist Exposure'])[1]    20s
    ${exposure_type}=    Get From Dictionary    ${current_row}    BA
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_HistoricalExposureForm_ddlExposureType'])[1]
    ...    ${exposure_type}
    ${exposure_amount}=    Get From Dictionary    ${current_row}    BB
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_HistoricalExposureForm_rptHistoricalExposures_ctl00_txtExposureAmount'])[1]
    ...    ${exposure_amount}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Individual Claims'])[1]    20s
    Click Element When Visible    (//span[normalize-space()='Yes'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_IndividualClaimsControl_radGridIndividualClaims_ctl00_ctl03_ctl01_btnInsert'])[1]
    ${policy_year}=    Get From Dictionary    ${current_row}    BC
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_IndividualClaimsControl_radGridIndividualClaims_ctl00_ctl04_txtPolicyYear'])[1]
    ...    ${policy_year}
    ${incident_date}=    Get From Dictionary    ${current_row}    BD
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_IndividualClaimsControl_radGridIndividualClaims_ctl00_ctl04_txtIncidentDate_textDate'])[1]
    ...    ${incident_date}
    ${currency}=    Get From Dictionary    ${current_row}    BE
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_IndividualClaimsControl_radGridIndividualClaims_ctl00_ctl04_ddlOrigCurrency'])[1]
    ...    ${currency}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Override'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    BF
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_OverrideControl_rptLiabilityOverrideOccupation_ctl01_tbRateOverride'])[1]
    ...    ${rate_override}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    BG
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ControlSummary_tbModelPremiumEL'])[1]
    ...    ${rate_override}
    ${rate_override}=    Get From Dictionary    ${current_row}    BH
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ControlSummary_tbModelPremiumPL'])[1]
    ...    ${rate_override}
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Final Summary'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Insured Details[Copy Risk]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Sleep    5s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListBranch'])[1]
    ...    London - Head Office
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]
    ...    LMUK
    Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    ${department}=    Get From Dictionary    ${current_row}    H
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListDepartment'])[1]
    ...    ${department}
    ${effdate}=    Get From Dictionary    ${current_row}    I
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
    ...    ${effdate}
    ${industry_code}=    Get From Dictionary    ${current_row}    J
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]
    ...    ${industry_code}
    ${industry_code_full}=    Get From Dictionary    ${current_row}    K
    Click Element When Visible    (//a[normalize-space()='${industry_code_full}'])[1]
    ${major_minor}=    Get From Dictionary    ${current_row}    M
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxMajorMinorCode'])[1]
    ...    ${major_minor}
    ${major_minor_full}=    Get From Dictionary    ${current_row}    M
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    ${brokfirm}=    Get From Dictionary    ${current_row}    O
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_TextBoxBrokerFirm'])[1]
    ...    ${brokfirm}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_ButtonBrokerFirmSearchContains'])[1]
    Sleep    10s
    ${table_elements}=    Get WebElements
    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]
    Log    Found ${table_elements.__len__()} elements in the broker info table
    Click Element    ${table_elements}[2]
    Execute Javascript    window.scrollTo(0, document.body.scrollHeight)
    Press Keys    None    PAGE_DOWN
    Sleep    2s
    Wait Until Keyword Succeeds
    ...    5x
    ...    5s
    ...    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    10s

Fill Pricing Details[Copy Risk]
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Property Setup'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Extensions'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Location Management'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Aggregate Locations'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rating Summary'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary of Pricing'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Policy Info'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Exposures'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Hist Exposure'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Individual Claims'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Override'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    AO
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_OverrideControl_rptLiabilityOverrideOccupation_ctl01_tbRateOverride'])[1]
    ...    ${rate_override}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Policy Info'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Exposures'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Hist Exposure'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Individual Claims'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Override'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    BF
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_OverrideControl_rptLiabilityOverrideOccupation_ctl01_tbRateOverride'])[1]
    ...    ${rate_override}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Final Summary'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Risk Creation[Renewal]
    Fill Insured Details[Renewal]
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_labelStatus'])[1]
    ...    Submission
    Fill Pricing Details[Renewal]

Fill Insured Details[Renewal]
    Wait Until Element Is Visible    (//a[normalize-space()='Insured Details'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Pricing Details[Renewal]
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Property Setup'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Extensions'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Location Management'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Aggregate Locations'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rating Summary'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary of Pricing'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Policy Info'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Exposures'])[1]    10s
    ${risk_managemnt}=    Get From Dictionary    ${current_row}    CE
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlRiskManagementGrade'])[1]
    ...    ${risk_managemnt}
    ${claims_history}=    Get From Dictionary    ${current_row}    CF
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbClaimsHistoryDiscount'])[1]
    ...    ${claims_history}
    ${comments}=    Get From Dictionary    ${current_row}    AZ
    Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbDiscountComments'])[1]
    ...    ${comments}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Hist Exposure'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Individual Claims'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Override'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    AO
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_OverrideControl_rptLiabilityOverrideOccupation_ctl01_tbRateOverride'])[1]
    ...    ${rate_override}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Policy Info'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Exposures'])[1]    10s
    ${risk_managemnt}=    Get From Dictionary    ${current_row}    CE
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlRiskManagementGrade'])[1]
    ...    ${risk_managemnt}
    ${claims_history}=    Get From Dictionary    ${current_row}    CF
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbClaimsHistoryDiscount'])[1]
    ...    ${claims_history}
    ${comments}=    Get From Dictionary    ${current_row}    AZ
    Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbDiscountComments'])[1]
    ...    ${comments}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Hist Exposure'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Individual Claims'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Override'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    BF
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_OverrideControl_rptLiabilityOverrideOccupation_ctl01_tbRateOverride'])[1]
    ...    ${rate_override}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rate Monitor EL'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rate Monitor PL'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Final Summary'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Pricing Details[Reissue]
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Property Setup'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Extensions'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Location Management'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Aggregate Locations'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rating Summary'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary of Pricing'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Policy Info'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Exposures'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Hist Exposure'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Individual Claims'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='EL Override'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    AO
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_OverrideControl_rptLiabilityOverrideOccupation_ctl01_tbRateOverride'])[1]
    ...    ${rate_override}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Policy Info'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Exposures'])[1]    10s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Hist Exposure'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Individual Claims'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='PL Override'])[1]
    ${rate_override}=    Get From Dictionary    ${current_row}    BF
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_OverrideControl_rptLiabilityOverrideOccupation_ctl01_tbRateOverride'])[1]
    ...    ${rate_override}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Final Summary'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Pre Bind Endorsement
    ${applicable_to}=    Get From Dictionary    ${current_row}    BI
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdSectionList'])[1]
    ...    ${applicable_to}
    Sleep    10s
    ${category_list}=    Get From Dictionary    ${current_row}    BJ
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]
    ...    ${category_list}
    Sleep    2s
    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    BK
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdEndorsementTypeList'])[1]
    ...    ${endorsement_type_list}
    Sleep    2s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btAddEndorsement'])[1]
    Sleep    3s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_EndorsementStandardButtons_btnSubmit'])[1]
    Sleep    3s
    Click Element When Visible    (//div[normalize-space()='Sample'])[1]
    Click Element When Visible    (//div[normalize-space()='Sample'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Remove'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    Click Element When Visible    (//span[normalize-space()='Yes'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s

Quote
    Wait Until Element Is Visible    (//a[normalize-space()='Quote Final'])[1]    20s
    ${cyber_type}=    Get From Dictionary    ${current_row}    BL
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlQuoteFinal_ddlCyberType'])[1]
    ...    ${cyber_type}
    ${cyber_limit}=    Get From Dictionary    ${current_row}    BM
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlQuoteFinal_tbCyberLimit'])[1]
    ...    ${cyber_limit}
    ${comments_under_quote}=    Get From Dictionary    ${current_row}    BN
    Input Text When Element Is Visible
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlQuoteFinal_tbFinalQuoteComments'])[1]
    ...    ${comments_under_quote}
    ${business_description}=    Get From Dictionary    ${current_row}    BO
    Input Text When Element Is Visible
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlQuoteFinal_tbBusinessDescription'])[1]
    ...    ${business_description}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_BtnGenQuote'])[1]

Quote[Copy Risk]
    Wait Until Element Is Visible    (//a[normalize-space()='Quote Final'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnGenQuote'])[1]

Pre Bind Endorsement[Reissue]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s
    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s

Quote[Reissue]
    Wait Until Element Is Visible    (//a[normalize-space()='Quote Final'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_BtnGenQuote'])[1]

Ready To Bind
    Click Element When Visible    (//a[normalize-space()='Bind'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s
    ${peer_underwriter}=    Get From Dictionary    ${current_row}    BP
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_BindControl_ddlPeerUnderwriter'])[1]
    ...    ${peer_underwriter}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_buttonBind'])[1]

Book
    [Arguments]    ${download_policy}
    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Book and Issue'])[1]
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    ${country_el}=    Get From Dictionary    ${current_row}    BQ
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGelTaxes_PremiumAndTaxControlEl_ctrlMultiCurrencyBreakdown_rptPremiumBreakdown_ctl00_ddlCountry'])[1]
    ...    ${country_el}
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[5]
    ${country_pl}=    Get From Dictionary    ${current_row}    BR
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGelTaxes_PremiumAndTaxControlPl_ctrlMultiCurrencyBreakdown_rptPremiumBreakdown_ctl00_ddlCountry'])[1]
    ...    ${country_pl}
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[6]
    Click Element When Visible    (//span[normalize-space()='Next'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s
    ${worldwide_el}=    Get From Dictionary    ${current_row}    BS
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_BookControl_ddlJurisdictionEl'])[1]
    ...    ${worldwide_el}
    ${worldwide_pl}=    Get From Dictionary    ${current_row}    BT
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_BookControl_ddlJurisdictionPl'])[1]
    ...    ${worldwide_pl}
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Contract Certainty'])[1]
    ...    30s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[2]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[3]
    ${source_of_business}=    Get From Dictionary    ${current_row}    BU
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_ddlSourceofBusiness'])[1]
    ...    ${source_of_business}
    ${placement_type}=    Get From Dictionary    ${current_row}    BV
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_ddlPlacementType'])[1]
    ...    ${placement_type}
    ${ncb_applicable}=    Get From Dictionary    ${current_row}    BW
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_ddlNCBApplicable'])[1]
    ...    ${ncb_applicable}
    ${risk_level}=    Get From Dictionary    ${current_row}    BX
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_ddlRiskLevel'])[1]
    ...    ${risk_level}
    ${eea_exposure}=    Get From Dictionary    ${current_row}    BY
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_txtEEAExposure'])[1]
    ...    ${eea_exposure}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonBook'])[1]
    ${country_el}=    Get From Dictionary    ${current_row}    BQ

Book[Reissue]
    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Book and Issue'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_buttonBook'])[1]
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
    Sleep    5s

View Risk
    Wait Until Element Is Visible    (//a[normalize-space()='View Risk'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='View Risk'])[1]
    Sleep    5s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonViewCurrent'])[1]
    Sleep    5s
    Verify Status    (//legend[@id='fsMainContentLegend'])[1]    View Risk

Post Bind Endorsement
    Wait Until Element Is Visible    (//a[normalize-space()='Endorsement'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Endorsement'])[1]
    ${applicable_to}=    Get From Dictionary    ${current_row}    BZ
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdSectionList'])[1]
    ...    ${applicable_to}
    Sleep    10s
    ${category_list}=    Get From Dictionary    ${current_row}    CA
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]
    ...    ${category_list}
    Sleep    2s
    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    CB
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdEndorsementTypeList'])[1]
    ...    ${endorsement_type_list}
    Sleep    2s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btEditEndorsement'])[1]
    Sleep    3s
    ${effective_date_post_bind}=    Get From Dictionary    ${current_row}    CC
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlEndorsementHeader_endorsementEffectiveDate_datepickerEffectiveDate_textDate'])[1]
    ...    ${effective_date_post_bind}
    TRY
        ${expiry_date}=    Get From Dictionary    ${current_row}    CD
        Input Text When Element Is Visible
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_datepickerPolicyExpiryDate_textDate'])[1]
        ...    ${expiry_date}
    EXCEPT
        ${expiry_date}=    Get From Dictionary    ${current_row}    CD
        Input Text When Element Is Visible
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_datepickerPolicyExpiryDate_textDate'])[1]
        ...    ${expiry_date}
    END
    Sleep    3s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_EndorsementStandardButtons_btnSubmit'])[1]
    Click Element When Visible    xpath=//div[@class='RiskHeaderColumnOne']//div[1]
    Sleep    3s

Renewal
    Wait Until Element Is Visible    (//a[normalize-space()='Renew'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Renew'])[1]
    Sleep    10s
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_labelStatus'])[1]
    ...    Submission
    Risk Creation[Renewal]
    Pre Bind Endorsement[Reissue]
    Quote[Reissue]
    Ready To Bind
    Sleep    60s
    Book    False

Reissue
    Wait Until Element Is Visible    (//a[normalize-space()='Reissue'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Reissue'])[1]
    Sleep    10s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[8]
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]
    ...    Quoted (Pending, In Revision, RI)
    Wait Until Element Is Visible    (//a[normalize-space()='Edit Quote'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Edit Quote'])[1]
    Fill Pricing Details[Reissue]
    Pre Bind Endorsement[Reissue]
    Quote[Reissue]
    Ready To Bind
    Sleep    60s
    Book    False

Copy Risk
    Wait Until Element Is Visible    (//a[normalize-space()='Copy Risk'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Copy Risk'])[1]
    Sleep    10s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[4]
    Fill Insured Details[Copy Risk]
    Fill Pricing Details[Copy Risk]
    Pre Bind Endorsement
    Quote
    Ready To Bind
    Sleep    60s
    Book    False

Continue on Next Risk
    Wait Until Element Is Visible    xpath=//a[normalize-space()='Create New Risk >']    20s
    Click Element When Visible    xpath=//a[normalize-space()='Create New Risk >']
    Click Element When Visible    xpath=//a[normalize-space()='Create New Risk >']
    Wait Until Element Is Visible    xpath=//a[normalize-space()='US Primary GL']    30s
    Click Element When Visible    xpath=//a[normalize-space()='US Primary GL']
    Click Element When Visible
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
