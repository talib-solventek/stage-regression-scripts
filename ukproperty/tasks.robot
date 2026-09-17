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
UK PROPERTY
    Open Workbook    ${DATA_SOURCE}
    ${excel_rows}=    Read Worksheet    UKProperty
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
    Copy Risk
    Sleep    900s
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
    Safe Click Element    xpath=(//a[normalize-space()='UK Property'])[1]
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
    Sleep    5s
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
    ${attribute}=    Get Element Attribute
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUKPostCode'])[1]
    ...    value
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUKPostCode'])[1]
    ...    ${zip1}
    Press Keys
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUKPostCode'])[1]
    ...    A+C+1+3+ +2+A+A
    ${return_value}=    Execute Javascript    document.getElementById('ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUKPostCode').value='AC13 2AA'
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
    ...    London - Head Office
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]
    ...    LMUK
    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    ${department}=    Get From Dictionary    ${current_row}    H
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListDepartment'])[1]
    ...    ${department}
    ${effdate}=    Get From Dictionary    ${current_row}    I
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
    ...    ${effdate}
    ${major_minor}=    Get From Dictionary    ${current_row}    J
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxMajorMinorCode'])[1]
    ...    ${major_minor}
    ${major_minor_full}=    Get From Dictionary    ${current_row}    K
    Safe Click Element    (//a[normalize-space()='${major_minor_full}'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    ${brokfirm}=    Get From Dictionary    ${current_row}    M
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_TextBoxBrokerFirm'])[1]
    ...    ${brokfirm}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_ButtonBrokerFirmSearchStartsWith'])[1]
    Sleep    10s
    Safe Click Element By Index
    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]
    ...    2
    Execute Javascript    window.scrollTo(0, document.body.scrollHeight)
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    ${yes_visible}=    Run Keyword And Return Status    Wait Until Element Is Visible    xpath=/html/body/div[4]/div[3]/div/button[1]/span    5s
    IF    ${yes_visible}
        Safe Click Element    xpath=/html/body/div[4]/div[3]/div/button[1]/span
    END
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    ${brokcont}=    Get From Dictionary    ${current_row}    N
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]
    ...    ${brokcont}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_ButtonBrokerContactSearchStartsWith'])[1]
    Sleep    20s
    Safe Click Element By Index
    ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]
    ...    2
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Pricing Details[First Run]
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    ${currency}=    Get From Dictionary    ${current_row}    O
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGeneralInformation_ddlCurrency'])[1]
    ...    ${currency}
    Safe Click Element    (//span[normalize-space()='Lead'])[1]
    ${documentation_type}=    Get From Dictionary    ${current_row}    P
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlGeneralInformation_ddlDocumentationType'])[1]
    ...    ${documentation_type}
    Sleep    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Property Setup'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Extensions'])[1]    20s
    ${personal_effects}=    Get From Dictionary    ${current_row}    Q
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExtensions_rptExtensionGroups_ctl00_ctrlExtensionGroup_rptStandardExtensions_ctl00_tbLimit'])[1]
    ...    ${personal_effects}
    ${fire_extinguishing}=    Get From Dictionary    ${current_row}    R
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExtensions_rptExtensionGroups_ctl00_ctrlExtensionGroup_rptStandardExtensions_ctl04_tbLimit'])[1]
    ...    ${fire_extinguishing}
    ${unspecified_customers}=    Get From Dictionary    ${current_row}    S
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlExtensions_rptExtensionGroups_ctl01_ctrlExtensionGroup_rptStandardExtensions_ctl04_tbLimit'])[1]
    ...    ${unspecified_customers}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Location Management'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlLocationManagement_btnImportTop'])[1]
    Sleep    10s
    Choose File
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlLocationManagement_fuLocationsCSV'])[1]
    ...    ${CURDIR}${/}Locations.csv
    Safe Click Element    (//span[normalize-space()='Import'])[1]
    Sleep    5s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ctrlLocationManagement$ctl01'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Location Details'])[1]    10s
    ${adl1}=    Get From Dictionary    ${current_row}    W
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlLocationDetails_ctrlAddress_txtAddressLine1'])[1]
    ...    ${adl1}
    ${postal_code}=    Get From Dictionary    ${current_row}    X
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlLocationDetails_ctrlAddress_txtPostalCode'])[1]
    ...    ${postal_code}
    ${construction_description}=    Get From Dictionary    ${current_row}    T
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlLocationDetails_ddlConstructionDescription'])[1]
    ...    ${construction_description}
    ${occupancy_description}=    Get From Dictionary    ${current_row}    U
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlLocationDetails_txtOccupancyDescription'])[1]
    ...    ${occupancy_description}
    ${occupancy_description_full}=    Get From Dictionary    ${current_row}    V
    Safe Click Element    (//a[normalize-space()='${occupancy_description_full}'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnSave'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Aggregate Locations'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rating Summary'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary of Pricing'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Insured Details[Copy Risk]
    Wait Until Element Is Visible    (//a[normalize-space()='Insured Details'])[1]    20s
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
    ${department}=    Get From Dictionary    ${current_row}    H
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListDepartment'])[1]
    ...    ${department}
    ${effdate}=    Get From Dictionary    ${current_row}    I
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
    ...    ${effdate}
    ${major_minor}=    Get From Dictionary    ${current_row}    J
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxMajorMinorCode'])[1]
    ...    ${major_minor}
    ${major_minor_full}=    Get From Dictionary    ${current_row}    K
    Safe Click Element    (//a[normalize-space()='${major_minor_full}'])[1]
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
    Safe Click Element By Index
    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]
    ...    2
    Execute Javascript    window.scrollTo(0, document.body.scrollHeight)
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    ${yes_visible}=    Run Keyword And Return Status    Wait Until Element Is Visible    xpath=/html/body/div[4]/div[3]/div/button[1]/span    5s
    IF    ${yes_visible}
        Safe Click Element    xpath=/html/body/div[4]/div[3]/div/button[1]/span
    END
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    ${brokcont}=    Get From Dictionary    ${current_row}    N
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]
    ...    ${brokcont}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_ButtonBrokerContactSearchStartsWith'])[1]
    Sleep    20s
    Safe Click Element By Index
    ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]
    ...    1
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Pricing Details[Copy Risk]
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s
    Sleep    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Property Setup'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Extensions'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Location Management'])[1]    20s
    Sleep    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Aggregate Locations'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rating Summary'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary of Pricing'])[1]    10s
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
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s
    Sleep    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Property Setup'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Extensions'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Location Management'])[1]    20s
    Sleep    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Aggregate Locations'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rating Summary'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary of Pricing'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Pricing Details[Reissue]
    Wait Until Element Is Visible    (//legend[normalize-space()='General Information'])[1]    20s
    Sleep    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Property Setup'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Policy Info'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Extensions'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Location Management'])[1]    20s
    Sleep    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Aggregate Locations'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Rating Summary'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary of Pricing'])[1]    10s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Pre Bind Endorsement
    ${category_list1}=    Get From Dictionary    ${current_row}    Y
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]
    ...    ${category_list1}
    Sleep    2s
    ${endorsement_type_list1}=    Get From Dictionary    ${current_row}    Z
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
    ${category_list2}=    Get From Dictionary    ${current_row}    AA
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]
    ...    ${category_list2}
    Sleep    2s
    ${endorsement_type_list2}=    Get From Dictionary    ${current_row}    AB
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdEndorsementTypeList'])[1]
    ...    ${endorsement_type_list2}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btAddEndorsement'])[1]
    Sleep    2s
    ${category_list3}=    Get From Dictionary    ${current_row}    AC
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]
    ...    ${category_list3}
    Sleep    2s
    ${endorsement_type_list3}=    Get From Dictionary    ${current_row}    AD
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdEndorsementTypeList'])[1]
    ...    ${endorsement_type_list3}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btAddEndorsement'])[1]
    Sleep    2s
    Safe Click Element    (//div[normalize-space()='Sample'])[1]
    Safe Click Element    (//div[normalize-space()='Sample'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Remove'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s

Quote
    Wait Until Element Is Visible    (//a[normalize-space()='Quote Final'])[1]    20s
    ${cyber_type}=    Get From Dictionary    ${current_row}    AE
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlQuoteFinal_ddlCyberType'])[1]
    ...    ${cyber_type}
    ${cyber_limit}=    Get From Dictionary    ${current_row}    AF
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlQuoteFinal_tbCyberLimit'])[1]
    ...    ${cyber_limit}
    ${comments_under_quote}=    Get From Dictionary    ${current_row}    AG
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlQuoteFinal_tbFinalQuoteComments'])[1]
    ...    ${comments_under_quote}
    ${business_description}=    Get From Dictionary    ${current_row}    AH
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlQuoteFinal_tbBusinessDescription'])[1]
    ...    ${business_description}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnGenQuote'])[1]

Quote[Copy Risk]
    Wait Until Element Is Visible    (//a[normalize-space()='Quote Final'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnGenQuote'])[1]

Pre Bind Endorsement[Reissue]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s
    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s

Quote[Reissue]
    Wait Until Element Is Visible    (//a[normalize-space()='Quote Final'])[1]    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnGenQuote'])[1]

Ready To Bind
    Safe Click Element    (//a[normalize-space()='Bind'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s
    ${peer_underwriter}=    Get From Dictionary    ${current_row}    AI
    Safe Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_BindControl_ddlPeerUnderwriter'])[1]
    ...    ${peer_underwriter}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_buttonBind'])[1]
    Sleep    90s

Book
    [Arguments]    ${download_policy}
    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Book and Issue'])[1]
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[4]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_buttonPreviewBinder'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Contract Certainty'])[1]
    ...    30s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[2]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[3]
    ${source_of_business}=    Get From Dictionary    ${current_row}    AJ
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_ddlSourceofBusiness'])[1]
    ...    ${source_of_business}
    ${placement_type}=    Get From Dictionary    ${current_row}    AK
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_ddlPlacementType'])[1]
    ...    ${placement_type}
    ${ncb_applicable}=    Get From Dictionary    ${current_row}    AL
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_ddlNCBApplicable'])[1]
    ...    ${ncb_applicable}
    ${risk_level}=    Get From Dictionary    ${current_row}    AM
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_ddlRiskLevel'])[1]
    ...    ${risk_level}
    ${eea_exposure}=    Get From Dictionary    ${current_row}    AN
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ConCertControl_txtEEAExposure'])[1]
    ...    ${eea_exposure}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonBook'])[1]

Book[Reissue]
    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Book and Issue'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_buttonBook'])[1]
    Safe Click Element    xpath=//div[@class='RiskHeaderColumnOne']//div[1]

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
    ${category_list}=    Get From Dictionary    ${current_row}    AO
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdCategoryList'])[1]
    ...    ${category_list}
    Sleep    2s
    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    AP
    Safe Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DdEndorsementTypeList'])[1]
    ...    ${endorsement_type_list}
    Sleep    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btEditEndorsement'])[1]
    Sleep    3s
    ${effective_date_post_bind}=    Get From Dictionary    ${current_row}    AQ
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlEndorsementHeader_endorsementEffectiveDate_datepickerEffectiveDate_textDate'])[1]
    ...    ${effective_date_post_bind}
    ${policy_expiry_date}=    Get From Dictionary    ${current_row}    AR
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_datepickerPolicyExpiryDate_textDate'])[1]
    ...    ${policy_expiry_date}
    Sleep    3s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_EndorsementStandardButtons_btnSubmit'])[1]
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
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]
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

Safe Click Element By Index
    [Arguments]    ${locator}    ${index}
    Wait Until Keyword Succeeds    5x    2s    Click Element By Index And Wait    ${locator}    ${index}

Click Element By Index And Wait
    [Arguments]    ${locator}    ${index}
    ${elements}=    Get WebElements    ${locator}
    Log    Found ${elements.__len__()} elements
    Click Element When Visible    ${elements}[${index}]
