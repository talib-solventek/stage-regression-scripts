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

CA Excess and Umbrella

    Open Workbook    ${DATA_SOURCE}

    ${excel_rows}=    Read Worksheet    CAExcessAndUmbrella

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

    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]

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

    Safe Click Element    xpath=//a[normalize-space()='CA Excess and Umbrella']

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

    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_labelStatus'])[1]

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

    Click Next And Wait For Element

    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]

    Click Next And Wait For Element    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]

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

    Click Next And Wait For Element    (//legend[normalize-space()='Broker Info'])[1]

    ${brokfirm}=    Get From Dictionary    ${current_row}    M

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_TextBoxBrokerFirm'])[1]

    ...    ${brokfirm}

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_ButtonBrokerFirmSearchContains'])[1]

    Wait Until Element Is Visible

    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]

    ...    30s

    ${table_elements}=    Get WebElements

    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]

    Log    Found ${table_elements.__len__()} elements in the broker info table

    Safe Click Element

    ...    xpath=(//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"])[1]

    ${value1}=    Is Element Visible    (//span[normalize-space()='Yes'])[1]

    IF    ${value1} == True

        Safe Click Element    (//span[normalize-space()='Yes'])[1]

    END

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

    Safe Click Element

    ...    xpath=(//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"])[1]

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

    Sleep    4s

    Wait Until Element Is Visible    (//a[normalize-space()='New Policy Info'])[1]    20s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]

    Sleep    2s

    ${policy_type}=    Get From Dictionary    ${current_row}    O

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlPolicyType'])[1]

    ...    ${policy_type}

    ${quote_type}=    Get From Dictionary    ${current_row}    P

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlQuoteType'])[1]

    ...    ${quote_type}

    ${exposure_base}=    Get From Dictionary    ${current_row}    Q

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ddlExposureBase'])[1]

    ...    ${exposure_base}

    Sleep    5s

    ${canadian_exposure}=    Get From Dictionary    ${current_row}    R

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbCanadianExposure'])[1]

    ...    ${canadian_exposure}

    ${us_exposure}=    Get From Dictionary    ${current_row}    S

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbUSExposure'])[1]

    ...    ${us_exposure}

    ${other_exposure}=    Get From Dictionary    ${current_row}    T

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbOtherExposure'])[1]

    ...    ${other_exposure}

    ${total_limit}=    Get From Dictionary    ${current_row}    U

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbTotalLimit'])[1]

    ...    ${total_limit}

    Click Next And Wait For Element    (//a[normalize-space()='Additional New Policy Info'])[1]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]

    ${class_code}=    Get From Dictionary    ${current_row}    V

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbSICCode'])[1]

    ...    ${class_code}

    ${class_code_full}=    Get From Dictionary    ${current_row}    W

    Safe Click Element    (//a[normalize-space()='${class_code_full}'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Schedule of Underlyers'])[1]

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_underLyerSched_underLyerList_0'])[1]

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_underLyerSched_ctrlCommercialGeneralLiability_btnAdd'])[1]

    Sleep    20s

    ${carrier}=    Get From Dictionary    ${current_row}    W

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_underLyerSched_ctrlCommercialGeneralLiability_AddEditPopup_tbCarrier'])[1]

    ...    ABCGF

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]

    Safe Click Element    (//span[normalize-space()='Add'])[1]

    Sleep    10s

    Click Next And Wait For Element    (//a[normalize-space()='Rating Input - CGL'])[1]

    ${cgl_limit}=    Get From Dictionary    ${current_row}    X

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRatingInputCGl_ctrlRatingInputCGlLiability_tbPrimaryGlLimit'])[1]

    ...    ${cgl_limit}

    ${cgl_premium}=    Get From Dictionary    ${current_row}    Y

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRatingInputCGl_ctrlRatingInputCGlLiability_tbPrimaryGlPremium'])[1]

    ...    ${cgl_premium}

    ${methodology}=    Get From Dictionary    ${current_row}    Z

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRatingInputCGl_ctrlRatingInputCGlLiability_ddlMetholodgy'])[1]

    ...    ${methodology}

    Click Next And Wait For Element    (//a[normalize-space()='Rating Summary'])[1]

    ${comments}=    Get From Dictionary    ${current_row}    AA

    Safe Input Text

    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ctrlRatingSummaryComments_tbPrimaryPremiumComment'])[1]

    ...    ${comments}

    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]



Risk Creation[Renewal]

    Fill Insured Details[Renewal]

    Fill Pricing Details[Renewal]



Fill Insured Details[Renewal]

    Click Next And Wait For Element

    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Submission Details'])[1]

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_txtNAICSCode'])[1]

    ...    111110 SOYBEAN FARMING

    Click Next And Wait For Element    (//a[normalize-space()='Broker Firm'])[1]

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

    Sleep    4s

    Wait Until Element Is Visible    (//a[normalize-space()='Broker Contacts'])[1]    20s

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



Fill Pricing Details[Renewal]

    Sleep    4s

    Wait Until Element Is Visible    (//a[normalize-space()='Expiring Policy Info'])[1]    20s

    Click Next And Wait For Element    (//a[normalize-space()='New Policy Info'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Additional New Policy Info'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Schedule of Underlyers'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Rating Input - CGL'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Exposure Comparison'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Rating Summary'])[1]

    ${comments}=    Get From Dictionary    ${current_row}    AA

    Safe Input Text

    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ctrlRatingSummaryComments_tbPrimaryPremiumComment'])[1]

    ...    ${comments}

    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]



Fill Insured Details[Copy Risk]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]

    Sleep    1s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[4]

    Click Next And Wait For Element

    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]

    Click Next And Wait For Element    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]

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

    Click Next And Wait For Element    (//legend[normalize-space()='Broker Info'])[1]

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

    Sleep    4s

    Wait Until Element Is Visible    (//a[normalize-space()='New Policy Info'])[1]    20s

    Click Next And Wait For Element    (//a[normalize-space()='Additional New Policy Info'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Schedule of Underlyers'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Rating Input - CGL'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Rating Summary'])[1]

    ${comments}=    Get From Dictionary    ${current_row}    AA

    Safe Input Text

    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ctrlRatingSummaryComments_tbPrimaryPremiumComment'])[1]

    ...    ${comments}

    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]



Fill Pricing Details[Reissue]

    Sleep    4s

    Wait Until Element Is Visible    (//a[normalize-space()='New Policy Info'])[1]    20s

    Click Next And Wait For Element    (//a[normalize-space()='Additional New Policy Info'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Schedule of Underlyers'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Rating Input - CGL'])[1]

    Click Next And Wait For Element    (//a[normalize-space()='Rating Summary'])[1]

    ${comments}=    Get From Dictionary    ${current_row}    AA

    Safe Input Text

    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_ctrlRatingSummaryComments_tbPrimaryPremiumComment'])[1]

    ...    ${comments}

    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]



Pre Bind Endorsement

    Sleep    4s

    Wait Until Element Is Visible    (//a[normalize-space()='Pre Bind Endorsements'])[1]    20s

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

    Click Next And Wait For Element    (//a[normalize-space()='Subjectivities'])[1]

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

    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    Sleep    5s



Quote

    Sleep    4s

    Wait Until Element Is Visible    (//a[normalize-space()='Quote Final'])[1]    20s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]

    Safe Click Element    (//span[normalize-space()='English Only'])[1]

    ${comments_under_quote}=    Get From Dictionary    ${current_row}    AD

    Safe Input Text

    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_form_tbComments'])[1]

    ...    ${comments_under_quote}

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonCheckQuote'])[1]



Pre Bind Endorsement[Reissue]

    Click Next And Wait For Element    (//a[normalize-space()='Subjectivities'])[1]

    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    Sleep    5s



Ready To Bind

    Sleep    4s

    Wait Until Element Is Visible    (//a[normalize-space()='Bind'])[1]    30s

    Safe Click Element    (//a[normalize-space()='Bind'])[1]

    Sleep    5s

    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s

    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonNext'])[1]

    Sleep    4s

    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonCheckBind'])[1]

    Sleep    20s



Book

    [Arguments]    ${download_policy}

    Sleep    4s

    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s

    Safe Click Element    (//a[normalize-space()='Book and Issue'])[1]

    Sleep    5s

    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s

    ${LOB_Code}=    Get From Dictionary    ${current_row}    AE

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlBook_ddlLOBCode'])[1]

    ...    ${LOB_Code}

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[4]

    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonBook'])[1]



Risk UW eFile

    [Arguments]    ${download_policy}

    Sleep    4s

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

    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnReload'])[1]

    IF    ${download_policy} == True

        Sleep    4s

        Wait Until Element Is Visible    (//a[normalize-space()='Policy.pdf'])[1]    60s

        Safe Click Element    (//a[normalize-space()='Policy.pdf'])[1]

        Safe Click Element If Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]

    ELSE

        Sleep    240s

    END

    Safe Click Element    xpath=//div[@class='RiskHeaderColumnOne']//div[1]



Copy Quote

    Sleep    4s

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



View Risk

    Sleep    4s

    Wait Until Element Is Visible    (//a[normalize-space()='View Risk'])[1]    30s

    Safe Click Element    (//a[normalize-space()='View Risk'])[1]

    Sleep    5s

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonViewCurrent'])[1]

    Sleep    5s

    Verify Status    (//legend[@id='fsMainContentLegend'])[1]    View Risk



Post Bind Endorsement

    Sleep    4s

    Wait Until Element Is Visible    (//a[normalize-space()='Endorsement'])[1]    30s

    Safe Click Element    (//a[normalize-space()='Endorsement'])[1]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='OK'])[1]

    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    AF

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_txtEndorsementType'])[1]

    ...    ${endorsement_type_list}

    Sleep    2s

    Safe Click Element    (//a[normalize-space()='${endorsement_type_list}'])[1]

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btEditEndorsement'])[1]

    Sleep    3s

    ${effective_date_post_bind}=    Get From Dictionary    ${current_row}    AG

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementEffectiveDate_datepickerEffectiveDate_textDate'])[1]

    ...    ${effective_date_post_bind}

    TRY

        ${expiry_date}=    Get From Dictionary    ${current_row}    AH

        Safe Input Text

        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_DatePickerPolicyExpiryDate_textDate'])[1]

        ...    ${expiry_date}

    EXCEPT

        ${expiry_date}=    Get From Dictionary    ${current_row}    AH

        Safe Input Text

        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_DatePickerPolicyExpiryDate_textDate'])[1]

        ...    ${expiry_date}

        Safe Click Element    (//span[normalize-space()='No Premium'])[1]

        Safe Click Element

        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementStandardButtons_btnSubmit'])[1]

    END

    Sleep    3s

    Safe Click Element    xpath=//div[@class='RiskHeaderColumnOne']//div[1]



Renewal

    Sleep    4s

    Wait Until Element Is Visible    (//a[normalize-space()='Renew'])[1]    30s

    Safe Click Element    (//a[normalize-space()='Renew'])[1]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='OK'])[1]

    Sleep    10s

    Verify Status

    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_labelStatus'])[1]

    ...    Submission

    Risk Creation[Renewal]

    Pre Bind Endorsement[Reissue]

    Quote

    Ready To Bind

    Book    False



Reissue

    Sleep    4s

    Wait Until Element Is Visible    (//a[normalize-space()='Reissue'])[1]    30s

    Safe Click Element    (//a[normalize-space()='Reissue'])[1]

    Sleep    10s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='OK'])[1]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[7]

    Verify Status

    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]

    ...    Quoted (Pending, In Revision, RI)

    Sleep    4s

    Wait Until Element Is Visible    (//a[normalize-space()='Edit Quote'])[1]    30s

    Safe Click Element    (//a[normalize-space()='Edit Quote'])[1]

    Fill Pricing Details[Reissue]

    Pre Bind Endorsement[Reissue]

    Quote

    Ready To Bind

    Book    False



Copy Risk

    Sleep    4s

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

    Sleep    4s

    Wait Until Element Is Visible    xpath=//a[normalize-space()='Create New Risk >']    20s

    Safe Click Element    xpath=//a[normalize-space()='Create New Risk >']

    Safe Click Element    xpath=//a[normalize-space()='Create New Risk >']

    Sleep    4s

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

    ${match}=    Evaluate

    ...    '${status_expected}' == '${status_actual}' or ('${status_expected}' == 'Submission (Pricing In Progress)' and '${status_actual}' == 'Submission (Pending)')

    Should Be True    ${match}    Expected status '${status_expected}' but got '${status_actual}'



Click Next And Wait For Element

    [Arguments]    ${expected_element_locator}

    Press Keys    None    PAGE_DOWN

    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    Sleep    3s

    Wait Until Element Is Visible    ${expected_element_locator}    20s



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

    ${passed}=    Run Keyword And Return Status    Wait Until Element Is Visible    ${locator}    5s

    IF    ${passed}    Safe Click Element    ${locator}

