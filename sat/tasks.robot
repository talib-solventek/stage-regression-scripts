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
SAT
    Open Workbook    ${DATA_SOURCE}
    ${excel_rows}=    Read Worksheet    SAT
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
        ${url}=    Get From Dictionary    ${current_row}    A
        IF    not $url
            BREAK
        END
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
    Book
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Bound
    Issue Policy
    Copy Risk
    Reissue
    Post Bind Endorsement
    Renewal
    View Risk
    # Close Browser

Login to App
    ${URL}=    Get From Dictionary    ${current_row}    A
    ${NUSER}=    Get From Dictionary    ${current_row}    B
    ${NUSERPASSWORD}=    Get From Dictionary    ${current_row}    C
    Open Browser    ${URL}    chrome    options=add_argument("--inprivate")
    Maximize Browser Window
    Set Selenium Implicit Wait    30s
    Click Element When Visible    xpath=//a[normalize-space()='Create New Risk >']
    Click Element When Visible    xpath=//a[normalize-space()='Stand Alone Terrorism']
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

Risk Creation[First Run]
    Fill Insured Details[First Run]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Submission
    Fill Pricing Details[First Run]

Fill Insured Details[First Run]
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
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'])[1]
    ...    ${zip1}
    Press Keys
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'])[1]
    ...    1+2+3+4+5+6+7+8+9
    ${return_value}=    Execute Javascript    var el = document.getElementById('ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'); if(el) el.value='12345-6789';
    Sleep    1s
    ${pobox}=    Get From Dictionary    ${current_row}    I
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSPoBox'])[1]
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
    ...    Atlanta Retail
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]
    ...    ISIC
    Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    ${effdate}=    Get From Dictionary    ${current_row}    J
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
    ...    ${effdate}
    ${Indcode}=    Get From Dictionary    ${current_row}    K
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]
    ...    ${Indcode}
    Extract Risk Number    First Run
    Scroll Element Into View
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Press Keys    None    PAGE_DOWN
    Wait Until Keyword Succeeds
    ...    3x
    ...    5s
    ...    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    ${brokfirm}=    Get From Dictionary    ${current_row}    M
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_TextBoxBrokerFirm'])[1]
    ...    ${brokfirm}
    Safe Click Element
    ...    xpath=(//input[contains(@id, 'CtrlBrokerFirm_ButtonBrokerFirmSearchContains') or contains(@id, 'CtrlBrokerFirm_ButtonBrokerFirmSearchStartsWith')])[1]
    Wait Until Element Is Visible
    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]
    ...    30s
    Safe Click Element By Index
    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]
    ...    1
    ${value1}=    Is Element Visible    (//span[normalize-space()='Yes'])[1]
    IF    ${value1} == True
        Click Element When Visible    (//span[normalize-space()='Yes'])[1]
    END
    Scroll Element Into View
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
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
        Click Element When Visible    (//span[normalize-space()='Continue'])[1]
    END
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    ${brokcont}=    Get From Dictionary    ${current_row}    N
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]
    ...    ${brokcont}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_ButtonBrokerContactSearchStartsWith'])[1]
    Wait Until Element Is Visible
    ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]
    ...    30s
    Safe Click Element By Index
    ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]
    ...    1
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]
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
            Scroll Element Into View
            ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
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
        Scroll Element Into View
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
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

Fill Pricing Details[First Run]
    Wait Until Element Is Visible    (//a[normalize-space()='Pricing'])[1]    20s
    ${layer_limit}=    Get From Dictionary    ${current_row}    O
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPricing_textBoxLayerLimit'])[1]
    ...    ${layer_limit}
    ${xs_of}=    Get From Dictionary    ${current_row}    P
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPricing_textBoxXSof'])[1]
    ...    ${xs_of}
    ${ironshore_limit}=    Get From Dictionary    ${current_row}    Q
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPricing_textBoxIronshoreGross'])[1]
    ...    ${ironshore_limit}
    ${layer_premium}=    Get From Dictionary    ${current_row}    R
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPricing_textBoxLayerPremium'])[1]
    ...    ${layer_premium}
    ${commision_percentage}=    Get From Dictionary    ${current_row}    S
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPricing_textBoxCommisionPercentage'])[1]
    ...    ${commision_percentage}
    ${minimum_earned_premium}=    Get From Dictionary    ${current_row}    T
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPricing_textBoxMinimumEarnedPremium'])[1]
    ...    ${minimum_earned_premium}
    ${technical_premium}=    Get From Dictionary    ${current_row}    U
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPricing_textBoxIronshoreTechnicalPremiumNet'])[1]
    ...    ${technical_premium}
    ${limit_of_liability}=    Get From Dictionary    ${current_row}    V
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPricing_textBoxAggregateLimitofLiability'])[1]
    ...    ${limit_of_liability}
    ${property_damage}=    Get From Dictionary    ${current_row}    W
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPricing_textBoxPropertyDamage'])[1]
    ...    ${property_damage}
    ${business_interruption}=    Get From Dictionary    ${current_row}    X
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPricing_textBoxBussinessInterruption'])[1]
    ...    ${business_interruption}
    ${others}=    Get From Dictionary    ${current_row}    Y
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPricing_textBoxOthers'])[1]
    ...    ${others}
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Deductible And Coverage'])[1]    20s
    ${policy_form}=    Get From Dictionary    ${current_row}    Z
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlDeductibleAndCoverages_ddlPolicyForm'])[1]
    ...    ${policy_form}
    ${peril_insured}=    Get From Dictionary    ${current_row}    AA
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlDeductibleAndCoverages_ddlPerilInsured'])[1]
    ...    ${peril_insured}
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Quote Items'])[1]
    ${valuation_property}=    Get From Dictionary    ${current_row}    AB
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlQuoteItems_ddlValuation'])[1]
    ...    ${valuation_property}
    ${business_interruption}=    Get From Dictionary    ${current_row}    AC
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlQuoteItems_ddlValuationBusinessInterruption'])[1]
    ...    ${business_interruption}
    Sleep    5s
    ${business_interruption_1}=    Get From Dictionary    ${current_row}    AD
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlQuoteItems_textBoxValuationBusinessInterruption'])[1]
    ...    ${business_interruption_1}
    ${perils}=    Get From Dictionary    ${current_row}    AE
    Input Text When Element Is Visible
    ...    (//textarea[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlQuoteItems_txtPerils'])[1]
    ...    ${perils}
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Additional Terms And Conditions'])[1]
    ${days_to_cancel}=    Get From Dictionary    ${current_row}    AF
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlTermsAndConditions_txtDaysToCancel'])[1]
    ...    ${days_to_cancel}
    ${premium_payable_days}=    Get From Dictionary    ${current_row}    AG
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlTermsAndConditions_ddlPremiumPayableDaysType'])[1]
    ...    ${premium_payable_days}
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

Risk Creation[Renewal]
    Fill Insured Details[Renewal]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Submission
    Fill Pricing Details[Renewal]

Fill Insured Details[Renewal]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Submission Details'])[1]    20s
    Extract Risk Number    Renewal
    Scroll Element Into View
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Press Keys    None    PAGE_DOWN
    Wait Until Keyword Succeeds
    ...    3x
    ...    5s
    ...    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    Scroll Element Into View
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
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
        Click Element When Visible    (//span[normalize-space()='Continue'])[1]
    END
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Scroll Element Into View
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Press Keys    None    PAGE_DOWN
    Wait Until Keyword Succeeds
    ...    3x
    ...    5s
    ...    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    10s
    ${value2}=    Is Element Visible    (//span[normalize-space()='Continue'])[1]
    IF    ${value2} == True
        Click Element When Visible    (//span[normalize-space()='Continue'])[1]
    END

Fill Pricing Details[Renewal]
    Wait Until Element Is Visible    (//a[normalize-space()='Pricing'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Deductible And Coverage'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Renewal Comparison'])[1]    20s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    ${comments}=    Get From Dictionary    ${current_row}    AQ
    Input Text When Element Is Visible
    ...    (//textarea[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlRenewalComparison_txtComment'])[1]
    ...    ${comments}
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Quote Items'])[1]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Additional Terms And Conditions'])[1]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

Fill Pricing Details[Reissue]
    Wait Until Element Is Visible    (//a[normalize-space()='Pricing'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Deductible And Coverage'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Quote Items'])[1]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Additional Terms And Conditions'])[1]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

Fill Insured Details[Copy Risk]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    ${effdate}=    Get From Dictionary    ${current_row}    J
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
    ...    ${effdate}
    Extract Risk Number    Copy Risk
    Scroll Element Into View
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Press Keys    None    PAGE_DOWN
    Wait Until Keyword Succeeds
    ...    3x
    ...    5s
    ...    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    Scroll Element Into View
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
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
        Click Element When Visible    (//span[normalize-space()='Continue'])[1]
    END
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Scroll Element Into View
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Press Keys    None    PAGE_DOWN
    Wait Until Keyword Succeeds
    ...    3x
    ...    5s
    ...    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    10s
    ${value2}=    Is Element Visible    (//span[normalize-space()='Continue'])[1]
    IF    ${value2} == True
        Click Element When Visible    (//span[normalize-space()='Continue'])[1]
    END

Pre Bind Endorsement
    Click Element When Visible    (//input[@id='btnShowAllEndorsementTypes'])[1]
    ${endorsement_type}=    Get From Dictionary    ${current_row}    AH
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPreBindEndorsements_txtEndorsementType'])[1]
    ...    ${endorsement_type}
    Sleep    2s
    Click Element When Visible    (//a[normalize-space()='${endorsement_type}'])[1]
    Sleep    5s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPreBindEndorsements_btAddEndorsement'])[1]
    Sleep    3s
    ${additional_insurable_limit}=    Get From Dictionary    ${current_row}    AI
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_txtBoxAdditionalInsurableLimit'])[1]
    ...    ${additional_insurable_limit}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementStandardButtons_btnSubmit'])[1]
    Sleep    3s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Sleep    5s
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    Check All Subjectivities
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Sleep    5s

Quote
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Finalize'])[1]    20s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    ${comments_under_quote}=    Get From Dictionary    ${current_row}    AJ
    Input Text When Element Is Visible
    ...    (//textarea[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlFinalize_txtComments'])[1]
    ...    ${comments_under_quote}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnGenerateQuote'])[1]

Risk UW eFile
    [Arguments]    ${download_policy}
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

Pre Bind Endorsement[Reissue]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Sleep    5s
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    Check All Subjectivities
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Sleep    5s

Quote[Reissue]
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Finalize'])[1]    20s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnGenerateQuote'])[1]

Quote[Renewal]
    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Finalize'])[1]    20s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[2]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    ${comments_under_quote}=    Get From Dictionary    ${current_row}    AJ
    Input Text When Element Is Visible
    ...    (//textarea[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlFinalize_txtComments'])[1]
    ...    ${comments_under_quote}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnGenerateQuote'])[1]

Ready To Bind
    Wait Until Element Is Visible    (//a[normalize-space()='Account Summary'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Account Summary'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnComplete'])[1]
    Sleep    5s
    Wait Until Element Is Visible    (//a[normalize-space()='Bind'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Bind'])[1]
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl01'])[1]
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnBind'])[1]

Ready To Bind[Reissue]
    Wait Until Element Is Visible    (//a[normalize-space()='Bind'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Bind'])[1]
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl01'])[1]
    Sleep    5s
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnBind'])[1]

Book
    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Book and Issue'])[1]
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    ${lead_carrier}=    Get From Dictionary    ${current_row}    AK
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlBookAndIssue_rptAdditionalLeadCarriers_ctl00_txtLeadCarrier'])[1]
    ...    ${lead_carrier}
    ${policy_number}=    Get From Dictionary    ${current_row}    AL
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlBookAndIssue_rptAdditionalLeadCarriers_ctl00_txtPolicyNumber'])[1]
    ...    ${policy_number}
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnDraft'])[1]

Issue Policy
    Wait Until Element Is Visible    (//a[normalize-space()='Issue Policy'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Issue Policy'])[1]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    ${endorsement_number}=    Get From Dictionary    ${current_row}    AM
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlBrokerEndorsements_txtEndorsementNo'])[1]
    ...    ${endorsement_number}
    ${endorsement_title}=    Get From Dictionary    ${current_row}    AN
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlBrokerEndorsements_txtEndorsementTitle'])[1]
    ...    ${endorsement_title}
    Choose File
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlBrokerEndorsements_fuExternalFormPdf'])[1]
    ...    ${CURDIR}${/}EndorsementBroker.pdf
    Sleep    3s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnIssue'])[1]

Book[Reissue]
    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Book and Issue'])[1]
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnDraft'])[1]

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
    Sleep    5s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonViewCurrent'])[1]
    Verify Status    (//legend[@id='fsMainContentLegend'])[1]    View Risk

Post Bind Endorsement
    Wait Until Element Is Visible    (//a[normalize-space()='Endorsement'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Endorsement'])[1]
    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    AO
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_txtEndorsementType'])[1]
    ...    ${endorsement_type_list}
    Sleep    2s
    Click Element When Visible    (//a[normalize-space()='${endorsement_type_list}'])[1]
    Sleep    5s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btEditEndorsement'])[1]
    Sleep    3s
    TRY
        ${policy_expiry_date}=    Get From Dictionary    ${current_row}    AP
        Input Text
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_DatePickerPolicyExpiryDate_textDate'])[1]
        ...    ${policy_expiry_date}
    EXCEPT
        ${policy_expiry_date}=    Get From Dictionary    ${current_row}    AP
        Input Text
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_DatePickerPolicyExpiryDate_textDate'])[1]
        ...    ${policy_expiry_date}
        Click Element When Visible    (//span[normalize-space()='Additional Premium'])[1]
        ${endorsement_premium}=    Get From Dictionary    ${current_row}    AO
        Input Text When Element Is Visible
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_endorsementPremiums_txtTotalEndorsmentPremiumValue'])[1]
        ...    1000
        Click Element When Visible    (//span[normalize-space()='Return Premium'])[1]
        Click Element When Visible
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementStandardButtons_btnSubmit'])[1]
    END
    Click Element When Visible    xpath=//div[@class='RiskHeaderColumnOne']//div[1]

Renewal
    Wait Until Element Is Visible    (//a[normalize-space()='Renew'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Renew'])[1]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Submission
    Risk Creation[Renewal]
    Pre Bind Endorsement[Reissue]
    Quote[Renewal]
    Ready To Bind
    Book[Reissue]
    Issue Policy

Reissue
    Wait Until Element Is Visible    (//a[normalize-space()='Reissue'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Reissue'])[1]
    Sleep    10s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[7]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Quoted (Pending, In Revision, RI)
    Wait Until Element Is Visible    (//a[normalize-space()='Edit Quote'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Edit Quote'])[1]
    Fill Pricing Details[Reissue]
    Pre Bind Endorsement[Reissue]
    Quote[Reissue]
    Ready To Bind[Reissue]
    Book[Reissue]
    Issue Policy

Copy Risk
    Wait Until Element Is Visible    (//a[normalize-space()='Copy Risk'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Copy Risk'])[1]
    Sleep    10s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[4]
    Fill Insured Details[Copy Risk]
    Fill Pricing Details[Reissue]
    Pre Bind Endorsement[Reissue]
    Quote[Reissue]
    Ready To Bind
    Book[Reissue]
    Issue Policy

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
    Wait Until Keyword Succeeds    45x    2s    Verify Status Matches    ${location}    ${status_expected}

Verify Status Matches
    [Arguments]    ${location}    ${status_expected}
    ${status_actual}=    RPA.Browser.Selenium.Get Text    ${location}
    ${status_actual}=    Strip String    ${status_actual}
    ${match}=    Evaluate    '${status_expected}' in '${status_actual}' or ('${status_expected}' == 'Submission (Pricing In Progress)' and ('Submission' in '${status_actual}' or 'Pending' in '${status_actual}'))
    Should Be True    ${match}    Expected status '${status_expected}' but got '${status_actual}'

Safe Click Element
    [Arguments]    ${locator}
    Wait Until Keyword Succeeds    15x    2s    Click Element When Visible    ${locator}

Safe Click Element By Index
    [Arguments]    ${locator}    ${index}
    Wait Until Keyword Succeeds    5x    2s    Click Element By Index And Wait    ${locator}    ${index}

Click Element By Index And Wait
    [Arguments]    ${locator}    ${index}
    ${elements}=    Get WebElements    ${locator}
    Log    Found ${elements.__len__()} elements
    ${count}=    Get Length    ${elements}
    ${target_index}=    Evaluate    ${index} if ${count} > ${index} else 0
    Click Element When Visible    ${elements}[${target_index}]

Check All Subjectivities
    FOR    ${index}    IN RANGE    1    20
        ${idx}=    Convert To String    ${index}
        ${idx}=    Run Keyword If    ${index} < 10    Set Variable    0${index}    ELSE    Set Variable    ${index}
        
        ${mandatory_xpath}=    Set Variable    //input[contains(@id, 'repeaterSubjectivities_ctl${idx}_cbMandatory')]
        ${satisfied_xpath}=    Set Variable    //input[contains(@id, 'repeaterSubjectivities_ctl${idx}_cbSatisfied')]
        
        ${exists}=    Run Keyword And Return Status    Page Should Contain Element    xpath=${mandatory_xpath}
        IF    ${exists} == True
            ${mandatory_checked}=    Execute Javascript    var el = document.evaluate("${mandatory_xpath}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; return el ? el.checked : false;
            IF    ${mandatory_checked} == False
                Execute Javascript    var el = document.evaluate("${mandatory_xpath}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; if(el) el.click();
                Sleep    1s
            END
            
            ${satisfied_checked}=    Execute Javascript    var el = document.evaluate("${satisfied_xpath}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; return el ? el.checked : false;
            IF    ${satisfied_checked} == False
                Execute Javascript    var el = document.evaluate("${satisfied_xpath}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; if(el) el.click();
                Sleep    1s
            END
        ELSE
            BREAK
        END
    END

Extract Risk Number
    [Arguments]    ${flow_type}
    Wait Until Keyword Succeeds    15x    2s    Wait Until Element Is Visible    xpath=//*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"] | //*[contains(@id, 'btnReturnToRiskSummary')]    15s
    ${risk_number}=    RPA.Browser.Selenium.Get Text    xpath=//*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"] | //*[contains(@id, 'btnReturnToRiskSummary')]
    Evaluate    open(r'${CURDIR}${/}risk_numbers.txt', 'a', encoding='utf-8').write($flow_type + ' - Risk Number: ' + $risk_number.strip() + chr(10))
    Log    Extracted ${flow_type} Risk Number: ${risk_number.strip()}
