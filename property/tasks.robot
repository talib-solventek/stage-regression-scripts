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
Property
    Open Workbook    ${DATA_SOURCE}
    ${excel_rows}=    Read Worksheet    Property
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
    ...    (//span[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]
    ...    Submission (Pricing In Progress)
    Quote
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]
    ...    Quoted
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
    Open Available Browser    ${URL}    maximized=${TRUE}    browser_selection=edge
    Set Selenium Implicit Wait    30s
    Click Element When Visible    xpath=//a[normalize-space()='Create New Risk >']
    Click Element When Visible    xpath=(//a[normalize-space()='Property'])[1]
    Sleep    10s
    Click Element When Visible
    ...    id:ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlClearanceSearch_CreateInsured
    Sleep    5s
    Set Anchor    type:WindowControl
    Sleep    5s
    ${locator}=    Send Keys    keys=${NUSER}{TAB}${NUSERPASSWORD}    send_enter=${TRUE}
    Send Keys    keys={RETURN}    send_enter=${TRUE}
    Sleep    10s

Risk Creation[First Run]
    Fill Insured Details[First Run]
    Verify Status
    ...    (//span[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]
    ...    Submission (SOV In Progress)
    Fill Pricing Details[First Run]

Fill Insured Details[First Run]
    Sleep    10s
    Click Element When Visible    (//span[normalize-space()='No'])[1]
    ${iname}=    Get From Dictionary    ${current_row}    D
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_TextBoxFirstNamedInsured1'])[1]
    ...    ${iname}
    ${country}=    Get From Dictionary    ${current_row}    E
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_ddCountry'])[1]
    ...    ${country}
    Sleep    5s
    ${adl1}=    Get From Dictionary    ${current_row}    F
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxROWAddressLine1'])[1]
    ...    ${adl1}
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
    ...    Adelaide
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]
    ...    AUS
    Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    Click Element When Visible    (//span[normalize-space()='Corporate'])[1]
    Sleep    5s
    ${department}=    Get From Dictionary    ${current_row}    G
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListDepartment'])[1]
    ...    Retail
    Sleep    20s
    ${effdate}=    Get From Dictionary    ${current_row}    H
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
    ...    ${effdate}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    ${brokfirm}=    Get From Dictionary    ${current_row}    J
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_TextBoxBrokerFirm'])[1]
    ...    ${brokfirm}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_ButtonBrokerFirmSearchContains'])[1]
    Sleep    10s
    ${table_elements}=    Get WebElements
    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]
    Log    Found ${table_elements.__len__()} elements in the broker info table
    Click Element    ${table_elements}[2]
    Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    ${brokcont}=    Get From Dictionary    ${current_row}    K
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]
    ...    ${brokcont}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_ButtonBrokerContactSearchStartsWith'])[1]
    Sleep    20s
    ${table_elements}=    Get WebElements
    ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]
    Log    Found ${table_elements.__len__()} elements in the broker contact table
    Click Element    ${table_elements}[0]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Pricing Details[First Run]
    Wait Until Element Is Visible    (//a[normalize-space()='Location Management'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlLocationManagement_btnImportLocationDialog'])[1]
    Sleep    10s
    Choose File
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlLocationManagement_fuLocationCSV'])[1]
    ...    ${CURDIR}${/}Location.xlsx
    Click Element When Visible    (//span[normalize-space()='Import'])[1]
    Sleep    5s
    Click Element When Visible    (//a[@class='btn btn-info'])[1]
    Input Text When Element Is Visible    (//input[@id='i0116'])[1]    n9970632@libertymutual.com
    Click Element When Visible    (//input[@id='idSIButton9'])[1]
    Input Text When Element Is Visible    (//input[@id='i0118'])[1]    wm$qY@32
    Click Element When Visible    (//input[@id='idSIButton9'])[1]
    Sleep    20s
	Capture Page Screenshot	Location
    Click Element When Visible    (//button[normalize-space()='Next'])[1]
    Sleep    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlLocationManagement_btnAddLocation'])[1]
    Input Text When Element Is Visible    (//input[@id='Name'])[1]    test
    Input Text When Element Is Visible    (//input[@id='country'])[1]    Bangladesh
    Input Text When Element Is Visible    (//input[@id='Latitude'])[1]    78
    Input Text When Element Is Visible    (//input[@id='Longitude'])[1]    170
    Click Element When Visible    (//button[normalize-space()='Match & Cleanse'])[1]
    Sleep    5s
    Click Element When Visible    (//button[normalize-space()='Next'])[1]
    Click Element When Visible    (//a[normalize-space()='Location Details'])[1]
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLocationInformation_ddlRatingSegment'])[1]
    ...    Accommodation
    Sleep    2s
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLocationInformation_ddlOccupancy'])[1]
    ...    Bed & Breakfast - good class
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLocationInformation_ddlAssetCurrency'])[1]
    ...    USD
    ${pd_building}=    Get From Dictionary    ${current_row}    P
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLocationInformation_txtPdBuilding'])[1]
    ...    500
    ${pd_mach_equip}=    Get From Dictionary    ${current_row}    P
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLocationInformation_txtPdMachineryAndEquipment'])[1]
    ...    500
    ${pd_contents}=    Get From Dictionary    ${current_row}    P
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLocationInformation_txtPdContents'])[1]
    ...    500
    ${pd_tiv}=    Get From Dictionary    ${current_row}    P
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLocationInformation_txtPDTIV'])[1]
    ...    1500
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[4]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnSave'])[1]
    Sleep    5s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl01'])[1]
    Sleep    60s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Ok'])[1]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Operational Policy Terms'])[1]    20s
    ${lead_follow}=    Get From Dictionary    ${current_row}    Z
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_ddlLeadFollow'])[1]
    ...    ${lead_follow}
    ${policy_wording}=    Get From Dictionary    ${current_row}    AA
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_ddlPolicyWording'])[1]
    ...    ${policy_wording}
    ${combined_loss_limit}=    Get From Dictionary    ${current_row}    AB
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_txtCombinedLossLimit'])[1]
    ...    ${combined_loss_limit}
    ${cbi}=    Get From Dictionary    ${current_row}    AC
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_txtSublimitsCbi'])[1]
    ...    ${cbi}
    ${extra_expense}=    Get From Dictionary    ${current_row}    AD
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_txtSublimitsExtraExpense'])[1]
    ...    ${extra_expense}
    ${PD}=    Get From Dictionary    ${current_row}    AH
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_txtDeductiblesPD'])[1]
    ...    ${PD}
    ${bi_days}=    Get From Dictionary    ${current_row}    AI
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_txtDeductiblesBIDays'])[1]
    ...    ${bi_days}
    ${cbi_days}=    Get From Dictionary    ${current_row}    AJ
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_txtDeductiblesCBIDays'])[1]
    ...    ${cbi_days}
    ${service_interuption}=    Get From Dictionary    ${current_row}    AK
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_ddlServiceInterruption'])[1]
    ...    ${service_interuption}
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Operational Location Terms'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Operational Rates'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Natural Perils Policy Terms'])[1]
    ${sublimit_type1}=    Get From Dictionary    ${current_row}    AL
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlNaturalPerilTerms_rptNaturalPerilsTerms_ctl01_ddlSublimitType'])[1]
    ...    ${sublimit_type1}
    ${sublimit_type2}=    Get From Dictionary    ${current_row}    AM
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlNaturalPerilTerms_rptNaturalPerilsTerms_ctl02_ddlSublimitType'])[1]
    ...    ${sublimit_type2}
    ${sublimit_type3}=    Get From Dictionary    ${current_row}    AN
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlNaturalPerilTerms_rptNaturalPerilsTerms_ctl03_ddlSublimitType'])[1]
    ...    ${sublimit_type3}
    ${bi_days1}=    Get From Dictionary    ${current_row}    AO
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlNaturalPerilTerms_rptNaturalPerilsTerms_ctl01_txtBiDays'])[1]
    ...    ${bi_days1}
    ${bi_days2}=    Get From Dictionary    ${current_row}    AP
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlNaturalPerilTerms_rptNaturalPerilsTerms_ctl02_txtBiDays'])[1]
    ...    ${bi_days2}
    ${bi_days3}=    Get From Dictionary    ${current_row}    AQ
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlNaturalPerilTerms_rptNaturalPerilsTerms_ctl03_txtBiDays'])[1]
    ...    ${bi_days3}
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Natural Perils Location Terms'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Natural Peril Rates'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Loss Experience'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Discretionary Modifiers'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Excess Of Loss'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary Of Pricing'])[1]    20s
    ${liu_share}=    Get From Dictionary    ${current_row}    AR
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlSummaryOfPricing_txtLiuSharePercentage'])[1]
    ...    ${liu_share}
    ${commission}=    Get From Dictionary    ${current_row}    AS
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlSummaryOfPricing_txtCommission'])[1]
    ...    ${commission}
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Fac'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='LibertyIndex'])[1]    20s
    Click Element When Visible    (//span[normalize-space()='+ve'])[1]
    ${renewal_years}=    Get From Dictionary    ${current_row}    AT
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlLibertyIndex_txtRenewalYears'])[1]
    ...    ${renewal_years}
    ${claims_approach}=    Get From Dictionary    ${current_row}    AU
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlLibertyIndex_ddlClaimsApproach'])[1]
    ...    ${claims_approach}
    ${opinion}=    Get From Dictionary    ${current_row}    AV
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlLibertyIndex_ddlExecOpinion'])[1]
    ...    ${opinion}
    ${cyber_type}=    Get From Dictionary    ${current_row}    AW
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlLibertyIndex_ddlCyberType'])[1]
    ...    ${cyber_type}
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Pre-Bind Allocation'])[1]    20s
    ${written%}=    Get From Dictionary    ${current_row}    AX
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlBind_txtWrittenLinePercent'])[1]
    ...    15
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Pre-Bind Deductions'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Tax'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

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
    Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Pricing Details[Copy Risk]
    Wait Until Element Is Visible    (//a[normalize-space()='Location Management'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Sleep    5s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Ok'])[1]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Operational Policy Terms'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Operational Location Terms'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Operational Rates'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Natural Perils Policy Terms'])[1]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Natural Perils Location Terms'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Natural Peril Rates'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Loss Experience'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Discretionary Modifiers'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Excess Of Loss'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary Of Pricing'])[1]    20s
    ${liu_share}=    Get From Dictionary    ${current_row}    AR
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlSummaryOfPricing_txtLiuSharePercentage'])[1]
    ...    ${liu_share}
    ${commission}=    Get From Dictionary    ${current_row}    AS
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlSummaryOfPricing_txtCommission'])[1]
    ...    ${commission}
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Fac'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='LibertyIndex'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Pre-Bind Allocation'])[1]    20s
    ${written%}=    Get From Dictionary    ${current_row}    AX
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlBind_txtWrittenLinePercent'])[1]
    ...    15
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Pre-Bind Deductions'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Tax'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

Risk Creation[Renewal]
    Fill Insured Details[Renewal]
    Verify Status
    ...    (//span[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1][1]
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
    Wait Until Element Is Visible    (//a[normalize-space()='Location Management'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Ok'])[1]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Operational Policy Terms'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Operational Location Terms'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Operational Rates'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Natural Perils Policy Terms'])[1]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Natural Perils Location Terms'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Natural Peril Rates'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Loss Experience'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Discretionary Modifiers'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Excess Of Loss'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary Of Pricing'])[1]    20s
    ${liu_share}=    Get From Dictionary    ${current_row}    AR
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlSummaryOfPricing_txtLiuSharePercentage'])[1]
    ...    ${liu_share}
    ${commission}=    Get From Dictionary    ${current_row}    AS
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlSummaryOfPricing_txtCommission'])[1]
    ...    ${commission}
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Fac'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='LibertyIndex'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Pre-Bind Allocation'])[1]    20s
    ${written}=    Get From Dictionary    ${current_row}    AX
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlBind_txtWrittenLinePercent'])[1]
    ...    15
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Pre-Bind Deductions'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Tax'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

Fill Pricing Details[Reissue]
    Wait Until Element Is Visible    (//a[normalize-space()='Location Management'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Sleep    5s
    Wait Until Element Is Visible    (//a[normalize-space()='Operational Policy Terms'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Operational Location Terms'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Operational Rates'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Natural Perils Policy Terms'])[1]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Natural Perils Location Terms'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Natural Peril Rates'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Loss Experience'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Discretionary Modifiers'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Excess Of Loss'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Summary Of Pricing'])[1]    20s
    ${liu_share}=    Get From Dictionary    ${current_row}    AR
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlSummaryOfPricing_txtLiuSharePercentage'])[1]
    ...    ${liu_share}
    ${commission}=    Get From Dictionary    ${current_row}    AS
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlSummaryOfPricing_txtCommission'])[1]
    ...    ${commission}
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Fac'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='LibertyIndex'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Pre-Bind Allocation'])[1]    20s
    ${written%}=    Get From Dictionary    ${current_row}    AX
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlBind_txtWrittenLinePercent'])[1]
    ...    15
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Pre-Bind Deductions'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Tax'])[1]    20s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

Pre Bind Endorsement
    Click Element When Visible    (//input[@id='btnShowAllEndorsementTypes'])[1]
    ${endorsement_type}=    Get From Dictionary    ${current_row}    AP
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPreBindEndorsements_txtEndorsementType'])[1]
    ...    ${endorsement_type}
    Sleep    2s
    Click Element When Visible    (//a[normalize-space()='${endorsement_type}'])[1]
    Sleep    2s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPreBindEndorsements_btAddEndorsement'])[1]
    Sleep    3s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Sleep    5s

Quote
    Wait Until Element Is Visible    (//a[normalize-space()='Quote Final'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnGenerateQuote'])[1]
    ${confirm}=    Is Element Visible    (//span[normalize-space()='Confirm'])[1]
    IF    ${confirm} == True
        Click Element When Visible    (//span[normalize-space()='Confirm'])[1]
    END

Quote[Copy Risk]
    Wait Until Element Is Visible    (//a[normalize-space()='Quote Final'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnGenerateQuote'])[1]
    ${confirm}=    Is Element Visible    (//span[normalize-space()='Confirm'])[1]
    IF    ${confirm} == True
        Click Element When Visible    (//span[normalize-space()='Confirm'])[1]
    END

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
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnGenerateQuote'])[1]
    ${confirm}=    Is Element Visible    (//span[normalize-space()='Confirm'])[1]
    IF    ${confirm} == True
        Click Element When Visible    (//span[normalize-space()='Confirm'])[1]
    END

Ready To Bind
    Click Element When Visible    (//a[normalize-space()='Bind'])[1]
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl01'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnNext'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnBind'])[1]

Book
    [Arguments]    ${download_policy}
    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Book and Issue'])[1]
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    20s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Sleep    5s
    ${due_days}=    Get From Dictionary    ${current_row}    AZ
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlBook_ctrlBookInstallmentPlan_tbDueDays'])[1]
    ...    ${due_days}
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ButtonBook'])[1]

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
    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    BA
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_txtEndorsementType'])[1]
    ...    ${endorsement_type_list}
    Sleep    2s
    Click Element When Visible    (//a[normalize-space()='${endorsement_type_list}'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btAddEndorsement'])[1]
    Sleep    3s
    ${endorsement_effective_date}=    Get From Dictionary    ${current_row}    BB
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_DatePickerEndorsementEffectiveDate_textDate'])[1]
    ...    ${endorsement_effective_date}
    TRY
        ${expiry_date}=    Get From Dictionary    ${current_row}    BC
        Input Text When Element Is Visible
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_DatePickerPolicyExpiryDate_textDate'])[1]
        ...    ${expiry_date}
    EXCEPT
        ${expiry_date}=    Get From Dictionary    ${current_row}    BC
        Input Text When Element Is Visible
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_DatePickerPolicyExpiryDate_textDate'])[1]
        ...    ${expiry_date}
    END
    TRY
        ${due_days}=    Get From Dictionary    ${current_row}    BD
        Input Text When Element Is Visible
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_tbDueDays'])[1]
        ...    ${due_days}
    EXCEPT
        ${due_days}=    Get From Dictionary    ${current_row}    BD
        Input Text When Element Is Visible
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_tbDueDays'])[1]
        ...    ${due_days}
        Click Element When Visible    (//span[normalize-space()='No Premium'])[1]
        Sleep    20s
        Click Element When Visible
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementStandardButtons_btnSubmit'])[1]
    END
    Sleep    3s
    Click Element When Visible    xpath=//div[@class='RiskHeaderColumnOne']//div[1]
    Sleep    100s

Renewal
    Wait Until Element Is Visible    (//a[normalize-space()='Renew'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Renew'])[1]
    Sleep    10s
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_labelStatus'])[1]
    ...    Submission
    Risk Creation[Renewal]
    Quote[Reissue]
    Ready To Bind
    Book    False

Reissue
    Wait Until Element Is Visible    (//a[normalize-space()='Reissue'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Reissue'])[1]
    Sleep    10s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[7]
    Verify Status
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]
    ...    Quoted (Pending, In Revision, RI)
    Wait Until Element Is Visible    (//a[normalize-space()='Edit Quote'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Edit Quote'])[1]
    Fill Pricing Details[Reissue]
    Quote[Reissue]
    Ready To Bind
    Book    False

Copy Risk
    Wait Until Element Is Visible    (//a[normalize-space()='Copy Risk'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Copy Risk'])[1]
    Sleep    10s
    Click Element When Visible    //span[contains(.,"Original Risks's Account")]
    Fill Insured Details[Copy Risk]
    Fill Pricing Details[Copy Risk]
    Quote[Copy Risk]
    Ready To Bind
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
